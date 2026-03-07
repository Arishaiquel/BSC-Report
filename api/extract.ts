import type { IncomingMessage, ServerResponse } from "http";
import { simpleParser } from "mailparser";
import { JSDOM } from "jsdom";
import ExcelJS from "exceljs";
import AdmZip from "adm-zip";

export const config = {
  api: {
    bodyParser: false,
  },
};

type UploadedFile = {
  buffer: Buffer;
  originalname: string;
  mimetype?: string;
};

function normalizeSpaces(value: string): string {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function cleanAdvisorName(raw: string): string {
  let cleaned = normalizeSpaces(raw);
  while (/\([^()]*\)\s*$/.test(cleaned)) {
    cleaned = cleaned.replace(/\s*\([^()]*\)\s*$/, "").trim();
  }
  cleaned = cleaned.replace(/\s*-\s*iPASS Approval.*$/i, "").trim();
  return cleaned;
}

function extractAdvisorName(doc: Document): string {
  const activityHeader = Array.from(doc.querySelectorAll("b, strong")).find(
    (el) => (el.textContent?.trim().toLowerCase() || "") === "activity and comments",
  );

  let activityTable: HTMLTableElement | null = null;
  if (activityHeader) {
    let current: Element | null =
      activityHeader.closest("p")?.nextElementSibling ||
      activityHeader.parentElement?.nextElementSibling ||
      activityHeader.nextElementSibling;

    while (current) {
      if (current.tagName === "TABLE") {
        activityTable = current as HTMLTableElement;
        break;
      }

      const nestedTable = current.querySelector("table");
      if (nestedTable) {
        activityTable = nestedTable as HTMLTableElement;
        break;
      }

      if (current.querySelector("b, strong")) break;
      current = current.nextElementSibling;
    }
  }

  if (!activityTable) {
    activityTable =
      (Array.from(doc.querySelectorAll("table")).find((table) => {
        const headerCells = Array.from(table.querySelectorAll("tr:first-child th, tr:first-child td"))
          .map((cell) => normalizeSpaces(cell.textContent || "").toLowerCase());
        return headerCells.some((header) => header === "author");
      }) as HTMLTableElement | undefined) || null;
  }

  if (!activityTable) return "";

  const rows = Array.from(activityTable.querySelectorAll("tr"));
  if (rows.length < 2) return "";

  const headers = Array.from(rows[0].querySelectorAll("th, td"))
    .map((cell) => normalizeSpaces(cell.textContent || "").toLowerCase());
  const authorIdx = headers.findIndex((header) => header === "author");
  if (authorIdx === -1) return "";

  for (let i = 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll("td"));
    if (authorIdx >= cells.length) continue;
    const value = cleanAdvisorName(cells[authorIdx].textContent || "");
    if (value) return value;
  }

  return "";
}

function sendJson(
  res: ServerResponse,
  statusCode: number,
  payload: Record<string, string>,
) {
  res.statusCode = statusCode;
  res.setHeader("Content-Type", "application/json");
  res.end(JSON.stringify(payload));
}

function readRequestBody(req: IncomingMessage): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];

    req.on("data", (chunk) => {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

async function parseUploads(req: IncomingMessage): Promise<UploadedFile[]> {
  const contentType = req.headers["content-type"];
  if (!contentType || !contentType.includes("multipart/form-data")) {
    return [];
  }

  const bodyBuffer = await readRequestBody(req);
  const multipartRequest = new Request("http://localhost/api/extract", {
    method: "POST",
    headers: { "content-type": contentType },
    body: bodyBuffer,
  });
  const formData = await multipartRequest.formData();

  const files: UploadedFile[] = [];
  for (const entry of formData.getAll("files")) {
    if (entry instanceof File) {
      files.push({
        buffer: Buffer.from(await entry.arrayBuffer()),
        originalname: entry.name || "",
        mimetype: entry.type || undefined,
      });
    }
  }

  return files;
}

export default async function handler(req: IncomingMessage, res: ServerResponse) {
  if (req.method !== "POST") {
    sendJson(res, 405, { message: "Method not allowed" });
    return;
  }

  try {
    const files = await parseUploads(req);
    if (!files.length) {
      sendJson(res, 400, { message: "No files uploaded" });
      return;
    }

    const results: Array<Record<string, string>> = [];
    const allFiles: UploadedFile[] = [];

    for (const file of files) {
      const originalName = file.originalname.toLowerCase();
      if (file.mimetype === "application/zip" || originalName.endsWith(".zip")) {
        const zip = new AdmZip(file.buffer);
        const zipEntries = zip.getEntries();
        for (const entry of zipEntries) {
          if (entry.entryName.toLowerCase().endsWith(".eml") && !entry.isDirectory) {
            allFiles.push({
              buffer: entry.getData(),
              originalname: entry.entryName,
            });
          }
        }
      } else if (originalName.endsWith(".eml")) {
        allFiles.push(file);
      }
    }

    for (const file of allFiles) {
      const parsed = await simpleParser(file.buffer);
      const html = parsed.html || parsed.textAsHtml || "";
      const dom = new JSDOM(String(html));
      const doc = dom.window.document;
      const advisorName = extractAdvisorName(doc);

      let policyNumber = "";
      const bodyText = doc.body.textContent || "";
      const policyMatch = bodyText.match(/\((P\d+)\)/);
      if (policyMatch) {
        policyNumber = policyMatch[1];
      }

      let submissionDate = "";
      const filename = file.originalname || "";
      const filenameDateMatch = filename.match(/_(\d{4})(\d{2})(\d{2})_/);
      if (filenameDateMatch) {
        const [, year, month, day] = filenameDateMatch;
        submissionDate = `${parseInt(day, 10)}/${parseInt(month, 10)}/${year}`;
      } else {
        submissionDate = parsed.date ? parsed.date.toLocaleDateString("en-GB") : "";
      }

      let buyAmount = 0;
      let buyForeign = "";
      const buyProducts: string[] = [];

      let rspAmount = 0;
      let rspForeign = "";
      const rspProducts: string[] = [];

      const allowedProducts = ["Company Portfolio", "DPMS", "Unit Trust"];

      const sections = Array.from(doc.querySelectorAll("b, strong")).filter((el) => {
        const text = el.textContent?.trim() || "";
        return (
          text === "Buy" ||
          text === "RSP Application" ||
          text.includes("Buy") ||
          text.includes("RSP Application")
        );
      });

      for (const header of sections) {
        const headerText = header.textContent?.trim() || "";
        const isBuy = headerText === "Buy";
        const isRsp = headerText === "RSP Application";

        if (!isBuy && !isRsp) continue;

        let current = header.parentElement?.nextElementSibling;

        while (current) {
          const currentText = current.textContent?.trim() || "";

          if (current.querySelector("b, strong")) {
            const innerText = current.querySelector("b, strong")?.textContent?.trim() || "";

            if (
              innerText !== headerText &&
              [
                "Switch",
                "Sell",
                "Rebalance",
                "RSP Amendment",
                "Consolidated Dividend",
                "ETF",
              ].some((s) => innerText.includes(s))
            ) {
              break;
            }

            if (
              innerText !== headerText &&
              (innerText === "Buy" || innerText === "RSP Application")
            ) {
              break;
            }
          }

          if (
            current.tagName === "P" &&
            current instanceof dom.window.HTMLElement &&
            current.style.textDecoration === "underline"
          ) {
            const pText = current.textContent?.trim() || "";
            if (
              pText !== headerText &&
              ["Switch", "Sell", "Rebalance", "RSP Amendment", "ETF"].some((s) =>
                pText.includes(s),
              )
            ) {
              break;
            }
          }

          if (current.tagName === "P") {
            const pText = currentText;
            if (pText && (pText.includes("ETF") || pText.includes("Exchange Traded Fund"))) {
              let next = current.nextElementSibling;
              while (
                next &&
                next.tagName !== "P" &&
                next.tagName !== "TABLE" &&
                !next.querySelector("b, strong")
              ) {
                next = next.nextElementSibling;
              }
              if (next && next.tagName === "TABLE") {
                current = next;
              }
            } else if (pText && allowedProducts.some((ap) => pText.includes(ap))) {
              const matchedProduct = allowedProducts.find((ap) => pText.includes(ap));
              if (matchedProduct) {
                if (isBuy) buyProducts.push(matchedProduct);
                else if (isRsp) rspProducts.push(matchedProduct);
              }
            }
          } else if (current.tagName === "TABLE") {
            const rows = current.querySelectorAll("tr");
            if (rows.length > 1) {
              const headers = Array.from(rows[0].querySelectorAll("th, td")).map(
                (h) => h.textContent?.trim() || "",
              );
              const amountIdx = headers.findIndex((h) => h.includes("Amount"));

              for (let i = 1; i < rows.length; i++) {
                const cols = Array.from(rows[i].querySelectorAll("td")).map(
                  (c) => c.textContent?.trim() || "",
                );
                if (amountIdx !== -1 && cols[amountIdx]) {
                  const textValue = cols[amountIdx];
                  const val = parseFloat(textValue.replace(/[^0-9.]/g, ""));
                  if (!isNaN(val)) {
                    const isForeign =
                      !textValue.includes("SGD") &&
                      (textValue.match(/[A-Z]{3}/) ||
                        cols.some((c) => c.match(/^(USD|EUR|GBP|AUD|JPY|HKD)$/)));

                    if (!isForeign) {
                      if (isBuy) buyAmount += val;
                      else if (isRsp) rspAmount += val;
                    } else {
                      const currency =
                        textValue.match(/[A-Z]{3}/)?.[0] ||
                        cols.find((c) => c.match(/^[A-Z]{3}$/)) ||
                        "Foreign";
                      const foreignEntry = `${currency} ${val.toFixed(2)}`;
                      if (isBuy) {
                        buyForeign = buyForeign ? `${buyForeign}, ${foreignEntry}` : foreignEntry;
                      } else if (isRsp) {
                        rspForeign = rspForeign ? `${rspForeign}, ${foreignEntry}` : foreignEntry;
                      }
                    }
                  }
                }
              }
            }
          }

          current = current.nextElementSibling;
        }
      }

      results.push({
        "Policy Number": policyNumber,
        "Submission Date": submissionDate,
        Buy:
          buyAmount > 0
            ? `SGD ${buyAmount.toLocaleString("en-US", {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              })}`
            : "",
        "RSP Application":
          rspAmount > 0
            ? `SGD ${rspAmount.toLocaleString("en-US", {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              })}`
            : "",
        "advisor name": advisorName,
        "Buy product": [...new Set(buyProducts)].join(", "),
        "RSP Application product": [...new Set(rspProducts)].join(", "),
        "Foreign Buy": buyForeign,
        "Foreign RSP Application": rspForeign,
      });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Extractions");

    worksheet.columns = [
      { header: "Policy Number", key: "Policy Number" },
      { header: "Submission Date", key: "Submission Date" },
      { header: "Buy", key: "Buy" },
      { header: "RSP Application", key: "RSP Application" },
      { header: "advisor name", key: "advisor name" },
      { header: "Buy product", key: "Buy product" },
      { header: "RSP Application product", key: "RSP Application product" },
      { header: "Foreign Buy", key: "Foreign Buy" },
      { header: "Foreign RSP Application", key: "Foreign RSP Application" },
    ];

    worksheet.addRows(results);

    const workbookData = await workbook.xlsx.writeBuffer();
    const buffer = Buffer.isBuffer(workbookData)
      ? workbookData
      : Buffer.from(workbookData);

    res.statusCode = 200;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader("Content-Disposition", "attachment; filename=extracted_data.xlsx");
    res.end(buffer);
  } catch (error) {
    console.error(error);
    sendJson(res, 500, { message: "Internal server error" });
  }
}
