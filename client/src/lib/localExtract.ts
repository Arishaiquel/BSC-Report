import JSZip from "jszip";
import PostalMime from "postal-mime";
import * as XLSX from "xlsx";

type LocalEmailFile = {
  originalname: string;
  raw: Uint8Array;
};

type ExtractionRow = {
  "Policy Number": string;
  "Submission Date": string;
  Buy: string;
  "RSP Application": string;
  "advisor name": string;
  "Buy product": string;
  "RSP Application product": string;
  "Foreign Buy": string;
  "Foreign RSP Application": string;
};

const ALLOWED_PRODUCTS = ["Company Portfolio", "DPMS", "Unit Trust"];
const STOP_SECTION_KEYWORDS = [
  "Switch",
  "Sell",
  "Rebalance",
  "RSP Amendment",
  "Consolidated Dividend",
  "ETF",
];

const WORKBOOK_COLUMNS: Array<keyof ExtractionRow> = [
  "Policy Number",
  "Submission Date",
  "Buy",
  "RSP Application",
  "advisor name",
  "Buy product",
  "RSP Application product",
  "Foreign Buy",
  "Foreign RSP Application",
];

function formatSgd(value: number): string {
  if (value <= 0) return "";
  return `${value.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function parseSubmissionDate(fileName: string, messageDate?: string): string {
  const filenameDateMatch = fileName.match(/_(\d{4})(\d{2})(\d{2})_/);
  if (filenameDateMatch) {
    const [, year, month, day] = filenameDateMatch;
    return `${parseInt(day, 10)}/${parseInt(month, 10)}/${year}`;
  }

  if (!messageDate) return "";
  const parsedDate = new Date(messageDate);
  if (Number.isNaN(parsedDate.getTime())) return "";
  return parsedDate.toLocaleDateString("en-GB");
}

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

      if (current.querySelector("b, strong")) {
        break;
      }
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

  const headerCells = Array.from(rows[0].querySelectorAll("th, td"))
    .map((cell) => normalizeSpaces(cell.textContent || "").toLowerCase());
  const authorIndex = headerCells.findIndex((header) => header === "author");
  if (authorIndex === -1) return "";

  for (let i = 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll("td"));
    if (authorIndex >= cells.length) continue;

    const value = cleanAdvisorName(cells[authorIndex].textContent || "");
    if (value) return value;
  }

  return "";
}

function parseEmailHtml(
  html: string,
  originalname: string,
  messageDate?: string,
): ExtractionRow {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html || "", "text/html");
  const advisorName = extractAdvisorName(doc);

  let policyNumber = "";
  const bodyText = doc.body.textContent || "";
  const policyMatch = bodyText.match(/\((P\d+)\)/);
  if (policyMatch) {
    policyNumber = policyMatch[1];
  }

  let buyAmount = 0;
  let buyForeign = "";
  const buyProducts: string[] = [];

  let rspAmount = 0;
  let rspForeign = "";
  const rspProducts: string[] = [];

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

    let current: Element | null = header.parentElement?.nextElementSibling || null;

    while (current) {
      const currentText = current.textContent?.trim() || "";

      const nestedHeader = current.querySelector("b, strong");
      if (nestedHeader) {
        const innerText = nestedHeader.textContent?.trim() || "";

        if (
          innerText !== headerText &&
          STOP_SECTION_KEYWORDS.some((keyword) => innerText.includes(keyword))
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
        current instanceof HTMLElement &&
        current.style.textDecoration === "underline"
      ) {
        const pText = current.textContent?.trim() || "";
        if (
          pText !== headerText &&
          ["Switch", "Sell", "Rebalance", "RSP Amendment", "ETF"].some((keyword) =>
            pText.includes(keyword),
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
        } else if (pText && ALLOWED_PRODUCTS.some((product) => pText.includes(product))) {
          const matchedProduct = ALLOWED_PRODUCTS.find((product) => pText.includes(product));
          if (matchedProduct) {
            if (isBuy) buyProducts.push(matchedProduct);
            else if (isRsp) rspProducts.push(matchedProduct);
          }
        }
      } else if (current.tagName === "TABLE") {
        const rows = current.querySelectorAll("tr");
        if (rows.length > 1) {
          const headers = Array.from(rows[0].querySelectorAll("th, td")).map(
            (cell) => cell.textContent?.trim() || "",
          );
          const amountIdx = headers.findIndex((headerCell) => headerCell.includes("Amount"));

          for (let i = 1; i < rows.length; i++) {
            const cols = Array.from(rows[i].querySelectorAll("td")).map(
              (cell) => cell.textContent?.trim() || "",
            );
            if (amountIdx !== -1 && cols[amountIdx]) {
              const textValue = cols[amountIdx];
              const val = parseFloat(textValue.replace(/[^0-9.]/g, ""));
              if (!Number.isNaN(val)) {
                const isForeign =
                  !textValue.includes("SGD") &&
                  (textValue.match(/[A-Z]{3}/) ||
                    cols.some((col) => col.match(/^(USD|EUR|GBP|AUD|JPY|HKD)$/)));

                if (!isForeign) {
                  if (isBuy) buyAmount += val;
                  else if (isRsp) rspAmount += val;
                } else {
                  const currency =
                    textValue.match(/[A-Z]{3}/)?.[0] ||
                    cols.find((col) => col.match(/^[A-Z]{3}$/)) ||
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

  return {
    "Policy Number": policyNumber,
    "Submission Date": parseSubmissionDate(originalname, messageDate),
    Buy: formatSgd(buyAmount),
    "RSP Application": formatSgd(rspAmount),
    "advisor name": advisorName,
    "Buy product": [...new Set(buyProducts)].join(", "),
    "RSP Application product": [...new Set(rspProducts)].join(", "),
    "Foreign Buy": buyForeign,
    "Foreign RSP Application": rspForeign,
  };
}

async function collectEmailFiles(files: FileList): Promise<LocalEmailFile[]> {
  const collectedFiles: LocalEmailFile[] = [];

  for (const file of Array.from(files)) {
    const lowerName = file.name.toLowerCase();

    if (lowerName.endsWith(".zip")) {
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      const zipEntries = Object.values(zip.files);

      for (const entry of zipEntries) {
        if (entry.dir || !entry.name.toLowerCase().endsWith(".eml")) continue;
        const entryData = await entry.async("uint8array");
        collectedFiles.push({
          originalname: entry.name,
          raw: entryData,
        });
      }
    } else if (lowerName.endsWith(".eml")) {
      collectedFiles.push({
        originalname: file.name,
        raw: new Uint8Array(await file.arrayBuffer()),
      });
    }
  }

  return collectedFiles;
}

export async function buildExtractionWorkbook(files: FileList): Promise<Blob> {
  const emailFiles = await collectEmailFiles(files);
  if (emailFiles.length === 0) {
    throw new Error("No valid .eml files found in selected upload.");
  }

  const rows: ExtractionRow[] = [];
  for (let i = 0; i < emailFiles.length; i++) {
    const file = emailFiles[i];
    const parsedEmail = await PostalMime.parse(file.raw);
    rows.push(parseEmailHtml(parsedEmail.html || "", file.originalname, parsedEmail.date));

    // Yield to the browser event loop periodically for better UX on large batches.
    if (i > 0 && i % 10 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: WORKBOOK_COLUMNS });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Extractions");

  const workbookArray = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Blob([workbookArray], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
