import * as MsgReaderModule from "@kenjiuno/msgreader";
import JSZip from "jszip";
import * as XLSX from "xlsx";

type LocalMsgFile = {
  originalname: string;
  raw: Uint8Array;
};

type MsgReaderCtor = new (data: ArrayBuffer) => {
  getFileData: () => {
    body?: string;
    bodyHtml?: string;
    bodyHTML?: string;
    messageDeliveryTime?: string;
    clientSubmitTime?: string;
  };
};

function resolveMsgReaderCtor(moduleValue: unknown): MsgReaderCtor {
  let current: unknown = moduleValue;
  for (let depth = 0; depth < 4; depth++) {
    if (typeof current === "function") {
      return current as MsgReaderCtor;
    }
    if (
      current &&
      typeof current === "object" &&
      "default" in (current as Record<string, unknown>)
    ) {
      current = (current as Record<string, unknown>).default;
      continue;
    }
    break;
  }
  throw new Error("Unable to initialize MSG parser.");
}

const MsgReader = resolveMsgReaderCtor(MsgReaderModule);

type FameExtractionRow = {
  "Policy Number": string;
  "Submission Date": string;
  Buy: string;
  RSP: string;
  "advisor name": string;
  "Buy product type": string;
  "RSP product type": string;
  "Foreign Buy": string;
  "Foreign RSP": string;
};

const WORKBOOK_COLUMNS: Array<keyof FameExtractionRow> = [
  "Policy Number",
  "Submission Date",
  "Buy",
  "RSP",
  "advisor name",
  "Buy product type",
  "RSP product type",
  "Foreign Buy",
  "Foreign RSP",
];

function formatAmount(value: number): string {
  if (value <= 0) return "";
  return value.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function normalizeCell(value: string): string {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function splitTableLine(line: string): string[] {
  const cells = line.split("\t").map((cell) => normalizeCell(cell));
  while (cells.length > 0 && cells[cells.length - 1] === "") {
    cells.pop();
  }
  return cells;
}

function normalizeHeader(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function parseGrossAmount(value: string): number {
  const normalized = value.replace(/,/g, "");
  const amount = Number.parseFloat(normalized.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(amount) ? amount : Number.NaN;
}

function parseTextDate(input: string): string {
  const match = input.trim().match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})$/);
  if (!match) return "";

  const [, dayRaw, monthRaw, year] = match;
  const monthMap: Record<string, number> = {
    jan: 1,
    feb: 2,
    mar: 3,
    apr: 4,
    may: 5,
    jun: 6,
    jul: 7,
    aug: 8,
    sep: 9,
    oct: 10,
    nov: 11,
    dec: 12,
  };
  const month = monthMap[monthRaw.toLowerCase()];
  if (!month) return "";
  return `${parseInt(dayRaw, 10)}/${month}/${year}`;
}

function parseSubmissionDate(bodyText: string, messageDate?: string): string {
  const bodyMatch = bodyText.match(
    /online transaction\s+[A-Za-z0-9]+\s+on\s+(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})/i,
  );
  if (bodyMatch) {
    const parsed = parseTextDate(bodyMatch[1]);
    if (parsed) return parsed;
  }

  if (!messageDate) return "";
  const parsedDate = new Date(messageDate);
  if (Number.isNaN(parsedDate.getTime())) return "";
  return parsedDate.toLocaleDateString("en-GB");
}

function formatForeignByCurrency(entries: Map<string, number>): string {
  if (entries.size === 0) return "";
  return Array.from(entries.entries())
    .map(([currency, amount]) => `${currency} ${formatAmount(amount)}`)
    .join(", ");
}

function parseMessageText(bodyText: string, messageDate?: string): FameExtractionRow {
  const policyNumber = normalizeCell(
    bodyText.match(/Your Account Number:\s*(\d+)/i)?.[1] || "",
  );

  const advisorName = normalizeCell(
    bodyText.match(/FA representative,\s*([^\r\n.]+)/i)?.[1] || "",
  );

  const lines = bodyText.split(/\r?\n/);
  const headerLineIndex = lines.findIndex((line) =>
    /Fund\s*\t\s*Transaction Type\s*\t\s*Fund Source\s*\t\s*Currency\s*\t\s*Gross Amt/i.test(
      line,
    ),
  );

  let buyAmount = 0;
  let rspAmount = 0;
  const buyProducts = new Set<string>();
  const rspProducts = new Set<string>();
  const foreignBuy = new Map<string, number>();
  const foreignRsp = new Map<string, number>();

  if (headerLineIndex !== -1) {
    const headers = splitTableLine(lines[headerLineIndex]);
    const headerIndexMap = new Map<string, number>();
    headers.forEach((header, idx) => {
      headerIndexMap.set(normalizeHeader(header), idx);
    });

    const fundIdx = headerIndexMap.get("fund");
    const typeIdx = headerIndexMap.get("transactiontype");
    const currencyIdx = headerIndexMap.get("currency");
    const grossIdx = headerIndexMap.get("grossamt");

    if (
      fundIdx !== undefined &&
      typeIdx !== undefined &&
      currencyIdx !== undefined &&
      grossIdx !== undefined
    ) {
      for (let i = headerLineIndex + 1; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();
        if (!trimmed) continue;
        if (/^#\s*Dividend option/i.test(trimmed)) break;
        if (/^To view the portfolio/i.test(trimmed)) break;
        if (/^For more information/i.test(trimmed)) break;
        if (!line.includes("\t")) continue;

        const cells = splitTableLine(line);
        if (
          fundIdx >= cells.length ||
          typeIdx >= cells.length ||
          currencyIdx >= cells.length ||
          grossIdx >= cells.length
        ) {
          continue;
        }

        const transactionType = cells[typeIdx].toLowerCase();
        if (transactionType !== "subscription" && transactionType !== "rsp") {
          continue;
        }

        const fund = cells[fundIdx];
        const currency = cells[currencyIdx].toUpperCase();
        const amount = parseGrossAmount(cells[grossIdx]);
        if (Number.isNaN(amount)) continue;

        const isSubscription = transactionType === "subscription";
        const isSgd = currency === "SGD";

        if (isSubscription) {
          if (fund) buyProducts.add(fund);
          if (isSgd) {
            buyAmount += amount;
          } else if (currency) {
            foreignBuy.set(currency, (foreignBuy.get(currency) || 0) + amount);
          }
        } else {
          if (fund) rspProducts.add(fund);
          if (isSgd) {
            rspAmount += amount;
          } else if (currency) {
            foreignRsp.set(currency, (foreignRsp.get(currency) || 0) + amount);
          }
        }
      }
    }
  }

  return {
    "Policy Number": policyNumber,
    "Submission Date": parseSubmissionDate(bodyText, messageDate),
    Buy: formatAmount(buyAmount),
    RSP: formatAmount(rspAmount),
    "advisor name": advisorName,
    "Buy product type": Array.from(buyProducts).join(", "),
    "RSP product type": Array.from(rspProducts).join(", "),
    "Foreign Buy": formatForeignByCurrency(foreignBuy),
    "Foreign RSP": formatForeignByCurrency(foreignRsp),
  };
}

function parseMessageHtml(
  htmlBody: string,
  messageDate?: string,
  fallbackBodyText = "",
): FameExtractionRow {
  const doc = new DOMParser().parseFromString(htmlBody || "", "text/html");
  const docText = doc.body.textContent || "";
  const combinedText = `${docText}\n${fallbackBodyText}`.trim();

  const policyNumber = normalizeCell(
    combinedText.match(/Your Account Number:\s*(\d+)/i)?.[1] || "",
  );

  const advisorName = normalizeCell(
    combinedText.match(/FA representative,\s*([^\r\n.]+)/i)?.[1] || "",
  );

  let buyAmount = 0;
  let rspAmount = 0;
  const buyProducts = new Set<string>();
  const rspProducts = new Set<string>();
  const foreignBuy = new Map<string, number>();
  const foreignRsp = new Map<string, number>();

  const tables = Array.from(doc.querySelectorAll("table"));

  for (const table of tables) {
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length < 2) continue;

    let headerRowIndex = -1;
    let fundIdx = -1;
    let typeIdx = -1;
    let currencyIdx = -1;
    let grossIdx = -1;

    for (let i = 0; i < rows.length; i++) {
      const headers = Array.from(rows[i].querySelectorAll("th, td")).map((cell) =>
        normalizeHeader(normalizeCell(cell.textContent || "")),
      );

      if (headers.length === 0) continue;

      fundIdx = headers.findIndex((header) => header === "fund");
      typeIdx = headers.findIndex((header) => header === "transactiontype");
      currencyIdx = headers.findIndex((header) => header === "currency");
      grossIdx = headers.findIndex((header) => header === "grossamt");

      if (fundIdx !== -1 && typeIdx !== -1 && currencyIdx !== -1 && grossIdx !== -1) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) continue;

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const cells = Array.from(rows[i].querySelectorAll("td")).map((cell) =>
        normalizeCell(cell.textContent || ""),
      );
      if (cells.length === 0) continue;

      if (
        fundIdx >= cells.length ||
        typeIdx >= cells.length ||
        currencyIdx >= cells.length ||
        grossIdx >= cells.length
      ) {
        continue;
      }

      const transactionType = cells[typeIdx].toLowerCase();
      if (transactionType !== "subscription" && transactionType !== "rsp") {
        continue;
      }

      const fund = cells[fundIdx];
      const currency = cells[currencyIdx].toUpperCase();
      const amount = parseGrossAmount(cells[grossIdx]);
      if (Number.isNaN(amount)) continue;

      const isSubscription = transactionType === "subscription";
      const isSgd = currency === "SGD";

      if (isSubscription) {
        if (fund) buyProducts.add(fund);
        if (isSgd) {
          buyAmount += amount;
        } else if (currency) {
          foreignBuy.set(currency, (foreignBuy.get(currency) || 0) + amount);
        }
      } else {
        if (fund) rspProducts.add(fund);
        if (isSgd) {
          rspAmount += amount;
        } else if (currency) {
          foreignRsp.set(currency, (foreignRsp.get(currency) || 0) + amount);
        }
      }
    }
  }

  return {
    "Policy Number": policyNumber,
    "Submission Date": parseSubmissionDate(combinedText, messageDate),
    Buy: formatAmount(buyAmount),
    RSP: formatAmount(rspAmount),
    "advisor name": advisorName,
    "Buy product type": Array.from(buyProducts).join(", "),
    "RSP product type": Array.from(rspProducts).join(", "),
    "Foreign Buy": formatForeignByCurrency(foreignBuy),
    "Foreign RSP": formatForeignByCurrency(foreignRsp),
  };
}

async function collectMsgFiles(files: FileList): Promise<LocalMsgFile[]> {
  const collectedFiles: LocalMsgFile[] = [];

  for (const file of Array.from(files)) {
    const lowerName = file.name.toLowerCase();

    if (lowerName.endsWith(".zip")) {
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      const zipEntries = Object.values(zip.files);

      for (const entry of zipEntries) {
        if (entry.dir || !entry.name.toLowerCase().endsWith(".msg")) continue;
        collectedFiles.push({
          originalname: entry.name,
          raw: await entry.async("uint8array"),
        });
      }
      continue;
    }

    if (lowerName.endsWith(".msg")) {
      collectedFiles.push({
        originalname: file.name,
        raw: new Uint8Array(await file.arrayBuffer()),
      });
    }
  }

  return collectedFiles;
}

export async function buildFameExtractionWorkbook(files: FileList): Promise<Blob> {
  const msgFiles = await collectMsgFiles(files);
  if (msgFiles.length === 0) {
    throw new Error("No valid .msg files found in selected upload.");
  }

  const rows: FameExtractionRow[] = [];
  for (let i = 0; i < msgFiles.length; i++) {
    const file = msgFiles[i];
    const buffer = file.raw.buffer.slice(file.raw.byteOffset, file.raw.byteOffset + file.raw.byteLength);
    const parsed = new MsgReader(buffer).getFileData();
    const htmlBody = ((parsed as { bodyHtml?: string; bodyHTML?: string }).bodyHtml ||
      (parsed as { bodyHtml?: string; bodyHTML?: string }).bodyHTML ||
      "") as string;
    const bodyText = (parsed.body || "") as string;
    const messageDate = parsed.messageDeliveryTime || parsed.clientSubmitTime || "";
    const bodyLooksLikeHtml = /<(html|body|table|tr|td|th|p|div)\b/i.test(bodyText);
    const htmlCandidate = htmlBody || (bodyLooksLikeHtml ? bodyText : "");

    if (htmlCandidate) {
      const htmlRow = parseMessageHtml(htmlCandidate, messageDate, bodyText);
      const hasNoExtractedAmounts =
        !htmlRow.Buy && !htmlRow.RSP && !htmlRow["Foreign Buy"] && !htmlRow["Foreign RSP"];
      rows.push(hasNoExtractedAmounts ? parseMessageText(bodyText, messageDate) : htmlRow);
    } else {
      rows.push(parseMessageText(bodyText, messageDate));
    }

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
