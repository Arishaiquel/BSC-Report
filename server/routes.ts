import type { Express, NextFunction, Request, Response } from "express";
import { createServer, type Server } from "http";
import multer from "multer";
import { simpleParser } from "mailparser";
import { JSDOM } from "jsdom";
import ExcelJS from "exceljs";
import AdmZip from "adm-zip";
import fs from "fs/promises";
import os from "os";
import path from "path";

type ParsedFile = {
  buffer: Buffer;
  originalname: string;
};

const MAX_UPLOAD_FILE_SIZE_BYTES = 550 * 1024 * 1024;
const MAX_UPLOAD_FILES = 20;
const uploadDir = path.join(os.tmpdir(), "bsc-report-uploads");

const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, uploadDir),
    filename: (_req, file, cb) => {
      const safeExt = path.extname(file.originalname || "").toLowerCase();
      const uniqueName = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}${safeExt}`;
      cb(null, uniqueName);
    },
  }),
  limits: {
    fileSize: MAX_UPLOAD_FILE_SIZE_BYTES,
    files: MAX_UPLOAD_FILES,
  },
});

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

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  await fs.mkdir(uploadDir, { recursive: true });

  const uploadFiles = (req: Request, res: Response, next: NextFunction) => {
    upload.array("files")(req, res, (err) => {
      if (!err) {
        next();
        return;
      }

      if (err instanceof multer.MulterError) {
        if (err.code === "LIMIT_FILE_SIZE") {
          res.status(413).json({
            message:
              "A file is too large. Keep each file under 550MB and retry.",
          });
          return;
        }

        if (err.code === "LIMIT_FILE_COUNT") {
          res.status(413).json({
            message: `Too many files. Maximum is ${MAX_UPLOAD_FILES} files per request.`,
          });
          return;
        }
      }

      next(err);
    });
  };

  app.get("/api/health", (_req, res) => {
    res.json({ ok: true });
  });

  app.post("/api/extract", uploadFiles, async (req, res) => {
    const uploadedFiles = (req.files as Express.Multer.File[]) || [];
    try {
      if (!uploadedFiles.length) {
        return res.status(400).json({ message: "No files uploaded" });
      }

      const results: Array<Record<string, string>> = [];
      const allFiles: ParsedFile[] = [];

      // Handle zip files if any
      for (const file of uploadedFiles) {
        if (file.mimetype === "application/zip" || file.originalname.endsWith(".zip")) {
          const zip = new AdmZip(file.path);
          const zipEntries = zip.getEntries();
          for (const entry of zipEntries) {
            if (entry.entryName.endsWith(".eml") && !entry.isDirectory) {
              allFiles.push({
                buffer: entry.getData(),
                originalname: entry.entryName
              });
            }
          }
        } else if (file.originalname.endsWith(".eml")) {
          allFiles.push({
            buffer: await fs.readFile(file.path),
            originalname: file.originalname,
          });
        }
      }

      if (!allFiles.length) {
        return res.status(400).json({ message: "No valid .eml files found in upload." });
      }

      for (const file of allFiles) {
        const parsed = await simpleParser(file.buffer);
        const html = parsed.html || parsed.textAsHtml || "";
        const dom = new JSDOM(String(html));
        const doc = dom.window.document;
        const advisorName = extractAdvisorName(doc);

        // Policy Number extraction
        let policyNumber = "";
        const bodyText = doc.body.textContent || "";
        const policyMatch = bodyText.match(/\((P\d+)\)/);
        if (policyMatch) {
          policyNumber = policyMatch[1];
        }

        // Date extraction from filename (format: ..._YYYYMMDD_...)
        let submissionDate = "";
        const filename = file.originalname || "";
        const filenameDateMatch = filename.match(/_(\d{4})(\d{2})(\d{2})_/);
        if (filenameDateMatch) {
          const [_, year, month, day] = filenameDateMatch;
          submissionDate = `${parseInt(day, 10)}/${parseInt(month, 10)}/${year}`;
        } else {
          submissionDate = parsed.date ? parsed.date.toLocaleDateString('en-GB') : "";
        }

        let buyAmount = 0;
        let buyForeign = "";
        const buyProducts: string[] = [];
        
        let rspAmount = 0;
        let rspForeign = "";
        const rspProducts: string[] = [];

        const allowedProducts = ["Company Portfolio", "DPMS", "Unit Trust"];

        const sections = Array.from(doc.querySelectorAll("b, strong")).filter(el => {
          const text = el.textContent?.trim() || "";
          return text === "Buy" || text === "RSP Application" || text.includes("Buy") || text.includes("RSP Application");
        });

        for (const header of sections) {
          const headerText = header.textContent?.trim() || "";
          // Strict matching for Buy and RSP Application only
          const isBuy = headerText === "Buy";
          const isRsp = headerText === "RSP Application";
          
          if (!isBuy && !isRsp) continue;

          let current = header.parentElement?.nextElementSibling;
          
          while (current) {
            const currentText = current.textContent?.trim() || "";
            
            // If we hit another main header, check if it's one we should skip
            if (current.querySelector("b, strong")) {
              const innerText = current.querySelector("b, strong")?.textContent?.trim() || "";
              
              // If it's a different section that is NOT Buy or RSP Application, we stop this section's processing.
              // This handles cases like Buy -> Switch -> RSP Application correctly.
              if (innerText !== headerText && 
                  ["Switch", "Sell", "Rebalance", "RSP Amendment", "Consolidated Dividend", "ETF"].some(s => innerText.includes(s))) {
                break;
              }
              
              // If we hit a DIFFERENT valid section (e.g. processing Buy and we hit RSP Application), 
              // we also stop because that section will be handled by its own iteration of the outer loop.
              if (innerText !== headerText && (innerText === "Buy" || innerText === "RSP Application")) {
                break;
              }
            }
            
            // Check for the specific underlined headers as well
            if (
              current.tagName === "P" &&
              current instanceof dom.window.HTMLElement &&
              current.style.textDecoration === "underline"
            ) {
              const pText = current.textContent?.trim() || "";
              if (pText !== headerText && ["Switch", "Sell", "Rebalance", "RSP Amendment", "ETF"].some(s => pText.includes(s))) {
                break;
              }
            }

            if (current.tagName === "P") {
              const pText = currentText;
              // Strict check to exclude ETF/Exchange Traded Fund sections
              if (pText && (pText.includes("ETF") || pText.includes("Exchange Traded Fund"))) {
                // If this is an ETF paragraph, we should skip the next table entirely
                let next = current.nextElementSibling;
                while (next && next.tagName !== "P" && next.tagName !== "TABLE" && !next.querySelector("b, strong")) {
                  next = next.nextElementSibling;
                }
                if (next && next.tagName === "TABLE") {
                  current = next; // Skip processing this table
                }
              } else if (pText && allowedProducts.some(ap => pText.includes(ap))) {
                const matchedProduct = allowedProducts.find(ap => pText.includes(ap));
                if (matchedProduct) {
                  if (isBuy) buyProducts.push(matchedProduct);
                  else if (isRsp) rspProducts.push(matchedProduct);
                }
              }
            } else if (current.tagName === "TABLE") {
              // Only process table if it's preceded by an allowed product or we're in a valid sub-section
              const rows = current.querySelectorAll("tr");
              if (rows.length > 1) {
                const headers = Array.from(rows[0].querySelectorAll("th, td")).map(h => h.textContent?.trim() || "");
                const amountIdx = headers.findIndex(h => h.includes("Amount"));
                
                for (let i = 1; i < rows.length; i++) {
                  const cols = Array.from(rows[i].querySelectorAll("td")).map(c => c.textContent?.trim() || "");
                  if (amountIdx !== -1 && cols[amountIdx]) {
                    const textValue = cols[amountIdx];
                    const val = parseFloat(textValue.replace(/[^0-9.]/g, ""));
                    if (!isNaN(val)) {
                      const isForeign = !textValue.includes("SGD") && (textValue.match(/[A-Z]{3}/) || cols.some(c => c.match(/^(USD|EUR|GBP|AUD|JPY|HKD)$/)));
                      
                      if (!isForeign) {
                        if (isBuy) buyAmount += val;
                        else if (isRsp) rspAmount += val;
                      } else {
                        const currency = textValue.match(/[A-Z]{3}/)?.[0] || cols.find(c => c.match(/^[A-Z]{3}$/)) || "Foreign";
                        const foreignEntry = `${currency} ${val.toFixed(2)}`;
                        if (isBuy) buyForeign = buyForeign ? `${buyForeign}, ${foreignEntry}` : foreignEntry;
                        else if (isRsp) rspForeign = rspForeign ? `${rspForeign}, ${foreignEntry}` : foreignEntry;
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
          "Buy": buyAmount > 0 ? `SGD ${buyAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : "",
          "RSP Application": rspAmount > 0 ? `SGD ${rspAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : "",
          "advisor name": advisorName,
          "Buy product": [...new Set(buyProducts)].join(", "),
          "RSP Application product": [...new Set(rspProducts)].join(", "),
          "Foreign Buy": buyForeign,
          "Foreign RSP Application": rspForeign
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
        { header: "Foreign RSP Application", key: "Foreign RSP Application" }
      ];

      worksheet.addRows(results);

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=extracted_data.xlsx");

      await workbook.xlsx.write(res);
      res.end();

    } catch (error) {
      console.error(error);
      res.status(500).json({ message: "Internal server error" });
    } finally {
      await Promise.allSettled(
        uploadedFiles
          .map((file) => file.path)
          .filter((filePath): filePath is string => typeof filePath === "string")
          .map((filePath) => fs.unlink(filePath)),
      );
    }
  });

  return httpServer;
}
