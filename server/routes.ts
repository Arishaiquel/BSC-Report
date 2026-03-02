import type { Express } from "express";
import { createServer, type Server } from "http";
import multer from "multer";
import { simpleParser } from "mailparser";
import { JSDOM } from "jsdom";
import ExcelJS from "exceljs";
import AdmZip from "adm-zip";

const upload = multer({ storage: multer.memoryStorage() });

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  
  app.post("/api/extract", upload.array("files"), async (req, res) => {
    try {
      const files = req.files as Express.Multer.File[];
      if (!files || files.length === 0) {
        return res.status(400).json({ message: "No files uploaded" });
      }

      const results = [];
      const allFiles = [];

      // Handle zip files if any
      for (const file of files) {
        if (file.mimetype === "application/zip" || file.originalname.endsWith(".zip")) {
          const zip = new AdmZip(file.buffer);
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
          allFiles.push(file);
        }
      }

      for (const file of allFiles) {
        const parsed = await simpleParser(file.buffer);
        const html = parsed.html || parsed.textAsHtml || "";
        const dom = new JSDOM(html);
        const doc = dom.window.document;

        // Policy Number extraction
        let policyNumber = "";
        const bodyText = doc.body.textContent || "";
        const policyMatch = bodyText.match(/\((P\d+)\)/);
        if (policyMatch) {
          policyNumber = policyMatch[1];
        }

        // Submission Date extraction from the email body if possible, fallback to parsed.date
        let submissionDate = "";
        const dateMatch = bodyText.match(/(\d{2}-\d{2}-\d{4})/);
        if (dateMatch) {
          submissionDate = dateMatch[1];
        } else {
          submissionDate = parsed.date ? parsed.date.toLocaleDateString('en-GB').replace(/\//g, '-') : "";
        }

        let buyAmount = 0;
        let buyForeign = "";
        let buyProducts: string[] = [];
        
        let rspAmount = 0;
        let rspForeign = "";
        let rspProducts: string[] = [];

        const sections = Array.from(doc.querySelectorAll("b, strong")).filter(el => {
          const text = el.textContent?.trim() || "";
          return text === "Buy" || text === "RSP Application" || text.includes("Buy") || text.includes("RSP Application");
        });

        for (const header of sections) {
          const isBuy = header.textContent?.trim().includes("Buy");
          const isRsp = header.textContent?.trim().includes("RSP Application");
          
          let current = header.parentElement?.nextElementSibling;
          
          while (current) {
            if (current.tagName === "P") {
              const pText = current.textContent?.trim() || "";
              if (pText && !["Buy", "RSP Application"].some(s => pText.includes(s))) {
                if (isBuy) buyProducts.push(pText);
                else if (isRsp) rspProducts.push(pText);
              }
            } else if (current.tagName === "TABLE") {
              const rows = current.querySelectorAll("tr");
              if (rows.length > 1) {
                const headers = Array.from(rows[0].querySelectorAll("th, td")).map(h => h.textContent?.trim() || "");
                const amountIdx = headers.findIndex(h => h.includes("Amount"));
                const currencyIdx = headers.findIndex(h => h.includes("Amount") || h.includes("Currency") || h.includes("Payment Mode"));
                
                for (let i = 1; i < rows.length; i++) {
                  const cols = Array.from(rows[i].querySelectorAll("td")).map(c => c.textContent?.trim() || "");
                  if (amountIdx !== -1 && cols[amountIdx]) {
                    const textValue = cols[amountIdx];
                    const val = parseFloat(textValue.replace(/[^0-9.]/g, ""));
                    if (!isNaN(val)) {
                      if (textValue.includes("SGD") || (!textValue.match(/[A-Z]{3}/) && !cols.some(c => c.match(/^(USD|EUR|GBP|AUD|JPY|HKD)$/)))) {
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
            } else if (current.querySelector("b, strong") && ["Buy", "RSP Application"].some(s => current?.textContent?.includes(s))) {
              break; // Hit next main section
            }
            current = current.nextElementSibling;
            if (current?.tagName === "P" && current.querySelector("b, strong") && ["Buy", "RSP Application"].some(s => current?.textContent?.includes(s))) break;
          }
        }

        results.push({
          "Policy Number": policyNumber,
          "Submission Date": submissionDate,
          "Buy": buyAmount > 0 ? `SGD ${buyAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : "",
          "Buy product": [...new Set(buyProducts)].join(", "),
          "RSP Application": rspAmount > 0 ? `SGD ${rspAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : "",
          "RSP Application product": [...new Set(rspProducts)].join(", "),
          "Foreign": [buyForeign, rspForeign].filter(Boolean).join("; ")
        });
      }

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Extractions");
      
      worksheet.columns = [
        { header: "Policy Number", key: "Policy Number" },
        { header: "Submission Date", key: "Submission Date" },
        { header: "Buy", key: "Buy" },
        { header: "Buy product", key: "Buy product" },
        { header: "RSP Application", key: "RSP Application" },
        { header: "RSP Application product", key: "RSP Application product" },
        { header: "Foreign", key: "Foreign" }
      ];

      worksheet.addRows(results);

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=extracted_data.xlsx");

      await workbook.xlsx.write(res);
      res.end();

    } catch (error) {
      console.error(error);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  return httpServer;
}
