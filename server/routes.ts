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

        // Date extraction from filename (format: ..._YYYYMMDD_...)
        let submissionDate = "";
        const filename = file.originalname || "";
        const filenameDateMatch = filename.match(/_(\d{4})(\d{2})(\d{2})_/);
        if (filenameDateMatch) {
          const [_, year, month, day] = filenameDateMatch;
          submissionDate = `${parseInt(day)}/${parseInt(month)}/${year}`;
        } else {
          submissionDate = parsed.date ? parsed.date.toLocaleDateString('en-GB') : "";
        }

        let buyAmount = 0;
        let buyForeign = "";
        let buyProducts: string[] = [];
        
        let rspAmount = 0;
        let rspForeign = "";
        let rspProducts: string[] = [];

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
                  ["Switch", "Sell", "Rebalance", "RSP Amendment", "Consolidated Dividend"].some(s => innerText.includes(s))) {
                break;
              }
              
              // If we hit a DIFFERENT valid section (e.g. processing Buy and we hit RSP Application), 
              // we also stop because that section will be handled by its own iteration of the outer loop.
              if (innerText !== headerText && (innerText === "Buy" || innerText === "RSP Application")) {
                break;
              }
            }
            
            // Check for the specific underlined headers as well
            if (current.tagName === "P" && current.style.textDecoration === "underline") {
              const pText = current.textContent?.trim() || "";
              if (pText !== headerText && ["Switch", "Sell", "Rebalance", "RSP Amendment"].some(s => pText.includes(s))) {
                break;
              }
            }

            if (current.tagName === "P") {
              const pText = currentText;
              if (pText && allowedProducts.some(ap => pText.includes(ap))) {
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
          "Buy product": [...new Set(buyProducts)].join(", "),
          "RSP Application": rspAmount > 0 ? `SGD ${rspAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : "",
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
        { header: "Buy product", key: "Buy product" },
        { header: "RSP Application", key: "RSP Application" },
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
    }
  });

  return httpServer;
}
