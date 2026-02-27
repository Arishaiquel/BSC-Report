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

        // Submission Date
        const submissionDate = parsed.date ? parsed.date.toISOString().split('T')[0] : "";

        let buyAmount = "";
        let buyProduct = "";
        let rspAmount = "";
        let rspProduct = "";

        // Buy section
        const buyHeader = Array.from(doc.querySelectorAll("b, strong")).find(el => el.textContent?.trim().includes("Buy"));
        if (buyHeader) {
          const parent = buyHeader.parentElement;
          let current = parent?.nextElementSibling;
          
          while (current && current.tagName !== "P") {
            current = current.nextElementSibling;
          }
          
          if (current && current.tagName === "P") {
            buyProduct = current.textContent?.trim() || "";
            current = current.nextElementSibling;
          }
          
          while (current && current.tagName !== "TABLE") {
            current = current.nextElementSibling;
          }
          
          if (current && current.tagName === "TABLE") {
            const rows = current.querySelectorAll("tr");
            if (rows.length > 1) {
              const headers = Array.from(rows[0].querySelectorAll("th, td")).map(h => h.textContent?.trim() || "");
              const amountIdx = headers.findIndex(h => h.includes("Investment Amount"));
              
              let totalBuy = 0;
              for (let i = 1; i < rows.length; i++) {
                const cols = Array.from(rows[i].querySelectorAll("td")).map(c => c.textContent?.trim() || "");
                if (amountIdx !== -1 && cols[amountIdx]) {
                  const val = parseFloat(cols[amountIdx].replace(/[^0-9.]/g, ""));
                  if (!isNaN(val)) totalBuy += val;
                }
              }
              if (totalBuy > 0) buyAmount = `SGD ${totalBuy.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
            }
          }
        }

        // RSP section
        const rspHeader = Array.from(doc.querySelectorAll("b, strong")).find(el => el.textContent?.trim().includes("RSP Application"));
        if (rspHeader) {
          const parent = rspHeader.parentElement;
          let current = parent?.nextElementSibling;
          
          // Skip potentially empty text nodes/tags until we find the wording <p>
          while (current && current.tagName !== "P") {
            current = current.nextElementSibling;
          }
          
          if (current && current.tagName === "P") {
            rspProduct = current.textContent?.trim() || "";
            current = current.nextElementSibling;
          }
          
          // Look for the table
          while (current && current.tagName !== "TABLE") {
            current = current.nextElementSibling;
          }
          
          if (current && current.tagName === "TABLE") {
            const rows = current.querySelectorAll("tr");
            if (rows.length > 1) {
              const headers = Array.from(rows[0].querySelectorAll("th, td")).map(h => h.textContent?.trim() || "");
              const amountIdx = headers.findIndex(h => h.includes("RSP Amount"));
              
              let totalRsp = 0;
              for (let i = 1; i < rows.length; i++) {
                const cols = Array.from(rows[i].querySelectorAll("td")).map(c => c.textContent?.trim() || "");
                if (amountIdx !== -1 && cols[amountIdx]) {
                  const val = parseFloat(cols[amountIdx].replace(/[^0-9.]/g, ""));
                  if (!isNaN(val)) totalRsp += val;
                }
              }
              if (totalRsp > 0) rspAmount = `SGD ${totalRsp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
            }
          }
        }

        results.push({
          "Policy Number": policyNumber,
          "Submission Date": submissionDate,
          "Buy": buyAmount,
          "Buy product": buyProduct,
          "RSP Application": rspAmount,
          "RSP Application product": rspProduct
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
        { header: "RSP Application product", key: "RSP Application product" }
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
