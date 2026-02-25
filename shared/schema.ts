import { pgTable, text, serial, timestamp } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const extractions = pgTable("extractions", {
  id: serial("id").primaryKey(),
  filename: text("filename").notNull(),
  policyNumber: text("policy_number"),
  submissionDate: text("submission_date"),
  buyWording: text("buy_wording"),
  buyAmount: text("buy_amount"),
  rspWording: text("rsp_wording"),
  rspAmount: text("rsp_amount"),
  createdAt: timestamp("created_at").defaultNow(),
});

export const insertExtractionSchema = createInsertSchema(extractions).omit({ 
  id: true, 
  createdAt: true 
});

export type Extraction = typeof extractions.$inferSelect;
export type InsertExtraction = z.infer<typeof insertExtractionSchema>;
