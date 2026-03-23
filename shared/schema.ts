import { sqliteTable, text, integer } from "drizzle-orm/sqlite-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const projects = sqliteTable("projects", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  name: text("name").notNull(),
  address: text("address").notNull(),
  client: text("client").notNull(),
  inspector: text("inspector").notNull(),
  afcReference: text("afc_reference").default(""),
  revision: text("revision").default("01"),

  inspectionNumber: text("inspection_number").default(""),
  inspectionDate: text("inspection_date").default(""),
  locationsCovered: text("locations_covered").default(""),
  elevations: text("elevations").default("[]"), // JSON array of strings: selected elevation labels for this project
  attendees: text("attendees").default("[]"), // JSON array: [{name, company}]
  createdAt: text("created_at").notNull(),
});

export const defects = sqliteTable("defects", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull(),
  uid: text("uid").notNull(),
  dateOpened: text("date_opened").notNull(),
  dateClosed: text("date_closed"),
  comment: text("comment").notNull(),
  actionRequired: text("action_required").notNull(),
  assignedTo: text("assigned_to").default(""),
  dueDate: text("due_date").default(""),
  verificationMethod: text("verification_method").notNull(),
  verificationPerson: text("verification_person").notNull(),
  status: text("status").notNull().default("open"), // open, complete
  recordType: text("record_type").notNull().default("defect"), // defect, observation
});

export const photos = sqliteTable("photos", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull(),
  filename: text("filename").notNull(),
  caption: text("caption"),
  slot: text("slot").notNull().default("wip1"), // wip1, wip2, wip3, wip4, wip5, complete
  createdAt: text("created_at").notNull(),
});

// Insert schemas
export const insertProjectSchema = createInsertSchema(projects).omit({ id: true });
export const insertDefectSchema = createInsertSchema(defects).omit({ id: true });
export const insertPhotoSchema = createInsertSchema(photos).omit({ id: true });

// Types
export type InsertProject = z.infer<typeof insertProjectSchema>;
export type Project = typeof projects.$inferSelect;
export type InsertDefect = z.infer<typeof insertDefectSchema>;
export type Defect = typeof defects.$inferSelect;
export type InsertPhoto = z.infer<typeof insertPhotoSchema>;
export type Photo = typeof photos.$inferSelect;

// Users (kept from template)
export const users = sqliteTable("users", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  username: text("username").notNull().unique(),
  password: text("password").notNull(),
});

export const insertUserSchema = createInsertSchema(users).pick({
  username: true,
  password: true,
});

export type InsertUser = z.infer<typeof insertUserSchema>;
export type User = typeof users.$inferSelect;
