import { sqliteTable, text, integer, real } from "drizzle-orm/sqlite-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const projects = sqliteTable("projects", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  name: text("name").notNull(),
  address: text("address").notNull(),
  client: text("client").notNull(),
  inspector: text("inspector").notNull(),
  afcReference: text("afc_reference").default(""),
  elevations: text("elevations").default("[]"), // JSON array of strings: selected elevation labels for this project
  createdAt: text("created_at").notNull(),
});

export const reports = sqliteTable("reports", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull(),
  inspectionNumber: text("inspection_number").default(""),
  inspectionDate: text("inspection_date").default(""),
  revision: text("revision").default("01"),
  locationsCovered: text("locations_covered").default(""),
  elevations: text("elevations").default("[]"), // JSON array of strings: elevation labels for this report
  attendees: text("attendees").default("[]"), // JSON array: [{name, company}]
  createdAt: text("created_at").notNull(),
});

export const defects = sqliteTable("defects", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull(),
  reportId: integer("report_id"), // nullable for backward compat
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
  updatedAt: text("updated_at"),
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
export const insertReportSchema = createInsertSchema(reports).omit({ id: true });
export const insertDefectSchema = createInsertSchema(defects).omit({ id: true });
export const insertPhotoSchema = createInsertSchema(photos).omit({ id: true });

// Types
export type InsertProject = z.infer<typeof insertProjectSchema>;
export type Project = typeof projects.$inferSelect;
export type InsertReport = z.infer<typeof insertReportSchema>;
export type Report = typeof reports.$inferSelect;
export type InsertDefect = z.infer<typeof insertDefectSchema>;
export type Defect = typeof defects.$inferSelect;
export type InsertPhoto = z.infer<typeof insertPhotoSchema>;
export type Photo = typeof photos.$inferSelect;

// Elevations (uploaded PDF/image drawings for annotation)
export const elevations = sqliteTable("elevations", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull(),
  name: text("name").notNull(),
  filename: text("filename").notNull(),
  fileType: text("file_type").notNull(), // 'image' or 'pdf'
  createdAt: text("created_at").notNull(),
});

// Markers (defect annotations on elevations)
export const markers = sqliteTable("markers", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  elevationId: integer("elevation_id").notNull(),
  defectId: integer("defect_id"), // nullable FK to defects for linked markers
  defectUid: text("defect_uid").notNull(), // display UID string
  status: text("status").notNull().default("open"),
  note: text("note"),
  xPercent: real("x_percent").notNull(),
  yPercent: real("y_percent").notNull(),
  createdAt: text("created_at").notNull(),
});

export const insertElevationSchema = createInsertSchema(elevations).omit({ id: true });
export const insertMarkerSchema = createInsertSchema(markers).omit({ id: true });

export type InsertElevation = z.infer<typeof insertElevationSchema>;
export type Elevation = typeof elevations.$inferSelect;
export type InsertMarker = z.infer<typeof insertMarkerSchema>;
export type Marker = typeof markers.$inferSelect;

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
