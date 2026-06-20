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
  // §2.3 Title page + §1 spec: separate Report Title field (full-width, no truncation).
  // Replaces the legacy combined "Report Title / Address" usage of `name` on the cover.
  reportTitle: text("report_title").default(""),
  // §1.1 Roles table — JSON array [{role, entity, contactDetails}].
  // contactDetails is a free-text multi-line string so a single cell can hold
  // "Lana Angel\nM | 0407 759 590\nE | lana@angelfacade.consulting" verbatim.
  roles: text("roles").default("[]"),
  // §1.2 Scope of works — JSON array [{areaRef, location, workItem, accessMethod}].
  scopeOfWorks: text("scope_of_works").default("[]"),
  // §1.4 Background information — JSON array of reference docs.
  // [{type: "drawing"|"specification"|"manual"|"report"|"other",
  //   originator, title, docNumbers?, revision?, date}]
  backgroundDocs: text("background_docs").default("[]"),
  // §1.5.1 Area Ref template (NEW projects only). Free-form pattern of UID part
  // placeholders + literal separators, e.g. "{elevation}{drop}-{level}" → "E4-7".
  // Placeholders: {elevation} {drop} {level} — substituted with each defect's codes.
  // Final UID becomes: <AreaRef>-<workItem>-<seq#>. Legacy projects leave this
  // empty/null and continue to render the original 5-part UID assembly.
  areaRefTemplate: text("area_ref_template").default(""),
  elevations: text("elevations").default("[]"), // JSON array of strings: selected elevation labels for this project
  customDrops: text("custom_drops").default("[]"), // JSON array of strings
  customLevels: text("custom_levels").default("[]"), // JSON array of strings
  customWorkTypes: text("custom_work_types").default("[]"), // JSON array of {code, label}
  enabledUidParts: text("enabled_uid_parts").default('{"elevation":true,"drop":true,"level":true,"workType":true}'),
  primaryWorkTypes: text("primary_work_types").default("[]"), // JSON array of codes
  // Ordered list of this project's location dimensions, JSON-encoded.
  // e.g. ["elevation","drop","level"] (East Elevation) or ["stage","level"] (waterproofing).
  // Drives the single formatLocation() helper so the register and card never disagree.
  locationDimensions: text("location_dimensions").default('["elevation","drop","level"]'),
  // SVR reformat (Stage 2) — when true, hide the "(prev. {legacy_id})" alias on defect
  // cards/register. User flips this on after cycling through the next inspection.
  hideLegacyAliases: integer("hide_legacy_aliases", { mode: "boolean" }).default(false),
  // Inspection-workflow (D3): when true, the export renders a "completed register" tail
  // section. Wire-only this round — rendering deferred (no UI checkbox yet).
  showCompletedRegister: integer("show_completed_register", { mode: "boolean" }).default(false),
  // Export-profiles (Pass 1): follow-up action categories, JSON [{code,label,isDefault?}].
  categories: text("categories").default("[]"),
  // Per-profile export config (contractor/client): filenameSuffix + ordered categoryTreatments.
  exportProfiles: text("export_profiles").default('{"contractor":{"filenameSuffix":"Contractor","categoryTreatments":[{"code":"RR","treatment":"itemise"},{"code":"WIP","treatment":"itemise"},{"code":"PI","treatment":"itemise"},{"code":"RD","treatment":"itemise"},{"code":"PN","treatment":"summarise"}]},"client":{"filenameSuffix":"Client","categoryTreatments":[{"code":"RD","treatment":"itemise"},{"code":"WIP","treatment":"itemise"},{"code":"PN","treatment":"itemise"},{"code":"PI","treatment":"itemise"},{"code":"RR","treatment":"summarise"}]}}'),
  createdAt: text("created_at").notNull(),
});

export const globalSettings = sqliteTable("global_settings", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  workTypes: text("work_types").default("[]"), // JSON array of {code, label}
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
  // §2.3 spec — frozen snapshot of project-setup data captured at report creation.
  // Editing the project later does NOT mutate existing reports. Snapshot is preserved
  // across revisions of the same SVR. Renderers prefer this over the live project row.
  // JSON shape: {name, address, reportTitle, client, inspector, afcReference,
  //              roles, scopeOfWorks, backgroundDocs, areaRefTemplate}.
  projectSnapshot: text("project_snapshot"),
  priorReportId: integer("prior_report_id"), // report this one was cloned from (Start Next Inspection). NULL for the first report.
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
  status: text("status").notNull().default("open"), // stored: open, complete, archived (displayStatus computes Open|Amended|Closed|Archived)
  recordType: text("record_type").notNull().default("defect"), // defect, observation
  // Structured UID parts — source of truth for the form so it never re-parses the assembled uid string.
  elevationCode: text("elevation_code"),
  dropCode: text("drop_code"),
  levelCode: text("level_code"),
  workTypeCode: text("work_type_code"),
  seqNumber: text("seq_number"),
  // Export-profiles (Pass 1): audience tag + follow-up category code.
  audience: text("audience").default("both"), // "both" | "contractor" | "client"
  categoryCode: text("category_code"), // references one of project.categories[].code; nullable -> "(uncategorised)"
  // §2.2 — AI-generated Action List summary.
  // actionSummary       — the cached one-sentence imperative used in the Action List table.
  // actionSummarySource — "ai" | "fallback" | "manual" (set by the regenerate endpoint or by manual edit).
  // actionSummaryInputHash — SHA-256 hex of the normalised prompt input (observation + actionRequired +
  //                          category + workType). Hash inputs MUST stay identical to prompt inputs;
  //                          a mismatch against the live data marks the summary stale.
  actionSummary: text("action_summary"),
  actionSummarySource: text("action_summary_source"),
  actionSummaryInputHash: text("action_summary_input_hash"),
  // SVR reformat (Stage A) additions:
  legacyId: text("legacy_id"), // NULL until Stage 2 (apply) populates it — DO NOT backfill in Stage 1
  locationStructured: text("location_structured"), // JSON: {elevation,drop,level} etc. Source of truth for location string.
  inspectionOpened: integer("inspection_opened"), // inspection_number (as INTEGER) where this record FIRST appeared
  updatedAt: text("updated_at"),
  createdAt: text("created_at"),
});

export const photos = sqliteTable("photos", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull(),
  reportId: integer("report_id"), // nullable for backward compat; set on upload
  originReportId: integer("origin_report_id"), // report where this photo FIRST appeared (survives carry-over cloning)
  filename: text("filename").notNull(),
  caption: text("caption"),
  slot: text("slot").notNull().default("wip1"), // wip<N> for any positive integer N, or 'complete'
  newOverride: text("new_override"), // "new" | "not-new" | null (auto-detect via originReportId)
  captureDate: text("capture_date"), // when the photo was taken (EXIF/manual); falls back to createdAt for display
  createdAt: text("created_at").notNull(),
});

export const shareLinks = sqliteTable("share_links", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  token: text("token").notNull().unique(),
  reportId: integer("report_id").notNull(),
  recipientName: text("recipient_name").default(""),
  createdAt: text("created_at").notNull(),
});

// Insert schemas
export const insertProjectSchema = createInsertSchema(projects).omit({ id: true });
export const insertReportSchema = createInsertSchema(reports).omit({ id: true });
export const insertDefectSchema = createInsertSchema(defects).omit({ id: true });
export const insertPhotoSchema = createInsertSchema(photos).omit({ id: true });
export const insertShareLinkSchema = createInsertSchema(shareLinks).omit({ id: true });

// Types
export type InsertProject = z.infer<typeof insertProjectSchema>;
export type Project = typeof projects.$inferSelect;
export type InsertReport = z.infer<typeof insertReportSchema>;
export type Report = typeof reports.$inferSelect;
export type InsertDefect = z.infer<typeof insertDefectSchema>;
export type Defect = typeof defects.$inferSelect;
export type InsertPhoto = z.infer<typeof insertPhotoSchema>;
export type Photo = typeof photos.$inferSelect;
export type ShareLink = typeof shareLinks.$inferSelect;
export type InsertShareLink = z.infer<typeof insertShareLinkSchema>;

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

// Observation history (tracks prior observation text across inspections)
export const observationHistory = sqliteTable("observation_history", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull().references(() => defects.id, { onDelete: "cascade" }),
  reportId: integer("report_id").notNull().references(() => reports.id),
  text: text("text").notNull(),
  createdAt: text("created_at").notNull(),
});

// Action history (tracks prior action required text across inspections)
export const actionHistory = sqliteTable("action_history", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull().references(() => defects.id, { onDelete: "cascade" }),
  reportId: integer("report_id").notNull().references(() => reports.id),
  text: text("text").notNull(),
  createdAt: text("created_at").notNull(),
});

export const insertObservationHistorySchema = createInsertSchema(observationHistory).omit({ id: true });
export const insertActionHistorySchema = createInsertSchema(actionHistory).omit({ id: true });

export type InsertObservationHistory = z.infer<typeof insertObservationHistorySchema>;
export type ObservationHistory = typeof observationHistory.$inferSelect;
export type InsertActionHistory = z.infer<typeof insertActionHistorySchema>;
export type ActionHistory = typeof actionHistory.$inferSelect;

// Defect Locations (additional locations for a defect)
export const defectLocations = sqliteTable("defect_locations", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull().references(() => defects.id, { onDelete: "cascade" }),
  uid: text("uid").default(""),
  drop: text("drop").default(""),
  elevation: text("elevation").default(""),
  level: text("level").default(""),
  description: text("description").default(""),
  displayOrder: integer("display_order").default(0),
  createdAt: text("created_at").notNull(),
  updatedAt: text("updated_at"),
});

export const insertDefectLocationSchema = createInsertSchema(defectLocations).omit({ id: true });
export type InsertDefectLocation = z.infer<typeof insertDefectLocationSchema>;
export type DefectLocation = typeof defectLocations.$inferSelect;

// Status history (tracks status changes across inspections)
export const statusHistory = sqliteTable("status_history", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull(),
  oldStatus: text("old_status"),
  newStatus: text("new_status").notNull(),
  reportId: integer("report_id"),
  createdAt: text("created_at"),
});

export const insertStatusHistorySchema = createInsertSchema(statusHistory).omit({ id: true });
export type InsertStatusHistory = z.infer<typeof insertStatusHistorySchema>;
export type StatusHistory = typeof statusHistory.$inferSelect;

// Inspection notes — explicit "add note for this inspection" entries, kept separate from
// the canonical observation/action text so a note never overwrites the observation.
export const inspectionNotes = sqliteTable("inspection_notes", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  defectId: integer("defect_id").notNull(),
  reportId: integer("report_id").notNull(),
  author: text("author").default(""),
  text: text("text").notNull(),
  createdAt: text("created_at").notNull(),
});

export const insertInspectionNoteSchema = createInsertSchema(inspectionNotes).omit({ id: true });
export type InsertInspectionNote = z.infer<typeof insertInspectionNoteSchema>;
export type InspectionNote = typeof inspectionNotes.$inferSelect;

// Project Status — Narrative/Hold blocks (multiple per project)
export const narrativeHolds = sqliteTable("narrative_holds", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull(),
  title: text("title").notNull().default(""),
  body: text("body").notNull().default(""),
  status: text("status").notNull().default("Active"), // "Active" | "Lifted" | "For information"
  dateRaised: text("date_raised").default(""),
  dateLifted: text("date_lifted"),
  figures: text("figures").default("[]"), // JSON: [{filename, caption}]
  audience: text("audience").notNull().default("both"), // "both" | "contractor" | "client"
  sortOrder: integer("sort_order").default(0),
  createdAt: text("created_at").notNull(),
  updatedAt: text("updated_at"),
});

// Project Status — Program/Schedule (single per project, UNIQUE projectId)
export const programSchedule = sqliteTable("program_schedule", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull().unique(),
  programImageFilename: text("program_image_filename"),
  asAtDate: text("as_at_date").default(""),
  varianceText: text("variance_text").default(""),
  projectedCompletion: text("projected_completion").default(""),
  statusNarrative: text("status_narrative").default(""),
  audience: text("audience").notNull().default("both"),
  updatedAt: text("updated_at"),
});

// Project Status — Stage Progress Map (single per project, UNIQUE projectId)
export const stageProgressMap = sqliteTable("stage_progress_map", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  projectId: integer("project_id").notNull().unique(),
  planImageFilename: text("plan_image_filename"),
  stages: text("stages").default("[]"), // JSON: [{stageName, status: "Not started"|"Underway"|"Complete"}]
  audience: text("audience").notNull().default("both"),
  updatedAt: text("updated_at"),
});

export const insertNarrativeHoldSchema = createInsertSchema(narrativeHolds).omit({ id: true });
export const insertProgramScheduleSchema = createInsertSchema(programSchedule).omit({ id: true });
export const insertStageProgressMapSchema = createInsertSchema(stageProgressMap).omit({ id: true });

export type InsertNarrativeHold = z.infer<typeof insertNarrativeHoldSchema>;
export type NarrativeHold = typeof narrativeHolds.$inferSelect;
export type InsertProgramSchedule = z.infer<typeof insertProgramScheduleSchema>;
export type ProgramSchedule = typeof programSchedule.$inferSelect;
export type InsertStageProgressMap = z.infer<typeof insertStageProgressMapSchema>;
export type StageProgressMap = typeof stageProgressMap.$inferSelect;

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
