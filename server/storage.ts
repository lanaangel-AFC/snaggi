import {
  type User, type InsertUser, users,
  type Project, type InsertProject, projects,
  type Report, type InsertReport, reports,
  type Defect, type InsertDefect, defects,
  type Photo, type InsertPhoto, photos,
} from "@shared/schema";
import { drizzle } from "drizzle-orm/better-sqlite3";
import Database from "better-sqlite3";
import { eq, and, desc } from "drizzle-orm";
import path from "path";
import fs from "fs";

// Use DATA_DIR env var for persistent storage (Railway volume), fallback to cwd
const dataDir = process.env.DATA_DIR || process.cwd();
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const dbPath = path.join(dataDir, "data.db");
const sqlite = new Database(dbPath);
sqlite.pragma("journal_mode = WAL");

// Auto-create tables on startup (no need for drizzle-kit push on deploy)
sqlite.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    address TEXT NOT NULL,
    client TEXT NOT NULL,
    inspector TEXT NOT NULL,
    afc_reference TEXT DEFAULT '',
    revision TEXT DEFAULT '01',

    inspection_number TEXT DEFAULT '',
    inspection_date TEXT DEFAULT '',
    locations_covered TEXT DEFAULT '',
    elevations TEXT DEFAULT '[]',
    attendees TEXT DEFAULT '[]',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    inspection_number TEXT DEFAULT '',
    inspection_date TEXT DEFAULT '',
    revision TEXT DEFAULT '01',
    locations_covered TEXT DEFAULT '',
    elevations TEXT DEFAULT '[]',
    attendees TEXT DEFAULT '[]',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS defects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    report_id INTEGER,
    uid TEXT NOT NULL,
    date_opened TEXT NOT NULL,
    date_closed TEXT,
    comment TEXT NOT NULL,
    action_required TEXT NOT NULL,
    assigned_to TEXT DEFAULT '',
    due_date TEXT DEFAULT '',
    verification_method TEXT NOT NULL,
    verification_person TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'open',
    record_type TEXT NOT NULL DEFAULT 'defect'
  );
  CREATE TABLE IF NOT EXISTS photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    filename TEXT NOT NULL,
    caption TEXT,
    slot TEXT NOT NULL DEFAULT 'wip1',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password TEXT NOT NULL
  );
`);

// Add new columns to existing tables (safe: ALTER TABLE ADD COLUMN IF NOT EXISTS via try/catch)
const safeAddColumn = (table: string, col: string, colDef: string) => {
  try { sqlite.exec(`ALTER TABLE ${table} ADD COLUMN ${col} ${colDef}`); } catch {}
};
safeAddColumn("projects", "locations_covered", "TEXT DEFAULT ''");
safeAddColumn("projects", "elevations", "TEXT DEFAULT '[]'");
safeAddColumn("reports", "elevations", "TEXT DEFAULT '[]'");
safeAddColumn("defects", "record_type", "TEXT NOT NULL DEFAULT 'defect'");
safeAddColumn("defects", "report_id", "INTEGER");

// Migration: for existing defects without reportId, create a default "Report 1" for each project
{
  const orphanRows = sqlite.prepare(
    `SELECT DISTINCT project_id FROM defects WHERE report_id IS NULL`
  ).all() as { project_id: number }[];
  for (const row of orphanRows) {
    // Check if this project already has a report
    const existing = sqlite.prepare(
      `SELECT id FROM reports WHERE project_id = ?`
    ).get(row.project_id) as { id: number } | undefined;
    let reportId: number;
    if (existing) {
      reportId = existing.id;
    } else {
      // Pull per-visit fields from old project row to populate the default report
      const proj = sqlite.prepare(`SELECT * FROM projects WHERE id = ?`).get(row.project_id) as any;
      const result = sqlite.prepare(
        `INSERT INTO reports (project_id, inspection_number, inspection_date, revision, locations_covered, attendees, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).run(
        row.project_id,
        proj?.inspection_number || "",
        proj?.inspection_date || "",
        proj?.revision || "01",
        proj?.locations_covered || "",
        proj?.attendees || "[]",
        new Date().toISOString(),
      );
      reportId = Number(result.lastInsertRowid);
    }
    sqlite.prepare(
      `UPDATE defects SET report_id = ? WHERE project_id = ? AND report_id IS NULL`
    ).run(reportId, row.project_id);
  }
}

export const db = drizzle(sqlite);
export { dataDir };

const uploadDir = path.join(dataDir, "uploads");

export interface IStorage {
  // Users
  getUser(id: number): Promise<User | undefined>;
  getUserByUsername(username: string): Promise<User | undefined>;
  createUser(user: InsertUser): Promise<User>;
  // Projects
  getProjects(): Promise<Project[]>;
  getProject(id: number): Promise<Project | undefined>;
  createProject(project: InsertProject): Promise<Project>;
  updateProject(id: number, project: Partial<InsertProject>): Promise<Project | undefined>;
  deleteProject(id: number): Promise<void>;
  // Reports
  getReportsByProject(projectId: number): Promise<Report[]>;
  getReport(id: number): Promise<Report | undefined>;
  createReport(report: InsertReport): Promise<Report>;
  updateReport(id: number, report: Partial<InsertReport>): Promise<Report | undefined>;
  deleteReport(id: number): Promise<void>;
  copyReport(reportId: number): Promise<Report>;
  // Defects
  getDefectsByProject(projectId: number): Promise<Defect[]>;
  getDefectsByReport(reportId: number): Promise<Defect[]>;
  getDefect(id: number): Promise<Defect | undefined>;
  createDefect(defect: InsertDefect): Promise<Defect>;
  updateDefect(id: number, defect: Partial<InsertDefect>): Promise<Defect | undefined>;
  deleteDefect(id: number): Promise<void>;
  getNextDefectUid(projectId: number, prefix?: string): Promise<string>;
  getDefectByUid(projectId: number, uid: string): Promise<Defect | undefined>;
  // Photos
  getPhotosByDefect(defectId: number): Promise<Photo[]>;
  createPhoto(photo: InsertPhoto): Promise<Photo>;
  updatePhotoCaption(id: number, caption: string): Promise<Photo | undefined>;
  deletePhoto(id: number): Promise<Photo | undefined>;
}

export class DatabaseStorage implements IStorage {
  // Users
  async getUser(id: number): Promise<User | undefined> {
    return db.select().from(users).where(eq(users.id, id)).get();
  }
  async getUserByUsername(username: string): Promise<User | undefined> {
    return db.select().from(users).where(eq(users.username, username)).get();
  }
  async createUser(insertUser: InsertUser): Promise<User> {
    return db.insert(users).values(insertUser).returning().get();
  }

  // Projects
  async getProjects(): Promise<Project[]> {
    return db.select().from(projects).orderBy(desc(projects.id)).all();
  }
  async getProject(id: number): Promise<Project | undefined> {
    return db.select().from(projects).where(eq(projects.id, id)).get();
  }
  async createProject(project: InsertProject): Promise<Project> {
    return db.insert(projects).values(project).returning().get();
  }
  async updateProject(id: number, project: Partial<InsertProject>): Promise<Project | undefined> {
    return db.update(projects).set(project).where(eq(projects.id, id)).returning().get();
  }
  async deleteProject(id: number): Promise<void> {
    // Delete all photos for all defects in this project
    const projectDefects = db.select().from(defects).where(eq(defects.projectId, id)).all();
    for (const d of projectDefects) {
      db.delete(photos).where(eq(photos.defectId, d.id)).run();
    }
    db.delete(defects).where(eq(defects.projectId, id)).run();
    db.delete(reports).where(eq(reports.projectId, id)).run();
    db.delete(projects).where(eq(projects.id, id)).run();
  }

  // Reports
  async getReportsByProject(projectId: number): Promise<Report[]> {
    return db.select().from(reports).where(eq(reports.projectId, projectId)).orderBy(desc(reports.id)).all();
  }
  async getReport(id: number): Promise<Report | undefined> {
    return db.select().from(reports).where(eq(reports.id, id)).get();
  }
  async createReport(report: InsertReport): Promise<Report> {
    return db.insert(reports).values(report).returning().get();
  }
  async updateReport(id: number, report: Partial<InsertReport>): Promise<Report | undefined> {
    return db.update(reports).set(report).where(eq(reports.id, id)).returning().get();
  }
  async deleteReport(id: number): Promise<void> {
    // Delete all photos for all defects in this report
    const reportDefects = db.select().from(defects).where(eq(defects.reportId, id)).all();
    for (const d of reportDefects) {
      db.delete(photos).where(eq(photos.defectId, d.id)).run();
    }
    db.delete(defects).where(eq(defects.reportId, id)).run();
    db.delete(reports).where(eq(reports.id, id)).run();
  }
  async copyReport(reportId: number): Promise<Report> {
    const source = db.select().from(reports).where(eq(reports.id, reportId)).get();
    if (!source) throw new Error("Source report not found");

    // Increment inspection number
    const currentNum = parseInt(source.inspectionNumber || "0", 10);
    const newInspectionNumber = String(currentNum + 1).padStart(2, "0");

    const newReport = db.insert(reports).values({
      projectId: source.projectId,
      inspectionNumber: newInspectionNumber,
      inspectionDate: new Date().toISOString().split("T")[0],
      revision: "01",
      locationsCovered: source.locationsCovered,
      elevations: source.elevations,
      attendees: source.attendees,
      createdAt: new Date().toISOString(),
    }).returning().get();

    // Copy all defects from source report
    const sourceDefects = db.select().from(defects).where(eq(defects.reportId, reportId)).all();
    for (const d of sourceDefects) {
      const newDefect = db.insert(defects).values({
        projectId: d.projectId,
        reportId: newReport.id,
        uid: d.uid,
        dateOpened: d.dateOpened,
        dateClosed: d.dateClosed,
        comment: d.comment,
        actionRequired: d.actionRequired,
        assignedTo: d.assignedTo,
        dueDate: d.dueDate,
        verificationMethod: d.verificationMethod,
        verificationPerson: d.verificationPerson,
        status: d.status,
        recordType: d.recordType,
      }).returning().get();

      // Copy photos
      const defectPhotos = db.select().from(photos).where(eq(photos.defectId, d.id)).all();
      for (const p of defectPhotos) {
        // Copy file on disk
        const srcPath = path.join(uploadDir, p.filename);
        if (fs.existsSync(srcPath)) {
          const ext = path.extname(p.filename);
          const newFilename = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}${ext}`;
          const destPath = path.join(uploadDir, newFilename);
          fs.copyFileSync(srcPath, destPath);
          db.insert(photos).values({
            defectId: newDefect.id,
            filename: newFilename,
            caption: p.caption,
            slot: p.slot,
            createdAt: new Date().toISOString(),
          }).run();
        }
      }
    }

    return newReport;
  }

  // Defects
  async getDefectsByProject(projectId: number): Promise<Defect[]> {
    return db.select().from(defects).where(eq(defects.projectId, projectId)).all();
  }
  async getDefectsByReport(reportId: number): Promise<Defect[]> {
    return db.select().from(defects).where(eq(defects.reportId, reportId)).all();
  }
  async getDefect(id: number): Promise<Defect | undefined> {
    return db.select().from(defects).where(eq(defects.id, id)).get();
  }
  async createDefect(defect: InsertDefect): Promise<Defect> {
    return db.insert(defects).values(defect).returning().get();
  }
  async updateDefect(id: number, defect: Partial<InsertDefect>): Promise<Defect | undefined> {
    return db.update(defects).set(defect).where(eq(defects.id, id)).returning().get();
  }
  async deleteDefect(id: number): Promise<void> {
    db.delete(photos).where(eq(photos.defectId, id)).run();
    db.delete(defects).where(eq(defects.id, id)).run();
  }
  async getNextDefectUid(projectId: number, prefix?: string): Promise<string> {
    if (!prefix) {
      return "01-01-CR-01";
    }
    // prefix is like "01-13-CR" — count existing UIDs that start with this prefix
    const existing = db.select().from(defects).where(eq(defects.projectId, projectId)).all();
    const matching = existing.filter((d) => d.uid.startsWith(prefix + "-"));
    const num = matching.length + 1;
    return `${prefix}-${String(num).padStart(2, "0")}`;
  }

  async getDefectByUid(projectId: number, uid: string): Promise<Defect | undefined> {
    // UIDs may appear in multiple reports (copies); return the most recent (highest reportId)
    const results = db
      .select()
      .from(defects)
      .where(and(eq(defects.projectId, projectId), eq(defects.uid, uid)))
      .orderBy(desc(defects.reportId))
      .all();
    return results[0];
  }

  // Photos
  async getPhotosByDefect(defectId: number): Promise<Photo[]> {
    return db.select().from(photos).where(eq(photos.defectId, defectId)).all();
  }
  async createPhoto(photo: InsertPhoto): Promise<Photo> {
    return db.insert(photos).values(photo).returning().get();
  }
  async updatePhotoCaption(id: number, caption: string): Promise<Photo | undefined> {
    return db.update(photos).set({ caption }).where(eq(photos.id, id)).returning().get();
  }
  async deletePhoto(id: number): Promise<Photo | undefined> {
    const photo = db.select().from(photos).where(eq(photos.id, id)).get();
    if (photo) {
      db.delete(photos).where(eq(photos.id, id)).run();
    }
    return photo;
  }
}

export const storage = new DatabaseStorage();
