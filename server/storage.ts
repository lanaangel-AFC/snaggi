import {
  type User, type InsertUser, users,
  type Project, type InsertProject, projects,
  type Defect, type InsertDefect, defects,
  type Photo, type InsertPhoto, photos,
} from "@shared/schema";
import { drizzle } from "drizzle-orm/better-sqlite3";
import Database from "better-sqlite3";
import { eq, desc } from "drizzle-orm";
import path from "path";
import fs from "fs";

// Use DATA_DIR env var for persistent storage (Railway volume), fallback to cwd
const dataDir = process.env.DATA_DIR || process.cwd();
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const dbPath = path.join(dataDir, "data.db");
const sqlite = new Database(dbPath);
sqlite.pragma("journal_mode = WAL");

export const db = drizzle(sqlite);
export { dataDir };

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
  // Defects
  getDefectsByProject(projectId: number): Promise<Defect[]>;
  getDefect(id: number): Promise<Defect | undefined>;
  createDefect(defect: InsertDefect): Promise<Defect>;
  updateDefect(id: number, defect: Partial<InsertDefect>): Promise<Defect | undefined>;
  deleteDefect(id: number): Promise<void>;
  getNextDefectUid(projectId: number, prefix?: string): Promise<string>;
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
    db.delete(projects).where(eq(projects.id, id)).run();
  }

  // Defects
  async getDefectsByProject(projectId: number): Promise<Defect[]> {
    return db.select().from(defects).where(eq(defects.projectId, projectId)).all();
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
