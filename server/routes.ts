import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage, dataDir, sqlite, getGlobalSettings, addGlobalWorkType, cleanupCustomWorkTypes } from "./storage";
import { insertProjectSchema, insertDefectSchema, insertReportSchema, insertElevationSchema, insertMarkerSchema, insertDefectLocationSchema } from "@shared/schema";
import { getLocationDimensions } from "@shared/location";
import multer from "multer";
import path from "path";
import fs from "fs";
import sharp from "sharp";
import archiver from "archiver";

const uploadDir = path.join(dataDir, "uploads");
const thumbDir = path.join(dataDir, "thumbs");
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
if (!fs.existsSync(thumbDir)) fs.mkdirSync(thumbDir, { recursive: true });

const diskStorage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, uploadDir),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    const name = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}${ext}`;
    cb(null, name);
  },
});

const upload = multer({
  storage: diskStorage,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype.startsWith("image/")) {
      cb(null, true);
    } else {
      cb(new Error("Only image files are allowed"));
    }
  },
});

// Separate multer for elevation uploads (accepts images + PDFs)
const elevationUpload = multer({
  storage: diskStorage,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype.startsWith("image/") || file.mimetype === "application/pdf") {
      cb(null, true);
    } else {
      cb(new Error("Only image and PDF files are allowed"));
    }
  },
});

// Format a photo's "age" annotation used by the export: "(added DD MMM YYYY)" when
// the photo originated in a report other than the current one, else null.
function formatPhotoAge(photo: any, currentReportId: number | null): string | null {
  const origin = photo.originReportId ?? photo.origin_report_id ?? null;
  if (origin == null || currentReportId == null || origin === currentReportId) return null;
  const raw = photo.createdAt ?? photo.created_at ?? photo.uploadedAt ?? null;
  if (!raw) return null;
  const d = new Date(raw);
  if (isNaN(d.getTime())) return null;
  const day = String(d.getDate()).padStart(2, "0");
  const mon = d.toLocaleString("en-AU", { month: "short" });
  return `(added ${day} ${mon} ${d.getFullYear()})`;
}

// COMPUTED display status for API responses. Stored status stays lowercase
// (open/complete/archived); the API surface uses proper-case display values.
// "Amended" is computed in the live UI from amendment state; the base mapping here
// returns Open/Closed/Archived and callers may override to "Amended" when relevant.
function computeDisplayStatus(status: string): "Open" | "Closed" | "Archived" {
  if (status === "complete") return "Closed";
  if (status === "archived") return "Archived";
  return "Open";
}

// Non-destructive API enrichment for a defect. ADDS fields only (never removes the
// existing ones the client already reads). Stage A: observation alias, displayStatus,
// locationStructured passthrough, and a shaped photos array. history[] deferred to
// Prompt 2 (empty array for now).
async function enrichDefect(d: any): Promise<any> {
  const photos = await storage.getPhotosByDefect(d.id);
  const shapedPhotos = photos.map((p: any) => ({
    ...p,
    url: `/api/uploads/${p.filename}`,
    caption: p.caption ?? null,
    capture_date: p.createdAt ?? p.created_at ?? null,
    age: formatPhotoAge(p, d.reportId ?? d.report_id ?? null),
  }));
  return {
    ...d,
    observation: d.comment, // alias per spec — DO NOT rename the stored column
    displayStatus: computeDisplayStatus(d.status),
    locationStructured: d.locationStructured ?? null,
    legacyId: d.legacyId ?? null, // Stage 1: always NULL (Stage 2 populates)
    history: [], // deferred to Prompt 2
    photos: shapedPhotos,
  };
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  // Serve static public assets (logo etc.)
  app.use("/api/public", (req, res, next) => {
    const publicDir = path.join(process.cwd(), "public");
    const filePath = path.join(publicDir, path.basename(req.path));
    if (fs.existsSync(filePath)) {
      res.sendFile(filePath);
    } else {
      res.status(404).json({ message: "File not found" });
    }
  });

  // Serve uploaded photos
  app.use("/api/uploads", (req, res, next) => {
    const filePath = path.join(uploadDir, path.basename(req.path));
    if (fs.existsSync(filePath)) {
      res.sendFile(filePath);
    } else {
      res.status(404).json({ message: "File not found" });
    }
  });

  // === GLOBAL SETTINGS ===
  app.get("/api/global-settings", async (_req, res) => {
    res.json(getGlobalSettings());
  });

  app.post("/api/global-settings/work-types", async (req, res) => {
    addGlobalWorkType(req.body);
    res.json({ ok: true });
  });

  // === PROJECTS ===
  app.get("/api/projects", async (_req, res) => {
    const projects = await storage.getProjects();
    res.json(projects);
  });

  app.get("/api/projects/:id", async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
    if (!project) return res.status(404).json({ message: "Project not found" });
    res.json(project);
  });

  app.post("/api/projects", async (req, res) => {
    const parsed = insertProjectSchema.safeParse(req.body);
    if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
    const project = await storage.createProject(parsed.data);
    res.status(201).json(project);
  });

  app.patch("/api/projects/:id", async (req, res) => {
    const project = await storage.updateProject(Number(req.params.id), req.body);
    if (!project) return res.status(404).json({ message: "Project not found" });
    res.json(project);
  });

  app.delete("/api/projects/:id", async (req, res) => {
    await storage.deleteProject(Number(req.params.id));
    res.status(204).end();
  });

  // === REPORTS ===
  app.get("/api/projects/:projectId/reports", async (req, res) => {
    const reports = await storage.getReportsByProject(Number(req.params.projectId));
    res.json(reports);
  });

  app.get("/api/reports/:id", async (req, res) => {
    const report = await storage.getReport(Number(req.params.id));
    if (!report) return res.status(404).json({ message: "Report not found" });
    res.json(report);
  });

  app.post("/api/projects/:projectId/reports", async (req, res) => {
    const projectId = Number(req.params.projectId);
    const data = { ...req.body, projectId, createdAt: new Date().toISOString() };
    const parsed = insertReportSchema.safeParse(data);
    if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
    const report = await storage.createReport(parsed.data);
    res.status(201).json(report);
  });

  app.patch("/api/reports/:id", async (req, res) => {
    const report = await storage.updateReport(Number(req.params.id), req.body);
    if (!report) return res.status(404).json({ message: "Report not found" });
    res.json(report);
  });

  app.delete("/api/reports/:id", async (req, res) => {
    await storage.deleteReport(Number(req.params.id));
    res.status(204).end();
  });

  app.post("/api/reports/:id/copy", async (req, res) => {
    try {
      const newReport = await storage.copyReport(Number(req.params.id));
      res.status(201).json(newReport);
    } catch (err: any) {
      res.status(400).json({ message: err.message });
    }
  });

  // === DEFECTS ===
  app.get("/api/projects/:projectId/defects", async (req, res) => {
    const defects = await storage.getDefectsByProject(Number(req.params.projectId));
    res.json(defects);
  });

  app.get("/api/reports/:reportId/defects", async (req, res) => {
    const defects = await storage.getDefectsByReport(Number(req.params.reportId));
    res.json(defects);
  });

  // Lookup defect by UID within a project (for deep-linking from external apps)
  app.get("/api/projects/:projectId/defects/by-uid/:uid", async (req, res) => {
    const defect = await storage.getDefectByUid(Number(req.params.projectId), req.params.uid);
    if (!defect) return res.status(404).json({ message: "Defect not found" });
    res.json(defect);
  });

  app.get("/api/defects/:id", async (req, res) => {
    const defect = await storage.getDefect(Number(req.params.id));
    if (!defect) return res.status(404).json({ message: "Defect not found" });
    res.json(await enrichDefect(defect));
  });

  app.post("/api/projects/:projectId/defects", async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);
      const uidPrefix = req.body.uidPrefix;
      if (!uidPrefix) return res.status(400).json({ message: "UID prefix (drop-level-worktype) is required" });
      // Use the client-provided UID if they set a custom number, otherwise auto-generate
      const uid = req.body.uidOverride || await storage.getNextDefectUid(projectId, uidPrefix);

      // Check for duplicate UID within the same report (not project-wide, since copies share UIDs)
      const reportIdVal = req.body.reportId ? Number(req.body.reportId) : null;
      if (reportIdVal) {
        const existingInReport = await storage.getDefectsByReport(reportIdVal);
        const duplicate = existingInReport.find((d) => d.uid === uid);
        if (duplicate) {
          return res.status(400).json({ message: `UID "${uid}" is already in use in this report. Please change the Number field.` });
        }
      }

      const { uidPrefix: _removed, uidOverride: _removed2, ...rest } = req.body;
      const recordType = rest.recordType || "defect";
      const reportId = rest.reportId != null && !isNaN(Number(rest.reportId)) ? Number(rest.reportId) : null;
      const data = { ...rest, projectId, uid, recordType, reportId };
      const parsed = insertDefectSchema.safeParse(data);
      if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
      const defect = await storage.createDefect(parsed.data);
      res.status(201).json(defect);
    } catch (err: any) {
      console.error("Error creating defect:", err);
      res.status(500).json({ message: err.message || "Failed to create defect" });
    }
  });

  app.patch("/api/defects/:id", async (req, res) => {
    const defectId = Number(req.params.id);

    // Capture history BEFORE applying the update
    const existing = await storage.getDefect(defectId);
    if (!existing) return res.status(404).json({ message: "Defect not found" });

    const now = new Date().toISOString();
    if (req.body.comment !== undefined && req.body.comment !== existing.comment && existing.comment?.trim()) {
      await storage.createObservationHistory({
        defectId,
        reportId: existing.reportId!,
        text: existing.comment,
        createdAt: now,
      });
    }
    if (req.body.actionRequired !== undefined && req.body.actionRequired !== existing.actionRequired && existing.actionRequired?.trim()) {
      await storage.createActionHistory({
        defectId,
        reportId: existing.reportId!,
        text: existing.actionRequired,
        createdAt: now,
      });
    }

    // Track status changes
    if (req.body.status !== undefined && req.body.status !== existing.status) {
      await storage.createStatusHistory({
        defectId,
        oldStatus: existing.status,
        newStatus: req.body.status,
        reportId: existing.reportId,
        createdAt: now,
      });
    }

    const defect = await storage.updateDefect(defectId, req.body);
    if (!defect) return res.status(404).json({ message: "Defect not found" });

    // Sync markers when uid or status changes
    const markerUpdates: Record<string, string> = {};
    if (req.body.uid) markerUpdates.defectUid = req.body.uid;
    if (req.body.status) markerUpdates.status = req.body.status;
    if (Object.keys(markerUpdates).length > 0) {
      await storage.updateMarkersByDefectId(defect.id, markerUpdates);
    }

    res.json(defect);
  });

  // === HISTORY ===
  app.get("/api/defects/:id/observation-history", async (req, res) => {
    const entries = await storage.getObservationHistory(Number(req.params.id));
    // Join report info for each entry
    const result = await Promise.all(entries.map(async (entry) => {
      const report = await storage.getReport(entry.reportId);
      return {
        ...entry,
        reportName: report ? `Insp-${report.inspectionNumber || "01"}` : "Unknown",
        reportDate: report?.inspectionDate || report?.createdAt || "",
      };
    }));
    res.json(result);
  });

  app.get("/api/defects/:id/action-history", async (req, res) => {
    const entries = await storage.getActionHistory(Number(req.params.id));
    const result = await Promise.all(entries.map(async (entry) => {
      const report = await storage.getReport(entry.reportId);
      return {
        ...entry,
        reportName: report ? `Insp-${report.inspectionNumber || "01"}` : "Unknown",
        reportDate: report?.inspectionDate || report?.createdAt || "",
      };
    }));
    res.json(result);
  });

  // === INSPECTION NOTES (explicit add-note for observation/action) ===
  app.post("/api/defects/:id/observation-note", async (req, res) => {
    const defectId = Number(req.params.id);
    const existing = await storage.getDefect(defectId);
    if (!existing) return res.status(404).json({ message: "Defect not found" });
    const text = req.body.text;
    if (!text || !text.trim()) return res.status(400).json({ message: "Text is required" });
    const now = new Date().toISOString();
    const reportId = req.body.reportId ? Number(req.body.reportId) : existing.reportId;
    // Save old observation as history entry
    if (existing.comment?.trim()) {
      await storage.createObservationHistory({
        defectId,
        reportId: reportId!,
        text: existing.comment,
        createdAt: now,
      });
    }
    // Update defect comment to the new note
    const updated = await storage.updateDefect(defectId, { comment: text });
    res.json(updated);
  });

  app.post("/api/defects/:id/action-note", async (req, res) => {
    const defectId = Number(req.params.id);
    const existing = await storage.getDefect(defectId);
    if (!existing) return res.status(404).json({ message: "Defect not found" });
    const text = req.body.text;
    if (!text || !text.trim()) return res.status(400).json({ message: "Text is required" });
    const now = new Date().toISOString();
    const reportId = req.body.reportId ? Number(req.body.reportId) : existing.reportId;
    // Save old action as history entry
    if (existing.actionRequired?.trim()) {
      await storage.createActionHistory({
        defectId,
        reportId: reportId!,
        text: existing.actionRequired,
        createdAt: now,
      });
    }
    // Update defect actionRequired to the new note
    const updated = await storage.updateDefect(defectId, { actionRequired: text });
    res.json(updated);
  });

  app.delete("/api/defects/:id", async (req, res) => {
    await storage.deleteDefect(Number(req.params.id));
    res.status(204).end();
  });

  // === PHOTOS ===
  app.get("/api/defects/:defectId/photos", async (req, res) => {
    const photos = await storage.getPhotosByDefect(Number(req.params.defectId));
    res.json(photos);
  });

  // Upload photo with slot (wip1, wip2, wip3, wip4, wip5, complete)
  app.post("/api/defects/:defectId/photos", upload.single("photo"), async (req, res) => {
    if (!req.file) return res.status(400).json({ message: "No file uploaded" });
    const slot = req.body.slot || "wip1";
    const validSlots = ["wip1", "wip2", "wip3", "wip4", "wip5", "complete"];
    if (!validSlots.includes(slot)) {
      return res.status(400).json({ message: "Invalid slot. Must be wip1, wip2, wip3, wip4, wip5, or complete" });
    }

    const defectId = Number(req.params.defectId);

    // If a photo already exists in this slot, replace it
    const existingPhotos = await storage.getPhotosByDefect(defectId);
    const existingInSlot = existingPhotos.find((p) => p.slot === slot);
    if (existingInSlot) {
      // Delete old photo file and record
      const oldPath = path.join(uploadDir, existingInSlot.filename);
      if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
      await storage.deletePhoto(existingInSlot.id);
    }

    const reportId = req.body.reportId ? Number(req.body.reportId) : null;
    const photo = await storage.createPhoto({
      defectId,
      reportId,
      originReportId: reportId, // new upload → origin is this report
      filename: req.file.filename,
      caption: req.body.caption || null,
      slot,
      createdAt: new Date().toISOString(),
    });
    // Generate compressed thumbnail for report exports
    try {
      const origPath = path.join(uploadDir, req.file.filename);
      const thumbPath = path.join(thumbDir, req.file.filename);
      await sharp(origPath)
        .resize(800, 600, { fit: "inside", withoutEnlargement: true })
        .jpeg({ quality: 70 })
        .toFile(thumbPath);
    } catch (err) {
      console.warn("Thumbnail generation failed:", err);
    }

    res.status(201).json(photo);
  });

  // Serve compressed thumbnails for report exports
  app.use("/api/thumbs", (req, res, next) => {
    const filePath = path.join(thumbDir, path.basename(req.path));
    if (fs.existsSync(filePath)) {
      res.sendFile(filePath);
    } else {
      const origPath = path.join(uploadDir, path.basename(req.path));
      if (fs.existsSync(origPath)) res.sendFile(origPath);
      else res.status(404).json({ message: "File not found" });
    }
  });

  // Download photos for a report as a zip.
  // scope=current (default): only photos with report_id matching this report.
  // scope=all: every photo across the entire project.
  app.get("/api/reports/:reportId/photos-zip", async (req, res) => {
    try {
      const reportId = Number(req.params.reportId);
      const report = await storage.getReport(reportId);
      if (!report) return res.status(404).json({ message: "Report not found" });
      const scope = (req.query.scope as string) || "current";

      // Sanitise names for filename
      const project = await storage.getProject(report.projectId);
      const safeName = (s: string) => s.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 40);
      const projectSlug = project ? safeName(project.name) : `project-${report.projectId}`;
      const reportSlug = report.inspectionNumber ? `Insp-${safeName(report.inspectionNumber)}` : `report-${reportId}`;

      // Collect defects: current report only for scope=current, entire project for scope=all
      const defectList = scope === "all"
        ? await storage.getDefectsByProject(report.projectId)
        : await storage.getDefectsByReport(reportId);

      const archive = archiver("zip", { zlib: { level: 1 } });
      let included = 0;
      const seen = new Set<string>(); // prevent duplicate filenames in zip

      for (const defect of defectList) {
        const defectPhotos = await storage.getPhotosByDefect(defect.id);
        for (const photo of defectPhotos) {
          // For scope=current, only include photos tagged to this report
          if (scope !== "all" && photo.reportId !== reportId) continue;
          const filePath = path.join(uploadDir, photo.filename);
          if (!fs.existsSync(filePath)) continue;
          let zipName = `${defect.uid}_${photo.slot}.jpg`;
          if (seen.has(zipName)) zipName = `${defect.uid}_${photo.slot}_${photo.id}.jpg`;
          seen.add(zipName);
          archive.file(filePath, { name: zipName });
          included++;
        }
      }

      // No silent fallback — if no photos match, return a clear error
      if (included === 0) {
        return res.status(404).json({
          message: scope === "all"
            ? "No photos found in this project"
            : "No photos uploaded during this inspection",
        });
      }

      const filenameSuffix = scope === "all" ? "all" : reportSlug;
      res.setHeader("Content-Type", "application/zip");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="photos-${projectSlug}-${filenameSuffix}.zip"`,
      );
      archive.pipe(res);
      await archive.finalize();
    } catch (err: any) {
      console.error("Zip error:", err);
      if (!res.headersSent) res.status(500).json({ message: "Failed to create zip" });
    }
  });

  // Update photo fields (caption, newOverride, reportId)
  app.patch("/api/photos/:id", async (req, res) => {
    const id = Number(req.params.id);
    const existing = await storage.getPhoto(id);
    if (!existing) return res.status(404).json({ message: "Photo not found" });
    if (req.body.caption !== undefined) {
      await storage.updatePhotoCaption(id, req.body.caption || "");
    }
    if (req.body.newOverride !== undefined) {
      await storage.updatePhotoNewOverride(id, req.body.newOverride);
    }
    if (req.body.reportId !== undefined) {
      await storage.updatePhotoReportId(id, req.body.reportId);
    }
    const photo = await storage.getPhoto(id);
    res.json(photo);
  });

  app.delete("/api/photos/:id", async (req, res) => {
    const photo = await storage.deletePhoto(Number(req.params.id));
    if (photo) {
      const filePath = path.join(uploadDir, photo.filename);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    }
    res.status(204).end();
  });

  // === NEXT UID ===
  app.get("/api/projects/:projectId/next-uid", async (req, res) => {
    const prefix = req.query.prefix as string | undefined;
    const uid = await storage.getNextDefectUid(Number(req.params.projectId), prefix);
    res.json({ uid });
  });

  // === PDF REPORT DATA (scoped to report) ===
  app.get("/api/reports/:id/report-data", async (req, res) => {
    const report = await storage.getReport(Number(req.params.id));
    if (!report) return res.status(404).json({ message: "Report not found" });
    const project = await storage.getProject(report.projectId);
    if (!project) return res.status(404).json({ message: "Project not found" });

    // Determine the inspection window: [report.createdAt, nextReport.createdAt)
    const allReports = await storage.getReportsByProject(report.projectId);
    const sortedReports = allReports.sort((a, b) => a.createdAt.localeCompare(b.createdAt));
    const reportIdx = sortedReports.findIndex(r => r.id === report.id);
    const windowStart = report.createdAt;
    const windowEnd = reportIdx >= 0 && reportIdx + 1 < sortedReports.length
      ? sortedReports[reportIdx + 1].createdAt
      : new Date("2099-12-31T23:59:59Z").toISOString();

    const reportDefects = await storage.getDefectsByReport(report.id);
    const defectsWithPhotos = await Promise.all(
      reportDefects.map(async (d) => {
        const defectPhotos = await storage.getPhotosByDefect(d.id);
        const locations = await storage.getDefectLocations(d.id);
        const obsHistory = await storage.getObservationHistory(d.id);
        const actHistory = await storage.getActionHistory(d.id);
        const statHistory = await storage.getStatusHistory(d.id);

        // Compute DefectEvents
        const isNew = !!(d.createdAt && d.createdAt >= windowStart && d.createdAt < windowEnd);
        const obsAmended = obsHistory.some(h => h.createdAt >= windowStart && h.createdAt < windowEnd);
        const actAmended = actHistory.some(h => h.createdAt >= windowStart && h.createdAt < windowEnd);
        const photosAdded = defectPhotos.filter(p =>
          p.newOverride === "new" ? true
          : p.newOverride === "not-new" ? false
          : (p.originReportId ?? p.reportId) === report.id
        );
        const locsAdded = locations.filter(l => l.createdAt && l.createdAt >= windowStart && l.createdAt < windowEnd);
        const locsAmended = locations.filter(l => (l as any).updatedAt && (l as any).updatedAt >= windowStart && (l as any).updatedAt < windowEnd);
        const statusChanges = statHistory.filter(s => s.createdAt && s.createdAt >= windowStart && s.createdAt < windowEnd);
        const lastStatusChange = statusChanges.length > 0 ? statusChanges[0] : undefined;

        const events = {
          isNew,
          amendedFields: {
            observation: obsAmended,
            action: actAmended,
            photos: isNew ? defectPhotos.length : photosAdded.length,
            locationsAdded: locsAdded.length,
            locationsAmended: locsAmended.length,
            statusChange: lastStatusChange ? { from: lastStatusChange.oldStatus || "", to: lastStatusChange.newStatus } : undefined,
          },
          photosAddedThisInspection: photosAdded.map(p => p.id),
        };

        // Mark each photo with isThisInspection flag (respects manual override)
        // Uses originReportId (first inspection photo appeared in) instead of reportId
        const photosWithFlag = defectPhotos.map(p => ({
          ...p,
          isThisInspection: p.newOverride === "new" ? true
            : p.newOverride === "not-new" ? false
            : (p.originReportId ?? p.reportId) === report.id,
        }));

        return { ...d, photos: photosWithFlag, locations, observationHistory: obsHistory, actionHistory: actHistory, statusHistory: statHistory, events };
      })
    );
    // Include all project defects for the cumulative summary section
    const allProjectDefects = await storage.getDefectsByProject(project.id);
    res.json({ project, report, defects: defectsWithPhotos, allProjectDefects });
  });

  // Legacy: project-level report-data (returns all defects across all reports)
  app.get("/api/projects/:id/report-data", async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
    if (!project) return res.status(404).json({ message: "Project not found" });
    const projectDefects = await storage.getDefectsByProject(project.id);
    const defectsWithPhotos = await Promise.all(
      projectDefects.map(async (d) => {
        const defectPhotos = await storage.getPhotosByDefect(d.id);
        const locations = await storage.getDefectLocations(d.id);
        return { ...d, photos: defectPhotos, locations };
      })
    );
    res.json({ project, defects: defectsWithPhotos });
  });

  // === ELEVATIONS ===
  app.get("/api/projects/:projectId/elevations", async (req, res) => {
    const elev = await storage.getElevationsByProject(Number(req.params.projectId));
    res.json(elev);
  });

  app.get("/api/elevations/:id", async (req, res) => {
    const elev = await storage.getElevation(Number(req.params.id));
    if (!elev) return res.status(404).json({ message: "Elevation not found" });
    res.json(elev);
  });

  app.post("/api/projects/:projectId/elevations", elevationUpload.single("file"), async (req, res) => {
    if (!req.file) return res.status(400).json({ message: "No file uploaded" });
    const projectId = Number(req.params.projectId);
    const name = req.body.name || req.file.originalname;
    const fileType = req.file.mimetype === "application/pdf" ? "pdf" : "image";
    const elev = await storage.createElevation({
      projectId,
      name,
      filename: req.file.filename,
      fileType,
      createdAt: new Date().toISOString(),
    });
    res.status(201).json(elev);
  });

  app.delete("/api/elevations/:id", async (req, res) => {
    const elev = await storage.getElevation(Number(req.params.id));
    if (elev) {
      const filePath = path.join(uploadDir, elev.filename);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    }
    await storage.deleteElevation(Number(req.params.id));
    res.status(204).end();
  });

  // === MARKERS ===
  app.get("/api/elevations/:elevationId/markers", async (req, res) => {
    const markerList = await storage.getMarkersByElevation(Number(req.params.elevationId));
    res.json(markerList);
  });

  app.post("/api/elevations/:elevationId/markers", async (req, res) => {
    const elevationId = Number(req.params.elevationId);
    const data = {
      ...req.body,
      elevationId,
      createdAt: new Date().toISOString(),
    };
    const parsed = insertMarkerSchema.safeParse(data);
    if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
    const marker = await storage.createMarker(parsed.data);
    res.status(201).json(marker);
  });

  app.patch("/api/markers/:id", async (req, res) => {
    const marker = await storage.updateMarker(Number(req.params.id), req.body);
    if (!marker) return res.status(404).json({ message: "Marker not found" });
    res.json(marker);
  });

  app.delete("/api/markers/:id", async (req, res) => {
    await storage.deleteMarker(Number(req.params.id));
    res.status(204).end();
  });

  // === DEFECT LOCATIONS ===
  app.get("/api/defects/:defectId/locations", async (req, res) => {
    const locations = await storage.getDefectLocations(Number(req.params.defectId));
    res.json(locations);
  });

  app.post("/api/defects/:defectId/locations", async (req, res) => {
    try {
      const defectId = Number(req.params.defectId);
      const defect = await storage.getDefect(defectId);
      if (!defect) return res.status(404).json({ message: "Defect not found" });

      const existing = await storage.getDefectLocations(defectId);
      const displayOrder = req.body.displayOrder ?? existing.length;

      const data = {
        ...req.body,
        defectId,
        displayOrder,
        createdAt: new Date().toISOString(),
      };
      const parsed = insertDefectLocationSchema.safeParse(data);
      if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
      const location = await storage.createDefectLocation(parsed.data);
      res.status(201).json(location);
    } catch (err: any) {
      res.status(500).json({ message: err.message || "Failed to create location" });
    }
  });

  app.patch("/api/defect-locations/:id", async (req, res) => {
    const location = await storage.updateDefectLocation(Number(req.params.id), req.body);
    if (!location) return res.status(404).json({ message: "Location not found" });
    res.json(location);
  });

  app.delete("/api/defect-locations/:id", async (req, res) => {
    await storage.deleteDefectLocation(Number(req.params.id));
    res.status(204).end();
  });

  // ==================== ADMIN: PHOTO AUDIT & CORRECTIVE BACKFILL ====================

  // Audit: show photo→report_id distribution for a project
  app.get("/api/admin/photo-report-audit", async (req, res) => {
    const projectId = Number(req.query.projectId);
    if (!projectId) return res.status(400).json({ message: "projectId query param required" });

    const total = sqlite.prepare(
      `SELECT COUNT(*) as cnt FROM photos p JOIN defects d ON d.id = p.defect_id WHERE d.project_id = ?`
    ).get(projectId) as { cnt: number };

    const byReportId = sqlite.prepare(
      `SELECT p.report_id, r.inspection_number, r.inspection_date, COUNT(*) as cnt
       FROM photos p
       JOIN defects d ON d.id = p.defect_id
       LEFT JOIN reports r ON r.id = p.report_id
       WHERE d.project_id = ?
       GROUP BY p.report_id
       ORDER BY p.report_id`
    ).all(projectId) as { report_id: number | null; inspection_number: string; inspection_date: string; cnt: number }[];

    const nullCount = sqlite.prepare(
      `SELECT COUNT(*) as cnt FROM photos p JOIN defects d ON d.id = p.defect_id WHERE d.project_id = ? AND p.report_id IS NULL`
    ).get(projectId) as { cnt: number };

    const byOriginReportId = sqlite.prepare(
      `SELECT p.origin_report_id, r.inspection_number, r.inspection_date, COUNT(*) as cnt
       FROM photos p
       JOIN defects d ON d.id = p.defect_id
       LEFT JOIN reports r ON r.id = p.origin_report_id
       WHERE d.project_id = ?
       GROUP BY p.origin_report_id
       ORDER BY p.origin_report_id`
    ).all(projectId) as { origin_report_id: number | null; inspection_number: string; inspection_date: string; cnt: number }[];

    const carryOverCount = sqlite.prepare(
      `SELECT COUNT(*) as cnt FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ? AND p.origin_report_id IS NOT NULL AND p.origin_report_id != p.report_id`
    ).get(projectId) as { cnt: number };

    const sample = sqlite.prepare(
      `SELECT p.id, p.defect_id, p.report_id, p.origin_report_id, p.created_at, p.slot, p.new_override, d.uid as defect_uid
       FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ?
       ORDER BY p.report_id, p.id
       LIMIT 50`
    ).all(projectId);

    const carryOverSample = sqlite.prepare(
      `SELECT p.id, p.defect_id, p.report_id, p.origin_report_id, p.slot, d.uid as defect_uid
       FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ? AND p.origin_report_id IS NOT NULL AND p.origin_report_id != p.report_id
       ORDER BY p.origin_report_id, p.id
       LIMIT 30`
    ).all(projectId);

    const reports = sqlite.prepare(
      `SELECT id, inspection_number, inspection_date, created_at FROM reports WHERE project_id = ? ORDER BY created_at ASC`
    ).all(projectId);

    res.json({
      projectId, total: total.cnt, nullReportIdCount: nullCount.cnt,
      byReportId, byOriginReportId,
      carryOverCount: carryOverCount.cnt,
      reports, sample, carryOverSample,
    });
  });

  // Corrective backfill: re-assign photo report_ids using timestamp windows.
  // Only touches photos for the specified project. Does NOT overwrite manual overrides.
  app.post("/api/admin/photo-backfill-correct", async (req, res) => {
    const projectId = Number(req.body.projectId);
    if (!projectId) return res.status(400).json({ message: "projectId required in body" });

    const allPhotos = sqlite.prepare(
      `SELECT p.id, p.defect_id, p.created_at, p.new_override, d.project_id
       FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ?`
    ).all(projectId) as { id: number; defect_id: number; created_at: string; new_override: string | null; project_id: number }[];

    const projectReports = sqlite.prepare(
      `SELECT id, created_at FROM reports WHERE project_id = ? ORDER BY created_at ASC`
    ).all(projectId) as { id: number; created_at: string }[];

    if (projectReports.length === 0) return res.json({ updated: 0, message: "No reports found for project" });

    const updateStmt = sqlite.prepare(`UPDATE photos SET report_id = ? WHERE id = ?`);
    let updated = 0;

    for (const photo of allPhotos) {
      const photoTs = new Date(photo.created_at).getTime();
      let matched: number | null = null;
      for (let i = 0; i < projectReports.length; i++) {
        const rStart = new Date(projectReports[i].created_at).getTime();
        const rEnd = i + 1 < projectReports.length ? new Date(projectReports[i + 1].created_at).getTime() : Infinity;
        if (photoTs >= rStart && photoTs < rEnd) {
          matched = projectReports[i].id;
          break;
        }
      }
      // If photo predates all reports, assign to first report
      if (matched === null) matched = projectReports[0].id;
      updateStmt.run(matched, photo.id);
      updated++;
    }

    // Record that corrective backfill ran
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('photo_backfill_corrective_p' || ?, ?)`).run(String(projectId), new Date().toISOString());

    res.json({ updated, totalPhotos: allPhotos.length, reportsUsed: projectReports.length });
  });

  // Origin-report backfill: (re-)compute origin_report_id for all photos in a project
  // using UID+slot matching. Can be triggered manually for any project.
  app.post("/api/admin/photo-origin-backfill", async (req, res) => {
    const projectId = req.body.projectId ? Number(req.body.projectId) : null;

    const whereClause = projectId ? `WHERE d.project_id = ?` : ``;
    const params = projectId ? [projectId] : [];

    const allPhotos = sqlite.prepare(`
      SELECT p.id, p.defect_id, p.report_id, p.slot, d.uid as defect_uid, d.project_id,
             r.inspection_number
      FROM photos p
      JOIN defects d ON d.id = p.defect_id
      LEFT JOIN reports r ON r.id = p.report_id
      ${whereClause}
      ORDER BY CAST(r.inspection_number AS INTEGER) ASC, p.id ASC
    `).all(...params) as {
      id: number; defect_id: number; report_id: number | null; slot: string;
      defect_uid: string; project_id: number; inspection_number: string | null;
    }[];

    const originMap = new Map<string, number>();
    const updateStmt = sqlite.prepare(`UPDATE photos SET origin_report_id = ? WHERE id = ?`);
    let updated = 0;
    let carryOvers = 0;

    for (const photo of allPhotos) {
      const key = `${photo.project_id}:${photo.defect_uid}:${photo.slot}`;
      const existingOrigin = originMap.get(key);
      if (existingOrigin !== undefined) {
        updateStmt.run(existingOrigin, photo.id);
        if (existingOrigin !== photo.report_id) carryOvers++;
      } else {
        const origin = photo.report_id;
        if (origin !== null) originMap.set(key, origin);
        updateStmt.run(origin, photo.id);
      }
      updated++;
    }

    res.json({
      updated,
      carryOvers,
      newPhotos: updated - carryOvers,
      projectId: projectId || "all",
    });
  });

  // Strip bad custom work types (LEV/LEVEL, single letters, built-in collisions) across all projects.
  // Idempotent — safe to run repeatedly.
  app.post("/api/admin/cleanup-custom-work-types", async (_req, res) => {
    const result = cleanupCustomWorkTypes();
    res.json({ ok: true, ...result });
  });

  // ==================== UID MIGRATION PREVIEW (Stage 1 — READ ONLY) ====================
  // PREVIEW ONLY. This endpoint writes NOTHING. It computes what the UID migration
  // WOULD produce. Stage 2 (apply) is gated on explicit user approval and is NOT
  // implemented. Project-general: all config flows from the project record.
  app.get("/api/admin/uid-migration-preview", async (req, res) => {
    try {
      const projectId = Number(req.query.projectId);
      if (!projectId) return res.status(400).json({ message: "projectId query param required" });

      const project = sqlite.prepare(`SELECT * FROM projects WHERE id = ?`).get(projectId) as any;
      if (!project) {
        // Unknown project — return an empty, well-formed preview (generality requirement).
        return res.json({
          projectId, projectName: null,
          locationDimensions: getLocationDimensions(null),
          uidProtocol: null, rows: [],
          summary: { totalRows: 0, uniqueLegacyIds: 0, rowsWhereProposedDiffersFromLegacy: 0, duplicateLegacyIdGroups: 0,
            duplicateResolutionStrategy: "Keep latest inspection's row as canonical; mark older copies as legacy duplicates" },
        });
      }

      const locationDimensions = getLocationDimensions(project.location_dimensions);
      // Which UID parts are enabled for this project (blank-allowed config from b166a50).
      let enabled: Record<string, boolean> = { elevation: true, drop: true, level: true, workType: true };
      try { enabled = { ...enabled, ...JSON.parse(project.enabled_uid_parts || "{}") }; } catch {}

      // Build a human-readable protocol string from the project's location dims + work type + seq.
      const protocolSegments = [
        ...locationDimensions.filter((d) => enabled[d] !== false),
        ...(enabled.workType !== false ? ["workType"] : []),
        "seq",
      ];
      const uidProtocol = protocolSegments.map((s) => `{${s}}`).join("-");

      // Pull all defects for the project, joined with their report's inspection metadata.
      const defectRows = sqlite.prepare(
        `SELECT d.id, d.uid, d.record_type, d.status, d.report_id,
                d.elevation_code, d.drop_code, d.level_code, d.work_type_code, d.seq_number,
                d.location_structured,
                r.inspection_number, r.created_at AS report_created_at
         FROM defects d
         LEFT JOIN reports r ON r.id = d.report_id
         WHERE d.project_id = ?
         ORDER BY d.id`
      ).all(projectId) as any[];

      // Compute the proposed UID from structured parts, skipping empty/disabled segments
      // (mirrors the b166a50 client-side assembly: segments.filter(s => s !== "")).
      const computeProposedUid = (d: any): string => {
        const segs: string[] = [];
        for (const dim of locationDimensions) {
          if (enabled[dim] === false) continue;
          let v: string | null = null;
          if (dim === "elevation") v = d.elevation_code;
          else if (dim === "drop") v = d.drop_code ? String(d.drop_code).padStart(2, "0") : null;
          else if (dim === "level") v = d.level_code ? String(d.level_code).padStart(2, "0") : null;
          else if (dim === "stage") v = d.drop_code ?? d.elevation_code;
          else v = (d as any)[dim] ?? null;
          if (v != null && String(v).trim() !== "") segs.push(String(v));
        }
        if (enabled.workType !== false && d.work_type_code) segs.push(String(d.work_type_code));
        if (d.seq_number) segs.push(String(d.seq_number).padStart(2, "0"));
        return segs.join("-");
      };

      // Group rows by their current UID (the "legacy" id, since defects are cloned per
      // inspection sharing a UID). The LATEST inspection's row is canonical.
      const groups = new Map<string, any[]>();
      for (const d of defectRows) {
        const key = d.uid || "";
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key)!.push(d);
      }
      // Determine the canonical (latest) row per group by report createdAt, then defect id.
      const canonicalIdByUid = new Map<string, number>();
      let duplicateGroupCount = 0;
      groups.forEach((list, uid) => {
        if (list.length > 1) duplicateGroupCount++;
        const sorted = list.slice().sort((a, b) => {
          const ta = a.report_created_at || ""; const tb = b.report_created_at || "";
          if (ta !== tb) return ta < tb ? 1 : -1; // latest first
          return b.id - a.id;
        });
        canonicalIdByUid.set(uid, sorted[0].id);
      });

      // Inspection label lookup for duplicate notes.
      const inspLabel = (d: any) => d.inspection_number ? `Insp-${d.inspection_number}` : `report ${d.report_id ?? "?"}`;
      const canonicalRowFor = (uid: string) => groups.get(uid)!.find((x) => x.id === canonicalIdByUid.get(uid));

      const displayStatus = (status: string): string => {
        if (status === "complete") return "Closed";
        if (status === "archived") return "Archived";
        return "Open"; // "Amended" is computed in the live app from amendment state; preview shows stored Open
      };

      let changedCount = 0;
      const rows = defectRows.map((d) => {
        const proposedUid = computeProposedUid(d);
        const legacyId = d.uid; // current stored UID is what would become legacy on apply
        const changed = proposedUid !== legacyId;
        if (changed) changedCount++;
        const isCanonical = canonicalIdByUid.get(d.uid) === d.id;
        const isDuplicate = (groups.get(d.uid)?.length || 0) > 1 && !isCanonical;
        let notes: string | null = null;
        if (isDuplicate) {
          const canon = canonicalRowFor(d.uid);
          notes = `Duplicate of canonical row in ${inspLabel(canon)} (defect ${canon.id}); will be marked as archived/historical on apply`;
        }
        // Location object from the project's declared dimensions.
        const location: Record<string, string> = {};
        let locStruct: Record<string, any> = {};
        try { locStruct = JSON.parse(d.location_structured || "{}"); } catch {}
        for (const dim of locationDimensions) {
          if (locStruct[dim] != null && String(locStruct[dim]).trim() !== "") location[dim] = String(locStruct[dim]);
        }
        return {
          defectId: d.id,
          legacyId,
          proposedUid,
          location,
          workType: d.work_type_code || null,
          type: d.record_type === "observation" ? "Observation" : "Defect",
          status: displayStatus(d.status),
          changed,
          duplicateLegacyId: isDuplicate,
          notes,
        };
      });

      res.json({
        projectId,
        projectName: project.name,
        locationDimensions,
        uidProtocol,
        rows,
        summary: {
          totalRows: rows.length,
          uniqueLegacyIds: groups.size,
          rowsWhereProposedDiffersFromLegacy: changedCount,
          duplicateLegacyIdGroups: duplicateGroupCount,
          duplicateResolutionStrategy: "Keep latest inspection's row as canonical; mark older copies as legacy duplicates",
        },
      });
    } catch (err: any) {
      console.error("Error in uid-migration-preview:", err);
      res.status(500).json({ message: err.message || "Failed to build preview" });
    }
  });

  // ==================== SHARE LINKS ====================
  app.post("/api/reports/:reportId/share-links", async (req, res) => {
    const { randomBytes } = await import("crypto");
    const token = randomBytes(24).toString("hex");
    const link = await storage.createShareLink({
      token,
      reportId: Number(req.params.reportId),
      recipientName: req.body.recipientName || "",
      createdAt: new Date().toISOString(),
    });
    res.status(201).json(link);
  });

  app.get("/api/reports/:reportId/share-links", async (req, res) => {
    const links = await storage.getShareLinksByReport(Number(req.params.reportId));
    res.json(links);
  });

  app.delete("/api/share-links/:id", async (req, res) => {
    await storage.deleteShareLink(Number(req.params.id));
    res.json({ ok: true });
  });

  // PUBLIC: read-only shared report
  app.get("/api/share/:token", async (req, res) => {
    const link = await storage.getShareLinkByToken(req.params.token);
    if (!link) return res.status(404).json({ message: "Share link not found" });
    const report = await storage.getReport(link.reportId);
    if (!report) return res.status(404).json({ message: "Report not found" });
    const project = await storage.getProject(report.projectId);
    if (!project) return res.status(404).json({ message: "Project not found" });
    const reportDefects = await storage.getDefectsByReport(report.id);
    const defectsWithPhotos = await Promise.all(reportDefects.map(async (d) => {
      const photos = await storage.getPhotosByDefect(d.id);
      return { ...d, photos };
    }));
    res.json({
      recipientName: link.recipientName,
      sharedAt: link.createdAt,
      project: { name: project.name, address: project.address, client: project.client, inspector: project.inspector, afcReference: (project as any).afcReference },
      report,
      defects: defectsWithPhotos,
    });
  });

  // PUBLIC: photos via share token
  app.get("/api/share/:token/photo/:filename", async (req, res) => {
    const link = await storage.getShareLinkByToken(req.params.token);
    if (!link) return res.status(404).send("Not found");
    const defects = await storage.getDefectsByReport(link.reportId);
    const photoFilenames = new Set<string>();
    for (const d of defects) {
      const photos = await storage.getPhotosByDefect(d.id);
      photos.forEach((p) => photoFilenames.add(p.filename));
    }
    if (!photoFilenames.has(req.params.filename)) return res.status(404).send("Not found");
    const filePath = path.join(uploadDir, req.params.filename);
    if (!fs.existsSync(filePath)) return res.status(404).send("Not found");
    res.sendFile(filePath);
  });

  return httpServer;
}
