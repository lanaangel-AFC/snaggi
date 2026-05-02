import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage, dataDir } from "./storage";
import { insertProjectSchema, insertDefectSchema, insertReportSchema, insertElevationSchema, insertMarkerSchema, insertDefectLocationSchema } from "@shared/schema";
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
    res.json(defect);
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

    const photo = await storage.createPhoto({
      defectId,
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
  // By default, only photos uploaded during the current inspection are included
  // (i.e. photo.createdAt >= report.createdAt). Pass ?scope=all to get every photo.
  app.get("/api/reports/:reportId/photos-zip", async (req, res) => {
    try {
      const reportId = Number(req.params.reportId);
      const report = await storage.getReport(reportId);
      if (!report) return res.status(404).json({ message: "Report not found" });
      const defects = await storage.getDefectsByReport(reportId);
      const scope = (req.query.scope as string) || "current";
      const reportStart = report.createdAt ? new Date(report.createdAt).getTime() : 0;

      const filenameSuffix = scope === "all" ? "all" : "current";
      res.setHeader("Content-Type", "application/zip");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="photos-report-${reportId}-${filenameSuffix}.zip"`,
      );

      const archive = archiver("zip", { zlib: { level: 1 } });
      archive.pipe(res);

      let included = 0;
      for (const defect of defects) {
        const photos = await storage.getPhotosByDefect(defect.id);
        for (const photo of photos) {
          if (scope !== "all") {
            const photoTs = photo.createdAt ? new Date(photo.createdAt).getTime() : 0;
            if (!photoTs || photoTs < reportStart) continue;
          }
          const filePath = path.join(uploadDir, photo.filename);
          if (fs.existsSync(filePath)) {
            archive.file(filePath, { name: `${defect.uid}_${photo.slot}.jpg` });
            included++;
          }
        }
      }

      // If the current-inspection scope is empty (e.g. legacy data with no
      // photos uploaded after the report's createdAt), fall back to all photos
      // so the user gets something useful instead of an empty zip.
      if (included === 0 && scope !== "all") {
        for (const defect of defects) {
          const photos = await storage.getPhotosByDefect(defect.id);
          for (const photo of photos) {
            const filePath = path.join(uploadDir, photo.filename);
            if (fs.existsSync(filePath)) {
              archive.file(filePath, { name: `${defect.uid}_${photo.slot}.jpg` });
            }
          }
        }
      }

      await archive.finalize();
    } catch (err: any) {
      console.error("Zip error:", err);
      if (!res.headersSent) res.status(500).json({ message: "Failed to create zip" });
    }
  });

  // Update photo caption
  app.patch("/api/photos/:id", async (req, res) => {
    const photo = await storage.updatePhotoCaption(Number(req.params.id), req.body.caption || "");
    if (!photo) return res.status(404).json({ message: "Photo not found" });
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
    const reportDefects = await storage.getDefectsByReport(report.id);
    const defectsWithPhotos = await Promise.all(
      reportDefects.map(async (d) => {
        const defectPhotos = await storage.getPhotosByDefect(d.id);
        const locations = await storage.getDefectLocations(d.id);
        return { ...d, photos: defectPhotos, locations };
      })
    );
    res.json({ project, report, defects: defectsWithPhotos });
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

  return httpServer;
}
