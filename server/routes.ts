import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage, dataDir } from "./storage";
import { insertProjectSchema, insertDefectSchema } from "@shared/schema";
import multer from "multer";
import path from "path";
import fs from "fs";

const uploadDir = path.join(dataDir, "uploads");
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, uploadDir),
    filename: (_req, file, cb) => {
      const ext = path.extname(file.originalname);
      const name = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}${ext}`;
      cb(null, name);
    },
  }),
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype.startsWith("image/")) {
      cb(null, true);
    } else {
      cb(new Error("Only image files are allowed"));
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

  // === DEFECTS ===
  app.get("/api/projects/:projectId/defects", async (req, res) => {
    const defects = await storage.getDefectsByProject(Number(req.params.projectId));
    res.json(defects);
  });

  app.get("/api/defects/:id", async (req, res) => {
    const defect = await storage.getDefect(Number(req.params.id));
    if (!defect) return res.status(404).json({ message: "Defect not found" });
    res.json(defect);
  });

  app.post("/api/projects/:projectId/defects", async (req, res) => {
    const projectId = Number(req.params.projectId);
    const uidPrefix = req.body.uidPrefix;
    if (!uidPrefix) return res.status(400).json({ message: "UID prefix (drop-level-worktype) is required" });
    // Use the client-provided UID if they set a custom number, otherwise auto-generate
    const uid = req.body.uidOverride || await storage.getNextDefectUid(projectId, uidPrefix);
    const { uidPrefix: _removed, uidOverride: _removed2, ...rest } = req.body;
    const recordType = rest.recordType || "defect";
    const data = { ...rest, projectId, uid, recordType };
    const parsed = insertDefectSchema.safeParse(data);
    if (!parsed.success) return res.status(400).json({ message: parsed.error.message });
    const defect = await storage.createDefect(parsed.data);
    res.status(201).json(defect);
  });

  app.patch("/api/defects/:id", async (req, res) => {
    const defect = await storage.updateDefect(Number(req.params.id), req.body);
    if (!defect) return res.status(404).json({ message: "Defect not found" });
    res.json(defect);
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
    res.status(201).json(photo);
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

  // === PDF REPORT DATA ===
  app.get("/api/projects/:id/report-data", async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
    if (!project) return res.status(404).json({ message: "Project not found" });
    const projectDefects = await storage.getDefectsByProject(project.id);
    const defectsWithPhotos = await Promise.all(
      projectDefects.map(async (d) => {
        const defectPhotos = await storage.getPhotosByDefect(d.id);
        return { ...d, photos: defectPhotos };
      })
    );
    res.json({ project, defects: defectsWithPhotos });
  });

  return httpServer;
}
