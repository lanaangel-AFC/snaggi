// ============================================================================
// render-helpers.ts — shared, renderer-agnostic helpers used by both the DOCX
// and PDF export renderers (and, where pure, by the tree-builder).
//
// Pass 2: extracted out of report-detail.tsx so the two renderers can share a
// single copy of the font/colour/safeText/image/location/work-type helpers and
// the page file shrinks. NONE of these helpers filter report content — filtering
// lives exclusively in buildReportTree (report-tree.ts). These are formatting
// and asset-loading utilities only.
// ============================================================================

import { formatLocation, getLocationDimensions } from "@shared/location";

// API base — mirrors the convention used elsewhere in the client.
const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

// ---------------------------------------------------------------------------
// Text safety
// ---------------------------------------------------------------------------

// Defensive helper: coerce any value to a string for docx TextRun / pdf text.
export const safeText = (v: unknown): string => (v == null ? "" : String(v));

// ---------------------------------------------------------------------------
// Overdue helpers (shared by DOCX + PDF Action List status columns)
// ---------------------------------------------------------------------------

// Today as a local YYYY-MM-DD string. dueDate is captured via <input type="date">
// which stores YYYY-MM-DD, so a plain string compare against todayISO() is both
// correct and timezone-safe (no Date parsing / UTC-shift surprises).
export const todayISO = (): string => {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${mm}-${dd}`;
};

// An Action List row is overdue iff it has a dueDate strictly before today AND
// the item is NOT closed/archived. Closed (status === "complete") and Archived
// items are NEVER flagged overdue even when their dueDate is in the past.
export const isDefectOverdue = (defect: any): boolean => {
  if (!defect) return false;
  if (defect.status === "complete" || defect.status === "archived") return false;
  const due = safeText(defect.dueDate);
  return !!due && due < todayISO();
};

// ---------------------------------------------------------------------------
// Image loading / compression
// ---------------------------------------------------------------------------

// Load compressed thumbnail for report exports (smaller file size).
export async function loadImageBlob(filename: string): Promise<Blob | null> {
  try {
    const res = await fetch(`${API_BASE}/api/thumbs/${filename}`);
    return await res.blob();
  } catch {
    return null;
  }
}

export async function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result as string);
    reader.readAsDataURL(blob);
  });
}

export async function blobToArrayBuffer(blob: Blob): Promise<ArrayBuffer> {
  return blob.arrayBuffer();
}

// Compress an image blob via Canvas — returns a smaller ArrayBuffer (for Word export).
export async function compressImageForExport(blob: Blob, maxWidth = 800, quality = 0.7): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const scale = Math.min(1, maxWidth / img.width);
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext("2d")!;
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      canvas.toBlob(
        (result) => {
          if (result) result.arrayBuffer().then(resolve).catch(reject);
          else reject(new Error("Canvas toBlob failed"));
        },
        "image/jpeg",
        quality,
      );
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image load failed"));
    };
    img.src = url;
  });
}

// Compress an image blob via Canvas — returns a data URL string (for PDF export).
export async function compressImageForPdfExport(blob: Blob, maxWidth = 800, quality = 0.7): Promise<string> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const scale = Math.min(1, maxWidth / img.width);
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext("2d")!;
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image load failed"));
    };
    img.src = url;
  });
}

// Load the AFC logo (both data URL for PDF and ArrayBuffer for DOCX).
export async function loadAfcLogo(): Promise<{ dataUrl: string; buffer: ArrayBuffer } | null> {
  try {
    const res = await fetch(`${API_BASE}/api/public/afc-logo.png`);
    if (!res.ok) return null;
    const blob = await res.blob();
    const [dataUrl, buffer] = await Promise.all([blobToDataUrl(blob), blobToArrayBuffer(blob)]);
    return { dataUrl, buffer };
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Work types
// ---------------------------------------------------------------------------

// Work type code to human-readable label (built-in fallback set).
export const WORK_TYPE_LABELS: Record<string, string> = {
  CR: "Concrete Repair", CK: "Caulking", PT: "Painting", WP: "Waterproofing",
  GL: "Glazing", CL: "Cladding", ST: "Structural", SE: "Sealant",
  FL: "Flashing", RR: "Render Repair", GK: "Gasket", SR: "Steel Repair",
  SM: "Steel Removal", BR: "Brick Repair", LR: "Lintel Repair",
  SS: "Sill Stabilisation", GW: "General Works", OT: "Other",
};

// Parse work type label from UID (variable-length).
export function getWorkTypeLabel(uid: string): string {
  const parts = uid.split("-");
  const wtIdx = parts.findIndex((p) => /^[A-Z]{2,3}$/i.test(p));
  const code = wtIdx >= 0 ? parts[wtIdx] : "";
  return WORK_TYPE_LABELS[code] || code;
}

// ---------------------------------------------------------------------------
// Location
// ---------------------------------------------------------------------------

// Elevation code to full name mapping.
export const ELEVATION_NAMES: Record<string, string> = {
  N: "North", S: "South", E: "East", W: "West",
  NE: "North East", NW: "North West", SE: "South East", SW: "South West",
};

// Derive location string from defect UID (variable-length) — legacy fallback.
export function deriveLocation(uid: string): string {
  const parts = uid.split("-");
  const wtIdx = parts.findIndex((p) => /^[A-Z]{2,3}$/i.test(p));
  if (wtIdx < 0) return "";
  const before = parts.slice(0, wtIdx);
  const alphaIdx = before.findIndex((p) => /^[A-Z]+$/i.test(p));
  const elevation = alphaIdx >= 0 ? before[alphaIdx] : "";
  const numerics = before.filter((_, i) => i !== alphaIdx);
  const dropVal = numerics[0] || "";
  const levelVal = numerics[1] || "";
  const segments: string[] = [];
  if (elevation) { segments.push(`${ELEVATION_NAMES[elevation] || elevation} Elevation`); }
  if (dropVal) segments.push(`Drop ${parseInt(dropVal, 10) || dropVal}`);
  if (levelVal) segments.push(`Level ${parseInt(levelVal, 10) || levelVal}`);
  return segments.join(", ");
}

// SINGLE source of truth for a defect's location string in BOTH the register and
// the card. Prefers the structured location object (location_structured) via the
// shared formatLocation() helper; falls back to UID parsing for legacy rows.
export function formatDefectLocation(defect: any, dims: string[]): string {
  const structured = defect?.locationStructured;
  if (structured) {
    const s = formatLocation(structured, dims);
    if (s) return s;
  }
  return deriveLocation(defect?.uid || "");
}

export { getLocationDimensions };

// ---------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------

// Long form date for cover/intro.
export function formatReportDate(dateStr?: string): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });
}

// Short form date for photo datestamps.
export function formatPhotoDate(dateStr?: string): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric" });
}

// ---------------------------------------------------------------------------
// UID parsing / sorting (shared by both renderers and the tree-builder)
// ---------------------------------------------------------------------------

export function parseUidParts(uid: string) {
  const parts = uid.split("-");
  const wtIdx = parts.findIndex((p: string) => /^[A-Z]{2,3}$/i.test(p));
  if (wtIdx < 0) return { elev: "", drop: 0, level: 0, work: "", num: 0 };
  const before = parts.slice(0, wtIdx);
  const alphaIdx = before.findIndex((p: string) => /^[A-Z]+$/i.test(p));
  const elev = alphaIdx >= 0 ? before[alphaIdx] : "";
  const numerics = before.filter((_: string, i: number) => i !== alphaIdx);
  return {
    elev,
    drop: parseInt(numerics[0] || "0", 10),
    level: parseInt(numerics[1] || "0", 10),
    work: parts[wtIdx] || "",
    num: parseInt(parts[wtIdx + 1] || "0", 10),
  };
}

// ---------------------------------------------------------------------------
// Photo slot ordering (shared by both renderers, the shared report, and the form)
// ---------------------------------------------------------------------------

// Compare two photo slot keys for ascending order. WIP slots sort numerically
// (wip1 < wip2 < ... < wip10 ...); the "complete" slot always sorts last.
// e.g. ["complete","wip3","wip1","wip7"] => ["wip1","wip3","wip7","complete"].
export function comparePhotoSlots(a: string, b: string): number {
  const aVal = a === "complete" ? Number.POSITIVE_INFINITY : parseInt(a.replace(/^wip/, ""), 10);
  const bVal = b === "complete" ? Number.POSITIVE_INFINITY : parseInt(b.replace(/^wip/, ""), 10);
  return aVal - bVal;
}

// Human-readable label for any photo slot key: "complete" => "Complete",
// "wip6" => "WIP 6". Falls back to the raw key for anything unexpected.
export function photoSlotLabel(slot: string): string {
  if (slot === "complete") return "Complete";
  const m = /^wip([0-9]+)$/.exec(slot);
  return m ? `WIP ${m[1]}` : slot;
}

// Sort by: Elevation, Drop (asc), Level (desc/highest first), WorkType, Number (asc).
export function sortByUid(a: any, b: any): number {
  const ap = parseUidParts(a.uid);
  const bp = parseUidParts(b.uid);
  if (ap.elev !== bp.elev) return ap.elev.localeCompare(bp.elev);
  if (ap.drop !== bp.drop) return ap.drop - bp.drop;
  if (ap.level !== bp.level) return bp.level - ap.level;
  if (ap.work !== bp.work) return ap.work.localeCompare(bp.work);
  return ap.num - bp.num;
}
