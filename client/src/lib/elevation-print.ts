// WP8: elevation drawings for section 7 of the new report layout. Loads each project
// elevation that carries a pin for a record in this report, and returns it as a JPEG
// with those pins drawn on, coloured by the record's status. Mirrors the pin drawing
// used by the elevation screen's own export (annotation-canvas.tsx).

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

export const PIN_COLOURS: Record<string, string> = {
  open: "#EF4444",
  in_progress: "#F59E0B",
  complete: "#22C55E",
  archived: "#9CA3AF",
  note: "#94A3B8",
};

export type PrintedElevation = {
  id: number;
  name: string;
  dataUrl: string; // image/jpeg
  width: number; // pixels
  height: number;
  pins: Array<{ uid: string; status: string }>;
};

const MAX_DIM = 2400;

async function loadDrawing(elevation: any): Promise<HTMLCanvasElement | null> {
  const res = await fetch(`${API_BASE}/api/uploads/${elevation.filename}`);
  if (!res.ok) return null;
  if (elevation.fileType === "pdf") {
    const pdfjsLib: any = await import("pdfjs-dist");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
    const pdf = await pdfjsLib.getDocument({ data: await res.arrayBuffer() }).promise;
    const page = await pdf.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(MAX_DIM / Math.max(base.width, base.height), 3);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext("2d")!;
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport, canvas } as any).promise;
    return canvas;
  }
  const bitmap = await createImageBitmap(await res.blob());
  const scale = Math.min(1, MAX_DIM / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  return canvas;
}

function drawPin(ctx: CanvasRenderingContext2D, x: number, y: number, k: number, colour: string, label: string) {
  ctx.save();
  ctx.translate(x, y);
  ctx.scale(k, k);
  ctx.beginPath();
  ctx.arc(0, -12, 8, Math.PI, 0, false);
  ctx.quadraticCurveTo(8, -4, 0, 4);
  ctx.quadraticCurveTo(-8, -4, -8, -12);
  ctx.closePath();
  ctx.fillStyle = colour;
  ctx.fill();
  ctx.strokeStyle = "#fff";
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(0, -12, 3, 0, Math.PI * 2);
  ctx.fillStyle = "#fff";
  ctx.fill();
  ctx.font = "bold 11px sans-serif";
  const w = ctx.measureText(label).width;
  const padX = 4;
  ctx.fillStyle = colour;
  ctx.beginPath();
  if ((ctx as any).roundRect) (ctx as any).roundRect(-w / 2 - padX, -28, w + padX * 2, 16, 3);
  else ctx.rect(-w / 2 - padX, -28, w + padX * 2, 16);
  ctx.fill();
  ctx.fillStyle = "#fff";
  ctx.textBaseline = "middle";
  ctx.fillText(label, -w / 2, -20);
  ctx.restore();
}

/** Elevations for section 7, in the project's elevation order. Never throws: a drawing
 *  that cannot be loaded is skipped so the rest of the report still exports. */
export async function loadReportElevations(projectId: number, reportId: number, recordIds: number[]): Promise<PrintedElevation[]> {
  const wanted = new Set(recordIds);
  const out: PrintedElevation[] = [];
  let elevations: any[] = [];
  try {
    const res = await fetch(`${API_BASE}/api/projects/${projectId}/elevations`);
    if (res.ok) elevations = await res.json();
  } catch {
    return out;
  }
  for (const el of elevations) {
    try {
      const res = await fetch(`${API_BASE}/api/elevations/${el.id}/markers/resolved?reportId=${reportId}`);
      if (!res.ok) continue;
      const markers: any[] = await res.json();
      const pins = markers.filter((m) => m.resolved && m.resolved.inReport && wanted.has(m.resolved.defectId));
      if (pins.length === 0) continue;
      const canvas = await loadDrawing(el);
      if (!canvas) continue;
      const ctx = canvas.getContext("2d")!;
      // Pins are drawn for a ~1000 px wide view; scale so they read the same on paper.
      const k = Math.max(1, Math.max(canvas.width, canvas.height) / 1000);
      const listed: PrintedElevation["pins"] = [];
      for (const m of pins) {
        const status = m.resolved.isNote ? "note" : String(m.resolved.status || m.status || "open");
        const uid = String(m.resolved.uid || m.defectUid || "");
        drawPin(ctx, (m.xPercent / 100) * canvas.width, (m.yPercent / 100) * canvas.height, k, PIN_COLOURS[status] || PIN_COLOURS.open, uid);
        listed.push({ uid, status });
      }
      out.push({ id: el.id, name: String(el.name || `Elevation ${el.id}`), dataUrl: canvas.toDataURL("image/jpeg", 0.85), width: canvas.width, height: canvas.height, pins: listed });
    } catch (err) {
      console.error("[elevation-print] skipped elevation", el?.id, err);
    }
  }
  return out;
}
