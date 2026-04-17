import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link, useParams, useLocation } from "wouter";
import { useState, useRef, useEffect, useCallback } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { ArrowLeft, ZoomIn, ZoomOut, RotateCcw, Crosshair, Eye, Download } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import type { Elevation, Marker, Defect } from "@shared/schema";

const STATUS_COLORS: Record<string, string> = {
  open: "#EF4444",
  in_progress: "#F59E0B",
  complete: "#22C55E",
};

export default function AnnotationCanvas() {
  const { projectId, elevationId } = useParams<{ projectId: string; elevationId: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();

  // Parse ?defect= from the hash (hash routing puts query params inside the hash)
  const defectParam = (() => {
    const hash = window.location.hash;
    const qIndex = hash.indexOf("?");
    if (qIndex === -1) return null;
    const params = new URLSearchParams(hash.substring(qIndex));
    return params.get("defect");
  })();

  // State
  const [imageSrc, setImageSrc] = useState<string | null>(null);
  const [imageLoaded, setImageLoaded] = useState(false);
  const [scale, setScale] = useState(1);
  const [translate, setTranslate] = useState({ x: 0, y: 0 });
  const [isPanning, setIsPanning] = useState(false);
  const [panStart, setPanStart] = useState({ x: 0, y: 0 });
  const [placingMode, setPlacingMode] = useState(false);
  const [markerDialog, setMarkerDialog] = useState<{
    open: boolean;
    x: number;
    y: number;
    editing?: Marker;
  }>({ open: false, x: 0, y: 0 });
  const [deleteConfirm, setDeleteConfirm] = useState<Marker | null>(null);
  const [formDefectId, setFormDefectId] = useState<string>(""); // "custom" or defect id
  const [formUid, setFormUid] = useState("");
  const [formStatus, setFormStatus] = useState("open");
  const [formNote, setFormNote] = useState("");

  const containerRef = useRef<HTMLDivElement>(null);
  const imageRef = useRef<HTMLImageElement>(null);
  const lastTouchDist = useRef<number | null>(null);
  const lastTouchCenter = useRef<{ x: number; y: number } | null>(null);

  // Queries
  const { data: elevation } = useQuery<Elevation>({
    queryKey: ["/api/elevations", elevationId],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/elevations/${elevationId}`);
      return res.json();
    },
  });

  const { data: markerList = [] } = useQuery<Marker[]>({
    queryKey: ["/api/elevations", elevationId, "markers"],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/elevations/${elevationId}/markers`);
      return res.json();
    },
  });

  const { data: defects = [] } = useQuery<Defect[]>({
    queryKey: [`/api/projects/${projectId}/defects`],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/projects/${projectId}/defects`);
      return res.json();
    },
  });

  // Auto-enter placing mode when arriving with ?defect= param
  const [defectParamHandled, setDefectParamHandled] = useState(false);
  useEffect(() => {
    if (!defectParam || defectParamHandled || defects.length === 0) return;
    setDefectParamHandled(true);

    // Pre-select the defect in the form
    const matchingDefect = defects.find((d) => d.uid === defectParam);
    if (matchingDefect) {
      setFormDefectId(String(matchingDefect.id));
      setFormUid(matchingDefect.uid);
      setFormStatus(matchingDefect.status === "complete" ? "complete" : "open");
      setFormNote(matchingDefect.comment || "");
    } else {
      // Defect UID not found in list — use it as custom
      setFormDefectId("custom");
      setFormUid(defectParam);
    }

    setPlacingMode(true);
    toast({ title: `Tap the drawing to mark the location of ${defectParam}` });
  }, [defectParam, defects, defectParamHandled]);

  // Load image or PDF
  useEffect(() => {
    if (!elevation) return;
    if (elevation.fileType === "image") {
      setImageSrc(`/api/uploads/${elevation.filename}`);
    } else if (elevation.fileType === "pdf") {
      loadPdf(elevation.filename);
    }
  }, [elevation]);

  const loadPdf = async (filename: string) => {
    try {
      const pdfjsLib = await import("pdfjs-dist");
      pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
      const url = `/api/uploads/${filename}`;
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Failed to fetch PDF: ${response.status}`);
      const arrayBuffer = await response.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const page = await pdf.getPage(1);
      const baseViewport = page.getViewport({ scale: 1 });
      const maxDim = Math.max(baseViewport.width, baseViewport.height);
      const renderScale = Math.min(2400 / maxDim, 3);
      const viewport = page.getViewport({ scale: renderScale });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d")!;
      await page.render({ canvasContext: ctx, viewport, canvas } as any).promise;
      setImageSrc(canvas.toDataURL("image/png"));
    } catch (err) {
      console.error("PDF load error:", err);
    }
  };

  // Mutations
  const createMarkerMut = useMutation({
    mutationFn: async (data: { defectId?: number | null; defectUid: string; status: string; note: string; xPercent: number; yPercent: number }) => {
      const res = await apiRequest("POST", `/api/elevations/${elevationId}/markers`, data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/elevations", elevationId, "markers"] });
      closeDialog();
    },
  });

  const updateMarkerMut = useMutation({
    mutationFn: async ({ id, ...data }: { id: number; defectId?: number | null; defectUid: string; status: string; note: string }) => {
      const res = await apiRequest("PATCH", `/api/markers/${id}`, data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/elevations", elevationId, "markers"] });
      closeDialog();
    },
  });

  const deleteMarkerMut = useMutation({
    mutationFn: async (id: number) => {
      await apiRequest("DELETE", `/api/markers/${id}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/elevations", elevationId, "markers"] });
      setDeleteConfirm(null);
    },
  });

  const closeDialog = () => {
    setMarkerDialog({ open: false, x: 0, y: 0 });
    setFormDefectId("");
    setFormUid("");
    setFormStatus("open");
    setFormNote("");
  };

  // Handle canvas click for placing markers
  const handleCanvasClick = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      if (!placingMode || !imageRef.current) return;
      const img = imageRef.current;
      const rect = img.getBoundingClientRect();
      const x = ((e.clientX - rect.left) / rect.width) * 100;
      const y = ((e.clientY - rect.top) / rect.height) * 100;
      if (x < 0 || x > 100 || y < 0 || y > 100) return;
      setMarkerDialog({ open: true, x, y });
      setPlacingMode(false);
    },
    [placingMode]
  );

  // Handle existing marker click
  const handleMarkerClick = (e: React.MouseEvent, marker: Marker) => {
    e.stopPropagation();
    // Find matching defect for the dropdown
    if (marker.defectId) {
      setFormDefectId(String(marker.defectId));
    } else {
      setFormDefectId("custom");
    }
    setFormUid(marker.defectUid);
    setFormStatus(marker.status);
    setFormNote(marker.note || "");
    setMarkerDialog({ open: true, x: marker.xPercent, y: marker.yPercent, editing: marker });
  };

  // Zoom controls
  const zoomIn = () => setScale((s) => Math.min(s * 1.3, 5));
  const zoomOut = () => setScale((s) => Math.max(s / 1.3, 0.5));
  const resetZoom = () => {
    setScale(1);
    setTranslate({ x: 0, y: 0 });
  };

  // Mouse/touch pan
  const handlePointerDown = (e: React.PointerEvent) => {
    if (placingMode) return;
    if (scale > 1) {
      setIsPanning(true);
      setPanStart({ x: e.clientX - translate.x, y: e.clientY - translate.y });
    }
  };

  const handlePointerMove = (e: React.PointerEvent) => {
    if (isPanning) {
      setTranslate({
        x: e.clientX - panStart.x,
        y: e.clientY - panStart.y,
      });
    }
  };

  const handlePointerUp = () => setIsPanning(false);

  // Wheel zoom
  const handleWheel = (e: React.WheelEvent) => {
    e.preventDefault();
    const delta = e.deltaY > 0 ? 0.9 : 1.1;
    setScale((s) => Math.max(0.5, Math.min(5, s * delta)));
  };

  // Touch pinch zoom
  const handleTouchStart = (e: React.TouchEvent) => {
    if (e.touches.length === 2) {
      const dx = e.touches[0].clientX - e.touches[1].clientX;
      const dy = e.touches[0].clientY - e.touches[1].clientY;
      lastTouchDist.current = Math.sqrt(dx * dx + dy * dy);
      lastTouchCenter.current = {
        x: (e.touches[0].clientX + e.touches[1].clientX) / 2,
        y: (e.touches[0].clientY + e.touches[1].clientY) / 2,
      };
    }
  };

  const handleTouchMove = (e: React.TouchEvent) => {
    if (e.touches.length === 2 && lastTouchDist.current !== null) {
      const dx = e.touches[0].clientX - e.touches[1].clientX;
      const dy = e.touches[0].clientY - e.touches[1].clientY;
      const dist = Math.sqrt(dx * dx + dy * dy);
      const scaleRatio = dist / lastTouchDist.current;
      setScale((s) => Math.max(0.5, Math.min(5, s * scaleRatio)));
      lastTouchDist.current = dist;
    }
  };

  const handleTouchEnd = () => {
    lastTouchDist.current = null;
    lastTouchCenter.current = null;
  };

  // Handle defect dropdown selection
  const handleDefectSelect = (value: string) => {
    setFormDefectId(value);
    if (value === "custom") {
      setFormUid("");
      setFormStatus("open");
      setFormNote("");
    } else {
      const defect = defects.find((d) => String(d.id) === value);
      if (defect) {
        setFormUid(defect.uid);
        setFormStatus(defect.status === "complete" ? "complete" : "open");
        // Auto-populate note from defect comment only for new markers
        if (!markerDialog.editing) {
          setFormNote(defect.comment || "");
        }
      }
    }
  };

  // Find defect info for a marker (for "View Defect" button)
  const findDefectForMarker = (marker: Marker): Defect | undefined => {
    if (marker.defectId) {
      return defects.find((d) => d.id === marker.defectId);
    }
    return defects.find((d) => d.uid === marker.defectUid);
  };

  // Form submit
  const handleFormSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formUid.trim()) return;
    const defectIdNum = formDefectId && formDefectId !== "custom" ? Number(formDefectId) : null;
    if (markerDialog.editing) {
      updateMarkerMut.mutate({
        id: markerDialog.editing.id,
        defectId: defectIdNum,
        defectUid: formUid.trim(),
        status: formStatus,
        note: formNote.trim(),
      });
    } else {
      createMarkerMut.mutate({
        defectId: defectIdNum,
        defectUid: formUid.trim(),
        status: formStatus,
        note: formNote.trim(),
        xPercent: markerDialog.x,
        yPercent: markerDialog.y,
      });
    }
  };

  const [exporting, setExporting] = useState(false);

  // Export elevation with markers as PNG — draw directly on canvas for reliability
  const handleExportElevation = async () => {
    if (!imageRef.current) return;
    setExporting(true);
    try {
      const img = imageRef.current;
      const canvas = document.createElement("canvas");
      const scale2x = 2; // render at 2x for sharpness
      canvas.width = img.naturalWidth * scale2x;
      canvas.height = img.naturalHeight * scale2x;
      const ctx = canvas.getContext("2d")!;
      ctx.scale(scale2x, scale2x);

      // Draw the elevation image
      ctx.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight);

      // Draw each marker
      const imgW = img.naturalWidth;
      const imgH = img.naturalHeight;

      for (const marker of markerList) {
        const mx = (marker.xPercent / 100) * imgW;
        const my = (marker.yPercent / 100) * imgH;
        const color = STATUS_COLORS[marker.status] || "#EF4444";
        const label = marker.defectUid || "";

        // Pin drop shape
        ctx.save();
        ctx.translate(mx, my);

        // Draw pin body
        ctx.beginPath();
        ctx.arc(0, -12, 8, Math.PI, 0, false);
        ctx.quadraticCurveTo(8, -4, 0, 4);
        ctx.quadraticCurveTo(-8, -4, -8, -12);
        ctx.closePath();
        ctx.fillStyle = color;
        ctx.fill();
        ctx.strokeStyle = "#fff";
        ctx.lineWidth = 1;
        ctx.stroke();

        // White dot in pin
        ctx.beginPath();
        ctx.arc(0, -12, 3, 0, Math.PI * 2);
        ctx.fillStyle = "#fff";
        ctx.fill();

        // UID label above pin
        ctx.font = "bold 11px monospace";
        const textWidth = ctx.measureText(label).width;
        const padX = 4;
        const padY = 3;
        const labelX = -textWidth / 2 - padX;
        const labelY = -28;
        const labelH = 16;

        // Label background
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.roundRect(labelX, labelY, textWidth + padX * 2, labelH, 3);
        ctx.fill();
        ctx.strokeStyle = "rgba(255,255,255,0.4)";
        ctx.lineWidth = 0.5;
        ctx.stroke();

        // Label text
        ctx.fillStyle = "#fff";
        ctx.textBaseline = "middle";
        ctx.fillText(label, labelX + padX, labelY + labelH / 2);

        ctx.restore();
      }

      // Download
      const link = document.createElement("a");
      link.download = `${elevation?.name || "elevation"}_marked.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();

      toast({ title: "Elevation exported" });
    } catch (err) {
      console.error("Export error:", err);
      toast({ title: "Failed to export elevation", variant: "destructive" });
    } finally {
      setExporting(false);
    }
  };

  // Counts
  const counts = {
    open: markerList.filter((m) => m.status === "open").length,
    in_progress: markerList.filter((m) => m.status === "in_progress").length,
    complete: markerList.filter((m) => m.status === "complete").length,
  };

  return (
    <div className="flex flex-col h-screen">
      {/* Header */}
      <div className="flex items-center gap-2 px-3 py-2 border-b bg-background/95 backdrop-blur z-20 flex-shrink-0">
        <Link href={`/projects/${projectId}`}>
          <Button variant="ghost" size="icon" className="h-8 w-8">
            <ArrowLeft className="w-4 h-4" />
          </Button>
        </Link>
        <div className="flex-1 min-w-0">
          <p className="text-sm font-medium truncate">
            {elevation?.name || "Loading..."}
          </p>
        </div>
        <div className="flex items-center gap-1">
          <Button variant="ghost" size="icon" className="h-8 w-8" onClick={zoomOut}>
            <ZoomOut className="w-4 h-4" />
          </Button>
          <span className="text-xs text-muted-foreground w-10 text-center">{Math.round(scale * 100)}%</span>
          <Button variant="ghost" size="icon" className="h-8 w-8" onClick={zoomIn}>
            <ZoomIn className="w-4 h-4" />
          </Button>
          <Button variant="ghost" size="icon" className="h-8 w-8" onClick={resetZoom}>
            <RotateCcw className="w-3.5 h-3.5" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-8 w-8"
            onClick={handleExportElevation}
            disabled={exporting || !imageLoaded}
            title="Export elevation with markers"
          >
            <Download className="w-4 h-4" />
          </Button>
        </div>
      </div>

      {/* Status bar */}
      <div className="flex items-center justify-between px-3 py-1.5 border-b bg-muted/50 flex-shrink-0">
        <div className="flex items-center gap-3 text-xs">
          <span className="flex items-center gap-1">
            <span className="w-2 h-2 rounded-full bg-[#EF4444]" />
            Open: {counts.open}
          </span>
          <span className="flex items-center gap-1">
            <span className="w-2 h-2 rounded-full bg-[#F59E0B]" />
            In Progress: {counts.in_progress}
          </span>
          <span className="flex items-center gap-1">
            <span className="w-2 h-2 rounded-full bg-[#22C55E]" />
            Complete: {counts.complete}
          </span>
        </div>
        <Button
          size="sm"
          variant={placingMode ? "default" : "outline"}
          className="h-7 text-xs gap-1"
          onClick={() => setPlacingMode(!placingMode)}
        >
          <Crosshair className="w-3 h-3" />
          {placingMode ? "Tap to place" : "Add Marker"}
        </Button>
      </div>

      {/* Canvas area */}
      <div
        ref={containerRef}
        className="flex-1 overflow-auto relative bg-muted/30"
        style={{ cursor: placingMode ? "crosshair" : scale > 1 ? "grab" : "default", touchAction: scale > 1 ? "none" : "auto" }}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
        onPointerLeave={handlePointerUp}
        onWheel={handleWheel}
        onTouchStart={handleTouchStart}
        onTouchMove={handleTouchMove}
        onTouchEnd={handleTouchEnd}
      >
        {imageSrc ? (
          <div
            className="relative inline-block origin-top-left"
            style={{
              transform: `translate(${translate.x}px, ${translate.y}px) scale(${scale})`,
              transition: isPanning ? "none" : "transform 0.15s ease-out",
            }}
            onClick={handleCanvasClick}
          >
            <img
              ref={imageRef}
              src={imageSrc}
              alt="Elevation drawing"
              className="w-full h-auto select-none"
              draggable={false}
              onLoad={() => setImageLoaded(true)}
            />
            {/* Markers overlay */}
            {imageLoaded &&
              markerList.map((marker) => (
                <div
                  key={marker.id}
                  className="absolute flex flex-col items-center pointer-events-auto"
                  style={{
                    left: `${marker.xPercent}%`,
                    top: `${marker.yPercent}%`,
                    transform: `translate(-50%, -100%) scale(${1 / scale})`,
                    transformOrigin: "bottom center",
                    zIndex: 10,
                  }}
                  onClick={(e) => handleMarkerClick(e, marker)}
                >
                  {/* Pin */}
                  <div className="flex flex-col items-center cursor-pointer group">
                    <span
                      className="text-xs font-mono font-bold px-1.5 py-0.5 rounded whitespace-nowrap mb-0.5 shadow-sm border border-white/30"
                      style={{
                        backgroundColor: STATUS_COLORS[marker.status] || "#EF4444",
                        color: "#fff",
                        fontSize: "13px",
                        letterSpacing: "0.02em",
                      }}
                    >
                      {marker.defectUid}
                    </span>
                    <svg width="18" height="24" viewBox="0 0 16 22" fill="none">
                      <path
                        d="M8 0C3.6 0 0 3.6 0 8c0 5.4 7.05 13.09 7.35 13.43a.87.87 0 001.3 0C8.95 21.09 16 13.4 16 8c0-4.4-3.6-8-8-8z"
                        fill={STATUS_COLORS[marker.status] || "#EF4444"}
                      />
                      <circle cx="8" cy="8" r="3" fill="white" />
                    </svg>
                  </div>
                </div>
              ))}
          </div>
        ) : (
          <div className="flex items-center justify-center h-full">
            <p className="text-muted-foreground text-sm">Loading drawing...</p>
          </div>
        )}
      </div>

      {/* Marker create/edit dialog */}
      <Dialog open={markerDialog.open} onOpenChange={(v) => !v && closeDialog()}>
        <DialogContent className="max-w-sm">
          <DialogHeader>
            <DialogTitle>{markerDialog.editing ? "Edit Marker" : "New Marker"}</DialogTitle>
          </DialogHeader>
          <form onSubmit={handleFormSubmit} className="space-y-4">
            <div>
              <Label>Link to Defect</Label>
              <Select value={formDefectId} onValueChange={handleDefectSelect}>
                <SelectTrigger>
                  <SelectValue placeholder="Select a defect or enter custom..." />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="custom">Custom UID (not in register)</SelectItem>
                  {defects.map((d) => (
                    <SelectItem key={d.id} value={String(d.id)}>
                      <span className="font-mono">{d.uid}</span>
                      <span className="text-muted-foreground ml-2 text-xs">
                        {d.comment ? d.comment.substring(0, 40) + (d.comment.length > 40 ? "..." : "") : ""}
                      </span>
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            {formDefectId === "custom" && (
              <div>
                <Label htmlFor="defectUid">Defect UID</Label>
                <Input
                  id="defectUid"
                  value={formUid}
                  onChange={(e) => setFormUid(e.target.value)}
                  placeholder="e.g. 03-04-CR-01"
                  className="font-mono"
                />
              </div>
            )}
            {formDefectId && formDefectId !== "custom" && (
              <div className="px-3 py-2 rounded-md bg-muted text-sm">
                <span className="font-mono font-medium">{formUid}</span>
              </div>
            )}
            <div>
              <Label>Status</Label>
              <Select value={formStatus} onValueChange={setFormStatus}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="open">
                    <span className="flex items-center gap-2">
                      <span className="w-2 h-2 rounded-full bg-[#EF4444]" />
                      Open
                    </span>
                  </SelectItem>
                  <SelectItem value="in_progress">
                    <span className="flex items-center gap-2">
                      <span className="w-2 h-2 rounded-full bg-[#F59E0B]" />
                      In Progress
                    </span>
                  </SelectItem>
                  <SelectItem value="complete">
                    <span className="flex items-center gap-2">
                      <span className="w-2 h-2 rounded-full bg-[#22C55E]" />
                      Complete
                    </span>
                  </SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label htmlFor="note">Note (optional)</Label>
              <Textarea
                id="note"
                value={formNote}
                onChange={(e) => setFormNote(e.target.value)}
                placeholder="e.g. Crack at window head"
                rows={2}
              />
            </div>
            {/* View Defect button — navigate within the app */}
            {markerDialog.editing && (() => {
              const defect = findDefectForMarker(markerDialog.editing!);
              if (!defect || !defect.reportId) return null;
              return (
                <Button
                  type="button"
                  variant="outline"
                  className="w-full gap-2"
                  onClick={() => {
                    closeDialog();
                    navigate(`/projects/${projectId}/reports/${defect.reportId}/defects/${defect.id}`);
                  }}
                >
                  <Eye className="w-4 h-4" />
                  View Defect
                </Button>
              );
            })()}
            <div className="flex gap-2">
              <Button
                type="submit"
                disabled={!formUid.trim() || createMarkerMut.isPending || updateMarkerMut.isPending}
                className="flex-1"
              >
                {createMarkerMut.isPending || updateMarkerMut.isPending ? "Saving..." : "Save"}
              </Button>
              {markerDialog.editing && (
                <Button
                  type="button"
                  variant="destructive"
                  onClick={() => setDeleteConfirm(markerDialog.editing!)}
                >
                  Delete
                </Button>
              )}
            </div>
          </form>
        </DialogContent>
      </Dialog>

      {/* Delete confirmation */}
      <AlertDialog open={!!deleteConfirm} onOpenChange={(v) => !v && setDeleteConfirm(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Delete Marker?</AlertDialogTitle>
            <AlertDialogDescription>
              Remove marker "{deleteConfirm?.defectUid}" from this elevation?
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                if (deleteConfirm) {
                  deleteMarkerMut.mutate(deleteConfirm.id);
                  closeDialog();
                }
              }}
            >
              Delete
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
