import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link, useParams, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { useToast } from "@/hooks/use-toast";
import {
  ArrowLeft, Plus, FileText, Camera, ChevronRight, Trash2,
  MapPin, User, UserCheck, AlertTriangle, CheckCircle2, Archive,
  ChevronDown, FileDown, Eye, Settings, X, ImageDown
} from "lucide-react";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import type { Project, Report, Defect, Photo } from "@shared/schema";
import { useState, useMemo } from "react";

const STANDARD_ELEVATIONS = [
  "North", "North East", "East", "South East",
  "South", "South West", "West", "North West",
];

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

// Helper to load an image as ArrayBuffer (for docx) or data URL (for pdf)
// Load compressed thumbnail for report exports (smaller file size)
async function loadImageBlob(filename: string): Promise<Blob | null> {
  try {
    const res = await fetch(`${API_BASE}/api/thumbs/${filename}`);
    return await res.blob();
  } catch {
    return null;
  }
}

async function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result as string);
    reader.readAsDataURL(blob);
  });
}

async function blobToArrayBuffer(blob: Blob): Promise<ArrayBuffer> {
  return blob.arrayBuffer();
}

// Compress an image blob via Canvas — returns a smaller ArrayBuffer (for Word export)
async function compressImageForExport(blob: Blob, maxWidth = 800, quality = 0.7): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const scale = Math.min(1, maxWidth / img.width);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext('2d')!;
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      canvas.toBlob(
        (result) => {
          if (result) result.arrayBuffer().then(resolve).catch(reject);
          else reject(new Error('Canvas toBlob failed'));
        },
        'image/jpeg',
        quality,
      );
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('Image load failed'));
    };
    img.src = url;
  });
}

// Compress an image blob via Canvas — returns a data URL string (for PDF export)
async function compressImageForPdfExport(blob: Blob, maxWidth = 800, quality = 0.7): Promise<string> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const scale = Math.min(1, maxWidth / img.width);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext('2d')!;
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('Image load failed'));
    };
    img.src = url;
  });
}

// Load the AFC logo
async function loadAfcLogo(): Promise<{ dataUrl: string; buffer: ArrayBuffer } | null> {
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

// Elevation code to full name mapping
const ELEVATION_NAMES: Record<string, string> = {
  N: "North", S: "South", E: "East", W: "West",
  NE: "North East", NW: "North West", SE: "South East", SW: "South West",
};

// Derive location string from defect UID
function deriveLocation(uid: string): string {
  const parts = uid.split("-");
  if (parts.length >= 5) {
    const elevName = ELEVATION_NAMES[parts[0]] || parts[0];
    const drop = parseInt(parts[1], 10);
    const level = parseInt(parts[2], 10);
    return `${elevName} Elevation, Drop ${drop}, Level ${level}`;
  }
  if (parts.length >= 2) {
    const drop = parseInt(parts[0], 10);
    const level = parseInt(parts[1], 10);
    return `Drop ${drop}, Level ${level}`;
  }
  return "";
}

// Format date for report
function formatReportDate(dateStr?: string): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });
}

export default function ReportDetail() {
  const { projectId, reportId } = useParams<{ projectId: string; reportId: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [generating, setGenerating] = useState<string | null>(null);
  const [editOpen, setEditOpen] = useState(false);
  const [editForm, setEditForm] = useState<Record<string, string>>({});
  const [customElevation, setCustomElevation] = useState("");

  const { data: project } = useQuery<Project>({
    queryKey: ["/api/projects", projectId],
  });

  const { data: report, isLoading: reportLoading } = useQuery<Report>({
    queryKey: ["/api/reports", reportId],
  });

  const { data: defects, isLoading: defectsLoading } = useQuery<Defect[]>({
    queryKey: [`/api/reports/${reportId}/defects`],
  });

  const openEditDialog = () => {
    if (!report) return;
    setEditForm({
      inspectionNumber: report.inspectionNumber || "",
      inspectionDate: report.inspectionDate || "",
      revision: report.revision || "01",
      locationsCovered: report.locationsCovered || "",
      elevations: (report as any).elevations || project?.elevations || "[]",
      attendees: report.attendees || "[]",
    });
    setCustomElevation("");
    setEditOpen(true);
  };

  const updateReportMutation = useMutation({
    mutationFn: async (data: Record<string, string>) => {
      const res = await apiRequest("PATCH", `/api/reports/${reportId}`, data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/reports", reportId] });
      setEditOpen(false);
      toast({ title: "Report updated" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: async (defectId: number) => {
      await apiRequest("DELETE", `/api/defects/${defectId}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/defects`] });
      toast({ title: "Defect deleted" });
    },
  });

  // Sort by: Drop (asc), Level (desc/highest first), WorkType (grouped), Number (asc)
  const sortByUid = (a: Defect, b: Defect): number => {
    const pa = a.uid.split("-");
    const pb = b.uid.split("-");
    const aHasElev = pa.length >= 5;
    const bHasElev = pb.length >= 5;
    const aDrop = parseInt(pa[aHasElev ? 1 : 0] || "0", 10);
    const bDrop = parseInt(pb[bHasElev ? 1 : 0] || "0", 10);
    if (aDrop !== bDrop) return aDrop - bDrop;
    const aLevel = parseInt(pa[aHasElev ? 2 : 1] || "0", 10);
    const bLevel = parseInt(pb[bHasElev ? 2 : 1] || "0", 10);
    if (aLevel !== bLevel) return bLevel - aLevel;
    const aWork = pa[aHasElev ? 3 : 2] || "";
    const bWork = pb[bHasElev ? 3 : 2] || "";
    if (aWork !== bWork) return aWork.localeCompare(bWork);
    const aNum = parseInt(pa[aHasElev ? 4 : 3] || "0", 10);
    const bNum = parseInt(pb[bHasElev ? 4 : 3] || "0", 10);
    return aNum - bNum;
  };

  const activeDefects = useMemo(() =>
    (defects?.filter((d) => d.status !== "complete" && d.recordType !== "observation") ?? []).sort(sortByUid), [defects]);
  const activeObservations = useMemo(() =>
    (defects?.filter((d) => d.status !== "complete" && d.recordType === "observation") ?? []).sort(sortByUid), [defects]);
  const completedAll = useMemo(() =>
    (defects?.filter((d) => d.status === "complete") ?? []).sort((a, b) => {
      const dateCompare = (b.dateClosed ?? "").localeCompare(a.dateClosed ?? "");
      if (dateCompare !== 0) return dateCompare;
      return sortByUid(a, b);
    }),
    [defects]);

  // ==================== SHARED: Render a single defect page (PDF) ====================
  const renderDefectPagePdf = async (
    doc: any, defect: any, margin: number, contentWidth: number, pageWidth: number, pageHeight: number,
    addHeader: () => void, addFooter: () => void,
    autoTable: any, DARK_TEXT: readonly number[], CAPTION_BLUE: readonly number[],
  ) => {
    doc.addPage();
    addHeader();
    addFooter();
    let y = 20;

    const slotOrder = ["wip1", "wip2", "wip3", "wip4", "wip5", "complete"];
    const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", wip4: "WIP 4", wip5: "WIP 5", complete: "Complete" };

    const typeLabel = (defect.recordType === "observation") ? "OBSERVATION" : "DEFECT";
    const typeBgColor = (defect.recordType === "observation") ? [59, 130, 246] : [217, 119, 6];

    doc.setFontSize(8);
    doc.setFont("helvetica", "bold");
    const badgeW = doc.getTextWidth(typeLabel) + 6;
    doc.setFillColor(typeBgColor[0], typeBgColor[1], typeBgColor[2]);
    doc.roundedRect(margin, y - 4, badgeW, 6, 1.5, 1.5, "F");
    doc.setTextColor(255, 255, 255);
    doc.text(typeLabel, margin + 3, y);

    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text(defect.uid, margin + badgeW + 4, y);

    const statusLabel = defect.status === "complete" ? "COMPLETE" : "OPEN";
    const statusColor = defect.status === "complete" ? [34, 139, 34] : [200, 150, 0];
    doc.setFontSize(10);
    doc.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
    doc.text(statusLabel, pageWidth - margin - doc.getTextWidth(statusLabel), y);
    doc.setTextColor(0);
    y += 5;

    doc.setFontSize(9);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(100);
    doc.text(deriveLocation(defect.uid), margin, y);
    y += 5;

    autoTable(doc, {
      startY: y,
      body: [
        ["Date Opened", defect.dateOpened],
        ["Date Completed", defect.dateClosed || "\u2014"],
        ["Observation", defect.comment],
        ["Action Required", defect.actionRequired],
        ["Assigned To", defect.assignedTo],
        ["Due Date", defect.dueDate],
        ["Verification Method", defect.verificationMethod],
        ["Verification Person", defect.verificationPerson],
      ],
      margin: { left: margin, right: margin },
      styles: { fontSize: 8, cellPadding: 2.5 },
      columnStyles: {
        0: { fontStyle: "bold", cellWidth: 38 },
        1: { cellWidth: contentWidth - 38 },
      },
      theme: "plain",
      didDrawCell: (cellData: any) => {
        if (cellData.column.index === 0) {
          doc.setDrawColor(220);
          doc.line(
            cellData.cell.x,
            cellData.cell.y + cellData.cell.height,
            cellData.cell.x + contentWidth,
            cellData.cell.y + cellData.cell.height
          );
        }
      },
    });

    y = (doc as any).lastAutoTable.finalY + 6;

    if (defect.photos && defect.photos.length > 0) {
      doc.setFontSize(9);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("Photos", margin, y);
      y += 4;

      const sortedPhotos = slotOrder.map((s: string) => defect.photos.find((p: any) => p.slot === s)).filter(Boolean);
      const imgW = 85;
      const imgH = imgW * 0.75;

      for (let i = 0; i < sortedPhotos.length; i++) {
        const photo = sortedPhotos[i];
        const col = i % 2;
        if (i > 0 && col === 0) y += imgH + 12;
        if (col === 0 && y + imgH + 12 > pageHeight - 20) {
          doc.addPage();
          addHeader();
          addFooter();
          y = 20;
        }
        const x = margin + col * (imgW + 6);

        const blob = await loadImageBlob(photo.filename);
        if (blob) {
          let dataUrl: string;
          try {
            dataUrl = await compressImageForPdfExport(blob, 800, 0.7);
          } catch {
            dataUrl = await blobToDataUrl(blob);
          }
          doc.addImage(dataUrl, "JPEG", x, y, imgW, imgH);
        } else {
          doc.setDrawColor(180);
          doc.rect(x, y, imgW, imgH);
          doc.setFontSize(7);
          doc.text("[Photo unavailable]", x + 4, y + imgH / 2);
        }

        doc.setFontSize(7);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(100);
        const photoLabel = slotLabels[photo.slot] || photo.slot;
        const photoCaption = photo.caption ? ` — ${photo.caption}` : "";
        doc.text(photoLabel + photoCaption, x, y + imgH + 3);
        doc.setTextColor(0);
      }
    }
  };

  // ==================== PDF GENERATION (AFC Template) ====================
  const handleGeneratePdf = async () => {
    setGenerating("pdf");
    try {
      const res = await apiRequest("GET", `/api/reports/${reportId}/report-data`);
      const data = await res.json();

      const { default: jsPDF } = await import("jspdf");
      const autoTable = (await import("jspdf-autotable")).default;

      const doc = new jsPDF("p", "mm", "a4");
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 12.7;
      const contentWidth = pageWidth - margin * 2;

      const logo = await loadAfcLogo();

      const DARK_BLUE = [10, 29, 48] as const;
      const TEAL = [0, 235, 230] as const;
      const ACCENT_BLUE = [69, 176, 225] as const;
      const DARK_TEXT = [58, 58, 58] as const;
      const CAPTION_BLUE = [14, 40, 65] as const;

      // ======= COVER PAGE =======
      doc.setFillColor(...TEAL);
      doc.rect(0, 0, 25, pageHeight, "F");

      if (logo) {
        doc.addImage(logo.dataUrl, "PNG", pageWidth - 60, 15, 45, 14.4);
      }

      const titleX = 35;
      let ty = pageHeight * 0.35;

      doc.setFontSize(42);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_BLUE);
      const titleLines = doc.splitTextToSize(data.project.name.toUpperCase(), contentWidth - 25);
      doc.text(titleLines, titleX, ty);
      ty += titleLines.length * 16 + 4;

      doc.setFontSize(26);
      doc.setFont("helvetica", "normal");
      doc.text("SITE VISIT REPORT", titleX, ty);
      ty += 14;

      doc.setFontSize(14);
      const rev = data.report.revision || "01";
      doc.text(`Revision ${rev}`, titleX, ty);
      ty += 20;

      doc.setFontSize(12);
      doc.setTextColor(100);
      doc.text(formatReportDate(new Date().toISOString()), titleX, ty);
      ty += 16;

      doc.setFontSize(11);
      doc.setTextColor(...DARK_TEXT);
      doc.text("Angel Façade Consulting", titleX, ty); ty += 6;
      doc.text(`${data.project.inspector} | 0407 759 590`, titleX, ty); ty += 6;
      const afcRef = data.project.afcReference || "AFC-24XXX";
      doc.text(afcRef, titleX, ty);

      // ======= HEADER for subsequent pages =======
      const addHeader = () => {
        doc.setFontSize(8);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...DARK_TEXT);
        doc.text(data.project.name.toUpperCase(), margin, 10);
        doc.setFont("helvetica", "normal");
        doc.text("ANGEL FAÇADE CONSULTING", pageWidth - margin, 10, { align: "right" });
        doc.setDrawColor(...ACCENT_BLUE);
        doc.setLineWidth(0.5);
        doc.line(margin, 12, pageWidth - margin, 12);
      };

      const addFooter = () => {
        const pn = doc.internal.pages.length - 1;
        doc.setFontSize(7);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(150);
        doc.text(`Page ${pn}`, pageWidth / 2, pageHeight - 8, { align: "center" });
        doc.text(afcRef, pageWidth - margin, pageHeight - 8, { align: "right" });
      };

      // ======= SECTION 1 - INTRODUCTION =======
      doc.addPage();
      addHeader();
      addFooter();
      let y = 20;

      doc.setFontSize(18);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("1. Introduction", margin, y);
      y += 10;

      doc.setFontSize(14);
      doc.text("1.1 General", margin, y);
      y += 8;
      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(60);
      const generalText = `Angel Façade Consulting (AFC) was engaged by ${data.project.client} to carry out a site visit inspection of the facade at ${data.project.address}.`;
      const genLines = doc.splitTextToSize(generalText, contentWidth);
      doc.text(genLines, margin, y);
      y += genLines.length * 4.5 + 6;

      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("1.2 Inspection", margin, y);
      y += 8;

      // Build inspection table rows from report data
      const inspRows: string[][] = [];
      if (data.report.inspectionDate) inspRows.push(["Date", formatReportDate(data.report.inspectionDate)]);
      else inspRows.push(["Date", formatReportDate(new Date().toISOString())]);
      if (data.report.inspectionNumber) inspRows.push(["Inspection Number", data.report.inspectionNumber]);
      inspRows.push(["Inspector", data.project.inspector]);
      inspRows.push(["Locations covered", data.report.locationsCovered || data.project.address]);
      inspRows.push(["Client", data.project.client]);

      try {
        const attendees = JSON.parse(data.report.attendees || "[]");
        if (attendees.length > 0) {
          const attendeeStr = attendees.map((a: any) => `${a.name} (${a.company})`).join(", ");
          inspRows.push(["Attendees", attendeeStr]);
        }
      } catch {}

      autoTable(doc, {
        startY: y,
        body: inspRows,
        margin: { left: margin, right: margin },
        styles: { fontSize: 9, cellPadding: 3 },
        columnStyles: {
          0: { fontStyle: "bold", cellWidth: 35 },
          1: { cellWidth: contentWidth - 35 },
        },
        theme: "plain",
        didDrawCell: (cellData: any) => {
          if (cellData.row.index < 4) {
            doc.setDrawColor(220);
            doc.line(
              cellData.cell.x,
              cellData.cell.y + cellData.cell.height,
              cellData.cell.x + contentWidth,
              cellData.cell.y + cellData.cell.height
            );
          }
        },
      });

      y = (doc as any).lastAutoTable.finalY + 10;

      // ======= SECTION 2 - DEFECT REGISTER & RECTIFICATION LOG =======
      const sortByUidExport = (a: any, b: any): number => {
        const pa = a.uid.split("-");
        const pb = b.uid.split("-");
        const aHasElev = pa.length >= 5;
        const bHasElev = pb.length >= 5;
        const aDrop = parseInt(pa[aHasElev ? 1 : 0] || "0", 10);
        const bDrop = parseInt(pb[bHasElev ? 1 : 0] || "0", 10);
        if (aDrop !== bDrop) return aDrop - bDrop;
        const aLevel = parseInt(pa[aHasElev ? 2 : 1] || "0", 10);
        const bLevel = parseInt(pb[bHasElev ? 2 : 1] || "0", 10);
        if (aLevel !== bLevel) return bLevel - aLevel;
        const aWork = pa[aHasElev ? 3 : 2] || "";
        const bWork = pb[bHasElev ? 3 : 2] || "";
        if (aWork !== bWork) return aWork.localeCompare(bWork);
        const aNum = parseInt(pa[aHasElev ? 4 : 3] || "0", 10);
        const bNum = parseInt(pb[bHasElev ? 4 : 3] || "0", 10);
        return aNum - bNum;
      };

      const allDefects = [...(data.defects || [])].sort(sortByUidExport);
      const defectsOnly = allDefects.filter((d: any) => d.recordType !== "observation");
      const observationsOnly = allDefects.filter((d: any) => d.recordType === "observation");
      const openData = defectsOnly.filter((d: any) => d.status !== "complete").sort(sortByUidExport);
      const completedData = defectsOnly.filter((d: any) => d.status === "complete").sort((a: any, b: any) => {
        const dc = (b.dateClosed ?? "").localeCompare(a.dateClosed ?? "");
        return dc !== 0 ? dc : sortByUidExport(a, b);
      });

      // Inspection-scoped filtering: only generate detail pages for defects touched in this inspection
      const reportCreatedAt = data.report.createdAt;
      const allHaveNullUpdatedAt = allDefects.every((d: any) => !d.updatedAt);
      const isCurrentInspection = (d: any) => allHaveNullUpdatedAt || (d.updatedAt && d.updatedAt >= reportCreatedAt);
      const openDetailData = openData.filter(isCurrentInspection);
      const completedDetailData = completedData.filter(isCurrentInspection);
      const observationsDetailData = observationsOnly.filter(isCurrentInspection);

      doc.addPage();
      addHeader();
      addFooter();
      y = 20;

      doc.setFontSize(18);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("2. Defect Register & Rectification Log", margin, y);
      y += 8;

      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(60);
      doc.text("Based on our observations, we recommend the following actions.", margin, y);
      y += 8;

      // Overdue detection: past due date AND not complete, evaluated against inspection date
      const inspectionDate = new Date(data.report.inspectionDate || data.report.createdAt);
      const isOverdue = (defect: any): boolean => {
        if (defect.status === "complete") return false;
        if (!defect.dueDate) return false;
        return new Date(defect.dueDate) < inspectionDate;
      };

      // Sort overdue rows to top, preserving existing sort within each group
      const overdueRows = allDefects.filter(isOverdue);
      const onTimeRows = allDefects.filter((d: any) => !isOverdue(d));
      const orderedForSummary = [...overdueRows, ...onTimeRows];

      const summaryHead = [["ID", "Type", "Location", "Observation", "Action Required", "By Date", "Status"]];
      const summaryBody = orderedForSummary.map((d: any) => [
        d.uid,
        d.recordType === "observation" ? "Obs" : "Defect",
        deriveLocation(d.uid),
        d.comment.length > 45 ? d.comment.substring(0, 42) + "..." : d.comment,
        d.actionRequired.length > 35 ? d.actionRequired.substring(0, 32) + "..." : d.actionRequired,
        d.dueDate || "—",
        d.status === "complete" ? "Complete" : "Open",
      ]);

      // Track which rows are overdue for red badge rendering
      const overdueFlags = orderedForSummary.map(isOverdue);

      if (summaryBody.length > 0) {
        autoTable(doc, {
          startY: y,
          head: summaryHead,
          body: summaryBody,
          margin: { left: margin, right: margin },
          styles: { fontSize: 6.5, cellPadding: 1.5, overflow: "linebreak", valign: "top" },
          headStyles: {
            fillColor: [255, 255, 255],
            textColor: [...CAPTION_BLUE],
            fontStyle: "bold",
            lineWidth: { bottom: 0.5 },
            lineColor: [...DARK_TEXT],
          },
          columnStyles: {
            0: { cellWidth: 20 },
            1: { cellWidth: 12 },
            2: { cellWidth: 24 },
            3: { cellWidth: 42 },
            4: { cellWidth: 34 },
            5: { cellWidth: 18 },
            6: { cellWidth: 16 },
          },
          didDrawPage: () => { addHeader(); addFooter(); },
          rowPageBreak: "auto",
          didDrawCell: (cellData: any) => {
            // Append " • OVERDUE" in red bold to the Status cell for overdue rows
            if (cellData.section === "body" && cellData.column.index === 6 && overdueFlags[cellData.row.index]) {
              const cell = cellData.cell;
              doc.setFont("helvetica", "bold");
              doc.setFontSize(5.5);
              doc.setTextColor(192, 57, 43); // #C0392B red
              doc.text(" \u2022 OVERDUE", cell.x + 1, cell.y + cell.height - 1.5);
              doc.setTextColor(0);
              doc.setFont("helvetica", "normal");
            }
          },
        });
      } else {
        doc.setFontSize(9);
        doc.text("No items recorded.", margin, y);
      }

      if (openDetailData.length > 0) {
        for (const defect of openDetailData) {
          await renderDefectPagePdf(doc, defect, margin, contentWidth, pageWidth, pageHeight, addHeader, addFooter, autoTable, DARK_TEXT, CAPTION_BLUE);
        }
      }

      if (completedData.length > 0) {
        doc.addPage();
        addHeader();
        addFooter();
        y = 20;

        doc.setFontSize(18);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...DARK_TEXT);
        doc.text("3. Completed Works Summary", margin, y);
        y += 7;
        doc.setFontSize(9);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100);
        doc.text("All defects that have been rectified and verified.", margin, y);

        for (const defect of completedDetailData) {
          await renderDefectPagePdf(doc, defect, margin, contentWidth, pageWidth, pageHeight, addHeader, addFooter, autoTable, DARK_TEXT, CAPTION_BLUE);
        }
      }

      if (observationsOnly.length > 0) {
        doc.addPage();
        addHeader();
        addFooter();
        y = 20;

        const obsSection = completedData.length > 0 ? "4" : "3";
        doc.setFontSize(18);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...DARK_TEXT);
        doc.text(`${obsSection}. Observations`, margin, y);
        y += 7;
        doc.setFontSize(9);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100);
        doc.text("General observations noted during inspection.", margin, y);

        for (const defect of observationsDetailData) {
          await renderDefectPagePdf(doc, defect, margin, contentWidth, pageWidth, pageHeight, addHeader, addFooter, autoTable, DARK_TEXT, CAPTION_BLUE);
        }
      }

      doc.save(`${data.project.name.replace(/[^a-zA-Z0-9]/g, "_")}_SVR.pdf`);
      toast({ title: "PDF report downloaded" });
    } catch (err) {
      console.error(err);
      toast({ title: "Error generating PDF", variant: "destructive" });
    } finally {
      setGenerating(null);
    }
  };

  // ==================== WORD GENERATION (AFC Template) ====================
  const handleGenerateWord = async () => {
    setGenerating("word");
    try {
      const res = await apiRequest("GET", `/api/reports/${reportId}/report-data`);
      const data = await res.json();

      const docxLib = await import("docx");
      const { saveAs } = await import("file-saver");
      const {
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        ImageRun, PageBreak, HeadingLevel, AlignmentType,
        WidthType, BorderStyle, ShadingType, TableLayoutType,
        Header, Footer, PageNumber, NumberFormat, Tab, TabStopType, TabStopPosition,
      } = docxLib;

      const logo = await loadAfcLogo();

      const sortUid = (a: any, b: any): number => {
        const pa = a.uid.split("-"); const pb = b.uid.split("-");
        const ae = pa.length >= 5; const be = pb.length >= 5;
        const ad = parseInt(pa[ae?1:0]||"0",10); const bd = parseInt(pb[be?1:0]||"0",10);
        if (ad !== bd) return ad - bd;
        const al = parseInt(pa[ae?2:1]||"0",10); const bl = parseInt(pb[be?2:1]||"0",10);
        if (al !== bl) return bl - al;
        const aw = pa[ae?3:2]||""; const bw = pb[be?3:2]||"";
        if (aw !== bw) return aw.localeCompare(bw);
        return parseInt(pa[ae?4:3]||"0",10) - parseInt(pb[be?4:3]||"0",10);
      };

      const allDefects = [...(data.defects || [])].sort(sortUid);
      const defectsOnly = allDefects.filter((d: any) => d.recordType !== "observation");
      const observationsOnly = allDefects.filter((d: any) => d.recordType === "observation");
      const openData = defectsOnly.filter((d: any) => d.status !== "complete").sort(sortUid);
      const completedData = defectsOnly.filter((d: any) => d.status === "complete").sort((a: any, b: any) => {
        const dc = (b.dateClosed ?? "").localeCompare(a.dateClosed ?? "");
        return dc !== 0 ? dc : sortUid(a, b);
      });

      // Inspection-scoped filtering: only generate detail pages for defects touched in this inspection
      const reportCreatedAt = data.report.createdAt;
      const allHaveNullUpdatedAt = allDefects.every((d: any) => !d.updatedAt);
      const isCurrentInspection = (d: any) => allHaveNullUpdatedAt || (d.updatedAt && d.updatedAt >= reportCreatedAt);
      const openDetailData = openData.filter(isCurrentInspection);
      const completedDetailData = completedData.filter(isCurrentInspection);
      const observationsDetailData = observationsOnly.filter(isCurrentInspection);

      const slotOrder = ["wip1", "wip2", "wip3", "complete"];
      const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", complete: "Complete" };

      const DARK_BLUE = "0A1D30";
      const TEAL = "00EBE6";
      const ACCENT_BLUE = "45B0E1";
      const DARK_TEXT = "3A3A3A";
      const CAPTION_BLUE = "0E2841";

      const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
      const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
      const accentBottomBorder = { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE };

      const reportDate = formatReportDate(new Date().toISOString());
      const rev = data.report.revision || "01";
      const afcRef = data.project.afcReference || "AFC-24XXX";

      const headerTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [new Paragraph({
                  children: [new TextRun({ text: data.project.name.toUpperCase(), size: 16, font: "Aptos", bold: true, color: DARK_TEXT })],
                })],
                borders: { top: noBorder, left: noBorder, right: noBorder, bottom: accentBottomBorder },
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [new Paragraph({
                  children: [new TextRun({ text: "ANGEL FAÇADE CONSULTING", size: 16, font: "Aptos", color: DARK_TEXT })],
                  alignment: AlignmentType.RIGHT,
                })],
                borders: { top: noBorder, left: noBorder, right: noBorder, bottom: accentBottomBorder },
              }),
            ],
          }),
        ],
      });

      const defaultHeader = new Header({ children: [headerTable] });

      const defaultFooter = new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: afcRef, size: 14, font: "Aptos", color: "999999" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "Page ", size: 14, font: "Aptos", color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 14, font: "Aptos", color: "999999" }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          }),
        ],
      });

      // ======= COVER PAGE =======
      const coverChildren: any[] = [];

      if (logo) {
        let logoBuffer: ArrayBuffer;
        try {
          logoBuffer = await compressImageForExport(new Blob([logo.buffer]), 360, 0.8);
        } catch {
          logoBuffer = logo.buffer;
        }
        coverChildren.push(new Paragraph({
          children: [
            new ImageRun({ data: logoBuffer, transformation: { width: 180, height: 58 }, type: "jpg" }),
          ],
          alignment: AlignmentType.RIGHT,
          spacing: { after: 600 },
        }));
      } else {
        coverChildren.push(new Paragraph({ spacing: { after: 600 } }));
      }

      coverChildren.push(new Paragraph({ spacing: { before: 1200 } }));

      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: data.project.name.toUpperCase(), bold: true, size: 72, font: "Aptos", color: DARK_BLUE })],
        spacing: { after: 100 },
      }));

      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: "SITE VISIT REPORT", size: 52, font: "Aptos", color: DARK_BLUE, smallCaps: true })],
        spacing: { after: 100 },
      }));

      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: `Revision ${rev}`, size: 28, font: "Aptos", color: DARK_TEXT })],
        spacing: { after: 600 },
      }));

      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: reportDate, size: 24, font: "Aptos", color: DARK_TEXT })],
        spacing: { after: 400 },
      }));

      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: "Angel Façade Consulting", size: 22, font: "Aptos", color: DARK_TEXT })],
        spacing: { after: 40 },
      }));
      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: `${data.project.inspector} | 0407 759 590`, size: 22, font: "Aptos", color: DARK_TEXT })],
        spacing: { after: 40 },
      }));
      coverChildren.push(new Paragraph({
        children: [new TextRun({ text: afcRef, size: 22, font: "Aptos", color: DARK_TEXT })],
        spacing: { after: 200 },
      }));

      // ======= SECTION 1 - INTRODUCTION =======
      const introChildren: any[] = [];

      introChildren.push(new Paragraph({
        children: [new TextRun({ text: "Introduction", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 200 },
      }));

      introChildren.push(new Paragraph({
        children: [new TextRun({ text: "General", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_2,
        spacing: { after: 100 },
      }));

      introChildren.push(new Paragraph({
        children: [new TextRun({
          text: `Angel Façade Consulting (AFC) was engaged by ${data.project.client} to carry out a site visit inspection of the facade at ${data.project.address}.`,
          size: 20, font: "Aptos",
        })],
        spacing: { after: 200 },
      }));

      introChildren.push(new Paragraph({
        children: [new TextRun({ text: "Inspection", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_2,
        spacing: { after: 100 },
      }));

      // Inspection table from report data
      const inspectionData: string[][] = [];
      if (data.report.inspectionDate) inspectionData.push(["Date", formatReportDate(data.report.inspectionDate)]);
      else inspectionData.push(["Date", reportDate]);
      if (data.report.inspectionNumber) inspectionData.push(["Inspection Number", data.report.inspectionNumber]);
      inspectionData.push(["Inspector", data.project.inspector]);
      inspectionData.push(["Locations covered", data.report.locationsCovered || data.project.address]);
      inspectionData.push(["Client", data.project.client]);

      try {
        const attendees = JSON.parse(data.report.attendees || "[]");
        if (attendees.length > 0) {
          attendees.forEach((a: any) => {
            inspectionData.push([a.company || a.name, a.name]);
          });
        }
      } catch {}

      introChildren.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: inspectionData.map(([label, value]) =>
          new TableRow({
            children: [
              new TableCell({
                width: { size: 25, type: WidthType.PERCENTAGE },
                children: [new Paragraph({
                  children: [new TextRun({ text: label, bold: true, size: 20, font: "Aptos", color: DARK_TEXT })],
                  spacing: { before: 60, after: 60 },
                })],
                borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
              }),
              new TableCell({
                width: { size: 75, type: WidthType.PERCENTAGE },
                children: [new Paragraph({
                  children: [new TextRun({ text: value, size: 20, font: "Aptos" })],
                  spacing: { before: 60, after: 60 },
                })],
                borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
              }),
            ],
          })
        ),
      }));

      introChildren.push(new Paragraph({ spacing: { before: 200 } }));

      // ======= HELPER: Build 1 defect page for Word =======
      const buildWordDefectPage = async (defect: any): Promise<any[]> => {
        const elements: any[] = [];

        const isObs = defect.recordType === "observation";
        elements.push(new Paragraph({
          children: [
            new TextRun({
              text: isObs ? " OBSERVATION " : " DEFECT ",
              bold: true, size: 16, font: "Aptos", color: "FFFFFF",
              shading: { type: ShadingType.SOLID, color: isObs ? "3B82F6" : "D97706" },
            }),
            new TextRun({ text: "  " }),
            new TextRun({ text: defect.uid, bold: true, size: 28, font: "Aptos", color: CAPTION_BLUE }),
            new TextRun({ text: "    " }),
            new TextRun({
              text: defect.status === "complete" ? "COMPLETE" : "OPEN",
              bold: true, size: 20,
              color: defect.status === "complete" ? "228B22" : "C89600",
            }),
          ],
          spacing: { after: 40 },
        }));

        elements.push(new Paragraph({
          children: [new TextRun({ text: deriveLocation(defect.uid), size: 18, font: "Aptos", color: "666666" })],
          spacing: { after: 200 },
        }));

        const infoRows = [
          ["Date Opened", defect.dateOpened],
          ["Date Completed", defect.dateClosed || "\u2014"],
          ["Observation", defect.comment],
          ["Action Required", defect.actionRequired],
          ["Assigned To", defect.assignedTo],
          ["Due Date", defect.dueDate],
          ["Verification Method", defect.verificationMethod],
          ["Verification Person", defect.verificationPerson],
        ];

        const bottomBorder = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };

        elements.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: infoRows.map(([label, value]) =>
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 28, type: WidthType.PERCENTAGE },
                  children: [new Paragraph({
                    children: [new TextRun({ text: label, bold: true, size: 18, font: "Aptos" })],
                    spacing: { before: 50, after: 50 },
                  })],
                  borders: { top: noBorder, left: noBorder, right: noBorder, bottom: bottomBorder },
                }),
                new TableCell({
                  width: { size: 72, type: WidthType.PERCENTAGE },
                  children: [new Paragraph({
                    children: [new TextRun({ text: value, size: 18, font: "Aptos" })],
                    spacing: { before: 50, after: 50 },
                  })],
                  borders: { top: noBorder, left: noBorder, right: noBorder, bottom: bottomBorder },
                }),
              ],
            })
          ),
        }));

        elements.push(new Paragraph({ spacing: { before: 150 } }));

        if (defect.photos && defect.photos.length > 0) {
          elements.push(new Paragraph({
            children: [new TextRun({ text: "Photos", bold: true, size: 20, font: "Aptos" })],
            spacing: { before: 100, after: 80 },
          }));

          const sortedPhotos = slotOrder
            .map((s: string) => defect.photos.find((p: any) => p.slot === s))
            .filter(Boolean);

          for (let i = 0; i < sortedPhotos.length; i += 2) {
            const photoCells: any[] = [];

            for (let j = 0; j < 2; j++) {
              const photo = sortedPhotos[i + j];
              if (photo) {
                const blob = await loadImageBlob(photo.filename);
                const cellChildren: any[] = [];

                if (blob) {
                  let buffer: ArrayBuffer;
                  try {
                    buffer = await compressImageForExport(blob, 800, 0.7);
                  } catch {
                    buffer = await blobToArrayBuffer(blob);
                  }
                  cellChildren.push(new Paragraph({
                    children: [
                      new ImageRun({ data: buffer, transformation: { width: 241, height: 181 }, type: "jpg" }),
                    ],
                    alignment: AlignmentType.CENTER,
                  }));
                }

                const captionText = photo.caption ? ` \u2014 ${photo.caption}` : "";
                cellChildren.push(new Paragraph({
                  children: [
                    new TextRun({
                      text: slotLabels[photo.slot] || photo.slot,
                      bold: true, size: 16, font: "Aptos",
                      color: photo.slot === "complete" ? "228B22" : "666666",
                    }),
                    ...(captionText ? [new TextRun({ text: captionText, size: 14, font: "Aptos", color: "888888" })] : []),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 40 },
                }));

                photoCells.push(new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  children: cellChildren,
                  borders: { top: noBorder, left: noBorder, right: noBorder, bottom: noBorder },
                }));
              } else {
                photoCells.push(new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  children: [new Paragraph("")],
                  borders: { top: noBorder, left: noBorder, right: noBorder, bottom: noBorder },
                }));
              }
            }

            elements.push(new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [new TableRow({ children: photoCells })],
            }));
            elements.push(new Paragraph({ spacing: { before: 60 } }));
          }
        }

        return elements;
      };

      // ======= SECTION 2 - DEFECT REGISTER =======
      const obsChildren: any[] = [];

      obsChildren.push(new Paragraph({
        children: [new TextRun({ text: "Defect Register & Rectification Log", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 100 },
      }));

      obsChildren.push(new Paragraph({
        children: [new TextRun({ text: "Based on our observations, we recommend the following actions.", size: 20, font: "Aptos" })],
        spacing: { after: 200 },
      }));

      const summaryHeaderLabels = ["ID", "Type", "Location", "Observation", "Action Required", "By Date", "Status"];
      const summaryHeaderWidths = [850, 600, 1200, 2800, 2300, 1000, 950];

      const summaryTableRows: any[] = [];

      summaryTableRows.push(new TableRow({
        tableHeader: true,
        children: summaryHeaderLabels.map((label, i) =>
          new TableCell({
            width: { size: summaryHeaderWidths[i], type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: label, size: 16, font: "Aptos", bold: true, color: CAPTION_BLUE })],
              spacing: { before: 30, after: 30 },
            })],
            borders: {
              top: noBorder, left: noBorder, right: noBorder,
              bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_TEXT },
            },
          })
        ),
      }));

      // Overdue detection for Word export
      const wordInspectionDate = new Date(data.report.inspectionDate || data.report.createdAt);
      const isOverdueWord = (defect: any): boolean => {
        if (defect.status === "complete") return false;
        if (!defect.dueDate) return false;
        return new Date(defect.dueDate) < wordInspectionDate;
      };

      // Sort overdue rows to top for summary table
      const wordOverdueRows = allDefects.filter(isOverdueWord);
      const wordOnTimeRows = allDefects.filter((d: any) => !isOverdueWord(d));
      const wordOrderedForSummary = [...wordOverdueRows, ...wordOnTimeRows];

      for (const defect of wordOrderedForSummary) {
        const statusText = defect.status === "complete" ? "Complete" : "Open";
        const typeText = defect.recordType === "observation" ? "Obs" : "Defect";
        const rowBorder = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
        const overdue = isOverdueWord(defect);
        const cellTexts = [
          defect.uid, typeText, deriveLocation(defect.uid),
          defect.comment.length > 60 ? defect.comment.substring(0, 57) + "..." : defect.comment,
          defect.actionRequired.length > 45 ? defect.actionRequired.substring(0, 42) + "..." : defect.actionRequired,
          defect.dueDate || "\u2014", statusText,
        ];
        summaryTableRows.push(new TableRow({
          children: cellTexts.map((text: string, i: number) =>
            new TableCell({
              width: { size: summaryHeaderWidths[i], type: WidthType.DXA },
              children: [new Paragraph({
                children: i === 6 && overdue
                  ? [
                      new TextRun({
                        text, size: 14, font: "Aptos",
                        color: "C89600",
                      }),
                      new TextRun({
                        text: " \u2022 OVERDUE", size: 14, font: "Aptos",
                        color: "C0392B", bold: true,
                      }),
                    ]
                  : [new TextRun({
                      text, size: 14, font: "Aptos",
                      color: i === 6 ? (statusText === "Complete" ? "228B22" : "C89600") : (i === 1 ? "666666" : undefined),
                      bold: i === 0,
                    })],
                spacing: { before: 25, after: 25 },
              })],
              borders: { top: rowBorder, left: noBorder, right: noBorder, bottom: rowBorder },
            })
          ),
        }));
      }

      if (allDefects.length > 0) {
        obsChildren.push(new Table({
          width: { size: 9700, type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          rows: summaryTableRows,
        }));
      } else {
        obsChildren.push(new Paragraph({
          children: [new TextRun({ text: "No items recorded.", size: 20, font: "Aptos", color: "999999", italics: true })],
        }));
      }

      for (let i = 0; i < openDetailData.length; i++) {
        if (i > 0 || defectsOnly.length > 0) {
          obsChildren.push(new Paragraph({ children: [new PageBreak()] }));
        }
        const defectEls = await buildWordDefectPage(openDetailData[i]);
        obsChildren.push(...defectEls);
      }

      // ======= SECTION 3 - COMPLETED WORKS =======
      const completedChildren: any[] = [];

      if (completedData.length > 0) {
        completedChildren.push(new Paragraph({
          children: [new TextRun({ text: "Completed Works Summary", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
          heading: HeadingLevel.HEADING_1,
          spacing: { after: 80 },
        }));
        completedChildren.push(new Paragraph({
          children: [new TextRun({ text: "All defects that have been rectified and verified.", size: 18, color: "666666", italics: true, font: "Aptos" })],
          spacing: { after: 200 },
        }));

        for (let i = 0; i < completedDetailData.length; i++) {
          if (i > 0) {
            completedChildren.push(new Paragraph({ children: [new PageBreak()] }));
          }
          const defectEls = await buildWordDefectPage(completedDetailData[i]);
          completedChildren.push(...defectEls);
        }
      }

      // ======= BUILD DOCUMENT =======
      const sectionProps = {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 720, right: 720, bottom: 720, left: 720 },
        },
        headers: { default: defaultHeader },
        footers: { default: defaultFooter },
      };

      const docSections: any[] = [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 },
              margin: { top: 851, right: 1440, bottom: 1440, left: 1440 },
            },
          },
          children: coverChildren,
        },
        { properties: sectionProps, children: introChildren },
        { properties: sectionProps, children: obsChildren },
      ];

      if (completedChildren.length > 0) {
        docSections.push({ properties: sectionProps, children: completedChildren });
      }

      if (observationsOnly.length > 0) {
        const obsOnlyChildren: any[] = [];
        obsOnlyChildren.push(new Paragraph({
          children: [new TextRun({ text: "Observations", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
          heading: HeadingLevel.HEADING_1,
          spacing: { after: 80 },
        }));
        obsOnlyChildren.push(new Paragraph({
          children: [new TextRun({ text: "General observations noted during inspection.", size: 18, color: "666666", italics: true, font: "Aptos" })],
          spacing: { after: 200 },
        }));
        for (let i = 0; i < observationsDetailData.length; i++) {
          if (i > 0) {
            obsOnlyChildren.push(new Paragraph({ children: [new PageBreak()] }));
          }
          const defectEls = await buildWordDefectPage(observationsDetailData[i]);
          obsOnlyChildren.push(...defectEls);
        }
        docSections.push({ properties: sectionProps, children: obsOnlyChildren });
      }

      const wordDoc = new Document({
        creator: "Angel Façade Consulting",
        title: `${data.project.name} — Site Visit Report`,
        styles: {
          default: {
            document: {
              run: { font: "Aptos", size: 20, color: DARK_TEXT },
            },
          },
        },
        sections: docSections,
      });

      const blob = await Packer.toBlob(wordDoc);
      const filename = `${new Date().toISOString().slice(0, 10).replace(/-/g, "")}-${afcRef}-${data.project.name.replace(/[^a-zA-Z0-9]/g, "_")}-SVR.docx`;
      saveAs(blob, filename);
      toast({ title: "Word report downloaded" });
    } catch (err) {
      console.error(err);
      toast({ title: "Error generating Word document", variant: "destructive" });
    } finally {
      setGenerating(null);
    }
  };

  if (reportLoading) {
    return (
      <div className="max-w-3xl mx-auto px-4 py-8">
        <div className="h-8 w-48 bg-muted animate-pulse rounded mb-4" />
        <div className="h-4 w-64 bg-muted animate-pulse rounded mb-8" />
        <div className="space-y-3">
          {[1, 2, 3].map((i) => (
            <div key={i} className="h-20 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      </div>
    );
  }

  if (!report || !project) return null;

  return (
    <div className="max-w-3xl mx-auto px-4 py-8">
      {/* Header */}
      <div className="mb-8">
        <Link href={`/projects/${projectId}`}>
          <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4">
            <ArrowLeft className="w-4 h-4" />
            Back to Project
          </button>
        </Link>
        <div className="flex items-start justify-between gap-4">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">
              {project.name}
            </h1>
            <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-sm text-muted-foreground">
              {report.inspectionNumber && (
                <span>Inspection #{report.inspectionNumber}</span>
              )}
              {report.inspectionDate && (
                <span>{formatReportDate(report.inspectionDate)}</span>
              )}
              {report.revision && (
                <span>Rev {report.revision}</span>
              )}
            </div>
            {report.locationsCovered && (
              <p className="text-sm text-muted-foreground mt-1">
                <MapPin className="w-3.5 h-3.5 inline mr-1" />
                {report.locationsCovered}
              </p>
            )}
          </div>
          <Button variant="ghost" size="icon" onClick={openEditDialog}>
            <Settings className="w-4 h-4" />
          </Button>
        </div>

        {/* Edit Report Dialog */}
        <Dialog open={editOpen} onOpenChange={setEditOpen}>
          <DialogContent className="max-h-[90vh] overflow-y-auto">
            <DialogHeader>
              <DialogTitle>Edit Report</DialogTitle>
            </DialogHeader>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                updateReportMutation.mutate(editForm);
              }}
              className="space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Inspection Number</Label>
                  <Input value={editForm.inspectionNumber || ""} onChange={(e) => setEditForm({ ...editForm, inspectionNumber: e.target.value })} />
                </div>
                <div>
                  <Label>Inspection Date</Label>
                  <Input type="date" value={editForm.inspectionDate || ""} onChange={(e) => setEditForm({ ...editForm, inspectionDate: e.target.value })} />
                </div>
              </div>
              <div>
                <Label>Revision</Label>
                <Input value={editForm.revision || ""} onChange={(e) => setEditForm({ ...editForm, revision: e.target.value })} />
              </div>
              <div>
                <Label>Locations Covered</Label>
                <Textarea value={editForm.locationsCovered || ""} onChange={(e) => setEditForm({ ...editForm, locationsCovered: e.target.value })} rows={2} />
              </div>
              {/* Elevations picker */}
              <div>
                <Label className="mb-2 block">Elevations</Label>
                <div className="grid grid-cols-2 gap-2">
                  {STANDARD_ELEVATIONS.map((elev) => {
                    const selected: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                    return (
                      <label key={elev} className="flex items-center gap-2 text-sm cursor-pointer">
                        <Checkbox
                          checked={selected.includes(elev)}
                          onCheckedChange={(checked) => {
                            const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                            if (checked) { sel.push(elev); } else { const idx = sel.indexOf(elev); if (idx !== -1) sel.splice(idx, 1); }
                            setEditForm({ ...editForm, elevations: JSON.stringify(sel) });
                          }}
                        />
                        {elev}
                      </label>
                    );
                  })}
                </div>
                {(() => {
                  const selected: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                  const custom = selected.filter((e) => !STANDARD_ELEVATIONS.includes(e));
                  return custom.length > 0 ? (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {custom.map((c) => (
                        <span key={c} className="inline-flex items-center gap-1 px-2 py-0.5 bg-accent rounded text-xs">
                          {c}
                          <button type="button" onClick={() => {
                            const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                            setEditForm({ ...editForm, elevations: JSON.stringify(sel.filter((e) => e !== c)) });
                          }} className="text-muted-foreground hover:text-destructive"><X className="w-3 h-3" /></button>
                        </span>
                      ))}
                    </div>
                  ) : null;
                })()}
                <div className="flex gap-2 mt-2">
                  <Input placeholder="Add custom elevation..." value={customElevation} onChange={(e) => setCustomElevation(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && customElevation.trim()) {
                        e.preventDefault();
                        const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                        if (!sel.includes(customElevation.trim())) { sel.push(customElevation.trim()); setEditForm({ ...editForm, elevations: JSON.stringify(sel) }); }
                        setCustomElevation("");
                      }
                    }}
                  />
                  <Button type="button" variant="outline" size="sm" className="shrink-0" onClick={() => {
                    if (customElevation.trim()) {
                      const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                      if (!sel.includes(customElevation.trim())) { sel.push(customElevation.trim()); setEditForm({ ...editForm, elevations: JSON.stringify(sel) }); }
                      setCustomElevation("");
                    }
                  }}><Plus className="w-4 h-4" /></Button>
                </div>
              </div>
              <div>
                <Label>Attendees</Label>
                <div className="space-y-2 mt-1">
                  {(() => {
                    try {
                      return (JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[]).map((attendee, index) => (
                        <div key={index} className="flex items-center gap-2">
                          <Input
                            placeholder="Name"
                            value={attendee.name}
                            onChange={(e) => {
                              const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                              attendees[index].name = e.target.value;
                              setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                            }}
                          />
                          <Input
                            placeholder="Company / Role"
                            value={attendee.company}
                            onChange={(e) => {
                              const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                              attendees[index].company = e.target.value;
                              setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                            }}
                          />
                          <Button type="button" variant="ghost" size="icon" className="shrink-0" onClick={() => {
                            const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                            attendees.splice(index, 1);
                            setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                          }}>
                            <X className="w-4 h-4" />
                          </Button>
                        </div>
                      ));
                    } catch { return null; }
                  })()}
                  <Button type="button" variant="outline" size="sm" onClick={() => {
                    try {
                      const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                      attendees.push({ name: "", company: "" });
                      setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                    } catch {}
                  }}>
                    <Plus className="w-4 h-4 mr-1" />
                    Add Attendee
                  </Button>
                </div>
              </div>
              <Button type="submit" className="w-full" disabled={updateReportMutation.isPending}>
                {updateReportMutation.isPending ? "Saving..." : "Save Changes"}
              </Button>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-4 gap-3 mb-6">
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold">{defects?.length ?? 0}</div>
          <div className="text-xs text-muted-foreground">Total</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-amber-600">{activeDefects.length}</div>
          <div className="text-xs text-muted-foreground">Defects</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-blue-600">{activeObservations.length}</div>
          <div className="text-xs text-muted-foreground">Observations</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-green-600">{completedAll.length}</div>
          <div className="text-xs text-muted-foreground">Completed</div>
        </Card>
      </div>

      {/* Actions */}
      <div className="flex gap-3 mb-6">
        <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-defect`}>
          <Button>
            <Plus className="w-4 h-4 mr-2" />
            Add Defect
          </Button>
        </Link>
        <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-observation`}>
          <Button variant="secondary">
            <Plus className="w-4 h-4 mr-2" />
            Add Observation
          </Button>
        </Link>

        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              variant="secondary"
              disabled={!!generating || !defects?.length}
            >
              <FileDown className="w-4 h-4 mr-2" />
              {generating ? "Generating..." : "Export Report"}
              <ChevronDown className="w-3.5 h-3.5 ml-1.5" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start">
            <DropdownMenuItem onClick={handleGenerateWord}>
              <FileText className="w-4 h-4 mr-2" />
              Word Document (.docx)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={handleGeneratePdf}>
              <FileText className="w-4 h-4 mr-2" />
              PDF (.pdf)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={() => {
              window.open(`${API_BASE}/api/reports/${reportId}/photos-zip?scope=current`, "_blank");
            }}>
              <ImageDown className="w-4 h-4 mr-2" />
              Export Images — this inspection (.zip)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={() => {
              window.open(`${API_BASE}/api/reports/${reportId}/photos-zip?scope=all`, "_blank");
            }}>
              <ImageDown className="w-4 h-4 mr-2" />
              Export Images — all photos (.zip)
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>

      {defectsLoading ? (
        <div className="space-y-3">
          {[1, 2].map((i) => (
            <div key={i} className="h-20 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      ) : !defects?.length ? (
        <div className="flex flex-col items-center justify-center py-16 text-center">
          <Camera className="w-10 h-10 text-muted-foreground/40 mb-4" />
          <h2 className="text-base font-medium mb-1">No entries recorded</h2>
          <p className="text-sm text-muted-foreground mb-6 max-w-xs">
            Start your inspection by adding defects or observations with photos and details.
          </p>
          <div className="flex gap-3">
            <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-defect`}>
              <Button>
                <Plus className="w-4 h-4 mr-2" />
                Add Defect
              </Button>
            </Link>
            <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-observation`}>
              <Button variant="secondary">
                <Plus className="w-4 h-4 mr-2" />
                Add Observation
              </Button>
            </Link>
          </div>
        </div>
      ) : (
        <div className="space-y-8">
          {activeDefects.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Defects ({activeDefects.length})
                </h2>
              </div>
              <div className="space-y-2">
                {activeDefects.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                  />
                ))}
              </div>
            </div>
          )}

          {activeObservations.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Eye className="w-4 h-4 text-blue-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Observations ({activeObservations.length})
                </h2>
              </div>
              <div className="space-y-2">
                {activeObservations.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                  />
                ))}
              </div>
            </div>
          )}

          {completedAll.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Archive className="w-4 h-4 text-green-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Completed ({completedAll.length})
                </h2>
              </div>
              <div className="space-y-2">
                {completedAll.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                  />
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function DefectCard({ defect, projectId, reportId, onDelete }: { defect: Defect; projectId: string; reportId: string; onDelete: () => void }) {
  const isComplete = defect.status === "complete";

  return (
    <Card className="group relative">
      <Link href={`/projects/${projectId}/reports/${reportId}/defects/${defect.id}`}>
        <div className={`flex items-center justify-between p-4 cursor-pointer hover:bg-accent/50 rounded-lg transition-colors ${isComplete ? "opacity-80" : ""}`}>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2 mb-1">
              <span className="font-mono text-sm font-semibold">{defect.uid}</span>
              <Badge
                variant="secondary"
                className={isComplete
                  ? "bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400"
                  : "bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-400"
                }
              >
                {isComplete ? (
                  <><CheckCircle2 className="w-3 h-3 mr-1" />Complete</>
                ) : (
                  <><AlertTriangle className="w-3 h-3 mr-1" />Open</>
                )}
              </Badge>
            </div>
            <p className="text-sm text-muted-foreground truncate">{defect.comment}</p>
            <div className="flex gap-3 mt-1 text-xs text-muted-foreground">
              <span>Assigned: {defect.assignedTo}</span>
              {isComplete && defect.dateClosed ? (
                <span className="text-green-600 dark:text-green-400">Completed: {defect.dateClosed}</span>
              ) : (
                <span>Due: {defect.dueDate}</span>
              )}
            </div>
          </div>
          <ChevronRight className="w-5 h-5 text-muted-foreground shrink-0 ml-3" />
        </div>
      </Link>
      <button
        onClick={(e) => {
          e.stopPropagation();
          if (confirm("Delete this defect and all its photos?")) {
            onDelete();
          }
        }}
        className="absolute top-3 right-12 p-1.5 rounded-md opacity-0 group-hover:opacity-100 hover:bg-destructive/10 hover:text-destructive transition-all"
      >
        <Trash2 className="w-4 h-4" />
      </button>
    </Card>
  );
}
