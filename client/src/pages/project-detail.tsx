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
  ChevronDown, FileDown
} from "lucide-react";
import type { Project, Defect, Photo } from "@shared/schema";
import { useState, useMemo } from "react";

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

// Helper to load an image as ArrayBuffer (for docx) or data URL (for pdf)
async function loadImageBlob(filename: string): Promise<Blob | null> {
  try {
    const res = await fetch(`${API_BASE}/api/uploads/${filename}`);
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

// Derive location string from defect UID (e.g. "01-13-CR-01" -> "Drop 1, Level 13")
function deriveLocation(uid: string): string {
  const parts = uid.split("-");
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

export default function ProjectDetail() {
  const { id } = useParams<{ id: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [generating, setGenerating] = useState<string | null>(null);

  const { data: project, isLoading: projectLoading } = useQuery<Project>({
    queryKey: ["/api/projects", id],
  });

  const { data: defects, isLoading: defectsLoading } = useQuery<Defect[]>({
    queryKey: [`/api/projects/${id}/defects`],
  });

  const deleteMutation = useMutation({
    mutationFn: async (defectId: number) => {
      await apiRequest("DELETE", `/api/defects/${defectId}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${id}/defects`] });
      toast({ title: "Defect deleted" });
    },
  });

  const activeDefects = useMemo(() =>
    defects?.filter((d) => d.status !== "complete") ?? [], [defects]);
  const completedDefects = useMemo(() =>
    defects?.filter((d) => d.status === "complete")
      .sort((a, b) => (b.dateClosed ?? "").localeCompare(a.dateClosed ?? "")) ?? [],
    [defects]);

  // ==================== PDF GENERATION (AFC Template) ====================
  const handleGeneratePdf = async () => {
    setGenerating("pdf");
    try {
      const res = await apiRequest("GET", `/api/projects/${id}/report-data`);
      const data = await res.json();

      const { default: jsPDF } = await import("jspdf");
      const autoTable = (await import("jspdf-autotable")).default;

      const doc = new jsPDF("p", "mm", "a4");
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 12.7; // 0.5 inch = 12.7mm
      const contentWidth = pageWidth - margin * 2;

      const logo = await loadAfcLogo();

      // Colors from AFC template
      const DARK_BLUE = [10, 29, 48] as const;  // #0A1D30
      const TEAL = [0, 235, 230] as const;        // #00EBE6
      const ACCENT_BLUE = [69, 176, 225] as const; // #45B0E1
      const DARK_TEXT = [58, 58, 58] as const;      // #3A3A3A
      const CAPTION_BLUE = [14, 40, 65] as const;   // #0E2841

      // ======= COVER PAGE =======
      // Teal left sidebar
      doc.setFillColor(...TEAL);
      doc.rect(0, 0, 25, pageHeight, "F");

      // AFC Logo top right
      if (logo) {
        doc.addImage(logo.dataUrl, "PNG", pageWidth - 60, 15, 45, 14.4);
      }

      // Title block - right side
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
      const rev = data.project.revision || "01";
      doc.text(`Revision ${rev}`, titleX, ty);
      ty += 20;

      // Date
      doc.setFontSize(12);
      doc.setTextColor(100);
      doc.text(formatReportDate(new Date().toISOString()), titleX, ty);
      ty += 16;

      // Company details
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

      // 1.1 General
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

      // 1.2 Inspection
      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("1.2 Inspection", margin, y);
      y += 8;

      autoTable(doc, {
        startY: y,
        body: [
          ["Date", formatReportDate(new Date().toISOString())],
          ["Inspector", data.project.inspector],
          ["Location", data.project.address],
          ["Client", data.project.client],
        ],
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

      // Observations table matching AFC template: ID | Location | Observation | Action Required | Status | Photo
      const allDefects = [...(data.defects || [])];
      const slotOrder = ["wip1", "wip2", "wip3", "complete"];

      // Load thumbnail for each defect (first available photo)
      const photoThumbnails: Record<string, string> = {};
      for (const defect of allDefects) {
        if (defect.photos && defect.photos.length > 0) {
          const sortedPhotos = slotOrder
            .map((s: string) => defect.photos.find((p: any) => p.slot === s))
            .filter(Boolean);
          if (sortedPhotos.length > 0) {
            const blob = await loadImageBlob(sortedPhotos[sortedPhotos.length - 1].filename);
            if (blob) {
              photoThumbnails[defect.uid] = await blobToDataUrl(blob);
            }
          }
        }
      }

      // Build table data
      const tableHead = [["ID", "Location", "Observation", "Action Required", "Status", "Photo"]];
      const tableBody = allDefects.map((d: any) => [
        d.uid,
        deriveLocation(d.uid),
        d.comment,
        d.actionRequired,
        d.status === "complete" ? "Complete" : "Open",
        "", // Photo cell will be drawn manually
      ]);

      if (tableBody.length > 0) {
        autoTable(doc, {
          startY: y,
          head: tableHead,
          body: tableBody,
          margin: { left: margin, right: margin },
          styles: { fontSize: 7, cellPadding: 2, overflow: "linebreak", valign: "top" },
          headStyles: {
            fillColor: [255, 255, 255],
            textColor: [...CAPTION_BLUE],
            fontStyle: "bold",
            lineWidth: { bottom: 0.5 },
            lineColor: [...DARK_TEXT],
          },
          columnStyles: {
            0: { cellWidth: 22 },
            1: { cellWidth: 28 },
            2: { cellWidth: 52 },
            3: { cellWidth: 28 },
            4: { cellWidth: 18 },
            5: { cellWidth: contentWidth - 148 },
          },
          didDrawPage: () => {
            addHeader();
            addFooter();
          },
          didDrawCell: (cellData: any) => {
            // Draw photo thumbnails in column 5
            if (cellData.section === "body" && cellData.column.index === 5) {
              const defect = allDefects[cellData.row.index];
              if (defect && photoThumbnails[defect.uid]) {
                const imgW = cellData.cell.width - 2;
                const imgH = Math.min(imgW * 0.75, cellData.cell.height - 2);
                doc.addImage(
                  photoThumbnails[defect.uid],
                  "JPEG",
                  cellData.cell.x + 1,
                  cellData.cell.y + 1,
                  imgW,
                  imgH
                );
              }
            }
          },
          rowPageBreak: "auto",
        });
      } else {
        doc.setFontSize(9);
        doc.text("No defects recorded.", margin, y);
      }

      // ======= APPENDIX A - SUPPLEMENTARY PHOTOGRAPHS =======
      doc.addPage();
      addHeader();
      addFooter();
      y = 20;

      doc.setFontSize(18);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...DARK_TEXT);
      doc.text("Appendix A — Supplementary Photographs", margin, y);
      y += 10;

      const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", complete: "Complete" };

      for (const defect of allDefects) {
        if (!defect.photos || defect.photos.length === 0) continue;

        const sortedPhotos = slotOrder
          .map((s: string) => defect.photos.find((p: any) => p.slot === s))
          .filter(Boolean);
        if (sortedPhotos.length === 0) continue;

        // Check if we need a new page
        if (y + 10 > pageHeight - 30) {
          doc.addPage();
          addHeader();
          addFooter();
          y = 20;
        }

        // Defect caption
        doc.setFontSize(9);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...CAPTION_BLUE);
        doc.text(`${defect.uid} — ${deriveLocation(defect.uid)}`, margin, y);
        y += 5;

        // Photo grid (2 columns)
        const imgW = (contentWidth - 6) / 2;
        const imgH = imgW * 0.7;

        for (let i = 0; i < sortedPhotos.length; i++) {
          const photo = sortedPhotos[i];
          const col = i % 2;
          if (i > 0 && col === 0) y += imgH + 10;
          if (col === 0 && y + imgH + 10 > pageHeight - 20) {
            doc.addPage();
            addHeader();
            addFooter();
            y = 20;
          }
          const x = margin + col * (imgW + 6);

          const blob = await loadImageBlob(photo.filename);
          if (blob) {
            const dataUrl = await blobToDataUrl(blob);
            doc.addImage(dataUrl, "JPEG", x, y, imgW, imgH);
          } else {
            doc.setDrawColor(180);
            doc.rect(x, y, imgW, imgH);
          }

          doc.setFontSize(7);
          doc.setFont("helvetica", "bold");
          doc.setTextColor(100);
          const caption = `${slotLabels[photo.slot] || photo.slot}`;
          doc.text(caption, x, y + imgH + 4);
          doc.setTextColor(0);
        }

        const lastCol = (sortedPhotos.length - 1) % 2;
        y += imgH + (lastCol === 1 ? 10 : 10);
        y += 4;
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
      const res = await apiRequest("GET", `/api/projects/${id}/report-data`);
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

      const allDefects = data.defects || [];
      const openData = allDefects.filter((d: any) => d.status !== "complete");
      const completedData = allDefects.filter((d: any) => d.status === "complete");
      const slotOrder = ["wip1", "wip2", "wip3", "complete"];
      const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", complete: "Complete" };

      // AFC Colors
      const DARK_BLUE = "0A1D30";
      const TEAL = "00EBE6";
      const ACCENT_BLUE = "45B0E1";
      const DARK_TEXT = "3A3A3A";
      const CAPTION_BLUE = "0E2841";

      // Common borders
      const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
      const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
      const accentBottomBorder = { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE };

      // Date formatting
      const reportDate = formatReportDate(new Date().toISOString());
      const rev = data.project.revision || "01";
      const afcRef = data.project.afcReference || "AFC-24XXX";

      // ======= HEADER (for all pages except first) =======
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

      const defaultHeader = new Header({
        children: [headerTable],
      });

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

      // ======= COVER PAGE SECTION =======
      const coverChildren: any[] = [];

      // Logo
      if (logo) {
        coverChildren.push(new Paragraph({
          children: [
            new ImageRun({
              data: logo.buffer,
              transformation: { width: 180, height: 58 },
              type: "png",
            }),
          ],
          alignment: AlignmentType.RIGHT,
          spacing: { after: 600 },
        }));
      } else {
        coverChildren.push(new Paragraph({ spacing: { after: 600 } }));
      }

      // Spacer
      coverChildren.push(new Paragraph({ spacing: { before: 1200 } }));

      // Title
      coverChildren.push(new Paragraph({
        children: [new TextRun({
          text: data.project.name.toUpperCase(),
          bold: true,
          size: 72,
          font: "Aptos",
          color: DARK_BLUE,
        })],
        spacing: { after: 100 },
      }));

      // Subtitle
      coverChildren.push(new Paragraph({
        children: [new TextRun({
          text: "SITE VISIT REPORT",
          size: 52,
          font: "Aptos",
          color: DARK_BLUE,
          smallCaps: true,
        })],
        spacing: { after: 100 },
      }));

      // Revision
      coverChildren.push(new Paragraph({
        children: [new TextRun({
          text: `Revision ${rev}`,
          size: 28,
          font: "Aptos",
          color: DARK_TEXT,
        })],
        spacing: { after: 600 },
      }));

      // Date
      coverChildren.push(new Paragraph({
        children: [new TextRun({
          text: reportDate,
          size: 24,
          font: "Aptos",
          color: DARK_TEXT,
        })],
        spacing: { after: 400 },
      }));

      // Company info
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

      // 1.1 General
      introChildren.push(new Paragraph({
        children: [new TextRun({ text: "General", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_2,
        spacing: { after: 100 },
      }));

      introChildren.push(new Paragraph({
        children: [new TextRun({
          text: `Angel Façade Consulting (AFC) was engaged by ${data.project.client} to carry out a site visit inspection of the facade at ${data.project.address}.`,
          size: 20,
          font: "Aptos",
        })],
        spacing: { after: 200 },
      }));

      // 1.2 Inspection
      introChildren.push(new Paragraph({
        children: [new TextRun({ text: "Inspection", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_2,
        spacing: { after: 100 },
      }));

      // Inspection table
      const inspectionData = [
        ["Date", reportDate],
        ["Inspector", data.project.inspector],
        ["Locations covered", data.project.address],
        ["Client", data.project.client],
      ];

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

      // ======= SECTION 2 - DEFECT REGISTER & RECTIFICATION LOG =======
      const obsChildren: any[] = [];

      obsChildren.push(new Paragraph({
        children: [new TextRun({ text: "Defect Register & Rectification Log", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 100 },
      }));

      obsChildren.push(new Paragraph({
        children: [new TextRun({
          text: "Based on our observations, we recommend the following actions.",
          size: 20,
          font: "Aptos",
        })],
        spacing: { after: 200 },
      }));

      // Build the 6-column observations table matching AFC template
      // Columns: ID | Location | Observation | Action Required | Status | Photo
      // Template widths (dxa): 846, 1139, 2833, 1136, 1276, 3260
      const obsTableRows: any[] = [];

      // Header row
      const headerLabels = ["ID", "Location", "Observation", "Action Required", "Status", "Photo"];
      const headerWidths = [846, 1139, 2833, 1136, 1276, 3260];

      obsTableRows.push(new TableRow({
        tableHeader: true,
        children: headerLabels.map((label, i) =>
          new TableCell({
            width: { size: headerWidths[i], type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: label, size: 18, font: "Aptos", bold: true, color: CAPTION_BLUE })],
              spacing: { before: 40, after: 40 },
            })],
            borders: {
              top: noBorder,
              left: noBorder,
              right: noBorder,
              bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_TEXT },
            },
          })
        ),
      }));

      // Data rows
      for (const defect of allDefects) {
        // Get the latest photo for the Photo column
        let photoImageRun: any = null;
        if (defect.photos && defect.photos.length > 0) {
          const sortedPhotos = slotOrder
            .map((s: string) => defect.photos.find((p: any) => p.slot === s))
            .filter(Boolean);
          if (sortedPhotos.length > 0) {
            const lastPhoto = sortedPhotos[sortedPhotos.length - 1];
            const blob = await loadImageBlob(lastPhoto.filename);
            if (blob) {
              const buffer = await blobToArrayBuffer(blob);
              photoImageRun = new ImageRun({
                data: buffer,
                transformation: { width: 150, height: 100 },
                type: "jpg",
              });
            }
          }
        }

        const statusText = defect.status === "complete" ? "Complete" : "Open";

        const cellData = [
          defect.uid,
          deriveLocation(defect.uid),
          defect.comment,
          defect.actionRequired,
          statusText,
        ];

        const rowBorder = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };

        obsTableRows.push(new TableRow({
          children: [
            ...cellData.map((text: string, i: number) =>
              new TableCell({
                width: { size: headerWidths[i], type: WidthType.DXA },
                children: [new Paragraph({
                  children: [new TextRun({
                    text,
                    size: 16,
                    font: "Aptos",
                    color: i === 4 ? (statusText === "Complete" ? "228B22" : "C89600") : undefined,
                    bold: i === 0,
                  })],
                  spacing: { before: 30, after: 30 },
                })],
                borders: { top: rowBorder, left: noBorder, right: noBorder, bottom: rowBorder },
              })
            ),
            // Photo cell
            new TableCell({
              width: { size: headerWidths[5], type: WidthType.DXA },
              children: [
                photoImageRun
                  ? new Paragraph({
                      children: [photoImageRun],
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 30, after: 30 },
                    })
                  : new Paragraph({
                      children: [new TextRun({ text: "—", size: 16, font: "Aptos", color: "AAAAAA" })],
                      spacing: { before: 30, after: 30 },
                    }),
              ],
              borders: { top: rowBorder, left: noBorder, right: noBorder, bottom: rowBorder },
            }),
          ],
        }));
      }

      if (allDefects.length > 0) {
        obsChildren.push(new Table({
          width: { size: 10490, type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          rows: obsTableRows,
        }));
      } else {
        obsChildren.push(new Paragraph({
          children: [new TextRun({ text: "No defects recorded.", size: 20, font: "Aptos", color: "999999", italics: true })],
        }));
      }

      // ======= APPENDIX A - SUPPLEMENTARY PHOTOGRAPHS =======
      const appendixChildren: any[] = [];

      appendixChildren.push(new Paragraph({
        children: [new TextRun({ text: "Supplementary Photographs", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 200 },
      }));

      for (const defect of allDefects) {
        if (!defect.photos || defect.photos.length === 0) continue;

        const sortedPhotos = slotOrder
          .map((s: string) => defect.photos.find((p: any) => p.slot === s))
          .filter(Boolean);
        if (sortedPhotos.length === 0) continue;

        // Defect sub-heading
        appendixChildren.push(new Paragraph({
          children: [new TextRun({
            text: `${defect.uid} — ${deriveLocation(defect.uid)}`,
            bold: true,
            size: 20,
            font: "Aptos",
            color: DARK_TEXT,
          })],
          spacing: { before: 200, after: 100 },
        }));

        // 2-column photo grid
        for (let i = 0; i < sortedPhotos.length; i += 2) {
          const photoCells: any[] = [];

          for (let j = 0; j < 2; j++) {
            const photo = sortedPhotos[i + j];
            if (photo) {
              const blob = await loadImageBlob(photo.filename);
              const cellChildren: any[] = [];

              if (blob) {
                const buffer = await blobToArrayBuffer(blob);
                cellChildren.push(new Paragraph({
                  children: [
                    new ImageRun({
                      data: buffer,
                      transformation: { width: 240, height: 180 },
                      type: "jpg",
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                }));
              }

              cellChildren.push(new Paragraph({
                children: [new TextRun({
                  text: `${slotLabels[photo.slot] || photo.slot}`,
                  bold: true,
                  size: 16,
                  font: "Aptos",
                  color: DARK_TEXT,
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 60 },
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

          appendixChildren.push(new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            rows: [new TableRow({ children: photoCells })],
          }));

          appendixChildren.push(new Paragraph({ spacing: { before: 80 } }));
        }
      }

      // ======= BUILD DOCUMENT =======
      const sectionProps = {
        page: {
          size: { width: 11906, height: 16838 }, // A4 in twips
          margin: { top: 720, right: 720, bottom: 720, left: 720 }, // 0.5 inch
        },
        headers: { default: defaultHeader },
        footers: { default: defaultFooter },
      };

      const wordDoc = new Document({
        creator: "Angel Façade Consulting",
        title: `${data.project.name} — Site Visit Report`,
        styles: {
          default: {
            document: {
              run: {
                font: "Aptos",
                size: 20,
                color: DARK_TEXT,
              },
            },
          },
        },
        sections: [
          // Cover page (no header/footer)
          {
            properties: {
              page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 851, right: 1440, bottom: 1440, left: 1440 },
              },
            },
            children: coverChildren,
          },
          // Introduction
          {
            properties: sectionProps,
            children: introChildren,
          },
          // Observations
          {
            properties: sectionProps,
            children: obsChildren,
          },
          // Appendix
          {
            properties: sectionProps,
            children: appendixChildren,
          },
        ],
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

  if (projectLoading) {
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

  if (!project) return null;

  return (
    <div className="max-w-3xl mx-auto px-4 py-8">
      {/* Header */}
      <div className="mb-8">
        <Link href="/">
          <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4" data-testid="button-back">
            <ArrowLeft className="w-4 h-4" />
            All Projects
          </button>
        </Link>
        <div className="flex items-start justify-between gap-4">
          <div>
            <h1 className="text-xl font-semibold tracking-tight" data-testid="text-project-name">
              {project.name}
            </h1>
            <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 text-sm text-muted-foreground">
              <span className="flex items-center gap-1.5">
                <MapPin className="w-3.5 h-3.5" />
                {project.address}
              </span>
              <span className="flex items-center gap-1.5">
                <User className="w-3.5 h-3.5" />
                {project.client}
              </span>
              <span className="flex items-center gap-1.5">
                <UserCheck className="w-3.5 h-3.5" />
                {project.inspector}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-3 gap-3 mb-6">
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold">{defects?.length ?? 0}</div>
          <div className="text-xs text-muted-foreground">Total</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-amber-600">{activeDefects.length}</div>
          <div className="text-xs text-muted-foreground">Active</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-green-600">{completedDefects.length}</div>
          <div className="text-xs text-muted-foreground">Completed</div>
        </Card>
      </div>

      {/* Actions */}
      <div className="flex gap-3 mb-6">
        <Link href={`/projects/${id}/defects/new`}>
          <Button data-testid="button-add-defect">
            <Plus className="w-4 h-4 mr-2" />
            Add Defect
          </Button>
        </Link>

        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              variant="secondary"
              disabled={!!generating || !defects?.length}
              data-testid="button-generate-report"
            >
              <FileDown className="w-4 h-4 mr-2" />
              {generating ? "Generating..." : "Export Report"}
              <ChevronDown className="w-3.5 h-3.5 ml-1.5" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start">
            <DropdownMenuItem onClick={handleGenerateWord} data-testid="button-export-word">
              <FileText className="w-4 h-4 mr-2" />
              Word Document (.docx)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={handleGeneratePdf} data-testid="button-export-pdf">
              <FileText className="w-4 h-4 mr-2" />
              PDF (.pdf)
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
          <h2 className="text-base font-medium mb-1">No defects recorded</h2>
          <p className="text-sm text-muted-foreground mb-6 max-w-xs">
            Start your inspection by adding defects with photos and details.
          </p>
          <Link href={`/projects/${id}/defects/new`}>
            <Button data-testid="button-empty-add-defect">
              <Plus className="w-4 h-4 mr-2" />
              Add Defect
            </Button>
          </Link>
        </div>
      ) : (
        <div className="space-y-8">
          {activeDefects.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Active ({activeDefects.length})
                </h2>
              </div>
              <div className="space-y-2">
                {activeDefects.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={id!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                  />
                ))}
              </div>
            </div>
          )}

          {completedDefects.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Archive className="w-4 h-4 text-green-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Completed ({completedDefects.length})
                </h2>
              </div>
              <div className="space-y-2">
                {completedDefects.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={id!}
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

function DefectCard({ defect, projectId, onDelete }: { defect: Defect; projectId: string; onDelete: () => void }) {
  const isComplete = defect.status === "complete";

  return (
    <Card className="group relative" data-testid={`card-defect-${defect.id}`}>
      <Link href={`/projects/${projectId}/defects/${defect.id}`}>
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
        data-testid={`button-delete-defect-${defect.id}`}
      >
        <Trash2 className="w-4 h-4" />
      </button>
    </Card>
  );
}
