// ============================================================================
// render-pdf-v2.ts — WP8 PDF layout for NEW projects (featureVersion >= 2).
// Older projects never reach this file: renderPdf() only hands over when the
// project uses the new wording, so their reports are unchanged.
//
//   Title page, Contents (clickable, with page numbers and PDF bookmarks), then
//   1 Introduction · 2 Project Status · 3 Project Instructions · 4 New This
//   Inspection · 5 Closed This Inspection · 6 Running List of Open Items ·
//   7 Elevations · Appendix A. Numbering comes from report-outline-v2.ts.
//
// Text is set in Inter, the closest openly licensed match to Aptos (Aptos itself
// cannot be bundled in an app). The Word export names Aptos directly.
// ============================================================================

import type { ReportTree, CategoryGroup } from "./report-tree";
import {
  loadImageBlob, blobToDataUrl, compressImageForPdfExport, loadAfcLogo,
  getWorkTypeLabel, ELEVATION_NAMES, deriveLocation, formatDefectLocation,
  getLocationDimensions, formatReportDate, formatPhotoDate, isDefectOverdue,
  comparePhotoSlots, photoSlotLabel, resolveActionSummary,
  parseProjectSnapshot, resolveProjectField, deriveAreaRef, compressImageWithSize,
  workTypeLabelFor, uidNomenclatureSentence, areaRefFor,
} from "./render-helpers";
import { INSTRUCTION_HEADERS, INSTRUCTION_SUBTITLE, instructionRows } from "./project-instructions";
import { buildOutlineV2, outlineHeading, scopeRowsV2, backgroundDocsV2, harvardReference, titlePageLinesV2, type OutlineEntry } from "./report-outline-v2";
import { loadReportElevations, PIN_COLOURS, type PrintedElevation } from "./elevation-print";

export const LIMITATIONS_V2: string[] = [
  "The extent of our inspection is limited to the external surfaces of the building where works are underway unless otherwise noted within this report.",
  "Only those works nominated above form part of AFC\u2019s scope for inspection.",
  "The Contractor remains wholly responsible for construction documentation, workmanship, testing, installation, certification and guarantees.",
  "Visual inspection of the facade was undertaken at safely accessible areas only. Harnesses, fall arrest and fall restraint systems and equipment were utilised where necessary.",
  "Our assessments are based on a limited visual inspection of the areas identified only and do not include dimensional and engineering checks. AFC does not accept liability for items that have not been inspected and not identified in the photographs.",
  "No materials sampling or testing, destructive investigations, water testing or structural analysis of the existing facade systems has been carried out by AFC.",
  "By virtue of the scope and scale of this work, AFC can\u2019t make comment on any possible structural inadequacies of the facade design, fabrication or installation.",
  "Our inspection will not allow assessment of other aspects of fa\u00e7ade performance such as acoustics, building sealing, damp and weatherproofing or solar/thermal performance.",
  "This report has been prepared for the exclusive use of the nominated Client and shall therefore not be relied upon by any third party without their express written consent.",
];

// ---- Font: Inter (SIL Open Font Licence, client/public/fonts/Inter-OFL.txt) ----
const INTER_FILES: Array<[string, string]> = [
  ["Inter-Regular.ttf", "normal"],
  ["Inter-Bold.ttf", "bold"],
  ["Inter-Italic.ttf", "italic"],
  ["Inter-BoldItalic.ttf", "bolditalic"],
];
let interCache: Array<[string, string, string]> | null = null;

function toBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
  return btoa(bin);
}

/** Registers Inter with the document; falls back to Helvetica if the files can't load. */
async function useInter(doc: any): Promise<string> {
  try {
    if (!interCache) {
      const loaded: Array<[string, string, string]> = [];
      for (const [file, style] of INTER_FILES) {
        const res = await fetch(`/fonts/${file}`);
        if (!res.ok) throw new Error(`font ${file}: ${res.status}`);
        loaded.push([file, style, toBase64(await res.arrayBuffer())]);
      }
      interCache = loaded;
    }
    for (const [file, style, b64] of interCache) {
      doc.addFileToVFS(file, b64);
      doc.addFont(file, "Inter", style);
    }
    return "Inter";
  } catch (err) {
    console.warn("[render-pdf-v2] Inter unavailable, using Helvetica", err);
    return "helvetica";
  }
}

export async function renderPdfV2(tree: ReportTree): Promise<Blob> {
  const v2 = tree.v2!;
  const data = { project: tree.project, report: tree.report, projectInstructions: tree.projectInstructions };
  const dims = getLocationDimensions(tree.project?.locationDimensions);

  const { default: jsPDF } = await import("jspdf");
  const autoTable = (await import("jspdf-autotable")).default;

  const doc: any = new jsPDF("p", "mm", "a4");
  const FONT = await useInter(doc);
  const margin = 12.7;
  const TOP = 20;
  const pw = () => doc.internal.pageSize.getWidth();
  const ph = () => doc.internal.pageSize.getHeight();
  const cw = () => pw() - margin * 2; // header/footer width
  const limit = () => ph() - 18;
  const pageNo = () => doc.getCurrentPageInfo().pageNumber;

  // Body text: 10 pt, black, single line spacing, one blank line between paragraphs.
  const BODY_PT = 10;
  const LINE = 4.4; // mm, single spacing at 10 pt
  const PARA_GAP = 3.5; // mm, one 10 pt line
  const HANG = 12.7; // mm, level 2 heading hanging indent (1.27 cm)

  const TEAL = [0, 235, 230] as const;
  const DARK_BLUE = [10, 29, 48] as const;
  const ACCENT_BLUE = [0, 205, 200] as const;
  const DARK_TEXT = [58, 58, 58] as const;
  const CAPTION_BLUE = [14, 40, 65] as const;
  const OVERDUE_RED = [192, 0, 0] as const;
  const LINK_BLUE = [5, 99, 193] as const;

  const snap = parseProjectSnapshot((data.report as any).projectSnapshot);
  const projAddress = resolveProjectField(snap, data.project, "address");
  const projName = resolveProjectField(snap, data.project, "name");
  const projReportTitle = resolveProjectField(snap, data.project, "reportTitle");
  const projInspector = resolveProjectField(snap, data.project, "inspector");
  const projClient = resolveProjectField(snap, data.project, "client");
  const afcRef = resolveProjectField(snap, data.project, "afcReference") || "AFC-24XXX";
  const areaRefTemplate: string = String(resolveProjectField(snap, data.project, "areaRefTemplate") || "");
  const inspNumRaw = String(data.report.inspectionNumber || "").trim();
  const inspNum = inspNumRaw ? (/^\d+$/.test(inspNumRaw) ? inspNumRaw.padStart(2, "0") : inspNumRaw) : "";

  const elevations: PrintedElevation[] = await loadReportElevations(tree.project.id, tree.report.id, v2.elevationRecordIds);
  const outline: OutlineEntry[] = buildOutlineV2(tree, elevations);
  const logo = await loadAfcLogo();

  // ---- Page furniture: drawn once per page, whatever the orientation ----
  const decorated = new Set<number>();
  const decorate = () => {
    const p = pageNo();
    if (decorated.has(p)) return;
    decorated.add(p);
    doc.setFontSize(8); doc.setFont(FONT, "bold"); doc.setTextColor(...DARK_TEXT);
    doc.text(String(projName || "").toUpperCase(), margin, 10);
    doc.setFont(FONT, "normal");
    doc.text("ANGEL FA\u00c7ADE CONSULTING", pw() - margin, 10, { align: "right" });
    doc.setDrawColor(...ACCENT_BLUE); doc.setLineWidth(0.5);
    doc.line(margin, 12, pw() - margin, 12);
    doc.line(margin, ph() - 12, pw() - margin, ph() - 12);
    doc.text(`${afcRef} | ${projAddress || ""}`, margin, ph() - 8);
    doc.text(`Page ${p}`, pw() - margin, ph() - 8, { align: "right" });
  };
  let y = TOP;
  const newPage = (orientation: "p" | "l" = "p") => { doc.addPage("a4", orientation); decorate(); y = TOP; return y; };
  const ensure = (h: number) => { if (y + h > limit()) newPage(); };
  const tableHook = { didDrawPage: () => decorate() };

  // ---- Destinations for the contents page, bookmarks and UID links ----
  const dests = new Map<string, { page: number; top: number }>();
  const recordDests = new Map<number, { page: number; top: number }>();
  const uidLinks: Array<{ page: number; x: number; y: number; w: number; h: number; id: number }> = [];

  // ---- Text helpers ----
  const body = (size = BODY_PT) => { doc.setFont(FONT, "normal"); doc.setFontSize(size); doc.setTextColor(0, 0, 0); };
  const para = (text: string, opts: { indent?: number; gap?: number } = {}) => {
    body();
    const indent = opts.indent ?? 0;
    const lines: string[] = doc.splitTextToSize(String(text), cw() - indent);
    for (const ln of lines) { ensure(LINE); doc.text(ln, margin + indent, y); y += LINE; }
    y += opts.gap ?? PARA_GAP;
  };
  const h1 = (key: string) => {
    const h = outlineHeading(outline, key);
    doc.setFont(FONT, "bold"); doc.setFontSize(18); doc.setTextColor(...DARK_TEXT);
    y += 2;
    dests.set(key, { page: pageNo(), top: y - 8 });
    doc.text(h.text, margin, y);
    y += 10;
  };
  const h2 = (key: string) => {
    const h = outlineHeading(outline, key);
    // No number (an appendix name): the title starts at the margin.
    const ind = h.no ? HANG : 0;
    doc.setFont(FONT, "bold"); doc.setFontSize(14);
    const lines: string[] = doc.splitTextToSize(h.title, cw() - ind);
    ensure(0.7 + lines.length * 6 + 14);
    // Set the heading font after ensure(): a page break draws the running header, which
    // leaves a small font behind (this is what shrank 1.4 at the top of a page).
    doc.setFont(FONT, "bold"); doc.setFontSize(14);
    y += 0.7; // 2 pt before
    y += 4.5;
    dests.set(key, { page: pageNo(), top: y - 6 });
    doc.setTextColor(...DARK_TEXT);
    if (h.no) doc.text(h.no, margin, y);
    lines.forEach((ln, i) => doc.text(ln, margin + ind, y + i * 6));
    y += (lines.length - 1) * 6 + 2.1 + 3.5; // 6 pt after
  };
  // Numbered level 3 heading (2.3.1 …): level 3 type, number and title on the level 2
  // hanging indent.
  const h3n = (key: string, keepWith = 24) => {
    const h = outlineHeading(outline, key);
    doc.setFont(FONT, "bold"); doc.setFontSize(12);
    const lines: string[] = doc.splitTextToSize(h.title, cw() - HANG);
    ensure(1 + lines.length * 5 + keepWith);
    doc.setFont(FONT, "bold"); doc.setFontSize(12); doc.setTextColor(...CAPTION_BLUE);
    y += 1;
    dests.set(key, { page: pageNo(), top: y - 5 });
    doc.text(h.no, margin, y);
    lines.forEach((ln, i) => doc.text(ln, margin + HANG, y + i * 5));
    y += (lines.length - 1) * 5 + 5;
  };
  // AFC numbered points: teal "1." at the margin, text on a 6.3 mm hanging indent (the
  // Word list uses 0.63 cm). numMark draws the number at the current line.
  const NUM_IND = 6.3;
  const numMark = (n: number) => {
    doc.setFont(FONT, "normal"); doc.setTextColor(...ACCENT_BLUE);
    doc.text(`${n}.`, margin, y);
    body();
  };
  // Unnumbered sub-heading (category groups, "Unique identifier", "Notes").
  const h3 = (label: string, keepWith = 30) => {
    ensure(keepWith);
    y += 1;
    doc.setFont(FONT, "bold"); doc.setFontSize(12); doc.setTextColor(...CAPTION_BLUE);
    doc.text(label, margin, y);
    y += 5; // the first record starts on the next line
  };
  const afterTable = (gap = PARA_GAP + 1) => { y = (doc as any).lastAutoTable.finalY + gap; };

  // ===================== TITLE PAGE =====================
  doc.setFillColor(...TEAL);
  doc.rect(0, 0, 25, ph(), "F");
  if (logo) doc.addImage(logo.dataUrl, "PNG", pw() - 60, 15, 45, 14.4);
  {
    const x = 35;
    let ty = ph() * 0.3;
    doc.setFontSize(32); doc.setFont(FONT, "bold"); doc.setTextColor(...DARK_BLUE);
    const addr: string[] = doc.splitTextToSize(String(projAddress || "").toUpperCase(), cw() - 25);
    doc.text(addr, x, ty); ty += addr.length * 12 + 4;
    // Top to bottom: address, project, report title, inspection and revision, then date.
    const tp = titlePageLinesV2({ projectName: projName, reportTitle: projReportTitle, inspectionNumber: inspNum, revision: data.report.revision });
    doc.setFont(FONT, "normal");
    if (tp.project) {
      doc.setFontSize(20);
      const t: string[] = doc.splitTextToSize(tp.project, cw() - 25);
      doc.text(t, x, ty); ty += t.length * 8 + 3;
    }
    doc.setFontSize(16);
    const rt: string[] = doc.splitTextToSize(tp.reportTitle, cw() - 25);
    doc.text(rt, x, ty); ty += rt.length * 6.5 + 6;
    doc.setFontSize(14); doc.setTextColor(...DARK_TEXT);
    doc.text(tp.revisionLine, x, ty); ty += 18;
    doc.setFontSize(12); doc.setTextColor(100);
    doc.text(formatReportDate(new Date().toISOString()), x, ty); ty += 14;
    doc.setFontSize(11); doc.setTextColor(...DARK_TEXT);
    doc.text("Angel Fa\u00e7ade Consulting", x, ty); ty += 6;
    doc.text(`${projInspector} | 0407 759 590`, x, ty); ty += 6;
    doc.text(String(afcRef), x, ty);
  }

  // ===================== CONTENTS (filled in at the end) =====================
  const TOC_ROW = 6.2;
  const tocRowsPerPage = Math.floor((ph() - 18 - TOP - 14) / TOC_ROW);
  const tocEntries = outline.filter((e) => e.level <= 2); // same levels as the Word contents
  const tocPageCount = Math.max(1, Math.ceil(tocEntries.length / tocRowsPerPage));
  const tocFirstPage = 2;
  for (let i = 0; i < tocPageCount; i++) newPage();

  // ===================== 1. INTRODUCTION =====================
  newPage();
  h1("1");
  h2("1.general");
  para(`Angel Fa\u00e7ade Consulting (AFC) was engaged by ${projClient} to carry out regular inspections of the remedial works underway at ${projAddress}. Below is a summary of pertinent project information.`);
  {
    let roles: Array<{ role?: string; entity?: string; contactDetails?: string }> = [];
    try { roles = JSON.parse(resolveProjectField(snap, data.project, "roles") || "[]") || []; } catch { roles = []; }
    if (Array.isArray(roles) && roles.length > 0) {
      autoTable(doc, {
        startY: y, head: [["Role", "Entity", "Contact Details"]],
        body: roles.map((r) => [String(r.role || ""), String(r.entity || ""), String(r.contactDetails || "")]),
        margin: { left: margin, right: margin, top: TOP },
        styles: { font: FONT, fontSize: 9, cellPadding: 2, valign: "top", overflow: "linebreak", textColor: 0 },
        headStyles: { fillColor: [242, 242, 242], textColor: 50, fontStyle: "bold" },
        columnStyles: { 0: { cellWidth: cw() * 0.25 }, 1: { cellWidth: cw() * 0.3 }, 2: { cellWidth: cw() * 0.45 } },
        theme: "grid", ...tableHook,
      });
      afterTable();
    }
  }
  const scope = scopeRowsV2(tree);
  if (scope.length > 0) {
    h2("1.scope");
    autoTable(doc, {
      startY: y, head: [["Area Ref", "Location", "Work item", "Access method"]],
      body: scope.map((s: any) => [String(s.areaRef || ""), String(s.location || ""), String(s.workItem || ""), String(s.accessMethod || "")]),
      margin: { left: margin, right: margin, top: TOP },
      styles: { font: FONT, fontSize: 9, cellPadding: 2, valign: "top", overflow: "linebreak", textColor: 0 },
      headStyles: { fillColor: [242, 242, 242], textColor: 50, fontStyle: "bold" },
      columnStyles: { 0: { cellWidth: cw() * 0.18 }, 1: { cellWidth: cw() * 0.27 }, 2: { cellWidth: cw() * 0.3 }, 3: { cellWidth: cw() * 0.25 } },
      theme: "grid", ...tableHook,
    });
    afterTable();
  }

  {
    const rows: string[][] = [];
    rows.push(["Date", data.report.inspectionDate ? formatReportDate(data.report.inspectionDate) : formatReportDate(new Date().toISOString())]);
    if (data.report.inspectionNumber) rows.push(["Inspection number", String(inspNum || data.report.inspectionNumber)]);
    if (data.report.revision) rows.push(["Revision", String(data.report.revision)]);
    rows.push(["Inspector", String(projInspector || "")]);
    rows.push(["Locations covered", String(data.report.locationsCovered || projAddress || "")]);
    try {
      const att = JSON.parse(data.report.attendees || "[]");
      if (Array.isArray(att)) {
        const lines = att.map((a: any) => (a.company ? `${a.name || ""} (${a.company})` : String(a.name || ""))).filter((s: string) => s.trim());
        if (lines.length) rows.push(["Attendees", lines.join("\n")]);
      }
    } catch { /* malformed attendees: omit the row */ }
    // Keep 1.3 together: if the heading and the whole particulars table won't fit in
    // the space left, start 1.3 on a new page rather than splitting it.
    {
      doc.setFont(FONT, "normal"); doc.setFontSize(BODY_PT);
      const lineH = (BODY_PT * doc.getLineHeightFactor()) / doc.internal.scaleFactor;
      let tableH = 0;
      for (const [label, value] of rows) {
        doc.setFont(FONT, "bold");
        const a = doc.splitTextToSize(label, 40 - 2).length;
        doc.setFont(FONT, "normal");
        const b = doc.splitTextToSize(value, cw() - 40 - 2).length;
        tableH += Math.max(a, b) * lineH + 1; // 0.5 mm padding top and bottom
      }
      ensure(11 + tableH + 2); // heading block + table + small allowance
    }
    h2("1.particulars");
    // Tighter rows than the old layout (change 2).
    autoTable(doc, {
      startY: y, body: rows, margin: { left: margin, right: margin, top: TOP },
      styles: { font: FONT, fontSize: BODY_PT, cellPadding: { top: 0.5, bottom: 0.5, left: 0, right: 2 }, valign: "top", overflow: "linebreak", textColor: 0, minCellHeight: 0 },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: 40 }, 1: { cellWidth: cw() - 40 } },
      theme: "plain", ...tableHook,
    });
    afterTable();
  }

  const bg = backgroundDocsV2(tree);
  if (bg.length > 0) {
    h2("1.background");
    para("The following documents have been reviewed and form the basis of this inspection.");
    // Harvard style: Originator (Year) Title. Doc. no. X, Rev. Y. The title is italic;
    // words are wrapped across the three styles. AFC numbered list (teal numbers).
    const IND = NUM_IND;
    let refNo = 0;
    for (const b of bg) {
      const ref = harvardReference(b);
      if (!ref) continue;
      body();
      refNo++;
      const words: Array<{ w: string; style: "normal" | "italic" }> = [];
      ([[ref.lead, "normal"], [ref.title, "italic"], [ref.tail, "normal"]] as Array<[string, "normal" | "italic"]>).forEach(([t, st]) => {
        for (const w of t.split(/\s+/).filter(Boolean)) words.push({ w, style: st });
      });
      const lines: Array<typeof words> = [[]];
      let lineW = 0;
      const space = () => { doc.setFont(FONT, "normal"); return doc.getTextWidth(" "); };
      for (const wd of words) {
        doc.setFont(FONT, wd.style);
        const ww = doc.getTextWidth(wd.w);
        const max = cw() - IND;
        const cur = lines[lines.length - 1];
        const add = cur.length ? space() + ww : ww;
        if (cur.length && lineW + add > max) { lines.push([wd]); lineW = ww; }
        else { cur.push(wd); lineW += add; }
      }
      ensure(lines.length * LINE);
      numMark(refNo);
      lines.forEach((ln, li) => {
        if (li > 0) ensure(LINE);
        let x = margin + IND;
        ln.forEach((wd, wi) => {
          if (wi > 0) x += space();
          doc.setFont(FONT, wd.style); doc.text(wd.w, x, y); x += doc.getTextWidth(wd.w);
        });
        y += LINE;
      });
      doc.setFont(FONT, "normal");
      y += PARA_GAP;
    }
  }

  h2("1.uid");
  h3("Unique identifier", 20);
  y += 0.5;
  para("Throughout the document, each record (observation or defect) is referred to by its unique identifier (UID).");
  para(uidNomenclatureSentence(data.project, [...v2.newThisInspection, ...v2.closedThisInspection, ...v2.runningOpen].flatMap((g: any) => g.defects || []).concat(v2.notes || [])));

  h2("1.limitations");
  LIMITATIONS_V2.forEach((b, i) => {
    body();
    const wrapped: string[] = doc.splitTextToSize(b, cw() - NUM_IND);
    ensure(wrapped.length * LINE);
    numMark(i + 1);
    wrapped.forEach((ln) => { doc.text(ln, margin + NUM_IND, y); y += LINE; });
    y += PARA_GAP;
  });

  // ===================== SHARED: tables =====================
  const tableStyles: any = { font: FONT, fontSize: 7.5, cellPadding: 1.5, overflow: "linebreak", valign: "top", textColor: 0 };
  const tableHead: any = { fillColor: [255, 255, 255], textColor: [...CAPTION_BLUE], fontStyle: "bold", lineWidth: { bottom: 0.5 }, lineColor: [...DARK_TEXT] };
  const statusText = (d: any) => (d.status === "complete" ? "Complete" : "Open");
  const detailIds = new Set(v2.detailIds);

  // 2.2 Action List: full width, Work Type and Action widened (change 6). UIDs of
  // records with a full entry in section 4 or 5 link to it (change 7).
  const actionTable = (defects: any[]) => {
    const rows: string[][] = []; const overdue: boolean[] = []; const ids: Array<number | null> = [];
    for (const d of defects) {
      const area = areaRefFor(d, areaRefTemplate, data.project, dims) || "";
      const od = isDefectOverdue(d);
      rows.push([d.uid, area, workTypeLabelFor(d), resolveActionSummary(d), d.assignedTo || "\u2014", d.dueDate || "\u2014", statusText(d)]);
      overdue.push(od); ids.push(detailIds.has(d.id) ? d.id : null);
      for (const loc of d.locations || []) {
        rows.push([loc.uid || "", area, workTypeLabelFor(d), "", d.assignedTo || "\u2014", d.dueDate || "\u2014", statusText(d)]);
        overdue.push(od); ids.push(detailIds.has(d.id) ? d.id : null);
      }
    }
    if (rows.length === 0) return;
    // 26 + 17 + 28 + 52.6 + 21 + 19 + 21 = 184.6 mm (header/footer width)
    autoTable(doc, {
      startY: y, head: [["ID", "Location", "Work Type", "Action", "Responsible", "By Date", "Status"]], body: rows,
      margin: { left: margin, right: margin, top: TOP }, styles: tableStyles, headStyles: tableHead,
      columnStyles: { 0: { cellWidth: 26, fontStyle: "bold" }, 1: { cellWidth: 17 }, 2: { cellWidth: 28 }, 3: { cellWidth: 52.6 }, 4: { cellWidth: 21 }, 5: { cellWidth: 19 }, 6: { cellWidth: 21 } },
      ...tableHook,
      didParseCell: (h: any) => { if (h.section === "body" && h.column.index === 0 && ids[h.row.index] != null) h.cell.styles.textColor = [...LINK_BLUE]; },
      didDrawCell: (h: any) => {
        if (h.section !== "body") return;
        const id = ids[h.row.index];
        if (h.column.index === 0 && id != null) uidLinks.push({ page: pageNo(), x: h.cell.x, y: h.cell.y, w: h.cell.width, h: h.cell.height, id });
        if (h.column.index === 6 && overdue[h.row.index]) {
          const base = Array.isArray(h.cell.text) ? h.cell.text.join("") : String(h.cell.text ?? "");
          const pos = h.cell.getTextPos();
          const w = doc.getTextWidth(base);
          doc.setTextColor(...OVERDUE_RED);
          doc.text(" - Overdue", pos.x + w, pos.y, { baseline: "top" });
          doc.setTextColor(0);
        }
      },
    });
    afterTable();
  };

  // Section 6 running list: full width.
  const registerTable = (defects: any[]) => {
    const rows: string[][] = []; const overdue: boolean[] = [];
    for (const d of defects) {
      const type = d.recordType === "observation" ? "Obs" : "Defect";
      rows.push([d.uid, type, formatDefectLocation(d, dims), workTypeLabelFor(d), d.assignedTo || "\u2014", d.dueDate || "\u2014", statusText(d)]);
      overdue.push(isDefectOverdue(d));
      for (const loc of d.locations || []) {
        rows.push([loc.uid || "", type, loc.uid ? deriveLocation(loc.uid) : "", workTypeLabelFor(d), d.assignedTo || "\u2014", d.dueDate || "\u2014", statusText(d)]);
        overdue.push(isDefectOverdue(d));
      }
    }
    if (rows.length === 0) return;
    // 26 + 13 + 48.6 + 34 + 24 + 19 + 20 = 184.6 mm
    autoTable(doc, {
      startY: y, head: [["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"]], body: rows,
      margin: { left: margin, right: margin, top: TOP }, styles: tableStyles, headStyles: tableHead,
      columnStyles: { 0: { cellWidth: 26, fontStyle: "bold" }, 1: { cellWidth: 13 }, 2: { cellWidth: 48.6 }, 3: { cellWidth: 34 }, 4: { cellWidth: 24 }, 5: { cellWidth: 19 }, 6: { cellWidth: 20 } },
      ...tableHook,
      didDrawCell: (h: any) => {
        if (h.section !== "body" || h.column.index !== 6 || !overdue[h.row.index]) return;
        const base = Array.isArray(h.cell.text) ? h.cell.text.join("") : String(h.cell.text ?? "");
        const pos = h.cell.getTextPos();
        const w = doc.getTextWidth(base);
        doc.setTextColor(...OVERDUE_RED);
        doc.text(" - Overdue", pos.x + w, pos.y, { baseline: "top" });
        doc.setTextColor(0);
      },
    });
    afterTable();
  };

  const summaryLine = (g: Extract<CategoryGroup, { kind: "summary" }>) => {
    para(g.note ? `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}. ${g.note}.` : `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}.`);
  };

  // ===================== 2. PROJECT STATUS =====================
  newPage();
  h1("2");
  if (tree.progressSummary) {
    const ps = tree.progressSummary;
    h2("2.progress");
    autoTable(doc, {
      startY: y,
      body: [["Total", String(ps.total)], ["Open", String(ps.open)], ["Closed this period", String(ps.closedThisPeriod)], ["Overdue", String(ps.overdue)]],
      margin: { left: margin, right: margin, top: TOP }, styles: { font: FONT, fontSize: 9, cellPadding: 2, textColor: 0 },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: 60 }, 1: { cellWidth: 30 } }, theme: "grid", ...tableHook,
    });
    afterTable(8); // space after the 2.1 table (change 5)
  }
  h2("2.action");
  para("Based on our observations, we recommend the following actions.");
  if (v2.actionList.length === 0) para("Nothing for this report.");
  for (const g of v2.actionList) {
    if (g.kind === "summary") { summaryLine(g); continue; }
    h3(g.label, 24);
    actionTable(g.defects);
  }
  if (tree.projectStatus.narratives.length > 0) h2("2.narratives");
  tree.projectStatus.narratives.forEach((n: any, i: number) => {
    h3n(`2.narrative.${i}`);
    if (n.status) para(`Status: ${n.status}`);
    if (n.body) String(n.body).split(/\n\s*\n/).forEach((p) => para(p.trim()));
  });
  const addPicture = async (filename: string) => {
    const blob = await loadImageBlob(filename);
    if (!blob) return;
    // Program / Stage Map: full page width at the image's true proportions (never
    // stretched), shrunk only if it would be taller than a page.
    let url: string; let ratio: number; let fmt = "JPEG";
    try {
      const im = await compressImageWithSize(blob, 2400, 0.9);
      url = im.dataUrl; ratio = im.height / im.width;
    } catch {
      url = await blobToDataUrl(blob);
      const props = doc.getImageProperties(url);
      ratio = props.height / props.width; fmt = props.fileType || "JPEG";
    }
    let w = cw(); let h = w * ratio;
    const maxH = limit() - TOP - 4;
    if (h > maxH) { h = maxH; w = h / ratio; }
    ensure(h + 4);
    doc.addImage(url, fmt, margin, y, w, h);
    y += h + PARA_GAP;
  };
  if (tree.projectStatus.program) {
    const p = tree.projectStatus.program;
    h2("2.program");
    const rows: string[][] = [];
    if (p.asAtDate) rows.push(["As at", String(p.asAtDate)]);
    if (p.varianceText) rows.push(["Variance", String(p.varianceText)]);
    if (p.projectedCompletion) rows.push(["Projected completion", String(p.projectedCompletion)]);
    if (p.statusNarrative) rows.push(["Status", String(p.statusNarrative)]);
    if (rows.length) {
      autoTable(doc, { startY: y, body: rows, margin: { left: margin, right: margin, top: TOP }, styles: { font: FONT, fontSize: 9, cellPadding: 1.5, textColor: 0 }, columnStyles: { 0: { fontStyle: "bold", cellWidth: 45 }, 1: { cellWidth: cw() - 45 } }, theme: "plain", ...tableHook });
      afterTable();
    }
    if (p.programImageFilename) await addPicture(p.programImageFilename);
  }
  if (tree.projectStatus.stageMap) {
    const sm = tree.projectStatus.stageMap;
    h2("2.stagemap");
    if (sm.planImageFilename) await addPicture(sm.planImageFilename);
    let stages: any[] = [];
    try { stages = Array.isArray(sm.stages) ? sm.stages : JSON.parse(sm.stages || "[]"); } catch { stages = []; }
    if (stages.length) {
      autoTable(doc, { startY: y, head: [["Stage", "Status"]], body: stages.map((s: any) => [String(s.stageName || ""), String(s.status || "")]), margin: { left: margin, right: margin, top: TOP }, styles: { font: FONT, fontSize: 8, cellPadding: 1.5, textColor: 0 }, theme: "grid", ...tableHook });
      afterTable();
    }
  }

  // ===================== 3. PROJECT INSTRUCTIONS =====================
  newPage();
  h1("3");
  para(INSTRUCTION_SUBTITLE);
  {
    const rows = instructionRows(data);
    if (rows.length === 0) para("No project instructions issued.");
    else {
      // Full width; Work Type(s), Instruction and Action Required widened (change 8).
      // 16 + 18 + 28 + 43.6 + 42 + 21 + 16 = 184.6 mm
      autoTable(doc, {
        startY: y, head: [INSTRUCTION_HEADERS], body: rows, margin: { left: margin, right: margin, top: TOP },
        styles: tableStyles, headStyles: tableHead,
        columnStyles: { 0: { cellWidth: 16, fontStyle: "bold" }, 1: { cellWidth: 18 }, 2: { cellWidth: 28 }, 3: { cellWidth: 43.6 }, 4: { cellWidth: 42 }, 5: { cellWidth: 21 }, 6: { cellWidth: 16 } },
        rowPageBreak: "avoid", ...tableHook,
        didParseCell: (h: any) => { if (h.section === "body" && h.column.index === 6 && h.cell.raw !== "Ongoing") h.cell.styles.textColor = [136, 136, 136]; },
      });
      afterTable();
    }
  }

  // ===================== Records (sections 4 and 5) =====================
  const renderPhotos = async (list: any[], added?: Set<number>) => {
    if (!list.length) return;
    const hasWip = list.some((p: any) => typeof p.wipNumber === "number");
    let photos: any[];
    if (hasWip) photos = list;
    else {
      const seen = new Set<string>();
      const inSlot = [...list].sort((a: any, b: any) => comparePhotoSlots(a.slot, b.slot)).filter((p: any) => (seen.has(p.slot) ? false : (seen.add(p.slot), true)));
      const ids = new Set(inSlot.map((p: any) => p.id));
      photos = [...inSlot, ...list.filter((p: any) => !ids.has(p.id))];
    }
    const gap = 6;
    const w = (cw() - gap) / 2; const h = w * 0.75;
    ensure(7 + h + 6);
    y += 3;
    doc.setFontSize(9); doc.setFont(FONT, "bold"); doc.setTextColor(0);
    doc.text("Photos", margin, y); y += 3;
    for (let i = 0; i < photos.length; i++) {
      const p = photos[i]; const col = i % 2;
      if (i > 0 && col === 0) y += h + 8;
      if (col === 0 && y + h + 8 > limit()) newPage();
      const x = margin + col * (w + gap);
      const blob = await loadImageBlob(p.filename);
      if (blob) {
        let url: string;
        try { url = await compressImageForPdfExport(blob, 800, 0.7); } catch { url = await blobToDataUrl(blob); }
        doc.addImage(url, "JPEG", x, y, w, h);
      } else {
        doc.setDrawColor(180); doc.rect(x, y, w, h); doc.setFontSize(7); doc.text("[Photo unavailable]", x + 4, y + h / 2);
      }
      doc.setFontSize(7); doc.setFont(FONT, "bold"); doc.setTextColor(100);
      const label = p.slot === "complete" ? photoSlotLabel(p.slot) : (typeof p.wipNumber === "number" ? `WIP ${p.wipNumber}` : photoSlotLabel(p.slot));
      const caption = p.caption ? ` \u2014 ${p.caption}` : "";
      const date = added && added.has(p.id) && p.createdAt ? ` (added ${formatPhotoDate(p.createdAt)})` : "";
      doc.text(doc.splitTextToSize(label + caption + date, w)[0], x, y + h + 3.5);
      doc.setTextColor(0);
    }
    y += h + 8;
  };

  // Rough height of a two-column detail table, so a record's heading and details are
  // kept together on one page (its photos may run on).
  const detailHeight = (rows: string[][], w0: number, w1: number, size: number) => {
    doc.setFontSize(size);
    const lh = size * 0.3528 * 1.2;
    let total = 0;
    for (const [a, b] of rows) {
      doc.setFont(FONT, "bold"); const la = doc.splitTextToSize(String(a), w0 - 3).length;
      doc.setFont(FONT, "normal"); const lb = doc.splitTextToSize(String(b), w1 - 3).length;
      total += Math.max(la, lb) * lh + 2.4;
    }
    return total;
  };

  // One record, flowing on from the previous one. A bold rule at the header/footer width
  // separates records (change 9 / answer 7); none at the top of a page.
  const renderRecord = async (d: any, opts: { note?: boolean; rule?: boolean } = {}) => {
    const asNote = !!opts.note;
    const multi = Array.isArray(d.locations) && d.locations.length > 0;
    const added = d.events ? new Set<number>(d.events.photosAddedThisInspection || []) : undefined;
    const info: string[][] = asNote
      ? [["Date Recorded", d.dateOpened || "\u2014"], ["Observation", d.comment || "\u2014"], ["Action Required", "No action required"]]
      : [
          ["Date Opened", d.dateOpened || "\u2014"],
          ["Date Completed", d.dateClosed || "\u2014"],
          ["Observation", d.comment || "\u2014"],
          ["Action Required", d.actionRequired || "\u2014"],
          ["Assigned To", d.assignedTo || "\u2014"],
          ["Due Date", d.dueDate || "\u2014"],
          ["Verification Method", d.verificationMethod || "\u2014"],
          ["Verification Person", d.verificationPerson || "\u2014"],
        ];
    const W0 = 40;
    const ruleH = opts.rule ? 7 : 0;
    const need = ruleH + 12 + detailHeight(info, W0, cw() - W0, 9);
    if (y + need > limit() && y > TOP + 1) newPage();
    if (opts.rule && y > TOP + 1) {
      y += 1.5;
      doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.8);
      doc.line(margin, y, pw() - margin, y);
      y += 6;
    }
    y += 4;
    recordDests.set(d.id, { page: pageNo(), top: y - 6 });
    const typeLabel = asNote ? "NOTE" : d.recordType === "observation" ? "OBSERVATION" : "DEFECT";
    const typeBg = asNote ? [148, 163, 184] : d.recordType === "observation" ? [59, 130, 246] : [217, 119, 6];
    doc.setFontSize(8); doc.setFont(FONT, "bold");
    const badgeW = doc.getTextWidth(typeLabel) + 6;
    doc.setFillColor(typeBg[0], typeBg[1], typeBg[2]);
    doc.roundedRect(margin, y - 4, badgeW, 6, 1.5, 1.5, "F");
    doc.setTextColor(255, 255, 255);
    doc.text(typeLabel, margin + 3, y);
    doc.setFontSize(14); doc.setTextColor(...DARK_TEXT);
    doc.text(multi ? `Multiple Entries for ${workTypeLabelFor(d)}` : d.uid, margin + badgeW + 4, y);
    if (!asNote) {
      const st = d.status === "complete" ? "COMPLETE" : "OPEN";
      doc.setFontSize(10);
      if (d.status === "complete") doc.setTextColor(34, 139, 34); else doc.setTextColor(200, 150, 0);
      doc.text(st, pw() - margin - doc.getTextWidth(st), y);
    }
    y += 5;
    const where = multi ? "" : formatDefectLocation(d, dims);
    if (where) {
      doc.setFontSize(9); doc.setFont(FONT, "normal"); doc.setTextColor(80);
      doc.text(where, margin, y); y += 3;
    }
    autoTable(doc, {
      startY: y, body: info, margin: { left: margin, right: margin, top: TOP },
      styles: { font: FONT, fontSize: 9, cellPadding: { top: 1.1, bottom: 1.1, left: 0, right: 2 }, overflow: "linebreak", valign: "top", textColor: 0 },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: W0 }, 1: { cellWidth: cw() - W0 } },
      theme: "plain", ...tableHook,
    });
    afterTable(3);
    if (multi) {
      const all = [
        { uid: d.uid, description: d.comment, elevation: "", photos: d.photos || [] },
        ...d.locations.map((l: any) => ({ uid: l.uid || "", description: l.description || "", elevation: l.elevation || "", photos: [] as any[] })),
      ];
      for (let i = 0; i < all.length; i++) {
        const loc = all[i];
        ensure(20);
        doc.setFontSize(11); doc.setFont(FONT, "bold"); doc.setTextColor(...CAPTION_BLUE);
        doc.text(loc.uid || "\u2014", margin, y); y += 5;
        if (loc.uid) { doc.setFontSize(8); doc.setFont(FONT, "normal"); doc.setTextColor(100); doc.text(deriveLocation(loc.uid), margin, y); y += 4; }
        if (i > 0 && loc.description) { body(9); const dl: string[] = doc.splitTextToSize(loc.description, cw()); doc.text(dl, margin, y); y += dl.length * 4 + 2; }
        if (loc.elevation) { doc.setFontSize(7.5); doc.setFont(FONT, "italic"); doc.setTextColor(120); doc.text(`Elevation: ${ELEVATION_NAMES[loc.elevation] || loc.elevation}`, margin, y); y += 4; }
        if (loc.photos?.length) await renderPhotos(loc.photos, added);
        y += 2;
      }
    } else {
      await renderPhotos(d.photos || [], added);
    }
  };

  // Bold rule above a category heading too, so every change of record is marked
  // (none at the top of a section or a page).
  const groupRule = (first: boolean) => {
    if (first) return;
    if (y + 48 > limit()) { newPage(); return; }
    if (y > TOP + 1) { y += 1.5; doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.8); doc.line(margin, y, pw() - margin, y); y += 4; }
  };
  const renderGroups = async (groups: CategoryGroup[]) => {
    for (let gi = 0; gi < groups.length; gi++) {
      const g = groups[gi];
      groupRule(gi === 0);
      h3(g.label, 40);
      if (g.kind === "summary") { summaryLine(g); continue; }
      for (let i = 0; i < g.defects.length; i++) await renderRecord(g.defects[i], { rule: i > 0 });
      y += 2;
    }
  };

  // ===================== 4. NEW THIS INSPECTION =====================
  newPage();
  h1("4");
  if (v2.newThisInspection.length === 0 && v2.notes.length === 0) para("No new records this inspection.");
  await renderGroups(v2.newThisInspection);
  if (v2.notes.length > 0) {
    groupRule(v2.newThisInspection.length === 0);
    h3("Notes", 40);
    for (let i = 0; i < v2.notes.length; i++) await renderRecord(v2.notes[i], { note: true, rule: i > 0 });
  }

  // ===================== 5. CLOSED THIS INSPECTION =====================
  newPage();
  h1("5");
  if (v2.closedThisInspection.length === 0) para("No records were closed this inspection.");
  await renderGroups(v2.closedThisInspection);

  // ===================== 6. RUNNING LIST OF OPEN ITEMS =====================
  newPage();
  h1("6");
  para("All open items from previous inspections.");
  if (v2.runningOpen.length === 0) para("No open items from previous inspections.");
  for (const g of v2.runningOpen) {
    if (g.kind === "summary") { summaryLine(g); continue; }
    h3(g.label, 24);
    registerTable(g.defects);
  }

  // ===================== 7. ELEVATIONS (A4 landscape) =====================
  if (elevations.length === 0) {
    newPage();
    h1("7");
    para("No elevation drawings have pins for records in this report.");
  }
  elevations.forEach((el, i) => {
    newPage("l");
    if (i === 0) h1("7");
    h2(`7.elevation.${el.id}`);
    const legendH = 7;
    const boxW = cw(); const boxH = limit() - y - legendH;
    const k = Math.min(boxW / el.width, boxH / el.height);
    const w = el.width * k; const h = el.height * k;
    doc.addImage(el.dataUrl, "JPEG", margin + (boxW - w) / 2, y, w, h);
    doc.setDrawColor(200); doc.setLineWidth(0.2); doc.rect(margin + (boxW - w) / 2, y, w, h);
    let lx = margin; const ly = y + h + 5;
    doc.setFontSize(8); doc.setFont(FONT, "normal");
    for (const [label, key] of [["Open", "open"], ["Closed", "complete"], ["Note", "note"]] as const) {
      const hex = PIN_COLOURS[key];
      doc.setFillColor(parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16));
      doc.circle(lx + 1.5, ly - 1, 1.5, "F");
      doc.setTextColor(0); doc.text(label, lx + 4.5, ly);
      lx += 6 + doc.getTextWidth(label) + 6;
    }
    doc.text(`${el.pins.length} pin${el.pins.length === 1 ? "" : "s"}`, pw() - margin, ly, { align: "right" });
  });

  // ===================== APPENDICES (last; no section number) =====================
  // Divider page per appendix: "Appendix A" (heading 1) then its name (heading 2).
  for (const e of outline) {
    if (!/^appendix\.\d+$/.test(e.key)) continue;
    newPage();
    h1(e.key);
    h2(`${e.key}.name`);
  }

  // ===================== Fill in contents, links and bookmarks =====================
  for (let p = 0; p < tocPageCount; p++) {
    doc.setPage(tocFirstPage + p);
    let ty = TOP;
    if (p === 0) {
      doc.setFont(FONT, "bold"); doc.setFontSize(18); doc.setTextColor(...DARK_TEXT);
      ty += 2; doc.text("Contents", margin, ty); ty += 12;
    }
    for (const e of tocEntries.slice(p * tocRowsPerPage, (p + 1) * tocRowsPerPage)) {
      const dest = dests.get(e.key);
      const top = e.level === 1;
      const x = margin + (top ? 0 : HANG);
      doc.setFont(FONT, top ? "bold" : "normal"); doc.setFontSize(10); doc.setTextColor(0, 0, 0);
      const label = e.no.startsWith("Appendix") ? (e.title ? `${e.no} \u2013 ${e.title}` : e.no) : e.no;
      const titleX = e.no.startsWith("Appendix") || !e.no ? x : x + HANG;
      if (top && ty > TOP + 14) ty += 1.5;
      doc.text(e.no.startsWith("Appendix") ? label : e.no, x, ty);
      const title = e.no.startsWith("Appendix") ? "" : e.title;
      const pageText = dest ? String(dest.page) : "";
      const maxTitle = pw() - margin - 12 - titleX;
      const shown = title ? doc.splitTextToSize(title, maxTitle)[0] : "";
      if (shown) doc.text(shown, titleX, ty);
      // Dot leader to the page number.
      const endX = (shown ? titleX + doc.getTextWidth(shown) : x + doc.getTextWidth(label)) + 2;
      const numW = doc.getTextWidth(pageText);
      doc.setFont(FONT, "normal"); doc.setTextColor(120);
      const dotW = doc.getTextWidth(". ");
      const dots = Math.max(0, Math.floor((pw() - margin - numW - 2 - endX) / dotW));
      if (dots > 0) doc.text(". ".repeat(dots), endX, ty);
      doc.setFont(FONT, top ? "bold" : "normal"); doc.setTextColor(0, 0, 0);
      doc.text(pageText, pw() - margin, ty, { align: "right" });
      if (dest) doc.link(margin, ty - 4, cw(), TOC_ROW, { pageNumber: dest.page, top: dest.top });
      ty += TOC_ROW;
    }
  }
  for (const l of uidLinks) {
    const t = recordDests.get(l.id);
    if (!t) continue;
    doc.setPage(l.page);
    doc.link(l.x, l.y, l.w, l.h, { pageNumber: t.page, top: t.top });
  }
  // PDF bookmarks mirror the contents page.
  try {
    let parent: any = null; let parent2: any = null;
    const tocDest = { pageNumber: tocFirstPage };
    doc.outline.add(null, "Contents", tocDest);
    for (const e of outline) {
      const dest = dests.get(e.key);
      if (!dest) continue;
      const text = outlineHeading(outline, e.key).text;
      if (e.level === 1) { parent = doc.outline.add(null, text, { pageNumber: dest.page }); parent2 = null; }
      else if (e.level === 2) parent2 = doc.outline.add(parent, text, { pageNumber: dest.page });
      else doc.outline.add(parent2 || parent, text, { pageNumber: dest.page });
    }
  } catch (err) {
    console.warn("[render-pdf-v2] bookmarks skipped", err);
  }

  return doc.output("blob");
}
