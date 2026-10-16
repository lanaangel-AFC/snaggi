// ============================================================================
// render-docx-v2.ts — WP8 Word layout for NEW projects (featureVersion >= 2).
// Same structure and numbering as render-pdf-v2.ts (both read report-outline-v2.ts).
// Older projects never reach this file, so their Word exports are unchanged.
//
//   Fonts: Aptos throughout. Body text 10 pt, automatic colour, single line spacing,
//   one blank line (10 pt) after each paragraph. Level 2 headings: left 0 cm, hanging
//   1.27 cm, 2 pt before, 6 pt after, no page break before.
//   Contents: a Word contents field with links. Word fills it in when it asks to
//   update fields on opening the file.
// ============================================================================

import type { ReportTree, CategoryGroup } from "./report-tree";
import {
  safeText, loadImageBlob, blobToArrayBuffer, compressImageForExport, loadAfcLogo,
  getWorkTypeLabel, ELEVATION_NAMES, deriveLocation, formatDefectLocation,
  getLocationDimensions, formatReportDate, formatPhotoDate, isDefectOverdue,
  comparePhotoSlots, photoSlotLabel, resolveActionSummary,
  parseProjectSnapshot, resolveProjectField, deriveAreaRef,
} from "./render-helpers";
import { INSTRUCTION_HEADERS, INSTRUCTION_SUBTITLE, instructionRows } from "./project-instructions";
import { buildOutlineV2, outlineHeading, scopeRowsV2, backgroundDocsV2, type OutlineEntry } from "./report-outline-v2";
import { loadReportElevations, PIN_COLOURS, type PrintedElevation } from "./elevation-print";
import { LIMITATIONS_V2 } from "./render-pdf-v2";

function dataUrlToBuffer(url: string): ArrayBuffer {
  const b64 = url.slice(url.indexOf(",") + 1);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out.buffer;
}

export async function renderDocxV2(tree: ReportTree): Promise<Blob> {
  const v2 = tree.v2!;
  const data = { project: tree.project, report: tree.report, projectInstructions: tree.projectInstructions };
  const dims = getLocationDimensions(tree.project?.locationDimensions);

  const docxLib = await import("docx");
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
    AlignmentType, WidthType, BorderStyle, ShadingType, TableLayoutType,
    Header, Footer, PageNumber, TabStopType, TableOfContents, BookmarkStart, BookmarkEnd, InternalHyperlink,
    PageOrientation, HeadingLevel,
  } = docxLib as any;

  const FONT = "Aptos";
  const AUTO = "auto";
  const DARK_BLUE = "0A1D30";
  const ACCENT_BLUE = "00CDC8";
  const DARK_TEXT = "3A3A3A";
  const CAPTION_BLUE = "0E2841";
  const OVERDUE_RED = "C00000";
  const PAGE_W = 11906; const PAGE_H = 16838; const MARGIN = 720; const V_MARGIN = 1134; // 2 cm top/bottom clears the header and footer
  const CONTENT_W = PAGE_W - MARGIN * 2; // 10466 twips, the header/footer width
  const CONTENT_W_LANDSCAPE = PAGE_H - MARGIN * 2;

  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const rowBorder = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
  const noBorders = { top: noBorder, left: noBorder, right: noBorder, bottom: noBorder };

  const snap = parseProjectSnapshot((data.report as any).projectSnapshot);
  const projAddress = resolveProjectField(snap, data.project, "address");
  const projName = resolveProjectField(snap, data.project, "name");
  const projReportTitle = resolveProjectField(snap, data.project, "reportTitle");
  const projInspector = resolveProjectField(snap, data.project, "inspector");
  const projClient = resolveProjectField(snap, data.project, "client");
  const afcRef = resolveProjectField(snap, data.project, "afcReference") || "AFC-24XXX";
  const areaRefTemplate: string = String(resolveProjectField(snap, data.project, "areaRefTemplate") || "");
  const inspNumRaw = safeText(data.report.inspectionNumber).trim();
  const inspNum = inspNumRaw ? (/^\d+$/.test(inspNumRaw) ? inspNumRaw.padStart(2, "0") : inspNumRaw) : "";

  const elevations: PrintedElevation[] = await loadReportElevations(tree.project.id, tree.report.id, v2.elevationRecordIds);
  const outline: OutlineEntry[] = buildOutlineV2(tree, elevations);
  const logo = await loadAfcLogo();

  // ---- Text helpers ----
  const run = (text: string, o: any = {}) => new TextRun({ text: safeText(text), font: FONT, ...o });
  const para = (text: string, o: any = {}) => new Paragraph({ children: [run(text)], ...o });
  const h1 = (key: string) => new Paragraph({ text: outlineHeading(outline, key).text, heading: HeadingLevel.HEADING_1 });
  const h2 = (key: string) => {
    const h = outlineHeading(outline, key);
    return new Paragraph({ text: `${h.no}\t${h.title}`, heading: HeadingLevel.HEADING_2 });
  };
  // Unnumbered sub-heading (category groups, "Unique identifier", "Notes").
  const h3 = (label: string) => new Paragraph({ text: safeText(label), heading: HeadingLevel.HEADING_3 });

  // ---- Header / footer at the page width of the section ----
  const header = () => new Header({
    children: [new Table({
      width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE } },
          children: [new Paragraph({ spacing: { after: 0 }, children: [run(safeText(projName).toUpperCase(), { size: 16, bold: true, color: DARK_TEXT })] })] }),
        new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE } },
          children: [new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.RIGHT, children: [run("ANGEL FA\u00c7ADE CONSULTING", { size: 16, color: DARK_TEXT })] })] }),
      ] })],
    })],
  });
  const footer = (width: number) => new Footer({
    children: [new Paragraph({
      spacing: { after: 0 },
      tabStops: [{ type: TabStopType.RIGHT, position: width }],
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE } },
      children: [
        run(`${safeText(afcRef)} | ${safeText(projAddress)}`, { size: 16, color: DARK_TEXT }),
        new TextRun({ text: "\t" }),
        run("Page ", { size: 16, color: DARK_TEXT }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, font: FONT, color: DARK_TEXT }),
      ],
    })],
  });
  // Headers/footers belong on the section itself, not inside its properties (the
  // library silently drops them there).
  const portrait = { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: V_MARGIN, right: MARGIN, bottom: V_MARGIN, left: MARGIN, header: 425, footer: 425 } } };
  const landscape = { page: { size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE }, margin: { top: V_MARGIN, right: MARGIN, bottom: V_MARGIN, left: MARGIN, header: 425, footer: 425 } } };
  const portraitHF = { headers: { default: header() }, footers: { default: footer(CONTENT_W) } };
  const landscapeHF = { headers: { default: header() }, footers: { default: footer(CONTENT_W_LANDSCAPE) } };

  // ---- Tables ----
  type Cell = string | any[]; // plain text, or ready-made runs/links
  const gridTable = (headers: string[], widths: number[], rows: Cell[][], o: { boldFirst?: boolean; cantSplit?: boolean } = {}) => new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((label, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_TEXT } },
        children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [run(label, { size: 15, bold: true, color: CAPTION_BLUE })] })],
      })) }),
      ...rows.map((cells) => new TableRow({ cantSplit: o.cantSplit ?? true, children: cells.map((c, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        borders: { top: rowBorder, left: noBorder, right: noBorder, bottom: rowBorder },
        children: [new Paragraph({ spacing: { before: 25, after: 25 }, children: Array.isArray(c) ? c : [run(c, { size: 15, bold: i === 0 && o.boldFirst !== false })] })],
      })) })),
    ],
  });
  const statusRuns = (d: any) => {
    const text = d.status === "complete" ? "Complete" : "Open";
    const runs = [run(text, { size: 15, color: d.status === "complete" ? "228B22" : "C89600" })];
    if (isDefectOverdue(d)) runs.push(run(" - Overdue", { size: 15, color: OVERDUE_RED }));
    return runs;
  };
  const anchorFor = (id: number) => `rec_${id}`;
  const detailIds = new Set(v2.detailIds);
  const uidCell = (uid: string, d: any): Cell => (detailIds.has(d.id)
    ? [new InternalHyperlink({ anchor: anchorFor(d.id), children: [run(uid, { size: 15, bold: true, style: "Hyperlink" })] })]
    : uid);

  // 2.2 Action List at the header/footer width, Work Type and Action widened.
  const ACTION_W = [1474, 964, 1587, 2982, 1191, 1077, 1191];
  const actionTable = (defects: any[]) => {
    const rows: Cell[][] = [];
    for (const d of defects) {
      const area = deriveAreaRef(d, areaRefTemplate) || "";
      rows.push([uidCell(safeText(d.uid), d), area, getWorkTypeLabel(safeText(d.uid)), resolveActionSummary(d), safeText(d.assignedTo) || "\u2014", safeText(d.dueDate) || "\u2014", statusRuns(d)]);
      for (const loc of d.locations || []) rows.push([uidCell(safeText(loc.uid), d), area, getWorkTypeLabel(safeText(d.uid)), "", safeText(d.assignedTo) || "\u2014", safeText(d.dueDate) || "\u2014", statusRuns(d)]);
    }
    return gridTable(["ID", "Location", "Work Type", "Action", "Responsible", "By Date", "Status"], ACTION_W, rows);
  };
  const REGISTER_W = [1474, 737, 2755, 1928, 1361, 1077, 1134];
  const registerTable = (defects: any[]) => {
    const rows: Cell[][] = [];
    for (const d of defects) {
      const type = d.recordType === "observation" ? "Obs" : "Defect";
      rows.push([safeText(d.uid), type, formatDefectLocation(d, dims), getWorkTypeLabel(safeText(d.uid)), safeText(d.assignedTo) || "\u2014", safeText(d.dueDate) || "\u2014", statusRuns(d)]);
      for (const loc of d.locations || []) rows.push([safeText(loc.uid), type, loc.uid ? deriveLocation(safeText(loc.uid)) : "", getWorkTypeLabel(safeText(d.uid)), safeText(d.assignedTo) || "\u2014", safeText(d.dueDate) || "\u2014", statusRuns(d)]);
    }
    return gridTable(["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"], REGISTER_W, rows);
  };
  const keyValueTable = (rows: Array<[string, string]>, labelPct: number, o: { tight?: boolean; keepNext?: boolean; size?: number } = {}) => {
    const LW = Math.round((CONTENT_W * labelPct) / 100); const VW = CONTENT_W - LW;
    return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [LW, VW], layout: TableLayoutType.FIXED,
    margins: { left: 0, right: 0 },
    rows: rows.map(([label, value], ri) => new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: LW, type: WidthType.DXA }, borders: noBorders,
        children: [new Paragraph({ keepNext: o.keepNext && ri < rows.length - 1, spacing: { before: o.tight ? 0 : 40, after: o.tight ? 0 : 40 }, children: [run(label, { bold: true, size: o.size ?? 20 })] })] }),
      new TableCell({ width: { size: VW, type: WidthType.DXA }, borders: noBorders,
        children: safeText(value).split(/\n/).map((line) => new Paragraph({ keepNext: o.keepNext && ri < rows.length - 1, spacing: { before: o.tight ? 0 : 40, after: o.tight ? 0 : 40 }, children: [run(line, { size: o.size ?? 20 })] })) }),
    ] })),
    });
  };
  const spacer = () => new Paragraph({ spacing: { after: 0 }, children: [] });
  const summaryLine = (g: Extract<CategoryGroup, { kind: "summary" }>) =>
    para(g.note ? `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}. ${g.note}.` : `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}.`);

  const sections: any[] = [];

  // ===================== TITLE PAGE =====================
  {
    const c: any[] = [];
    if (logo) {
      let buf: ArrayBuffer;
      try { buf = await compressImageForExport(new Blob([logo.buffer]), 360, 0.8); } catch { buf = logo.buffer; }
      c.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 600 }, children: [new ImageRun({ data: buf, transformation: { width: 180, height: 58 }, type: "jpg" })] }));
    }
    c.push(new Paragraph({ spacing: { before: 1200, after: 0 } }));
    c.push(new Paragraph({ spacing: { after: 120 }, children: [run(safeText(projAddress).toUpperCase(), { bold: true, size: 56, color: DARK_BLUE })] }));
    if (projReportTitle) c.push(new Paragraph({ spacing: { after: 120 }, children: [run(projReportTitle, { size: 36, color: DARK_BLUE })] }));
    c.push(new Paragraph({ spacing: { after: 100 }, children: [run(inspNum ? `SITE VISIT REPORT ${inspNum}` : "SITE VISIT REPORT", { size: 44, color: DARK_BLUE, smallCaps: true })] }));
    c.push(new Paragraph({ spacing: { after: 600 }, children: [run(`Revision ${safeText(data.report.revision || "01")}`, { size: 28, color: DARK_TEXT })] }));
    c.push(new Paragraph({ spacing: { after: 400 }, children: [run(formatReportDate(new Date().toISOString()), { size: 24, color: DARK_TEXT })] }));
    c.push(new Paragraph({ spacing: { after: 40 }, children: [run("Angel Fa\u00e7ade Consulting", { size: 22, color: DARK_TEXT })] }));
    c.push(new Paragraph({ spacing: { after: 40 }, children: [run(`${safeText(projInspector)} | 0407 759 590`, { size: 22, color: DARK_TEXT })] }));
    c.push(new Paragraph({ spacing: { after: 200 }, children: [run(afcRef, { size: 22, color: DARK_TEXT })] }));
    sections.push({ properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 851, right: 1440, bottom: 1440, left: 1440 } } }, children: c });
  }

  // ===================== CONTENTS =====================
  sections.push({ properties: portrait, ...portraitHF, children: [
    new Paragraph({ spacing: { after: 240 }, children: [run("Contents", { bold: true, size: 40, color: DARK_TEXT })] }),
    new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  ] });

  // ===================== 1. INTRODUCTION =====================
  {
    const c: any[] = [h1("1"), h2("1.general")];
    c.push(para(`Angel Fa\u00e7ade Consulting (AFC) was engaged by ${safeText(projClient)} to carry out regular inspections of the remedial works underway at ${safeText(projAddress)}. Below is a summary of pertinent project information.`));
    let roles: Array<{ role?: string; entity?: string; contactDetails?: string }> = [];
    try { roles = JSON.parse(resolveProjectField(snap, data.project, "roles") || "[]") || []; } catch { roles = []; }
    if (Array.isArray(roles) && roles.length > 0) {
      const W = [Math.round(CONTENT_W * 0.25), Math.round(CONTENT_W * 0.3)]; W.push(CONTENT_W - W[0] - W[1]);
      c.push(gridTable(["Role", "Entity", "Contact Details"], W, roles.map((r) => [safeText(r.role), safeText(r.entity), safeText(r.contactDetails)]), { boldFirst: false }));
      c.push(spacer());
    }
    const scope = scopeRowsV2(tree);
    if (scope.length > 0) {
      c.push(h2("1.scope"));
      const W = [0.18, 0.27, 0.3].map((f) => Math.round(CONTENT_W * f)); W.push(CONTENT_W - W[0] - W[1] - W[2]);
      c.push(gridTable(["Area Ref", "Location", "Work item", "Access method"], W, scope.map((s: any) => [safeText(s.areaRef), safeText(s.location), safeText(s.workItem), safeText(s.accessMethod)]), { boldFirst: false }));
      c.push(spacer());
    }
    c.push(h2("1.particulars"));
    const rows: Array<[string, string]> = [];
    rows.push(["Date", data.report.inspectionDate ? formatReportDate(data.report.inspectionDate) : formatReportDate(new Date().toISOString())]);
    if (data.report.inspectionNumber) rows.push(["Inspection number", safeText(inspNum || data.report.inspectionNumber)]);
    if (data.report.revision) rows.push(["Revision", safeText(data.report.revision)]);
    rows.push(["Inspector", safeText(projInspector)]);
    rows.push(["Locations covered", safeText(data.report.locationsCovered) || safeText(projAddress)]);
    try {
      const att = JSON.parse(data.report.attendees || "[]");
      if (Array.isArray(att)) {
        const lines = att.map((a: any) => (a.company ? `${safeText(a.name)} (${safeText(a.company)})` : safeText(a.name))).filter((s: string) => s.trim());
        if (lines.length) rows.push(["Attendees", lines.join("\n")]);
      }
    } catch { /* malformed attendees: omit the row */ }
    c.push(keyValueTable(rows, 25, { tight: true })); // tighter rows (change 2)
    c.push(spacer());
    const bg = backgroundDocsV2(tree);
    if (bg.length > 0) {
      c.push(h2("1.background"));
      c.push(para("The following documents have been reviewed and form the basis of this inspection."));
      const yearOf = (d?: string) => { if (!d) return ""; const m = String(d).match(/\b(19|20)\d{2}\b/); return m ? m[0] : String(d); };
      for (const b of bg) {
        const originator = safeText(b.originator); const year = yearOf(b.date); const title = safeText(b.title);
        const tail = [safeText(b.type) ? `[${safeText(b.type)}]` : "", safeText(b.docNumbers) ? `Doc no. ${safeText(b.docNumbers)}` : "", safeText(b.revision) ? `Rev. ${safeText(b.revision)}` : ""].filter(Boolean);
        const head = originator ? (year ? `${originator} (${year}). ` : `${originator}. `) : (year ? `(${year}). ` : "");
        const titleText = title ? `${title}.` : "";
        const tailText = tail.length ? ` ${tail.join(", ")}.` : "";
        if (!`${head}${titleText}${tailText}`.trim()) continue;
        c.push(new Paragraph({ indent: { left: 360, hanging: 360 }, children: [
          ...(head ? [run(head, { bold: true })] : []), ...(titleText ? [run(titleText, { italics: true })] : []), ...(tailText ? [run(tailText)] : []),
        ] }));
      }
    }
    c.push(h2("1.uid"));
    c.push(h3("Unique identifier"));
    c.push(para("Throughout the document, each record (observation or defect) is referred to by its unique identifier (UID)."));
    c.push(para("The UID comprises the following components: Area Ref \u2013 Work item \u2013 Sequential identifier."));
    c.push(h2("1.limitations"));
    for (const b of LIMITATIONS_V2) c.push(new Paragraph({ indent: { left: 284, hanging: 284 }, tabStops: [{ type: TabStopType.LEFT, position: 284 }], children: [run(`\u2022\t${b}`)] }));
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== 2. PROJECT STATUS =====================
  {
    const c: any[] = [h1("2")];
    if (tree.progressSummary) {
      const ps = tree.progressSummary;
      c.push(h2("2.progress"));
      c.push(new Table({
        width: { size: 4600, type: WidthType.DXA }, columnWidths: [3000, 1600], layout: TableLayoutType.FIXED,
        rows: ([["Total", ps.total], ["Open", ps.open], ["Closed this period", ps.closedThisPeriod], ["Overdue", ps.overdue]] as Array<[string, number]>).map(([l, v]) => new TableRow({ children: [
          new TableCell({ width: { size: 3000, type: WidthType.DXA }, borders: { ...noBorders, bottom: thinBorder }, children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [run(l, { bold: true, size: 18 })] })] }),
          new TableCell({ width: { size: 1600, type: WidthType.DXA }, borders: { ...noBorders, bottom: thinBorder }, children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [run(String(v), { size: 18 })] })] }),
        ] })),
      }));
      c.push(new Paragraph({ spacing: { after: 120 }, children: [] })); // space after the 2.1 table (change 5)
    }
    c.push(h2("2.action"));
    c.push(para("Based on our observations, we recommend the following actions."));
    if (v2.actionList.length === 0) c.push(para("Nothing for this report."));
    for (const g of v2.actionList) {
      if (g.kind === "summary") { c.push(summaryLine(g)); continue; }
      c.push(h3(g.label));
      c.push(actionTable(g.defects));
    }
    c.push(spacer());
    const picture = async (filename: string) => {
      const blob = await loadImageBlob(filename);
      if (!blob) return null;
      let buf: ArrayBuffer;
      try { buf = await compressImageForExport(blob, 1000, 0.7); } catch { buf = await blobToArrayBuffer(blob); }
      return new Paragraph({ children: [new ImageRun({ data: buf, transformation: { width: 600, height: 336 }, type: "jpg" })] });
    };
    tree.projectStatus.narratives.forEach((n: any, i: number) => {
      c.push(h2(`2.narrative.${i}`));
      if (n.status) c.push(para(`Status: ${safeText(n.status)}`));
      if (n.body) String(n.body).split(/\n\s*\n/).forEach((p) => c.push(para(p.trim())));
    });
    if (tree.projectStatus.program) {
      const p = tree.projectStatus.program;
      c.push(h2("2.program"));
      const rows: Array<[string, string]> = [];
      if (p.asAtDate) rows.push(["As at", safeText(p.asAtDate)]);
      if (p.varianceText) rows.push(["Variance", safeText(p.varianceText)]);
      if (p.projectedCompletion) rows.push(["Projected completion", safeText(p.projectedCompletion)]);
      if (p.statusNarrative) rows.push(["Status", safeText(p.statusNarrative)]);
      if (rows.length) { c.push(keyValueTable(rows, 30, { size: 18 })); c.push(spacer()); }
      if (p.programImageFilename) { const pic = await picture(p.programImageFilename); if (pic) c.push(pic); }
    }
    if (tree.projectStatus.stageMap) {
      const sm = tree.projectStatus.stageMap;
      c.push(h2("2.stagemap"));
      if (sm.planImageFilename) { const pic = await picture(sm.planImageFilename); if (pic) c.push(pic); }
      let stages: any[] = [];
      try { stages = Array.isArray(sm.stages) ? sm.stages : JSON.parse(sm.stages || "[]"); } catch { stages = []; }
      if (stages.length) c.push(gridTable(["Stage", "Status"], [Math.round(CONTENT_W * 0.6), CONTENT_W - Math.round(CONTENT_W * 0.6)], stages.map((s: any) => [safeText(s.stageName), safeText(s.status)]), { boldFirst: false }));
    }
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== 3. PROJECT INSTRUCTIONS =====================
  {
    const c: any[] = [h1("3"), para(INSTRUCTION_SUBTITLE)];
    const rows = instructionRows(data);
    if (rows.length === 0) c.push(para("No project instructions issued."));
    else {
      // Full width; Work Type(s), Instruction and Action Required widened (change 8).
      const W = [907, 1077, 1587, 2268, 2324, 1361, 942];
      c.push(gridTable(INSTRUCTION_HEADERS, W, rows.map((cells) => cells.map((t, i) => (i === 6 && t !== "Ongoing" ? [run(t, { size: 15, color: "888888" })] : t)))));
    }
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== Records (sections 4 and 5) =====================
  const photosBlock = async (list: any[], added?: Set<number>) => {
    const out: any[] = [];
    if (!list.length) return out;
    const hasWip = list.some((p: any) => typeof p.wipNumber === "number");
    let photos: any[];
    if (hasWip) photos = list;
    else {
      const seen = new Set<string>();
      const inSlot = [...list].sort((a: any, b: any) => comparePhotoSlots(a.slot, b.slot)).filter((p: any) => (seen.has(p.slot) ? false : (seen.add(p.slot), true)));
      const ids = new Set(inSlot.map((p: any) => p.id));
      photos = [...inSlot, ...list.filter((p: any) => !ids.has(p.id))];
    }
    out.push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [run("Photos", { bold: true, size: 18 })] }));
    const half = Math.floor(CONTENT_W / 2);
    for (let i = 0; i < photos.length; i += 2) {
      const cells: any[] = [];
      for (let j = 0; j < 2; j++) {
        const p = photos[i + j];
        const kids: any[] = [];
        if (p) {
          const blob = await loadImageBlob(p.filename);
          if (blob) {
            let buf: ArrayBuffer;
            try { buf = await compressImageForExport(blob, 800, 0.7); } catch { buf = await blobToArrayBuffer(blob); }
            kids.push(new Paragraph({ spacing: { after: 0 }, children: [new ImageRun({ data: buf, transformation: { width: 320, height: 240 }, type: "jpg" })] }));
          } else kids.push(para("[Photo unavailable]"));
          const label = p.slot === "complete" ? photoSlotLabel(p.slot) : (typeof p.wipNumber === "number" ? `WIP ${p.wipNumber}` : photoSlotLabel(p.slot));
          const caption = p.caption ? ` \u2014 ${safeText(p.caption)}` : "";
          const date = added && added.has(p.id) && p.createdAt ? ` (added ${formatPhotoDate(p.createdAt)})` : "";
          kids.push(new Paragraph({ spacing: { before: 40, after: 120 }, children: [run(label, { bold: true, size: 14, color: "666666" }), run(caption + date, { size: 14, color: "666666" })] }));
        } else kids.push(new Paragraph({ children: [] }));
        cells.push(new TableCell({ width: { size: half, type: WidthType.DXA }, borders: noBorders, children: kids }));
      }
      out.push(new Table({ width: { size: half * 2, type: WidthType.DXA }, columnWidths: [half, half], layout: TableLayoutType.FIXED, rows: [new TableRow({ cantSplit: true, children: cells })] }));
    }
    return out;
  };

  // One record, flowing on from the previous one. A bold rule at the header/footer width
  // separates records; the heading and details are kept together, photos may run on.
  const recordBlock = async (d: any, o: { note?: boolean; rule?: boolean } = {}) => {
    const out: any[] = [];
    const asNote = !!o.note;
    const multi = Array.isArray(d.locations) && d.locations.length > 0;
    const isObs = d.recordType === "observation";
    const added = d.events ? new Set<number>(d.events.photosAddedThisInspection || []) : undefined;
    if (o.rule) out.push(new Paragraph({ spacing: { before: 120, after: 120 }, border: { top: { style: BorderStyle.SINGLE, size: 12, color: AUTO, space: 1 } }, children: [] }));
    out.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [
      run(asNote ? " NOTE " : isObs ? " OBSERVATION " : " DEFECT ", { bold: true, size: 16, color: "FFFFFF", shading: { type: ShadingType.SOLID, color: asNote ? "94A3B8" : isObs ? "3B82F6" : "D97706" } }),
      new TextRun({ text: "  " }),
      // Bookmark parts built by hand: the library's Bookmark gives every bookmark id 1.
      new BookmarkStart(anchorFor(d.id), d.id),
      run(multi ? `Multiple Entries for ${getWorkTypeLabel(safeText(d.uid))}` : safeText(d.uid), { bold: true, size: 28, color: CAPTION_BLUE }),
      new BookmarkEnd(d.id),
      ...(asNote ? [] : [new TextRun({ text: "    " }), run(d.status === "complete" ? "COMPLETE" : "OPEN", { bold: true, size: 20, color: d.status === "complete" ? "228B22" : "C89600" })]),
    ] }));
    const where = multi ? "" : formatDefectLocation(d, dims);
    if (where) out.push(new Paragraph({ keepNext: true, spacing: { after: 60 }, children: [run(where, { size: 18, color: "666666" })] }));
    const info: Array<[string, string]> = asNote
      ? [["Date Recorded", safeText(d.dateOpened) || "\u2014"], ["Observation", safeText(d.comment) || "\u2014"], ["Action Required", "No action required"]]
      : [
          ["Date Opened", safeText(d.dateOpened) || "\u2014"],
          ["Date Completed", safeText(d.dateClosed) || "\u2014"],
          ["Observation", safeText(d.comment) || "\u2014"],
          ["Action Required", safeText(d.actionRequired) || "\u2014"],
          ["Assigned To", safeText(d.assignedTo) || "\u2014"],
          ["Due Date", safeText(d.dueDate) || "\u2014"],
          ["Verification Method", safeText(d.verificationMethod) || "\u2014"],
          ["Verification Person", safeText(d.verificationPerson) || "\u2014"],
        ];
    out.push(keyValueTable(info, 25, { keepNext: true, size: 18 }));
    if (multi) {
      const all = [
        { uid: d.uid, description: d.comment, elevation: "", photos: d.photos || [] },
        ...d.locations.map((l: any) => ({ uid: l.uid || "", description: l.description || "", elevation: l.elevation || "", photos: [] as any[] })),
      ];
      for (let i = 0; i < all.length; i++) {
        const loc = all[i];
        out.push(new Paragraph({ keepNext: true, spacing: { before: 120, after: 20 }, children: [run(loc.uid || "\u2014", { bold: true, size: 22, color: CAPTION_BLUE })] }));
        if (loc.uid) out.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [run(deriveLocation(safeText(loc.uid)), { size: 16, color: "666666" })] }));
        if (i > 0 && loc.description) out.push(para(loc.description));
        if (loc.elevation) out.push(new Paragraph({ spacing: { after: 60 }, children: [run(`Elevation: ${ELEVATION_NAMES[loc.elevation] || loc.elevation}`, { size: 15, italics: true, color: "888888" })] }));
        out.push(...(await photosBlock(loc.photos || [], added)));
      }
    } else {
      out.push(...(await photosBlock(d.photos || [], added)));
    }
    return out;
  };
  const ruleLine = () => new Paragraph({ spacing: { before: 120, after: 0 }, border: { top: { style: BorderStyle.SINGLE, size: 12, color: AUTO, space: 1 } }, children: [] });
  const groupBlocks = async (groups: CategoryGroup[]) => {
    const out: any[] = [];
    for (let gi = 0; gi < groups.length; gi++) {
      const g = groups[gi];
      if (gi > 0) out.push(ruleLine()); // every change of record is marked, including a new category
      out.push(h3(g.label));
      if (g.kind === "summary") { out.push(summaryLine(g)); continue; }
      for (let i = 0; i < g.defects.length; i++) out.push(...(await recordBlock(g.defects[i], { rule: i > 0 })));
    }
    return out;
  };

  // ===================== 4. NEW THIS INSPECTION =====================
  {
    const c: any[] = [h1("4")];
    if (v2.newThisInspection.length === 0 && v2.notes.length === 0) c.push(para("No new records this inspection."));
    c.push(...(await groupBlocks(v2.newThisInspection)));
    if (v2.notes.length > 0) {
      if (v2.newThisInspection.length > 0) c.push(ruleLine());
      c.push(h3("Notes"));
      for (let i = 0; i < v2.notes.length; i++) c.push(...(await recordBlock(v2.notes[i], { note: true, rule: i > 0 })));
    }
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== 5. CLOSED THIS INSPECTION =====================
  {
    const c: any[] = [h1("5")];
    if (v2.closedThisInspection.length === 0) c.push(para("No records were closed this inspection."));
    c.push(...(await groupBlocks(v2.closedThisInspection)));
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== 6. RUNNING LIST OF OPEN ITEMS =====================
  {
    const c: any[] = [h1("6"), para("All open items from previous inspections.")];
    if (v2.runningOpen.length === 0) c.push(para("No open items from previous inspections."));
    for (const g of v2.runningOpen) {
      if (g.kind === "summary") { c.push(summaryLine(g)); continue; }
      c.push(h3(g.label));
      c.push(registerTable(g.defects));
    }
    sections.push({ properties: portrait, ...portraitHF, children: c });
  }

  // ===================== 7. ELEVATIONS (A4 landscape, one per page) =====================
  if (elevations.length === 0) {
    sections.push({ properties: portrait, ...portraitHF, children: [h1("7"), para("No elevation drawings have pins for records in this report.")] });
  }
  const twipsToPx = (t: number) => (t / 1440) * 96;
  elevations.forEach((el, i) => {
    const c: any[] = [];
    if (i === 0) c.push(h1("7"));
    c.push(h2(`7.elevation.${el.id}`));
    const maxW = twipsToPx(CONTENT_W_LANDSCAPE);
    const maxH = twipsToPx(PAGE_W - V_MARGIN * 2) - (i === 0 ? 150 : 100) - 60;
    const k = Math.min(maxW / el.width, maxH / el.height);
    c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new ImageRun({ data: dataUrlToBuffer(el.dataUrl), transformation: { width: Math.round(el.width * k), height: Math.round(el.height * k) }, type: "jpg" })] }));
    c.push(new Paragraph({ spacing: { after: 0 }, children: [
      ...([["Open", "open"], ["Closed", "complete"], ["Note", "note"]] as const).flatMap(([label, key]) => [run("\u25CF ", { size: 18, color: PIN_COLOURS[key].slice(1) }), run(`${label}     `, { size: 16 })]),
      run(`${el.pins.length} pin${el.pins.length === 1 ? "" : "s"}`, { size: 16 }),
    ] }));
    sections.push({ properties: landscape, ...landscapeHF, children: c });
  });

  // ===================== APPENDIX A (last; no section number) =====================
  sections.push({ properties: portrait, ...portraitHF, children: [
    h1("appendixA"),
    para(tree.appendixMode === "reference"
      ? "Refer to the contractor report for full technical method statements."
      : "Technical guidance, Technical Data Sheets (TDS) and method statements applicable to the works are included with this report."),
  ] });

  const wordDoc = new Document({
    creator: "Angel Fa\u00e7ade Consulting",
    title: `${safeText(data.project.name)} \u2014 Site Visit Report`,
    features: { updateFields: true },
    styles: {
      default: {
        document: { run: { font: FONT, size: 20, color: AUTO }, paragraph: { spacing: { before: 0, after: 200, line: 240 } } },
        heading1: { run: { font: FONT, size: 40, bold: true, color: DARK_TEXT }, paragraph: { spacing: { before: 0, after: 200 }, keepNext: true } },
        // Level 2: Left 0 cm, Hanging 1.27 cm, 2 pt before, 6 pt after, no page break before.
        heading2: { run: { font: FONT, size: 32, bold: true, color: DARK_TEXT }, paragraph: { indent: { left: 720, hanging: 720 }, tabStops: [{ type: TabStopType.LEFT, position: 720 }], spacing: { before: 40, after: 120 }, keepNext: true, pageBreakBefore: false } },
        heading3: { run: { font: FONT, size: 24, bold: true, color: CAPTION_BLUE }, paragraph: { spacing: { before: 160, after: 60 }, keepNext: true } },
      },
    },
    sections,
  });
  return Packer.toBlob(wordDoc);
}
