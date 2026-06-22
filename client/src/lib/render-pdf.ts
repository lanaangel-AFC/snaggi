// ============================================================================
// render-pdf.ts — PDF renderer (jsPDF + jspdf-autotable). Consumes a fully-
// resolved ReportTree and walks it in master section order. It does NOT filter,
// group, or trim — the tree is the contract (see report-tree.ts). Both profiles
// share this one renderer; profile-aware branches are presentational only
// (appendix mode + client-only Progress Summary), both flagged on the tree.
//
// Returns a Blob (the caller downloads it) so this mirrors renderDocx().
// ============================================================================

import type { ReportTree, CategoryGroup } from "./report-tree";
import {
  loadImageBlob, blobToDataUrl, compressImageForPdfExport, loadAfcLogo,
  getWorkTypeLabel, ELEVATION_NAMES, deriveLocation, formatDefectLocation,
  getLocationDimensions, formatReportDate, formatPhotoDate, isDefectOverdue,
  comparePhotoSlots, photoSlotLabel,
  truncateWordBoundary, resolveActionSummary,
  parseProjectSnapshot, resolveProjectField,
} from "./render-helpers";

export async function renderPdf(tree: ReportTree, _opts: { profile: "contractor" | "client" }): Promise<Blob> {
  const data = { project: tree.project, report: tree.report };
  const pdfDims = getLocationDimensions(tree.project?.locationDimensions);

  const { default: jsPDF } = await import("jspdf");
  const autoTable = (await import("jspdf-autotable")).default;

  const doc = new jsPDF("p", "mm", "a4");
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12.7;
  const contentWidth = pageWidth - margin * 2;

  const logo = await loadAfcLogo();

  const TEAL = [0, 235, 230] as const;
  const DARK_BLUE = [10, 29, 48] as const;
  const ACCENT_BLUE = [0, 205, 200] as const;
  const DARK_TEXT = [58, 58, 58] as const;
  const CAPTION_BLUE = [14, 40, 65] as const;
  const OVERDUE_RED = [192, 0, 0] as const; // matches DOCX OVERDUE_RED (#C00000)

  // §2.3 — prefer frozen snapshot over live project row so historical reports
  // keep their original wording even after the project is later edited.
  const snap = parseProjectSnapshot((data.report as any).projectSnapshot);
  const projAddress     = resolveProjectField(snap, data.project, "address");
  const projName        = resolveProjectField(snap, data.project, "name");
  const projReportTitle = resolveProjectField(snap, data.project, "reportTitle");
  const projInspector   = resolveProjectField(snap, data.project, "inspector");
  const afcRef          = resolveProjectField(snap, data.project, "afcReference") || "AFC-24XXX";

  // §1.7 inspection number, zero-padded to 2 digits for the title-page heading.
  const inspNumRawPdf = String(data.report.inspectionNumber || "").trim();
  const inspNumPaddedPdf = inspNumRawPdf
    ? (/^\d+$/.test(inspNumRawPdf) ? inspNumRawPdf.padStart(2, "0") : inspNumRawPdf)
    : "";
  const siteVisitHeadingPdf = inspNumPaddedPdf ? `SITE VISIT REPORT ${inspNumPaddedPdf}` : "SITE VISIT REPORT";

  // ===================== COVER PAGE =====================
  // §Title-page spec — vertical stack (top → bottom):
  //   1) Address (large, dark blue)
  //   2) Report Title (full, no truncation)
  //   3) Site Visit Report NN
  //   4) Revision N
  //   5) Date
  //   6) Angel Façade Consulting
  //   7) {Inspector} | 0407 759 590
  //   8) AFC reference
  doc.setFillColor(...TEAL);
  doc.rect(0, 0, 25, pageHeight, "F");
  if (logo) doc.addImage(logo.dataUrl, "PNG", pageWidth - 60, 15, 45, 14.4);

  const titleX = 35;
  let ty = pageHeight * 0.30;

  // 1) Address (primary heading)
  doc.setFontSize(32);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_BLUE);
  const addrLines = doc.splitTextToSize(String(projAddress || "").toUpperCase(), contentWidth - 25);
  doc.text(addrLines, titleX, ty);
  ty += addrLines.length * 12 + 4;

  // 2) Report Title (full, no truncation — wraps as needed)
  if (projReportTitle) {
    doc.setFontSize(18);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...DARK_BLUE);
    const titleLines = doc.splitTextToSize(String(projReportTitle), contentWidth - 25);
    doc.text(titleLines, titleX, ty);
    ty += titleLines.length * 7 + 4;
  }

  // 3) Site Visit Report NN
  doc.setFontSize(22);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(...DARK_BLUE);
  doc.text(siteVisitHeadingPdf, titleX, ty); ty += 12;

  // 4) Revision N
  doc.setFontSize(14);
  doc.setTextColor(...DARK_TEXT);
  const rev = data.report.revision || "01";
  doc.text(`Revision ${rev}`, titleX, ty); ty += 18;

  // 5) Date
  doc.setFontSize(12);
  doc.setTextColor(100);
  doc.text(formatReportDate(new Date().toISOString()), titleX, ty); ty += 14;

  // 6) Consultant + 7) Inspector | phone + 8) AFC ref
  doc.setFontSize(11);
  doc.setTextColor(...DARK_TEXT);
  doc.text("Angel Façade Consulting", titleX, ty); ty += 6;
  doc.text(`${projInspector} | 0407 759 590`, titleX, ty); ty += 6;
  doc.text(String(afcRef), titleX, ty);

  const addHeader = () => {
    doc.setFontSize(8);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text(String(projName || "").toUpperCase(), margin, 10);
    doc.setFont("helvetica", "normal");
    doc.text("ANGEL FAÇADE CONSULTING", pageWidth - margin, 10, { align: "right" });
    doc.setDrawColor(...ACCENT_BLUE);
    doc.setLineWidth(0.5);
    doc.line(margin, 12, pageWidth - margin, 12);
  };
  const addFooter = () => {
    const pn = doc.internal.pages.length - 1;
    // Teal rule above the footer, mirroring the header rule.
    doc.setDrawColor(...ACCENT_BLUE);
    doc.setLineWidth(0.5);
    doc.line(margin, pageHeight - 12, pageWidth - margin, pageHeight - 12);
    // Footer text styled to match the header (fontSize 8, DARK_TEXT).
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...DARK_TEXT);
    doc.text(`${afcRef} | ${projAddress || ""}`, margin, pageHeight - 8);
    doc.text(`Page ${pn}`, pageWidth - margin, pageHeight - 8, { align: "right" });
  };

  const newSection = () => { doc.addPage(); addHeader(); addFooter(); return 20; };

  // ===================== SECTION 1 — INTRODUCTION =====================
  // Snapshot-resolved fields for §1.1 boilerplate substitution.
  const projClient = resolveProjectField(snap, data.project, "client");

  let y = newSection();
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_TEXT);
  doc.text("1. Introduction", margin, y); y += 10;

  // §1.1 General — boilerplate per AFC template. Wording is fixed; only CLIENT
  // and ADDRESS are substituted, sourced from the frozen snapshot when present.
  doc.setFontSize(14);
  doc.text("1.1 General", margin, y); y += 8;
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(60);
  const generalText = `Angel Façade Consulting (AFC) was engaged by ${projClient} to carry out regular inspections of the remedial works underway at ${projAddress}. Below is a summary of pertinent project information.`;
  const genLines = doc.splitTextToSize(generalText, contentWidth);
  doc.text(genLines, margin, y); y += genLines.length * 4.5 + 6;

  // §1.1 Roles table (Role / Entity / Contact Details). From snapshot when present.
  const rolesJsonPdf = resolveProjectField(snap, data.project, "roles") || "[]";
  let rolesPdf: Array<{ role?: string; entity?: string; contactDetails?: string }> = [];
  try { rolesPdf = JSON.parse(rolesJsonPdf) || []; } catch { rolesPdf = []; }
  if (Array.isArray(rolesPdf) && rolesPdf.length > 0) {
    autoTable(doc, {
      startY: y,
      head: [["Role", "Entity", "Contact Details"]],
      body: rolesPdf.map((r) => [
        String(r.role || ""),
        String(r.entity || ""),
        String(r.contactDetails || ""),
      ]),
      margin: { left: margin, right: margin },
      styles: { fontSize: 9, cellPadding: 3, valign: "top", overflow: "linebreak" },
      headStyles: { fillColor: [242, 242, 242], textColor: 50, fontStyle: "bold" },
      columnStyles: { 0: { cellWidth: contentWidth * 0.25 }, 1: { cellWidth: contentWidth * 0.30 }, 2: { cellWidth: contentWidth * 0.45 } },
      theme: "grid",
    });
    y = (doc as any).lastAutoTable.finalY + 8;
  }

  // §1.2 Scope of works — Area Ref / Location / Work Item / Access Method.
  // Sourced from the frozen snapshot when present.
  const scopeJsonPdf = resolveProjectField(snap, data.project, "scopeOfWorks") || "[]";
  let scopePdf: Array<{ areaRef?: string; location?: string; workItem?: string; accessMethod?: string }> = [];
  try { scopePdf = JSON.parse(scopeJsonPdf) || []; } catch { scopePdf = []; }
  if (Array.isArray(scopePdf) && scopePdf.length > 0) {
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text("1.2 Scope of works", margin, y); y += 8;
    autoTable(doc, {
      startY: y,
      head: [["Area Ref", "Location", "Work item", "Access method"]],
      body: scopePdf.map((s) => [
        String(s.areaRef || ""),
        String(s.location || ""),
        String(s.workItem || ""),
        String(s.accessMethod || ""),
      ]),
      margin: { left: margin, right: margin },
      styles: { fontSize: 9, cellPadding: 3, valign: "top", overflow: "linebreak" },
      headStyles: { fillColor: [242, 242, 242], textColor: 50, fontStyle: "bold" },
      columnStyles: {
        0: { cellWidth: contentWidth * 0.18 },
        1: { cellWidth: contentWidth * 0.27 },
        2: { cellWidth: contentWidth * 0.30 },
        3: { cellWidth: contentWidth * 0.25 },
      },
      theme: "grid",
    });
    y = (doc as any).lastAutoTable.finalY + 10;
  }

  // §1.3 Inspection particulars (renumbered from previous §1.2 Inspection).
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_TEXT);
  doc.text("1.3 Inspection particulars", margin, y); y += 8;

  // §1.3 Inspection particulars — ONLY inspection-specific fields. Client and
  // address are already in the §1.1 boilerplate, so they're omitted here. Each
  // attendee gets its own line so multi-person inspections stay legible.
  const inspRows: string[][] = [];
  inspRows.push(["Date", data.report.inspectionDate ? formatReportDate(data.report.inspectionDate) : formatReportDate(new Date().toISOString())]);
  if (data.report.inspectionNumber) inspRows.push(["Inspection number", String(inspNumPaddedPdf || data.report.inspectionNumber)]);
  if (data.report.revision) inspRows.push(["Revision", String(data.report.revision)]);
  inspRows.push(["Inspector", String(projInspector || "")]);
  inspRows.push(["Locations covered", String(data.report.locationsCovered || projAddress || "")]);
  try {
    const attendees = JSON.parse(data.report.attendees || "[]");
    if (Array.isArray(attendees) && attendees.length > 0) {
      const lines = attendees
        .map((a: any) => {
          const name = String(a.name || "");
          const company = String(a.company || "");
          return company ? `${name} (${company})` : name;
        })
        .filter((s: string) => s.trim() !== "");
      if (lines.length > 0) inspRows.push(["Attendees", lines.join("\n")]);
    }
  } catch {}
  autoTable(doc, {
    startY: y, body: inspRows, margin: { left: margin, right: margin },
    // overflow:linebreak makes \n inside a cell render as a hard line break,
    // which is what we want for the multi-line Attendees row.
    styles: { fontSize: 9, cellPadding: 3, valign: "top", overflow: "linebreak" },
    columnStyles: { 0: { fontStyle: "bold", cellWidth: 35 }, 1: { cellWidth: contentWidth - 35 } },
    theme: "plain",
  });
  y = (doc as any).lastAutoTable.finalY + 10;

  // §1.4 Background information — flat Harvard reference list. Mirrors the DOCX
  // §1.4 block: each entry in backgroundDocs JSON renders as one bibliographic
  // line of the form: Originator (Year). Title. [Type], Doc no. XYZ, Rev. A.
  // Empty parts are silently omitted so partial entries still render cleanly.
  // jsPDF cannot mix bold + italic + roman in a single text() call, so each
  // entry is drawn as head (bold) + title (italic) + tail (roman) using width
  // measurements to position the segments side-by-side on the same baseline.
  const bgJsonPdf = resolveProjectField(snap, data.project, "backgroundDocs") || "[]";
  let bgDocsPdf: Array<{ type?: string; originator?: string; title?: string; docNumbers?: string; revision?: string; date?: string }> = [];
  try { bgDocsPdf = JSON.parse(bgJsonPdf) || []; } catch { bgDocsPdf = []; }
  if (Array.isArray(bgDocsPdf) && bgDocsPdf.length > 0) {
    // Heading style mirrors the §1.3 heading immediately above.
    if (y + 20 > pageHeight - 20) y = newSection();
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text("1.4 Background information", margin, y); y += 8;
    // Intro line — same wording as the DOCX renderer.
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...DARK_TEXT);
    const bgIntro = "The following documents have been reviewed and form the basis of this inspection.";
    const bgIntroLines = doc.splitTextToSize(bgIntro, contentWidth);
    bgIntroLines.forEach((ln: string) => { doc.text(ln, margin, y); y += 5; });
    y += 2;
    // 4-digit year regex helper, matches the DOCX yearOf().
    const yearOfPdf = (d?: string): string => {
      if (!d) return "";
      const m = String(d).match(/\b(19|20)\d{2}\b/);
      return m ? m[0] : String(d);
    };
    const sText = (v: any): string => (v == null ? "" : String(v));
    const hangIndent = 6; // mm — hanging indent for wrapped reference lines.
    bgDocsPdf.forEach((b) => {
      const originator = sText(b.originator).trim();
      const year = yearOfPdf(b.date).trim();
      const title = sText(b.title).trim();
      const docNo = sText(b.docNumbers).trim();
      const rev = sText(b.revision).trim();
      const type = sText(b.type).trim();
      const tailParts: string[] = [];
      if (type) tailParts.push(`[${type}]`);
      if (docNo) tailParts.push(`Doc no. ${docNo}`);
      if (rev) tailParts.push(`Rev. ${rev}`);
      const head = originator ? (year ? `${originator} (${year}). ` : `${originator}. `) : (year ? `(${year}). ` : "");
      const titleText = title ? `${title}.` : "";
      const tail = tailParts.length > 0 ? ` ${tailParts.join(", ")}.` : "";
      const fullPlain = `${head}${titleText}${tail}`;
      if (!fullPlain.trim()) return;
      // Wrap the full plain-text version for page-break + multi-line layout.
      const wrapWidth = contentWidth - hangIndent;
      const wrapped = doc.splitTextToSize(fullPlain, wrapWidth);
      const entryHeight = wrapped.length * 5 + 2;
      if (y + entryHeight > pageHeight - 20) y = newSection();
      doc.setFontSize(10);
      doc.setTextColor(...DARK_TEXT);
      if (wrapped.length === 1) {
        // Single line: draw the three styled segments side-by-side.
        let x = margin;
        if (head) {
          doc.setFont("helvetica", "bold");
          doc.text(head, x, y);
          x += doc.getTextWidth(head);
        }
        if (titleText) {
          doc.setFont("helvetica", "italic");
          doc.text(titleText, x, y);
          x += doc.getTextWidth(titleText);
        }
        if (tail) {
          doc.setFont("helvetica", "normal");
          doc.text(tail, x, y);
        }
        y += 5;
      } else {
        // Multi-line: try to keep head bold on line 1, then render the rest as
        // plain roman with hanging indent. Best-effort styling for long titles.
        const firstLine = wrapped[0];
        if (head && firstLine.startsWith(head)) {
          doc.setFont("helvetica", "bold");
          doc.text(head, margin, y);
          const headW = doc.getTextWidth(head);
          doc.setFont("helvetica", "normal");
          doc.text(firstLine.slice(head.length), margin + headW, y);
        } else {
          doc.setFont("helvetica", "normal");
          doc.text(firstLine, margin, y);
        }
        y += 5;
        doc.setFont("helvetica", "normal");
        for (let i = 1; i < wrapped.length; i++) {
          if (y + 5 > pageHeight - 20) y = newSection();
          doc.text(wrapped[i], margin + hangIndent, y);
          y += 5;
        }
      }
      y += 2;
    });
    y += 4;
    // Restore default font state so downstream renderers aren't affected.
    doc.setFont("helvetica", "normal");
  }

  // §1.5 Defect/observation UID nomenclature — fixed boilerplate from the AFC
  // SVR template. Renders unconditionally (every report includes this section).
  // §1.5.1 "Unique identifier" follows as a bolded subsection heading.
  if (y + 24 > pageHeight - 20) y = newSection();
  doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.setTextColor(...DARK_TEXT);
  doc.text("1.5 Defect/observation UID nomenclature", margin, y); y += 8;
  doc.setFontSize(12); doc.setFont("helvetica", "bold");
  doc.text("1.5.1 Unique identifier", margin, y); y += 6;
  doc.setFontSize(10); doc.setFont("helvetica", "normal");
  const uidParas: string[] = [
    "Throughout the document, observations and defects are referred to by their unique identifier (UID).",
    "The UID comprises the following components: Area Ref \u2013 Work item \u2013 Sequential identifier.",
  ];
  uidParas.forEach((para) => {
    const lines = doc.splitTextToSize(para, contentWidth);
    const h = lines.length * 5 + 2;
    if (y + h > pageHeight - 20) y = newSection();
    lines.forEach((ln: string) => { doc.text(ln, margin, y); y += 5; });
    y += 2;
  });
  y += 4;

  // §1.6 Limitations — fixed 9-bullet list, verbatim from the template. Each
  // bullet is rendered with a leading "\u2022" glyph and a hanging indent so
  // wrapped lines align with the bullet text, matching the DOCX layout.
  if (y + 20 > pageHeight - 20) y = newSection();
  doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.setTextColor(...DARK_TEXT);
  doc.text("1.6 Limitations", margin, y); y += 8;
  doc.setFontSize(10); doc.setFont("helvetica", "normal");
  const limitationsBulletsPdf: string[] = [
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
  const bulletIndent = 5; // mm — hanging indent for wrapped bullet lines.
  limitationsBulletsPdf.forEach((bullet) => {
    const wrapped = doc.splitTextToSize(bullet, contentWidth - bulletIndent);
    const h = wrapped.length * 5 + 2;
    if (y + h > pageHeight - 20) y = newSection();
    // Bullet glyph on the first line; subsequent wrapped lines indent to align.
    doc.text("\u2022", margin, y);
    wrapped.forEach((ln: string) => {
      doc.text(ln, margin + bulletIndent, y);
      y += 5;
    });
    y += 1;
  });
  y += 4;

  // ===================== SHARED RENDER HELPERS =====================
  const summaryHead = [["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"]];
  // Action List adds a "Category" column immediately after "Responsible".
  const summaryHeadCat = [["ID", "Type", "Location", "Work Type", "Responsible", "Category", "By Date", "Status"]];
  // Action List (§2.2): 8-column AI layout. Sub-location rows leave the
  // Observation/Action cells blank since the parent row carries the cached
  // AI summary for that defect.
  const summaryHeadAct = [["UID", "Location", "Work item", "Observation", "Action", "Responsible", "Due Date", "Status"]];
  type SummaryMode = { showCategory?: boolean; showAction?: boolean };
  // showAction is true only for the Action List; showCategory is kept for
  // older Carried-forward callers. Default {} → original 7-column layout.
  const renderSummaryTable = (defects: any[], startY: number, mode: SummaryMode = {}) => {
    const showCategory = !!mode.showCategory;
    const showAction = !!mode.showAction;
    const body: string[][] = [];
    // Per-body-row overdue flag (parallel to `body`). A defect's sub-location
    // rows inherit the parent's overdue state since they share the same status.
    const overdueByRow: boolean[] = [];
    // Status column index: 7 for both action and category modes, 6 otherwise.
    const statusColIdx = showAction ? 7 : showCategory ? 7 : 6;
    for (const d of defects) {
      const cat = d.categoryLabel || "(uncategorised)";
      // Overdue := dueDate < today AND not closed/archived (shared helper).
      // Closed items are never flagged overdue. The base status text is left
      // intact; the red " - Overdue" suffix is drawn in didDrawCell below.
      const overdue = isDefectOverdue(d);
      const statusText = d.status === "complete" ? "Complete" : "Open";
      if (showAction) {
        // Parent row: full Observation (truncated) + AI/fallback Action.
        body.push([
          d.uid,
          formatDefectLocation(d, pdfDims),
          getWorkTypeLabel(d.uid),
          truncateWordBoundary(d.comment, { maxWords: 12, maxChars: 80 }),
          resolveActionSummary(d),
          d.assignedTo || "\u2014",
          d.dueDate || "\u2014",
          statusText,
        ]);
        overdueByRow.push(overdue);
        if (d.locations && d.locations.length > 0) for (const loc of d.locations) {
          const locUid = loc.uid || "";
          // Sub-location rows leave Observation + Action blank to avoid duplication.
          body.push([
            locUid,
            locUid ? deriveLocation(locUid) : "",
            getWorkTypeLabel(d.uid),
            "",
            "",
            d.assignedTo || "\u2014",
            d.dueDate || "\u2014",
            statusText,
          ]);
          overdueByRow.push(overdue);
        }
        continue;
      }
      const baseRow = [d.uid, d.recordType === "observation" ? "Obs" : "Defect", formatDefectLocation(d, pdfDims), getWorkTypeLabel(d.uid), d.assignedTo || "\u2014"];
      const tail = [d.dueDate || "\u2014", statusText];
      body.push(showCategory ? [...baseRow, cat, ...tail] : [...baseRow, ...tail]);
      overdueByRow.push(overdue);
      if (d.locations && d.locations.length > 0) for (const loc of d.locations) {
        const locUid = loc.uid || "";
        const locBase = [locUid, d.recordType === "observation" ? "Obs" : "Defect", locUid ? deriveLocation(locUid) : "", getWorkTypeLabel(d.uid), d.assignedTo || "\u2014"];
        body.push(showCategory ? [...locBase, cat, ...tail] : [...locBase, ...tail]);
        overdueByRow.push(overdue);
      }
    }
    if (body.length === 0) return startY;
    // Original widths sum to 166 (20+12+24+34+30+26+20). With Category added,
    // shrink Work Type (description) to keep the table within the same span.
    // Action mode widths also sum to 166 (14+22+14+36+38+22+10+10) so the
    // table spans the same content width as the other two modes.
    const columnStyles: Record<number, { cellWidth: number }> = showAction
      ? { 0: { cellWidth: 14 }, 1: { cellWidth: 22 }, 2: { cellWidth: 14 }, 3: { cellWidth: 36 }, 4: { cellWidth: 38 }, 5: { cellWidth: 22 }, 6: { cellWidth: 10 }, 7: { cellWidth: 10 } }
      : showCategory
        ? { 0: { cellWidth: 20 }, 1: { cellWidth: 12 }, 2: { cellWidth: 24 }, 3: { cellWidth: 14 }, 4: { cellWidth: 30 }, 5: { cellWidth: 20 }, 6: { cellWidth: 26 }, 7: { cellWidth: 20 } }
        : { 0: { cellWidth: 20 }, 1: { cellWidth: 12 }, 2: { cellWidth: 24 }, 3: { cellWidth: 34 }, 4: { cellWidth: 30 }, 5: { cellWidth: 26 }, 6: { cellWidth: 20 } };
    autoTable(doc, {
      startY, head: showAction ? summaryHeadAct : showCategory ? summaryHeadCat : summaryHead, body, margin: { left: margin, right: margin },
      styles: { fontSize: 6.5, cellPadding: 1.5, overflow: "linebreak", valign: "top" },
      headStyles: { fillColor: [255, 255, 255], textColor: [...CAPTION_BLUE], fontStyle: "bold", lineWidth: { bottom: 0.5 }, lineColor: [...DARK_TEXT] },
      columnStyles,
      didDrawPage: () => { addHeader(); addFooter(); },
      // For overdue Action List rows, append a red " - Overdue" suffix after the
      // base status text in the Status column. autoTable still paints the base
      // status (e.g. "Open"); we only draw the suffix here, positioned right
      // after it using the cell's own text position so it shares autoTable's
      // baseline/alignment exactly. Closed items are never overdue (filtered in
      // overdueByRow), so they never get the suffix.
      didDrawCell: (hookData: any) => {
        if (hookData.section !== "body") return;
        if (hookData.column.index !== statusColIdx) return;
        if (!overdueByRow[hookData.row.index]) return;
        const cell = hookData.cell;
        const baseText = Array.isArray(cell.text) ? cell.text.join("") : String(cell.text ?? "");
        const pos = cell.getTextPos();
        const baseWidth = doc.getTextWidth(baseText);
        doc.setTextColor(...OVERDUE_RED);
        // jspdf-autotable v5 removed doc.autoTableText, so draw the suffix
        // directly with jsPDF. The autoTable styles use valign: "top", so use
        // baseline: "top" here to line up with the base status text.
        doc.text(" - Overdue", pos.x + baseWidth, pos.y, { baseline: "top" });
        doc.setTextColor(0);
      },
      rowPageBreak: "auto",
    });
    return (doc as any).lastAutoTable.finalY + 6;
  };

  const renderPhotos = async (photoList: any[], startY: number, photosAddedIds?: Set<number>): Promise<number> => {
    let yy = startY;
    if (photoList.length === 0) return yy;
    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text("Photos", margin, yy); yy += 4;
    // When the server supplies wipNumber (Open items: cumulative date-sorted timeline),
    // render in the order given — the timeline may legitimately contain several photos
    // per slot, which slot-bucketing would collapse. Otherwise fall back to the legacy
    // slot-order layout (one per slot, then any extras) for non-Open items.
    const hasWip = photoList.some((p: any) => typeof p.wipNumber === "number");
    let sortedPhotos: any[];
    if (hasWip) {
      sortedPhotos = photoList;
    } else {
      // One photo per slot, ordered by slot (wip1..wipN then complete), with any
      // extras (e.g. duplicate-slot rows) appended in their original order.
      const seen = new Set<string>();
      const inSlot = [...photoList]
        .sort((a: any, b: any) => comparePhotoSlots(a.slot, b.slot))
        .filter((p: any) => (seen.has(p.slot) ? false : (seen.add(p.slot), true)));
      const inSlotIds = new Set(inSlot.map((p: any) => p.id));
      const extras = photoList.filter((p: any) => !inSlotIds.has(p.id));
      sortedPhotos = [...inSlot, ...extras];
    }
    const imgW = 85; const imgH = imgW * 0.75;
    for (let i = 0; i < sortedPhotos.length; i++) {
      const photo = sortedPhotos[i];
      const col = i % 2;
      if (i > 0 && col === 0) yy += imgH + 12;
      if (col === 0 && yy + imgH + 12 > pageHeight - 20) { doc.addPage(); addHeader(); addFooter(); yy = 20; }
      const x = margin + col * (imgW + 6);
      const blob = await loadImageBlob(photo.filename);
      if (blob) {
        let dataUrl: string;
        try { dataUrl = await compressImageForPdfExport(blob, 800, 0.7); } catch { dataUrl = await blobToDataUrl(blob); }
        doc.addImage(dataUrl, "JPEG", x, yy, imgW, imgH);
      } else {
        doc.setDrawColor(180); doc.rect(x, yy, imgW, imgH); doc.setFontSize(7); doc.text("[Photo unavailable]", x + 4, yy + imgH / 2);
      }
      doc.setFontSize(7); doc.setFont("helvetica", "bold"); doc.setTextColor(100);
      // Prefer the server-computed wipNumber (1-based position in the cumulative,
      // date-sorted timeline). Falls back to the stored slot label when absent.
      const photoLabel = photo.slot === "complete"
        ? photoSlotLabel(photo.slot)
        : (typeof photo.wipNumber === "number" ? `WIP ${photo.wipNumber}` : photoSlotLabel(photo.slot));
      const photoCaption = photo.caption ? ` — ${photo.caption}` : "";
      const isNewPhoto = photosAddedIds ? photosAddedIds.has(photo.id) : false;
      const dateSuffix = isNewPhoto && photo.createdAt ? ` (added ${formatPhotoDate(photo.createdAt)})` : "";
      doc.text(photoLabel + photoCaption + dateSuffix, x, yy + imgH + 3);
      doc.setTextColor(0);
    }
    yy += (sortedPhotos.length > 0 ? imgH + 12 : 0);
    return yy;
  };

  const renderDefectPage = async (defect: any, options?: { showChangeSummary?: boolean }) => {
    const hasMultipleLocations = defect.locations && defect.locations.length > 0;
    const events = defect.events;
    const photosAddedIds = events ? new Set<number>(events.photosAddedThisInspection || []) : undefined;
    let yy = newSection();

    const typeLabel = (defect.recordType === "observation") ? "OBSERVATION" : "DEFECT";
    const typeBgColor = (defect.recordType === "observation") ? [59, 130, 246] : [217, 119, 6];
    doc.setFontSize(8); doc.setFont("helvetica", "bold");
    const badgeW = doc.getTextWidth(typeLabel) + 6;
    doc.setFillColor(typeBgColor[0], typeBgColor[1], typeBgColor[2]);
    doc.roundedRect(margin, yy - 4, badgeW, 6, 1.5, 1.5, "F");
    doc.setTextColor(255, 255, 255);
    doc.text(typeLabel, margin + 3, yy);
    doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.setTextColor(...DARK_TEXT);
    if (hasMultipleLocations) doc.text(`Multiple Entries for ${getWorkTypeLabel(defect.uid)}`, margin + badgeW + 4, yy);
    else doc.text(defect.uid, margin + badgeW + 4, yy);
    const statusLabel = defect.status === "complete" ? "COMPLETE" : "OPEN";
    const statusColor = defect.status === "complete" ? [34, 139, 34] : [200, 150, 0];
    doc.setFontSize(10); doc.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
    doc.text(statusLabel, pageWidth - margin - doc.getTextWidth(statusLabel), yy);
    doc.setTextColor(0); yy += 5;

    if (!hasMultipleLocations) {
      doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(100);
      doc.text(formatDefectLocation(defect, pdfDims), margin, yy); yy += 5;
    } else { yy += 2; }

    if (options?.showChangeSummary && events && !events.isNew) {
      const af = events.amendedFields;
      const lines: string[] = [];
      if (af.observation) lines.push("Observation amended");
      if (af.action) lines.push("Action amended");
      if (af.photos > 0) lines.push(`${af.photos} new photo${af.photos > 1 ? "s" : ""}`);
      if (af.locationsAdded > 0) lines.push("New location added");
      if (af.locationsAmended > 0) lines.push("Location amended");
      if (af.statusChange) lines.push(`Status changed: ${af.statusChange.from || "\u2014"} \u2192 ${af.statusChange.to}`);
      if (lines.length > 0) {
        doc.setFillColor(254, 243, 199);
        doc.roundedRect(margin, yy - 2, contentWidth, 5 + lines.length * 4, 1, 1, "F");
        doc.setFontSize(8); doc.setFont("helvetica", "bold"); doc.setTextColor(146, 64, 14);
        doc.text("Changes this inspection:", margin + 2, yy + 2); yy += 5;
        doc.setFont("helvetica", "normal"); doc.setFontSize(7.5);
        for (const line of lines) { doc.text(`  \u2022  ${line}`, margin + 2, yy + 1); yy += 4; }
        doc.setTextColor(0); yy += 3;
      }
    }

    const obsLabel = (events && !events.isNew && events.amendedFields.observation) ? "Observation (amended)" : "Observation";
    const actLabel = (events && !events.isNew && events.amendedFields.action) ? "Action Required (amended)" : "Action Required";
    const infoBody = [
      ["Date Opened", defect.dateOpened || "\u2014"],
      ["Date Completed", defect.dateClosed || "\u2014"],
      [obsLabel, defect.comment || "\u2014"],
      [actLabel, defect.actionRequired || "\u2014"],
      ["Assigned To", defect.assignedTo || "\u2014"],
      ["Due Date", defect.dueDate || "\u2014"],
      ["Verification Method", defect.verificationMethod || "\u2014"],
      ["Verification Person", defect.verificationPerson || "\u2014"],
    ];
    autoTable(doc, {
      startY: yy, body: infoBody, margin: { left: margin, right: margin },
      styles: { fontSize: 8, cellPadding: 2, overflow: "linebreak", valign: "top" },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: 40 }, 1: { cellWidth: contentWidth - 40 } },
      theme: "plain",
      didDrawPage: () => { addHeader(); addFooter(); },
    });
    yy = (doc as any).lastAutoTable.finalY + 6;

    if (hasMultipleLocations) {
      const allLocs = [
        { uid: defect.uid, description: defect.comment, elevation: "", photos: defect.photos || [] },
        ...defect.locations.map((l: any) => ({ uid: l.uid || "", description: l.description || "", elevation: l.elevation || "", photos: [] as any[] })),
      ];
      for (let li = 0; li < allLocs.length; li++) {
        const loc = allLocs[li];
        if (yy + 20 > pageHeight - 20) { yy = newSection(); }
        doc.setFontSize(11); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE);
        doc.text(loc.uid || "\u2014", margin, yy); yy += 5;
        if (loc.uid) { doc.setFontSize(8); doc.setFont("helvetica", "normal"); doc.setTextColor(100); doc.text(deriveLocation(loc.uid), margin, yy); yy += 4; }
        if (li > 0 && loc.description) { doc.setFontSize(8); doc.setFont("helvetica", "normal"); doc.setTextColor(60); const dl = doc.splitTextToSize(loc.description, contentWidth); doc.text(dl, margin, yy); yy += dl.length * 4 + 2; }
        if (loc.elevation) { doc.setFontSize(7.5); doc.setFont("helvetica", "italic"); doc.setTextColor(120); doc.text(`Elevation: ${ELEVATION_NAMES[loc.elevation] || loc.elevation}`, margin, yy); yy += 4; }
        if (loc.photos && loc.photos.length > 0) yy = await renderPhotos(loc.photos, yy, photosAddedIds);
        yy += 4;
      }
    } else {
      yy = await renderPhotos(defect.photos || [], yy, photosAddedIds);
    }
  };

  const sectionHeading = (title: string, subtitle: string, startY: number) => {
    let yy = startY;
    doc.setFontSize(18); doc.setFont("helvetica", "bold"); doc.setTextColor(...DARK_TEXT);
    doc.text(title, margin, yy); yy += 7;
    if (subtitle) { doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(100); doc.text(subtitle, margin, yy); yy += 6; }
    return yy;
  };

  const groupSubheading = (label: string, startY: number) => {
    let yy = startY;
    if (yy + 12 > pageHeight - 20) yy = newSection();
    doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE);
    doc.text(label, margin, yy); yy += 6;
    return yy;
  };

  const summaryNoteLine = (g: Extract<CategoryGroup, { kind: "summary" }>, startY: number) => {
    let yy = startY;
    if (yy + 8 > pageHeight - 20) yy = newSection();
    const text = g.note ? `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}. ${g.note}.` : `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}.`;
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(...DARK_TEXT);
    doc.text(text, margin, yy); yy += 6;
    return yy;
  };

  // ===================== SECTION 2 — ACTION LIST =====================
  y = newSection();
  y = sectionHeading("2. Action List — This Inspection", "Based on our observations, we recommend the following actions.", y);
  let actionHasContent = false;
  for (const g of tree.actionList.groups) {
    actionHasContent = true;
    if (g.kind === "summary") { y = summaryNoteLine(g, y); continue; }
    y = groupSubheading(g.label, y);
    y = renderSummaryTable(g.defects, y, { showAction: true });
  }
  if (!actionHasContent) { doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(120); doc.text("Nothing for this report.", margin, y); }

  // ===================== SECTION 3 — PROJECT STATUS =====================
  y = newSection();
  y = sectionHeading("3. Project Status", "", y);
  if (tree.projectStatus.empty) {
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(120); doc.text("Nothing for this report.", margin, y); y += 6;
  } else {
    for (const n of tree.projectStatus.narratives) {
      if (y + 16 > pageHeight - 20) y = newSection();
      doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE);
      doc.text(`${n.title || "Narrative"}${n.status ? `  [${n.status}]` : ""}`, margin, y); y += 5;
      if (n.body) { doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(60); const bl = doc.splitTextToSize(String(n.body), contentWidth); doc.text(bl, margin, y); y += bl.length * 4.5 + 4; }
    }
    if (tree.projectStatus.program) {
      const p = tree.projectStatus.program;
      if (y + 16 > pageHeight - 20) y = newSection();
      doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Program", margin, y); y += 6;
      const pRows: string[][] = [];
      if (p.asAtDate) pRows.push(["As at", String(p.asAtDate)]);
      if (p.varianceText) pRows.push(["Variance", String(p.varianceText)]);
      if (p.projectedCompletion) pRows.push(["Projected completion", String(p.projectedCompletion)]);
      if (p.statusNarrative) pRows.push(["Status", String(p.statusNarrative)]);
      if (pRows.length > 0) {
        autoTable(doc, { startY: y, body: pRows, margin: { left: margin, right: margin }, styles: { fontSize: 8, cellPadding: 2 }, columnStyles: { 0: { fontStyle: "bold", cellWidth: 45 }, 1: { cellWidth: contentWidth - 45 } }, theme: "plain", didDrawPage: () => { addHeader(); addFooter(); } });
        y = (doc as any).lastAutoTable.finalY + 4;
      }
      if (p.programImageFilename) {
        const blob = await loadImageBlob(p.programImageFilename);
        if (blob) { let dataUrl: string; try { dataUrl = await compressImageForPdfExport(blob, 1000, 0.7); } catch { dataUrl = await blobToDataUrl(blob); } if (y + 90 > pageHeight - 20) y = newSection(); doc.addImage(dataUrl, "JPEG", margin, y, contentWidth, contentWidth * 0.56); y += contentWidth * 0.56 + 6; }
      }
    }
    if (tree.projectStatus.stageMap) {
      const sm = tree.projectStatus.stageMap;
      if (y + 16 > pageHeight - 20) y = newSection();
      doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Stage Map", margin, y); y += 6;
      if (sm.planImageFilename) {
        const blob = await loadImageBlob(sm.planImageFilename);
        if (blob) { let dataUrl: string; try { dataUrl = await compressImageForPdfExport(blob, 1000, 0.7); } catch { dataUrl = await blobToDataUrl(blob); } if (y + 90 > pageHeight - 20) y = newSection(); doc.addImage(dataUrl, "JPEG", margin, y, contentWidth, contentWidth * 0.56); y += contentWidth * 0.56 + 6; }
      }
      let stages: any[] = [];
      try { stages = Array.isArray(sm.stages) ? sm.stages : JSON.parse(sm.stages || "[]"); } catch { stages = []; }
      if (stages.length > 0) {
        autoTable(doc, { startY: y, head: [["Stage", "Status"]], body: stages.map((s: any) => [String(s.stageName || ""), String(s.status || "")]), margin: { left: margin, right: margin }, styles: { fontSize: 8, cellPadding: 2 }, theme: "grid", didDrawPage: () => { addHeader(); addFooter(); } });
        y = (doc as any).lastAutoTable.finalY + 4;
      }
    }
  }

  // ===================== CLIENT-ONLY — PROGRESS SUMMARY =====================
  if (tree.progressSummary) {
    const ps = tree.progressSummary;
    if (y + 30 > pageHeight - 20) y = newSection();
    doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.setTextColor(...DARK_TEXT); doc.text("Progress Summary", margin, y); y += 6;
    autoTable(doc, {
      startY: y, body: [["Open", String(ps.open)], ["Closed this period", String(ps.closedThisPeriod)], ["Overdue", String(ps.overdue)], ["Total", String(ps.total)]],
      margin: { left: margin, right: margin }, styles: { fontSize: 9, cellPadding: 2.5 },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: 60 }, 1: { cellWidth: 30 } }, theme: "grid",
      didDrawPage: () => { addHeader(); addFooter(); },
    });
    y = (doc as any).lastAutoTable.finalY + 6;
  }

  // ===================== SECTION 4 — THIS INSPECTION =====================
  const renderBucket = async (title: string, subtitle: string, groups: CategoryGroup[], showChangeSummary: boolean) => {
    if (groups.length === 0) return;
    let yy = newSection();
    yy = sectionHeading(title, subtitle, yy);
    for (const g of groups) {
      yy = groupSubheading(g.label, yy);
      if (g.kind === "summary") { yy = summaryNoteLine(g, yy); continue; }
      for (const d of g.defects) { await renderDefectPage(d, { showChangeSummary }); }
    }
  };
  if (tree.thisInspection.empty) {
    y = newSection();
    y = sectionHeading("4. This Inspection", "", y);
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(120); doc.text("Nothing for this report.", margin, y);
  } else {
    await renderBucket("NEW THIS INSPECTION", "Items added during this inspection.", tree.thisInspection.new, false);
    await renderBucket("AMENDED THIS INSPECTION", "Existing items updated during this inspection.", tree.thisInspection.amended, true);
    await renderBucket("COMPLETED THIS INSPECTION", "Items marked complete during this inspection.", tree.thisInspection.completed, true);
  }

  // ===================== SECTION 5 — CARRIED-FORWARD REGISTER =====================
  y = newSection();
  y = sectionHeading("5. Carried-forward Register", "Open items not covered in this inspection's locations.", y);
  if (tree.carriedForward.empty) {
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(120); doc.text("Nothing for this report.", margin, y);
  } else {
    for (const g of tree.carriedForward.groups) {
      if (g.kind === "summary") { y = summaryNoteLine(g, y); continue; }
      y = groupSubheading(g.label, y);
      y = renderSummaryTable(g.defects, y);
    }
  }

  // ===================== SECTION 6 — APPENDICES =====================
  y = newSection();
  y = sectionHeading("6. Appendices", "", y);
  if (tree.appendixMode === "reference") {
    doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Technical Method Statements", margin, y); y += 6;
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(60);
    doc.text(doc.splitTextToSize("Refer to the contractor report for full technical method statements.", contentWidth), margin, y); y += 10;
    doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Coverage Drawings", margin, y); y += 6;
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(60);
    doc.text(doc.splitTextToSize(String(data.report.locationsCovered || "Refer to inspection locations covered."), contentWidth), margin, y);
  } else {
    doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Technical Guidance & Method Statements", margin, y); y += 6;
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(60);
    doc.text(doc.splitTextToSize("Technical guidance, Technical Data Sheets (TDS) and method statements applicable to the works are included with this report.", contentWidth), margin, y); y += 12;
    doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.setTextColor(...CAPTION_BLUE); doc.text("Coverage Drawings", margin, y); y += 6;
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(60);
    doc.text(doc.splitTextToSize(String(data.report.locationsCovered || "Refer to inspection locations covered."), contentWidth), margin, y);
  }

  return doc.output("blob");
}
