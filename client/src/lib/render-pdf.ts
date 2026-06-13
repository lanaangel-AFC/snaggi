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
  getLocationDimensions, formatReportDate, formatPhotoDate,
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
  const ACCENT_BLUE = [69, 176, 225] as const;
  const DARK_TEXT = [58, 58, 58] as const;
  const CAPTION_BLUE = [14, 40, 65] as const;

  const afcRef = data.project.afcReference || "AFC-24XXX";

  // ===================== COVER PAGE =====================
  doc.setFillColor(...TEAL);
  doc.rect(0, 0, 25, pageHeight, "F");
  if (logo) doc.addImage(logo.dataUrl, "PNG", pageWidth - 60, 15, 45, 14.4);

  const titleX = 35;
  let ty = pageHeight * 0.35;
  doc.setFontSize(42);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_BLUE);
  const titleLines = doc.splitTextToSize(String(data.project.name || "").toUpperCase(), contentWidth - 25);
  doc.text(titleLines, titleX, ty);
  ty += titleLines.length * 16 + 4;
  doc.setFontSize(26);
  doc.setFont("helvetica", "normal");
  doc.text("SITE VISIT REPORT", titleX, ty); ty += 14;
  doc.setFontSize(14);
  const rev = data.report.revision || "01";
  doc.text(`Revision ${rev}`, titleX, ty); ty += 20;
  doc.setFontSize(12);
  doc.setTextColor(100);
  doc.text(formatReportDate(new Date().toISOString()), titleX, ty); ty += 16;
  doc.setFontSize(11);
  doc.setTextColor(...DARK_TEXT);
  doc.text("Angel Façade Consulting", titleX, ty); ty += 6;
  doc.text(`${data.project.inspector} | 0407 759 590`, titleX, ty); ty += 6;
  doc.text(String(afcRef), titleX, ty);

  const addHeader = () => {
    doc.setFontSize(8);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text(String(data.project.name || "").toUpperCase(), margin, 10);
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
    doc.text(String(afcRef), pageWidth - margin, pageHeight - 8, { align: "right" });
  };

  const newSection = () => { doc.addPage(); addHeader(); addFooter(); return 20; };

  // ===================== SECTION 1 — INTRODUCTION =====================
  let y = newSection();
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_TEXT);
  doc.text("1. Introduction", margin, y); y += 10;
  doc.setFontSize(14);
  doc.text("1.1 General", margin, y); y += 8;
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(60);
  const generalText = `Angel Façade Consulting (AFC) was engaged by ${data.project.client} to carry out a site visit inspection of the facade at ${data.project.address}.`;
  const genLines = doc.splitTextToSize(generalText, contentWidth);
  doc.text(genLines, margin, y); y += genLines.length * 4.5 + 6;
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...DARK_TEXT);
  doc.text("1.2 Inspection", margin, y); y += 8;

  const inspRows: string[][] = [];
  if (data.report.inspectionDate) inspRows.push(["Date", formatReportDate(data.report.inspectionDate)]);
  else inspRows.push(["Date", formatReportDate(new Date().toISOString())]);
  if (data.report.inspectionNumber) inspRows.push(["Inspection Number", String(data.report.inspectionNumber)]);
  inspRows.push(["Inspector", String(data.project.inspector || "")]);
  inspRows.push(["Locations covered", String(data.report.locationsCovered || data.project.address || "")]);
  inspRows.push(["Client", String(data.project.client || "")]);
  try {
    const attendees = JSON.parse(data.report.attendees || "[]");
    if (attendees.length > 0) inspRows.push(["Attendees", attendees.map((a: any) => `${a.name} (${a.company})`).join(", ")]);
  } catch {}
  autoTable(doc, {
    startY: y, body: inspRows, margin: { left: margin, right: margin },
    styles: { fontSize: 9, cellPadding: 3 },
    columnStyles: { 0: { fontStyle: "bold", cellWidth: 35 }, 1: { cellWidth: contentWidth - 35 } },
    theme: "plain",
  });
  y = (doc as any).lastAutoTable.finalY + 10;

  // ===================== SHARED RENDER HELPERS =====================
  const summaryHead = [["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"]];
  const renderSummaryTable = (defects: any[], startY: number) => {
    const body: string[][] = [];
    for (const d of defects) {
      body.push([d.uid, d.recordType === "observation" ? "Obs" : "Defect", formatDefectLocation(d, pdfDims), getWorkTypeLabel(d.uid), d.assignedTo || "\u2014", d.dueDate || "\u2014", d.status === "complete" ? "Complete" : "Open"]);
      if (d.locations && d.locations.length > 0) for (const loc of d.locations) {
        const locUid = loc.uid || "";
        body.push([locUid, d.recordType === "observation" ? "Obs" : "Defect", locUid ? deriveLocation(locUid) : "", getWorkTypeLabel(d.uid), d.assignedTo || "\u2014", d.dueDate || "\u2014", d.status === "complete" ? "Complete" : "Open"]);
      }
    }
    if (body.length === 0) return startY;
    autoTable(doc, {
      startY, head: summaryHead, body, margin: { left: margin, right: margin },
      styles: { fontSize: 6.5, cellPadding: 1.5, overflow: "linebreak", valign: "top" },
      headStyles: { fillColor: [255, 255, 255], textColor: [...CAPTION_BLUE], fontStyle: "bold", lineWidth: { bottom: 0.5 }, lineColor: [...DARK_TEXT] },
      columnStyles: { 0: { cellWidth: 20 }, 1: { cellWidth: 12 }, 2: { cellWidth: 24 }, 3: { cellWidth: 34 }, 4: { cellWidth: 30 }, 5: { cellWidth: 26 }, 6: { cellWidth: 20 } },
      didDrawPage: () => { addHeader(); addFooter(); },
      rowPageBreak: "auto",
    });
    return (doc as any).lastAutoTable.finalY + 6;
  };

  const renderPhotos = async (photoList: any[], startY: number, photosAddedIds?: Set<number>): Promise<number> => {
    let yy = startY;
    const slotOrder = ["wip1", "wip2", "wip3", "wip4", "wip5", "complete"];
    const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", wip4: "WIP 4", wip5: "WIP 5", complete: "Complete" };
    if (photoList.length === 0) return yy;
    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...DARK_TEXT);
    doc.text("Photos", margin, yy); yy += 4;
    // Tree already trimmed; render slot photos then any extras (e.g. most-recent OLD).
    const inSlot = slotOrder.map((s: string) => photoList.find((p: any) => p.slot === s)).filter(Boolean);
    const inSlotIds = new Set(inSlot.map((p: any) => p.id));
    const extras = photoList.filter((p: any) => !inSlotIds.has(p.id));
    const sortedPhotos = [...inSlot, ...extras];
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
      const photoLabel = slotLabels[photo.slot] || photo.slot;
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
    y = renderSummaryTable(g.defects, y);
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
