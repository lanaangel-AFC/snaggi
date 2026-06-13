// ============================================================================
// render-docx.ts — DOCX renderer. Consumes a fully-resolved ReportTree and walks
// it in master section order. It does NOT filter, group, or trim — the tree is
// the contract (see report-tree.ts). Both profiles share this one renderer; the
// only profile-aware branches here are presentational (appendix mode + the
// client-only Progress Summary block, both flagged on the tree).
//
// All text content is wrapped in safeText() per the DOCX rendering constraint.
// ============================================================================

import type { ReportTree, CategoryGroup } from "./report-tree";
import {
  safeText, loadImageBlob, blobToArrayBuffer, compressImageForExport, loadAfcLogo,
  getWorkTypeLabel, ELEVATION_NAMES, deriveLocation, formatDefectLocation,
  getLocationDimensions, formatReportDate, formatPhotoDate,
} from "./render-helpers";

export async function renderDocx(tree: ReportTree, _opts: { profile: "contractor" | "client" }): Promise<Blob> {
  const data = { project: tree.project, report: tree.report };
  const wordDims = getLocationDimensions(tree.project?.locationDimensions);

  const docxLib = await import("docx");
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    ImageRun, PageBreak, HeadingLevel, AlignmentType,
    WidthType, BorderStyle, ShadingType, TableLayoutType,
    Header, Footer, PageNumber, TabStopType, TabStopPosition,
  } = docxLib as any;

  const logo = await loadAfcLogo();

  const slotOrder = ["wip1", "wip2", "wip3", "complete"];
  const slotLabels: Record<string, string> = { wip1: "WIP 1", wip2: "WIP 2", wip3: "WIP 3", complete: "Complete" };

  const DARK_BLUE = "0A1D30";
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
              children: [new TextRun({ text: safeText(data.project.name).toUpperCase(), size: 16, font: "Aptos", bold: true, color: DARK_TEXT })],
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
          new TextRun({ text: safeText(afcRef), size: 14, font: "Aptos", color: "999999" }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: "Page ", size: 14, font: "Aptos", color: "999999" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 14, font: "Aptos", color: "999999" }),
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      }),
    ],
  });

  // ===================== COVER PAGE =====================
  const coverChildren: any[] = [];
  if (logo) {
    let logoBuffer: ArrayBuffer;
    try { logoBuffer = await compressImageForExport(new Blob([logo.buffer]), 360, 0.8); }
    catch { logoBuffer = logo.buffer; }
    coverChildren.push(new Paragraph({
      children: [new ImageRun({ data: logoBuffer, transformation: { width: 180, height: 58 }, type: "jpg" })],
      alignment: AlignmentType.RIGHT,
      spacing: { after: 600 },
    }));
  } else {
    coverChildren.push(new Paragraph({ spacing: { after: 600 } }));
  }
  coverChildren.push(new Paragraph({ spacing: { before: 1200 } }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: safeText(data.project.name).toUpperCase(), bold: true, size: 72, font: "Aptos", color: DARK_BLUE })],
    spacing: { after: 100 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: "SITE VISIT REPORT", size: 52, font: "Aptos", color: DARK_BLUE, smallCaps: true })],
    spacing: { after: 100 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: `Revision ${safeText(rev)}`, size: 28, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 600 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: safeText(reportDate), size: 24, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 400 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: "Angel Façade Consulting", size: 22, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 40 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: `${safeText(data.project.inspector)} | 0407 759 590`, size: 22, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 40 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: safeText(afcRef), size: 22, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 200 },
  }));

  // ===================== SECTION 1 — INTRODUCTION =====================
  const introChildren: any[] = [];
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "Introduction", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_1, spacing: { after: 200 },
  }));
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "General", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_2, spacing: { after: 100 },
  }));
  introChildren.push(new Paragraph({
    children: [new TextRun({
      text: `Angel Façade Consulting (AFC) was engaged by ${safeText(data.project.client)} to carry out a site visit inspection of the facade at ${safeText(data.project.address)}.`,
      size: 20, font: "Aptos",
    })],
    spacing: { after: 200 },
  }));
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "Inspection", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_2, spacing: { after: 100 },
  }));

  const inspectionData: string[][] = [];
  if (data.report.inspectionDate) inspectionData.push(["Date", formatReportDate(data.report.inspectionDate)]);
  else inspectionData.push(["Date", reportDate]);
  if (data.report.inspectionNumber) inspectionData.push(["Inspection Number", safeText(data.report.inspectionNumber)]);
  inspectionData.push(["Inspector", safeText(data.project.inspector)]);
  inspectionData.push(["Locations covered", safeText(data.report.locationsCovered) || safeText(data.project.address)]);
  inspectionData.push(["Client", safeText(data.project.client)]);
  try {
    const attendees = JSON.parse(data.report.attendees || "[]");
    if (attendees.length > 0) {
      attendees.forEach((a: any) => { inspectionData.push([safeText(a.company) || safeText(a.name), safeText(a.name)]); });
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
            children: [new Paragraph({ children: [new TextRun({ text: safeText(label), bold: true, size: 20, font: "Aptos", color: DARK_TEXT })], spacing: { before: 60, after: 60 } })],
            borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
          }),
          new TableCell({
            width: { size: 75, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: safeText(value), size: 20, font: "Aptos" })], spacing: { before: 60, after: 60 } })],
            borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
          }),
        ],
      })
    ),
  }));
  introChildren.push(new Paragraph({ spacing: { before: 200 } }));

  // ===================== HELPERS: photos + defect page =====================
  const buildWordPhotos = async (photoList: any[], photosAddedIds?: Set<number>): Promise<any[]> => {
    const photoElements: any[] = [];
    if (photoList.length === 0) return photoElements;
    photoElements.push(new Paragraph({
      children: [new TextRun({ text: "Photos", bold: true, size: 20, font: "Aptos" })],
      spacing: { before: 100, after: 80 },
    }));
    // The tree already trimmed which photos appear. Render them in slot order, then
    // any photos that fall outside the known slot list (e.g. the most-recent OLD
    // photo on a client export) so nothing the tree included is dropped here.
    const inSlot = slotOrder.map((s: string) => photoList.find((p: any) => p.slot === s)).filter(Boolean);
    const inSlotIds = new Set(inSlot.map((p: any) => p.id));
    const extras = photoList.filter((p: any) => !inSlotIds.has(p.id));
    const sortedPhotos = [...inSlot, ...extras];

    for (let i = 0; i < sortedPhotos.length; i += 2) {
      const photoCells: any[] = [];
      for (let j = 0; j < 2; j++) {
        const photo = sortedPhotos[i + j];
        if (photo) {
          const cellChildren: any[] = [];
          const blob = await loadImageBlob(photo.filename);
          if (blob) {
            let buffer: ArrayBuffer;
            try { buffer = await compressImageForExport(blob, 800, 0.7); }
            catch { buffer = await blobToArrayBuffer(blob); }
            cellChildren.push(new Paragraph({
              children: [new ImageRun({ data: buffer, transformation: { width: 241, height: 181 }, type: "jpg" })],
              alignment: AlignmentType.CENTER,
            }));
          }
          const isNewPhoto = photosAddedIds ? photosAddedIds.has(photo.id) : false;
          const dateSuffix = isNewPhoto && photo.createdAt ? ` (added ${formatPhotoDate(photo.createdAt)})` : "";
          const captionText = photo.caption ? ` \u2014 ${safeText(photo.caption)}` : "";
          cellChildren.push(new Paragraph({
            children: [
              new TextRun({ text: safeText(slotLabels[photo.slot] || photo.slot), bold: true, size: 16, font: "Aptos", color: photo.slot === "complete" ? "228B22" : "666666" }),
              ...(captionText ? [new TextRun({ text: safeText(captionText), size: 14, font: "Aptos", color: "888888" })] : []),
              ...(dateSuffix ? [new TextRun({ text: safeText(dateSuffix), size: 13, font: "Aptos", color: "2563EB", italics: true })] : []),
            ],
            alignment: AlignmentType.CENTER, spacing: { before: 40 },
          }));
          photoCells.push(new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: cellChildren, borders: { top: noBorder, left: noBorder, right: noBorder, bottom: noBorder } }));
        } else {
          photoCells.push(new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph("")], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: noBorder } }));
        }
      }
      photoElements.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED, rows: [new TableRow({ children: photoCells })] }));
      photoElements.push(new Paragraph({ spacing: { before: 60 } }));
    }
    return photoElements;
  };

  const buildWordDefectPage = async (defect: any, options?: { showChangeSummary?: boolean }): Promise<any[]> => {
    const elements: any[] = [];
    const hasMultipleLocations = defect.locations && defect.locations.length > 0;
    const events = defect.events;
    const photosAddedIds = events ? new Set<number>(events.photosAddedThisInspection || []) : undefined;
    const isObs = defect.recordType === "observation";
    const headingText = hasMultipleLocations
      ? `Multiple Entries for ${getWorkTypeLabel(safeText(defect.uid)) || "Entry"}`
      : safeText(defect.uid);

    elements.push(new Paragraph({
      children: [
        new TextRun({ text: isObs ? " OBSERVATION " : " DEFECT ", bold: true, size: 16, font: "Aptos", color: "FFFFFF", shading: { type: ShadingType.SOLID, color: isObs ? "3B82F6" : "D97706" } }),
        new TextRun({ text: "  " }),
        new TextRun({ text: headingText, bold: true, size: 28, font: "Aptos", color: CAPTION_BLUE }),
        new TextRun({ text: "    " }),
        new TextRun({ text: defect.status === "complete" ? "COMPLETE" : "OPEN", bold: true, size: 20, color: defect.status === "complete" ? "228B22" : "C89600" }),
      ],
      spacing: { after: 40 },
    }));

    if (!hasMultipleLocations) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: formatDefectLocation(defect, wordDims), size: 18, font: "Aptos", color: "666666" })],
        spacing: { after: options?.showChangeSummary ? 100 : 200 },
      }));
    } else {
      elements.push(new Paragraph({ spacing: { after: 100 } }));
    }

    if (options?.showChangeSummary && events && !events.isNew) {
      const af = events.amendedFields;
      const lines: string[] = [];
      if (af.observation) lines.push("Observation amended");
      if (af.action) lines.push("Action amended");
      if (af.photos > 0) lines.push(`${af.photos} new photo${af.photos > 1 ? "s" : ""}`);
      if (af.locationsAdded > 0) lines.push(`New location added`);
      if (af.locationsAmended > 0) lines.push(`Location amended`);
      if (af.statusChange) lines.push(`Status changed: ${af.statusChange.from || "\u2014"} \u2192 ${af.statusChange.to}`);
      if (lines.length > 0) {
        elements.push(new Paragraph({
          children: [new TextRun({ text: "Changes this inspection:", bold: true, size: 17, font: "Aptos", color: "92400E" })],
          spacing: { before: 40, after: 30 }, shading: { type: ShadingType.SOLID, color: "FEF3C7" },
        }));
        for (const line of lines) {
          elements.push(new Paragraph({
            children: [new TextRun({ text: `  \u2022  ${line}`, size: 16, font: "Aptos", color: "92400E" })],
            spacing: { before: 10, after: 10 }, shading: { type: ShadingType.SOLID, color: "FEF3C7" },
          }));
        }
        elements.push(new Paragraph({ spacing: { before: 80 } }));
      }
    }

    const obsLabel = (events && !events.isNew && events.amendedFields.observation) ? "Observation (amended this inspection)" : "Observation";
    const actLabel = (events && !events.isNew && events.amendedFields.action) ? "Action Required (amended this inspection)" : "Action Required";
    const infoRows = [
      ["Date Opened", safeText(defect.dateOpened)],
      ["Date Completed", safeText(defect.dateClosed) || "\u2014"],
      [obsLabel, safeText(defect.comment)],
      [actLabel, safeText(defect.actionRequired)],
      ["Assigned To", safeText(defect.assignedTo)],
      ["Due Date", safeText(defect.dueDate)],
      ["Verification Method", safeText(defect.verificationMethod)],
      ["Verification Person", safeText(defect.verificationPerson)],
    ];
    const bottomBorder = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
    elements.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: infoRows.map(([label, value]) =>
        new TableRow({
          children: [
            new TableCell({ width: { size: 28, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(label), bold: true, size: 18, font: "Aptos" })], spacing: { before: 50, after: 50 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: bottomBorder } }),
            new TableCell({ width: { size: 72, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(value), size: 18, font: "Aptos" })], spacing: { before: 50, after: 50 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: bottomBorder } }),
          ],
        })
      ),
    }));
    elements.push(new Paragraph({ spacing: { before: 150 } }));

    if (hasMultipleLocations) {
      const allLocs = [
        { uid: defect.uid, description: defect.comment, elevation: "", photos: defect.photos || [] },
        ...defect.locations.map((l: any) => ({ uid: l.uid || "", description: l.description || "", elevation: l.elevation || "", photos: [] as any[] })),
      ];
      for (let li = 0; li < allLocs.length; li++) {
        const loc = allLocs[li];
        if (li > 0) {
          elements.push(new Paragraph({ children: [], spacing: { before: 100 }, border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } } }));
          elements.push(new Paragraph({ spacing: { before: 60 } }));
        }
        elements.push(new Paragraph({ children: [new TextRun({ text: loc.uid || "\u2014", bold: true, size: 22, font: "Aptos", color: CAPTION_BLUE })], spacing: { after: 30 } }));
        if (loc.uid) elements.push(new Paragraph({ children: [new TextRun({ text: deriveLocation(loc.uid), size: 16, font: "Aptos", color: "666666" })], spacing: { after: 40 } }));
        if (li > 0 && loc.description) elements.push(new Paragraph({ children: [new TextRun({ text: loc.description, size: 18, font: "Aptos" })], spacing: { after: 60 } }));
        if (loc.elevation) elements.push(new Paragraph({ children: [new TextRun({ text: `Elevation: ${ELEVATION_NAMES[loc.elevation] || loc.elevation}`, size: 15, font: "Aptos", color: "888888", italics: true })], spacing: { after: 60 } }));
        if (loc.photos && loc.photos.length > 0) { const photoEls = await buildWordPhotos(loc.photos, photosAddedIds); elements.push(...photoEls); }
      }
    } else {
      const photoEls = await buildWordPhotos(defect.photos || [], photosAddedIds);
      elements.push(...photoEls);
    }
    return elements;
  };

  // Render a summary (count-line) group node.
  const summaryGroupParagraph = (g: Extract<CategoryGroup, { kind: "summary" }>): any => {
    const text = g.note
      ? `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}. ${g.note}.`
      : `${g.label}: ${g.count} item${g.count === 1 ? "" : "s"}.`;
    return new Paragraph({
      children: [new TextRun({ text: safeText(text), size: 20, font: "Aptos", color: DARK_TEXT })],
      spacing: { before: 60, after: 60 },
    });
  };

  // ===================== SECTION 2 — ACTION LIST =====================
  const actionChildren: any[] = [];
  actionChildren.push(new Paragraph({
    children: [new TextRun({ text: "Action List — This Inspection", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_1, spacing: { after: 100 },
  }));
  actionChildren.push(new Paragraph({
    children: [new TextRun({ text: "Based on our observations, we recommend the following actions.", size: 20, font: "Aptos" })],
    spacing: { after: 200 },
  }));

  const summaryHeaderLabels = ["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"];
  const summaryHeaderWidths = [950, 600, 1300, 2300, 1850, 1300, 1400];
  const buildSummaryHeaderRow = () => new TableRow({
    tableHeader: true,
    children: summaryHeaderLabels.map((label, i) =>
      new TableCell({
        width: { size: summaryHeaderWidths[i], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: label, size: 16, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { before: 30, after: 30 } })],
        borders: { top: noBorder, left: noBorder, right: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_TEXT } },
      })
    ),
  });
  const buildSummaryRow = (rowUid: string, defect: any) => {
    const statusText = defect.status === "complete" ? "Complete" : "Open";
    const typeText = defect.recordType === "observation" ? "Obs" : "Defect";
    const rowBorder = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
    const locationText = rowUid === defect.uid ? formatDefectLocation(defect, wordDims) : deriveLocation(safeText(rowUid));
    const cellTexts = [safeText(rowUid), typeText, locationText, getWorkTypeLabel(safeText(defect.uid)), safeText(defect.assignedTo) || "\u2014", safeText(defect.dueDate) || "\u2014", statusText];
    return new TableRow({
      children: cellTexts.map((text: string, i: number) =>
        new TableCell({
          width: { size: summaryHeaderWidths[i], type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: safeText(text), size: 14, font: "Aptos", color: i === 6 ? (statusText === "Complete" ? "228B22" : "C89600") : (i === 1 ? "666666" : undefined), bold: i === 0 })], spacing: { before: 25, after: 25 } })],
          borders: { top: rowBorder, left: noBorder, right: noBorder, bottom: rowBorder },
        })
      ),
    });
  };

  // Emit each category group: itemise => labelled sub-table; summarise => count line.
  let actionHasContent = false;
  for (const g of tree.actionList.groups) {
    if (g.kind === "summary") {
      actionChildren.push(summaryGroupParagraph(g));
      actionHasContent = true;
      continue;
    }
    actionHasContent = true;
    actionChildren.push(new Paragraph({
      children: [new TextRun({ text: safeText(g.label), bold: true, size: 24, font: "Aptos", color: CAPTION_BLUE })],
      spacing: { before: 160, after: 60 },
    }));
    const rows: any[] = [buildSummaryHeaderRow()];
    for (const d of g.defects) {
      rows.push(buildSummaryRow(d.uid, d));
      if (d.locations && d.locations.length > 0) for (const loc of d.locations) rows.push(buildSummaryRow(loc.uid || "", d));
    }
    actionChildren.push(new Table({ width: { size: 9700, type: WidthType.DXA }, layout: TableLayoutType.FIXED, rows }));
  }
  if (!actionHasContent) {
    actionChildren.push(new Paragraph({ children: [new TextRun({ text: "Nothing for this report.", size: 20, font: "Aptos", color: "999999", italics: true })] }));
  }

  // ===================== SECTION 3 — PROJECT STATUS =====================
  const statusChildren: any[] = [];
  statusChildren.push(new Paragraph({
    children: [new TextRun({ text: "Project Status", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_1, spacing: { after: 120 },
  }));
  if (tree.projectStatus.empty) {
    statusChildren.push(new Paragraph({ children: [new TextRun({ text: "Nothing for this report.", size: 20, font: "Aptos", color: "999999", italics: true })] }));
  } else {
    for (const n of tree.projectStatus.narratives) {
      statusChildren.push(new Paragraph({
        children: [
          new TextRun({ text: safeText(n.title) || "Narrative", bold: true, size: 26, font: "Aptos", color: CAPTION_BLUE }),
          ...(n.status ? [new TextRun({ text: `   [${safeText(n.status)}]`, size: 18, font: "Aptos", color: "888888" })] : []),
        ],
        spacing: { before: 120, after: 40 },
      }));
      if (n.body) statusChildren.push(new Paragraph({ children: [new TextRun({ text: safeText(n.body), size: 20, font: "Aptos" })], spacing: { after: 80 } }));
    }
    if (tree.projectStatus.program) {
      const p = tree.projectStatus.program;
      statusChildren.push(new Paragraph({ children: [new TextRun({ text: "Program", bold: true, size: 26, font: "Aptos", color: CAPTION_BLUE })], spacing: { before: 120, after: 40 } }));
      const pRows: string[][] = [];
      if (p.asAtDate) pRows.push(["As at", safeText(p.asAtDate)]);
      if (p.varianceText) pRows.push(["Variance", safeText(p.varianceText)]);
      if (p.projectedCompletion) pRows.push(["Projected completion", safeText(p.projectedCompletion)]);
      if (p.statusNarrative) pRows.push(["Status", safeText(p.statusNarrative)]);
      if (pRows.length > 0) {
        statusChildren.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
          rows: pRows.map(([l, v]) => new TableRow({ children: [
            new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(l), bold: true, size: 18, font: "Aptos" })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
            new TableCell({ width: { size: 70, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(v), size: 18, font: "Aptos" })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
          ] })),
        }));
      }
      if (p.programImageFilename) {
        const blob = await loadImageBlob(p.programImageFilename);
        if (blob) {
          let buffer: ArrayBuffer; try { buffer = await compressImageForExport(blob, 1000, 0.7); } catch { buffer = await blobToArrayBuffer(blob); }
          statusChildren.push(new Paragraph({ children: [new ImageRun({ data: buffer, transformation: { width: 480, height: 270 }, type: "jpg" })], spacing: { before: 80, after: 80 } }));
        }
      }
    }
    if (tree.projectStatus.stageMap) {
      const sm = tree.projectStatus.stageMap;
      statusChildren.push(new Paragraph({ children: [new TextRun({ text: "Stage Map", bold: true, size: 26, font: "Aptos", color: CAPTION_BLUE })], spacing: { before: 120, after: 40 } }));
      let stages: any[] = [];
      try { stages = Array.isArray(sm.stages) ? sm.stages : JSON.parse(sm.stages || "[]"); } catch { stages = []; }
      if (sm.planImageFilename) {
        const blob = await loadImageBlob(sm.planImageFilename);
        if (blob) {
          let buffer: ArrayBuffer; try { buffer = await compressImageForExport(blob, 1000, 0.7); } catch { buffer = await blobToArrayBuffer(blob); }
          statusChildren.push(new Paragraph({ children: [new ImageRun({ data: buffer, transformation: { width: 480, height: 270 }, type: "jpg" })], spacing: { before: 80, after: 80 } }));
        }
      }
      if (stages.length > 0) {
        statusChildren.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
          rows: stages.map((s: any) => new TableRow({ children: [
            new TableCell({ width: { size: 60, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(s.stageName), size: 18, font: "Aptos" })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
            new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(s.status), size: 18, font: "Aptos", color: "666666" })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
          ] })),
        }));
      }
    }
  }

  // ===================== CLIENT-ONLY — PROGRESS SUMMARY =====================
  if (tree.progressSummary) {
    const ps = tree.progressSummary;
    statusChildren.push(new Paragraph({
      children: [new TextRun({ text: "Progress Summary", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
      heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 },
    }));
    const psRows: string[][] = [
      ["Open", String(ps.open)],
      ["Closed this period", String(ps.closedThisPeriod)],
      ["Overdue", String(ps.overdue)],
      ["Total", String(ps.total)],
    ];
    statusChildren.push(new Table({
      width: { size: 60, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
      rows: psRows.map(([l, v]) => new TableRow({ children: [
        new TableCell({ width: { size: 60, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(l), bold: true, size: 18, font: "Aptos" })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
        new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: safeText(v), size: 18, font: "Aptos", color: CAPTION_BLUE })], spacing: { before: 40, after: 40 } })], borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder } }),
      ] })),
    }));
  }

  // ===================== SECTION 4 — THIS INSPECTION =====================
  // Walk new/amended/completed; within each, walk category groups in profile order.
  const buildBucketSection = async (title: string, subtitle: string, groups: CategoryGroup[], showChangeSummary: boolean): Promise<any[]> => {
    const children: any[] = [];
    children.push(new Paragraph({ children: [new TextRun({ text: title, size: 40, font: "Aptos", bold: true, color: DARK_TEXT })], heading: HeadingLevel.HEADING_1, spacing: { after: 80 } }));
    children.push(new Paragraph({ children: [new TextRun({ text: subtitle, size: 18, color: "666666", italics: true, font: "Aptos" })], spacing: { after: 200 } }));
    let first = true;
    for (const g of groups) {
      children.push(new Paragraph({ children: [new TextRun({ text: safeText(g.label), bold: true, size: 26, font: "Aptos", color: CAPTION_BLUE })], spacing: { before: first ? 0 : 160, after: 80 } }));
      first = false;
      if (g.kind === "summary") {
        children.push(summaryGroupParagraph(g));
      } else {
        for (let i = 0; i < g.defects.length; i++) {
          if (i > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
          const els = await buildWordDefectPage(g.defects[i], { showChangeSummary });
          children.push(...els);
        }
      }
    }
    return children;
  };

  // ===================== SECTION 5 — CARRIED-FORWARD REGISTER =====================
  const carriedChildren: any[] = [];
  carriedChildren.push(new Paragraph({ children: [new TextRun({ text: "Carried-forward Register", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })], heading: HeadingLevel.HEADING_1, spacing: { after: 100 } }));
  carriedChildren.push(new Paragraph({ children: [new TextRun({ text: "Open items not covered in this inspection's locations.", size: 18, color: "666666", italics: true, font: "Aptos" })], spacing: { after: 160 } }));
  if (tree.carriedForward.empty) {
    carriedChildren.push(new Paragraph({ children: [new TextRun({ text: "Nothing for this report.", size: 20, font: "Aptos", color: "999999", italics: true })] }));
  } else {
    for (const g of tree.carriedForward.groups) {
      // Carried-forward always itemises (tree enforced ignoreTreatment) — but guard.
      if (g.kind === "summary") { carriedChildren.push(summaryGroupParagraph(g)); continue; }
      carriedChildren.push(new Paragraph({ children: [new TextRun({ text: safeText(g.label), bold: true, size: 24, font: "Aptos", color: CAPTION_BLUE })], spacing: { before: 160, after: 60 } }));
      const rows: any[] = [buildSummaryHeaderRow()];
      for (const d of g.defects) {
        rows.push(buildSummaryRow(d.uid, d));
        if (d.locations && d.locations.length > 0) for (const loc of d.locations) rows.push(buildSummaryRow(loc.uid || "", d));
      }
      carriedChildren.push(new Table({ width: { size: 9700, type: WidthType.DXA }, layout: TableLayoutType.FIXED, rows }));
    }
  }

  // ===================== SECTION 6 — APPENDICES =====================
  const appendixChildren: any[] = [];
  appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Appendices", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })], heading: HeadingLevel.HEADING_1, spacing: { after: 120 } }));
  if (tree.appendixMode === "reference") {
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Technical Method Statements", size: 28, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { after: 60 } }));
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Refer to the contractor report for full technical method statements.", size: 20, font: "Aptos" })], spacing: { after: 160 } }));
    // Coverage drawings STAY for client (communicative, not commercial).
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Coverage Drawings", size: 28, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { after: 60 } }));
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: safeText(data.report.locationsCovered) || "Refer to inspection locations covered.", size: 20, font: "Aptos" })], spacing: { after: 120 } }));
  } else {
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Technical Guidance & Method Statements", size: 28, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { after: 60 } }));
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Technical guidance, Technical Data Sheets (TDS) and method statements applicable to the works are included with this report.", size: 20, font: "Aptos" })], spacing: { after: 120 } }));
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: "Coverage Drawings", size: 28, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { after: 60 } }));
    appendixChildren.push(new Paragraph({ children: [new TextRun({ text: safeText(data.report.locationsCovered) || "Refer to inspection locations covered.", size: 20, font: "Aptos" })], spacing: { after: 120 } }));
  }

  // ===================== BUILD DOCUMENT =====================
  const sectionProps = {
    page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } },
    headers: { default: defaultHeader },
    footers: { default: defaultFooter },
  };

  const docSections: any[] = [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 851, right: 1440, bottom: 1440, left: 1440 } } }, children: coverChildren },
    { properties: sectionProps, children: introChildren },
    { properties: sectionProps, children: actionChildren },
    { properties: sectionProps, children: statusChildren },
  ];

  // This Inspection — emit each non-empty bucket as its own section.
  if (tree.thisInspection.new.length > 0) {
    docSections.push({ properties: sectionProps, children: await buildBucketSection("NEW THIS INSPECTION", "Items added during this inspection.", tree.thisInspection.new, false) });
  }
  if (tree.thisInspection.amended.length > 0) {
    docSections.push({ properties: sectionProps, children: await buildBucketSection("AMENDED THIS INSPECTION", "Existing items updated during this inspection.", tree.thisInspection.amended, true) });
  }
  if (tree.thisInspection.completed.length > 0) {
    docSections.push({ properties: sectionProps, children: await buildBucketSection("COMPLETED THIS INSPECTION", "Items marked complete during this inspection.", tree.thisInspection.completed, true) });
  }
  if (tree.thisInspection.empty) {
    docSections.push({ properties: sectionProps, children: [
      new Paragraph({ children: [new TextRun({ text: "This Inspection", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })], heading: HeadingLevel.HEADING_1, spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Nothing for this report.", size: 20, font: "Aptos", color: "999999", italics: true })] }),
    ] });
  }

  docSections.push({ properties: sectionProps, children: carriedChildren });
  docSections.push({ properties: sectionProps, children: appendixChildren });

  const wordDoc = new Document({
    creator: "Angel Façade Consulting",
    title: `${safeText(data.project.name)} — Site Visit Report`,
    styles: { default: { document: { run: { font: "Aptos", size: 20, color: DARK_TEXT } } } },
    sections: docSections,
  });

  return Packer.toBlob(wordDoc);
}
