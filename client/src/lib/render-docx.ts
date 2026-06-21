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
  getLocationDimensions, formatReportDate, formatPhotoDate, isDefectOverdue,
  comparePhotoSlots, photoSlotLabel,
  truncateWordBoundary, resolveActionSummary,
  parseProjectSnapshot, resolveProjectField,
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

  const DARK_BLUE = "0A1D30";
  const ACCENT_BLUE = "00CDC8";
  const DARK_TEXT = "3A3A3A";
  const CAPTION_BLUE = "0E2841";
  const OVERDUE_RED = "C00000";

  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const accentBottomBorder = { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE };
  const accentTopBorder = { style: BorderStyle.SINGLE, size: 6, color: ACCENT_BLUE };

  const reportDate = formatReportDate(new Date().toISOString());
  const rev = data.report.revision || "01";

  // §2.3 — prefer frozen snapshot over live project row so historical reports
  // keep their original wording even after the project is later edited.
  const snap = parseProjectSnapshot((data.report as any).projectSnapshot);
  const projAddress     = resolveProjectField(snap, data.project, "address");
  const projName        = resolveProjectField(snap, data.project, "name");
  const projReportTitle = resolveProjectField(snap, data.project, "reportTitle");
  const projInspector   = resolveProjectField(snap, data.project, "inspector");
  const afcRef          = resolveProjectField(snap, data.project, "afcReference") || "AFC-24XXX";

  // §1.7 inspection number, zero-padded to 2 digits for the title-page heading.
  const inspNumRaw = safeText(data.report.inspectionNumber).trim();
  const inspNumPadded = inspNumRaw
    ? (/^\d+$/.test(inspNumRaw) ? inspNumRaw.padStart(2, "0") : inspNumRaw)
    : "";
  const siteVisitHeading = inspNumPadded ? `SITE VISIT REPORT ${inspNumPadded}` : "SITE VISIT REPORT";

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [new Paragraph({
              children: [new TextRun({ text: safeText(projName).toUpperCase(), size: 16, font: "Aptos", bold: true, color: DARK_TEXT })],
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
          new TextRun({ text: `${safeText(afcRef)} | ${safeText(projAddress)}`, size: 16, font: "Aptos", color: DARK_TEXT }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: "Page ", size: 16, font: "Aptos", color: DARK_TEXT }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Aptos", color: DARK_TEXT }),
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        border: { top: accentTopBorder },
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
  // §Title-page spec — vertical stack (top → bottom):
  //   1) Address (large, dark blue)
  //   2) Report Title (full, no truncation)
  //   3) Site Visit Report NN (zero-padded inspection number)
  //   4) Revision N
  //   5) Date
  //   6) Angel Façade Consulting
  //   7) {Inspector} | 0407 759 590
  //   8) AFC reference
  coverChildren.push(new Paragraph({ spacing: { before: 1200 } }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: safeText(projAddress).toUpperCase(), bold: true, size: 56, font: "Aptos", color: DARK_BLUE })],
    spacing: { after: 120 },
  }));
  if (projReportTitle) {
    coverChildren.push(new Paragraph({
      children: [new TextRun({ text: safeText(projReportTitle), size: 36, font: "Aptos", color: DARK_BLUE })],
      spacing: { after: 120 },
    }));
  }
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: siteVisitHeading, size: 44, font: "Aptos", color: DARK_BLUE, smallCaps: true })],
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
    children: [new TextRun({ text: `${safeText(projInspector)} | 0407 759 590`, size: 22, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 40 },
  }));
  coverChildren.push(new Paragraph({
    children: [new TextRun({ text: safeText(afcRef), size: 22, font: "Aptos", color: DARK_TEXT })],
    spacing: { after: 200 },
  }));

  // ===================== SECTION 1 — INTRODUCTION =====================
  // Snapshot-resolved fields for §1.1 boilerplate substitution.
  const projClient = resolveProjectField(snap, data.project, "client");

  const introChildren: any[] = [];
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "Introduction", size: 40, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_1, spacing: { after: 200 },
  }));

  // §1.1 General — boilerplate per AFC template. Wording is fixed; only CLIENT
  // and ADDRESS are substituted, sourced from the frozen snapshot when present.
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "General", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_2, spacing: { after: 100 },
  }));
  introChildren.push(new Paragraph({
    children: [new TextRun({
      text: `Angel Façade Consulting (AFC) was engaged by ${safeText(projClient)} to carry out regular inspections of the remedial works underway at ${safeText(projAddress)}. Below is a summary of pertinent project information.`,
      size: 20, font: "Aptos",
    })],
    spacing: { after: 200 },
  }));

  // §1.1 Roles table (Role / Entity / Contact Details). Sourced from the frozen
  // snapshot's roles JSON; falls back to the live project row when snapshot is
  // missing (legacy reports). Renders only when there is at least one role.
  const rolesJson = resolveProjectField(snap, data.project, "roles") || "[]";
  let roles: Array<{ role?: string; entity?: string; contactDetails?: string }> = [];
  try { roles = JSON.parse(rolesJson) || []; } catch { roles = []; }
  if (Array.isArray(roles) && roles.length > 0) {
    const headerCell = (text: string, widthPct: number) => new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: "Aptos", color: DARK_TEXT })], spacing: { before: 60, after: 60 } })],
      borders: { top: thinBorder, left: noBorder, right: noBorder, bottom: thinBorder },
      shading: { fill: "F2F2F2" } as any,
    });
    const bodyCell = (text: string, widthPct: number) => new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      children: safeText(text).split(/\n/).map((line) => new Paragraph({
        children: [new TextRun({ text: line, size: 20, font: "Aptos" })],
        spacing: { before: 40, after: 40 },
      })),
      borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
    });
    introChildren.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          children: [
            headerCell("Role", 25),
            headerCell("Entity", 30),
            headerCell("Contact Details", 45),
          ],
          tableHeader: true,
        }),
        ...roles.map((r) => new TableRow({
          children: [
            bodyCell(safeText(r.role), 25),
            bodyCell(safeText(r.entity), 30),
            bodyCell(safeText(r.contactDetails), 45),
          ],
        })),
      ],
    }));
    introChildren.push(new Paragraph({ spacing: { before: 200 } }));
  }

  // §1.2 Scope of works — table of locations covered by the engagement.
  // Columns: Area Ref / Location (Elevation/Floor) / Work Item / Access Method.
  // Sourced from the frozen snapshot's scopeOfWorks JSON.
  const scopeJson = resolveProjectField(snap, data.project, "scopeOfWorks") || "[]";
  let scope: Array<{ areaRef?: string; location?: string; workItem?: string; accessMethod?: string }> = [];
  try { scope = JSON.parse(scopeJson) || []; } catch { scope = []; }
  if (Array.isArray(scope) && scope.length > 0) {
    introChildren.push(new Paragraph({
      children: [new TextRun({ text: "Scope of works", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
      heading: HeadingLevel.HEADING_2, spacing: { after: 100 },
    }));
    const sHeader = (text: string, widthPct: number) => new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: "Aptos", color: DARK_TEXT })], spacing: { before: 60, after: 60 } })],
      borders: { top: thinBorder, left: noBorder, right: noBorder, bottom: thinBorder },
      shading: { fill: "F2F2F2" } as any,
    });
    const sBody = (text: string, widthPct: number) => new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      children: safeText(text).split(/\n/).map((line) => new Paragraph({
        children: [new TextRun({ text: line, size: 20, font: "Aptos" })],
        spacing: { before: 40, after: 40 },
      })),
      borders: { top: noBorder, left: noBorder, right: noBorder, bottom: thinBorder },
    });
    introChildren.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          children: [
            sHeader("Area Ref", 18),
            sHeader("Location", 27),
            sHeader("Work item", 30),
            sHeader("Access method", 25),
          ],
          tableHeader: true,
        }),
        ...scope.map((s) => new TableRow({
          children: [
            sBody(safeText(s.areaRef), 18),
            sBody(safeText(s.location), 27),
            sBody(safeText(s.workItem), 30),
            sBody(safeText(s.accessMethod), 25),
          ],
        })),
      ],
    }));
    introChildren.push(new Paragraph({ spacing: { before: 200 } }));
  }

  // §1.3 Inspection particulars (renumbered from previous §1.2 Inspection).
  introChildren.push(new Paragraph({
    children: [new TextRun({ text: "Inspection particulars", size: 32, font: "Aptos", bold: true, color: DARK_TEXT })],
    heading: HeadingLevel.HEADING_2, spacing: { after: 100 },
  }));

  const inspectionData: string[][] = [];
  if (data.report.inspectionDate) inspectionData.push(["Date", formatReportDate(data.report.inspectionDate)]);
  else inspectionData.push(["Date", reportDate]);
  if (data.report.inspectionNumber) inspectionData.push(["Inspection Number", safeText(data.report.inspectionNumber)]);
  inspectionData.push(["Inspector", safeText(projInspector)]);
  inspectionData.push(["Locations covered", safeText(data.report.locationsCovered) || safeText(projAddress)]);
  inspectionData.push(["Client", safeText(projClient)]);
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
          // Prefer the server-computed wipNumber (1-based position in the cumulative,
          // date-sorted timeline). Falls back to the stored slot label when absent.
          const photoLabel = photo.slot === "complete"
            ? photoSlotLabel(photo.slot)
            : (typeof photo.wipNumber === "number" ? `WIP ${photo.wipNumber}` : photoSlotLabel(photo.slot));
          cellChildren.push(new Paragraph({
            children: [
              new TextRun({ text: safeText(photoLabel), bold: true, size: 16, font: "Aptos", color: photo.slot === "complete" ? "228B22" : "666666" }),
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

  // The Action List uses a dedicated 8-column AI layout (§2.2):
  //   UID | Location | Work item | Observation (truncated) | Action (AI) |
  //   Responsible | Due Date | Status
  // Carried-forward keeps the original Category-aware layout below.
  const summaryHeaderLabels = ["ID", "Type", "Location", "Work Type", "Responsible", "By Date", "Status"];
  const summaryHeaderWidths = [950, 600, 1300, 2300, 1850, 1300, 1400];
  // Category-aware layout: insert "Category" after "Responsible" (index 5) and
  // shrink Work Type (the description column) to keep the total width constant.
  const summaryHeaderLabelsCat = ["ID", "Type", "Location", "Work Type", "Responsible", "Category", "By Date", "Status"];
  const summaryHeaderWidthsCat = [950, 600, 1300, 900, 1850, 1400, 1300, 1400];
  // Action List AI layout. Widths sum to 9700 DXA (same total as the other modes).
  const summaryHeaderLabelsAct = ["UID", "Location", "Work item", "Observation", "Action", "Responsible", "Due Date", "Status"];
  const summaryHeaderWidthsAct = [900, 1100, 900, 1900, 2100, 1200, 800, 800];
  type SummaryMode = { showCategory?: boolean; showAction?: boolean };
  const buildSummaryHeaderRow = (mode: SummaryMode = {}) => {
    const labels = mode.showAction
      ? summaryHeaderLabelsAct
      : mode.showCategory
        ? summaryHeaderLabelsCat
        : summaryHeaderLabels;
    const widths = mode.showAction
      ? summaryHeaderWidthsAct
      : mode.showCategory
        ? summaryHeaderWidthsCat
        : summaryHeaderWidths;
    return new TableRow({
      tableHeader: true,
      children: labels.map((label, i) =>
        new TableCell({
          width: { size: widths[i], type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: label, size: 16, font: "Aptos", bold: true, color: CAPTION_BLUE })], spacing: { before: 30, after: 30 } })],
          borders: { top: noBorder, left: noBorder, right: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_TEXT } },
        })
      ),
    });
  };
  const buildSummaryRow = (rowUid: string, defect: any, mode: SummaryMode = {}) => {
    const showCategory = !!mode.showCategory;
    const showAction = !!mode.showAction;
    const statusText = defect.status === "complete" ? "Complete" : "Open";
    // Overdue := dueDate exists AND dueDate < today AND the item is NOT closed.
    // Closed/Complete (and Archived) items are NEVER flagged overdue even if
    // their dueDate is in the past. (Shared helper — see render-helpers.ts.)
    const overdue = isDefectOverdue(defect);
    const typeText = defect.recordType === "observation" ? "Obs" : "Defect";
    const rowBorder = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
    const locationText = rowUid === defect.uid ? formatDefectLocation(defect, wordDims) : deriveLocation(safeText(rowUid));
    const widths = showAction
      ? summaryHeaderWidthsAct
      : showCategory
        ? summaryHeaderWidthsCat
        : summaryHeaderWidths;
    const categoryText = safeText(defect.categoryLabel) || "(uncategorised)";
    // Action List rows: only emit the cached AI summary / fallback for parent
    // rows (rowUid === defect.uid). Sub-location rows leave Observation/Action
    // blank to avoid duplication — the parent row already conveys the work for
    // that defect.
    const isParentRow = rowUid === defect.uid;
    const observationCell = showAction
      ? (isParentRow ? truncateWordBoundary(defect.comment, { maxWords: 12, maxChars: 80 }) : "")
      : "";
    const actionCell = showAction
      ? (isParentRow ? resolveActionSummary(defect) : "")
      : "";
    const cellTexts = showAction
      ? [safeText(rowUid), locationText, getWorkTypeLabel(safeText(defect.uid)), observationCell, actionCell, safeText(defect.assignedTo) || "\u2014", safeText(defect.dueDate) || "\u2014", statusText]
      : showCategory
        ? [safeText(rowUid), typeText, locationText, getWorkTypeLabel(safeText(defect.uid)), safeText(defect.assignedTo) || "\u2014", categoryText, safeText(defect.dueDate) || "\u2014", statusText]
        : [safeText(rowUid), typeText, locationText, getWorkTypeLabel(safeText(defect.uid)), safeText(defect.assignedTo) || "\u2014", safeText(defect.dueDate) || "\u2014", statusText];
    const statusIdx = showAction ? 7 : showCategory ? 7 : 6;
    // Type column index: index 1 in non-action modes, absent (-1) in action mode.
    const typeColIdx = showAction ? -1 : 1;
    const statusColor = statusText === "Complete" ? "228B22" : "C89600";
    // The Status cell shows the base status, and for overdue items appends a
    // red " - Overdue" suffix as a second TextRun (e.g. "Open - Overdue").
    const buildStatusCellRuns = () => {
      const runs = [new TextRun({ text: statusText, size: 14, font: "Aptos", color: statusColor, bold: false })];
      if (overdue) {
        runs.push(new TextRun({ text: " - Overdue", size: 14, font: "Aptos", color: OVERDUE_RED, bold: false }));
      }
      return runs;
    };
    return new TableRow({
      children: cellTexts.map((text: string, i: number) =>
        new TableCell({
          width: { size: widths[i], type: WidthType.DXA },
          children: [new Paragraph({
            children: i === statusIdx
              ? buildStatusCellRuns()
              : [new TextRun({ text: safeText(text), size: 14, font: "Aptos", color: i === typeColIdx ? "666666" : undefined, bold: i === 0 })],
            spacing: { before: 25, after: 25 },
          })],
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
    const rows: any[] = [buildSummaryHeaderRow({ showAction: true })];
    for (const d of g.defects) {
      rows.push(buildSummaryRow(d.uid, d, { showAction: true }));
      if (d.locations && d.locations.length > 0) for (const loc of d.locations) rows.push(buildSummaryRow(loc.uid || "", d, { showAction: true }));
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
