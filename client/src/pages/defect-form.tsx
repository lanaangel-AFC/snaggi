import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link, useParams, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";
import { ArrowLeft, Camera, Upload, X, ImageIcon, Save, CheckCircle2, Clock, Wrench, Mic, Download, MapPin, ChevronDown, History, Copy, PenLine, Plus, Trash2 } from "lucide-react";
import { Collapsible, CollapsibleTrigger, CollapsibleContent } from "@/components/ui/collapsible";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { DictationButton } from "@/components/DictationButton";
import type { Defect, Photo, Project, Elevation, DefectLocation } from "@shared/schema";
import { useState, useRef, useEffect, useMemo, useCallback } from "react";

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

// Work type codes for facade repair
const WORK_TYPES = [
  { code: "CR", label: "Concrete Repair" },
  { code: "CK", label: "Caulking" },
  { code: "PT", label: "Painting" },
  { code: "WP", label: "Waterproofing" },
  { code: "GL", label: "Glazing" },
  { code: "CL", label: "Cladding" },
  { code: "ST", label: "Structural" },
  { code: "SE", label: "Sealant" },
  { code: "FL", label: "Flashing" },
  { code: "RR", label: "Render Repair" },
  { code: "GK", label: "Gasket" },
  { code: "SR", label: "Steel Repair" },
  { code: "SM", label: "Steel Removal" },
  { code: "BR", label: "Brick Repair" },
  { code: "LR", label: "Lintel Repair" },
  { code: "SS", label: "Sill Stabilisation" },
  { code: "GW", label: "General Works" },
  { code: "OT", label: "Other" },
];

const BUILTIN_WORK_TYPE_CODES = WORK_TYPES.map((w) => w.code);
const FORBIDDEN_WORK_TYPE_CODES = new Set(["LEV", "LEVEL", "L"]);

// Reject custom work-type codes that collide with the level field or a built-in code.
// Returns an error message if invalid, or null if the code is acceptable.
function validateCustomWorkTypeCode(code: string): string | null {
  const c = code.trim().toUpperCase();
  if (c.length < 2) return "Work type code must be at least 2 characters.";
  if (FORBIDDEN_WORK_TYPE_CODES.has(c)) return `"${c}" is reserved (it collides with the Level field). Choose a different code.`;
  if (BUILTIN_WORK_TYPE_CODES.includes(c)) return `"${c}" is already a built-in work type. Pick it from the list instead.`;
  return null;
}

// Elevation options for UID
const ELEVATIONS = [
  { code: "N", label: "North" },
  { code: "S", label: "South" },
  { code: "E", label: "East" },
  { code: "W", label: "West" },
  { code: "NE", label: "North East" },
  { code: "NW", label: "North West" },
  { code: "SE", label: "South East" },
  { code: "SW", label: "South West" },
];

// Standardised person roles for dropdowns
const PERSON_ROLES = [
  "Consultant",
  "Contractor",
  "Client",
  "Project Manager",
  "Facility Manager",
];

const PHOTO_SLOTS = [
  { key: "wip1", label: "WIP 1", description: "Progress photo 1" },
  { key: "wip2", label: "WIP 2", description: "Progress photo 2" },
  { key: "wip3", label: "WIP 3", description: "Progress photo 3" },
  { key: "wip4", label: "WIP 4", description: "Progress photo 4" },
  { key: "wip5", label: "WIP 5", description: "Progress photo 5" },
  { key: "complete", label: "Complete", description: "Finished repair" },
] as const;

type SlotKey = "wip1" | "wip2" | "wip3" | "wip4" | "wip5" | "complete";

function formatHistoryDate(dateStr?: string): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric" });
}

export default function DefectForm() {
  const { projectId, reportId, defectId } = useParams<{ projectId: string; reportId: string; defectId?: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const fileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const cameraInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  // Determine if editing or creating, and record type from URL path
  const isEdit = !!defectId && defectId !== "new-defect" && defectId !== "new-observation";
  const isNewObservation = defectId === "new-observation" || (typeof window !== "undefined" && window.location.hash.includes("/new-observation"));

  // Two-step flow: Step 1 = pick work type, Step 2 = full form
  const [formStep, setFormStep] = useState<"workType" | "form">(isEdit ? "form" : "workType");
  const [showOtherWorkTypes, setShowOtherWorkTypes] = useState(false);

  const [form, setForm] = useState({
    dateOpened: new Date().toISOString().split("T")[0],
    dateClosed: "",
    comment: "",
    actionRequired: "",
    assignedTo: "",
    dueDate: "",
    verificationMethod: "",
    verificationPerson: "",
    status: "open",
  });

  const [recordType, setRecordType] = useState(() => isNewObservation ? "observation" : "defect");

  // Fetch project and report to get configured elevations
  const { data: project } = useQuery<Project>({
    queryKey: ["/api/projects", projectId],
  });

  const { data: report } = useQuery<any>({
    queryKey: ["/api/reports", reportId],
    enabled: !!reportId,
  });

  const { data: globalSettings } = useQuery<{ workTypes: { code: string; label: string }[] }>({
    queryKey: ["/api/global-settings"],
  });

  const enabledUidParts = useMemo(() => {
    try { return JSON.parse((project as any)?.enabledUidParts || '{}'); } catch { return { elevation: true, drop: true, level: true, workType: true }; }
  }, [(project as any)?.enabledUidParts]);

  const primaryWorkTypeCodes: string[] = useMemo(() => {
    try { return JSON.parse((project as any)?.primaryWorkTypes || '[]'); } catch { return []; }
  }, [(project as any)?.primaryWorkTypes]);

  // Skip work type step if workType is disabled in UID parts
  useEffect(() => {
    if (!isEdit && formStep === "workType" && project) {
      try {
        const parts = JSON.parse((project as any)?.enabledUidParts || '{}');
        if (parts.workType === false) setFormStep("form");
      } catch {}
    }
  }, [project, isEdit, formStep]);

  // Build elevation options — prefer report-level, fall back to project-level
  const elevationOptions = useMemo(() => {
    const codeMap: Record<string, string> = {
      "North": "N", "South": "S", "East": "E", "West": "W",
      "North East": "NE", "North West": "NW", "South East": "SE", "South West": "SW",
    };
    // Try report elevations first, then project
    const sources = [report?.elevations, project?.elevations];
    for (const src of sources) {
      if (!src) continue;
      try {
        const configured: string[] = JSON.parse(src as string);
        if (configured.length > 0) {
          return configured.map((label) => ({
            code: codeMap[label] || label.substring(0, 3).toUpperCase(),
            label,
          }));
        }
      } catch {}
    }
    return ELEVATIONS;
  }, [report?.elevations, project?.elevations]);

  const [elevation, setElevation] = useState("");
  const [drop, setDrop] = useState("01");
  const [level, setLevel] = useState("");
  const [workType, setWorkType] = useState("");
  const [seqNumber, setSeqNumber] = useState("01");

  const [uid, setUid] = useState("");
  const [photos, setPhotos] = useState<Photo[]>([]);
  const [uploading, setUploading] = useState<string | null>(null);
  // Track whether we've already initialized the form from the loaded defect.
  // Without this, refetches (e.g. on window focus after returning from the photo
  // picker) would overwrite unsaved edits to Drop / Elevation / Level / Comment.
  const hasInitializedRef = useRef(false);

  // "Mark on Elevation" state
  const [elevationPickerOpen, setElevationPickerOpen] = useState(false);
  const [projectElevations, setProjectElevations] = useState<Elevation[]>([]);
  const [obsHistoryOpen, setObsHistoryOpen] = useState(false);
  const [actHistoryOpen, setActHistoryOpen] = useState(false);
  const [obsNoteOpen, setObsNoteOpen] = useState(false);
  const [obsNoteText, setObsNoteText] = useState("");
  const [actNoteOpen, setActNoteOpen] = useState(false);
  const [actNoteText, setActNoteText] = useState("");
  const [priorHistoryOpen, setPriorHistoryOpen] = useState(false);
  const [inspNoteOpen, setInspNoteOpen] = useState(false);
  const [inspNoteText, setInspNoteText] = useState("");
  const commentRef = useRef<HTMLTextAreaElement>(null);
  const actionRef = useRef<HTMLTextAreaElement>(null);

  // Autosave for comment/actionRequired fields
  const formRef = useRef(form);
  useEffect(() => { formRef.current = form; }, [form]);
  const autosaveTimer = useRef<NodeJS.Timeout | null>(null);
  const [autosaveStatus, setAutosaveStatus] = useState<"idle" | "saving" | "saved">("idle");

  // === Additional Locations ===
  const { data: existingLocations } = useQuery<DefectLocation[]>({
    queryKey: [`/api/defects/${defectId}/locations`],
    enabled: isEdit,
  });
  const [additionalLocations, setAdditionalLocations] = useState<DefectLocation[]>([]);
  const additionalLocationsRef = useRef<DefectLocation[]>([]);
  // Keep the ref in sync with state so debounced callbacks always read current values
  useEffect(() => { additionalLocationsRef.current = additionalLocations; }, [additionalLocations]);
  const locationPatchTimers = useRef<Record<number, NodeJS.Timeout>>({});

  useEffect(() => {
    if (existingLocations) {
      setAdditionalLocations(existingLocations);
    }
  }, [existingLocations]);

  // Compute a suggested UID for a location based on parent defect's work type/number
  const computeLocationUid = (locElevation: string, locDrop: string, locLevel: string): string => {
    const wt = workType || "";
    const num = seqNumber || "01";
    if (!locElevation && !locDrop && !locLevel) return "";
    const elev = locElevation || elevation || "";
    const dd = locDrop ? locDrop.padStart(2, "0") : "";
    const ll = locLevel ? locLevel.padStart(2, "0") : "";
    const segments = [elev, dd, ll, wt, num ? num.padStart(2, "0") : ""].filter((s) => s !== "");
    return segments.join("-");
  };

  const addAdditionalLocation = async () => {
    if (!defectId || defectId === "new-defect" || defectId === "new-observation") {
      toast({ title: "Save the defect first before adding locations", variant: "destructive" });
      return;
    }
    try {
      const res = await apiRequest("POST", `/api/defects/${defectId}/locations`, {
        drop: "", elevation: "", level: "", description: "", uid: "",
      });
      const newLoc: DefectLocation = await res.json();
      setAdditionalLocations(prev => [...prev, newLoc]);
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/locations`] });
    } catch (err: any) {
      toast({ title: err.message || "Failed to add location", variant: "destructive" });
    }
  };

  const updateLocationField = (locId: number, field: string, value: string) => {
    setAdditionalLocations(prev => prev.map(l => {
      if (l.id !== locId) return l;
      const updated = { ...l, [field]: value };
      // Auto-compute UID when drop/elevation/level change
      if (field === "drop" || field === "elevation" || field === "level") {
        updated.uid = computeLocationUid(
          field === "elevation" ? value : (l.elevation || ""),
          field === "drop" ? value : (l.drop || ""),
          field === "level" ? value : (l.level || ""),
        );
      }
      return updated;
    }));
    // Debounced PATCH — read from ref inside the timer to avoid stale closures
    if (locationPatchTimers.current[locId]) clearTimeout(locationPatchTimers.current[locId]);
    locationPatchTimers.current[locId] = setTimeout(async () => {
      const loc = additionalLocationsRef.current.find(l => l.id === locId);
      if (!loc) return;
      const patch: Record<string, string> = { [field]: value };
      if (field === "drop" || field === "elevation" || field === "level") {
        patch.uid = computeLocationUid(
          field === "elevation" ? value : (loc.elevation || ""),
          field === "drop" ? value : (loc.drop || ""),
          field === "level" ? value : (loc.level || ""),
        );
      }
      try {
        await apiRequest("PATCH", `/api/defect-locations/${locId}`, patch);
      } catch {}
    }, 800);
  };

  // Flush all pending debounced location PATCHes immediately (awaitable).
  // Must be called before the parent defect save so location writes complete first.
  const flushPendingLocationPatches = useCallback(async () => {
    const timerIds = Object.keys(locationPatchTimers.current).map(Number);
    if (timerIds.length === 0) return;
    const promises: Promise<void>[] = [];
    for (const locId of timerIds) {
      clearTimeout(locationPatchTimers.current[locId]);
      delete locationPatchTimers.current[locId];
      const loc = additionalLocationsRef.current.find(l => l.id === locId);
      if (!loc) continue;
      // Send full location state so nothing is lost
      const patch: Record<string, string> = {};
      if (loc.elevation) patch.elevation = loc.elevation;
      if (loc.drop) patch.drop = loc.drop;
      if (loc.level) patch.level = loc.level;
      if (loc.description) patch.description = loc.description;
      patch.uid = loc.uid || computeLocationUid(loc.elevation || "", loc.drop || "", loc.level || "");
      promises.push(
        apiRequest("PATCH", `/api/defect-locations/${locId}`, patch).then(() => {}).catch(() => {}),
      );
    }
    await Promise.all(promises);
  }, []);

  // Debounced autosave for comment and actionRequired — fires 800ms after last keystroke
  const triggerAutosave = useCallback(() => {
    if (!isEdit || !defectId) return;
    if (autosaveTimer.current) clearTimeout(autosaveTimer.current);
    autosaveTimer.current = setTimeout(async () => {
      const current = formRef.current;
      setAutosaveStatus("saving");
      try {
        await apiRequest("PATCH", `/api/defects/${defectId}`, {
          comment: current.comment,
          actionRequired: current.actionRequired,
        });
        setAutosaveStatus("saved");
        setTimeout(() => setAutosaveStatus("idle"), 2000);
      } catch {
        setAutosaveStatus("idle");
      }
    }, 800);
  }, [isEdit, defectId]);

  // Cancel autosave timer on unmount
  useEffect(() => () => { if (autosaveTimer.current) clearTimeout(autosaveTimer.current); }, []);

  const deleteAdditionalLocation = async (locId: number) => {
    if (!confirm("Delete this location?")) return;
    try {
      await apiRequest("DELETE", `/api/defect-locations/${locId}`);
      setAdditionalLocations(prev => prev.filter(l => l.id !== locId));
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/locations`] });
      toast({ title: "Location removed" });
    } catch (err: any) {
      toast({ title: err.message || "Failed to remove location", variant: "destructive" });
    }
  };

  const handleMarkLocationOnElevation = async (loc: DefectLocation) => {
    try {
      const res = await apiRequest("GET", `/api/projects/${projectId}/elevations`);
      const elevationsList: Elevation[] = await res.json();
      const locUid = loc.uid || computeLocationUid(loc.elevation || "", loc.drop || "", loc.level || "");
      if (elevationsList.length === 0) {
        toast({ title: "No elevations uploaded for this project yet.", variant: "destructive" });
      } else if (elevationsList.length === 1) {
        navigate(`/projects/${projectId}/elevations/${elevationsList[0].id}?defect=${encodeURIComponent(locUid)}&locationId=${loc.id}`);
      } else {
        setProjectElevations(elevationsList);
        setLocationPickerTarget(loc);
        setElevationPickerOpen(true);
      }
    } catch {
      toast({ title: "Failed to load elevations", variant: "destructive" });
    }
  };

  // Track which location triggered the elevation picker (null = parent defect)
  const [locationPickerTarget, setLocationPickerTarget] = useState<DefectLocation | null>(null);

  const handleMarkOnElevation = async () => {
    try {
      const res = await apiRequest("GET", `/api/projects/${projectId}/elevations`);
      const elevations: Elevation[] = await res.json();
      const defectUid = uid || assembledUid;
      if (elevations.length === 0) {
        toast({ title: "No elevations uploaded for this project yet. Upload one from the project page first.", variant: "destructive" });
      } else if (elevations.length === 1) {
        navigate(`/projects/${projectId}/elevations/${elevations[0].id}?defect=${encodeURIComponent(defectUid)}`);
      } else {
        setProjectElevations(elevations);
        setElevationPickerOpen(true);
      }
    } catch {
      toast({ title: "Failed to load elevations", variant: "destructive" });
    }
  };

  // Fetch all project defects for text suggestions (across all reports)
  const { data: allProjectDefects } = useQuery<Defect[]>({
    queryKey: [`/api/projects/${projectId}/defects`],
  });

  // Build suggestion lists filtered by selected work type
  const textSuggestions = useMemo(() => {
    if (!allProjectDefects || !workType) return { comments: [], actions: [] };
    const matching = allProjectDefects.filter((d) => {
      const parts = d.uid.split("-");
      return parts.some((p) => p === workType);
    });
    const commentSet = new Set<string>();
    const actionSet = new Set<string>();
    matching.forEach((d) => {
      if (d.comment?.trim()) commentSet.add(d.comment.trim());
      if (d.actionRequired?.trim()) actionSet.add(d.actionRequired.trim());
    });
    return { comments: Array.from(commentSet), actions: Array.from(actionSet) };
  }, [allProjectDefects, workType]);

  // Build custom option lists from project
  const dropOptions = useMemo(() => {
    try {
      const custom: string[] = JSON.parse((project as any)?.customDrops || "[]");
      if (custom.length > 0) return custom;
    } catch {}
    return Array.from({ length: 10 }, (_, i) => String(i + 1).padStart(2, "0"));
  }, [(project as any)?.customDrops]);

  const levelOptions = useMemo(() => {
    try {
      const custom: string[] = JSON.parse((project as any)?.customLevels || "[]");
      if (custom.length > 0) return custom;
    } catch {}
    return Array.from({ length: 20 }, (_, i) => String(i + 1).padStart(2, "0"));
  }, [(project as any)?.customLevels]);

  const allWorkTypes = useMemo(() => {
    const custom: { code: string; label: string }[] = (() => {
      try { return JSON.parse((project as any)?.customWorkTypes || "[]"); } catch { return []; }
    })();
    return [...WORK_TYPES, ...custom];
  }, [(project as any)?.customWorkTypes]);

  // Flexible UID builder — only includes present segments
  const uidPrefix = useMemo(() => {
    const segments = [
      elevation || "",
      drop ? drop.padStart(2, "0") : "",
      level ? level.padStart(2, "0") : "",
      workType || "",
    ].filter((s) => s !== "");
    return segments.join("-");
  }, [elevation, drop, level, workType]);

  // Full assembled UID — includes number if present
  const assembledUid = useMemo(() => {
    const segments = [
      elevation || "",
      drop ? drop.padStart(2, "0") : "",
      level ? level.padStart(2, "0") : "",
      workType || "",
      seqNumber ? seqNumber.padStart(2, "0") : "",
    ].filter((s) => s !== "");
    return segments.join("-");
  }, [elevation, drop, level, workType, seqNumber]);

  // Fetch all defects in this report to check for duplicate UIDs
  const { data: reportDefects } = useQuery<Defect[]>({
    queryKey: [`/api/reports/${reportId}/defects`],
    enabled: !isEdit && !!reportId,
  });

  // Check if assembled UID matches an existing entry in this report
  const [duplicateAlertShown, setDuplicateAlertShown] = useState<string | null>(null);

  useEffect(() => {
    if (isEdit || !assembledUid || !reportDefects || assembledUid === duplicateAlertShown) return;
    const match = reportDefects.find((d) => d.uid === assembledUid);
    if (match) {
      setDuplicateAlertShown(assembledUid);
      const typeLabel = (match as any).recordType === "observation" ? "observation" : "defect";
      const confirmed = window.confirm(
        `A ${typeLabel} with UID "${assembledUid}" already exists:\n\n"${match.comment?.substring(0, 80)}..."\n\nWould you like to open it instead?`
      );
      if (confirmed) {
        navigate(`/projects/${projectId}/reports/${reportId}/defects/${match.id}`, { replace: true });
      }
    }
  }, [assembledUid, reportDefects, isEdit, duplicateAlertShown]);

  // Fetch existing defect if editing
  const { data: existingDefect } = useQuery<Defect>({
    queryKey: ["/api/defects", defectId],
    enabled: isEdit,
  });

  // Fetch photos for existing defect
  const { data: existingPhotos } = useQuery<Photo[]>({
    queryKey: [`/api/defects/${defectId}/photos`],
    enabled: isEdit,
  });

  // Fetch observation and action history for existing defect
  const { data: obsHistory } = useQuery<Array<{ id: number; defectId: number; reportId: number; text: string; createdAt: string; reportName: string; reportDate: string }>>({
    queryKey: [`/api/defects/${defectId}/observation-history`],
    enabled: isEdit,
  });
  const { data: actHistory } = useQuery<Array<{ id: number; defectId: number; reportId: number; text: string; createdAt: string; reportName: string; reportDate: string }>>({
    queryKey: [`/api/defects/${defectId}/action-history`],
    enabled: isEdit,
  });
  // Unified, merged history (observation + action + note + status), newest-first, with
  // age (current vs prior) computed relative to the report being viewed.
  const { data: unifiedHistory } = useQuery<Array<{ kind: string; date: string; inspectionNumber: string; author: string; text: string; age: "current" | "prior"; reportId: number | null }>>({
    queryKey: [`/api/defects/${defectId}/history`, reportId],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/defects/${defectId}/history?currentReportId=${reportId}`);
      return res.json();
    },
    enabled: isEdit,
  });

  // Auto-suggest next number when workType changes (based on existing defects in report)
  useEffect(() => {
    if (isEdit || !workType || !reportDefects) return;
    // Find existing defects with this workType and compute next number
    const existingNums = reportDefects
      .filter((d) => {
        const parts = d.uid.split("-");
        return parts.some((p) => p === workType);
      })
      .map((d) => {
        const parts = d.uid.split("-");
        const lastPart = parts[parts.length - 1];
        return parseInt(lastPart, 10);
      })
      .filter((n) => !isNaN(n));
    const next = existingNums.length > 0 ? Math.max(...existingNums) + 1 : 1;
    setSeqNumber(String(next).padStart(2, "0"));
  }, [workType, reportDefects, isEdit]);

  useEffect(() => {
    // Only seed the form from the server response on the FIRST load. Any
    // subsequent refetch of `existingDefect` (e.g. React Query refetch-on-focus
    // after picking a photo on mobile) must not clobber the user's in-progress
    // edits to fields like Drop / Elevation / Level / Comment / Action.
    // Wait until project + globalSettings have loaded before parsing the UID
    if (existingDefect && project && globalSettings && !hasInitializedRef.current) {
      hasInitializedRef.current = true;
      setForm({
        dateOpened: existingDefect.dateOpened,
        dateClosed: existingDefect.dateClosed || "",
        comment: existingDefect.comment,
        actionRequired: existingDefect.actionRequired,
        assignedTo: existingDefect.assignedTo,
        dueDate: existingDefect.dueDate,
        verificationMethod: existingDefect.verificationMethod,
        verificationPerson: existingDefect.verificationPerson,
        status: existingDefect.status,
      });
      setUid(existingDefect.uid);
      if ((existingDefect as any).recordType) {
        setRecordType((existingDefect as any).recordType);
      }
      // Prefer the structured UID columns (source of truth) — no fragile re-parsing.
      // A record counts as "structured" if ANY of the part columns is populated; this is
      // true for all records after the uid_parts_backfill_v1 migration.
      const ed = existingDefect as any;
      const hasStructured = ed.elevationCode != null || ed.dropCode != null ||
        ed.levelCode != null || ed.workTypeCode != null || ed.seqNumber != null;
      if (hasStructured) {
        setElevation(ed.elevationCode || "");
        setDrop(ed.dropCode || "");
        setLevel(ed.levelCode || "");
        setWorkType(ed.workTypeCode || "");
        setSeqNumber(ed.seqNumber || "");
        setFormStep("form");
        return;
      }
      // Fallback: parse variable-length UID into components using KNOWN lists from the project
      // UID format: [Elevation]-[Drop]-[Level]-[WorkType]-[Number]
      // Any segment may be omitted. Number is always the LAST segment.
      // WorkType is always the SECOND-TO-LAST segment if it matches a known work type code.
      const parts = existingDefect.uid.split("-");
      if (parts.length > 0) {
        // Build known elevation codes from project + report elevations (use the same codeMap)
        const codeMap: Record<string, string> = {
          "North": "N", "South": "S", "East": "E", "West": "W",
          "North East": "NE", "North West": "NW", "South East": "SE", "South West": "SW",
        };
        const knownElevCodes = new Set<string>();
        for (const src of [(report as any)?.elevations, (project as any)?.elevations]) {
          if (!src) continue;
          try {
            const arr: string[] = JSON.parse(src as string);
            arr.forEach((label) => {
              const code = codeMap[label] || label.substring(0, 3).toUpperCase();
              knownElevCodes.add(code);
              knownElevCodes.add(label); // also accept the full label in the UID
            });
          } catch {}
        }
        // Build known work type codes from project customWorkTypes + global defaults
        const knownWtCodes = new Set<string>();
        try {
          const customWT = JSON.parse((project as any)?.customWorkTypes || "[]");
          customWT.forEach((wt: any) => knownWtCodes.add(wt.code));
        } catch {}
        try {
          if (globalSettings?.workTypes) {
            globalSettings.workTypes.forEach((wt: any) => knownWtCodes.add(wt.code));
          }
        } catch {}
        // Always-known fallback work types (do NOT include codes that could be elevations like LEV)
        ["CR","CK","PT","WP","GL","CL","ST","SE","FL","RR","GK","OT","SR","SM","BR","LR","SS","GW"].forEach((c) => knownWtCodes.add(c));

        // Last part is always the Number
        const lastIdx = parts.length - 1;
        setSeqNumber(parts[lastIdx]);

        // Search for work type — prefer the second-to-last position, else any match
        let wtIdx = -1;
        if (lastIdx >= 1 && knownWtCodes.has(parts[lastIdx - 1])) {
          wtIdx = lastIdx - 1;
        } else {
          // Try other positions but only match known codes (not just any alpha)
          for (let i = lastIdx - 1; i >= 0; i--) {
            if (knownWtCodes.has(parts[i])) { wtIdx = i; break; }
          }
        }

        if (wtIdx >= 0) {
          setWorkType(parts[wtIdx]);
          const before = parts.slice(0, wtIdx);
          // Now identify Elevation in `before` using knownElevCodes
          let elevIdx = -1;
          for (let i = 0; i < before.length; i++) {
            if (knownElevCodes.has(before[i])) { elevIdx = i; break; }
          }
          if (elevIdx >= 0) {
            setElevation(before[elevIdx]);
            before.splice(elevIdx, 1);
          }
          // Remaining parts are Drop then Level (in that order)
          if (before.length >= 1) setDrop(before[0]);
          if (before.length >= 2) setLevel(before[1]);
        }
      }
      setFormStep("form");
    }
  }, [existingDefect, project, globalSettings]);

  // When navigating between defects within the same form route, reset the
  // initialization flag so the next defect's data does seed its form.
  useEffect(() => {
    hasInitializedRef.current = false;
  }, [defectId]);

  useEffect(() => {
    if (existingPhotos) setPhotos(existingPhotos);
  }, [existingPhotos]);

  // Helper to get photo for a specific slot
  const getPhotoForSlot = (slot: SlotKey): Photo | undefined => {
    return photos.find((p) => p.slot === slot);
  };

  const saveMutation = useMutation({
    mutationFn: async () => {
      // Structured UID parts — stored as the source of truth so reopening never re-parses the uid string.
      const uidParts = {
        elevationCode: elevation || null,
        dropCode: drop || null,
        levelCode: level || null,
        workTypeCode: workType || null,
        seqNumber: seqNumber || null,
      };
      if (isEdit) {
        // Include updated UID if the fields were changed (open items only)
        const updatedUid = assembledUid || uid;
        const res = await apiRequest("PATCH", `/api/defects/${defectId}`, { ...form, uid: updatedUid, ...uidParts });
        return res.json();
      } else {
        const res = await apiRequest("POST", `/api/projects/${projectId}/defects`, {
          ...form,
          reportId: Number(reportId),
          uidPrefix,
          uidOverride: assembledUid,
          recordType,
          ...uidParts,
        });
        return res.json();
      }
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/defects`] });
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/defects`] });
      // Invalidate the singular defect query so the form's source of truth stays consistent
      if (isEdit && defectId) {
        queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
      }
      const typeLabel = recordType === "observation" ? "Observation" : "Defect";
      toast({ title: isEdit ? `${typeLabel} updated` : `${typeLabel} created` });
      if (!isEdit) {
        navigate(`/projects/${projectId}/reports/${reportId}/defects/${data.id}`, { replace: true });
      }
    },
    onError: (err: Error) => {
      toast({ title: err.message || "Failed to save", variant: "destructive" });
    },
  });

  // Toggle a record's classification between Defect and Observation on an existing record.
  const recordTypeMutation = useMutation({
    mutationFn: async (newType: "defect" | "observation") => {
      const res = await apiRequest("PATCH", `/api/defects/${defectId}`, { recordType: newType });
      return res.json();
    },
    onSuccess: (_data, newType) => {
      // Invalidate everything that buckets a record by type: report list (DEFECTS vs OBSERVATIONS),
      // project-wide list, the singular defect, and the assembled report export data.
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/defects`] });
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/defects`] });
      queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/report-data`] });
      toast({ title: `Changed to ${newType === "observation" ? "Observation" : "Defect"}` });
    },
    onError: () => {
      toast({ title: "Failed to change record type", variant: "destructive" });
      // Re-seed from server truth on failure
      queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
    },
  });

  const obsNoteMutation = useMutation({
    mutationFn: async (text: string) => {
      const res = await apiRequest("POST", `/api/defects/${defectId}/observation-note`, { text, reportId: Number(reportId) });
      return res.json();
    },
    onSuccess: (data) => {
      setForm((prev) => ({ ...prev, comment: data.comment }));
      setObsNoteOpen(false);
      setObsNoteText("");
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/observation-history`] });
      queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/report-data`] });
      toast({ title: "Inspection note added" });
    },
  });

  const actNoteMutation = useMutation({
    mutationFn: async (text: string) => {
      const res = await apiRequest("POST", `/api/defects/${defectId}/action-note`, { text, reportId: Number(reportId) });
      return res.json();
    },
    onSuccess: (data) => {
      setForm((prev) => ({ ...prev, actionRequired: data.actionRequired }));
      setActNoteOpen(false);
      setActNoteText("");
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/action-history`] });
      queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/report-data`] });
      toast({ title: "Action update added" });
    },
  });

  // "Add note for this inspection" — inserts an inspectionNotes row (does NOT overwrite
  // the canonical observation/action). Flips the "Amended this inspection" badge.
  const inspectionNoteMutation = useMutation({
    mutationFn: async (text: string) => {
      const res = await apiRequest("POST", `/api/defects/${defectId}/notes`, { text, reportId: Number(reportId) });
      return res.json();
    },
    onSuccess: () => {
      setInspNoteOpen(false);
      setInspNoteText("");
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/history`, reportId] });
      queryClient.invalidateQueries({ queryKey: ["/api/defects", defectId] });
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/report-data`] });
      toast({ title: "Note added for this inspection" });
    },
  });

  const handlePhotoUpload = async (file: File, slot: SlotKey) => {
    if (!defectId) {
      toast({ title: "Save the defect first before adding photos", variant: "destructive" });
      return;
    }
    setUploading(slot);
    try {
      const formData = new FormData();
      formData.append("photo", file);
      formData.append("slot", slot);
      if (reportId) formData.append("reportId", reportId);

      const res = await fetch(`${API_BASE}/api/defects/${defectId}/photos`, {
        method: "POST",
        body: formData,
      });

      if (!res.ok) throw new Error("Upload failed");
      const photo = await res.json();
      // Replace existing photo in this slot, or add new
      setPhotos((prev) => {
        const filtered = prev.filter((p) => p.slot !== slot);
        return [...filtered, photo];
      });
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/photos`] });
      toast({ title: `Photo added to ${slot === "complete" ? "Complete" : slot.toUpperCase()}` });

      // Auto-mark as complete when uploading to Complete slot
      if (slot === "complete" && defectId) {
        const dateClosed = new Date().toISOString().split("T")[0];
        setForm((prev) => ({ ...prev, status: "complete", dateClosed }));
        try {
          await apiRequest("PATCH", `/api/defects/${defectId}`, { status: "complete", dateClosed });
          queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/defects`] });
          queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/defects`] });
          toast({ title: `${recordType === "observation" ? "Observation" : "Defect"} marked as complete` });
        } catch {
          toast({ title: "Photo uploaded but failed to auto-complete", variant: "destructive" });
        }
      }
    } catch {
      toast({ title: "Failed to upload photo", variant: "destructive" });
    } finally {
      setUploading(null);
    }
  };

  // Caption state - keyed by photo id
  const [captions, setCaptions] = useState<Record<number, string>>({});

  // Initialize captions from loaded photos
  useEffect(() => {
    if (photos.length > 0) {
      const caps: Record<number, string> = {};
      photos.forEach((p) => { caps[p.id] = p.caption || ""; });
      setCaptions(caps);
    }
  }, [photos]);

  const handleCaptionSave = async (photoId: number, caption: string) => {
    try {
      await apiRequest("PATCH", `/api/photos/${photoId}`, { caption });
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/photos`] });
    } catch {
      toast({ title: "Failed to save caption", variant: "destructive" });
    }
  };

  // Debounced save for caption edits
  const captionTimers = useRef<Record<number, NodeJS.Timeout>>({});
  const updateCaption = useCallback((photoId: number, text: string) => {
    setCaptions((prev) => ({ ...prev, [photoId]: text }));
    if (captionTimers.current[photoId]) clearTimeout(captionTimers.current[photoId]);
    captionTimers.current[photoId] = setTimeout(() => handleCaptionSave(photoId, text), 800);
  }, [defectId]);

  // Compute whether a photo is "new this inspection" (respects newOverride)
  // Uses originReportId (first inspection photo appeared in) for auto-detect
  const isPhotoNew = (photo: Photo): boolean => {
    if (photo.newOverride === "new") return true;
    if (photo.newOverride === "not-new") return false;
    return (photo.originReportId ?? photo.reportId) === Number(reportId);
  };

  const handleToggleNewOverride = async (photo: Photo) => {
    const currentlyNew = isPhotoNew(photo);
    // Toggle: if currently new → mark "not-new"; if currently not new → mark "new"
    // If there's already an override matching the auto-detect result, clear it instead
    const newVal = currentlyNew ? "not-new" : "new";
    try {
      await apiRequest("PATCH", `/api/photos/${photo.id}`, { newOverride: newVal });
      // Update local state
      setPhotos(prev => prev.map(p =>
        p.id === photo.id ? { ...p, newOverride: newVal } : p
      ));
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/photos`] });
      toast({ title: newVal === "new" ? "Marked as new this inspection" : "Marked as not new" });
    } catch {
      toast({ title: "Failed to update photo", variant: "destructive" });
    }
  };

  const handleDeletePhoto = async (photoId: number) => {
    try {
      await apiRequest("DELETE", `/api/photos/${photoId}`);
      setPhotos((prev) => prev.filter((p) => p.id !== photoId));
      queryClient.invalidateQueries({ queryKey: [`/api/defects/${defectId}/photos`] });
      toast({ title: "Photo removed" });
    } catch {
      toast({ title: "Failed to remove photo", variant: "destructive" });
    }
  };

  const handleMarkComplete = () => {
    setForm((prev) => ({
      ...prev,
      status: "complete",
      dateClosed: new Date().toISOString().split("T")[0],
    }));
  };

  const handleReopen = () => {
    setForm((prev) => ({
      ...prev,
      status: "open",
      dateClosed: "",
    }));
  };

  const set = (field: string) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    setForm((prev) => ({ ...prev, [field]: e.target.value }));
    if (field === "comment" || field === "actionRequired") triggerAutosave();
  };

  // Helpers to add custom values to project lists
  const addCustomElevation = async () => {
    const val = window.prompt("Custom elevation label (e.g. Podium, Roof):");
    if (!val || !val.trim()) return;
    const trimmed = val.trim();
    try {
      const elevs: string[] = (() => { try { return JSON.parse((project as any)?.elevations || "[]"); } catch { return []; } })();
      if (!elevs.includes(trimmed)) { elevs.push(trimmed); }
      await apiRequest("PATCH", `/api/projects/${projectId}`, { elevations: JSON.stringify(elevs) });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      const codeMap: Record<string, string> = {
        "North": "N", "South": "S", "East": "E", "West": "W",
        "North East": "NE", "North West": "NW", "South East": "SE", "South West": "SW",
      };
      setElevation(codeMap[trimmed] || trimmed.substring(0, 3).toUpperCase());
    } catch { toast({ title: "Failed to add custom elevation", variant: "destructive" }); }
  };

  const addCustomDrop = async () => {
    const val = window.prompt("Custom drop value (e.g. Podium, Roof, 15):");
    if (!val || !val.trim()) return;
    const trimmed = val.trim();
    try {
      const drops: string[] = (() => { try { return JSON.parse((project as any)?.customDrops || "[]"); } catch { return []; } })();
      if (!drops.includes(trimmed)) { drops.push(trimmed); }
      await apiRequest("PATCH", `/api/projects/${projectId}`, { customDrops: JSON.stringify(drops) });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      setDrop(trimmed);
    } catch { toast({ title: "Failed to add custom drop", variant: "destructive" }); }
  };

  const addCustomLevel = async () => {
    const val = window.prompt("Custom level value (e.g. Basement, Roof, 25):");
    if (!val || !val.trim()) return;
    const trimmed = val.trim();
    try {
      const levels: string[] = (() => { try { return JSON.parse((project as any)?.customLevels || "[]"); } catch { return []; } })();
      if (!levels.includes(trimmed)) { levels.push(trimmed); }
      await apiRequest("PATCH", `/api/projects/${projectId}`, { customLevels: JSON.stringify(levels) });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      setLevel(trimmed);
    } catch { toast({ title: "Failed to add custom level", variant: "destructive" }); }
  };

  const canSave = isEdit || !!assembledUid;
  const isComplete = form.status === "complete";

  return (
    <div className="max-w-2xl mx-auto px-4 py-6 pb-24">
      {/* Header */}
      <Link href={`/projects/${projectId}/reports/${reportId}`}>
        <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4" data-testid="button-back-to-project">
          <ArrowLeft className="w-4 h-4" />
          Back to Report
        </button>
      </Link>

      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <h1 className="text-xl font-semibold tracking-tight">
            {isEdit ? `Edit ${recordType === "observation" ? "Observation" : "Defect"}` : `New ${recordType === "observation" ? "Observation" : "Defect"}`}
          </h1>
          {isEdit && (
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
                <><Clock className="w-3 h-3 mr-1" />Open</>
              )}
            </Badge>
          )}
          {isEdit && (
            <div className="flex items-center gap-1.5">
              <span className="text-[10px] uppercase tracking-wide text-muted-foreground">Record type</span>
              <div className="inline-flex rounded-md border overflow-hidden" data-testid="record-type-toggle">
                {(["defect", "observation"] as const).map((type) => (
                  <button
                    key={type}
                    type="button"
                    disabled={recordTypeMutation.isPending}
                    onClick={() => {
                      if (recordType === type) return;
                      setRecordType(type);
                      recordTypeMutation.mutate(type, {
                        onError: () => setRecordType(recordType),
                      });
                    }}
                    className={`px-2.5 py-1 text-xs font-medium transition-colors disabled:opacity-50 ${
                      recordType === type
                        ? "bg-primary text-primary-foreground"
                        : "bg-background text-muted-foreground hover:bg-accent/60"
                    }`}
                    data-testid={`button-record-type-${type}`}
                  >
                    {type === "defect" ? "Defect" : "Observation"}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
        {isEdit && (
          <div>
            {!isComplete ? (
              <Button variant="secondary" onClick={handleMarkComplete} data-testid="button-mark-complete">
                <CheckCircle2 className="w-4 h-4 mr-2" />
                Mark Complete
              </Button>
            ) : (
              <Button variant="secondary" onClick={handleReopen} data-testid="button-reopen-defect">
                <Clock className="w-4 h-4 mr-2" />
                Reopen
              </Button>
            )}
          </div>
        )}
      </div>

      {/* Mark on Elevation — only for existing defects */}
      {isEdit && (
        <div className="mb-4">
          <Button
            type="button"
            variant="outline"
            size="sm"
            className="gap-2"
            onClick={handleMarkOnElevation}
            data-testid="button-mark-on-elevation"
          >
            <MapPin className="w-4 h-4" />
            Mark on Elevation
          </Button>
        </div>
      )}

      {/* Elevation picker dialog (when multiple elevations exist) */}
      <Dialog open={elevationPickerOpen} onOpenChange={(v) => { setElevationPickerOpen(v); if (!v) setLocationPickerTarget(null); }}>
        <DialogContent className="max-w-sm">
          <DialogHeader>
            <DialogTitle>Choose Elevation</DialogTitle>
          </DialogHeader>
          <div className="space-y-2">
            {projectElevations.map((el) => (
              <Button
                key={el.id}
                variant="outline"
                className="w-full justify-start"
                onClick={() => {
                  setElevationPickerOpen(false);
                  if (locationPickerTarget) {
                    const locUid = locationPickerTarget.uid || computeLocationUid(locationPickerTarget.elevation || "", locationPickerTarget.drop || "", locationPickerTarget.level || "");
                    navigate(`/projects/${projectId}/elevations/${el.id}?defect=${encodeURIComponent(locUid)}&locationId=${locationPickerTarget.id}`);
                    setLocationPickerTarget(null);
                  } else {
                    const defectUid = uid || assembledUid;
                    navigate(`/projects/${projectId}/elevations/${el.id}?defect=${encodeURIComponent(defectUid)}`);
                  }
                }}
              >
                {el.name}
              </Button>
            ))}
          </div>
        </DialogContent>
      </Dialog>

      <form
        onSubmit={async (e) => {
          e.preventDefault();
          if (!canSave) {
            toast({ title: "Select a Work Type to generate the entry ID", variant: "destructive" });
            return;
          }
          // Cancel any pending autosave — the explicit save covers everything
          if (autosaveTimer.current) { clearTimeout(autosaveTimer.current); autosaveTimer.current = null; }
          // Flush pending location PATCHes so they complete before the parent save
          await flushPendingLocationPatches();
          saveMutation.mutate();
        }}
        className="space-y-6"
      >
        {/* Step 1: Work Type Picker (shown for new entries) */}
        {formStep === "workType" && (
          <Card className="p-5 space-y-4">
            <h3 className="text-sm font-medium">Select Work Type</h3>
            {(() => {
              const globalWts = globalSettings?.workTypes || [];
              const primaryWts = primaryWorkTypeCodes.length > 0
                ? primaryWorkTypeCodes.map((code) => globalWts.find((wt) => wt.code === code) || { code, label: code }).filter(Boolean)
                : allWorkTypes;
              const otherWts = globalWts.filter((wt) => !primaryWorkTypeCodes.includes(wt.code));
              const customProjectWts: { code: string; label: string }[] = (() => {
                try { return JSON.parse((project as any)?.customWorkTypes || "[]"); } catch { return []; }
              })();

              return (
                <>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                    {primaryWts.map((wt) => (
                      <button
                        key={wt.code}
                        type="button"
                        onClick={() => { setWorkType(wt.code); setShowOtherWorkTypes(false); setFormStep("form"); }}
                        className="flex flex-col items-start p-3 rounded-lg border hover:bg-accent/60 transition-colors text-left"
                      >
                        <span className="font-mono font-semibold text-sm">{wt.code}</span>
                        <span className="text-xs text-muted-foreground">{wt.label}</span>
                      </button>
                    ))}
                    {customProjectWts.map((wt) => (
                      <button
                        key={wt.code}
                        type="button"
                        onClick={() => { setWorkType(wt.code); setShowOtherWorkTypes(false); setFormStep("form"); }}
                        className="flex flex-col items-start p-3 rounded-lg border hover:bg-accent/60 transition-colors text-left"
                      >
                        <span className="font-mono font-semibold text-sm">{wt.code}</span>
                        <span className="text-xs text-muted-foreground">{wt.label}</span>
                      </button>
                    ))}
                  </div>

                  {primaryWorkTypeCodes.length > 0 && otherWts.length > 0 && (
                    <>
                      <button
                        type="button"
                        onClick={() => setShowOtherWorkTypes(!showOtherWorkTypes)}
                        className="text-xs text-muted-foreground hover:text-foreground underline underline-offset-2"
                      >
                        {showOtherWorkTypes ? "Hide other work types" : `Other (${otherWts.length} more)`}
                      </button>
                      {showOtherWorkTypes && (
                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                          {otherWts.map((wt) => (
                            <button
                              key={wt.code}
                              type="button"
                              onClick={() => {
                                setWorkType(wt.code);
                                setShowOtherWorkTypes(false);
                                setFormStep("form");
                                if (window.confirm(`Add "${wt.code} — ${wt.label}" to this project's primary work types?`)) {
                                  const updated = [...primaryWorkTypeCodes, wt.code];
                                  apiRequest("PATCH", `/api/projects/${projectId}`, { primaryWorkTypes: JSON.stringify(updated) })
                                    .then(() => queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] }));
                                }
                              }}
                              className="flex flex-col items-start p-3 rounded-lg border hover:bg-accent/60 transition-colors text-left opacity-80"
                            >
                              <span className="font-mono font-semibold text-sm">{wt.code}</span>
                              <span className="text-xs text-muted-foreground">{wt.label}</span>
                            </button>
                          ))}
                        </div>
                      )}
                    </>
                  )}

                  <button
                    type="button"
                    onClick={async () => {
                      const code = window.prompt("Work type code (2-3 characters):");
                      if (!code || code.trim().length < 2) return;
                      const trimCode = code.trim().toUpperCase().slice(0, 3);
                      const validationError = validateCustomWorkTypeCode(trimCode);
                      if (validationError) {
                        toast({ title: validationError, variant: "destructive" });
                        return;
                      }
                      const label = window.prompt("Work type label:");
                      if (!label || !label.trim()) return;
                      const trimLabel = label.trim();
                      try {
                        const existing: { code: string; label: string }[] = (() => {
                          try { return JSON.parse((project as any)?.customWorkTypes || "[]"); } catch { return []; }
                        })();
                        existing.push({ code: trimCode, label: trimLabel });
                        await apiRequest("PATCH", `/api/projects/${projectId}`, { customWorkTypes: JSON.stringify(existing) });
                        queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
                        if (window.confirm(`Also add "${trimCode} — ${trimLabel}" to the global work types list?`)) {
                          await apiRequest("POST", "/api/global-settings/work-types", { code: trimCode, label: trimLabel });
                          queryClient.invalidateQueries({ queryKey: ["/api/global-settings"] });
                        }
                        setWorkType(trimCode);
                        setShowOtherWorkTypes(false);
                        setFormStep("form");
                      } catch {
                        toast({ title: "Failed to add custom work type", variant: "destructive" });
                      }
                    }}
                    className="flex items-center gap-2 p-3 rounded-lg border border-dashed hover:bg-accent/40 transition-colors w-full text-left"
                  >
                    <Plus className="w-4 h-4 text-muted-foreground" />
                    <span className="text-xs text-muted-foreground">Add Custom Work Type</span>
                  </button>
                </>
              );
            })()}
          </Card>
        )}

        {/* Step 2: UID builder + full form */}
        {formStep === "form" && (<>
        <Card className="p-4 space-y-4 bg-accent/30">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-medium">Defect ID</h3>
            {(assembledUid || uid) && (
              <span className="font-mono text-sm font-semibold text-primary" data-testid="text-defect-uid">
                {isEdit ? uid : assembledUid}
              </span>
            )}
          </div>
          <div className="grid grid-cols-3 sm:grid-cols-5 gap-3">
            {enabledUidParts.elevation !== false && (
            <div>
              <Label htmlFor="elevation" className="text-xs">Elevation</Label>
              <Select
                value={elevation}
                onValueChange={setElevation}
              >
                <SelectTrigger className="font-mono" data-testid="select-elevation">
                  <SelectValue placeholder="—" />
                </SelectTrigger>
                <SelectContent>
                  {elevationOptions.map((el) => (
                    <SelectItem key={el.code} value={el.code}>
                      {el.code} — {el.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <div className="flex items-center justify-between mt-0.5">
                <button type="button" onClick={addCustomElevation} className="text-[10px] text-primary hover:underline">+ Add Custom</button>
                {elevation && (
                  <button type="button" onClick={() => setElevation("")} className="text-[10px] text-muted-foreground hover:text-foreground" data-testid="button-clear-elevation">Clear</button>
                )}
              </div>
            </div>
            )}
            {enabledUidParts.drop !== false && (
            <div>
              <Label htmlFor="drop" className="text-xs">Drop</Label>
              <Select value={drop} onValueChange={setDrop}>
                <SelectTrigger className="font-mono" data-testid="input-drop">
                  <SelectValue placeholder="—" />
                </SelectTrigger>
                <SelectContent>
                  {dropOptions.map((d) => (
                    <SelectItem key={d} value={d}>{d}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <div className="flex items-center justify-between mt-0.5">
                <button type="button" onClick={addCustomDrop} className="text-[10px] text-primary hover:underline">+ Add Custom</button>
                {drop && (
                  <button type="button" onClick={() => setDrop("")} className="text-[10px] text-muted-foreground hover:text-foreground" data-testid="button-clear-drop">Clear</button>
                )}
              </div>
            </div>
            )}
            <div>
              <Label htmlFor="level" className="text-xs">Level</Label>
              <Select value={level} onValueChange={setLevel}>
                <SelectTrigger className="font-mono" data-testid="input-level">
                  <SelectValue placeholder="—" />
                </SelectTrigger>
                <SelectContent>
                  {levelOptions.map((l) => (
                    <SelectItem key={l} value={l}>{l}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <div className="flex items-center justify-between mt-0.5">
                <button type="button" onClick={addCustomLevel} className="text-[10px] text-primary hover:underline">+ Add Custom</button>
                {level && (
                  <button type="button" onClick={() => setLevel("")} className="text-[10px] text-muted-foreground hover:text-foreground" data-testid="button-clear-level">Clear</button>
                )}
              </div>
            </div>
            {enabledUidParts.workType !== false && (
            <div>
              <Label htmlFor="workType" className="text-xs">Work Type</Label>
              <button
                type="button"
                onClick={() => setFormStep("workType")}
                className="w-full h-9 px-3 rounded-md border bg-background text-left font-mono text-sm hover:bg-accent/40 transition-colors"
                data-testid="select-work-type"
              >
                {workType || "—"}
              </button>
              {workType && (
                <button type="button" onClick={() => setWorkType("")} className="text-[10px] text-muted-foreground hover:text-foreground mt-0.5 block" data-testid="button-clear-work-type">Clear</button>
              )}
            </div>
            )}
            <div>
              <Label htmlFor="seqNum" className="text-xs">Number</Label>
              <Input
                id="seqNum"
                placeholder="01"
                value={seqNumber}
                onChange={(e) => {
                  const v = e.target.value.replace(/\D/g, "").slice(0, 2);
                  setSeqNumber(v);
                }}
                maxLength={2}
                className="font-mono text-center"
                data-testid="input-seq-number"
              />
            </div>
          </div>
          <p className="text-xs text-muted-foreground">
            All fields optional — empty segments are skipped. The number auto-suggests but you can change it.
          </p>
        </Card>

        {/* Dates */}
        <div className="grid grid-cols-2 gap-4">
          <div>
            <Label htmlFor="dateOpened">Date Opened</Label>
            <Input
              id="dateOpened"
              type="date"
              value={form.dateOpened}
              onChange={set("dateOpened")}
              required
              data-testid="input-date-opened"
            />
          </div>
          <div>
            <Label htmlFor="dateClosed">Date Completed</Label>
            <Input
              id="dateClosed"
              type="date"
              value={form.dateClosed}
              onChange={set("dateClosed")}
              data-testid="input-date-closed"
            />
          </div>
        </div>

        {/* Comment */}
        <div>
          <div className="flex items-center gap-2 mb-1">
            <Label htmlFor="comment">Observation</Label>
            <DictationButton
              onTranscript={(text) => { setForm((prev) => ({ ...prev, comment: prev.comment + (prev.comment ? " " : "") + text })); triggerAutosave(); }}
            />
            {isEdit && autosaveStatus === "saving" && (
              <span className="text-[10px] text-muted-foreground animate-pulse">Saving...</span>
            )}
            {isEdit && autosaveStatus === "saved" && (
              <span className="text-[10px] text-green-600">Saved</span>
            )}
          </div>
          {textSuggestions.comments.length > 0 && !form.comment && (
            <div className="mb-1.5">
              <p className="text-[10px] text-muted-foreground mb-1">Previously used:</p>
              <div className="flex flex-wrap gap-1">
                {textSuggestions.comments.map((s, i) => (
                  <button
                    key={i}
                    type="button"
                    onClick={() => { setForm((prev) => ({ ...prev, comment: s })); triggerAutosave(); }}
                    className="text-xs px-2 py-1 rounded-md bg-accent hover:bg-accent/80 text-left truncate max-w-full border"
                    title={s}
                  >
                    {s.length > 60 ? s.substring(0, 57) + "..." : s}
                  </button>
                ))}
              </div>
            </div>
          )}
          <Textarea
            id="comment"
            ref={commentRef}
            placeholder="Describe the defect observed..."
            value={form.comment}
            onChange={set("comment")}
            rows={3}
            required
            data-testid="input-comment"
          />
          {isEdit && obsHistory && obsHistory.length > 0 && (
            <Collapsible open={obsHistoryOpen} onOpenChange={setObsHistoryOpen} className="mt-2">
              <CollapsibleTrigger className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground transition-colors">
                <ChevronDown className={`w-3.5 h-3.5 transition-transform ${obsHistoryOpen ? "rotate-0" : "-rotate-90"}`} />
                <History className="w-3.5 h-3.5" />
                History ({obsHistory.length} {obsHistory.length === 1 ? "entry" : "entries"})
              </CollapsibleTrigger>
              <CollapsibleContent className="mt-2 space-y-2">
                {obsHistory.map((entry) => (
                  <div key={entry.id} className="rounded-md border bg-muted/40 p-2.5">
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-xs font-medium text-muted-foreground">
                        {entry.reportName} — {formatHistoryDate(entry.reportDate)}
                      </span>
                      <div className="flex gap-1">
                        <Button
                          type="button"
                          variant="secondary"
                          size="sm"
                          onClick={() => { setForm((prev) => ({ ...prev, comment: entry.text })); triggerAutosave(); }}
                          title="Copy this text as the current observation"
                        >
                          <Copy className="w-3 h-3" /> Still applicable
                        </Button>
                        <Button
                          type="button"
                          variant="secondary"
                          size="sm"
                          onClick={() => {
                            setForm((prev) => ({ ...prev, comment: "" }));
                            commentRef.current?.focus();
                          }}
                          title="Clear the field and write a new observation"
                        >
                          <PenLine className="w-3 h-3" /> New observation
                        </Button>
                      </div>
                    </div>
                    <p className="text-xs text-muted-foreground whitespace-pre-wrap">{entry.text}</p>
                  </div>
                ))}
              </CollapsibleContent>
            </Collapsible>
          )}
          {isEdit && (
            <div className="mt-2">
              {!obsNoteOpen ? (
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => setObsNoteOpen(true)}
                  className="text-xs"
                >
                  <Plus className="w-3 h-3 mr-1" />
                  Add inspection note
                </Button>
              ) : (
                <div className="rounded-md border border-emerald-300 bg-emerald-50/50 dark:bg-emerald-900/10 p-2.5 space-y-2">
                  <p className="text-xs font-medium text-emerald-800 dark:text-emerald-400">New observation entry</p>
                  <Textarea
                    placeholder="Add a new dated observation for this inspection..."
                    value={obsNoteText}
                    onChange={(e) => setObsNoteText(e.target.value)}
                    rows={2}
                    autoFocus
                  />
                  <div className="flex gap-2 justify-end">
                    <Button type="button" variant="ghost" size="sm" onClick={() => { setObsNoteOpen(false); setObsNoteText(""); }}>
                      Cancel
                    </Button>
                    <Button
                      type="button"
                      size="sm"
                      disabled={!obsNoteText.trim() || obsNoteMutation.isPending}
                      onClick={() => obsNoteMutation.mutate(obsNoteText.trim())}
                    >
                      {obsNoteMutation.isPending ? "Saving..." : "Save note"}
                    </Button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Action Required */}
        <div>
          <div className="flex items-center gap-2 mb-1">
            <Label htmlFor="actionRequired">Action Required</Label>
            <DictationButton
              onTranscript={(text) => { setForm((prev) => ({ ...prev, actionRequired: prev.actionRequired + (prev.actionRequired ? " " : "") + text })); triggerAutosave(); }}
            />
          </div>
          {textSuggestions.actions.length > 0 && !form.actionRequired && (
            <div className="mb-1.5">
              <p className="text-[10px] text-muted-foreground mb-1">Previously used:</p>
              <div className="flex flex-wrap gap-1">
                {textSuggestions.actions.map((s, i) => (
                  <button
                    key={i}
                    type="button"
                    onClick={() => { setForm((prev) => ({ ...prev, actionRequired: s })); triggerAutosave(); }}
                    className="text-xs px-2 py-1 rounded-md bg-accent hover:bg-accent/80 text-left truncate max-w-full border"
                    title={s}
                  >
                    {s.length > 60 ? s.substring(0, 57) + "..." : s}
                  </button>
                ))}
              </div>
            </div>
          )}
          <Textarea
            id="actionRequired"
            ref={actionRef}
            placeholder="What needs to be done to rectify this defect?"
            value={form.actionRequired}
            onChange={set("actionRequired")}
            rows={2}
            required
            data-testid="input-action-required"
          />
          {isEdit && actHistory && actHistory.length > 0 && (
            <Collapsible open={actHistoryOpen} onOpenChange={setActHistoryOpen} className="mt-2">
              <CollapsibleTrigger className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground transition-colors">
                <ChevronDown className={`w-3.5 h-3.5 transition-transform ${actHistoryOpen ? "rotate-0" : "-rotate-90"}`} />
                <History className="w-3.5 h-3.5" />
                History ({actHistory.length} {actHistory.length === 1 ? "entry" : "entries"})
              </CollapsibleTrigger>
              <CollapsibleContent className="mt-2 space-y-2">
                {actHistory.map((entry) => (
                  <div key={entry.id} className="rounded-md border bg-muted/40 p-2.5">
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-xs font-medium text-muted-foreground">
                        {entry.reportName} — {formatHistoryDate(entry.reportDate)}
                      </span>
                      <div className="flex gap-1">
                        <Button
                          type="button"
                          variant="secondary"
                          size="sm"
                          onClick={() => { setForm((prev) => ({ ...prev, actionRequired: entry.text })); triggerAutosave(); }}
                          title="Copy this text as the current action"
                        >
                          <Copy className="w-3 h-3" /> Still applicable
                        </Button>
                        <Button
                          type="button"
                          variant="secondary"
                          size="sm"
                          onClick={() => {
                            setForm((prev) => ({ ...prev, actionRequired: "" }));
                            actionRef.current?.focus();
                          }}
                          title="Clear the field and write a new action"
                        >
                          <PenLine className="w-3 h-3" /> Add new
                        </Button>
                      </div>
                    </div>
                    <p className="text-xs text-muted-foreground whitespace-pre-wrap">{entry.text}</p>
                  </div>
                ))}
              </CollapsibleContent>
            </Collapsible>
          )}
          {isEdit && (
            <div className="mt-2">
              {!actNoteOpen ? (
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => setActNoteOpen(true)}
                  className="text-xs"
                >
                  <Plus className="w-3 h-3 mr-1" />
                  Add action update
                </Button>
              ) : (
                <div className="rounded-md border border-blue-300 bg-blue-50/50 dark:bg-blue-900/10 p-2.5 space-y-2">
                  <p className="text-xs font-medium text-blue-800 dark:text-blue-400">New action entry</p>
                  <Textarea
                    placeholder="Add an updated action for this inspection..."
                    value={actNoteText}
                    onChange={(e) => setActNoteText(e.target.value)}
                    rows={2}
                    autoFocus
                  />
                  <div className="flex gap-2 justify-end">
                    <Button type="button" variant="ghost" size="sm" onClick={() => { setActNoteOpen(false); setActNoteText(""); }}>
                      Cancel
                    </Button>
                    <Button
                      type="button"
                      size="sm"
                      disabled={!actNoteText.trim() || actNoteMutation.isPending}
                      onClick={() => actNoteMutation.mutate(actNoteText.trim())}
                    >
                      {actNoteMutation.isPending ? "Saving..." : "Save update"}
                    </Button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Inspection history — unified log (observation/action/note/status), newest-first.
            Current cluster at full strength; prior cluster collapsed and muted. */}
        {isEdit && (() => {
          const all = unifiedHistory ?? [];
          const current = all.filter((e) => e.age === "current");
          const prior = all.filter((e) => e.age === "prior");
          const kindLabel = (k: string) =>
            k === "observation" ? "Observation" : k === "action" ? "Action" : k === "note" ? "Note" : "Status";
          const renderEntry = (e: typeof all[number], idx: number, muted: boolean) => {
            const meta = `[Insp-${e.inspectionNumber || "?"}] ${formatHistoryDate(e.date)}${e.author ? ` \u00b7 ${e.author}` : ""} \u00b7 ${kindLabel(e.kind)}`;
            return (
              <div key={`${e.kind}-${e.reportId}-${idx}`} className={`rounded-md border p-2.5 ${muted ? "bg-muted/30 opacity-70" : "bg-emerald-50/40 dark:bg-emerald-900/10 border-emerald-200 dark:border-emerald-900"}`}>
                <div className="flex items-center justify-between mb-1 gap-2">
                  <span className="text-xs font-medium text-muted-foreground">{meta}</span>
                </div>
                <p className="text-xs text-muted-foreground whitespace-pre-wrap">{e.text}</p>
              </div>
            );
          };
          return (
            <div className="border-t pt-4">
              <div className="flex items-center gap-1.5 mb-2">
                <History className="w-4 h-4 text-muted-foreground" />
                <h3 className="text-sm font-semibold">Inspection history</h3>
              </div>

              {/* Add note for this inspection */}
              <div className="mb-3">
                {!inspNoteOpen ? (
                  <Button type="button" variant="outline" size="sm" onClick={() => setInspNoteOpen(true)} className="text-xs">
                    <Plus className="w-3 h-3 mr-1" />
                    Add note for this inspection
                  </Button>
                ) : (
                  <div className="rounded-md border border-emerald-300 bg-emerald-50/50 dark:bg-emerald-900/10 p-2.5 space-y-2">
                    <p className="text-xs font-medium text-emerald-800 dark:text-emerald-400">Add note for this inspection</p>
                    <Textarea
                      placeholder="Note about this defect for the current inspection..."
                      value={inspNoteText}
                      onChange={(e) => setInspNoteText(e.target.value)}
                      rows={2}
                      autoFocus
                    />
                    <div className="flex gap-2 justify-end">
                      <Button type="button" variant="ghost" size="sm" onClick={() => { setInspNoteOpen(false); setInspNoteText(""); }}>
                        Cancel
                      </Button>
                      <Button
                        type="button"
                        size="sm"
                        disabled={!inspNoteText.trim() || inspectionNoteMutation.isPending}
                        onClick={() => inspectionNoteMutation.mutate(inspNoteText.trim())}
                      >
                        {inspectionNoteMutation.isPending ? "Saving..." : "Save note"}
                      </Button>
                    </div>
                  </div>
                )}
              </div>

              {current.length > 0 && (
                <div className="space-y-2 mb-2">
                  {current.map((e, i) => renderEntry(e, i, false))}
                </div>
              )}
              {prior.length > 0 && (
                <Collapsible open={priorHistoryOpen} onOpenChange={setPriorHistoryOpen}>
                  <CollapsibleTrigger className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground transition-colors">
                    <ChevronDown className={`w-3.5 h-3.5 transition-transform ${priorHistoryOpen ? "rotate-0" : "-rotate-90"}`} />
                    Show prior history ({prior.length} {prior.length === 1 ? "entry" : "entries"})
                  </CollapsibleTrigger>
                  <CollapsibleContent className="mt-2 space-y-2">
                    {prior.map((e, i) => renderEntry(e, i, true))}
                  </CollapsibleContent>
                </Collapsible>
              )}
              {current.length === 0 && prior.length === 0 && (
                <p className="text-xs text-muted-foreground">No history yet for this defect.</p>
              )}
            </div>
          );
        })()}

        {/* By Whom / By When */}
        <div className="grid grid-cols-2 gap-4">
          <div>
            <Label htmlFor="assignedTo">By Whom</Label>
            <Select
              value={form.assignedTo}
              onValueChange={(val) => setForm((prev) => ({ ...prev, assignedTo: val }))}
            >
              <SelectTrigger data-testid="input-assigned-to">
                <SelectValue placeholder="Select..." />
              </SelectTrigger>
              <SelectContent>
                {PERSON_ROLES.map((role) => (
                  <SelectItem key={role} value={role}>{role}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div>
            <Label htmlFor="dueDate">By When</Label>
            <Input
              id="dueDate"
              type="date"
              value={form.dueDate}
              onChange={set("dueDate")}
              data-testid="input-due-date"
            />
          </div>
        </div>

        {/* Verification */}
        <Card className="p-4 space-y-4 bg-accent/30">
          <h3 className="text-sm font-medium">Verification to Close Out</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <Label htmlFor="verificationMethod">Method</Label>
              <Select
                value={form.verificationMethod}
                onValueChange={(val) => setForm((prev) => ({ ...prev, verificationMethod: val }))}
              >
                <SelectTrigger data-testid="select-verification-method">
                  <SelectValue placeholder="Select method" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="visual_inspection">Visual Inspection</SelectItem>
                  <SelectItem value="photographic_evidence">Photographic Evidence</SelectItem>
                  <SelectItem value="testing">Testing</SelectItem>
                  <SelectItem value="third_party_review">Third Party Review</SelectItem>
                  <SelectItem value="documentation_review">Documentation Review</SelectItem>
                  <SelectItem value="other">Other</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label htmlFor="verificationPerson">Person</Label>
              <Select
                value={form.verificationPerson}
                onValueChange={(val) => setForm((prev) => ({ ...prev, verificationPerson: val }))}
              >
                <SelectTrigger data-testid="input-verification-person">
                  <SelectValue placeholder="Select..." />
                </SelectTrigger>
                <SelectContent>
                  {PERSON_ROLES.map((role) => (
                    <SelectItem key={role} value={role}>{role}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </div>
        </Card>

        {/* Additional Locations */}
        {isEdit && (
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-medium">Additional Locations</h3>
              <Button
                type="button"
                variant="outline"
                size="sm"
                className="gap-1.5"
                onClick={addAdditionalLocation}
              >
                <Plus className="w-3.5 h-3.5" />
                Add Location
              </Button>
            </div>
            {additionalLocations.length === 0 && (
              <p className="text-xs text-muted-foreground bg-accent/40 p-3 rounded-lg">
                No additional locations yet. Add a location if this defect occurs at multiple places.
              </p>
            )}
            {additionalLocations.map((loc) => {
              const locUid = loc.uid || computeLocationUid(loc.elevation || "", loc.drop || "", loc.level || "");
              return (
                <Card key={loc.id} className="p-4 space-y-3 bg-accent/20">
                  <div className="flex items-center justify-between">
                    <span className="font-mono text-xs font-semibold text-primary">
                      {locUid || "—"}
                    </span>
                    <div className="flex gap-1.5">
                      <Button
                        type="button"
                        variant="ghost"
                        size="icon"
                        className="h-7 w-7"
                        onClick={() => handleMarkLocationOnElevation(loc)}
                        title="Mark on Elevation"
                      >
                        <MapPin className="w-3.5 h-3.5" />
                      </Button>
                      <Button
                        type="button"
                        variant="ghost"
                        size="icon"
                        className="h-7 w-7 hover:text-destructive"
                        onClick={() => deleteAdditionalLocation(loc.id)}
                        title="Delete Location"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </Button>
                    </div>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div>
                      <Label className="text-xs">Elevation</Label>
                      <Select
                        value={loc.elevation || ""}
                        onValueChange={(v) => updateLocationField(loc.id, "elevation", v)}
                      >
                        <SelectTrigger className="font-mono h-8 text-xs">
                          <SelectValue placeholder="—" />
                        </SelectTrigger>
                        <SelectContent>
                          {elevationOptions.map((el) => (
                            <SelectItem key={el.code} value={el.code}>{el.code}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                    <div>
                      <Label className="text-xs">Drop</Label>
                      <Input
                        value={loc.drop || ""}
                        onChange={(e) => updateLocationField(loc.id, "drop", e.target.value.replace(/\D/g, "").slice(0, 2))}
                        placeholder="01"
                        maxLength={2}
                        className="font-mono text-center h-8 text-xs"
                      />
                    </div>
                    <div>
                      <Label className="text-xs">Level</Label>
                      <Input
                        value={loc.level || ""}
                        onChange={(e) => updateLocationField(loc.id, "level", e.target.value.replace(/\D/g, "").slice(0, 2))}
                        placeholder="01"
                        maxLength={2}
                        className="font-mono text-center h-8 text-xs"
                      />
                    </div>
                  </div>
                  <div>
                    <Label className="text-xs">Description</Label>
                    <Input
                      value={loc.description || ""}
                      onChange={(e) => updateLocationField(loc.id, "description", e.target.value)}
                      placeholder="Optional note about this location..."
                      className="h-8 text-xs"
                    />
                  </div>
                </Card>
              );
            })}
          </div>
        )}
        {!isEdit && (
          <p className="text-xs text-muted-foreground bg-accent/40 p-3 rounded-lg">
            Save the defect first, then you can add additional locations.
          </p>
        )}

        {/* Photo slots — only shown when editing (defect must exist first) */}
        {isEdit && (
          <div className="space-y-4">
            <div className="flex items-center gap-2">
              <Camera className="w-4 h-4 text-muted-foreground" />
              <h3 className="text-sm font-medium">Photos</h3>
              <span className="text-xs text-muted-foreground">({photos.length}/6)</span>
            </div>

            <div className="grid grid-cols-2 gap-3">
              {PHOTO_SLOTS.map((slot) => {
                const photo = getPhotoForSlot(slot.key);
                const isSlotUploading = uploading === slot.key;
                const isCompleteSlot = slot.key === "complete";

                return (
                  <Card
                    key={slot.key}
                    className={`relative overflow-hidden ${isCompleteSlot ? "border-green-300 dark:border-green-700" : ""}`}
                    data-testid={`photo-slot-${slot.key}`}
                  >
                    {photo ? (
                      <div>
                        <div className="relative">
                          <img
                            src={`${API_BASE}/api/uploads/${photo.filename}`}
                            alt={slot.label}
                            className={`w-full aspect-[4/3] object-cover ${isPhotoNew(photo) ? "" : "opacity-[0.55]"}`}
                          />
                          <button
                            type="button"
                            onClick={() => handleDeletePhoto(photo.id)}
                            className="absolute top-2 right-2 p-1 rounded-full bg-black/60 text-white hover:bg-black/80 transition-colors"
                            data-testid={`button-delete-photo-${slot.key}`}
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                          <a
                            href={`${API_BASE}/api/uploads/${photo.filename}`}
                            download={`${uid || defectId}_${slot.key}.jpg`}
                            className="absolute top-2 left-2 p-1 rounded-full bg-black/60 text-white hover:bg-black/80 transition-colors"
                            data-testid={`button-download-photo-${slot.key}`}
                          >
                            <Download className="w-3.5 h-3.5" />
                          </a>
                          {/* New/Not-new toggle badge */}
                          <button
                            type="button"
                            onClick={() => handleToggleNewOverride(photo)}
                            className={`absolute top-2 left-10 px-1.5 py-0.5 rounded text-[10px] font-semibold transition-colors ${
                              isPhotoNew(photo)
                                ? "bg-blue-500/90 text-white hover:bg-blue-600"
                                : "bg-gray-500/70 text-white/80 hover:bg-gray-600"
                            }`}
                            title={isPhotoNew(photo) ? "Click to mark as NOT new this inspection" : "Click to mark as NEW this inspection"}
                          >
                            {isPhotoNew(photo) ? "NEW" : "OLD"}
                          </button>
                          <div className={`absolute bottom-0 left-0 right-0 px-2 py-1.5 text-xs font-medium ${isCompleteSlot ? "bg-green-600/90 text-white" : "bg-black/50 text-white"}`}>
                            {slot.label}
                          </div>
                        </div>
                        <div className="flex gap-1 p-1.5 bg-muted/30">
                          <input
                            type="text"
                            placeholder="Add comment..."
                            value={captions[photo.id] ?? ""}
                            onChange={(e) => updateCaption(photo.id, e.target.value)}
                            className="flex-1 text-xs px-2 py-1 rounded border bg-background placeholder:text-muted-foreground/60"
                            data-testid={`input-caption-${slot.key}`}
                          />
                          <DictationButton
                            onTranscript={(text) => updateCaption(photo.id, (captions[photo.id] ?? "") + (captions[photo.id] ? " " : "") + text)}
                            className="shrink-0"
                          />
                        </div>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center justify-center aspect-[4/3] bg-muted/50 p-3">
                        {isSlotUploading ? (
                          <div className="text-xs text-muted-foreground animate-pulse">Uploading...</div>
                        ) : (
                          <>
                            {isCompleteSlot ? (
                              <CheckCircle2 className="w-6 h-6 text-green-400/60 mb-2" />
                            ) : (
                              <Wrench className="w-6 h-6 text-muted-foreground/40 mb-2" />
                            )}
                            <span className={`text-xs font-medium mb-1 ${isCompleteSlot ? "text-green-600" : "text-muted-foreground"}`}>
                              {slot.label}
                            </span>
                            <span className="text-[10px] text-muted-foreground text-center leading-tight">
                              {slot.description}
                            </span>
                            <div className="flex gap-1.5 mt-2">
                              <input
                                ref={(el) => { cameraInputRefs.current[slot.key] = el; }}
                                type="file"
                                accept="image/*"
                                capture="environment"
                                className="hidden"
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) handlePhotoUpload(file, slot.key);
                                  e.target.value = "";
                                }}
                              />
                              <input
                                ref={(el) => { fileInputRefs.current[slot.key] = el; }}
                                type="file"
                                accept="image/*"
                                className="hidden"
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) handlePhotoUpload(file, slot.key);
                                  e.target.value = "";
                                }}
                              />
                              <button
                                type="button"
                                onClick={() => cameraInputRefs.current[slot.key]?.click()}
                                className="p-1.5 rounded-md bg-background border hover:bg-accent transition-colors"
                                data-testid={`button-camera-${slot.key}`}
                              >
                                <Camera className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => fileInputRefs.current[slot.key]?.click()}
                                className="p-1.5 rounded-md bg-background border hover:bg-accent transition-colors"
                                data-testid={`button-upload-${slot.key}`}
                              >
                                <Upload className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </>
                        )}
                      </div>
                    )}
                  </Card>
                );
              })}
            </div>
          </div>
        )}

        {!isEdit && (
          <p className="text-sm text-muted-foreground bg-accent/40 p-3 rounded-lg">
            Save the defect first, then you can add photos.
          </p>
        )}

        {/* Submit */}
        <Button type="submit" className="w-full" disabled={saveMutation.isPending || !canSave} data-testid="button-save-defect">
          <Save className="w-4 h-4 mr-2" />
          {saveMutation.isPending ? "Saving..." : isEdit ? "Update Defect" : "Save Defect"}
        </Button>

        </>)}
      </form>
    </div>
  );
}
