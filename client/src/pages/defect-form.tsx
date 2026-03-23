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
import { ArrowLeft, Camera, Upload, X, ImageIcon, Save, CheckCircle2, Clock, Wrench, Mic } from "lucide-react";
import { DictationButton } from "@/components/DictationButton";
import type { Defect, Photo, Project } from "@shared/schema";
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

export default function DefectForm() {
  const { projectId, defectId } = useParams<{ projectId: string; defectId?: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const fileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const cameraInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const isEdit = !!defectId;

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

  // Read query params from hash URL
  const hashParams = useMemo(() => {
    if (typeof window === "undefined") return new URLSearchParams();
    const hash = window.location.hash;
    const qIdx = hash.indexOf("?");
    if (qIdx === -1) return new URLSearchParams();
    return new URLSearchParams(hash.slice(qIdx));
  }, []);

  const [recordType, setRecordType] = useState(() => hashParams.get("type") || "defect");

  // Fetch project to get configured elevations
  const { data: project } = useQuery<Project>({
    queryKey: ["/api/projects", projectId],
  });

  // Build elevation options from project config
  const elevationOptions = useMemo(() => {
    if (!project?.elevations) return [];
    try {
      const configured: string[] = JSON.parse(project.elevations as string);
      // Map full names to short codes for UID
      const codeMap: Record<string, string> = {
        "North": "N", "South": "S", "East": "E", "West": "W",
        "North East": "NE", "North West": "NW", "South East": "SE", "South West": "SW",
      };
      return configured.map((label) => ({
        code: codeMap[label] || label.substring(0, 3).toUpperCase(),
        label,
      }));
    } catch { return []; }
  }, [project?.elevations]);

  const [elevation, setElevation] = useState("");
  const [drop, setDrop] = useState("01");
  const [level, setLevel] = useState("");
  const [workType, setWorkType] = useState("");
  const [seqNumber, setSeqNumber] = useState("01");

  const [uid, setUid] = useState("");
  const [photos, setPhotos] = useState<Photo[]>([]);
  const [uploading, setUploading] = useState<string | null>(null);

  // Build the UID prefix from the components
  const uidPrefix = useMemo(() => {
    if (!elevation || !drop || !level || !workType) return "";
    const dd = drop.padStart(2, "0");
    const ll = level.padStart(2, "0");
    return `${elevation}-${dd}-${ll}-${workType}`;
  }, [elevation, drop, level, workType]);

  // Full assembled UID for display
  const assembledUid = useMemo(() => {
    if (!uidPrefix) return "";
    return `${uidPrefix}-${seqNumber.padStart(2, "0")}`;
  }, [uidPrefix, seqNumber]);

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

  // Fetch next sequence number when prefix changes (new defect only)
  const { data: nextUidData } = useQuery<{ uid: string }>({
    queryKey: [`/api/projects/${projectId}/next-uid`, uidPrefix],
    queryFn: async () => {
      if (!uidPrefix) return { uid: "" };
      const res = await fetch(`${API_BASE}/api/projects/${projectId}/next-uid?prefix=${encodeURIComponent(uidPrefix)}`);
      if (!res.ok) throw new Error("Failed to fetch next UID");
      return res.json();
    },
    enabled: !isEdit && !!uidPrefix,
  });

  // When we get the next UID from the server, extract the sequence number
  useEffect(() => {
    if (nextUidData?.uid && !isEdit) {
      setUid(nextUidData.uid);
      const parts = nextUidData.uid.split("-");
      if (parts.length >= 4) {
        setSeqNumber(parts[parts.length - 1]);
      }
    }
  }, [nextUidData, isEdit]);

  useEffect(() => {
    if (existingDefect) {
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
      const parts = existingDefect.uid.split("-");
      if (parts.length >= 5) {
        // New format: Elevation-Drop-Level-WorkType-Number
        setElevation(parts[0]);
        setDrop(parts[1]);
        setLevel(parts[2]);
        setWorkType(parts[3]);
        setSeqNumber(parts[4]);
      } else if (parts.length >= 4) {
        // Old format: Drop-Level-WorkType-Number
        setDrop(parts[0]);
        setLevel(parts[1]);
        setWorkType(parts[2]);
        setSeqNumber(parts[3]);
      }
    }
  }, [existingDefect]);

  useEffect(() => {
    if (existingPhotos) setPhotos(existingPhotos);
  }, [existingPhotos]);

  // Helper to get photo for a specific slot
  const getPhotoForSlot = (slot: SlotKey): Photo | undefined => {
    return photos.find((p) => p.slot === slot);
  };

  const saveMutation = useMutation({
    mutationFn: async () => {
      if (isEdit) {
        const res = await apiRequest("PATCH", `/api/defects/${defectId}`, form);
        return res.json();
      } else {
        const res = await apiRequest("POST", `/api/projects/${projectId}/defects`, {
          ...form,
          uidPrefix,
          uidOverride: assembledUid,
          recordType,
        });
        return res.json();
      }
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/defects`] });
      const typeLabel = recordType === "observation" ? "Observation" : "Defect";
      toast({ title: isEdit ? `${typeLabel} updated` : `${typeLabel} created` });
      if (!isEdit) {
        navigate(`/projects/${projectId}/defects/${data.id}`, { replace: true });
      }
    },
    onError: (err: Error) => {
      toast({ title: err.message || "Failed to save", variant: "destructive" });
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

  const set = (field: string) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
    setForm((prev) => ({ ...prev, [field]: e.target.value }));

  const canSave = isEdit || (!!elevation && !!uidPrefix && !!seqNumber);
  const isComplete = form.status === "complete";

  return (
    <div className="max-w-2xl mx-auto px-4 py-6 pb-24">
      {/* Header */}
      <Link href={`/projects/${projectId}`}>
        <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4" data-testid="button-back-to-project">
          <ArrowLeft className="w-4 h-4" />
          Back to Project
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
            <button
              type="button"
              onClick={async () => {
                const newType = recordType === "defect" ? "observation" : "defect";
                setRecordType(newType);
                try {
                  await apiRequest("PATCH", `/api/defects/${defectId}`, { recordType: newType });
                  queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/defects`] });
                  toast({ title: `Converted to ${newType}` });
                } catch {
                  toast({ title: "Failed to convert", variant: "destructive" });
                  setRecordType(recordType);
                }
              }}
              className="text-xs text-muted-foreground hover:text-foreground underline underline-offset-2"
              data-testid="button-convert-type"
            >
              Convert to {recordType === "defect" ? "Observation" : "Defect"}
            </button>
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

      <form
        onSubmit={(e) => {
          e.preventDefault();
          if (!canSave) {
            toast({ title: "Fill in Elevation, Drop, Level, and Work Type to generate the defect ID", variant: "destructive" });
            return;
          }
          saveMutation.mutate();
        }}
        className="space-y-6"
      >
        {/* Defect ID builder */}
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
            <div>
              <Label htmlFor="elevation" className="text-xs">Elevation</Label>
              <Select
                value={elevation}
                onValueChange={setElevation}
                disabled={isEdit}
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
            </div>
            <div>
              <Label htmlFor="drop" className="text-xs">Drop</Label>
              <Input
                id="drop"
                placeholder="01"
                value={drop}
                onChange={(e) => {
                  const v = e.target.value.replace(/\D/g, "").slice(0, 2);
                  setDrop(v);
                }}
                maxLength={2}
                required
                disabled={isEdit}
                className="font-mono text-center"
                data-testid="input-drop"
              />
            </div>
            <div>
              <Label htmlFor="level" className="text-xs">Level</Label>
              <Input
                id="level"
                placeholder="13"
                value={level}
                onChange={(e) => {
                  const v = e.target.value.replace(/\D/g, "").slice(0, 2);
                  setLevel(v);
                }}
                maxLength={2}
                required
                disabled={isEdit}
                className="font-mono text-center"
                data-testid="input-level"
              />
            </div>
            <div>
              <Label htmlFor="workType" className="text-xs">Work Type</Label>
              <Select
                value={workType}
                onValueChange={setWorkType}
                disabled={isEdit}
              >
                <SelectTrigger className="font-mono" data-testid="select-work-type">
                  <SelectValue placeholder="—" />
                </SelectTrigger>
                <SelectContent>
                  {WORK_TYPES.map((wt) => (
                    <SelectItem key={wt.code} value={wt.code}>
                      {wt.code} — {wt.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
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
                required
                disabled={isEdit}
                className="font-mono text-center"
                data-testid="input-seq-number"
              />
            </div>
          </div>
          <p className="text-xs text-muted-foreground">
            Format: Elevation-Drop-Level-WorkType-Number. The number auto-suggests but you can change it.
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
              onTranscript={(text) => setForm((prev) => ({ ...prev, comment: prev.comment + (prev.comment ? " " : "") + text }))}
            />
          </div>
          <Textarea
            id="comment"
            placeholder="Describe the defect observed..."
            value={form.comment}
            onChange={set("comment")}
            rows={3}
            required
            data-testid="input-comment"
          />
        </div>

        {/* Action Required */}
        <div>
          <div className="flex items-center gap-2 mb-1">
            <Label htmlFor="actionRequired">Action Required</Label>
            <DictationButton
              onTranscript={(text) => setForm((prev) => ({ ...prev, actionRequired: prev.actionRequired + (prev.actionRequired ? " " : "") + text }))}
            />
          </div>
          <Textarea
            id="actionRequired"
            placeholder="What needs to be done to rectify this defect?"
            value={form.actionRequired}
            onChange={set("actionRequired")}
            rows={2}
            required
            data-testid="input-action-required"
          />
        </div>

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
              required
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
                            className="w-full aspect-[4/3] object-cover"
                          />
                          <button
                            type="button"
                            onClick={() => handleDeletePhoto(photo.id)}
                            className="absolute top-2 right-2 p-1 rounded-full bg-black/60 text-white hover:bg-black/80 transition-colors"
                            data-testid={`button-delete-photo-${slot.key}`}
                          >
                            <X className="w-3.5 h-3.5" />
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

      </form>
    </div>
  );
}
