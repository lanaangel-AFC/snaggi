// Controlled editors for the §1.1 Roles, §1.2 Scope of Works, and §1.4 Background
// Documents fields added in commit 1. Pure controlled components — they hold no
// internal state for the data; the parent owns the JSON string. This lets the
// same component slot into both the project-create dialog (project-list.tsx) and
// the project-edit dialog (project-detail.tsx) without any new API endpoints.
//
// Each editor stores its value as a JSON-encoded array of plain objects matching
// the shapes documented in shared/schema.ts:
//   roles            JSON [{role, entity, contactDetails}]
//   scopeOfWorks     JSON [{areaRef, location, workItem, accessMethod}]
//   backgroundDocs   JSON [{type, originator, title, docNumbers?, revision?, date}]

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { Plus, Trash2 } from "lucide-react";
import { parseUidFormat, describeUidFormat, exampleUid, partKeyFromLabel } from "@shared/uid-format";
import { useMemo } from "react";

// ---- helpers ---------------------------------------------------------------

function parseRows<T>(json: string, fallback: T[] = []): T[] {
  try {
    const v = JSON.parse(json || "[]");
    return Array.isArray(v) ? v : fallback;
  } catch {
    return fallback;
  }
}

// ---- Roles -----------------------------------------------------------------

type Role = { role: string; entity: string; contactDetails: string };

export function RolesEditor({ value, onChange }: { value: string; onChange: (next: string) => void }) {
  const rows = useMemo(() => parseRows<Role>(value), [value]);

  const update = (next: Role[]) => onChange(JSON.stringify(next));
  const set = (idx: number, patch: Partial<Role>) => {
    const next = [...rows];
    next[idx] = { ...next[idx], ...patch };
    update(next);
  };
  const add = () => update([...rows, { role: "", entity: "", contactDetails: "" }]);
  const remove = (idx: number) => update(rows.filter((_, i) => i !== idx));

  return (
    <div>
      <Label className="mb-2 block">Roles (§1.1)</Label>
      <p className="text-xs text-muted-foreground mb-2">
        Principal, Contractor, Principal's Engineer, etc. Contact Details accepts
        multiple lines (Name, M | phone, E | email).
      </p>
      <div className="space-y-2">
        {rows.map((row, idx) => (
          <div key={idx} className="grid grid-cols-12 gap-2 items-start">
            <Input
              className="col-span-3"
              placeholder="Role"
              value={row.role}
              onChange={(e) => set(idx, { role: e.target.value })}
            />
            <Input
              className="col-span-3"
              placeholder="Entity"
              value={row.entity}
              onChange={(e) => set(idx, { entity: e.target.value })}
            />
            <Textarea
              className="col-span-5 min-h-[64px] text-sm"
              placeholder={"Name\nM | 0407 759 590\nE | name@example.com"}
              value={row.contactDetails}
              onChange={(e) => set(idx, { contactDetails: e.target.value })}
            />
            <Button
              type="button"
              variant="ghost"
              size="icon"
              className="col-span-1"
              onClick={() => remove(idx)}
              aria-label="Remove role"
            >
              <Trash2 className="w-4 h-4" />
            </Button>
          </div>
        ))}
        <Button type="button" variant="outline" size="sm" onClick={add}>
          <Plus className="w-3.5 h-3.5 mr-1.5" /> Add role
        </Button>
      </div>
    </div>
  );
}

// ---- Scope of Works --------------------------------------------------------

type Scope = { areaRef: string; location: string; workItem: string; accessMethod: string };

export function ScopeOfWorksEditor({ value, onChange }: { value: string; onChange: (next: string) => void }) {
  const rows = useMemo(() => parseRows<Scope>(value), [value]);

  const update = (next: Scope[]) => onChange(JSON.stringify(next));
  const set = (idx: number, patch: Partial<Scope>) => {
    const next = [...rows];
    next[idx] = { ...next[idx], ...patch };
    update(next);
  };
  const add = () => update([...rows, { areaRef: "", location: "", workItem: "", accessMethod: "" }]);
  const remove = (idx: number) => update(rows.filter((_, i) => i !== idx));

  return (
    <div>
      <Label className="mb-2 block">Scope of works under review (§1.2)</Label>
      <p className="text-xs text-muted-foreground mb-2">
        Columns: Area ref · Location (Elevation / Floor) · Work item · Access method.
      </p>
      <div className="space-y-2">
        {rows.map((row, idx) => (
          <div key={idx} className="grid grid-cols-12 gap-2 items-start">
            <Input
              className="col-span-2"
              placeholder="Area ref"
              value={row.areaRef}
              onChange={(e) => set(idx, { areaRef: e.target.value })}
            />
            <Input
              className="col-span-3"
              placeholder="Location"
              value={row.location}
              onChange={(e) => set(idx, { location: e.target.value })}
            />
            <Input
              className="col-span-3"
              placeholder="Work item"
              value={row.workItem}
              onChange={(e) => set(idx, { workItem: e.target.value })}
            />
            <Input
              className="col-span-3"
              placeholder="Access method"
              value={row.accessMethod}
              onChange={(e) => set(idx, { accessMethod: e.target.value })}
            />
            <Button
              type="button"
              variant="ghost"
              size="icon"
              className="col-span-1"
              onClick={() => remove(idx)}
              aria-label="Remove row"
            >
              <Trash2 className="w-4 h-4" />
            </Button>
          </div>
        ))}
        <Button type="button" variant="outline" size="sm" onClick={add}>
          <Plus className="w-3.5 h-3.5 mr-1.5" /> Add row
        </Button>
      </div>
    </div>
  );
}

// ---- Background documents --------------------------------------------------

type BgDoc = {
  type: string;
  originator: string;
  title: string;
  docNumbers?: string;
  revision?: string;
  date: string;
};

const BG_TYPES = [
  { value: "drawing", label: "Drawing set" },
  { value: "specification", label: "Specification" },
  { value: "manual", label: "Manual" },
  { value: "report", label: "Report" },
  { value: "other", label: "Other" },
];

export function BackgroundDocsEditor({ value, onChange }: { value: string; onChange: (next: string) => void }) {
  const rows = useMemo(() => parseRows<BgDoc>(value), [value]);

  const update = (next: BgDoc[]) => onChange(JSON.stringify(next));
  const set = (idx: number, patch: Partial<BgDoc>) => {
    const next = [...rows];
    next[idx] = { ...next[idx], ...patch };
    update(next);
  };
  const add = () =>
    update([...rows, { type: "drawing", originator: "", title: "", docNumbers: "", revision: "", date: "" }]);
  const remove = (idx: number) => update(rows.filter((_, i) => i !== idx));

  return (
    <div>
      <Label className="mb-2 block">Background documents (§1.4)</Label>
      <p className="text-xs text-muted-foreground mb-2">
        Drawings, specifications, manuals, reports. Rendered as a Harvard-style
        reference list inline in §1.4.
      </p>
      <div className="space-y-3">
        {rows.map((row, idx) => (
          <div key={idx} className="border rounded-md p-3 space-y-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 w-full max-w-xs">
                <Label className="text-xs w-16 shrink-0">Type</Label>
                <Select value={row.type} onValueChange={(v) => set(idx, { type: v })}>
                  <SelectTrigger className="h-8">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {BG_TYPES.map((t) => (
                      <SelectItem key={t.value} value={t.value}>
                        {t.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <Button
                type="button"
                variant="ghost"
                size="icon"
                onClick={() => remove(idx)}
                aria-label="Remove document"
              >
                <Trash2 className="w-4 h-4" />
              </Button>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <Input
                placeholder="Originator (e.g. UDA)"
                value={row.originator}
                onChange={(e) => set(idx, { originator: e.target.value })}
              />
              <Input
                placeholder="Title"
                value={row.title}
                onChange={(e) => set(idx, { title: e.target.value })}
              />
              <Input
                placeholder="Document number(s) (optional)"
                value={row.docNumbers || ""}
                onChange={(e) => set(idx, { docNumbers: e.target.value })}
              />
              <Input
                placeholder="Revision (optional)"
                value={row.revision || ""}
                onChange={(e) => set(idx, { revision: e.target.value })}
              />
              <Input
                placeholder="Date (YYYY or DD/MM/YYYY)"
                value={row.date}
                onChange={(e) => set(idx, { date: e.target.value })}
              />
            </div>
          </div>
        ))}
        <Button type="button" variant="outline" size="sm" onClick={add}>
          <Plus className="w-3.5 h-3.5 mr-1.5" /> Add document
        </Button>
      </div>
    </div>
  );
}

// ---- Area Ref template (NEW projects only) --------------------------------

export function AreaRefTemplateEditor({ value, onChange }: { value: string; onChange: (next: string) => void }) {
  // Live preview: substitute sample codes so the user sees exactly what a real
  // defect UID will look like before saving the template.
  const sample = (value || "")
    .replace(/\{elevation\}/g, "E")
    .replace(/\{drop\}/g, "04")
    .replace(/\{level\}/g, "07");
  const hasPlaceholder = /\{(elevation|drop|level)\}/.test(value || "");
  const isLiteral = !!value && !hasPlaceholder;
  return (
    <div>
      <Label className="mb-2 block">Area Ref template (§1.5.1)</Label>
      <p className="text-xs text-muted-foreground mb-2">
        Type a <strong>pattern</strong>, not a literal value. Use the placeholders{" "}
        <code>{"{elevation}"}</code>, <code>{"{drop}"}</code>, <code>{"{level}"}</code>
        {" "}with literal separators. Example: <code>{"{elevation}{drop}-{level}"}</code>{" "}
        → <code>E04-07</code>. Final UID becomes <code>AreaRef-WorkItem-Seq#</code>.
        Leave blank to keep the legacy 5-part UID.
      </p>
      <Input
        placeholder="{elevation}{drop}-{level}"
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
      {value ? (
        <div className="mt-2 text-xs">
          <span className="text-muted-foreground">Preview (with sample codes E/04/07/CR/01):</span>{" "}
          <code className="font-mono">{sample}-CR-01</code>
        </div>
      ) : null}
      {isLiteral ? (
        <p className="mt-2 text-xs text-amber-600 dark:text-amber-400">
          ⚠ Your template has no <code>{"{...}"}</code> placeholders, so this exact text
          will be reused for every defect (e.g. all UIDs become <code>{sample}-CR-01</code>,
          <code>{sample}-CR-02</code>, …). To vary the Area Ref per defect, include at
          least one placeholder like <code>{"{elevation}"}</code>.
        </p>
      ) : null}
    </div>
  );
}

// ---- Record UID format (NEW projects only) --------------------------------
// Optional. When set, records in this project are identified by the project's own
// parts and pattern (e.g. Level, Fin, Bracket → "{level}.{fin}-{bracket}" → 9.14-B)
// instead of Elevation-Drop-Level-WorkType-Number. Stored as JSON, see shared/uid-format.ts.

type UidPartDraft = { label: string; options: string };

function readUidFormatDraft(value: string): { pattern: string; parts: UidPartDraft[] } {
  try {
    const obj = value ? JSON.parse(value) : null;
    if (obj && typeof obj === "object") {
      return {
        pattern: String(obj.pattern || ""),
        parts: (Array.isArray(obj.parts) ? obj.parts : []).map((p: any) => ({
          label: String(p?.label || p?.key || ""),
          options: typeof p?.optionsText === "string" ? p.optionsText : (Array.isArray(p?.options) ? p.options.join(", ") : ""),
        })),
      };
    }
  } catch { /* fall through */ }
  return { pattern: "", parts: [] };
}

function writeUidFormatDraft(pattern: string, parts: UidPartDraft[]): string {
  if (!pattern.trim() && parts.every((p) => !p.label.trim())) return "";
  return JSON.stringify({
    pattern,
    parts: parts.map((p) => ({
      key: partKeyFromLabel(p.label),
      label: p.label,
      options: p.options.split(",").map((o) => o.trim()).filter(Boolean),
      optionsText: p.options, // raw text so typing "A, " isn't tidied away mid-edit
    })),
  });
}

export function UidFormatEditor({ value, onChange }: { value: string; onChange: (next: string) => void }) {
  const draft = readUidFormatDraft(value || "");
  const fmt = parseUidFormat(value || "");
  const update = (pattern: string, parts: UidPartDraft[]) => onChange(writeUidFormatDraft(pattern, parts));
  const setPart = (i: number, patch: Partial<UidPartDraft>) => {
    const parts = draft.parts.map((p, j) => (j === i ? { ...p, ...patch } : p));
    update(draft.pattern, parts);
  };
  const tokens = draft.parts.map((p) => partKeyFromLabel(p.label)).filter(Boolean);
  const unused = tokens.filter((k) => !draft.pattern.includes(`{${k}}`));
  return (
    <div data-testid="uid-format-editor">
      <Label className="mb-2 block">Record UID format</Label>
      <p className="text-xs text-muted-foreground mb-2">
        Optional. Set your own UID parts and how they join, instead of
        Elevation-Drop-Level-Work type-Number. Each part becomes the record's location;
        the pick list is optional and you can always type a new value. Example: parts
        Level, Fin, Bracket with pattern <code>{"{level}.{fin}-{bracket}"}</code> gives{" "}
        <code>9.14-B</code>. Leave empty to keep the standard UID.
      </p>
      <div className="space-y-2">
        {draft.parts.map((p, i) => (
          <div key={i} className="grid grid-cols-[7rem_1fr_auto] gap-2 items-center">
            <Input
              placeholder="Part (e.g. Fin)"
              value={p.label}
              onChange={(e) => setPart(i, { label: e.target.value })}
              data-testid={`input-uid-part-label-${i}`}
            />
            <Input
              placeholder="Optional list, e.g. A, B, C"
              value={p.options}
              onChange={(e) => setPart(i, { options: e.target.value })}
              data-testid={`input-uid-part-options-${i}`}
            />
            <Button type="button" variant="ghost" size="icon"
              onClick={() => update(draft.pattern, draft.parts.filter((_, j) => j !== i))}
              aria-label="Remove part">
              <Trash2 className="w-4 h-4" />
            </Button>
          </div>
        ))}
        <Button type="button" variant="outline" size="sm"
          onClick={() => update(draft.pattern, [...draft.parts, { label: "", options: "" }])}
          data-testid="button-add-uid-part">
          <Plus className="w-4 h-4 mr-1" /> Add part
        </Button>
      </div>
      {draft.parts.length > 0 && (
        <div className="mt-3">
          <Label className="text-xs mb-1 block">Pattern</Label>
          <div className="flex gap-2">
            <Input
              className="font-mono"
              placeholder={tokens.map((k) => `{${k}}`).join("-") || "{level}.{fin}-{bracket}"}
              value={draft.pattern}
              onChange={(e) => update(e.target.value, draft.parts)}
              data-testid="input-uid-pattern"
            />
            <Button type="button" variant="outline" size="sm"
              onClick={() => update(tokens.map((k) => `{${k}}`).join("-"), draft.parts)}>
              Fill
            </Button>
          </div>
          <p className="text-xs text-muted-foreground mt-1">
            Placeholders: {tokens.map((k) => <code key={k} className="mr-1">{`{${k}}`}</code>)}
            Anything else is typed as-is, e.g. <code>.</code> or <code>-</code>.
          </p>
          {fmt ? (
            <p className="mt-1 text-xs">
              <span className="text-muted-foreground">Format:</span> {describeUidFormat(fmt)}{" "}
              <span className="text-muted-foreground">e.g.</span>{" "}
              <code className="font-mono">{exampleUid(fmt)}</code>
            </p>
          ) : null}
          {unused.length > 0 ? (
            <p className="mt-1 text-xs text-amber-600 dark:text-amber-400">
              Not in the pattern yet: {unused.map((k) => `{${k}}`).join(", ")}
            </p>
          ) : null}
        </div>
      )}
    </div>
  );
}
