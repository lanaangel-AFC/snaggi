import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { Checkbox } from "@/components/ui/checkbox";
import { Plus, Trash2, ArrowUp, ArrowDown, Star, BookmarkPlus, X } from "lucide-react";
import { useState, useEffect } from "react";

type Category = { code: string; label: string; isDefault?: boolean };
type Treatment = { code: string; treatment: string };
type Profiles = {
  contractor: { filenameSuffix: string; categoryTreatments: Treatment[] };
  client: { filenameSuffix: string; categoryTreatments: Treatment[] };
};

const TREATMENTS = [
  { value: "itemise", label: "Itemise" },
  { value: "summarise", label: "Summarise" },
  { value: "hide", label: "Hide" },
];

const emptyProfiles: Profiles = {
  contractor: { filenameSuffix: "Contractor", categoryTreatments: [] },
  client: { filenameSuffix: "Client", categoryTreatments: [] },
};

// Project setup editors for Categories + Export Profiles. Self-contained: fetches and
// saves via the dedicated category/profile endpoints (auto-append handled server-side).
export default function ProjectSetupEditors({ projectId }: { projectId: string }) {
  const { toast } = useToast();

  const { data: serverCategories } = useQuery<Category[]>({ queryKey: [`/api/projects/${projectId}/categories`] });
  const { data: serverProfiles } = useQuery<Profiles>({ queryKey: [`/api/projects/${projectId}/export-profiles`] });

  const [categories, setCategories] = useState<Category[]>([]);
  const [profiles, setProfiles] = useState<Profiles>(emptyProfiles);
  const [newCode, setNewCode] = useState("");
  const [newLabel, setNewLabel] = useState("");
  const [newToLibrary, setNewToLibrary] = useState(false);

  // Generic category library, shared by every project.
  const { data: globalSettings } = useQuery<{ categoryLibrary?: Category[] }>({ queryKey: ["/api/global-settings"] });
  const library: Category[] = globalSettings?.categoryLibrary ?? [];
  const inLibrary = (code: string) => library.some((c) => c.code === code);
  const inProject = (code: string) => categories.some((c) => c.code === code);
  const unsaved = JSON.stringify(categories) !== JSON.stringify(serverCategories ?? []);

  const saveToLibrary = async (cat: { code: string; label: string }) => {
    try {
      const res = await apiRequest("POST", "/api/global-settings/categories", { code: cat.code, label: cat.label });
      const out = await res.json();
      queryClient.invalidateQueries({ queryKey: ["/api/global-settings"] });
      if (out.added) toast({ title: `${cat.code} saved to the generic library` });
      else if (out.existing) toast({ title: `${cat.code} is already in the library as "${out.existing.label}"` });
    } catch (e: any) {
      toast({ title: e?.message || "Could not save to the library", variant: "destructive" });
    }
  };

  const removeFromLibrary = async (cat: Category) => {
    if (!window.confirm(`Remove ${cat.code} (${cat.label}) from the generic library? Projects already using it keep it.`)) return;
    try {
      await apiRequest("DELETE", `/api/global-settings/categories/${encodeURIComponent(cat.code)}`);
      queryClient.invalidateQueries({ queryKey: ["/api/global-settings"] });
      toast({ title: `${cat.code} removed from the generic library` });
    } catch (e: any) {
      toast({ title: e?.message || "Could not remove from the library", variant: "destructive" });
    }
  };

  const toggleLibraryCategory = (cat: Category, on: boolean) => {
    if (on) {
      if (!inProject(cat.code)) setCategories([...categories, { code: cat.code, label: cat.label }]);
    } else {
      setCategories(categories.filter((c) => c.code !== cat.code));
    }
  };

  useEffect(() => { if (serverCategories) setCategories(serverCategories); }, [serverCategories]);
  useEffect(() => {
    if (serverProfiles) {
      setProfiles({
        contractor: serverProfiles.contractor || emptyProfiles.contractor,
        client: serverProfiles.client || emptyProfiles.client,
      });
    }
  }, [serverProfiles]);

  const saveCategoriesMutation = useMutation({
    mutationFn: async (cats: Category[]) => {
      const res = await apiRequest("PUT", `/api/projects/${projectId}/categories`, cats);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/categories`] });
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/export-profiles`] });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      toast({ title: "Categories saved" });
    },
    onError: (e: any) => toast({ title: e?.message || "Save failed", variant: "destructive" }),
  });

  const saveProfilesMutation = useMutation({
    mutationFn: async (p: Profiles) => {
      const res = await apiRequest("PUT", `/api/projects/${projectId}/export-profiles`, p);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/export-profiles`] });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      toast({ title: "Export profiles saved" });
    },
    onError: (e: any) => toast({ title: e?.message || "Save failed", variant: "destructive" }),
  });

  const addCategory = () => {
    const code = newCode.trim().toUpperCase();
    const label = newLabel.trim();
    if (!code || !label) { toast({ title: "Code and label required", variant: "destructive" }); return; }
    if (categories.some((c) => c.code === code)) { toast({ title: "Code already exists", variant: "destructive" }); return; }
    setCategories([...categories, { code, label }]);
    if (newToLibrary) saveToLibrary({ code, label });
    setNewCode(""); setNewLabel(""); setNewToLibrary(false);
  };

  const moveCategory = (idx: number, dir: -1 | 1) => {
    const next = [...categories];
    const j = idx + dir;
    if (j < 0 || j >= next.length) return;
    [next[idx], next[j]] = [next[j], next[idx]];
    setCategories(next);
  };

  const labelFor = (code: string) => categories.find((c) => c.code === code)?.label || code;

  const moveTreatment = (key: "contractor" | "client", idx: number, dir: -1 | 1) => {
    const list = [...profiles[key].categoryTreatments];
    const j = idx + dir;
    if (j < 0 || j >= list.length) return;
    [list[idx], list[j]] = [list[j], list[idx]];
    setProfiles({ ...profiles, [key]: { ...profiles[key], categoryTreatments: list } });
  };

  const setTreatment = (key: "contractor" | "client", idx: number, treatment: string) => {
    const list = [...profiles[key].categoryTreatments];
    list[idx] = { ...list[idx], treatment };
    setProfiles({ ...profiles, [key]: { ...profiles[key], categoryTreatments: list } });
  };

  const renderProfileColumn = (key: "contractor" | "client", title: string) => (
    <div className="flex-1 min-w-0">
      <h4 className="font-medium text-sm mb-2">{title}</h4>
      <div className="mb-2">
        <Label className="text-xs">Filename suffix</Label>
        <Input
          className="h-8"
          value={profiles[key].filenameSuffix}
          onChange={(e) => setProfiles({ ...profiles, [key]: { ...profiles[key], filenameSuffix: e.target.value } })}
        />
      </div>
      <div className="space-y-1.5">
        {profiles[key].categoryTreatments.map((t, idx) => (
          <div key={t.code} className="flex items-center gap-1 border rounded p-1.5">
            <span className="font-mono text-xs w-10 shrink-0">{t.code}</span>
            <span className="text-xs truncate flex-1 min-w-0">{labelFor(t.code)}</span>
            <Select value={t.treatment} onValueChange={(v) => setTreatment(key, idx, v)}>
              <SelectTrigger className="h-7 w-24 text-xs"><SelectValue /></SelectTrigger>
              <SelectContent>
                {TREATMENTS.map((tr) => <SelectItem key={tr.value} value={tr.value}>{tr.label}</SelectItem>)}
              </SelectContent>
            </Select>
            <button type="button" onClick={() => moveTreatment(key, idx, -1)} className="p-0.5 text-muted-foreground hover:text-foreground"><ArrowUp className="w-3.5 h-3.5" /></button>
            <button type="button" onClick={() => moveTreatment(key, idx, 1)} className="p-0.5 text-muted-foreground hover:text-foreground"><ArrowDown className="w-3.5 h-3.5" /></button>
          </div>
        ))}
      </div>
    </div>
  );

  return (
    <div className="space-y-6 border-t pt-4">
      {/* ===== Categories editor ===== */}
      <div>
        <Label className="mb-2 block">Categories</Label>
        <p className="text-xs text-muted-foreground mb-2">
          Follow-up action groups. Codes are fixed once created; labels and order can change. Adding a category appends it to both export profiles (itemise).
        </p>

        {/* Generic library: tick to use in this project */}
        <div className="rounded border bg-muted/30 p-2 mb-3" data-testid="category-library">
          <p className="text-xs font-medium mb-1.5">Generic categories</p>
          {library.length === 0 ? (
            <p className="text-xs text-muted-foreground">The library is empty. Save a category to it below.</p>
          ) : (
            <div className="space-y-1">
              {library.map((c) => (
                <div key={c.code} className="flex items-center gap-2">
                  <Checkbox
                    id={`lib-${c.code}`}
                    checked={inProject(c.code)}
                    onCheckedChange={(v) => toggleLibraryCategory(c, v === true)}
                    data-testid={`library-toggle-${c.code}`}
                  />
                  <label htmlFor={`lib-${c.code}`} className="flex-1 min-w-0 flex items-center gap-2 text-sm cursor-pointer">
                    <span className="font-mono text-xs w-10 shrink-0">{c.code}</span>
                    <span className="truncate">{c.label}</span>
                  </label>
                  <button
                    type="button"
                    title="Remove from the generic library"
                    onClick={() => removeFromLibrary(c)}
                    className="p-0.5 text-muted-foreground hover:text-destructive"
                    data-testid={`library-remove-${c.code}`}
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        <p className="text-xs font-medium mb-1.5">Used in this project</p>
        {categories.length === 0 && (
          <p className="text-xs text-muted-foreground mb-1.5">
            No categories yet. Tick generic ones above or add your own below, then save.
          </p>
        )}
        <div className="space-y-1.5">
          {categories.map((c, idx) => (
            <div key={c.code} className="flex items-center gap-1.5 border rounded p-1.5">
              <span className="font-mono text-xs w-10 shrink-0">{c.code}</span>
              <Input
                className="h-8 flex-1"
                value={c.label}
                onChange={(e) => { const next = [...categories]; next[idx] = { ...next[idx], label: e.target.value }; setCategories(next); }}
              />
              <button
                type="button"
                title="Mark as default"
                onClick={() => setCategories(categories.map((x, i) => ({ ...x, isDefault: i === idx ? true : undefined })))}
                className={`p-1 ${c.isDefault ? "text-amber-500" : "text-muted-foreground hover:text-foreground"}`}
              >
                <Star className="w-4 h-4" fill={c.isDefault ? "currentColor" : "none"} />
              </button>
              <button type="button" onClick={() => moveCategory(idx, -1)} className="p-0.5 text-muted-foreground hover:text-foreground"><ArrowUp className="w-4 h-4" /></button>
              <button type="button" onClick={() => moveCategory(idx, 1)} className="p-0.5 text-muted-foreground hover:text-foreground"><ArrowDown className="w-4 h-4" /></button>
              {!inLibrary(c.code) && (
                <button
                  type="button"
                  title="Save to the generic library for future projects"
                  onClick={() => saveToLibrary(c)}
                  className="p-0.5 text-muted-foreground hover:text-foreground"
                  data-testid={`save-to-library-${c.code}`}
                >
                  <BookmarkPlus className="w-4 h-4" />
                </button>
              )}
              <button type="button" onClick={() => setCategories(categories.filter((_, i) => i !== idx))} className="p-0.5 text-muted-foreground hover:text-destructive"><Trash2 className="w-4 h-4" /></button>
            </div>
          ))}
        </div>
        <div className="flex gap-2 mt-2">
          <Input className="h-8 w-20 font-mono" placeholder="CODE" value={newCode} onChange={(e) => setNewCode(e.target.value)} data-testid="input-new-category-code" />
          <Input className="h-8 flex-1 min-w-0" placeholder="Label" value={newLabel} onChange={(e) => setNewLabel(e.target.value)} data-testid="input-new-category-label" />
          <Button type="button" size="sm" variant="outline" onClick={addCategory} data-testid="button-add-category"><Plus className="w-4 h-4" /></Button>
        </div>
        <div className="flex items-center gap-2 mt-1.5">
          <Checkbox id="new-to-library" checked={newToLibrary} onCheckedChange={(v) => setNewToLibrary(v === true)} data-testid="checkbox-new-to-library" />
          <label htmlFor="new-to-library" className="text-xs text-muted-foreground cursor-pointer">Also save to the generic library for future projects</label>
        </div>
        <div className="flex items-center gap-3 mt-2">
          <Button type="button" size="sm" onClick={() => saveCategoriesMutation.mutate(categories)} disabled={saveCategoriesMutation.isPending} data-testid="button-save-categories">
            {saveCategoriesMutation.isPending ? "Saving..." : "Save Categories"}
          </Button>
          {unsaved && <span className="text-xs text-amber-600" data-testid="categories-unsaved">Unsaved changes</span>}
        </div>
      </div>

      {/* ===== Export Profiles editor ===== */}
      <div>
        <Label className="mb-2 block">Export Profiles</Label>
        <p className="text-xs text-muted-foreground mb-2">
          Order controls which category leads each export. Treatment shapes content (Pass 1 renders all as itemise).
        </p>
        <div className="flex flex-col sm:flex-row gap-4">
          {renderProfileColumn("contractor", "Contractor")}
          {renderProfileColumn("client", "Client")}
        </div>
        <Button type="button" size="sm" className="mt-2" onClick={() => saveProfilesMutation.mutate(profiles)} disabled={saveProfilesMutation.isPending}>
          {saveProfilesMutation.isPending ? "Saving..." : "Save Export Profiles"}
        </Button>
      </div>
    </div>
  );
}
