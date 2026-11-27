import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Plus, X } from "lucide-react";

type Attendee = { name: string; company: string };

function parseAttendees(raw: string | null | undefined): Attendee[] {
  try {
    const v = JSON.parse(raw || "[]");
    return Array.isArray(v) ? v.map((a: any) => ({ name: String(a?.name ?? ""), company: String(a?.company ?? "") })) : [];
  } catch {
    return [];
  }
}

// Index of the attendee who is the inspector: the first whose name matches it.
export function inspectorAttendeeIndex(attendees: Attendee[], inspector: string): number {
  const name = (inspector || "").trim();
  if (!name) return -1;
  return attendees.findIndex((a) => a.name.trim() === name);
}

// Attendees + Inspector for one inspection. The inspector is usually the AFC attendee:
// tick "Inspector" on that attendee and the name follows them. Otherwise type a name.
export function InspectionPeople({
  attendees,
  inspector,
  onChange,
}: {
  attendees: string;
  inspector: string;
  onChange: (next: { attendees: string; inspector: string }) => void;
}) {
  const list = parseAttendees(attendees);
  const tickedIdx = inspectorAttendeeIndex(list, inspector);

  const emit = (next: Attendee[], nextInspector: string) =>
    onChange({ attendees: JSON.stringify(next), inspector: nextInspector });

  return (
    <div className="space-y-4">
      <div>
        <Label>Attendees</Label>
        <div className="space-y-3 mt-1">
          {list.map((attendee, index) => (
            <div key={index} className="space-y-1.5">
              <div className="flex items-center gap-2">
                <Input
                  placeholder="Name"
                  value={attendee.name}
                  onChange={(e) => {
                    const next = list.map((a) => ({ ...a }));
                    next[index].name = e.target.value;
                    // Renaming the ticked attendee renames the inspector with them.
                    emit(next, index === tickedIdx ? e.target.value : inspector);
                  }}
                  data-testid={`input-attendee-name-${index}`}
                />
                <Input
                  placeholder="Company / Role"
                  value={attendee.company}
                  onChange={(e) => {
                    const next = list.map((a) => ({ ...a }));
                    next[index].company = e.target.value;
                    emit(next, inspector);
                  }}
                  data-testid={`input-attendee-company-${index}`}
                />
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  className="shrink-0"
                  onClick={() => {
                    const next = list.filter((_, i) => i !== index);
                    emit(next, index === tickedIdx ? "" : inspector);
                  }}
                  aria-label="Remove attendee"
                >
                  <X className="w-4 h-4" />
                </Button>
              </div>
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <Checkbox
                  checked={index === tickedIdx}
                  disabled={!attendee.name.trim()}
                  onCheckedChange={(checked) => emit(list, checked ? attendee.name.trim() : "")}
                  data-testid={`checkbox-attendee-inspector-${index}`}
                />
                Inspector
              </label>
            </div>
          ))}
          <Button
            type="button"
            variant="outline"
            size="sm"
            onClick={() => emit([...list, { name: "", company: "" }], inspector)}
            data-testid="button-add-attendee"
          >
            <Plus className="w-4 h-4 mr-1" />
            Add Attendee
          </Button>
        </div>
      </div>
      <div>
        <Label htmlFor="inspection-inspector">Inspector</Label>
        <Input
          id="inspection-inspector"
          placeholder="Inspector name"
          value={inspector}
          readOnly={tickedIdx >= 0}
          className={tickedIdx >= 0 ? "bg-muted" : undefined}
          onChange={(e) => onChange({ attendees, inspector: e.target.value })}
          data-testid="input-inspection-inspector"
        />
        <p className="text-xs text-muted-foreground mt-1">
          {tickedIdx >= 0
            ? "Taken from the ticked attendee. Untick to enter a different name."
            : "Tick Inspector against an attendee, or type the inspector's name."}
        </p>
      </div>
    </div>
  );
}
