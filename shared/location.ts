// Single source of truth for rendering a defect/observation location string.
//
// Both the summary register and the detail card MUST call this one helper so they
// can never disagree. The input is the structured location object stored on the
// defect (location_structured) plus the project's configured location dimensions.
//
// This is project-general: no location dimension is hard-coded. The dimensions
// come from the project record (projects.locationDimensions, JSON array).

// Default dimensions when a project has none configured (matches the East
// Elevation protocol: elevation / drop / level).
export const DEFAULT_LOCATION_DIMENSIONS = ["elevation", "drop", "level"];

// Full-name mapping for the standard compass elevation codes. Anything not in
// this map is rendered as-is (so custom elevation labels still display).
const ELEVATION_NAMES: Record<string, string> = {
  N: "North", S: "South", E: "East", W: "West",
  NE: "North East", NW: "North West", SE: "South East", SW: "South West",
};

export type LocationStructured = Record<string, string | undefined | null>;

// Human label for a single dimension value.
function formatDimension(dim: string, value: string): string {
  const v = String(value).trim();
  if (!v) return "";
  switch (dim) {
    case "elevation":
      return `${ELEVATION_NAMES[v] || v} Elevation`;
    case "drop":
      return `Drop ${parseInt(v, 10) || v}`;
    case "level":
      return `Level ${parseInt(v, 10) || v}`;
    case "stage":
      return `Stage ${parseInt(v, 10) || v}`;
    default: {
      // Generic fallback for any project-defined dimension (e.g. "zone", "bay").
      const label = dim.charAt(0).toUpperCase() + dim.slice(1);
      return `${label} ${v}`;
    }
  }
}

// Parse the stored location_structured value (JSON string or object) into an object.
export function parseLocationStructured(raw: unknown): LocationStructured {
  if (!raw) return {};
  if (typeof raw === "object") return raw as LocationStructured;
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  }
  return {};
}

// THE single location formatter. Register, card, export, and any other consumer
// all call this. Renders dimensions in the order the project declares them.
export function formatLocation(
  locationStructured: unknown,
  locationDimensions: string[] = DEFAULT_LOCATION_DIMENSIONS,
): string {
  const loc = parseLocationStructured(locationStructured);
  const dims = Array.isArray(locationDimensions) && locationDimensions.length > 0
    ? locationDimensions
    : DEFAULT_LOCATION_DIMENSIONS;
  const segments: string[] = [];
  for (const dim of dims) {
    const value = loc[dim];
    if (value != null && String(value).trim() !== "") {
      segments.push(formatDimension(dim, String(value)));
    }
  }
  return segments.join(", ");
}

// Parse a project's locationDimensions config. Falls back to the default
// (elevation/drop/level) when missing or malformed.
export function getLocationDimensions(rawConfig: unknown): string[] {
  if (Array.isArray(rawConfig) && rawConfig.length > 0) return rawConfig as string[];
  if (typeof rawConfig === "string" && rawConfig.trim()) {
    try {
      const parsed = JSON.parse(rawConfig);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed as string[];
    } catch {
      // fall through to default
    }
  }
  return DEFAULT_LOCATION_DIMENSIONS;
}
