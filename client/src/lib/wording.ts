// WP2: record wording. Projects created after this change (featureVersion >= 2) say
// "record" in general wording, and the record form follows its type ("observation" or
// "defect"). Older projects keep their original wording. Internal names, data and web
// addresses keep the word "defect" either way.

export const RECORD_WORDING_VERSION = 2;

/** True when the project uses the new wording. Screens outside any project always do. */
export function usesRecordWording(project: { featureVersion?: number | null } | null | undefined): boolean {
  return Number(project?.featureVersion ?? 1) >= RECORD_WORDING_VERSION;
}

type Words = { one: string; One: string; many: string; Many: string };

const RECORD: Words = { one: "record", One: "Record", many: "records", Many: "Records" };
const DEFECT: Words = { one: "defect", One: "Defect", many: "defects", Many: "Defects" };
const OBSERVATION: Words = { one: "observation", One: "Observation", many: "observations", Many: "Observations" };

/** General wording for any record, whatever its type. */
export function recordWords(project: { featureVersion?: number | null } | null | undefined): Words {
  return usesRecordWording(project) ? RECORD : DEFECT;
}

/** Wording inside a record's own form: follows its type in new projects. */
export function typeWords(
  project: { featureVersion?: number | null } | null | undefined,
  recordType: string | null | undefined,
): Words {
  if (!usesRecordWording(project)) return DEFECT;
  return recordType === "observation" ? OBSERVATION : DEFECT;
}
