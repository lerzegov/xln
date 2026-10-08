// L3 audit: checks C1–C15 over a workbook snapshot, as findings, a name census and a spill census.
export { audit, checkTitle, DEFAULT_CONSTANTS, DEFAULT_LIMITS, DEFAULT_SENTINEL_ABOVE } from "./audit.js";
export { auditOptions, CONFIG_FILE, CONFIG_KEYS, configKeyRange, defaultConfigText, embedSetting, NAMES_SCOPE_NOTE, parseConfig, unknownKeyMessage, type AuditSettings, type BuildSettings, type ConfigIssue, type ProjectConfig } from "./config.js";
export { RULES, CHECK_TITLES, fill, type RuleSpec } from "./rules.js";
export { nameCensus, nameFamilies, globMatch, CENSUS_EXPLANATION, type FamilyAgreement } from "./census.js";
export { callUses, type CallUse } from "./walk.js";
export { afeStatus, afeStoreLabel, afeNamesTouched, type AfeEntryState, type AfeEntryStatus, type AfeStoreStatus } from "./afe.js";
export { renderAuditReport, placeText, type RenderedAudit, type AuditLink, type RenderAuditOptions } from "./report.js";
export { CHECK_IDS } from "./types.js";
export type * from "./types.js";
