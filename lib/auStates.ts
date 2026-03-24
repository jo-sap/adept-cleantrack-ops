/** Australian states/territories — same vocabulary as contract sites (SiteManager). */
export const AU_STATES = ["ACT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA"] as const;

export type AustralianState = (typeof AU_STATES)[number];
