/** A tenant/tag entry read from the "MALL FÖR DEBITERING" template file. */
export type Tenant = {
  tagNr: string;
  namn: string;
  /** Full apartment / hyresavi number, e.g. "204044001-0097". */
  apartment: string;
  /** RFID badge id, normalized to lowercase hex — the join key to the Hager export. */
  badgeId: string;
  /** Price per kWh in SEK (3.00 for members, 3.80 for the external Grannförening). */
  pricePerKwh: number;
  /** True for the external "Grannföreningen" entry (invoiced separately, not on rent). */
  external: boolean;
};

/** Summed consumption for one badge id, aggregated from the Hager export. */
export type BadgeConsumption = {
  badgeId: string;
  /** Total kWh across all charging sessions in the export. */
  kwh: number;
  /** Free-text badge name from the export (col G) — for display/debugging only. */
  badgeName: string;
  /** Number of charging sessions. */
  sessions: number;
};

/** A tenant joined with their consumption and computed costs — one output row. */
export type ReportRow = {
  tenant: Tenant;
  kwh: number;
  /** kwh * pricePerKwh, in SEK. */
  cost: number;
  /** cost * 0.25 (moms 25%). */
  moms: number;
  /** cost + moms. */
  total: number;
};

/** Result of parsing the Hager export. */
export type ExportResult = {
  consumption: BadgeConsumption[];
  /** Detected billing period from the transaction dates. */
  period: { year: number; month: number } | null;
};

/** A charged badge with no matching tenant, plus an optional likely-typo hint. */
export type UnmatchedBadge = {
  consumption: BadgeConsumption;
  /** A registry badge that is one edit away — likely a typo in the MALL file. */
  suggestion: { badgeId: string; namn: string } | null;
};

/** Everything needed to render the preview and build the output workbook. */
export type BillingData = {
  rows: ReportRow[];
  /** Badges that charged but have no matching tenant in the MALL file. */
  unmatched: UnmatchedBadge[];
  /** Tenants in the MALL that did not charge this period (excluded from the report). */
  nonCharging: Tenant[];
};
