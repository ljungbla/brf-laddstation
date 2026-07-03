# Debitering Billaddare – Brf Lavetten

A tiny, fully client-side web tool that produces the monthly EV-charging billing
spreadsheet. You upload two files, it does the summing and VAT maths, and you download
a ready report. No server, no database — everything runs in the browser and no data
leaves your machine.

## How to use

1. Open the app.
2. Upload the **MALL FÖR DEBITERING** file (`.xlsx`). This is the registry of active
   RFID tags — the app reads each tenant's Tag Nr, name, apartment number, Badge Id and
   price per kWh from it.
3. Upload the **Hager export** (`.xlsm` or `.xlsx`).
4. Check the detected **month/year** (auto-filled from the export dates; editable).
5. Review the preview table and click **Ladda ner rapport**. You get
   `Debitering Billaddare {Månad} {År}.xlsx` with one row per tenant who charged, plus a
   totals row.

Only tenants who actually charged that period are included.

### The "saknas i MALL-filen" warning

If a tag charged but its Badge Id isn't found in the MALL file, it is **not** billed and
is listed in a warning box instead (the tool never guesses who to charge). When a listed
badge is one character off from a registry badge, it suggests the likely tenant so you
can fix the typo in the MALL file. In the sample data, Emma Pihlgren's MALL Badge Id
`e0eccfa` is missing a character versus the export's `e0eccefa` — fix it in the MALL
file and she'll be billed correctly.

## How the numbers are computed

Per tenant: `kWh` (summed from the export, matched by Badge Id) →
`kostnad = kWh × pris` → `moms = kostnad × 25%` → `att betala = kostnad + moms`.
Price per kWh is read from the tenant's block in the MALL file (3.00 for members;
the external "Grannförening" uses 3.80 and is flagged as invoiced separately).

## Develop

```bash
npm install
npm run dev      # local dev server
npm run build    # type-check + production build into dist/
npm run preview  # preview the production build
```

`verify.mjs` is a dev harness that runs the parsers against the sample files:
`npx tsx verify.mjs`.

## Deploy (GitHub Pages)

Pushing to `main` triggers `.github/workflows/deploy.yml`, which builds and publishes to
GitHub Pages. One-time setup: in the repo **Settings → Pages**, set **Source** to
**GitHub Actions**.

> **Important:** `base` in `vite.config.ts` must match the repo name so assets resolve
> when hosted at `https://<user>.github.io/<repo>/`. It's currently `/brf-laddstation/`;
> change it if the repo is named differently.

## Notes

- Recommended browser: any modern one (Chrome, Edge, Firefox, Safari) — the app is
  static and needs no special permissions.
- Nothing is uploaded anywhere; file processing happens entirely in the browser.
