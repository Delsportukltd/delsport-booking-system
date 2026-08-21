# Delsport UK Booking System

A multi-site facility booking and invoicing app for Delsport UK, who broker sports facility
lettings across several partner schools/clubs (St Augustine's Catholic High, St Benedict's
Catholic High, Stourport High School, Baxter College, Stourport Cricket Club, Stourport Rugby
Club). Staff log bookings per facility, track members/hirers, and raise PDF invoices per site.

**This is a React + Vite single-page app, not a Node.js/Express backend.** There is no custom
API server and no chat feature — the entire backend is Supabase (Postgres + Auth + Storage),
called directly from the browser via the Supabase JS client. Keep that framing in mind when
picking up work here.

Live at: https://delsportukltd.github.io/delsport-booking-system/ (migrated from Netlify in
August 2026 after the Netlify team account hit its usage-credit limit and blocked new deploys —
see git history for the old `netlify.toml`-based setup if reviving that path)

## 1. Tech stack

- **React 18** + **Vite 5** — single-page app, no router (page switching is a `page` state
  string in `BookingApp`, not URL-based)
- **Supabase** (`@supabase/supabase-js`) — Postgres database, Auth (email/password), and
  Storage (private `invoices` bucket for raised invoice PDFs). Anon key only — no service-role
  access, so all schema/storage changes must be run manually by the user in Supabase's SQL Editor
- **jsPDF** — client-side generation of invoice and booking-confirmation PDFs (no server-side
  rendering)
- **JSZip** — bundles multiple invoice PDFs into one `.zip` when raising more than one at a time
- **xlsx** — CSV/Excel import (members, bookings) and export (members to Excel)
- **lucide-react** — icon set
- **recharts** — Reports page charts (revenue trend, bar charts)
- **GitHub Pages** — hosting, served from the `Delsportukltd/delsport-booking-system` repo (public,
  required for Pages on a free org plan — no member/booking/invoice data lives in the repo, only
  app source). Deploys automatically via the `.github/workflows/deploy.yml` GitHub Actions
  workflow on every push to `main` (`VITE_SUPABASE_URL`/`VITE_SUPABASE_ANON_KEY` are set as repo
  secrets, injected at build time). `vite.config.js` sets `base: "/delsport-booking-system/"` to
  match the Pages sub-path — any hardcoded root-absolute asset path referenced from JS (not just
  `index.html`, which Vite rewrites automatically) needs `import.meta.env.BASE_URL` prefixed
  instead, or it will 404 under this sub-path (bit us once with the trust-logo image paths used
  in invoice PDF generation)
- No CSS framework — everything is inline `style={{}}` objects against a shared color/token
  object (`C`), plus one `<style>` block injected in `BookingApp` for things inline styles can't
  do (media queries, `@font-face`, `:has()` selectors)

Node/npm are managed via nvm on this machine — if `npm` isn't on `PATH` in a fresh shell, run:
```bash
export PATH="$HOME/.nvm/versions/node/v24.19.0/bin:$PATH"
```

## 2. Folder & file structure

The app is almost entirely one file. This is deliberate (it grew this way and hasn't needed
splitting), not an oversight — don't refactor into multiple files unless asked.

```
Booking System/
├── CLAUDE.md                  # this file
├── index.html                 # favicon links, theme-color meta, mounts #root
├── vite.config.js             # just the React plugin, nothing custom
├── netlify.toml                # build command + publish dir, nothing else
├── package.json
├── .env                        # VITE_SUPABASE_URL, VITE_SUPABASE_ANON_KEY (not committed)
├── supabase-schema_2.sql       # original schema dump — reference only, NOT authoritative;
│                                # many columns/tables have been added by later migrations
│                                # (see §3 "Schema drift" below) that only exist in the live DB
├── public/
│   ├── favicon.ico, favicon-16x16.png, favicon-32x32.png, apple-touch-icon.png,
│   │   icon-192.png, icon-512.png        # custom favicon: Delsport "D" + runner mark
│   │                                      # with a navy calendar badge, bottom-right
│   ├── magnificat-trust-logo.jpeg         # trust logo for St Augustine's / St Benedict's invoices
│   └── saet-trust-logo.png                # trust logo for Stourport High / Baxter College invoices
├── src/
│   ├── main.jsx                # createRoot(...).render(<App/>), nothing else
│   ├── lib/supabaseClient.js   # createClient() using the two VITE_ env vars
│   └── App.jsx                 # ~4,360 lines — everything else lives here
└── dist/                       # build output, not committed
```

### Map of `App.jsx` (in file order)

| Area | What's there |
|---|---|
| ~L1–50 | Imports, `C` color/token object, status metadata |
| ~L50–430 | Pure helper functions: date/time math, `.ics` export, CSV import/export/parsing, blackout-rule matching, clash/capacity detection, `bookingsInGroup`/`buildDocumentRows` (multi-facility helpers) |
| ~L440–615 | Generic UI primitives: `ConfirmDeleteModal`, `Modal`, `Field`, `Btn`, `StatusPill`, `EmptyState` |
| ~L540–615 | Generic Supabase sync helpers: `camelToSnake`/`snakeToCamel`, `toRow`, `TABLE_COLUMNS` |
| ~L616–675 | `SiteScopeSelector` — sidebar/drawer site picker |
| ~L675–964 | `BookingApp` — the root shell: data loading, the generic sync effect, sidebar + mobile top bar/bottom nav/drawer, page routing |
| ~L964–1060 | `LoginScreen` |
| ~L1060–1130 | `BackupModal` — JSON export/import of all tables |
| ~L1144–1420 | `Admin` (blackout rules + `TeamPanel`), `BlackoutModal` |
| ~L1419–1531 | `Dashboard` — stat cards, upcoming/cancelled panels, per-site "Upcoming lettings" table |
| ~L1544–2314 | `Bookings` page: list/week/grid views, `PlannerGrid`, `BookingsList`, `BookingsCalendar`, `ImportBookingsModal` |
| ~L2314–2656 | `BookingModal` — the new/edit booking form, including multi-facility selection |
| ~L2698–3057 | `Members` page, `MemberDetailModal`, `MemberModal`, `ImportMembersModal` |
| ~L3057–3322 | `SitesFacilities` page, `SiteModal`, `FacilityModal` |
| ~L3350–3648 | `Reports` page |
| ~L3648–4080 | Invoice/confirmation PDF builders: `nextInvoiceNumber`, `buildPaymentReference`, `paymentReferenceForInvoice`, `drawDocHeader`/`drawDocFooter`/`drawInfoColumns`, `buildInvoicePDF`, `buildConfirmationPDF` |
| ~L4080–end | `Invoices` page — eligibility preview, raise all/selected, history, delete |

## 3. Key architecture decisions

**Auth** — Supabase Auth, email/password only (no OAuth, no magic links). `Root` component
checks `supabase.auth.getSession()` on load and subscribes to `onAuthStateChange`. A 20-minute
inactivity timer signs the user out automatically (`Root`'s effect watching mousemove/keydown/click).
A `profiles` table (id, display_name) supplements the auth user with a display name; there's no
self-serve account management in the UI — accounts are created/removed from the Supabase
dashboard directly (see `TeamPanel`'s copy).

**No custom API layer** — every read/write goes straight from the browser to Supabase via the
JS client (`supabase.from(table)...`), relying on Postgres RLS policies (not shown in the repo —
configured per-table in the Supabase dashboard/SQL Editor) for access control. There are no
Netlify Functions currently (an earlier attempt at server-side email sending via Netlify
Functions + Nodemailer was built and then fully removed after Outlook SMTP auth failed for
personal accounts — don't reintroduce that path without checking history first).

**Generic sync layer, not per-table CRUD code** — `BookingApp` holds all six tables (`sites`,
`facilities`, `members`, `bookings`, `blackouts`, `invoices`) as React state, fetched once on
mount. A single effect diffs current state against a `prevRef` snapshot on every change and
calls a generic `syncTable(table, items, previous)` per table, which upserts changed rows and
deletes removed ones. `TABLE_COLUMNS` is the whitelist of camelCase fields that get written back
per table — **if you add a new field to a table, it must be added to `TABLE_COLUMNS` or it will
silently never persist.** `toRow`/`camelToSnake`/`snakeToCamel` handle the naming conversion.

**Multi-site scoping** — a sidebar/drawer `SiteScopeSelector` lets staff narrow the whole app to
one or more sites (stored in `localStorage`, not synced — it's a per-device preference). Every
page receives already-filtered `visible*` props computed once in `BookingApp` (`visibleSites`,
`visibleBookings`, etc.) rather than filtering per-page.

**Multi-facility bookings** — a booking can span more than one facility (e.g. "Sports Hall +
Classroom" booked together as one event). Implemented as multiple real booking rows — one per
facility — sharing a `groupId` (mirrors the existing `recurringId` pattern for weekly-repeat
bookings). Only the first row in the group carries the actual price; the others are £0, so
revenue sums elsewhere never double-count. `bookingsInGroup()` finds all rows in a group;
`buildDocumentRows()` collapses a group back into one combined line (e.g. "Sports Hall +
Classroom") for invoices/confirmations. Cancel/delete on any one leg cascades to the whole group
(see `setStatus`/`remove`/`removeSelected` in the `Bookings` component).

Combinable with a one-off booking or a "pick specific dates" series (**not** with weekly-repeat —
a multi-facility series running for months was judged too complicated, so `BookingModal` disables
the weekly-repeat button whenever extra facilities are selected). When combined with picked
dates, each date gets its *own* `groupId` (so cancelling one date's booking doesn't touch the
other dates), and all dates additionally share one `recurringId` linking the whole series — see
the `multiFacilityIds && data.repeatMode === "dates"` branch in `Bookings`' `save()`. This is
also why `BookingModal`'s `handleSave` hard-drops `extraFacilityIds` whenever repeat mode is
`"weekly"`: without that, a stale selection from before switching modes could otherwise slip
through and silently create the wrong bookings (this happened once — see git history).

**Delete confirmations, everywhere** — every delete action in the app (single or bulk) routes
through `ConfirmDeleteModal` before anything is actually removed. This was added deliberately
after an incident where bulk-deleting rows directly in the Supabase Table Editor (not through
this app) wiped ~95 bookings with no way to recover them (Supabase free tier has no
backups/point-in-time recovery). If you add a new delete action, wire it through
`ConfirmDeleteModal` — don't call a `remove`/`delete` function directly from a click handler.

**Invoicing** — one invoice PDF per member per site per billing period, generated client-side
with jsPDF and uploaded to a private Supabase Storage bucket (`invoices`), with a `pdf_path`
column on the `invoices` table so it can be re-downloaded via a signed URL later (not just a
one-time browser download). Four sites (St Augustine's, St Benedict's, Stourport High, Baxter
College) belong to a diocese/trust and get an extra trust logo on their invoices plus a
reconciliation-friendly **payment reference** (e.g. `STAmwad001` = site code + member initials +
running count) shown in place of the generic `DEL-2026-XXXX` invoice number, both on the PDF and
in the Invoice History table. `SITE_PAYMENT_REF_CODES` / `SITE_TRUST_LOGO_URLS` are the two maps
controlling which sites get this treatment — add a site to both maps to extend it elsewhere.
Payment references are **not stored** — they're recomputed from invoice history each time
(`buildPaymentReference` for a new one, `paymentReferenceForInvoice` for displaying an existing
one), so they stay correct even across a JSON backup/restore.

**Backup/restore** — a manual JSON export/import (`BackupModal`) covering all six tables,
including `invoices` (records only — the actual PDF bytes stay in Supabase Storage and are NOT
included in the backup file). This is a point-in-time snapshot, not automatic — there is
currently no scheduled backup. Supabase is on the free tier: no automatic backups, no
point-in-time recovery, 500MB database cap (current usage is trivial, ~350KB), 1GB file storage
cap (the more realistic constraint — invoice PDFs run ~1–2MB each due to embedded logos).

**Mobile layout** — the sidebar is desktop-only (`.dp-sidebar`, hidden under 767px). Mobile gets
a fixed top bar + bottom tab bar (Home/Bookings/Members/Invoices/Menu) + a slide-in drawer with
the full nav, all rendered unconditionally but CSS-hidden on desktop via a `<style>` media query
block in `BookingApp`, so nothing here needs JS viewport detection. Tables scroll horizontally
via `*:has(> table) { overflow-x: auto }` rather than `table { display: block }` — the latter was
tried first and broke column-width alignment (rows with a long cell value would get a huge,
wrongly-sized row). Several grids get a shared className (`.dp-stat-grid`, `.dp-dash-grid`,
`.dp-form-grid`) that collapses to fewer/one column under the same breakpoint.

### Known gotchas worth knowing before touching this code

- **`doc.splitTextToSize()` in jsPDF measures using whatever font/size is currently set on the
  `doc` object** — call `setFont`/`setFontSize` *before* measuring, not after, or wrapped text
  will be sized against the wrong font and overflow its column. Bit us once in the invoice
  facility-name wrapping.
- **The sandboxed browser preview used for testing doesn't save downloads to the real macOS
  `~/Downloads` folder.** A "missing" downloaded PDF during dev testing is almost always this,
  not a bug — verify by patching `URL.createObjectURL` to capture the blob directly instead of
  trusting the filesystem.
- **Checkbox clicks in the preview browser sometimes don't fire React's onChange when done via
  the `computer` tool's coordinate-based click** — use a real `element.click()` via
  `javascript_tool` instead, or `form_input` (which sets the DOM property directly and can drift
  out of sync with React state if mixed with real clicks).
- `TABLE_COLUMNS` must be kept in sync with any new column — see "Generic sync layer" above.

### Schema drift from `supabase-schema_2.sql`

The committed schema file is a snapshot from early in the project and is **out of date**. Columns/
tables added since via manual SQL migrations (given to the user to run, not tracked as files):
- `bookings.group_id` (uuid, nullable) — multi-facility booking linking
- `invoices` table — the whole table was added later (`invoice_number`, `member_id`, `site_id`,
  `period_start`, `period_end`, `total`, `booking_ids` jsonb, `pdf_path`, `created_at`), plus a
  private Storage bucket also named `invoices` with authenticated-only RLS policies
- `sites.bank_account_name` / `bank_sort_code` / `bank_account_number` / `vat_number` — per-site
  payment details for invoicing
- If you need to know the *actual* current schema, don't trust the `.sql` file — ask the user to
  run a quick introspection query, or check `TABLE_COLUMNS` in `App.jsx` (that list reflects what
  the app actually reads/writes, though it won't show DB-only columns like `created_at`).

## 4. Current implementation status

**Built and live in production:**
- Bookings: list/week/grid views, recurring (weekly or picked-dates), multi-facility combo
  bookings, blackout rules, capacity/clash detection, CSV import, .ics export, bulk + single
  delete with confirmation, search-by-member filter
- Members: CRUD, per-site restriction, CSV import, Excel export, booking history per member
- Sites & Facilities: CRUD, per-facility custom fields, per-site bank/VAT details, "load starter
  data from delsportuk.com"
- Invoices: eligibility preview, raise all/selected (PDF generated + uploaded to Storage +
  zipped if multiple), invoice history with signed-URL re-download, bulk/single delete,
  trust-site payment references and logos
- Booking confirmation PDFs (client-side download; automated email sending was attempted and
  abandoned — see below)
- Reports: revenue trend, top members, revenue by site/facility, utilisation, busiest times
- Dashboard: stat cards, upcoming/cancelled panels, per-site "upcoming lettings" table
- Admin: blackout rules, read-only team panel
- Backup & restore (manual JSON, all tables including invoice records)
- Mobile-responsive nav and layout
- Custom favicon/app icon

**Explicitly deferred / not built:**
- **Automated email sending** — tried SMTP (Outlook basic auth disabled for personal accounts),
  Azure OAuth2 app registration (blocked — personal accounts can't register apps outside a
  directory), and Power Automate (blocked by a Microsoft tenant error). Abandoned in favour of
  manual sending: bookings download a confirmation PDF instead of emailing it; a `mailto:` link
  still exists for the cancellation flow only. Don't restart this without checking whether the
  user has since set up a Microsoft 365 tenant or Google Workspace account — the blockers were
  all account-type issues, not code issues.
- **Google Drive invoice archiving** — the plan (agreed with the user) is a "Send to Google
  Drive" button to archive old invoice PDFs and free up Supabase Storage. Blocked on the user
  creating a Google Cloud project + OAuth Client ID (instructions already given to them). **Pick
  this up once they provide a Client ID** — implementation should use Google Identity Services
  (`accounts.google.com/gsi/client`) token client with the `drive.file` scope, upload via the
  Drive v3 API, and keep deletion-from-Storage as a separate deliberate step afterward (through
  the existing `ConfirmDeleteModal` pattern), not an automatic cleanup.

**No other feature is currently queued.** If you're picking this project up cold, the natural
next question to ask the user is simply "what would you like to work on next?" rather than
assuming — the last substantial feature shipped was multi-facility bookings, and everything
since then in this file reflects the state as of that point.
