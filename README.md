# Butler Property Management (BPM) — Owner's Operations Guide

Welcome to your **Butler Property Management (`BPM`)** operations platform. This system connects tenant maintenance requests, field crew time & store receipts, client invoicing & owner settlements, rent & lease tracking, and crew payroll into one closed loop backed by your **`BPM Update`** Google Sheet — so **nothing falls through the cracks or goes unbilled**.

---

## Quick Links & First-Time Setup

| Resource | Purpose | Link |
| :--- | :--- | :--- |
| **BPM Web App** | Owner dashboard, Crew portal, Billing, Rent Roll & Payroll | [Open BPM Web App](https://script.google.com/macros/s/AKfycbzzsVEAFGBOB14VA0fcuTDs5LNS6tnvxwckNY9LtcNuta1IfBUFpUfZrutRFy8RCimC/exec) |
| **Tenant Maintenance Portal** | Public no-login repair request link for residents (`?view=request`) | [Open Tenant Request Form](https://script.google.com/macros/s/AKfycbzzsVEAFGBOB14VA0fcuTDs5LNS6tnvxwckNY9LtcNuta1IfBUFpUfZrutRFy8RCimC/exec?view=request) |
| **BPM Update Google Sheet** | Live database behind the app (20 structured tabs) | [Open BPM Update Sheet](https://docs.google.com/spreadsheets/d/1bELSUI7xnIyqNxUPMOFTU8uRc9sGsIhkEre6LkVenmI/edit) |
| **Apps Script Editor** | Script hosting, permissions & deployments | [Open Apps Script Project](https://script.google.com/u/0/home/projects/1kQ-N0lL7Y883BUzayVnQ52ffLsKtag24PgN-r5bA1PodLilZwA7XsdBZ/edit) |

### First-Time Owner Checklist (2 Minutes)
1. **Authorize Script Permissions (One Time):** Open the **Apps Script Project** link above, select `setupDatabase` in the top dropdown, click **Run**, and approve **Review Permissions -> Allow** under `BGross19@gmail.com`. (This lets tenants submit requests via `?view=request` and lets the app send emails and create PDFs on your behalf.)
2. **Sign In & Set Your Quick PIN:** Open the **BPM Web App** link, switch to **Email Code (OTP)**, enter `BGross19@gmail.com`, and verify with the 6-digit email code. Set a 4–8 digit **Quick PIN** — your browser stays signed in for **30 days**.
3. **Configure Payment Link & Weekly Automation:** Inside **Settings -> Company & Automation**, add your Venmo/Zelle/Stripe **Payment Link** (`PAYMENT_URL`) if you have one, update your phone/address, click **Save Settings**, and click **Enable Automated Triggers**.

---

## How Access & Team Roles Work

- **Owner / Admin (`BGross19@gmail.com`)**: Full access to all 7 tabs (`Dashboard`, `Work Orders`, `Time & Expenses`, `Billing & Invoices`, `Rentals & Leases`, `Payroll`, `Settings`).
- **Field Crew (e.g., Jason R.)**: Automatically sees a focused mobile view with only **My Work Orders**, **Log Time / Receipt & History**, and **My Pay Summaries / Paystubs**.
  - Crew **never** sees owner billing rates, rent rolls, company revenue, or other crew members' pay.
  - Crew can edit their own **Pending / Unbilled** logs or receipts, but **cannot** approve their own payroll or change client billing rates.
- **Tenants (`?view=request`)**: Public link with zero login required. Inside the app header, click **Tenant Link** any time to copy the request URL to text, email, or add to welcome packets.

---

## Day-to-Day & Monthly Workflows

### 1. Dashboard — Your "Nothing Missed" Command Center
Every time you open the app, the top banners and 6 KPI cards immediately surface any revenue or maintenance gaps:
- **Red / Amber Alerts**: Highlights **Emergency Work Orders**, **Completed Work Orders with 0 logged hours/expenses** (work finished in the field that has not been logged for billing), **Unbilled Backlog**, **Pending Approvals**, and **Overdue Invoices**.
- **6-Month Financial Activity Chart**: Compares **Invoices Billed**, **Rent Collected**, and **Payroll Paid** side-by-side.

### 2. Work Orders & Tenant Request Intake
1. **Tenant Submission (`?view=request`)**: Tenants choose their property, unit/room, contact info, **Permission to Enter** (`Lockbox/key`, `Knock first`, or `Call/text first`), urgency (**Normal**, **Urgent**, **Emergency**), description, and up to **3 photos** (automatically compressed and saved in your Google Drive).
   - **Emergency & Urgent Alerts**: Immediately email your Dispatch Email (`[EMERGENCY]` / `[URGENT]`) with full tenant contact info and photo links, and issue a ticket number (`WO-1001`, etc.).
2. **Admin Dispatch**: Under **Work Orders**, click **+ New Work Order** or **Edit** on any ticket to assign a crew member, set a due date, or set a **Not-To-Exceed (NTE) Limit**.
3. **Closed-Loop Completion**: When crew finish a ticket and log time/receipts directly from it (**Log Time / Receipt**), the ticket automatically moves cleanly from `In Progress` -> `Completed` -> `Billed` once invoiced.

### 3. Time & Material Receipts (Including 0-Hour Store Runs)
Under **+ Log Work / Receipt** or **Time & Expenses**, you and your crew can record three types of entries:
- **Time + Materials**: Logs hours worked plus optional hardware store receipts.
- **0-Hr Expense Only**: Built specifically for Home Depot/Menards/Lowes runs or contractor receipts **without** creating fake `0`-hour labor rows.
- **Flat-Rate Job**: Overrides hourly billing with an agreed flat client charge (for example, a `$150` flat drain snake or turnover fee).

Key billing and reimbursement controls on every entry:
- **Paid By**:
  - **Company Account / Card**: Client-billable material, **not** reimbursed to crew in payroll.
  - **Employee Out-of-Pocket**: Client-billable material **and** queued into the employee's next payroll run once approved.
- **Charge Target**: Defaults to **Owner** (or switch to **Tenant** / **BPM Internal**). Only `BPM Internal` is excluded from client invoicing so billable work never gets skipped accidentally.
- **Approvals**: Check rows in the **Work & Expense Ledger** and click **Approve Selected** before running payroll.

### 4. Billing, Client Rates, Owner Settlements & Payments

#### Separate Crew Pay vs. 5-Tier Client Billing Rates
Your team's hourly pay (`$25/hr`, `$35/hr`, etc.) is completely separate from what owners are billed. For every labor log, BPM automatically resolves the client billing rate in strict priority order:
1. **Custom Flat / Job Override** on that specific log (highest priority)
2. **Property + Task Override** (e.g., special HVAC rate for a specific house)
3. **Owner + Task / Owner Default Override**
4. **Employee + Task / Employee Rate Override**
5. **Task Default Rate** (set in `Tasks & Rates`, `$65–$85/hr`)
6. **Global Fallback Rate** (`$65/hr` in Settings)

#### Generating Invoices & Monthly Owner Settlements
Under **Billing & Invoices**:
- **Unbilled Pipeline** groups every unbilled labor hour and receipt cleanly by **Owner & Property**, with warning badges for any **Completed Work Order missing time/expenses**.
- Click **Invoice Property** (single house), **Invoice Owner (Consolidated)** (all unbilled properties for an owner), or **Create Monthly Settlement**:
  - **Direct Invoice**: `Labor + Materials (+ optional Adjustments)`
  - **Monthly Owner Settlement** (e.g., co-owned properties like `710 42nd` w/ Heath): `Rent Collected - Management Fee % - Maintenance Expenses (+ optional Adjustments)` = exact net statement or payout balance.
- Clicking **Generate PDF & Save** creates a formatted PDF in Drive, locks those logs/expenses as `Invoiced` (preventing duplicate billing), and opens the preview where you can **Download PDF** or **Email Owner / Client** (with your payment link & instructions included).

#### Lump-Sum Owner Payments & Safe Voiding
- **Record Lump-Sum Owner Payment**: If an owner sends one combined ACH, Zelle, or check covering multiple invoices, select the owner and enter the payment amount — click **Auto-Distribute** (or type custom splits) to pay down multiple invoices at once (`Paid` / `Partial`).
- **Voiding an Invoice**: If you need to redo an invoice, click **Void** on the invoice row. BPM reverses the invoice safely and **automatically returns all linked labor logs, material expenses, and work orders back to `Unbilled` / `Completed`**.

### 5. Rentals, Leases & Academic / Calendar Rent Roll
Under **Rentals & Leases**, your **`BPM Rent` `26-27` leases** (`7` leases, `31` tenants, deposits, waivers, and prepayments) are live in an interactive matrix:
- **Academic (`May–Apr`) vs. Calendar (`Jan–Dec`) Toggle**: Switch cycles anytime without splitting leases — leases whose active dates overlap a calendar or academic window appear automatically with `Outside` shaded cleanly outside their lease start/end.
- **Color-Coded Rent Cells**: Click **any month cell** on any tenant row to take quick actions:
  - **Record Rent Payment** (date, method, note; partial payments show exact remaining balance).
  - **Upfront Prepay (Semester / Full Year)**: Select covered months (for example, Aug–Dec or May–Apr) or enter a lump sum to roll credit **FIFO** automatically across upcoming open months.
  - **Waive Month (`$0 Waived`)** (e.g., June concession or `"FREE per Jeff"`).
  - **Add Late Fee / Charge** or **Credit Adjustment**.
- **Lease Renewals & Move-Out Deposit Settlements**:
  - Click **Renew** on an expiring lease to pre-fill next cycle's lease + active tenants in seconds.
  - Click **Deposit Settlement** on a tenant row to itemize deposit held minus cleaning/damage/unpaid rent deductions and generate a printable **Move-Out Security Deposit Settlement PDF**.

### 6. Crew Payroll & Out-of-Pocket Reimbursements
Under **Payroll**:
1. Pick a date range and click **Run Payroll** for a crew member.
2. BPM strictly pulls **only `Approved` labor logs and `Approved` `Employee Out-of-Pocket` receipts** (`Pending` entries stay safely held back until reviewed).
3. Add any optional bonus/advance adjustment, choose payment method & reference number, and click **Complete Payroll Run**.
4. BPM marks those entries `Paid`, saves the run in **Completed Payroll Runs**, and generates a downloadable **Paystub PDF** visible both to you and in the crew member's portal.

### 7. Settings, Bulk Import & Automated Weekly Safeguards
Inside **Settings**:
- **Company & Automation**: Edit business info, default grace day (`5th`), late fee (`$50`), default management fee (`8%`), material markup (`%`), invoice due terms (`15 days`), and `PAYMENT_URL`.
  - **Monday 7:00 AM Digest Email**: Automatically emails your Dispatch Email a Monday morning checklist of any Emergency Work Orders, Completed-Unbilled tickets, Pending Approvals, Unbilled Revenue Ready, Overdue Invoices, and Past-Due Rent.
  - **Sunday 2:00 AM Safe Archive**: Moves **only closed, billed, and paid** records older than **90 days** into the `archive_*` sheets to keep the app fast forever — open or unbilled items are **never** archived.
- **Owners / Properties / Tasks & Rates / Team & PINs**: Add or edit owners (`Direct Invoice` vs. `Monthly Settlement`), properties (`Aliases`, gate/lockbox notes), standard tasks, custom **Scope Rate Overrides**, and team members / reset PINs.
- **Bulk Import**: Paste CSV/spreadsheet rows or upload a `.csv` file to batch-add Properties, Tasks, Owners, or Employees with a live validation preview before saving.

---

## Best Practices for Your `BPM Update` Google Sheet

- **Use the Web App for daily edits** so IDs, balances, invoice links, and permissions stay synchronized automatically.
- **Do not rename Sheet tabs or Row 1 column headers** — the app reads every tab by exact column header name.
- **Google Drive Folders**: On first receipt upload or PDF generation, BPM automatically creates a **`BPM Documents`** folder in your Google Drive containing **`Receipts & Intake Photos`** and **`Invoices & Settlements`** subfolders and stores their folder IDs in `sys_Settings` (`RECEIPTS_FOLDER_ID` and `INVOICES_FOLDER_ID`).
- **Legacy Workbooks**: Your previous sheets (**BPM Rent**, **BPM MAINTENANCE**, **BPM 2025 Statements**, **bpm owner fees**) were used read-only for data setup and were never altered.
