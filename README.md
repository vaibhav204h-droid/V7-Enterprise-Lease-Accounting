# Ind AS 116 Lease Calculation Software V7.3

## V7.3 error remediation release
This build preserves the existing V6/V6.1/V7 architecture, Streamlit workflow, database model, calculation entry points and report structure. Changes are surgical additions/corrections driven by the supplied error workbook.

### V7.3 repaired controls
- Pre-commencement leases are kept outside recognised closing balances, ROU, deposits, restoration provisions and journal entries while remaining visible in contractual maturity analysis.
- Initial measurement distinguishes unpaid lease liability from advance/prepaid cash.
- Escalation engine supports Fixed %, Fixed Amount, CPI, WPI and Other Index events with boundary checks and cumulative event application.
- Current/non-current split reconciles to reporting-date liability using a configurable lease-level tolerance.
- Security deposits support FV/EIR carrying amounts and dynamic ECL using PD × LGD / horizon or configured loss-rate methodology.
- Restoration provisions support discounted cost estimates, unwinding and final-date true-up, with date-gated recognition.
- Modifications preserve the next contractual payment date instead of inventing a payment at the modification effective date; reassessed lease term is used for ROU depreciation.
- FX reporting exposes functional balances and optional FX gain/loss working based on prior/initial rates.
- Journal entries are event-gated and dynamically validated for balance.
- Input validation now reports field-specific reasons for failure and validates rates, dates, escalation events, deposits, FX and tolerance.
- Calculation snapshots are stored for traceability, alongside existing audit-trail records.
- Excel report controls are formula-driven and independently compare engine calculations with a separate validation engine.
- New import template: `templates/MNC_Lease_Import_Template_V7_3.xlsx`.

### Validation
- Existing deterministic V7.0/V7.1/V7.2 regression suites pass.
- New V7.3 remediation suite: **20 tests passed**.
- Full randomized suites remain available in the existing test suite; they intentionally contain 10,000 deterministic scenarios and are materially heavier than the fast CI regression set.

## Run
`streamlit run app_v5.py`

For Windows, use `start_windows.bat`.

## Windows startup
Double-click `start_windows.bat`. The first launch creates a fresh local `.venv` for the current computer and installs the packages once. Later launches reuse that environment. Do not copy a `.venv` from another computer because Windows virtual environments contain machine-specific paths.

For diagnostics, run `start_windows_debug.bat`; it keeps the console open if startup fails so the error can be read.

## Important V7.3 data-source control
The final distribution does not bundle the previous 10-lease demo database or legacy report packs. Import your current `MNC_Lease_Import_Template_V7_3` file through **Import Existing Leases → Validate Upload → Confirm & Commit Import** before running **Calculate Portfolio**. V7.3 blocks calculation if the active Lease Master IDs differ from the last committed import batch.


## V7.3 targeted patch — 14-Sep-2026

This patch implements the supplied **Lease Accounting Software – Required Changes & Improvements** as targeted corrections; existing sheet names, input fields and core architecture are retained.

### Implemented
- Payment dates are generated from the original first-payment anchor (`anchor + n × frequency`), preventing 29/30/31 month-end drift.
- A single `lease_term_months()` calculation is used by ROU depreciation and exposed in the report; L-007 style 08-Sep-to-07-Sep terms are treated consistently as 156 monthly periods.
- Final ROU depreciation true-up forces the closing ROU to exactly zero at lease end.
- Independent validation reconstructs payment dates and assessed expiry independently rather than relying on the engine payment schedule.
- Independent validation now compares independent initial liability, reporting liability, current liability and initial ROU.
- Assessed effective expiry explicitly considers reasonably-certain extension/termination options; contractual expiry is retained separately.
- Future-effective modifications are excluded from reporting-date measurement while still being validated as metadata.
- Escalation schedule validation is aligned to assessed effective expiry rather than prematurely rejecting events in an assessed extension period.
- Validation output supports PASS / WARNING / FAIL, while retaining the existing zero-liability active-lease control.
- Master report now exposes contractual/assessed/modification-adjusted expiry, lease-term months and term-assessment reasoning.
- Existing actual-day Ind AS 109 deposit methodology, escalation engine, current/non-current methodology and workbook/report architecture are preserved.

### Verification
Targeted regression suite: **43 tests passed**.
