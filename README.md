\# CheckMarq Payroll Production



> \*\*Production Apps Script repository for CheckMarq payroll, invoice, and audit automation.\*\*



This repository contains the full Google Apps Script codebase that powers the \*\*CheckMarq Payroll Production\*\* spreadsheet. It is treated as \*\*production code\*\* and follows a `dev → main` workflow with branch protections enabled.



---



\## ⚠️ Important Production Notes



\* \*\*This repo is production-backed.\*\* Changes pushed to Apps Script affect live payroll behavior.

\* \*\*Do not edit code directly in the Apps Script UI\*\* unless explicitly necessary.

\* All changes should be:



&nbsp; 1. Made locally

&nbsp; 2. Pushed via `clasp push`

&nbsp; 3. Committed to GitHub (preferably on `dev`)

&nbsp; 4. Merged to `main` via Pull Request



---



\## 🔗 Bound Spreadsheet



This Apps Script project is \*\*container-bound\*\* to the CheckMarq Payroll Production Google Sheet.



\* The Sheet defines:



&nbsp; \* Data structure

&nbsp; \* Configuration tabs

&nbsp; \* Source visit data

&nbsp; \* Output destinations



> The spreadsheet is the \*\*runtime host\*\*; this repo is the \*\*source of truth\*\* for logic.



---



\## 🗂️ File Overview



\### Core Orchestration



\* \*\*`Code.js`\*\*

&nbsp; Entry point and menu wiring. Defines `onOpen()` menus and routes user actions to the appropriate modules.



\* \*\*`Master\_Rebuild.js`\*\*

&nbsp; High-level orchestration logic used for full rebuilds or coordinated recalculations.



---



\### Payroll \& Invoice Logic



\* \*\*`Payroll\_Build.js`\*\*

&nbsp; Core payroll calculation logic. Handles aggregation, rate application, and payroll row generation.



\* \*\*`Invoice\_Build.js`\*\*

&nbsp; Builds invoice-side calculations derived from payroll data and visit structures.



\* \*\*`Audit\_Build.js`\*\*

&nbsp; Generates audit checks, summaries, and validation outputs used to verify payroll integrity.



---



\### PDF Generation



\* \*\*`PayrollPDF.js`\*\*

&nbsp; Responsible for generating payroll PDF outputs.



\* \*\*`InvoicePDF.js`\*\*

&nbsp; Responsible for generating invoice PDF outputs.



> PDF generation depends on spreadsheet layout and named ranges. Changes should be tested carefully.



---



\### Email Automation



\* \*\*`Email\_Build.js`\*\*

&nbsp; Constructs email payloads and routing logic.



\* \*\*`Email\_Drafts.js`\*\*

&nbsp; Creates draft emails for review before sending.



---



\### Utilities \& Helpers



\* \*\*`Sheet\_Helpers.js`\*\*

&nbsp; Shared utilities for reading/writing sheets, ranges, formatting, and guards.



\* \*\*`FinalPaidVisits.js`\*\*

&nbsp; Final export / snapshot logic for paid visit records.



\* \*\*`Reset\_PayPeriod.js`\*\*

&nbsp; Controlled reset logic for starting a new pay period.



---



\### Configuration



\* \*\*`appsscript.json`\*\*

&nbsp; Apps Script manifest. Defines runtime, scopes, and advanced services.



---



\## 🌱 Branching \& Workflow



\* \*\*`main`\*\*

&nbsp; Production branch. Protected. PRs required.



\* \*\*`dev`\*\*

&nbsp; Development branch. All changes should start here.



\### Standard workflow



```bash

git checkout dev

\# make changes

clasp push



git add .

git commit -m "Describe change"

git push

```



Then open a PR:



\* \*\*Base:\*\* `main`

\* \*\*Compare:\*\* `dev`



---



\## 🔐 Safety \& Guardrails



\* Branch protection enforced on `main`

\* Force-pushes blocked

\* PR approval required

\* Conversation resolution required



These guardrails exist because \*\*this code impacts payroll and billing\*\*.



---



\## 🧪 Testing Philosophy



There is no separate automated test harness.



Testing is performed by:



\* Using known historical pay periods

\* Verifying row counts and totals

\* Reviewing audit outputs

\* Validating PDFs before distribution



> A future enhancement may include a dedicated \*\*staging spreadsheet\*\* bound to a clone of this script.



---



\## 📌 Common Gotchas



\* Sheet names are case-sensitive

\* Many calculations assume sorted and filtered visit data

\* Menu actions assume specific tab visibility

\* Reset operations are \*\*destructive\*\* by design



Proceed carefully.



---



\## 📄 Ownership \& Maintenance



\* \*\*Owner:\*\* CheckMarq

\* \*\*Environment:\*\* Production

\* \*\*Change discipline:\*\* Conservative



When in doubt:



\* Branch

\* Test

\* Review

\* Merge deliberately



---



\## 🚧 Future Improvements (Optional)



\* Staging spreadsheet + script

\* Dry-run / audit-only mode

\* Modular folder structure

\* Automated snapshot exports

\* CI checks for accidental scope changes



---



\*\*This repository is intentionally explicit and guarded.\*\*

It exists to protect payroll integrity while still allowing controlled evolution.



