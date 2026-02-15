# Dispatch Portal: Current State and Minimal Change Plan

## Current State

- **Entry points:**
  - `onFormSubmit(e)` in `dispatch.js` (form-submit trigger)
  - `doGet(e)` in `webApp.js` (web app endpoint)
  - `onOpen()` in `dispatch.js` (Spreadsheet UI menu hook)
  - `updateAllCompanyPages()` in `dispatch.js` (manual/scheduled rebuild helper)
  - `cleanupOldDispatches()` and `deleteOldReplacedDispatches()` in `maintenance.js` (maintenance jobs)
  - No `doPost` and no `main` function are present.
- **Google Sheets usage / tab names:**
  - No sheet ID is hardcoded.
  - No sheet tab names are referenced.
  - Spreadsheet APIs are only used for UI (`SpreadsheetApp.getUi()`), not for reading/writing rows.
- **Drive IDs / template IDs currently expected:**
  - Dispatch archives root folder: `1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE`
  - Template doc IDs:
    - `1dS1Q0s55yBZHBGEUp4mHkIEroEWmtFPowRBDMLZQ84c` for `MCRC`/`CMPM`
    - `1S4mTIjbAGT-xTLAabY7wWzy92dkxKuri4FNHOxIV6aQ` for all other jobs
  - Truck archive folders:
    - `DT02`: `1k3LqwTyN4VprEZp4WjAwCutnTPDn76FR`
    - `RT03`: `1OHpftrKXETwNesEijjrlxO8WGzVEWjvL`
    - `RT12`: `1pm1UVd42sIAKGvnY6Hv7nqOBb3LPsefN`
    - `EXCL1`: `-1ixUnhC3UHhu6i5mFzPiFdjXUGpt_Y4tu`
    - `WAJA01`: `1VofxiA7vNfDUYMBXVAVUvxKI6DQuJPQE`
    - `WAJA03`: `1SakeHV9M9Ly_Yx1dY_Q6bTazXxW9Oekc`
    - `WMBA11`: `1kZZym6R-b33j7_2SQtGLw6eiIHzURQYn`
- **How the portal currently finds dispatch docs:**
  1. Form submit creates a new dispatch Google Doc from a template and stores it in a truck archive folder.
  2. It then rebuilds the company HTML page by scanning truck folders in the company folder under the archive root.
  3. The rebuild includes non-HTML files from the last 18 days, parses date/time from the file name, then categorizes into `Upcoming`, `Today`, and `Past`.
  4. It writes/updates `<COMPANY>_dispatch_list.html` into the company folder.
  5. Web app `doGet` takes `?company=...`, recursively finds `<company>_dispatch_list.html` under the archive root, and serves it.

## Proposed Changes (Smallest viable plan)

1. **Add two sheet-backed tables** (new tabs in the bound spreadsheet):
   - `Dispatches` tab: append one row per generated dispatch doc with key metadata + doc URL.
   - `Confirmations` tab: append one row each time a driver confirms receipt.
2. **Write Dispatches row during existing generation path**:
   - In `onFormSubmit`, after archive copy creation, append to `Dispatches` so existing doc generation remains unchanged.
3. **Add a lightweight confirmation endpoint**:
   - Add `doPost(e)` in `webApp.js` to receive confirmation payload (dispatch ID/url + truck + optional name).
   - Validate payload, append to `Confirmations`, and return JSON/text success.
4. **Add “Confirm Receipt” button in generated portal HTML**:
   - In `updateCompanyDispatchPage`, add a button next to each dispatch link.
   - Button calls a small inline script that POSTs to the web app endpoint.
5. **Keep existing listing logic intact**:
   - Do not alter folder scanning, 18-day filtering, naming conventions, or template selection.

## Exact Files to Edit

- `dispatch.js`
  - Add sheet write helper(s) and call to log to `Dispatches` in `onFormSubmit`.
  - Add “Confirm Receipt” button markup + minimal JS in generated HTML in `updateCompanyDispatchPage`.
- `webApp.js`
  - Add `doPost(e)` confirmation handler and sheet append logic to `Confirmations`.

No changes required in `maintenance.js` or `appsscript.json` for the minimal plan.
