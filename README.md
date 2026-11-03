# TBML Reporting Portal – Channel Management Department

Collects the monthly TBML Key Risk Indicators (Circular No. 1996/2026, Annex I) **from every branch**, through
25 regional Forms, validates them, stores them month by month and produces the consolidated Annex I report
(Present / Previous / YTD) as Excel and PDF, with region-level and branch-level views.

## How the process works

* **25 Microsoft Forms, one per region.** Each Form lists only that region's branches. The RHO is a collaborator on
  their own Form, so the RHO sees only their own branches' responses. Branches use the same link every month.
* **Branches** submit by the 5th. **The RHO** opens *Responses → Open in Excel*, checks the branches (asks a branch to
  resubmit, or corrects the figure in Excel), and e-mails the file to Channel Management by the 10th.
* **You** upload all 25 Excel files at once in the portal; it validates, adds up the branches (the RHO does not
  consolidate), and produces the Annex I report for IBD and Compliance by the 15th.

## 1. Install

**Windows (DGM's PC)** – install Python 3.11/3.12 (64-bit, tick *Add to PATH*), copy the `tbml_portal` folder,
double-click `Install_Portal.bat` (internet needed once), then use the desktop shortcut *TBML Reporting Portal*.
For an offline PC: run `pip download -r requirements.txt -d wheels` elsewhere and copy the `wheels` folder next to
`Install_Portal.bat`.

**Mac (for testing)** – in Terminal:
```
cd ~/Documents/tbml_portal
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
Practice with fake data first: `TBML_DATA_DIR=demo_data streamlit run app.py` (Windows: `Start_Demo_Portal.bat`),
then upload the 25 files in `sample_files\` together (they use your real region and branch codes and names, with random figures), choose **October 2026**.

Optional login: put a password in `settings.json` (`"password"`).

## 2. One-time set-up

1. Portal → **Master Data**: import `Regions_Master_Template.xlsx` (regions with RHO e-mail; branches with code and
   name), check, **Save master data**.
2. Build the Form from `MS_Form_Question_Sheet.xlsx` (sheet *Form Settings* has the step-by-step checklist, sheet
   *Per-region tracker* lets you tick off all 25).
3. For each region copy the branch list from **Master Data → Branch dropdown** into that region's Form.

## 3. Every month

| When | What |
|---|---|
| 1st | Reminder e-mail to the 25 RHOs (*Submission Tracker → Reminders* gives the addresses) |
| by 5th | Branches submit in their region's Form |
| 5th–10th | RHO downloads the Excel, checks/corrects, e-mails it with their confirmation |
| after 10th | **Upload & Validate** → select all files → review flags → type each RHO's name → **Save** |
| | **Submission Tracker** → chase outstanding branches; late files can be added later ("Update only the branches found") |
| by 15th | **Annex I Report** → sign-off names → **Prepare Excel & PDF** → send to IBD and Compliance |
| any time | **Combined Report** → choose From / To month (e.g. October to December) → **Prepare combined Excel** = one master file |

## Combining months

You never merge months by hand. Every month you upload is kept as its own tab in `TBML_Records.xlsx`, and the portal
reads all saved months automatically (that is how Previous and YTD are produced). For a multi-month master file use
**Combined Report**: pick the From and To months and download one Excel with these sheets: *Summary* (bank-wide by month
plus the period total), *Region by Month*, *Region Total*, *Branch Total* and *All Data* (every branch of every month in
one table, ready for filters or pivots). Period totals add the months together, except KRI-02 and KRI-03 (period-end
balances), which show the last month. Months with no saved data are listed as a warning and left out.

## 4. Where the data is

* `data/TBML_Records.xlsx` – one tab per month (`Oct-2026`, …), **one row per branch**, TOTAL row, plus `Legend` and `Log`.
  Close it in Excel before saving in the portal.
* `data/Regions_Master.xlsx` – regions and branches.
* `data/backups/` – automatic copy before every save (last 30 kept). Copy `data/` to a bank-approved location regularly.

## 5. Rules and assumptions

* **The form:** every branch enters, for each of the 18 questions, three figures: *Previous Month*, *Present Month* and
  *Year To Date* (exactly as Annex I). The portal reads the columns by the question wording, so do not rename questions.
* **Annex I = the sum of the branch figures** for Present, Previous and YTD. Branches that have not submitted are not in the totals.
* **Cross-checks (warnings, not blocks):** a branch's *Previous Month* is compared with what it reported as *Present* last
  month; YTD smaller than the present month is flagged (not for KRI-02/03, which are balances); branch name vs branch code,
  region code vs the branch's region, sender vs branch e-mail (if registered), late after the 5th / 10th.
* **Errors (row un-ticked):** unknown branch, not a number, negative, decimals in a count, unreadable month/year.
* Several submissions from one branch: the latest is kept, earlier ones are *Superseded*. Responses for other months in a file are ignored.
* **Combined Report:** period totals add the *Present* figures of the months; KRI-02/03 (balances) show the last month.
* The circular says the RHO amalgamates branch data; here the portal does it and the RHO reviews and confirms. Keep the RHO's
  e-mail as proof and get your DGM's approval for this variation.

## 6. Files

`app.py` (portal) · `tbml/config.py` (18 indicators, wording from the circular) · `tbml/core.py` (validation, storage,
calculations) · `tbml/exports.py` (Excel/PDF) · `settings.json` · `make_samples.py` (re-creates demo data).
