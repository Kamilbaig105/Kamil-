# EMS Energy Dashboard

Static dashboard (GitHub Pages) built from the daily KWH report.

## One-time setup
1. Create a GitHub repository (e.g. `ems`). Note: on a free account the site is public.
2. Upload everything from this folder to the repository root (keep the `.github` folder).
3. Go to **Settings → Pages**, choose **Deploy from a branch**, branch `main`, folder `/ (root)`, and save.
4. After a minute the site is live at `https://<your-username>.github.io/<repo-name>/`.

## Daily update – automatic
1. Save your updated `KWH_REPORT.xlsx` in Excel (so formula results are stored).
2. In the repository click **Add file → Upload files**, upload `KWH_REPORT.xlsx` (same name), and commit.
3. The "Update dashboard data" action rebuilds `data.js`; the site refreshes within a couple of minutes.
   Do not delete `data.js` – it is the file the dashboard reads.

## Daily update – manual
```
pip install openpyxl pandas numpy
python build_data.py KWH_REPORT.xlsx
```
Then upload the new `data.js` to the repository.

## Privacy
The repository (`KWH_REPORT.xlsx`, `data.js`) and the site show your plant's consumption data.
For a private site use a paid GitHub plan, or keep the repository private and run the manual route without uploading the Excel file.
