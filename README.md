# Factsheets: fund managers and AUM

This folder reads each SIF's monthly factsheet PDF and pulls out, for every strategy in `scheme_master_mapping/sif_scheme_master.xlsx`:

- its fund managers and the date each started managing it;
- its month-end AUM and its monthly average AUM;
- the date the AUM is as on.

It is a port of `historical-mf/excel/factsheets`. The docling conversion, the LangGraph agent, the checks, the code rules and the caching all work the same way.

## Commands

Run everything from the repo root, inside the project's virtual environment.

```powershell
.venv\Scripts\activate
pip install -r requirements.txt     # first time only
```

### Full run, every SIF

Run these in order after adding files to `data/monthly_portfolio/` and `data/factsheets/` (see [Input](#input)):

```powershell
python factsheets\factsheet_portfolio.py    # portfolios of SIFs without a disclosure, read from their factsheet
python sif_parser.py                        # holdings -> output/holdings.json, output/categories/
python factsheets\factsheet_parser.py       # managers and AUM -> factsheets/output/factsheet_fund_managers.csv
python factsheets\manager_exits.py          # exit dates printed on the pages
python merge_portfolio_factsheet.py         # holdings + factsheets -> output/merged/merged.json, coverage.csv
```

### Only some SIFs

`--amc` matches any SIF folder whose name contains the text, and can be repeated. Give `sif_parser.py` its own output folder: a filtered run into the default `output/` replaces the full `holdings.json` and `categories/` with partial ones.

This is the run for the SIFs with full data as of the Sep 2026 factsheets (Diviniti since its May 2026 factsheet was added, Apex since Nov 2026 although its factsheets cover only the Hybrid fund from June 2026; Arthaya, DynaSIF, Infinity, Sapphire, WSIF and Altiva since Nov 2026 (Altiva has no factsheet for Hybrid OCT 2025 or Equity Long-Short SEP 2026), see [Portfolios from the factsheet](#portfolios-from-the-factsheet)):

```powershell
python factsheets\factsheet_portfolio.py
python sif_parser.py data\monthly_portfolio output\full_data --amc iSIF --amc Magnum --amc Prism --amc RedHex --amc Summit --amc Arudha --amc Platinum --amc Diviniti --amc Apex --amc Arthaya --amc DynaSIF --amc Infinity --amc Sapphire --amc WSIF --amc Altiva
python factsheets\factsheet_parser.py --amc iSIF --amc Magnum --amc Prism --amc RedHex --amc Summit --amc Arudha --amc Platinum --amc Diviniti --amc Apex --amc Arthaya --amc DynaSIF --amc Infinity --amc Sapphire --amc WSIF --amc Altiva
python merge_portfolio_factsheet.py --holdings output\full_data\holdings.json --output output\full_data\merged.json --amc iSIF --amc Magnum --amc Prism --amc RedHex --amc Summit --amc Arudha --amc Platinum --amc Diviniti --amc Apex --amc Arthaya --amc DynaSIF --amc Infinity --amc Sapphire --amc WSIF --amc Altiva
```

Run the last three with the same `--amc` list: the merge only keeps SIFs that are in `output\full_data\holdings.json`, or in the factsheets for months with no holdings (Infinity, WSIF Jun-Sep 2026).

### Portfolios from the factsheet

Arthaya, DynaSIF, Sapphire and WSIF publish no portfolio disclosure, but print every holding in the factsheet. `factsheets\factsheet_portfolio.py` reads it and writes `<Scheme> (from factsheet).xlsx` into the SIF's `data/monthly_portfolio/<YEAR>/<MON>/<SIF>/` folder, laid out like a disclosure, which `sif_parser.py` then parses like any other:

1. It finds each strategy's pages and the model reads the portfolio table (`portfolio_extractor.py`). The answer is checked against the printed subtotals and the grand total, and the model fixes it or confirms it.
2. Code corrects slips the model is known to make: a total returned as a holding, a short shown as negative equity, or a cash line listed only as a total.
3. Each company name is matched to its ISIN, symbol and market cap in the AMFI lists (`instrument_match.py`). Both sides are standardised first, so `The Indian Hotels Company Limited` matches `Indian Hotels Co. Ltd`. Matched rows are written under the AMFI name.
4. Market value is % to net assets × the month-end AUM printed with the portfolio. The rounding gap to 100% goes to net receivables, at most 0.004 % points per row.

No workbook is written when the factsheet prints only the top holdings (WSIF from June 2026) or the rows do not add up. Infinity is not read this way: its factsheet prints the debt half as category weights only. A real disclosure added later wins: generated workbooks are parsed last.

```powershell
python factsheets\factsheet_portfolio.py --year 2026 --month SEP --amc Dyna   # one SIF month
python factsheets\factsheet_portfolio.py --retry-unresolved                    # re-ask the model where checks still fail
```

Each run writes `factsheets/output/<YEAR>/<MON>/<SIF>/portfolio.md`, which has the totals, the corrections and every row with its printed and matched name. `factsheets/output/factsheet_portfolio_matches.csv` has every match.

Names matched `fuzzy` or `unmatched` need a look. Fix them in `factsheets/factsheet_instrument_alias.csv` (`fund_house, printed_name, isin, note`) and re-run; answers are cached, so no model call is made. Listings newer than the latest half-year AMFI list come from historical-mf's AMFI list in `market_cap_mapping/supplementary/`, which `sif_parser.py` also uses for their market cap.

It writes `merged.json` and `coverage_selected.csv` to `output/full_data/`. The factsheet step still rebuilds `factsheets/output/factsheet_fund_managers.csv` with every SIF processed so far; the merge's `--amc` does the filtering.

### Narrower runs

```powershell
python sif_parser.py data\monthly_portfolio output\sep_2026 --year 2026 --month SEP
python factsheets\factsheet_parser.py --year 2026 --month SEP --amc Magnum
python factsheets\factsheet_parser.py --category "Hybrid Long-Short"
python factsheets\factsheet_parser.py --force          # redo docling and the model
python factsheets\factsheet_parser.py --force-llm      # redo the model only
python merge_portfolio_factsheet.py --year 2026                # writes merged_2026.json, keeps merged.json
python merge_portfolio_factsheet.py --year 2026 --month SEP    # writes merged_2026_september.json
```

### Checks

```powershell
python factsheets\rules_eval.py                        # score the code rules against the model's answers
python ingestion-pipeline\neo4j_parser.py --dry-run    # AUM per snapshot and manager tenures, no Neo4j connection
```

### Load into Neo4j

These write to the database:

```powershell
python ingestion-pipeline\neo4j_parser.py                                         # category files, AUM and managers from the factsheet CSV
python ingestion-pipeline\neo4j_parser.py --input output\merged\merged.json       # from the merged file; also creates factsheet-only months
python ingestion-pipeline\neo4j_parser.py --input output\full_data\merged.json    # only the full-data SIFs' snapshots
python ingestion-pipeline\neo4j_parser.py --managers-only                         # only SIF_FUND_MANAGER and SIF_MANAGES
```

`--input` also limits the managers: they are loaded only for the SIFs in the input files, so the `full_data` load above adds no SIF_FUND nodes for SIFs without full data. `--managers-only` reloads just the managers of the SIFs in `--input`. Add `--all-managers` to load managers for every SIF in `factsheet_fund_managers.csv`.

## Input

Input files live under `data/`, sorted by `<YEAR>/<MON>/`:

```
data/
  monthly_portfolio/<YEAR>/<MON>/<SIF>/*.xls*        portfolio disclosures, read by sif_parser.py
  factsheets/<YEAR>/<MON>/<Name> Factsheet.pdf       factsheets, read by factsheets/factsheet_parser.py
```

- **Disclosures** go in a `<SIF>` folder named exactly as the SIF's `fund_house` in the scheme master, for example `Magnum SIF` or `iSIF`.
- **Factsheets** sit directly in the month folder, one PDF per SIF, named `<Name> Factsheet.pdf`. `<Name>` is the `fund_house`, or the start of it if only one SIF matches: `Dyna Factsheet.pdf` is DynaSIF, `Magnum Factsheet.pdf` is Magnum SIF. A PDF named any other way is reported and skipped (`factsheet_house()` in `factsheet_parser.py`).
- Altiva publishes one PDF per strategy. Merge them into one `Altiva Factsheet.pdf` (Hybrid first); the per-strategy originals up to Sep 2026 are in `archive/altiva_factsheets_per_strategy/`.

`data/factsheets/` holds the input PDFs; `factsheets/` at the repo root is the code, and `factsheets/output/<YEAR>/<MON>/<SIF>/` its results. `data/` and `archive/` are not in git.

## How it works, and how it differs from MF

1. **Pages.** MF finds each fund through the factsheet's index. Most SIF factsheets have no index, so every page is scored for each strategy instead, from the PDF text layer. A page scores for:
   - printing the strategy's name, with or without the brand. Apex's details page prints only `Name Hybrid Long Short Fund`;
   - printing the name as a title line;
   - carrying fund-manager and AUM labels;
   - printing few other strategies' titles.

   The best page and the following pages that still name the strategy go to pass 2. A strategy is reported as missing in `index.md` in two cases:
   - no page names it, because it has not launched yet or is printed under another name;
   - it is named without any fund-manager or AUM details, as on Apex's July cover.
2. **Funds.** This works as in MF. Docling converts the strategies' pages, then an agent per strategy reads them, reads or searches more pages if needed, and submits an answer. Plain-code checks review the answer before it is accepted. The code rules in `rules/` answer only when the model can't: no credentials, an API error, or no answer submitted.
3. **Parent AMC.** Each SIF is run by a mutual fund house: Altiva by Edelweiss, Magnum by SBI, iSIF by ICICI Prudential, RedHex by HSBC, and so on. The house's name is passed to the rules along with the SIF's (`PARENT_AMC` in `factsheet_parser.py`), so MF's rules for that house's layout also run on its SIF. Rules for layouts only SIFs print are in `rules/sif.py`.
4. **Unreadable text layer.** Magnum DEC 2025 prints its labels in a font whose text layer shifts every letter by 29 (`0DJQXP` for `Magnum`). The decoded words are added to the page's text, and the page is also sent to the model as a picture.

## Output (`factsheets/output/`)

| File | What it holds |
|---|---|
| `<YEAR>/<MON>/<SIF>/index.md` | Each PDF, the pages each strategy was found on, and the strategies not found |
| `<YEAR>/<MON>/<SIF>/funds.md` | Managers and AUM per strategy, for review |
| `<YEAR>/<MON>/<SIF>/fund_details.csv` | The same, one row per fund manager |
| `<YEAR>/<MON>/<SIF>/funds/*.md` | Docling markdown of each strategy's pages |
| `<YEAR>/<MON>/<SIF>/pages/`, `llm/` | Docling cache per page, and the model's answer per strategy |
| `factsheet_fund_managers.csv` | Every factsheet, one row per fund manager |
| `factsheet_manager_exits.csv` | Exit dates printed on the pages (`manager_exits.py`) |

`factsheet_fund_managers.csv` has these columns:

- `scheme_code` (`SIF-13`), `scheme_name` and `category`, as in `sif_parser.py`'s output.
- `as_on_date` and `portfolio_month`, read from the pages rather than the folder. Join on `scheme_code` + `portfolio_month` to the holdings.
- `aum_month_end_cr` and `aum_monthly_avg_cr`, in Rs crore.
- `aum_check`, which flags a missing AUM, or month-end and average AUM more than about 30% apart. New strategies often trip it because their average covers only part of the month.
- `fund_manager`, `managing_since`, `managing_since_raw` and `manager_role`.
- `notes`, with anything the model or the checks found ambiguous.

Docling output and the model's answers are cached, so a re-run makes no API call for unchanged pages. A PDF replaced under the same file name is detected by its content and read again. `--force` redoes both, and `--force-llm` redoes only the model step. Bump `PROMPT_VERSION` in `fund_extractor.py` after changing a prompt.

## Settings

The model is the GPT-5 deployment on Azure OpenAI. It reads `API_KEY`, `endpoint` and `deployment_name` from the repo-root `.env`, the same settings as historical-mf. Without them, the run carries on with the code rules alone and notes "code rules only" on any fund they could not finish.

`FACTSHEET_LLM_RPM` (default 60) caps requests per minute, and `FACTSHEET_LLM_REASONING_EFFORT` (default `low`) sets GPT-5's reasoning effort.

Docling's models load offline (`HF_HUB_OFFLINE=1`). On a new machine, set `HF_HUB_OFFLINE=0` once to download them.

## Adding data

| Task | What to do |
|---|---|
| New month | Put the portfolio file in `data/monthly_portfolio/<YEAR>/<MON>/<SIF>/` and the factsheet as `data/factsheets/<YEAR>/<MON>/<Name> Factsheet.pdf`, then do the full run above |
| Disclosure printed under a strategy's old name | Add it to `FORMER_NAMES` in `scheme_master_mapping/build_sif_scheme_master.py` and to the `former_names` sheet of `sif_scheme_master.xlsx` (`former_name, scheme_code, scheme_name, note`), then re-run `sif_parser.py` |
| New SIF | Rebuild the scheme master (`build_sif_scheme_master.py`); add its parent house to `PARENT_AMC` if its layout is that house's |
| Merge holdings and factsheets into one JSON | Run `python merge_portfolio_factsheet.py` (repo root), optionally with `--year 2026` and/or `--month SEP`. It writes `output/merged/merged.json`: each fund gets `fund_manager` (its tenures), and each month gets `AUM` (month-end, Rs crore) and `portfolio` (the holdings). Months are joined on SIF code and the month the data is as on. It also writes `coverage.csv`, showing what each fund-month has. `ingestion-pipeline/neo4j_parser.py --input output/merged/merged.json` loads it. |
| Strategy printed under another name | Add a row to `factsheet_scheme_alias.csv` (`fund_house, factsheet_name, scheme_code`) |
| Manager printed under several spellings | Add a row to `factsheet_manager_alias.csv` (`fund_house, printed_name, manager_name`) |
| New layout for the rules | Add a `@rule(...)` function to `rules/sif.py` and check it with `rules_eval.py` |
