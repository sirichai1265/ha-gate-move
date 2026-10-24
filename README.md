# H-A Gate Move

Daily summary of empty-container GATE IN / GATE OUT for the Heung-A yards, checked against the system export.

| Yard | Location | Report file |
|---|---|---|
| SMART BKK | BKK25 | `Report in-out H-A Final <yyyy-mm-dd> BKK21.xlsx` (GATE IN / GATE OUT) |
| BC | BKK27 | `FINAL_BC2_HeungA_Daily_Report_<dd_mm_yyyy>_*.xls` (GATE IN / GATE OUT) |
| HAST | LCH27 | `GATE IN <dd.mm.yyyy>.xls`, `GATE OUT  <dd.mm.yyyy>.xls` |
| CELLO | LCH28 | `FINAL CELLO H-A GATE IN_OUT <dd.mm.yyyy>.xlsx` (GATE IN REPORT / GATE OUT REPORT) |
| PW | LCHY5 | `FINAL_PW_HeungA_Daily_Report_<dd_mm_yyyy>_*.xls` (GATE IN / GATE OUT) |

System export: any `<m>-<d>-*.xls` (e.g. `10-2-GATE.xls`, `10-1-GAF2.xls`) with Location / Container No / Move / TPSZ / GateDate columns (a blank Location header is accepted). Files named with the next day's date (exports pulled the next morning) are also checked; only files whose rows are (almost) all of the run date count. All system files of the day are combined; for the same move the newest file wins.

## Rules

- TPSZ: 22GP, 42GP, 45GP, 22RE, 45RE, 22UT, 42UT, 22PC, 42PC (20'GP/DC → 22GP, 40'GP/DC → 42GP, 40'HC/HQ → 45GP, 20'RF/RE → 22RE, 40'RH → 45RE, 20'/40' OT → 22UT/42UT, 20'/40' FR → 22PC/42PC)
- GATE IN = IEC, IED, IEP · GATE OUT = OEV, OEP · IER / OER are not counted
- GATE IN: Agent when the row has a haulage / shore job no. (`MT-TWP-SEP-28`, `I20261725`, `26BI01624`, `HAL1009`, `HASLC0126090530`, `C3260577867`), otherwise Consignee
- GATE OUT: Shipper when the row has a 16-character booking (`HASLS00000000000`), otherwise Agent

## Run

```bash
python yard_gate_summary.py                                   # reports on the Desktop (moved into GATE MOVE\<dd.mm.yyyy>)
python yard_gate_summary.py 02.10.2026                        # re-run a day already filed
python yard_gate_summary.py 01.10.2026 "C:\path\to\folder"    # reports in another folder (copied, originals kept)
python yard_gate_summary.py 12.10.2026 --check                # compare + MISSING only, into <day>\CHECK (no database / dashboard, files stay)
```

`corrections.csv` (local, git-ignored): `Date,Location,Move,Container No,TPSZ,Note` fixes a TPSZ the yard report got wrong (e.g. 20' DC that is really 22RE).

BC and PW midday reports (`BC2_1ST_…`, `BC2_2ND_…`, `PW_1ST_…`, `PW_2ND_…`) are accepted; when a day has several reports for a yard, the newest file is used.

Output, all in `Desktop\GATE MOVE`:

- `<dd.mm.yyyy>\HA Gate In-Out Summary <dd.mm.yyyy>.xlsx`: per-yard and total summary, system compare, container list
- `<dd.mm.yyyy>\HA Gate Diff <dd.mm.yyyy>.xlsx`: containers that differ from the system
- `<dd.mm.yyyy>\<m>-<d>-MISSING.xls`: containers in the yard reports but not in the system, in the system upload layout (Location, Container No, Move, GateDate, GateTime, BL No). GATE OUT: OEV (OEP for `MT-…` jobs), BL No = booking / job no. GATE IN: IED (IEP for `I########` jobs), BL No = the yard BKG# (HAST BKG#, CELLO BOOKING IN), blank when the yard leaves it empty
- `GATE MOVE DATABASE.xlsx`: running database (re-running a day replaces that day for the yards in the run)
- `HA Gate Dashboard.html`: dashboard built from the database with `dashboard_template.html`

Requires Python with `pandas`, `openpyxl`, `xlrd` (read .xls) and `xlwt` (write .xls).

Data files are git-ignored on purpose; only the code is tracked.
