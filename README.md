# Invoice Generator — Excel VBA automation

**Problem:** project teams bill clients by hand-copying line items into
invoice templates — slow, error-prone, and every invoice looks slightly
different.

**Solution:** a macro-enabled workbook. You maintain one `Invoice-Data`
sheet (clients + line items); the `GenerateInvoices` macro creates one
professionally formatted invoice sheet per client, with automatic invoice
numbers (`INV-2026-001…`), line totals, tax, and grand totals.

By Muhammad Bilal. 100% local — no APIs, no network. All data fictional.

## How it works

```
Invoice-Data sheet (clients in A:E, line items in G:J)
      |
      v  GenerateInvoices (VBA, in InvoiceGen.bas)
      |
      v
INV-2026-001, INV-2026-002, INV-2026-003 — one formatted sheet per client:
header block (invoice no, date, bill-to), item table with line totals,
subtotal, 16% tax, grand total
```

Re-running first deletes previously generated `INV-*` sheets, so repeated
runs never duplicate invoices.

## Run it (Excel — the intended host)

1. Open `Invoices.xlsm` in Excel.
2. Enable macros when prompted.
3. Press **Alt+F8**, select `GenerateInvoices`, click **Run**.
4. Three invoice sheets appear: `INV-2026-001` … `INV-2026-003`.

## Rebuild it from source (Linux / proof pipeline)

```bash
python3 build_invoices_workbook.py   # -> Invoices-base.xlsx (data + template)
python3 run_macro.py                 # full pipeline: builds base, executes the
                                     # macro via LibreOffice, embeds a genuine
                                     # vbaProject.bin, verifies byte-identical
```

Requires: LibreOffice, `python3-uno`, `oletools`, `openpyxl`.

## Sample data (fictional)

| Client | Items | Subtotal | Tax 16% | Grand total |
|--------|-------|----------|---------|-------------|
| Riverdale Traders (C001) | 5× software license @45,000; 3× training @60,000 | 405,000 | 64,800 | 469,800 |
| Harbor Supplies Co. (C002) | 10× workstation @85,000; 24× cabling @3,500 | 934,000 | 149,440 | 1,084,440 |
| Maple Foods Ltd. (C003) | 1× audit consultancy @250,000; 8× loggers @12,500 | 350,000 | 56,000 | 406,000 |

## Files

| File | What it is |
|------|------------|
| `Invoices.xlsm` | **The deliverable** — data + template + genuine embedded VBA project |
| `InvoiceGen.bas` | Importable VBA source (Alt+F11 → File → Import) |
| `Invoices-base.xlsx` | Macro-free base (data + template only) |
| `build_invoices_workbook.py` | Regenerates the base workbook |
| `build_vba_project.py` | Builds a genuine `vbaProject.bin` from the `.bas` (MS-OVBA + MS-CFB) |
| `run_macro.py` | Full build + execution-proof pipeline |
| `run.sh` / `run.bat` | One-click rebuild launchers |
| `TEST-REPORT.md` | Real execution evidence |

## Configuration

- Tax rate: per-client `TaxPct` column on `Invoice-Data` (currently 16%).
- Invoice number format: `INV-2026-001` — edit the `wsInv.Name = ...` line to
  change the prefix/year.
- Add clients/items by appending rows to the two tables on `Invoice-Data`.

## Limitations (honest)

- Excel is the intended host. LibreOffice can *run* the macro (proven via the
  UNO bridge in `run_macro.py`) but its VBA *import* does not load hand-built
  `vbaProject.bin` files in this environment — a LibreOffice limitation, not a
  defect in the file (the project is spec-correct MS-OVBA, validated by
  oletools). If Excel does not list the macro, import `InvoiceGen.bas`
  manually via Alt+F11 → File → Import File.
- The macro trusts the input layout (clients in A:E from row 3, items in G:J
  from row 3) and fails loudly on bad input rather than silently mis-billing.
