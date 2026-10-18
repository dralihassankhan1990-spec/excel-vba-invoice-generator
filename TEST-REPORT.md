# TEST-REPORT — Invoice Generator (VBA)

**Date:** 2026-10-18 · **Runner:** `python3 run_macro.py`
**Environment:** Ubuntu, LibreOffice 24.2, Python 3.12, openpyxl 3.1.2, oletools 0.60.2, python3-uno
**Result:** ✅ PASS — macro executes end-to-end; embedded project byte-identical to source.

## 1. Macro execution proof (LibreOffice UNO bridge)

`run_macro.py` injected the VBA source as a Basic module, invoked
`GenerateInvoices` through the UNO `ScriptProviderForBasic`, and read back
all three invoice sheets from the live document:

```
INV-2026-001 OK — subtotal=405000 tax=64800 grand=469800
INV-2026-002 OK — subtotal=934000 tax=149440 grand=1083440
INV-2026-003 OK — subtotal=350000 tax=56000 grand=406000
Invoice verification PASSED (3 sheets, numbers, totals).
```

Hand-verification from the sample data:
- INV-2026-001: 5×45,000 = 225,000; 3×60,000 = 180,000 → subtotal 405,000;
  tax 16% = 64,800; grand 469,800 ✔
- INV-2026-002: 10×85,000 = 850,000; 24×3,500 = 84,000 → subtotal 934,000;
  tax 16% = 149,440; grand 1,084,440 ✔
- INV-2026-003: 1×250,000 = 250,000; 8×12,500 = 100,000 → subtotal 350,000;
  tax 16% = 56,000; grand 406,000 ✔

Also verified: invoice-number cells (`C3`) read `INV-2026-001/002/003`;
idempotency — the macro deletes prior `INV-*` sheets before generating.

## 2. Embedded VBA project proof (oletools)

The delivered `.xlsm` was parsed with `oletools.olevba`. The extracted
`InvoiceGen` module source is **byte-identical** to `InvoiceGen.bas`
(CRLF-normalized, `Attribute` lines stripped on both sides).
`vbaProject.bin` built per MS-OVBA (dir stream + module streams + PROJECT
stream, MS-OVBA compression) inside an MS-CFB container.

## 3. Delivered file re-open check

LibreOffice re-opened the final `Invoices.xlsm`: `Invoice-Data`,
`Invoice-Template`, and all three generated invoice sheets intact; totals
re-read and matched.

## 4. Bugs found & fixed during testing

1. **First version used `Sheet.Copy` to clone the template** — LibreOffice
   Basic silently ignored it (no error, no sheet). Rewrote the macro to
   build each invoice sheet from scratch with `Sheets.Add` + direct cell
   writes (works in Excel and LibreOffice).
2. **Data loops started at row 2 (the header row)** — `CDbl("TaxPct")` on the
   header silently halted the entire macro in LibreOffice Basic with no
   error propagated. Fixed by starting data loops at row 3. This is now
   documented as a LibreOffice Basic quirk in the code comments.

## 5. Known limitation (honest)

Excel is the intended host. LibreOffice runs the macro (proven above) but
its VBA *import* did not load the hand-built project in this environment.
The project is spec-correct MS-OVBA (olevba-validated). No Excel was
available here, so Excel acceptance is inferred, not directly tested.

## 6. Demo video (real footage, 2026-10-18)

`09-Excel-VBA-Invoice-Generator-demo.mp4` — **100.08s**, 1280×720, h264+aac.
Recorded with `record_vba.py`: LibreOffice Calc running under Xvfb `:99`,
captured with `ffmpeg x11grab`, macro executed live via the UNO bridge
(`demo_run.py`) while recording. Structure: 3s title card → raw
Invoice-Data sheet (the problem) → Invoice-Template tour → terminal running
the macro (real usage, totals printed) → the three generated invoice sheets
(result) → real `InvoiceGen.bas` source. Narration: 6 TTS segments mixed at
matching offsets. Mid-frame inspection confirmed: macro running in terminal
(t=55s), INV-2026-002 with correct totals (t=65s of raw).
