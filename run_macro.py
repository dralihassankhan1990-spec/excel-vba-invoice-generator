"""Full build + proof pipeline for Invoices.xlsm.

1. build_invoices_workbook.py -> Invoices-base.xlsx (Invoice-Data + Invoice-Template)
2. LibreOffice UNO: open base, inject InvoiceGen.bas as a Basic module,
   execute GenerateInvoices, verify the 3 invoice sheets + totals in memory.
3. Save macro-built workbook, inject hand-built vbaProject.bin
   (build_vba_project.py, MS-OVBA + MS-CFB) -> final Invoices.xlsm.
4. olevba: embedded macro byte-identical to InvoiceGen.bas; LO re-opens
   the delivered file and confirms sheets + totals.

100% local. by Muhammad Bilal.
"""
import logging
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_XLSX = os.path.join(HERE, "Invoices-base.xlsx")
TMP_XLSX = os.path.join(HERE, "_tmp_after_macro.xlsx")
OUT_XLSM = os.path.join(HERE, "Invoices.xlsm")
BAS_FILE = os.path.join(HERE, "InvoiceGen.bas")
PORT = 2003

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("run_macro")

# hand-verified expected totals: (subtotal, tax, grand)
EXPECTED = {
    "INV-2026-001": (405000, 64800, 469800),
    "INV-2026-002": (934000, 149440, 1083440),
    "INV-2026-003": (350000, 56000, 406000),
}


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_soffice():
    proc = subprocess.Popen(
        ["soffice", "--headless", "--nologo", "--norestore",
         f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ServiceManager",
         "--nofirststartwizard"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    log.info("soffice started (pid %s), waiting for UNO socket...", proc.pid)
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
            log.info("UNO bridge connected.")
            return proc, ctx
        except Exception:
            time.sleep(1)
    raise RuntimeError("Could not connect to LibreOffice UNO bridge")


def verify_invoices(doc):
    sheets = doc.getSheets()
    for name, (sub, tax, grand) in EXPECTED.items():
        assert sheets.hasByName(name), f"{name} was not created!"
        sh = sheets.getByName(name)
        got_sub = sh.getCellByPosition(4, 14).getValue()
        got_tax = sh.getCellByPosition(4, 15).getValue()
        got_grand = sh.getCellByPosition(4, 16).getValue()
        assert abs(got_sub - sub) < 0.01, f"{name} subtotal {got_sub} != {sub}"
        assert abs(got_tax - tax) < 0.01, f"{name} tax {got_tax} != {tax}"
        assert abs(got_grand - grand) < 0.01, f"{name} grand {got_grand} != {grand}"
        inv_no = sh.getCellByPosition(2, 2).getString()
        assert inv_no == name, f"invoice no cell = {inv_no!r}, expected {name!r}"
        log.info("%s OK — subtotal=%d tax=%d grand=%d", name, got_sub, got_tax, got_grand)
    log.info("Invoice verification PASSED (3 sheets, numbers, totals).")


def main() -> int:
    log.info("Step 1: building base workbook...")
    r = subprocess.run([sys.executable, os.path.join(HERE, "build_invoices_workbook.py")],
                       cwd=HERE, capture_output=True, text=True)
    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        log.error("build_invoices_workbook.py failed:\n%s", r.stderr)
        return 1

    with open(BAS_FILE, encoding="utf-8") as f:
        vba = f.read()
    vba_lines = [ln for ln in vba.splitlines() if not ln.startswith("Attribute ")]
    module_code = "Option VBASupport 1\n" + "\n".join(vba_lines)
    log.info("Step 2: running GenerateInvoices via UNO (VBA source: %d lines)...", len(vba_lines))

    proc, ctx = start_soffice()
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(BASE_XLSX), "_blank", 0, ())
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        std = libs.getByName("Standard")
        if std.hasByName("InvoiceGen"):
            std.removeByName("InvoiceGen")
        std.insertByName("InvoiceGen", module_code)
        sp = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.ScriptProviderForBasic", ctx)
        sp.initialize((doc,))
        script = sp.getScript(
            "vnd.sun.star.script:Standard.InvoiceGen.GenerateInvoices"
            "?language=Basic&location=document")
        script.invoke((), (), ())
        log.info("GenerateInvoices returned without error.")
        verify_invoices(doc)
        if os.path.exists(TMP_XLSX):
            os.remove(TMP_XLSX)
        doc.storeAsURL(uno.systemPathToFileUrl(TMP_XLSX),
                       (pv("FilterName", "Calc MS Excel 2007 XML"), pv("Overwrite", True)))
        log.info("Macro-built workbook saved as %s.", TMP_XLSX)
        doc.close(True)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
        log.info("soffice stopped.")

    log.info("Step 3: building vbaProject.bin + injecting into macro-built workbook...")
    sys.path.insert(0, HERE)
    from build_vba_project import build_vba_project_bin, inject_into_xlsm
    vba_bin = build_vba_project_bin(BAS_FILE, module_name="InvoiceGen")
    log.info("vbaProject.bin built (%d bytes).", len(vba_bin))
    if os.path.exists(OUT_XLSM):
        os.remove(OUT_XLSM)
    inject_into_xlsm(TMP_XLSX, vba_bin, OUT_XLSM)
    log.info("Final %s.", OUT_XLSM)

    log.info("Step 4: verifying delivered .xlsm...")
    from oletools.olevba import VBA_Parser
    p = VBA_Parser(OUT_XLSM)
    assert p.detect_vba_macros(), "olevba: no macros detected in delivered .xlsm!"
    mods = list(p.extract_macros())
    std_mods = [m for m in mods if m[2] == "InvoiceGen.bas"]
    assert len(std_mods) == 1, f"expected InvoiceGen module, got {[m[2] for m in mods]}"
    expected = "\r\n".join(vba_lines) + "\r\n"
    assert std_mods[0][3] == expected, "extracted macro differs from InvoiceGen.bas!"
    log.info("olevba: embedded macro is BYTE-IDENTICAL to InvoiceGen.bas.")

    proc, ctx = start_soffice()
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(OUT_XLSM), "_blank", 0, ())
        sheets = doc.getSheets()
        assert sheets.hasByName("Invoice-Data") and sheets.hasByName("Invoice-Template"), \
            "base sheets missing from delivered .xlsm"
        verify_invoices(doc)
        doc.close(True)
        log.info("Delivered .xlsm re-opened: all sheets intact, invoice totals hold.")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()

    if os.path.exists(TMP_XLSX):
        os.remove(TMP_XLSX)
    log.info("DONE — %s ready.", OUT_XLSM)
    return 0


if __name__ == "__main__":
    sys.exit(main())
