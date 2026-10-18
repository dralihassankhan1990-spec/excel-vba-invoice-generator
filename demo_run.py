"""Video demo driver: injects InvoiceGen into the open workbook on the
DISPLAY=:99 LibreOffice instance, runs GenerateInvoices, then pages through
the generated invoice sheets with pauses so the screen recording shows them.
Usage: DISPLAY=:99 python3 demo_run.py  (soffice must be running with
  --accept socket on port 2008)
100% local. by Muhammad Bilal.
"""
import time
import uno

PORT = 2008

print("Connecting to LibreOffice on :99 ...")
local_ctx = uno.getComponentContext()
resolver = local_ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_ctx)
ctx = None
for _ in range(30):
    try:
        ctx = resolver.resolve(
            f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
if ctx is None:
    raise SystemExit("Could not connect to LibreOffice")
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

# find the open Invoices workbook
doc = None
for comp in desktop.getComponents():
    try:
        if comp.getSheets().hasByName("Invoice-Data"):
            doc = comp
            break
    except Exception:
        pass
if doc is None:
    raise SystemExit("Invoices workbook not found among open documents")
print("Found workbook. Injecting InvoiceGen module...")

with open("/home/hatch/workspace/automation-build-2/09-Excel-VBA-Invoice-Generator/InvoiceGen.bas",
          encoding="utf-8") as f:
    vba = f.read()
lines = [ln for ln in vba.splitlines() if not ln.startswith("Attribute ")]
code = "Option VBASupport 1\n" + "\n".join(lines)
libs = doc.BasicLibraries
if not libs.hasByName("Standard"):
    libs.createLibrary("Standard")
std = libs.getByName("Standard")
if std.hasByName("InvoiceGen"):
    std.removeByName("InvoiceGen")
std.insertByName("InvoiceGen", code)

print("Running GenerateInvoices macro...")
sp = smgr.createInstanceWithContext(
    "com.sun.star.script.provider.ScriptProviderForBasic", ctx)
sp.initialize((doc,))
script = sp.getScript(
    "vnd.sun.star.script:Standard.InvoiceGen.GenerateInvoices"
    "?language=Basic&location=document")
script.invoke((), (), ())
print("Macro finished. Showing generated invoices...")

controller = doc.getCurrentController()
for name in ("INV-2026-001", "INV-2026-002", "INV-2026-003"):
    sheet = doc.getSheets().getByName(name)
    controller.setActiveSheet(sheet)
    total = sheet.getCellByPosition(4, 16).getValue()
    print(f"  {name}: grand total = {total:,.0f} PKR")
    time.sleep(6)

print("DONE — 3 invoices generated.")
