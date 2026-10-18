"""Record the REAL VBA invoice-generator demo on Xvfb :99.
One continuous take: Calc walkthrough -> terminal runs macro -> invoices appear.
Usage: python3 record_vba.py
100% local. by Muhammad Bilal.
"""
import os
import subprocess
import time
import uno

D = "/home/hatch/workspace/automation-build-2/09-Excel-VBA-Invoice-Generator"
PROF = "/tmp/lo-prof-09"
PORT = 2008
RAW = "/tmp/vba-raw.mp4"

import shutil
shutil.rmtree(PROF, ignore_errors=True)
os.makedirs(PROF, exist_ok=True)

env = dict(os.environ)
env["DISPLAY"] = ":99"

print("Starting LibreOffice on :99 ...")
lo = subprocess.Popen(
    ["soffice", f"-env:UserInstallation=file://{PROF}", "--calc",
     "--nologo", "--norestore", "--nofirststartwizard",
     f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ServiceManager",
     os.path.join(D, "Invoices-base.xlsx")],
    env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(10)

# connect UNO
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
assert ctx, "no UNO connection"
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = None
for _ in range(30):
    for comp in desktop.getComponents():
        try:
            if comp.getSheets().hasByName("Invoice-Data"):
                doc = comp
                break
        except Exception:
            pass
    if doc:
        break
    time.sleep(1)
assert doc, "workbook not open"
controller = doc.getCurrentController()

# try to maximize the LO window to fill 1280x720
try:
    win = controller.getFrame().getContainerWindow()
    win.setPosSize(0, 0, 1280, 720, 15)
except Exception as e:
    print("window resize failed:", e)

print("Starting ffmpeg recording ...")
ff = subprocess.Popen(
    ["ffmpeg", "-y", "-f", "x11grab", "-video_size", "1280x720",
     "-framerate", "15", "-i", ":99", "-pix_fmt", "yuv420p",
     "-c:v", "libx264", "-preset", "veryfast", RAW],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(8)  # show Invoice-Data: the raw data (the problem)

# tour the template sheet
controller.setActiveSheet(doc.getSheets().getByName("Invoice-Template"))
time.sleep(7)

def xterm_is_visible():
    """Screenshot :99 and check something rendered (brightness heuristic)."""
    p = subprocess.run(
        ["ffmpeg", "-y", "-f", "x11grab", "-video_size", "1280x720",
         "-i", ":99", "-frames:v", "1", "/tmp/xt-check.png"],
        capture_output=True)
    try:
        from PIL import Image
        import numpy as np
        img = np.array(Image.open("/tmp/xt-check.png").convert("L"))
        return img.mean() > 8
    except Exception:
        return False


def launch_xterm(cmd, fs=13, hold=5):
    for attempt in range(3):
        xt = subprocess.Popen(
            ["xterm", "-geometry", "160x45+0+0", "-fa", "Monospace",
             "-fs", str(fs), "-e", "bash", "-c", cmd + f"; sleep {hold}"],
            env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(3)
        if xterm_is_visible():
            return xt
        print(f"xterm not visible, retry {attempt + 1}...")
        xt.terminate()
        time.sleep(1)
    raise RuntimeError("xterm would not render")


# open terminal, run the macro driver (real usage)
print("Opening terminal, running macro ...")
xt = launch_xterm(f"cd {D} && echo '$ python3 demo_run.py' && python3 demo_run.py")
xt.wait()  # ~30s: macro runs live, sheets appear, totals print
time.sleep(2)

# page through the three invoices slowly (result)
for name in ("INV-2026-001", "INV-2026-002", "INV-2026-003"):
    controller.setActiveSheet(doc.getSheets().getByName(name))
    time.sleep(9)

# show the real VBA source in terminal
print("Showing source ...")
xt2 = launch_xterm(
    f"cd {D} && ls -la Invoices.xlsm InvoiceGen.bas && echo '--- InvoiceGen.bas (head) ---' && head -40 InvoiceGen.bas",
    fs=12, hold=8)
xt2.wait()
time.sleep(2)

print("Stopping recording ...")
ff.terminate()
ff.wait(timeout=15)
lo.terminate()
print("Raw footage:", RAW)
