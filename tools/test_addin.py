"""End-to-end test for the LOESS Calc Add-In (build/CalcLoessAddin.oxt).

Run with LibreOffice's bundled Python (it ships the `uno` module) against a
headless instance listening on a UNO socket:

    soffice --headless --invisible --norestore --accept="socket,host=localhost,port=2002;urp;" &
    /usr/lib64/libreoffice/program/python tools/test_addin.py

Prints RESULT: PASS / FAIL and exits non-zero on failure. Requires the
extension to already be installed (see build_addin.sh + `unopkg add`).
"""
import sys
import time
import uno


def connect(port=2002, tries=40):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % port
    last = None
    for _ in range(tries):
        try:
            return resolver.resolve(url)
        except Exception as e:
            last = e
            time.sleep(0.5)
    raise SystemExit("could not connect to LibreOffice: %s" % last)


def close_enough(a, b, tol=1e-6):
    return abs(a - b) < tol


def main():
    ctx = connect()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 1. LOESS is registered as a real Add-In (drives Function Wizard/autocomplete).
    fdescs = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionDescriptions", ctx)
    registered = any(
        {p.Name: p.Value for p in fdescs.getByIndex(i)}.get("Name") == "LOESS"
        for i in range(fdescs.Count)
    )

    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    try:
        sheet = doc.Sheets.getByIndex(0)
        # Linear data y = 2x: a degree-1 fit should reproduce it exactly.
        for i in range(1, 21):
            sheet.getCellByPosition(0, i - 1).setValue(i)
            sheet.getCellByPosition(1, i - 1).setValue(2 * i)

        full = sheet.getCellByPosition(3, 0)
        full.setFormula("=LOESS($A$1:$A$20;$B$1:$B$20;10;0.5;1;0)")

        defaults = sheet.getCellByPosition(3, 1)
        defaults.setFormula("=LOESS($A$1:$A$20;$B$1:$B$20;10)")

        doc.calculateAll()

        full_ok = full.getError() == 0 and close_enough(full.getValue(), 20.0)
        defaults_ok = defaults.getError() == 0 and close_enough(defaults.getValue(), 20.0)
    finally:
        doc.close(False)
        desktop.terminate()

    print("Registered in FunctionDescriptions:", registered)
    print("Full-args value :", full.getValue() if full_ok else full.getError(), "->", "PASS" if full_ok else "FAIL")
    print("Default-args value:", defaults.getValue() if defaults_ok else defaults.getError(), "->", "PASS" if defaults_ok else "FAIL")

    ok = registered and full_ok and defaults_ok
    print("RESULT:", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
