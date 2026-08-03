#!/usr/bin/env python3
"""Write a password-protected .docx with LibreOffice, inside the container.

Runs against a headless soffice over UNO. The command line's ``--convert-to``
cannot set a document password (its filter options reach FilterData, while the
password is a MediaDescriptor property), so the file is stored through
``storeToURL`` instead.

Used by scripts/make-encrypted-fixture.sh. Not part of the build or the tests.
"""
import os
import subprocess
import sys
import time

import uno  # provided by python3-uno
from com.sun.star.beans import PropertyValue

PORT = 2002


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect(timeout=90):
    """Resolve a UNO context, waiting for soffice to finish starting."""
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    url = f"uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext"
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        try:
            return resolver.resolve(url)
        except Exception as err:  # noqa: BLE001 — retry until the socket is up
            last = err
            time.sleep(1)
    raise SystemExit(f"could not reach soffice on port {PORT}: {last}")


def main():
    if len(sys.argv) != 4:
        raise SystemExit("usage: encrypt-fixture.py <in.docx> <out.docx> <password>")
    src, dst, password = sys.argv[1], sys.argv[2], sys.argv[3]

    soffice = subprocess.Popen(
        [
            "soffice",
            "--headless",
            "--norestore",
            "--nolockcheck",
            f"--accept=socket,host=127.0.0.1,port={PORT};urp;",
        ]
    )
    try:
        ctx = connect()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(src)),
            "_blank",
            0,
            (prop("Hidden", True),),
        )
        if doc is None:
            raise SystemExit(f"LibreOffice could not open {src}")
        try:
            doc.storeToURL(
                uno.systemPathToFileUrl(os.path.abspath(dst)),
                (
                    prop("FilterName", "MS Word 2007 XML"),
                    prop("EncryptFile", True),
                    prop("Password", password),
                ),
            )
        finally:
            doc.close(False)
        print(f"wrote {dst}")
    finally:
        soffice.terminate()
        try:
            soffice.wait(timeout=20)
        except subprocess.TimeoutExpired:
            soffice.kill()


if __name__ == "__main__":
    main()
