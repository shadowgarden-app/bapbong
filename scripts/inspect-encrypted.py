#!/usr/bin/env python3
"""Report what an encrypted .docx actually is — used by make-encrypted-fixture.sh.

WHICH SCHEME matters as much as whether the file got encrypted at all. 4.4 is
Agile (Word 2010+); 3.2 is the older Standard (Word 2007 — and what LibreOffice
writes). They are different formats: AES-CBC with a SHA-512 KDF versus AES-ECB
with SHA-1, so a reader that handles one refuses the other.

The version header is found by scanning rather than by parsing the container:
the repo's own container reader is part of what these fixtures exist to test,
so a diagnostic must not depend on it.
"""
import struct
import sys

OLE_SIGNATURE = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])


def main() -> int:
    data = open(sys.argv[1], "rb").read()
    print(f"Size: {len(data)} bytes")
    if data[:8] != OLE_SIGNATURE:
        print("Container: NOT OLE — the password was ignored, nothing was encrypted")
        return 1
    print("Container: OLE (encrypted)")

    for off in range(0, len(data) - 16, 2):
        major, minor = struct.unpack_from("<HH", data, off)
        flags = struct.unpack_from("<I", data, off + 4)[0]
        if (major, minor) == (4, 4) and flags == 0x40:
            print("Scheme: 4.4 Agile (AES-CBC + SHA-512 KDF)")
            return 0
        if (major, minor) == (3, 2) and flags & 0x20:
            algid, alghash, keybits = struct.unpack_from("<III", data, off + 20)
            print(
                f"Scheme: 3.2 Standard (AES-ECB + SHA-1) — AlgID={algid:#x} "
                f"hash={alghash:#x} keyBits={keybits}"
            )
            print("  NOTE: a DIFFERENT format from Agile, not a variant of it.")
            return 0
    print("Scheme: could not identify — inspect EncryptionInfo by hand")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
