# LibreOffice, headless, for generating encrypted-.docx test fixtures.
#
# Built from Debian rather than a prebuilt LibreOffice image because most of
# those are amd64-only — on an arm64 machine they run under emulation, slowly
# or not at all. Debian ships libreoffice on both.
#
# trixie, NOT bookworm: bookworm's LibreOffice 7.4 accepts a password on the
# .docx filter and silently writes the file unencrypted (its ODF filter does
# encrypt, which is how that was pinned down). 25.2 writes a real encrypted
# container.
#
# Used only by scripts/make-encrypted-fixture.sh; nothing in the test suite or
# CI depends on this image (the fixtures it produces are committed).
FROM debian:trixie-slim

# libreoffice-writer: the .docx filters. python3-uno: the scripting bridge —
# the CLI's --convert-to cannot set a document password, which is the whole
# point here, so the fixture is written through UNO's storeToURL instead.
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-writer \
      python3-uno \
      python3 \
 && rm -rf /var/lib/apt/lists/*

# LibreOffice needs a writable HOME for its user profile.
ENV HOME=/tmp
WORKDIR /work
