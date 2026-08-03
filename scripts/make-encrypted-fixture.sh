#!/usr/bin/env bash
#
# Generate a password-protected .docx test fixture using LibreOffice in Docker.
#
# WHY THIS EXISTS
#   packages/docx opens password-protected documents (MS-OFFCRYPTO). Proving
#   that against files we encrypt ourselves is circular — a wrong constant on
#   both sides cancels out. This produces a fixture from an INDEPENDENT
#   implementation, so the test suite can check the direction that matters:
#   somebody else's encryption, our decryption.
#
# YOU PROBABLY DO NOT NEED TO RUN THIS
#   The fixture it produces is committed (see packages/docx/src/lib/__fixtures__).
#   Tests read that file; neither they nor CI need Docker. Run this only to
#   regenerate — a different password, a newer LibreOffice, another variant.
#
# USAGE
#   scripts/make-encrypted-fixture.sh [output.docx] [password]
#
# Requires: Docker (first run builds a ~1 GB image; later runs reuse it).
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT="${1:-$REPO_ROOT/packages/docx/src/lib/__fixtures__/libreoffice-encrypted.docx}"
PASSWORD="${2:-bapbong-test}"
IMAGE="bapbong-libreoffice-fixture"

if ! docker info >/dev/null 2>&1; then
  echo "Docker is not running — start Docker Desktop and try again." >&2
  exit 1
fi

if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
  echo "Building $IMAGE (first run only; a few minutes)…"
  docker build -q -t "$IMAGE" -f "$REPO_ROOT/scripts/libreoffice.Dockerfile" \
    "$REPO_ROOT/scripts" >/dev/null
fi

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
cp "$REPO_ROOT/scripts/encrypt-fixture.py" "$WORK/"

# The source document is written here rather than taken from the repo, so the
# fixture is self-contained and unambiguously ours (no third-party content).
cat > "$WORK/source.txt" <<'TXT'
bapbong encrypted fixture
This document was encrypted by LibreOffice, not by bapbong.
Decrypting it is what proves our MS-OFFCRYPTO reader agrees with an
independent implementation.
TXT

docker run --rm -v "$WORK:/work" "$IMAGE" bash -lc "
  set -e
  # LibreOffice writes the plain .docx first; the password is applied on the
  # second pass, through UNO (the CLI cannot set one).
  soffice --headless --norestore --convert-to docx --outdir /work /work/source.txt >/dev/null 2>&1
  python3 /work/encrypt-fixture.py /work/source.docx /work/encrypted.docx '$PASSWORD'
"

mkdir -p "$(dirname "$OUT")"
cp "$WORK/encrypted.docx" "$OUT"

echo
echo "Fixture: $OUT"
echo "Password: $PASSWORD"
python3 "$REPO_ROOT/scripts/inspect-encrypted.py" "$OUT"

echo
echo "Commit the fixture; tests read it directly (no Docker needed)."
