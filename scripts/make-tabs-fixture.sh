#!/usr/bin/env bash
#
# Generate a .docx exercising TAB STOPS, using LibreOffice in Docker.
#
# WHY THIS EXISTS
#   apps/playground/public/sample.docx claims to demo w:tabs but carries none:
#   it was re-exported through Google Docs at some point, which flattened every
#   tab into plain text (0 × w:tab, 0 × w:tabs in that file). So the tab path
#   had no fixture at all — nothing to look at in the playground, nothing to
#   regression-test against.
#
#   Authoring the .docx ourselves would be circular: bapbong's exporter writing
#   the tabs its importer then reads proves nothing. LibreOffice is an
#   independent implementation, same reasoning as make-encrypted-fixture.sh.
#
# USAGE
#   scripts/make-tabs-fixture.sh [output.docx]
#
# Writes TWO copies of the same bytes, on purpose:
#   packages/docx/…/__fixtures__  the canonical one, read by the test suite
#   apps/playground/public        so the fixture is one click away in the
#                                 playground (it serves only from public/)
# Regenerating rewrites both, so they cannot drift.
#
# Requires: Docker (the first run builds the image; later runs reuse it).
# Reclaim that space afterwards with scripts/clean-fixture-image.sh.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
# shellcheck source=_fixture-common.sh
source "$SCRIPT_DIR/_fixture-common.sh"

OUT="${1:-$REPO_ROOT/packages/docx/src/lib/__fixtures__/tab-stops.docx}"
PLAYGROUND_COPY="$REPO_ROOT/apps/playground/public/tab-stops.docx"

require_docker

if ! docker image inspect "$FIXTURE_IMAGE" >/dev/null 2>&1; then
  echo "Building $FIXTURE_IMAGE (first run only; a few minutes)…"
  docker build -q -t "$FIXTURE_IMAGE" -f "$SCRIPT_DIR/libreoffice.Dockerfile" \
    "$SCRIPT_DIR" >/dev/null
fi

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
cp "$SCRIPT_DIR/tabs-fixture.fodt" "$WORK/source.fodt"

docker run --rm -v "$WORK:/work" "$FIXTURE_IMAGE" bash -lc "
  set -e
  soffice --headless --norestore --convert-to docx --outdir /work /work/source.fodt >/dev/null 2>&1
"

mkdir -p "$(dirname "$OUT")"
cp "$WORK/source.docx" "$OUT"
if [ "$OUT" != "$PLAYGROUND_COPY" ]; then
  mkdir -p "$(dirname "$PLAYGROUND_COPY")"
  cp "$WORK/source.docx" "$PLAYGROUND_COPY"
fi

echo
echo "Fixture: $OUT"
[ "$OUT" != "$PLAYGROUND_COPY" ] && echo "Playground copy: $PLAYGROUND_COPY"
# Report what LibreOffice actually emitted — a silent conversion that dropped
# the tabs would otherwise look like success.
python3 - "$OUT" <<'PY'
import re, sys, zipfile
xml = zipfile.ZipFile(sys.argv[1]).read('word/document.xml').decode('utf8')
stops = re.findall(r'<w:tab\s[^>]*w:val="(\w+)"[^>]*?(?:w:leader="(\w+)")?[^>]*/>', xml)
kinds = {}
for val, leader in stops:
    key = f"{val}+{leader}" if leader else val
    kinds[key] = kinds.get(key, 0) + 1
print(f"  w:tabs blocks : {len(re.findall(r'<w:tabs>', xml))}")
print(f"  tab stops     : {kinds or 'NONE — conversion dropped them'}")
print(f"  <w:tab/> chars: {len(re.findall(r'<w:tab/>', xml))}")
PY

echo
echo "Commit the fixture; nothing in the test suite needs Docker."
