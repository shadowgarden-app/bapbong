#!/usr/bin/env bash
#
# Generate a .docx exercising CHARACTER TRACKING (w:spacing) and HORIZONTAL
# GLYPH SCALING (w:w), using LibreOffice in Docker.
#
# WHY THIS EXISTS
#   The two properties combine, and the combination has one open question:
#   does the scale apply to the tracking as well as to the glyphs? bapbong
#   assumes not — width = glyphs × scale + n × tracking — and ECMA-376's
#   wording backs that (w:w scales "the character size"; w:spacing is an
#   absolute twips measure "added after each character").
#
#   Asserting that against a .docx bapbong wrote would prove nothing: the
#   exporter would just be agreeing with the importer. LibreOffice is an
#   independent implementation, same reasoning as make-tabs-fixture.sh.
#
# USAGE
#   scripts/make-glyph-scale-fixture.sh [output.docx]
#
# Writes TWO copies of the same bytes, like the other fixtures:
#   packages/docx/…/__fixtures__  the canonical one, read by the test suite
#   apps/playground/public        so it is one click away in the playground
#
# Requires: Docker (the first run builds the image; later runs reuse it).
# Reclaim that space afterwards with scripts/clean-fixture-image.sh.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
# shellcheck source=_fixture-common.sh
source "$SCRIPT_DIR/_fixture-common.sh"

OUT="${1:-$REPO_ROOT/packages/docx/src/lib/__fixtures__/glyph-scale.docx}"
PLAYGROUND_COPY="$REPO_ROOT/apps/playground/public/glyph-scale.docx"

require_docker

if ! docker image inspect "$FIXTURE_IMAGE" >/dev/null 2>&1; then
  echo "Building $FIXTURE_IMAGE (first run only; a few minutes)…"
  docker build -q -t "$FIXTURE_IMAGE" -f "$SCRIPT_DIR/libreoffice.Dockerfile" \
    "$SCRIPT_DIR" >/dev/null
fi

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
cp "$SCRIPT_DIR/glyph-scale-fixture.fodt" "$WORK/source.fodt"

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
# Report what LibreOffice actually emitted. A conversion that silently dropped
# either property would otherwise look like success, and the interesting line
# is the one carrying BOTH.
python3 - "$OUT" <<'PY'
import re, sys, zipfile
xml = zipfile.ZipFile(sys.argv[1]).read('word/document.xml').decode('utf8')
rprs = re.findall(r'<w:rPr>(.*?)</w:rPr>', xml, re.S)
sp = [m.group(1) for r in rprs for m in [re.search(r'<w:spacing w:val="(-?\d+)"/>', r)] if m]
sc = [m.group(1) for r in rprs for m in [re.search(r'<w:w w:val="([\d%]+)"/>', r)] if m]
both = sum(1 for r in rprs if '<w:spacing' in r and '<w:w ' in r)
print(f"  w:spacing values : {sp or 'NONE — conversion dropped tracking'}")
print(f"  w:w values       : {sc or 'NONE — conversion dropped scaling'}")
print(f"  rPr with BOTH    : {both}  <- the case the fixture exists for")
PY

# ── The interaction, settled from LibreOffice's own rendering ─────────
# The .docx carries markup, not widths, so measuring it with OUR measurer
# would just be our assumption agreeing with itself. Render the same source to
# PDF instead and read the glyph positioning LibreOffice chose.
#
# LibreOffice puts the horizontal scale in the text matrix (Tm) and the
# tracking in the TJ array. A TJ adjustment sits in TEXT space, so the matrix
# scales it too — which makes the two possible answers distinguishable:
#
#   tracking IS scaled      -> same TJ adjustment with and without the scale
#   tracking is NOT scaled  -> TJ adjustment DIVIDED by the scale, so that it
#                              survives the matrix at its absolute value
echo
echo "Checking the tracking/scale interaction against LibreOffice's own layout..."
docker run --rm -v "$WORK:/work" "$FIXTURE_IMAGE" bash -lc "
  set -e
  soffice --headless --norestore --convert-to pdf --outdir /work /work/source.fodt >/dev/null 2>&1
"
python3 - "$WORK/source.pdf" <<'PDFCHECK'
import re, statistics, sys, zlib
raw = open(sys.argv[1], 'rb').read()
parts = []
for m in re.finditer(rb'stream\r?\n(.*?)\r?\nendstream', raw, re.S):
    try: parts.append(zlib.decompress(m.group(1)))
    except Exception: pass
txt = b'\n'.join(parts).decode('latin-1', 'ignore')
runs = []
for b in re.findall(r'BT\s(.*?)\sET', txt, re.S):
    tm = re.search(r'([\d.]+) 0 0 1 [\d.]+ [\d.]+ Tm', b)
    tj = re.search(r'\[(.*?)\]\s*TJ', b, re.S)
    if not tj: continue
    adj = [int(x) for x in re.findall(r'>(-?\d+)<', tj.group(1))]
    # The four sample lines share one long text; the labels are short.
    if len(adj) >= 10:
        runs.append((float(tm.group(1)) if tm else 1.0, statistics.median(adj)))
if len(runs) != 4:
    print(f"  could not identify the four sample runs (found {len(runs)})")
    raise SystemExit(1)
(_, plain), (_, tracked), (scale, scaled), (_, both) = runs
tr, bo = tracked - plain, both - scaled
print(f"  tracking, no scale  : {tr:>8.1f}  ({tr/1000*12:+.2f}pt at 12pt)")
print(f"  tracking, scale {scale} : {bo:>8.1f}")
print(f"    if scaled too     : {tr:>8.1f}")
print(f"    if left absolute  : {tr/scale:>8.1f}")
verdict = 'LEFT ABSOLUTE' if abs(bo - tr/scale) < abs(bo - tr) else 'SCALED WITH THE GLYPHS'
print(f"  LibreOffice: tracking is {verdict}")
print("  bapbong agrees: applyGlyphSpec sets letterSpacing = tracking / scale")
PDFCHECK

echo
echo "Commit the fixture; nothing in the test suite needs Docker."
