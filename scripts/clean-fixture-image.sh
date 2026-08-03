#!/usr/bin/env bash
#
# Remove the LibreOffice image that make-encrypted-fixture.sh builds.
#
# That image exists only to REGENERATE fixtures — the fixtures
# themselves are committed, so tests and CI never touch it. Deleting it costs
# nothing but the next regeneration's build time.
#
# USAGE
#   scripts/clean-fixture-image.sh
#
# Safe to run twice: it reports and exits cleanly when there is nothing there.
#
# It does NOT touch the BuildKit cache. Reclaiming that means
# `docker builder prune`, which is global — it would take every other
# project's cache with it, and that is not this script's call to make.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=_fixture-common.sh
source "$SCRIPT_DIR/_fixture-common.sh"

case "${1:-}" in
  '') ;;
  -h | --help)
    sed -n '2,17p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
    exit 0
    ;;
  *)
    echo "unknown option: $1 (try --help)" >&2
    exit 1
    ;;
esac

require_docker

if ! docker image inspect "$FIXTURE_IMAGE" >/dev/null 2>&1; then
  echo "Nothing to remove — $FIXTURE_IMAGE is not present."
  exit 0
fi

SIZE_MB=$(($(docker image inspect "$FIXTURE_IMAGE" --format '{{.Size}}') / 1024 / 1024))
docker image rm "$FIXTURE_IMAGE" >/dev/null
echo "Removed $FIXTURE_IMAGE (~${SIZE_MB} MB)."
echo "Untagged leftovers from older builds: docker image prune"
echo "Regenerate a fixture any time with scripts/make-encrypted-fixture.sh."
