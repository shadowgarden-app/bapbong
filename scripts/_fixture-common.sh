# Shared by make-encrypted-fixture.sh and clean-fixture-image.sh.
#
# The image name lives here so the two cannot drift: a rename in one place
# would otherwise leave the cleanup script reporting "nothing to remove" while
# a gigabyte of LibreOffice sits on disk.
#
# Sourced, not executed.

FIXTURE_IMAGE="bapbong-libreoffice-fixture"

# Print a clear message and fail when the daemon is not up. Both scripts need
# it, and "Cannot connect to the Docker daemon" on its own sends people
# looking for the wrong problem.
require_docker() {
  if ! docker info >/dev/null 2>&1; then
    echo "Docker is not running — start Docker Desktop and try again." >&2
    return 1
  fi
}
