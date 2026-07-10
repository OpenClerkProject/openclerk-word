#!/usr/bin/env bash
# Standalone macOS installer for the OpenClerk Word add-in manifest -- no Node.js required.
# Copies manifest.xml into Word's shared sideloading folder on macOS
# (~/Library/Containers/com.Microsoft.OsfWebHost/Data/documents/wef), the same folder Word for
# Mac scans for manifests regardless of which Office app is used to sideload them.
set -euo pipefail

manifest_path="manifest.xml"
target=""
dry_run=0

usage() {
  echo "Usage: install-openclerk.sh [--manifest <path>] [--target <dir>] [--dry-run]" >&2
}

while [ $# -gt 0 ]; do
  case "$1" in
    --manifest)
      manifest_path="$2"
      shift 2
      ;;
    --target)
      target="$2"
      shift 2
      ;;
    --dry-run)
      dry_run=1
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage
      exit 1
      ;;
  esac
done

# $script_dir is the installer/ subfolder inside the extracted release package.
# The manifest lives one level up, at the package root.
script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
package_root="$(cd "$script_dir/.." && pwd)"

if [ -z "$target" ]; then
  if [ "$(uname -s)" = "Darwin" ] && [ -n "${HOME:-}" ]; then
    target="$HOME/Library/Containers/com.Microsoft.OsfWebHost/Data/documents/wef"
  else
    target="$package_root/.installer-test/wef"
  fi
fi

case "$manifest_path" in
  /*) manifest_full_path="$manifest_path" ;;
  *) manifest_full_path="$package_root/$manifest_path" ;;
esac

if [ ! -f "$manifest_full_path" ]; then
  echo "Manifest not found: $manifest_full_path" >&2
  exit 1
fi

target_manifest="$target/openclerk-manifest.xml"

echo "Source manifest: $manifest_full_path"
echo "Install target:  $target_manifest"

if [ "$dry_run" -eq 1 ]; then
  echo "Dry run enabled. No files were copied."
else
  mkdir -p "$target"
  cp "$manifest_full_path" "$target_manifest"
  echo "Manifest installed successfully."
fi

echo "Done."
