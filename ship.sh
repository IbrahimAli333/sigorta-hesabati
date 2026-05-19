#!/usr/bin/env bash
# ship.sh — Rebuild the macOS .app and push the code so GitHub Actions
# can build the Windows .exe. Usage:
#
#   ./ship.sh "What you changed in plain language"
#
# If you forget the message it'll ask for one.

set -euo pipefail

cd "$(dirname "$0")"

MSG="${1:-}"
if [ -z "$MSG" ]; then
  printf "Commit message: "
  read -r MSG
  if [ -z "$MSG" ]; then
    echo "❌ Commit message required. Aborting."
    exit 1
  fi
fi

echo "================================================================"
echo "  Step 1/3 — Rebuilding macOS .app (this takes ~30 seconds)"
echo "================================================================"
python3 -m PyInstaller sigorta.spec --clean --noconfirm >/tmp/ship-build.log 2>&1 || {
  echo "❌ PyInstaller build failed. Full log: /tmp/ship-build.log"
  tail -20 /tmp/ship-build.log
  exit 1
}
echo "✅ Mac app rebuilt → dist/Sigorta Hesabati.app"
echo

echo "================================================================"
echo "  Step 2/3 — Staging and committing changes"
echo "================================================================"
git add -A
if git diff --cached --quiet; then
  echo "ℹ️  No file changes to commit — skipping push."
  echo
  echo "Done. Your local Mac .app is rebuilt but GitHub already has the latest code."
  exit 0
fi
git commit -m "$MSG"
echo

echo "================================================================"
echo "  Step 3/3 — Pushing to GitHub (triggers Windows .exe build)"
echo "================================================================"
git push

echo
echo "================================================================"
echo "  ✅ Done"
echo "================================================================"
echo
echo "  Mac     →  dist/Sigorta Hesabati.app  (ready now)"
echo "  Windows →  https://github.com/IbrahimAli333/sigorta-hesabati/actions"
echo "             Wait ~5 min for the green ✓, then click the run and"
echo "             download the 'Sigorta-Hesabati-Windows' artifact."
echo
