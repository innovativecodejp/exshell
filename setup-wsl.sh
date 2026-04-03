#!/usr/bin/env bash
# Exshell WSL setup script
# Source this file from your shell profile or current shell.
# Example: source /mnt/d/dev/exshell/setup-wsl.sh

EXSHELL_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXSHELL_BIN="$EXSHELL_ROOT/src/Exshell/bin/Release/net8.0-windows/win-x64/publish/exshell.exe"

if [[ ! -f "$EXSHELL_BIN" ]]; then
  echo "warning: exshell.exe not found: $EXSHELL_BIN" >&2
  echo "run 'dotnet publish -c Release -r win-x64 --self-contained false' first." >&2
  return 0 2>/dev/null || exit 0
fi

eopen() { "$EXSHELL_BIN" eopen "$@"; }
els()   { "$EXSHELL_BIN" els "$@"; }
ecat()  { "$EXSHELL_BIN" ecat "$@"; }
cate()  { "$EXSHELL_BIN" cate "$@"; }
ediff() { "$EXSHELL_BIN" ediff "$@"; }
einfo() { "$EXSHELL_BIN" einfo "$@"; }

echo "Exshell loaded. Commands: eopen, els, ecat, cate, ediff, einfo"
