#!/usr/bin/env bash
# ppt-intake skill installer (cross-agent)
# Usage:  ./install.sh [target...]      no args = list targets
#         ./install.sh claude-user zcode-user
#         ./install.sh all
# Policy: an existing ppt-intake install at the target is REPLACED.
#         Junctions/symlinks are detected and left untouched (never delete through a link).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
SRC="$ROOT/skills/ppt-intake"
[ -f "$SRC/SKILL.md" ] || { echo "ERROR: $SRC/SKILL.md not found"; exit 1; }

DESTS=(
  "zcode-project|$ROOT/.zcode/skills"
  "zcode-user|$HOME/.zcode/skills"
  "claude-project|$ROOT/.claude/skills"
  "claude-user|$HOME/.claude/skills"
  "agents-project|$ROOT/.agents/skills"
  "agents-user|$HOME/.agents/skills"
)

is_link() {
  [ -L "$1" ] && return 0
  command -v cygpath >/dev/null 2>&1 || return 1
  fsutil reparsepoint query "$(cygpath -w "$1")" >/dev/null 2>&1
}

install_to() {
  local dest_root="$1" dest
  dest="$dest_root/ppt-intake"
  if [ -e "$dest" ] || [ -L "$dest" ]; then
    if is_link "$dest"; then
      echo "SKIP (junction/symlink, managed elsewhere): $dest"
      return
    fi
    rm -rf "$dest"
  fi
  mkdir -p "$dest_root"
  cp -R "$SRC" "$dest"
  echo "installed -> $dest"
}

usage() {
  echo "Install targets:"
  local e
  for e in "${DESTS[@]}"; do
    printf '  %-16s %s\n' "${e%%|*}" "${e#*|}"
  done
  echo "Usage: ./install.sh <target...> | all"
}

[ $# -eq 0 ] && { usage; exit 0; }

if [ "$1" = "all" ]; then
  set -- "${DESTS[@]%%|*}"
fi

for want in "$@"; do
  matched=""
  for entry in "${DESTS[@]}"; do
    name="${entry%%|*}"; root="${entry#*|}"
    if [ "$want" = "$name" ]; then
      install_to "$root"; matched=1
    fi
  done
  if [ -z "$matched" ]; then
    echo "unknown target: $want"; usage; exit 1
  fi
done

echo "done. dependency: pip install python-pptx"
