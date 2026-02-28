#!/usr/bin/env bash
set -euo pipefail

SOURCE_DIR="${1:-$(pwd)}"
TARGET_DIR="${2:-/workspace/DANFES}"

if [[ ! -d "$SOURCE_DIR" ]]; then
  echo "Source directory not found: $SOURCE_DIR" >&2
  exit 1
fi

mkdir -p "$TARGET_DIR"

rsync -a \
  --exclude '.git' \
  --exclude 'DANFES' \
  --exclude 'node_modules' \
  --exclude '__pycache__' \
  --exclude '*.pyc' \
  "$SOURCE_DIR/" "$TARGET_DIR/"

echo "DANFES project created at: $TARGET_DIR"
