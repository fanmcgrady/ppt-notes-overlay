#!/usr/bin/env bash
set -euo pipefail
cd "/mnt/c/Users/fangzy/Desktop/AI工作区/ppt-notes-overlay"

git add .
git status

cat > /tmp/ppt-notes-commitmsg.txt <<'EOF'
Initial commit: PowerPoint notes overlay for defense recording

Floating speaker notes, rehearsal timer, Word export, and exe build scripts.
EOF

git commit -F /tmp/ppt-notes-commitmsg.txt
git log -1 --format='%h %an %ae %s'
git status
