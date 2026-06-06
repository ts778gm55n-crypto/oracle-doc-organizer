#!/bin/bash
# Oracle Document Organizer
# =========================
# Scans INBOX folder, classifies, renames and moves Oracle documents
#
# Usage:
#   Manual:    bash /volume1/Claude/Scripts/oracle-doc-organizer/oracle_docs_organizer.sh
#   Dry run:   bash /volume1/Claude/Scripts/oracle-doc-organizer/oracle_docs_organizer.sh --dry-run

SCRIPT="/volume1/AI/Claude/Scripts/oracle-doc-organizer/organizer.py"
DOCS_ROOT="/volume3/Oracle Documents/OFC White Papers"

echo "========================================"
echo " Oracle Document Organizer"
echo " $(date '+%Y-%m-%d %H:%M:%S')"
echo "========================================"
echo ""

python3 "$SCRIPT" "$DOCS_ROOT" "$@"

echo ""
echo "Done."
