import json
from datetime import datetime
from openpyxl import Workbook
import subprocess

# Load PR data from file
with open("prs.json", "r") as f:
    prs = json.load(f)

# Excel file setup
excel_file = "new_sheet.xlsx"
wb = Workbook()
ws = wb.active
ws.append(["Source Branch", "Author", "Action", "Comment", "Date", "Change ID"])

# Deadline to filter PRs
deadline = datetime.strptime("2025-06-25", "%Y-%m-%d")
logged_count = 0

for pr in prs:
    if not pr.get("mergedAt"):
        continue

    merged_date = datetime.strptime(pr["mergedAt"][:10], "%Y-%m-%d")
    if merged_date > deadline:
        continue

    source_branch = pr.get("headRefName")
    author = pr.get("author", {}).get("login")
    comment = pr.get("body") or "N/A"
    merged_at = pr["mergedAt"][:10]

    # Action detection (assume 'Squashed' if title or commit format hints so)
    if pr.get("title", "").lower().startswith("squash") or "[squash]" in pr.get("title", "").lower():
        action = "Squashed"
    else:
        action = "Merged"

    # Use sha as change ID (mocked fallback if unavailable)
    sha = pr.get("mergeCommit") or pr.get("mergeCommitSha") or "N/A"
    change_id = sha if sha != "N/A" else "N/A"

    ws.append([source_branch, author, action, comment, merged_at, change_id])
    logged_count += 1

wb.save(excel_file)
print(f"✅ Logged {logged_count} PRs into {excel_file}")
