import json
from pathlib import Path
import requests

tokens = json.loads(Path("/home/ross/.local/lib/ms-graph-explorer/session_state/tokens.json").read_text())
token = tokens["tokens"]["outlook.office.com"]

resp = requests.get(
    "https://outlook.office.com/api/v2.0/me/tasks",
    headers={"Authorization": f"Bearer {token}"},
    timeout=15,
)
print(f"Status: {resp.status_code}")
if resp.status_code == 200:
    tasks = resp.json().get("value", [])
    print(f"Tasks found: {len(tasks)}")
    for t in tasks[:5]:
        print(f"- {t.get('Subject')} (Status: {t.get('Status')})")
else:
    print(resp.text[:500])
