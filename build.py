#!/usr/bin/env python3
"""Build the maintenance dashboard from a SAP PM xlsx export."""
import sys, json, openpyxl
from datetime import datetime, timedelta

def build(xlsx_path: str, out_path: str = 'index.html'):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb['SAPUI5 Export']
    events = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        notif, asset, start, end, dur, txt, order = r
        if not asset:
            continue
        events.append({
            'notification': notif,
            'asset': asset,
            'start': start.isoformat() if start else None,
            'end': end.isoformat() if end else None,
            'duration_hrs': dur,
            'text': (txt or '').replace('\n', ' | '),
            'order': order,
            'still_down': end is None,
        })
    latest = max(datetime.fromisoformat(e['start']) for e in events)
    now = (latest + timedelta(hours=6)).isoformat()
    payload = json.dumps({'events': events, 'now': now})

    template_path = '/home/turbo/.hermes/skills/productivity/sap-pm-maintenance-dashboard/templates/dashboard.html'
    with open(template_path) as f:
        html = f.read()
    html = html.replace('/*DATA_PLACEHOLDER*/null', payload)
    with open(out_path, 'w') as f:
        f.write(html)
    print(f'Wrote {out_path}: {len(events)} events, now={now}')

if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else 'notifications.xlsx',
          sys.argv[2] if len(sys.argv) > 2 else 'index.html')
