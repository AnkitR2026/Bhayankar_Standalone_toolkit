#!/usr/bin/env python3
r"""Watch input\ and run the extract + gate + build + test workflow whenever a workbook lands or changes.

Leave this window open (Run_Watch.bat) and the workflow starts the moment a copied-in workbook is complete.
The result is output\<workbook>\WORKFLOW_STATUS.md.

  python bhayankar_watch.py            watch until Ctrl+C
  python bhayankar_watch.py --once     run whatever is waiting, then exit

Python 3 standard library only.
"""
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bhayankar_run as run  # noqa: E402

POLL_SECONDS = 3
SETTLE_SECONDS = 4  # a file still being copied (OneDrive, large .twbx) changes size: wait until it stops


def snapshot():
    wbs, tds, _ = run.find_inputs()
    return {p: (os.path.getsize(p), os.path.getmtime(p)) for p in [*wbs.values(), *tds]
            if os.path.exists(p)}


def process_waiting():
    for r in run.scan():
        if r['state'] in ('new', 'changed'):
            print(time.strftime('%H:%M:%S'), f'running the workflow for "{r["name"]}" ...', flush=True)
            st = run.run_one(r['name'])
            print(time.strftime('%H:%M:%S'), f'{r["name"]}: {st["verdict"]}  ->  output\\{r["name"]}\\{run.STATUS_MD}', flush=True)


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    os.makedirs(run.INPUT_DIR, exist_ok=True)
    process_waiting()
    if '--once' in sys.argv:
        return
    print(f'Watching {run.INPUT_DIR} for Tableau workbooks. Press Ctrl+C to stop.', flush=True)
    last, quiet_since = snapshot(), None
    try:
        while True:
            time.sleep(POLL_SECONDS)
            now = snapshot()
            if now != last:
                last, quiet_since = now, time.time()  # something changed: wait for it to finish copying
            elif quiet_since and time.time() - quiet_since >= SETTLE_SECONDS:
                quiet_since = None
                process_waiting()
    except KeyboardInterrupt:
        print('Stopped.')


if __name__ == '__main__':
    main()
