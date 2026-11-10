"""Archive source observations; collection time and travel dates stay separate.

python -m collector.history --git-ref origin/main
python -m collector.history data/prices.json
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import datetime
import gzip
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent.parent
HISTORY = ROOT / "data/history"
MODES = ("air", "rail", "bus", "hotels")
SOURCES = {"air": "aviasales", "rail": "tutu", "bus": "atlasbus", "hotels": "ostrovok"}


def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def observations(snapshot):
    """Only successful, newly observed source blocks; stale copies add no observations."""
    seen = snapshot["generated_at"]
    seen_dt = datetime.fromisoformat(seen.replace("Z", "+00:00"))
    if seen_dt.tzinfo is None:
        raise ValueError("generated_at must include a timezone")
    origin = snapshot.get("origin")
    if not origin or not origin.get("city"):
        raise ValueError("Snapshot has no explicit origin")
    for city in snapshot.get("cities", []):
        dest = {k: city.get(k) for k in ("id", "city", "country", "country_code", "abroad", "lat", "lon")}
        for mode in MODES:
            block = city.get(mode) or {}
            if block.get("status") != "ok" or block.get("stale"):
                continue
            at = block.get("updated_at")
            if not at:
                continue  # no fabricated collection date for legacy data
            at_dt = datetime.fromisoformat(at.replace("Z", "+00:00"))
            if at_dt.tzinfo is None:
                raise ValueError(f"Invalid observation time: {at}")
            data = {k: v for k, v in block.items() if k not in ("updated_at", "stale", "last_attempt")}
            # Hotel prices belong to a destination, independent of the departure city.
            row = {"observed_at": at_dt.isoformat(timespec="seconds"),
                   "origin": None if mode == "hotels" else origin,
                   "destination": dest, "mode": mode, "currency": "RUB",
                   "source": block.get("source") or SOURCES[mode], "data": data}
            row["id"] = hashlib.sha256(canonical(row).encode()).hexdigest()
            # Legacy generated_at was set at collection start, before some blocks finished.
            row["first_seen_at"] = max(seen_dt, at_dt).isoformat(timespec="seconds")
            yield row


def load_history(directory=HISTORY):
    rows = {}
    for path in sorted(Path(directory).glob("????-??.jsonl.gz")):
        with gzip.open(path, "rt", encoding="utf-8") as f:
            for line in f:
                row = json.loads(line)
                rows[row["id"]] = row
    return rows


def archive(snapshots, directory=HISTORY):
    directory = Path(directory)
    rows = load_history(directory)
    before = len(rows)
    touched = set()
    for snapshot in snapshots:
        for row in observations(snapshot):
            if row["id"] not in rows:
                rows[row["id"]] = row
                touched.add(row["observed_at"][:7])
    directory.mkdir(parents=True, exist_ok=True)
    months = defaultdict(list)
    for row in rows.values():
        months[row["observed_at"][:7]].append(row)
    for month in sorted(touched):
        data = "".join(canonical(r) + "\n" for r in sorted(months[month], key=lambda r: r["id"]))
        target = directory / f"{month}.jsonl.gz"
        temp = target.with_suffix(target.suffix + ".tmp")
        temp.write_bytes(gzip.compress(data.encode(), mtime=0))
        temp.replace(target)
    index = {"schema_version": 1, "observations": len(rows),
             "first_observed_at": min((r["observed_at"] for r in rows.values()), default=None),
             "last_observed_at": max((r["observed_at"] for r in rows.values()), default=None),
             "months": [{"observed_month": m, "file": f"{m}.jsonl.gz", "observations": len(rs),
                         "by_mode": dict(sorted(Counter(r["mode"] for r in rs).items()))}
                        for m, rs in sorted(months.items())],
             "note": "Months describe collection dates, not a complete history of travel-month prices."}
    target = directory / "index.json"
    temp = target.with_suffix(".json.tmp")
    temp.write_text(canonical(index) + "\n", encoding="utf-8")
    temp.replace(target)
    return {"added": len(rows) - before, **index}


def git_snapshots(ref, root=ROOT):
    """Replay immutable file revisions, oldest first; never execute commands from data."""
    root = Path(root).resolve()
    prefix = ["git", "-c", f"safe.directory={root.as_posix()}", "-C", str(root)]
    def git(*args):
        return subprocess.check_output([*prefix, *args])
    commit = git("rev-parse", "--verify", f"{ref}^{{commit}}").decode().strip()
    commits = git("log", "--reverse", "--format=%H", commit, "--", "data/prices.json").decode().splitlines()
    last_blob = None
    for sha in commits:
        blob = git("rev-parse", f"{sha}:data/prices.json").decode().strip()
        if blob != last_blob:
            yield json.loads(git("show", f"{sha}:data/prices.json"))
            last_blob = blob


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("snapshots", nargs="*", type=Path)
    ap.add_argument("--git-ref")
    ap.add_argument("--out", type=Path, default=HISTORY)
    args = ap.parse_args()
    def inputs():
        if args.git_ref:
            yield from git_snapshots(args.git_ref)
        for path in args.snapshots:
            yield json.loads(path.read_text(encoding="utf-8"))
    if not args.git_ref and not args.snapshots:
        ap.error("Supply snapshot paths or --git-ref")
    print(json.dumps(archive(inputs(), args.out), ensure_ascii=False))


if __name__ == "__main__":
    main()
