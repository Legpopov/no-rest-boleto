"""Resumable issue #2 collection. Checkpoint every completed route outside git.

Existing unrelated working-tree inputs are preserved. A campaign freezes its exact
scope and baseline, and refuses to resume under another scope. HTTP sessions keep
the existing per-host delays; at most eight independent workers are permitted.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
from pathlib import Path
import threading

from .branch_scope import endpoints
from .route_scope import route_scope
from .surface import ReceiptSession, collect_route, now


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    temporary.replace(path)


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--state-dir', type=Path, required=True)
    parser.add_argument('--workers', type=int, default=4, choices=range(1, 9))
    parser.add_argument('--out', type=Path, default=Path('data/surface/issue-2.json'))
    parser.add_argument('--shard', type=int, default=0)
    parser.add_argument('--of', type=int, default=1)
    args = parser.parse_args()
    if args.of < 1 or not 0 <= args.shard < args.of:
        parser.error('Invalid shard')
    registry = endpoints()
    scope = route_scope(registry)
    scope['shard'] = args.shard
    scope['of'] = args.of
    args.state_dir.mkdir(parents=True, exist_ok=True)
    frozen = args.state_dir/'scope.json'
    # JSON roundtrip normalizes tuples before comparing a resumed scope.
    scope = json.loads(json.dumps(scope, ensure_ascii=False))
    if frozen.exists():
        previous = read(frozen)
        previous.setdefault('shard', 0)
        previous.setdefault('of', 1)
        if previous != scope:
            raise ValueError('Campaign scope changed; start a separate campaign')
    else:
        write(frozen, scope)
        for name in ('planning-coverage.json', 'surface/latest.json', 'air-current/latest.json'):
            source = Path('data')/name
            write(args.state_dir/('baseline-'+source.parent.name+'-'+source.name), read(source))
        write(args.state_dir/'started.json', dict(started_at=now()))
    journal = args.state_dir/'surface.jsonl'
    rows = {}
    if journal.exists():
        # A write interrupted before its newline is not a committed checkpoint.
        content = journal.read_bytes()
        if content and not content.endswith(b'\n'):
            content = content[:content.rfind(b'\n')+1]
            journal.write_bytes(content)
        for line in content.decode('utf-8').splitlines():
            row = json.loads(line)
            key = (row['origin'], row['destination'], row['mode'])
            if key in rows:
                raise ValueError('Duplicate campaign result')
            rows[key] = row
    requested = [(a,b,m) for a,b in scope['directed_pairs'] for m in ('rail','bus')][args.shard::args.of]
    local = threading.local()
    def collect(key):
        if not hasattr(local, 'session'):
            local.session = ReceiptSession(args.state_dir/'raw')
        return collect_route(local.session, *key, registry)
    pending = [key for key in requested if key not in rows]
    print(f'Resuming {len(rows)}/{len(requested)}; workers={args.workers}', flush=True)
    with journal.open('a', encoding='utf-8') as log, ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(collect, key): key for key in pending}
        for future in as_completed(futures):
            key = futures[future]
            row = future.result()
            log.write(json.dumps(row, ensure_ascii=False)+'\n')
            log.flush()
            rows[key] = row
            if len(rows) % 50 == 0 or len(rows) == len(requested):
                print(f'{now()} {len(rows)}/{len(requested)}', flush=True)
    write(args.out, dict(schema_version=2, complete=True,
                        started_at=read(args.state_dir/'started.json')['started_at'], updated_at=now(),
                        scope=sorted(set(scope['origins']) | set(scope['destinations'])),
                        requested_routes=requested, routes=[rows[key] for key in requested]))


if __name__ == '__main__':
    main()
