from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
from pathlib import Path

from common import (
    active_profile,
    active_record,
    deep_merge,
    flatten,
    load_preset,
    read_json,
    set_dotted,
    state_root,
    validate_override,
    write_json_atomic,
)


def paths() -> tuple[Path, Path, Path]:
    root = state_root()
    return root / "active.json", root / "draft.json", root / "history"


def make_record(overrides: dict) -> dict:
    return {
        "schema_version": 1,
        "base_profile": load_preset()["profile_id"],
        "overrides": overrides,
    }


def archive_active() -> None:
    active, _, history = paths()
    if not active.exists():
        return
    history.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    shutil.copy2(active, history / f"profile-{stamp}.json")
    files = sorted(history.glob("profile-*.json"), reverse=True)
    for old in files[5:]:
        old.unlink()


def command_view(_: argparse.Namespace) -> None:
    record = active_record()
    print(json.dumps({
        "source": "user_profile" if record else "built_in_preset",
        "profile": active_profile(),
    }, ensure_ascii=False, indent=2))


def command_begin(args: argparse.Namespace) -> None:
    _, draft, _ = paths()
    if draft.exists() and not args.replace:
        raise SystemExit("A draft already exists; use diff, confirm, cancel, or begin --replace")
    current = active_record()
    overrides = {} if args.fresh or not current else current.get("overrides", {})
    write_json_atomic(draft, make_record(overrides))
    print(json.dumps({"status": "draft_created", "fresh": bool(args.fresh)}, ensure_ascii=False))


def command_set(args: argparse.Namespace) -> None:
    _, draft, _ = paths()
    if not draft.exists():
        raise SystemExit("No draft exists; run begin first")
    record = read_json(draft)
    try:
        value = json.loads(args.value)
    except json.JSONDecodeError:
        value = args.value
    set_dotted(record.setdefault("overrides", {}), args.path, value)
    validate_override(record["overrides"])
    write_json_atomic(draft, record)
    print(json.dumps({"status": "draft_updated", "path": args.path, "value": value}, ensure_ascii=False))


def command_diff(_: argparse.Namespace) -> None:
    _, draft, _ = paths()
    if not draft.exists():
        raise SystemExit("No draft exists")
    before = flatten(active_profile())
    record = read_json(draft)
    after = flatten(deep_merge(load_preset(), record.get("overrides", {})))
    changes = [
        {"field": key, "current": before.get(key), "proposed": after.get(key)}
        for key in sorted(set(before) | set(after))
        if before.get(key) != after.get(key)
    ]
    print(json.dumps({"changes": changes}, ensure_ascii=False, indent=2))


def command_confirm(_: argparse.Namespace) -> None:
    active, draft, _ = paths()
    if not draft.exists():
        raise SystemExit("No draft exists")
    record = read_json(draft)
    validate_override(record.get("overrides", {}))
    archive_active()
    write_json_atomic(active, record)
    draft.unlink()
    print(json.dumps({"status": "saved"}, ensure_ascii=False))


def command_cancel(_: argparse.Namespace) -> None:
    _, draft, _ = paths()
    if draft.exists():
        draft.unlink()
    print(json.dumps({"status": "draft_cancelled"}, ensure_ascii=False))


def command_restore(_: argparse.Namespace) -> None:
    active, draft, _ = paths()
    archive_active()
    if active.exists():
        active.unlink()
    if draft.exists():
        draft.unlink()
    print(json.dumps({"status": "built_in_preset_restored"}, ensure_ascii=False))


def command_history(_: argparse.Namespace) -> None:
    _, _, history = paths()
    items = []
    for path in sorted(history.glob("profile-*.json"), reverse=True) if history.exists() else []:
        items.append({"version": path.name, "record": read_json(path)})
    print(json.dumps({"history": items}, ensure_ascii=False, indent=2))


def command_rollback(args: argparse.Namespace) -> None:
    active, draft, history = paths()
    source = history / Path(args.version).name
    if not source.exists() or source.parent.resolve() != history.resolve():
        raise SystemExit("History version not found")
    record = read_json(source)
    validate_override(record.get("overrides", {}))
    archive_active()
    write_json_atomic(active, record)
    if draft.exists():
        draft.unlink()
    print(json.dumps({"status": "rolled_back", "version": source.name}, ensure_ascii=False))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage the persistent formatting-only profile")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("view").set_defaults(func=command_view)
    begin = sub.add_parser("begin")
    begin.add_argument("--fresh", action="store_true")
    begin.add_argument("--replace", action="store_true")
    begin.set_defaults(func=command_begin)
    setter = sub.add_parser("set")
    setter.add_argument("path")
    setter.add_argument("value")
    setter.set_defaults(func=command_set)
    sub.add_parser("diff").set_defaults(func=command_diff)
    sub.add_parser("confirm").set_defaults(func=command_confirm)
    sub.add_parser("cancel").set_defaults(func=command_cancel)
    sub.add_parser("restore").set_defaults(func=command_restore)
    sub.add_parser("history").set_defaults(func=command_history)
    rollback = sub.add_parser("rollback")
    rollback.add_argument("version")
    rollback.set_defaults(func=command_rollback)
    return parser


if __name__ == "__main__":
    arguments = build_parser().parse_args()
    try:
        arguments.func(arguments)
    except (KeyError, TypeError, ValueError) as exc:
        raise SystemExit(str(exc)) from exc
