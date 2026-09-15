#!/usr/bin/env python3
"""Codex-friendly JSON CLI; no database connection or SQL execution."""
import argparse
import json
import sys
from pathlib import Path
from datadict.metadata import read_excel
from datadict.relations import infer
from datadict.search import search


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build")
    build.add_argument("excel")
    build.add_argument("--output", required=True)
    build.add_argument("--config")
    build.add_argument("--sql", nargs="*", default=[])
    build.add_argument("--dialect", default=None)
    query = sub.add_parser("search")
    query.add_argument("model")
    query.add_argument("query")
    query.add_argument("--limit", type=int, default=10)
    context = sub.add_parser("context")
    context.add_argument("model")
    context.add_argument("tables", nargs="+")
    args = parser.parse_args()
    try:
        if args.command == "build":
            model = read_excel(args.excel)
            config = load(args.config) if args.config else {}
            if not isinstance(config, dict):
                raise ValueError("config must be a JSON object")
            synonyms = config.get("synonyms", {})
            if not isinstance(synonyms, dict) or any(not isinstance(v, list) or any(not isinstance(x, str) for x in v) for v in synonyms.values()):
                raise ValueError("synonyms must map strings to string arrays")
            if not isinstance(config.get("kingdee_rules", True), bool):
                raise ValueError("kingdee_rules must be a boolean")
            observations, warnings = [], []
            if args.sql:
                from datadict.sql_history import extract
                for path in args.sql:
                    from sqlglot.errors import SqlglotError
                    try:
                        found, notes = extract(Path(path).read_text(encoding="utf-8-sig"), Path(path).name, args.dialect)
                    except SqlglotError as error:
                        raise ValueError(f"Cannot parse SQL file {Path(path).name}") from error
                    observations.extend(found)
                    warnings.extend(notes)
            model.update(config=config, relationships=infer(model, config, observations), warnings=warnings)
            Path(args.output).parent.mkdir(parents=True, exist_ok=True)
            Path(args.output).write_text(json.dumps(model, ensure_ascii=False, indent=2), encoding="utf-8")
            result = dict(output=args.output, tables=len(model["tables"]), relationships=len(model["relationships"]), warnings=warnings)
        else:
            model = load(args.model)
            if args.command == "search":
                result = search(model, args.query, model.get("config"), args.limit)
            else:
                names = {name.lower() for name in args.tables}
                available = {t["name"] for t in model["tables"]}
                if names - available:
                    raise ValueError(f"Unknown tables: {sorted(names - available)}")
                edges = [r for r in model["relationships"] if r["source_table"] in names or r["target_table"] in names]
                neighbors = {r[k] for r in edges for k in ("source_table", "target_table")} - names
                result = dict(tables=[t for t in model["tables"] if t["name"] in names],
                              relationships=edges, candidate_neighbors=sorted(neighbors))
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (ValueError, OSError, KeyError, ImportError) as error:
        parser.exit(2, f"Error: {error}\n")


if __name__ == "__main__":
    main()
