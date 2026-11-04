#!/usr/bin/env python3
"""Generate the legacy outputs plus semantic metadata for each workbook."""
import argparse
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path")
    parser.add_argument("--sql", nargs="*", default=[])
    parser.add_argument("--dialect", default=None)
    parser.add_argument("--config")
    args = parser.parse_args()
    path = Path(args.path).resolve()
    files = sorted(p for p in path.glob("*.xlsx") if not p.name.startswith("~$")) if path.is_dir() else [path]
    if not files or any(not p.is_file() for p in files):
        parser.error("No Excel files found")
    scripts = Path(__file__).resolve().parent
    for file in files:
        model = file.with_name(file.stem + "_semantic.json")
        command = [sys.executable, str(scripts / "semantic.py"), "build", str(file), "--output", str(model)]
        if args.config: command += ["--config", args.config]
        if args.dialect: command += ["--dialect", args.dialect]
        if args.sql: command += ["--sql", *args.sql]
        subprocess.run(command, check=True)
        subprocess.run([sys.executable, str(scripts / "generate_sql.py"), str(file)], check=True)
    print("Done")


if __name__ == "__main__":
    main()
