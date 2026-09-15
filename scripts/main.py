#!/usr/bin/env python3
"""Generate the legacy outputs plus semantic metadata for each workbook."""
import argparse
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path")
    args = parser.parse_args()
    path = Path(args.path).resolve()
    files = sorted(p for p in path.glob("*.xlsx") if not p.name.startswith("~$")) if path.is_dir() else [path]
    if not files or any(not p.is_file() for p in files):
        parser.error("No Excel files found")
    scripts = Path(__file__).resolve().parent
    for file in files:
        for name in ("generate_sql.py", "generate_card.py", "generate_search_graph.py"):
            # Separate graph directories prevent multi-workbook index/chunk collisions.
            if name == "generate_search_graph.py":
                import shutil
                target = file.parent / (file.stem + "_graph")
                target.mkdir(exist_ok=True)
                import tempfile
                with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=target, delete=False) as temporary:
                    graph_input = Path(temporary.name)
                shutil.copy2(file, graph_input)
                try:
                    subprocess.run([sys.executable, str(scripts / name), str(graph_input)], check=True)
                finally:
                    graph_input.unlink()
            else:
                subprocess.run([sys.executable, str(scripts / name), str(file)], check=True)
        subprocess.run([sys.executable, str(scripts / "semantic.py"), "build", str(file),
                        "--output", str(file.with_name(file.stem + "_semantic.json"))], check=True)
    print("Done")


if __name__ == "__main__":
    main()
