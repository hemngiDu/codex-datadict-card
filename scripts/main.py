#!/usr/bin/env python3
"""Data Dictionary Card Generator - entry point"""
import sys, os, subprocess
path = sys.argv[1] if len(sys.argv) > 1 else None
if not path or not os.path.exists(path):
    print("Usage: python main.py <Excel file or directory>"); sys.exit(1)
if os.path.isdir(path):
    files = [f for f in os.listdir(path) if f.endswith(".xlsx") and not f.startswith("~$")]
    if not files: print("No xlsx found"); sys.exit(1)
    path = os.path.join(path, files[0])
print("Processing:", path)
base = os.path.dirname(os.path.abspath(__file__))
scripts = ["generate_sql.py", "generate_card.py", "generate_html_graph.py"]
for s in scripts:
    sp = os.path.join(base, s)
    r = subprocess.run([sys.executable, sp, path], capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0: print("ERROR:", r.stderr.strip())
print("\nDone!")