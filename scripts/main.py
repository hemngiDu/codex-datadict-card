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

print("\n[1/3] Generating SQL comments...")
r = subprocess.run([sys.executable, os.path.join(base, "generate_sql.py"), path], capture_output=True, text=True)
print(r.stdout.strip())

print("[2/3] Generating knowledge card...")
r = subprocess.run([sys.executable, os.path.join(base, "generate_card.py"), path], capture_output=True, text=True)
print(r.stdout.strip())

print("[3/3] Generating search graph (on-demand loading)...")
r = subprocess.run([sys.executable, os.path.join(base, "generate_search_graph.py"), path], capture_output=True, text=True)
print(r.stdout.strip())

print("\nDone! Files generated alongside your Excel file.")
