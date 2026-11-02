import re


def family(value):
    value = re.sub(r"\(.*", "", value).strip()
    if value in {"int", "integer", "bigint", "smallint", "int4", "int8"}:
        return "integer"
    if value in {"varchar", "nvarchar", "char", "text", "character varying"}:
        return "text"
    return value


def infer(model, config=None, observations=()):
    config = config or {}
    tables = {t["name"]: t for t in model["tables"]}
    columns = {n: {c["name"]: c for c in t["columns"]} for n, t in tables.items()}
    targets = {}
    for name in tables:
        targets.setdefault(name.rsplit("_", 1)[-1], []).append(name)
    edges = {}
    def add(src, col, dst, key, confidence, evidence):
        if src not in columns or dst not in columns or col not in columns[src] or key not in columns[dst]:
            return
        a, b = columns[src][col], columns[dst][key]
        if a["type"] and b["type"] and family(a["type"]) != family(b["type"]):
            return
        identity = (src, col, dst, key)
        edge = edges.setdefault(identity, dict(source_table=src, source_column=col,
            target_table=dst, target_column=key, confidence=confidence, evidence=[],
            cardinality="unknown", status="candidate"))
        edge["confidence"] = max(edge["confidence"], confidence)
        if evidence not in edge["evidence"]:
            edge["evidence"].append(evidence)
    for name, table in tables.items():
        parents = {r.get("parent") for r in table.get("references", []) if r.get("parent")}
        if table.get("parent"):
            parents.add(table["parent"])
        # Link tables are not always assigned a catalog parent. Infer the parent
        # only when the exact base name exists and its single primary key matches.
        if name.endswith("_lk") and name[:-3] in tables:
            parents.add(name[:-3])
        if config.get("kingdee_rules", True):
            for parent in sorted(parents):
                if parent not in tables: continue
                keys = [c["name"] for c in tables[parent]["columns"] if c["primary_key"]]
                if len(keys) != 1: continue
                key = keys[0]
                candidate = "fdataid" if name.endswith("_u") else key
                # Do not treat an entry's own key as a link to its main table.
                if candidate == "fentryid" and not (name.endswith("_r") or name.endswith("_lk")):
                    continue
                source = "name_suffix" if name.endswith("_lk") and parent == name[:-3] else "kingdee_catalog"
                add(name, candidate, parent, key, .7 if source == "name_suffix" else .8,
                    {"source": source, "detail": "parent table and matching single primary key; verify with model"})
        for col in table["columns"]:
            stem = re.sub(r"id$", "", col["name"]).lstrip("f").strip("_")
            if not col["name"].endswith("id") or len(stem) < 3:
                continue
            for dst in targets.get(stem, []):
                target = tables[dst]
                if dst == name:
                    continue
                keys = [c["name"] for c in target["columns"] if c["primary_key"]]
                if len(keys) == 1:
                    add(name, col["name"], dst, keys[0], .65, {"source": "naming", "detail": stem})
    for rel in config.get("relationships", []):
        args = [rel[k] for k in ("source_table", "source_column", "target_table", "target_column")]
        if any(t not in columns or c not in columns[t] for t, c in (args[:2], args[2:])):
            raise ValueError(f"Unknown configured relationship: {args}")
        add(*args, .95, {"source": "manual", "detail": rel.get("evidence", "configured by user")})
    for rel in observations:
        add(*(rel[k] for k in ("source_table", "source_column", "target_table", "target_column")),
            .75, {"source": "sql_join", "detail": rel["origin"]})
    minimum = config.get("min_confidence", .5)
    if not isinstance(minimum, (int, float)) or not 0 <= minimum <= 1:
        raise ValueError("min_confidence must be between 0 and 1")
    return [edges[k] for k in sorted(edges) if edges[k]["confidence"] >= minimum]
