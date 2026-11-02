import re


def tokens(text):
    text = text.casefold()
    result = set(re.findall(r"[a-z0-9_]+", text))
    for phrase in re.findall(r"[\u3400-\u9fff]+", text):
        result.add(phrase)
        result.update(phrase[i:i+2] for i in range(len(phrase)-1))
    return result


def search(model, query, config=None, limit=10):
    if limit < 1:
        raise ValueError("limit must be positive")
    terms = tokens(query)
    if not terms:
        return []
    config = config or {}
    for alias, expansions in config.get("synonyms", {}).items():
        if alias.casefold() in query.casefold():
            for expansion in expansions:
                terms.update(tokens(expansion))
    results = []
    for table in model["tables"]:
        matches, score = [], 0
        for field, weight in (("name", 8), ("business_name", 6), ("description", 4)):
            value = " ".join(str(ref.get(field, "")) for ref in table.get("references", [])) + " " + table.get(field, "")
            value = value.casefold()
            hits = sorted(t for t in terms if t in value)
            if hits:
                score += weight * len(hits)
                matches.append(dict(field=field, terms=hits))
        columns = []
        for col in table["columns"]:
            value = " ".join(str(col.get(k, "")) for k in ("name", "description", "remarks")).casefold()
            hits = sorted(t for t in terms if t in value)
            if hits:
                columns.append(dict(name=col["name"], terms=hits))
        score += min(12, sum(len(c["terms"]) * 2 for c in columns))
        if score:
            results.append(dict(table=table["name"], business_name=table["business_name"],
                                score=score, matches=matches, columns=columns))
    return sorted(results, key=lambda r: (-r["score"], r["table"]))[:limit]
