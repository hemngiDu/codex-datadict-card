"""Conservative SQL history adapter: flat SELECT JOIN ON equalities only."""
import sqlglot
from sqlglot import exp


def extract(sql, origin="inline", dialect=None):
    observations, warnings = [], []
    for index, tree in enumerate(sqlglot.parse(sql, read=dialect), 1):
        if tree is None:
            continue
        if not isinstance(tree, exp.Select) or tree.args.get("with_") or any(tree.find_all(exp.Subquery)):
            warnings.append(f"{origin}:{index}: skipped non-flat SELECT (CTE/subquery/union/DML)")
            continue
        aliases = {}
        for table in tree.find_all(exp.Table):
            aliases[table.alias_or_name.lower()] = ".".join(p.name.lower() for p in table.parts)
        for join in tree.find_all(exp.Join):
            on = join.args.get("on")
            if on is None or any(on.find_all(exp.Or)):
                warnings.append(f"{origin}:{index}: skipped JOIN without simple conjunctive ON")
                continue
            def conjuncts(node):
                if isinstance(node, exp.Paren):
                    yield from conjuncts(node.this)
                elif isinstance(node, exp.And):
                    yield from conjuncts(node.left)
                    yield from conjuncts(node.right)
                else:
                    yield node
            for eq in conjuncts(on):
                if not isinstance(eq, exp.EQ) or not isinstance(eq.left, exp.Column) or not isinstance(eq.right, exp.Column):
                    continue
                a, b = eq.left, eq.right
                if a.table.lower() not in aliases or b.table.lower() not in aliases or a.table.lower() == b.table.lower():
                    continue
                observations.append(dict(source_table=aliases[a.table.lower()], source_column=a.name.lower(),
                    target_table=aliases[b.table.lower()], target_column=b.name.lower(), origin=f"{origin}:{index}"))
    return observations, warnings
