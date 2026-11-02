#!/usr/bin/env python3
"""Export a semantic model as small, searchable knowledge files."""
import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("model")
    parser.add_argument("output")
    args = parser.parse_args()
    model = json.loads(Path(args.model).read_text(encoding="utf-8-sig"))
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    tables = sorted(model["tables"], key=lambda t: t["name"])
    edges = model.get("relationships", [])
    conflicts = model.get("metadata_conflicts", [])
    index = []
    for offset in range(0, len(tables), 50):
        chunk = tables[offset:offset+50]
        names = {t["name"] for t in chunk}
        filename = f"tables-{offset//50+1:03}.json"
        data = {"source": model["source"], "tables": chunk,
                "relationships": [r for r in edges if r["source_table"] in names or r["target_table"] in names],
                "metadata_conflicts": [c for c in conflicts if c["table"] in names]}
        (out / filename).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        for t in chunk:
            index.append({"name":t["name"], "references":t.get("references", []), "file":filename})
    (out / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "README.md").write_text("""# 供应链云数据字典知识库

来源：金蝶数据字典导出 Excel。先搜索 index.json 的表名或业务对象，再读取对应 tables-NNN.json。每个文件包含表字段、业务对象引用、相关候选关系及字段定义差异。字段差异保留原始工作表名称，需结合原始导出核对。

关系为规则推断候选，不是已验证的外键或基数。confidence 是规则分数而非概率。生成 SQL 前须核对连接字段、组织、租户、语言和业务状态；不能确定时明确说明。表目录父子归属不等于外键。没有 SQL 历史或数据库约束作为证据。

相同物理表可属于多个业务对象：references 保存各目录引用；顶层业务名称与父表是首条记录，不代表唯一归属。字段首次出现的定义在 columns，冲突的其他定义在 metadata_conflicts；字段缺失不应猜测。
""", encoding="utf-8")
    print(f"Exported {len(tables)} tables in {(len(tables)+49)//50} files")

if __name__ == "__main__":
    main()
