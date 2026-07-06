# 数据字典速查卡 (Data Dictionary Card)

从金蝶云苍穹数据字典导出 Excel 中，一键生成三件套。

## 功能

1. **SQL 注释** — comment on table / comment on column 语句
2. **知识卡片 (Markdown)** — 含 Mermaid ER 图的业务对象目录和关联模型
3. **交互式 HTML 图谱** — D3.js 力导向图，节点连线展示表间关联

## 使用方法

```
pip install -r requirements.txt
python scripts/generate_all.py <Excel文件路径或目录>
```

## 输出

| 文件 | 说明 |
|---|---|
| *_comments.sql | PostgreSQL 注释 SQL |
| *_datadict_card.md | Mermaid ER 图 + 业务对象目录 + 表索引 |
| *_graph.html | D3.js 交互式力导向图 |

## 依赖

- Python 3.7+
- openpyxl

## 数据模型

数据字典遵循固定层级模式：主表 → 多语言表(_l) → 使用范围表(_u) → 子表(entry) → 多选基础资料表
