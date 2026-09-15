# Codex 数据库语义理解 Skill

从金蝶云苍穹 Excel 数据字典查找表、定位字段、读取局部上下文并发现关联候选，兼容原有 SQL Comment、Markdown/Mermaid ER 和 D3 图谱。

## 给 Codex 使用（推荐）

安装到 Codex 后，把金蝶数据字典 Excel 文件交给 Codex，并直接说：

> 这是供应链数据字典。帮我找“采购订单的供应商、物料和采购数量”涉及哪些表、字段和关联；标出推断关系的证据。

Codex 会自行读取 Excel、建立语义模型、搜索候选表并解释结果。只有你要求“生成卡片/图谱/注释 SQL”时才生成这些文件。无须记住 build、search、context 命令。关联是候选，实际 SQL 还需核对业务条件。

首次安装：把此仓库放到 Codex 的 `skills/codex-datadict-card` 目录，让 Codex 重新加载技能。你也可以直接让 Codex 帮你安装。安装一次后，日常只需提供 Excel 和问题；Python 依赖由 Codex 按 Skill 指引处理。

## 命令行使用（可选）

Python 3.10+：

```bash
pip install -r requirements.txt
python examples/make_sample.py demo/sample.xlsx
python scripts/generate_all.py demo/sample.xlsx
python scripts/semantic.py search demo/sample_semantic.json "采购数量" --limit 5
python scripts/semantic.py context demo/sample_semantic.json t_pm_purorder
```

样例为合成数据。检索“采购数量”优先返回采购分录；订单上下文提供供应商和分录候选邻居。

## 能力

- 中文业务词、表名、字段名、说明和备注检索；可配置同义词，返回命中依据。
- 金蝶目录父子关系、字段命名和目标单列主键推断；过滤已知类型冲突，输出 confidence/evidence。
- 基础 SQL JOIN ON 等值条件抽取；支持别名及 schema 名，使用 SQLGlot 语法树，SQL 不会执行。
- context 只输出指定表完整字段及直接关系，供 Codex 分阶段读取。
- 旧生成器继续可用；补齐文档中的 generate_all.py 入口，失败不再被忽略。

```bash
python scripts/semantic.py build demo/sample.xlsx --output demo/model.json --config examples/config.json --sql examples/joins.sql --dialect postgres
```

关系评分是启发式排序值，不是经过数据校验的概率。历史 JOIN 不证明外键、方向或基数；输出均标为 candidate/unknown。复杂 SQL（CTE、子查询、UNION、DML）整句跳过并报告 warnings。详见 [模型与扩展说明](references/semantic-model.md)。

## 兼容输出

| 入口 | 输出 |
|---|---|
| generate_all.py / main.py | 每个 Excel 的 SQL、Markdown、语义 JSON 和独立 `<名称>_graph/` 目录 |
| generate_sql.py | Excel 同目录 `*_comments.sql`（PostgreSQL） |
| generate_card.py | Excel 同目录 `*_datadict_card.md`，含 Mermaid ER |
| generate_search_graph.py | Excel 同目录 graph.html、index.json、chunks/、columns/ |

`generate_all.py` 也接受目录，按文件名顺序处理全部 xlsx，忽略 `~$` 临时文件。独立图谱目录防止多工作簿输出互相覆盖。图谱运行：

```bash
cd demo/sample_graph
python -m http.server 8080
```

浏览器访问 http://localhost:8080/graph.html，需要联网加载 D3。其他历史图谱脚本保留。旧 ER/D3 的库存跨表关系仍是历史规则，不应视为经过验证的关系。

## 目录与测试

```text
scripts/datadict/       元数据适配器、检索、关系推断、SQL 历史适配器
scripts/semantic.py     build / search / context JSON 接口
scripts/generate_*.py   原有产物生成器及兼容入口
examples/              可重复生成的合成工作簿、配置、JOIN 示例
tests/                 单元与端到端兼容验证
references/            数据契约与扩展说明
```

```bash
python -m unittest discover -s tests -v
```

测试覆盖候选表排序、字段命中、类型冲突、复合主键保守处理、SQL 作用域与表达式过滤及 Excel→SQL/Markdown/D3/语义 JSON。当前未用真实客户工作簿或数据库验证。检索逐表扫描，内存保存元数据，尚未对 5 万表做性能承诺；大型部署可按同一模型接口引入倒排索引。
