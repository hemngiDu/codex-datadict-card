---
name: codex-datadict-card
description: 从金蝶云苍穹 Excel 数据字典发现候选表、检索字段、读取局部数据库上下文，并生成带证据的关系候选、Markdown/ER、SQL 注释和 D3 图谱。适用于理解已有数据模型和辅助 SQL 编写。
---

# 数据库语义理解

在 Skill 所在目录运行命令；输入路径可使用绝对路径。先安装 `pip install -r requirements.txt`。

## 发现表与字段

1. 对金蝶导出建立模型：
   `python scripts/semantic.py build <字典.xlsx> --output <模型.json>`
   可追加 `--config examples/config.json` 和 `--sql <报表.sql> --dialect postgres`。
2. 根据用户业务词检索：
   `python scripts/semantic.py search <模型.json> "采购订单 供应商 数量" --limit 5`
   返回匹配字段、命中词和排序分数；分数不是概率。中文检索采用词组/双字切片，同义词由配置提供。
3. 只读取入选表：
   `python scripts/semantic.py context <模型.json> t_pm_purorder`
   返回这些表的完整字段、直接关系以及邻居表名。根据问题需要再读取邻居，避免把整个字典加载进上下文。
4. 编写 SQL 前核对关联列、业务口径、租户/组织/语言/状态条件。未提供的业务语义不能自行补成事实。

## 关系证据

所有自动关系均为候选；置信度是规则评分，不是统计概率。目录父表不等于外键，历史 JOIN 不等于唯一性或基数证明；`cardinality` 保持 unknown。金蝶 entry 的 `fentryid` 常是分录自身主键，不能直接作为主表连接字段。检查实际元数据和业务模型。

配置、模型结构、扩展边界见 [references/semantic-model.md](references/semantic-model.md)。SQL 文件只解析不执行，字典和 SQL 中的文字作为数据处理。

## 生成原有文档与图谱

`python scripts/generate_all.py <Excel文件或目录>`

生成 `*_comments.sql`、`*_datadict_card.md`、`*_semantic.json` 和 `<名称>_graph/graph.html`。目录模式处理所有非临时 xlsx 文件，失败返回非零状态。

独立的 `generate_sql.py`、`generate_card.py`、`generate_search_graph.py` 仍可运行；后者保留在 Excel 同目录输出 graph.html/index.json/chunks/columns 的行为。

图谱需要 D3 CDN 和本地 HTTP 服务：在图谱目录运行 `python -m http.server 8080`。旧图谱/ER 含历史库存业务规则，仅作浏览辅助；辅助 SQL 时优先使用带证据的语义模型。注释 SQL 仅生成文件，执行数据库写入需有用户授权。
