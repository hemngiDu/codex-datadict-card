# 语义模型与扩展

## 数据契约（schema_version = 1）

build 读取目录第 7 行起 D:H（编号、业务对象、说明、物理表、父表），明细识别“表名：”“数据表名：”“编号/列名”等标记。字段保留名称、说明、类型、主键标记和备注；不截断。重复表或字段显式报错，缺少目录报错；目录中无字段详情的表可被发现，但不会猜出缺失字段。

模型包含 `tables`、`relationships`、`config`、`warnings`。tables 包含 name/business_name/description/parent/columns；columns 包含 name/description/type/primary_key/remarks。标识符统一小写，适用于常见金蝶导出；大小写敏感数据库标识符需要另行适配。

关系包含 source_table/source_column/target_table/target_column/confidence/evidence/status/cardinality。source/target 是记录端点顺序，不能解释为经过验证的外键方向。候选必须两端表字段存在，且已知类型兼容。无基数证明时 cardinality=unknown。

## 配置

参考 [examples/config.json](../examples/config.json)：

- `kingdee_rules`：是否启用目录父表 + `_u.fdataid` 或其他表 `.fid` 的保守候选。
- `min_confidence`：0 到 1 的输出阈值，默认 0.5。
- `synonyms`：查询别名到扩展词数组的映射；不调用外部模型。
- `relationships`：人工配置的关系数组，字段为四个端点名称及可选 evidence 字符串；端点不存在会报错。

评分：人工配置 0.95，目录与金蝶规则 0.80，SQL JOIN 0.75，命名与目标单列主键 0.65。相同有向端点合并 evidence、保留最大评分，不累加分数。反向 SQL JOIN 可单独保留，因为并未证明方向等价。多目标命名匹配保留歧义供人工判断。未接入数据库 FK 或采样，不能声称已验证外键。

## SQL 历史适配器

`sql_history.extract` 返回 observations 和 warnings，再由 infer 检查元数据端点及类型。只接受 flat SELECT 中直接列到列的 JOIN ON 等式及 AND 组合；OR、函数内等式不用于关系推断。多语句分别处理；CTE/子查询/UNION/DML 明确跳过。USING、无别名的未限定列不推断。非法 SQL 使构建失败；仅保留来源文件名与语句编号，不把原 SQL 或字面值复制到模型。

下一步可添加作用域/列血缘解析，或独立 FK/采样适配器；每个适配器输出端点与来源证据，经统一关系层校验，避免在检索或展示层重新猜关系。增加真实查询用例和标注关系后，才考虑校准评分或推断基数。

## 检索与上下文

search 返回排序分数、匹配的表属性和字段，不代表语义向量检索。中文词组 + 双字切片提高无分词器环境的召回，但也可能误匹配。context 返回指定表全部字段、直接关系及候选邻居名称；不会递归扩展，以控制上下文体积。业务域、指标口径、常用过滤条件尚不自动生成，防止把词法相似当作业务事实。
