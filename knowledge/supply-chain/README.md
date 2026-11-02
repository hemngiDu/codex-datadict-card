# 供应链云数据字典知识库

来源：金蝶数据字典导出 Excel。先搜索 index.json 的表名或业务对象，再读取对应 tables-NNN.json。每个文件包含表字段、业务对象引用、相关候选关系及字段定义差异。字段差异保留原始工作表名称，需结合原始导出核对。

关系为规则推断候选，不是已验证的外键或基数。confidence 是规则分数而非概率。生成 SQL 前须核对连接字段、组织、租户、语言和业务状态；不能确定时明确说明。表目录父子归属不等于外键。没有 SQL 历史或数据库约束作为证据。

相同物理表可属于多个业务对象：references 保存各目录引用；顶层业务名称与父表是首条记录，不代表唯一归属。字段首次出现的定义在 columns，冲突的其他定义在 metadata_conflicts；字段缺失不应猜测。

## WorkBuddy 查询与 PostgreSQL 建表

用户输入业务对象名时，先在 index.json 精确匹配 references.business_name，列出该对象全部物理表；相似对象（如采购计划变更单）分别列为候选，不自动混入。读取这些表所在分块，返回每表全部字段的类型、长度、精度、主键、非空标记、默认值、中文注释以及关联候选与 evidence。metadata_conflicts 非空时展示差异，不能自动挑选。列缺失或类型不明时明确说明，不能编造。

如需 PostgreSQL 建表，运行 `python scripts/business_ddl.py knowledge/supply-chain 采购计划 --output 采购计划.sql`。脚本遇字段冲突、缺失或不支持的类型和默认值会报错停止；仅将候选关系写为注释，不创建外键。用户需核对冲突、数据库版本、schema、索引和实际业务关联；生成的 DDL 不包含字典未提供的索引或权限。执行 SQL 需由用户决定。

## 跨字典关系

查询供应链表时同时读取 knowledge/supply-chain/base-service-relations.json，按 source_table 匹配。relationships 由字段备注实体标识、基础服务云工作表实体标识、主表和单列主键推断，均为待验证候选；unresolved 保存无法唯一解析的实体，禁止猜测目标表或 JOIN。报告原始备注和证据，具体主键连接仍须业务核实。
