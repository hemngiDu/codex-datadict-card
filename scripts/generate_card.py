#!/usr/bin/env python3
"""Generate Markdown knowledge card from Kingdee data dictionary Excel"""
import sys, os, openpyxl
fp = sys.argv[1] if len(sys.argv) > 1 else None
if not fp or not os.path.exists(fp): print("Usage: ..."); sys.exit(1)
base = os.path.splitext(os.path.basename(fp))[0]
outdir = os.path.dirname(os.path.abspath(fp))
wb = openpyxl.load_workbook(fp)
ws = wb["表目录"]
rows = []
curr_no = None; curr_bo = None
for r in range(7, ws.max_row + 1):
    c4 = ws.cell(r,4).value; c5 = ws.cell(r,5).value
    c6 = ws.cell(r,6).value; c7 = ws.cell(r,7).value; c8 = ws.cell(r,8).value
    if c4:
        try: curr_no = int(str(c4).strip())
        except: pass
    if c5: curr_bo = str(c5).strip()
    if c6 and c7:
        rows.append({"no":curr_no,"bo":curr_bo,"desc":str(c6).strip(),"tbl":str(c7).strip(),"rel":str(c8).strip() if c8 else ""})
DG = [["配置与方案",["库存计划方案","安全库存统计方案","日均消耗统计方案","调度方案"]],["因子与水位",["库存水位信息","库存水位因子","因子数据维护","因子取数方案","因子取数规则","补货日期"]],["算法与计算",["算法方案配置","算法注册配置","库存水位更新方案","智能计算设置","因子分批计算参数","快速搭建模型","资源注册模型"]],["执行与结果",["库存计划建议","供需匹配明细","安全库存记录","匹配映射配置","日均消耗记录","客户服务水平"]],["日志审计",["消耗量计算日志","安全库存计算日志","库存计划运算日志"]],["辅助测试",["组织测试"]]]
CR = [["t_invp_scheme","t_invp_model_register","需求/供应来源模型"],["t_invp_scheme","t_invp_algoconfig","算法方案"],["t_invp_scheme","t_invp_calrulecfg","水位更新方案"],["t_invp_scheme","t_invp_invlevel","库存水位信息"],["t_invp_scheme","t_invp_matchconfig","匹配映射配置"],["t_invp_planadvice","t_invp_scheme","所属计划方案"],["t_invp_invlevel","t_invp_levelfactor","水位因子"],["t_invp_matchdetail","t_invp_scheme","计划方案"],["t_invp_smartcalccfg","t_invp_invlevel","库存水位"],["t_invp_smartcalccfg","t_invp_calrulecfg","因子计算规则"],["t_invp_smartcalccfg","t_invp_queryschema","因子取数方案"],["t_invp_invleveldata","t_invp_levelfactor","水位因子"],["t_invp_ssrecord","t_invp_ssdayscheme","安全库存方案"],["t_invp_dac_record","t_invp_dailycomsumption","日均消耗方案"],["t_invp_safestock_callog","t_invp_ssdayscheme","安全库存方案"]]
def gd(bo):
    for d, objs in DG:
        if bo in objs: return d
    return "其他"
bos = {}
for row in rows:
    bo = row["bo"]
    if bo not in bos: bos[bo] = []
    bos[bo].append(row)
lines = []
lines.append("# 数据字典速查卡")
lines.append("")
lines.append("共 **" + str(len(bos)) + "** 个业务对象 / **" + str(len(rows)) + "** 张物理表")
lines.append("")
lines.append("## 表结构模式")
lines.append("")
lines.append("| 层级 | 后缀 | 说明 | 关联方式 |")
lines.append("|---|---|---|---|")
lines.append("| 主表 | (无) | 业务对象主实体 | — |")
lines.append("| 多语言表 | _l | 多语言名称/描述 | fid→主表fid |")
lines.append("| 使用范围表 | _u | 组织/角色可见范围 | fdataid→主表fid |")
lines.append("| 子表/分录 | entry/_entry | 明细行、参数行 | fentryid→主表fid |")
lines.append("| 多选基础资料表 | (表名描述) | N:M 关联基础资料 | fpkid→主表fid |")
lines.append("")
lines.append("## 关联模型")
lines.append("")
lines.append("```mermaid")
lines.append("erDiagram")
ents = set()
for row in rows:
    if not row["rel"]:
        lines.append("    " + row["tbl"].upper() + " { }")
    ents.add(row["tbl"].upper())
    if row["rel"]: ents.add(row["rel"].upper())
for row in rows:
    if row["rel"] and row["rel"] != row["tbl"]:
        lbl = row["desc"][:12]
        lines.append("    " + row["rel"].upper() + " ||--o{ " + row["tbl"].upper() + " : \"" + lbl + "\"")
for src, tgt, lbl in CR:
    if src.upper() in ents and tgt.upper() in ents:
        lines.append("    " + src.upper() + " }o--|| " + tgt.upper() + " : \"" + lbl + "\"")
lines.append("```")
lines.append("")
lines.append("## 业务对象目录")
lines.append("")
for domain, bo_list in DG:
    has_any = any(bo in bos for bo in bo_list)
    if not has_any: continue
    lines.append("### " + domain)
    lines.append("")
    for bo_name in bo_list:
        if bo_name not in bos: continue
        tbls = bos[bo_name]
        main_tbl = [r for r in tbls if not r["rel"]]
        sub_tbls = [r for r in tbls if r["rel"]]
        if main_tbl:
            lines.append("**" + bo_name + "** \u2014 `" + main_tbl[0]["tbl"] + "` (" + main_tbl[0]["desc"] + ")")
        if sub_tbls:
            lines.append("")
            lines.append("| 数据表名 | 描述 | 关联父表 |")
            lines.append("|---|---|---|")
            for r in sub_tbls:
                lines.append("| `" + r["tbl"] + "` | " + r["desc"] + " | `" + r["rel"] + "` |")
        lines.append("")
lines.append("## 表名索引")
lines.append("")
lines.append("| 数据表名 | 业务对象 | 描述 | 关联父表 |")
lines.append("|---|---|---|---|")
for bo_name in sorted(bos.keys()):
    for r in bos[bo_name]:
        lines.append("| `" + r["tbl"] + "` | " + bo_name + " | " + r["desc"] + " | " + (r["rel"] or "—") + " |")
lines.append("")
outpath = os.path.join(outdir, base + "_datadict_card.md")
with open(outpath, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("Card: " + outpath + " (" + str(len(lines)) + " lines)")