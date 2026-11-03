#!/usr/bin/env python3
"""Find references from field remarks to entities in another workbook."""
import argparse, json, re
from pathlib import Path
import openpyxl

def main():
 p=argparse.ArgumentParser(); p.add_argument('model'); p.add_argument('target_excel'); p.add_argument('--output',required=True); a=p.parse_args()
 model=json.loads(Path(a.model).read_text(encoding='utf-8-sig')); wb=openpyxl.load_workbook(a.target_excel,read_only=True,data_only=True)
 entities={}
 for ws in wb:
  match=re.search(r'-([a-z][a-z0-9_]+)$',ws.title.lower())
  if not match: continue
  if ws.max_row==1 and ws.max_column==1: ws.reset_dimensions()
  entity=match.group(1); desc=''; mains=[]; current=None; keys=[]; in_cols=False
  for row in ws.iter_rows(max_col=8,values_only=True):
   marker=str(row[1] or '').strip()
   if marker=='表名：':
    desc=str(row[4] or '').strip(); current=None; in_cols=False
   elif marker=='数据表名：':
    current=str(row[4] or '').strip().lower(); in_cols=False
    if desc.endswith('-主表') and current: mains.append({'table':current,'primary_keys':[]})
   elif marker=='编号' and str(row[2] or '').strip()=='列名': in_cols=True
   elif in_cols and current and mains and mains[-1]['table']==current and str(row[7] or '').strip().lower() in ('是','√','true','1','yes','y','pk'):
    name=str(row[2] or '').strip().lower()
    if name: mains[-1]['primary_keys'].append(name)
  if mains: entities.setdefault(entity,[]).extend(mains)
 wb.close(); relations=[]; unresolved=[]
 for t in model['tables']:
  for c in t['columns']:
   remark=c.get('remarks',''); tokens=re.findall(r'(?<![a-z0-9_])([a-z][a-z0-9]*(?:_[a-z0-9]+)+)(?![a-z0-9_])',remark.lower())
   for entity in dict.fromkeys(tokens):
    targets=entities.get(entity,[])
    item={'source_table':t['name'],'source_column':c['name'],'entity':entity,'remark':remark,'target_workbook':Path(a.target_excel).name,'targets':targets}
    if '枚举' in remark:
     item['reason']='enumeration or polymorphic type; not a direct foreign-key value'
     unresolved.append(item)
    elif len(targets)==1 and len(targets[0]['primary_keys'])==1:
     item.update(target_table=targets[0]['table'],target_column=targets[0]['primary_keys'][0],status='candidate',evidence='field remark + exact worksheet entity identifier + main table primary key')
     relations.append(item)
    else:
     item['reason']='entity or unique main-table primary key not found' if not targets or any(len(x['primary_keys'])!=1 for x in targets) else 'ambiguous main tables'
     unresolved.append(item)
 Path(a.output).write_text(json.dumps({'source':model['source'],'target':Path(a.target_excel).name,'relationships':relations,'unresolved':unresolved},ensure_ascii=False,indent=2),encoding='utf-8')
 print('resolved',len(relations),'unresolved',len(unresolved))
if __name__=='__main__':main()
