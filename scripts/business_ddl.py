#!/usr/bin/env python3
"""Create PostgreSQL DDL for an exact business object from a semantic model."""
import argparse, json, re
from pathlib import Path

def ident(s):
    return '"' + s.replace('"', '""') + '"'
def literal(s):
    return "'" + str(s).replace("'", "''") + "'"
def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('model'); p.add_argument('business'); p.add_argument('--output', required=True)
    a=p.parse_args(); source=Path(a.model)
    if source.is_dir():
        chunks=[json.loads(f.read_text(encoding='utf-8-sig')) for f in sorted(source.glob('tables-*.json'))]
        if not chunks: p.error('No knowledge chunks found')
        m={'source':chunks[0]['source'], 'tables':[], 'relationships':[], 'metadata_conflicts':[]}
        seen=set()
        for chunk in chunks:
            m['tables'].extend(chunk['tables']); m['metadata_conflicts'].extend(chunk['metadata_conflicts'])
            for rel in chunk['relationships']:
                key=(rel['source_table'],rel['source_column'],rel['target_table'],rel['target_column'])
                if key not in seen: m['relationships'].append(rel); seen.add(key)
    else: m=json.loads(source.read_text(encoding='utf-8-sig'))
    tables=[t for t in m['tables'] if any(r['business_name']==a.business for r in t.get('references', []))]
    if not tables: p.error('No exact business object found')
    names={t['name'] for t in tables}
    problems=[]; statements=[]
    conflicts=[c for c in m.get('metadata_conflicts', []) if c['table'] in names]
    if conflicts: problems.append(f'{len(conflicts)} conflicting field definitions')
    for t in tables:
        if not t['columns']: problems.append(t['name']+': missing columns'); continue
        cols=[]; keys=[]
        for c in t['columns']:
            typ=c['type']; length=c.get('length',''); scale=c.get('scale','')
            if typ in ('varchar','bpchar','char','character varying'):
                if not length.isdigit() or int(length)<1: problems.append(t['name']+'.'+c['name']+': invalid length'); continue
                typ += '('+length+')'
            elif typ in ('numeric','decimal'):
                if not length.isdigit() or not scale.isdigit() or int(scale)>int(length) or int(length)<1: problems.append(t['name']+'.'+c['name']+': invalid precision/scale'); continue
                typ += '('+length+','+scale+')'
            elif typ not in ('int2','int4','int8','integer','bigint','smallint','text','timestamp','date','bool','boolean','float4','float8','bytea'):
                problems.append(t['name']+'.'+c['name']+': unsupported type '+typ); continue
            nullable=c.get('nullable_marker','')
            if nullable not in ('是','否'): problems.append(t['name']+'.'+c['name']+': unknown nullability'); continue
            default=c.get('default','')
            if default and not re.fullmatch(r"(?:[-+]?\d+(?:\.\d+)?|'(?:[^']|'')*'|NULL|true|false)",default,re.I):
                problems.append(t['name']+'.'+c['name']+': unsupported default'); continue
            col='    '+ident(c['name'])+' '+typ+(' NOT NULL' if nullable=='是' else '')+(' DEFAULT '+default if default else '')
            cols.append(col)
            if c['primary_key']: keys.append(ident(c['name']))
        if keys: cols.append('    PRIMARY KEY ('+', '.join(keys)+')')
        q=ident(t['name']); statements.append('CREATE TABLE '+q+' (\n'+',\n'.join(cols)+'\n);')
        desc=next((r['description'] for r in t['references'] if r['business_name']==a.business),t['description'])
        statements.append('COMMENT ON TABLE '+q+' IS '+literal(desc)+';')
        for c in t['columns']:
            statements.append('COMMENT ON COLUMN '+q+'.'+ident(c['name'])+' IS '+literal(c['description'] or c['remarks'])+';')
    if problems: p.error('; '.join(problems[:30]))
    lines=['-- Generated from '+m['source']+'; business object: '+a.business,
           '-- Candidate relationships below are not foreign keys and must be verified.']
    for r in m.get('relationships',[]):
        if r['source_table'] in names or r['target_table'] in names:
            lines.append('-- candidate: '+r['source_table']+'.'+r['source_column']+' = '+r['target_table']+'.'+r['target_column']+'; evidence: '+json.dumps(r['evidence'],ensure_ascii=False))
    Path(a.output).write_text('\n'.join(lines+['']+statements)+'\n',encoding='utf-8')
    print(f'{len(tables)} tables written to {a.output}')
if __name__=='__main__': main()
