#!/usr/bin/env python3
"""Independent outline audit; writes reports and a ZIP-preserving copy only."""
from pathlib import Path
from collections import Counter
from decimal import Decimal
from copy import copy
import json, re, hashlib, zipfile, xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

SOURCE=Path('/Users/tmsxhc/.hermes/cache/documents/native-inputs/107e0d0e979a37295b09392744736dc8e2ba38d58cb2c609831d612f865d6317/doc_c1a3b6a16de6_工作簿1.xlsx')
OUT=Path('/Users/tmsxhc/Desktop/Excel分级校验加载项/audit')
OUT.mkdir(parents=True,exist_ok=True)
SHEET='F主体-分标'
THRESHOLD=Decimal('0.05')
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
W=openpyxl.load_workbook(SOURCE, data_only=False)
D=openpyxl.load_workbook(SOURCE, data_only=True)
s=W[SHEET]; cache=D[SHEET]
source_hash=hashlib.sha256(SOURCE.read_bytes()).hexdigest()
def dec(v):
    return Decimal(str(v)) if isinstance(v,(int,float)) and not isinstance(v,bool) else None

def amount(r,col):
    v=s.cell(r,col).value
    if v is None: return None,'空白'
    n=dec(v)
    if n is not None: return n,'原始数值'
    if s.cell(r,col).data_type=='f':
        cv=dec(cache.cell(r,col).value)
        if cv is not None: return cv,'原文件公式缓存（非本次重算）'
        # Deliberately limited grammar: same-row F*E/10000, optional unary +/$.
        patt=rf'^=\+?\$?F\$?{r}\*\$?E\$?{r}/10000$'
        if col==7 and re.fullmatch(patt,v,re.I):
            f,e=dec(s.cell(r,6).value),dec(s.cell(r,5).value)
            if f is not None and e is not None: return f*e/Decimal(10000),'限定语法解释 F*E/10000（非Excel重算）'
        return None,'无法解释的无缓存公式'
    return None,'非数值'

# Every physical row is retained, including title/header/blank rows.
rows=[]; nodes={}
for r in range(1,s.max_row+1):
    meaningful=any(s.cell(r,c).value is not None for c in range(2,11))
    status={'row':r,'level':s.row_dimensions[r].outlineLevel,'name':s.cell(r,3).value,
            'classification':'数据' if r>=4 and meaningful else ('标题表头' if r<=3 else '空白'),
            'G_original':s.cell(r,7).value,'J_original':s.cell(r,10).value}
    rows.append(status)
    if status['classification']=='数据': nodes[r]=status

# Nearest-lower stack gives spans, but only +1 level edges are legal direct edges.
stack=[]; jumps=[]
for row in rows[3:]:
    r=row['row']; level=row['level']
    if row['classification']!='数据':
        while stack: nodes[stack.pop()]['span_end']=r-1
        continue
    while stack and nodes[stack[-1]]['level']>=level:
        nodes[stack.pop()]['span_end']=r-1
    candidate=stack[-1] if stack else None
    row.update(candidate_parent=candidate,strict_parent=None,children=[],candidate_children=[])
    if candidate:
        nodes[candidate]['candidate_children'].append(r)
        if level==nodes[candidate]['level']+1:
            row['strict_parent']=candidate
            nodes[candidate]['children'].append(r)
        else:
            jumps.append({'row':r,'candidate_parent':candidate,'parent_level':nodes[candidate]['level'],
                          'level':level,'missing_levels':list(range(nodes[candidate]['level']+1,level)),
                          'type':'向上跳级起点' if r==candidate+1 else '跳级区同级孤立分支'})
    elif level!=0:
        jumps.append({'row':r,'candidate_parent':None,'parent_level':None,'level':level,'missing_levels':list(range(level)), 'type':'非0级孤立根'})
    stack.append(r)
while stack: nodes[stack.pop()]['span_end']=s.max_row

jump_regions=[]; jump_region_rows=set(); blocked_ancestors=set()
for jump in jumps:
    p=jump['candidate_parent'] or jump['row']
    end=nodes[p]['span_end']
    reg=(p,end)
    if reg not in jump_regions: jump_regions.append(reg)
    jump_region_rows.update(r for r in nodes if p<=r<=end)
    parent=nodes[p]['candidate_parent']
    while parent:
        blocked_ancestors.add(parent)
        parent=nodes[parent]['candidate_parent']

for r,row in nodes.items():
    g,gs=amount(r,7); j,js=amount(r,10)
    row.update(G_value=g,G_source=gs,J_value=j,J_source=js,
               is_parent=bool(row['candidate_children']),
               jump_region=r in jump_region_rows,blocked_ancestor=r in blocked_ancestors)

for r in reversed(list(nodes)):
    row=nodes[r]; children=row['children']; candidate=row['candidate_children']
    def aggregate(child_rows,key):
        vals=[nodes[c][key] for c in child_rows]
        return sum(vals,Decimal(0)) if all(v is not None for v in vals) else None
    if not candidate:
        row['leaf_sum_G']=row['G_value']
        row['strict_recursive_leaf_sum_G']=row['G_value']
    else:
        row['leaf_sum_G']=aggregate(candidate,'leaf_sum_G')
        row['strict_recursive_leaf_sum_G']=aggregate(children,'strict_recursive_leaf_sum_G') if len(children)==len(candidate) else None
    child_j=[nodes[c]['J_value'] for c in children]
    row['direct_children_J_complete']=bool(children) and all(v is not None for v in child_j)
    row['direct_children_J_sum']=sum(child_j,Decimal(0)) if row['direct_children_J_complete'] else None
    row['direct_children_J_known_partial_sum']=sum((v for v in child_j if v is not None),Decimal(0)) if children else None
    for key, val in [('recursive',row['strict_recursive_leaf_sum_G']),('diagnostic_span',row['leaf_sum_G']),('direct_J',row['direct_children_J_sum'])]:
        diff=val-row['J_value'] if val is not None and row['J_value'] is not None else None
        row[key+'_minus_original_J']=diff
        row[key+'_J_difference_ge_005']=diff is not None and abs(diff)>=THRESHOLD
    row['single_child']=row['is_parent'] and len(children)==1 and len(candidate)==1
    row['blank_G_parent']=row['is_parent'] and row['G_original'] is None
    row['old_plugin_single_child_skip']=row['blank_G_parent'] and row['single_child']
    reasons=[]
    if row['jump_region']: reasons.append('跳级区域人工复核')
    if row['blocked_ancestor']: reasons.append('下级跳级导致依赖不完整')
    if row['is_parent'] and row['strict_recursive_leaf_sum_G'] is None and not reasons: reasons.append('叶子G缺失或无法解释')
    if row['is_parent'] and any(nodes[c]['G_original'] is None and not nodes[c].get('safe_fill',False) for c in children): reasons.append('直接子项G空白且无法安全补全')
    row['safe_fill']=row['blank_G_parent'] and bool(children) and not reasons
    row['block_reasons']=reasons
    row['proposed_G_formula']='=SUM('+','.join('G'+str(c) for c in children)+')' if row['safe_fill'] else None
    issues=[]
    if row['blank_G_parent']: issues.append('父项空G（可能为测试清空）')
    if row['old_plugin_single_child_skip']: issues.append('旧至少2项策略会跳过单子项')
    if row['jump_region']: issues.append('层级跳跃人工复核')
    if row['blocked_ancestor']: issues.append('跳级影响上级')
    if row['is_parent'] and children and not row['direct_children_J_complete']: issues.append('直接子项J缺失，完整和不可比较')
    if row['J_original'] is None: issues.append('J空白-'+('父项' if row['is_parent'] else '叶子'))
    if row['recursive_J_difference_ge_005']: issues.append('递归叶子和与J差异>=0.05')
    if row['direct_J_J_difference_ge_005']: issues.append('直接子项J和与J差异>=0.05')
    if row['diagnostic_span_J_difference_ge_005'] and row['strict_recursive_leaf_sum_G'] is None: issues.append('含跳级诊断叶子和与J差异>=0.05')
    row['issues']=issues

parents=[n for n in nodes.values() if n['is_parent']]
leaves=[n for n in nodes.values() if not n['is_parent']]
filled={n['row']:n['proposed_G_formula'] for n in parents if n['safe_fill']}
# Patch exact G cell XML while retaining all other archive members byte-for-byte.
COPY=OUT/'工作簿1_安全补全副本.xlsx'
with zipfile.ZipFile(SOURCE) as zin:
    original_sheet=zin.read('xl/worksheets/sheet1.xml').decode('utf-8')
    text=original_sheet
    for r,formula in filled.items():
        patt=rf'<c\b[^>]*\br="G{r}"[^>]*?(?:/>|>.*?</c>)'
        matches=list(re.finditer(patt,text,re.S))
        assert len(matches)==1,(r,len(matches))
        old=matches[0].group(0)
        el=ET.fromstring(old)
        assert el.find('f') is None and el.find('v') is None and el.find('is') is None,(r,old)
        opening=re.match(r'<c\b[^>]*',old).group(0).rstrip('/')
        opening=re.sub(r'\s+t="[^"]*"','',opening)
        new=opening+'><f>'+formula[1:]+'</f></c>'
        text=text[:matches[0].start()]+new+text[matches[0].end():]
    workbook_xml=zin.read('xl/workbook.xml').decode('utf-8')
    def calc_replace(m):
        opening=m.group(0).removesuffix('/>')
        opening=re.sub(r'\s+(fullCalcOnLoad|forceFullCalc|calcOnSave)="[^"]*"','',opening)
        return opening+' fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="1"/>'
    workbook_new,ncalc=re.subn(r'<calcPr\b[^>]*/>',calc_replace,workbook_xml)
    assert ncalc==1
    with zipfile.ZipFile(COPY,'w') as zout:
        for info in zin.infolist():
            value=zin.read(info.filename)
            if info.filename=='xl/worksheets/sheet1.xml': value=text.encode('utf-8')
            elif info.filename=='xl/workbook.xml': value=workbook_new.encode('utf-8')
            zout.writestr(info,value)

# Verify persisted artifacts: exact original cells/styles/outlines and untouched ZIP parts.
V=openpyxl.load_workbook(COPY,data_only=False)
v=V[SHEET]
non_g_diff=[]; g_diff=[]; style_diff=[]; dim_diff=[]
for r in range(1,s.max_row+1):
    for c in range(1,s.max_column+1):
        a,b=s.cell(r,c),v.cell(r,c)
        if a.value!=b.value: (g_diff if c==7 else non_g_diff).append(a.coordinate)
        if a._style!=b._style: style_diff.append(a.coordinate)
    if dict(s.row_dimensions[r])!=dict(v.row_dimensions[r]): dim_diff.append(r)
assert not non_g_diff and not style_diff and not dim_diff
assert set(g_diff)=={'G'+str(r) for r in filled}
assert all(v.cell(r,7).value==formula and s.cell(r,7).value is None for r,formula in filled.items())
assert set(map(str,s.merged_cells.ranges))==set(map(str,v.merged_cells.ranges))
assert s.sheet_properties.outlinePr==v.sheet_properties.outlinePr
with zipfile.ZipFile(SOURCE) as za,zipfile.ZipFile(COPY) as zb:
    changed_parts=[f for f in za.namelist() if za.read(f)!=zb.read(f)]
    assert za.namelist()==zb.namelist()
    assert set(changed_parts)=={'xl/worksheets/sheet1.xml','xl/workbook.xml'}
    # Restore every modified G element in the copy and compare the full worksheet XML.
    restored=zb.read('xl/worksheets/sheet1.xml').decode('utf-8')
    for r in filled:
        patt=rf'<c\b[^>]*\br="G{r}"[^>]*?(?:/>|>.*?</c>)'
        orig=re.search(patt,original_sheet,re.S).group(0)
        restored=re.sub(patt,lambda m:orig,restored,count=1,flags=re.S)
    assert restored==original_sheet
assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==source_hash

stats={
 'physical_rows':s.max_row,'physical_columns':s.max_column,
 'physical_rows_audited':len(rows),'header_rows':3,'blank_rows':sum(x['classification']=='空白' for x in rows),
 'data_rows':len(nodes),'parent_rows':len(parents),'leaf_rows':len(leaves),
 'blank_G_parents':sum(n['blank_G_parent'] for n in parents),
 'blank_G_leaves':sum(n['G_original'] is None for n in leaves),
 'single_child_parents':sum(n['single_child'] for n in parents),
 'old_plugin_single_child_skips':sum(n['old_plugin_single_child_skip'] for n in parents),
 'upward_jump_transitions':sum(j['type']=='向上跳级起点' for j in jumps),
 'strict_orphan_rows':len(jumps),'jump_regions':len(jump_regions),'jump_region_rows':len(jump_region_rows),
 'jump_affected_ancestors':len(blocked_ancestors),
 'safe_G_formulas_written':len(filled),'safe_single_child_formulas':sum(nodes[r]['single_child'] for r in filled),
 'blank_G_parents_not_written':sum(n['blank_G_parent'] and not n['safe_fill'] for n in parents),
 'J_blank_leaves':sum(n['J_original'] is None for n in leaves),'J_blank_parents':sum(n['J_original'] is None for n in parents),
 'strict_recursive_J_difference_rows':sum(n['recursive_J_difference_ge_005'] for n in nodes.values()),
 'direct_children_J_difference_rows':sum(n['direct_J_J_difference_ge_005'] for n in parents),
 'diagnostic_span_J_difference_rows':sum(n['diagnostic_span_J_difference_ge_005'] for n in nodes.values()),
 'G_formula_sources':dict(Counter(n['G_source'] for n in nodes.values() if s.cell(n['row'],7).data_type=='f')),
 'level_counts':dict(Counter(n['level'] for n in nodes.values()))}
verification={
 'source_sha256_before_and_after':source_hash,'source_unchanged':True,'G_changed_cells':g_diff,
 'non_G_value_formula_differences':non_g_diff,'style_differences':style_diff,'row_dimension_differences':dim_diff,
 'merged_range_count_before':len(s.merged_cells.ranges),'merged_range_count_after':len(v.merged_cells.ranges),
 'outline_summaryBelow_before':s.sheet_properties.outlinePr.summaryBelow,'outline_summaryBelow_after':v.sheet_properties.outlinePr.summaryBelow,
 'archive_member_count':len(zipfile.ZipFile(SOURCE).namelist()),'changed_archive_parts':changed_parts,
 'all_other_archive_parts_byte_identical':True,'worksheet_outside_filled_G_byte_identical':True,
 'formula_results_recalculated_by_Excel':False,'new_formulas_have_no_cached_results':True,
 'existing_formula_caches_preserved':True,'fullCalcOnLoad':V.calculation.fullCalcOnLoad,
 'forceFullCalc':V.calculation.forceFullCalc,'calcOnSave':V.calculation.calcOnSave}
notes=[
 '所有1840个物理行逐行落盘；数据行按B:J存在非空值识别，标题表头及空白单列。',
 'summaryBelow=False，父在上；严格直接子项须位于连续后续较深区间且level=父level+1。仅对子项求和，不同时计入孙节点。',
 'nearest-lower候选关系只用于界定区间和跳级诊断，不把跳级节点认定为合法直接子项；严格递归结果遇缺级为未知。',
 'G父项空白可能是用户有意清空的测试数据，只统计待求和位置，不一概断言用户漏填。',
 '旧插件至少2项规则的漏项为结构推定，未运行或修改插件源码。',
 '递归叶子和使用G原始数值/原缓存；缓存缺失仅允许同一行=+F*E/10000限定语法解释，并注明来源。',
 '直接子项J和遇任一J缺失为不可完整比较，已知部分单列；J空白本身作为问题，不默认为0。',
 '差额=计算值-原J，绝对值>=0.05万元列为差异；含跳级区间叶子和仅供人工诊断。',
 '跳级区域和受影响上级不写G；安全副本只在原G为空的父项写直接子项SUM，单子项写=SUM(G行号)。',
 '副本定点修改ZIP中工作表G单元格和workbook calcPr，其余归档成员字节保持一致；J、批注、外链、样式、打印设置等原样保留。',
 'openpyxl仅用于读取/报告生成/验证，不是Excel计算引擎；新公式无计算缓存，需Excel打开重算，报告中的Python复核值非Excel重算结果。']

def encode(obj):
    if isinstance(obj,Decimal): return float(obj)
    raise TypeError(type(obj).__name__)
payload={'source':str(SOURCE),'sheet':SHEET,'threshold':str(THRESHOLD),'stats':stats,
         'notes':notes,'jump_edges':jumps,'jump_regions':jump_regions,'blocked_ancestor_rows':sorted(blocked_ancestors),
         'verification':verification,'rows':rows,'files':{'copy':str(COPY),'report':str(OUT/'全表层级核查报告.xlsx')}}
JSON=OUT/'全表层级核查统计.json'
JSON.write_text(json.dumps(payload,ensure_ascii=False,indent=2,default=encode),encoding='utf-8')

R=openpyxl.Workbook(); overview=R.active; overview.title='说明与统计'
overview.append(['项目','值'])
for k,value in stats.items(): overview.append([k,json.dumps(value,ensure_ascii=False) if isinstance(value,dict) else value])
for note in notes: overview.append(['核查口径/限制',note])
for k,value in verification.items():
    if k=='G_changed_cells': continue
    overview.append(['验证:'+k,json.dumps(value,ensure_ascii=False) if isinstance(value,(dict,list)) else value])
HEADERS=[('row','原行号'),('level','大纲级别'),('name','名称'),('classification','行分类'),('is_parent','父项'),
 ('strict_parent','严格父行'),('candidate_parent','候选父行_诊断'),('span_end','区间末行'),
 ('children','严格直接子行'),('candidate_children','候选子行_含跳级诊断'),('G_original','原G值或公式'),('G_value','G复核数值'),('G_source','G取值来源'),
 ('J_original','原J值或公式'),('J_value','J数值'),('strict_recursive_leaf_sum_G','严格递归叶子G和'),('recursive_minus_original_J','严格叶子和减J'),
 ('direct_children_J_sum','直接子项J和_完整'),('direct_children_J_known_partial_sum','直接子项J已知部分和'),('direct_J_minus_original_J','直接子J和减原J'),
 ('leaf_sum_G','区间叶子G和_跳级仅诊断'),('diagnostic_span_minus_original_J','区间叶子和减J_诊断'),
 ('single_child','严格单子项'),('blank_G_parent','父项空G'),('jump_region','跳级区'),('blocked_ancestor','跳级影响上级'),
 ('safe_fill','副本已安全补全'),('proposed_G_formula','副本新G公式'),('block_reasons','不补全原因'),('issues','状态')]
def safe_cell(v):
    if isinstance(v,Decimal): return float(v)
    if isinstance(v,(list,dict)): return json.dumps(v,ensure_ascii=False)
    return v

def add_sheet(title,selected):
    ws=R.create_sheet(title)
    ws.append([label for key,label in HEADERS])
    for row in selected:
        ws.append([safe_cell(row.get(key)) for key,label in HEADERS])
        rr=ws.max_row
        ws.cell(rr,1).hyperlink=f"{COPY.name}#'{SHEET}'!G{row['row']}"
        ws.cell(rr,1).style='Hyperlink'
        # Formula text is explicitly text in audit reports, never executed.
        for cc in (11,14,28):
            cell=ws.cell(rr,cc)
            if isinstance(cell.value,str) and cell.value.startswith('='): cell.data_type='s'
    return ws
add_sheet('全表逐行',rows)
add_sheet('父项空G',[n for n in parents if n['blank_G_parent']])
add_sheet('旧策略单子项漏项',[n for n in parents if n['old_plugin_single_child_skip']])
add_sheet('层级跳跃人工复核',[n for n in nodes.values() if n['jump_region'] or n['blocked_ancestor']])
add_sheet('J对照差额',[n for n in nodes.values() if n['recursive_J_difference_ge_005'] or n['direct_J_J_difference_ge_005'] or n['diagnostic_span_J_difference_ge_005']])
add_sheet('J空白叶子与父项',[n for n in nodes.values() if n['J_original'] is None])
add_sheet('副本G补全清单',[nodes[r] for r in filled])
for ws in R:
    ws.freeze_panes='D2' if ws is not overview else 'B2'
    ws.auto_filter.ref=ws.dimensions
    ws.row_dimensions[1].height=32
    for cell in ws[1]:
        cell.fill=PatternFill('solid',fgColor='17365D');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(wrap_text=True)
    for col in ws.columns:
        letter=col[0].column_letter
        ws.column_dimensions[letter].width=15
    if ws is overview:
        ws.column_dimensions['A'].width=48;ws.column_dimensions['B'].width=110
    else:
        for letter,width in {'C':38,'I':28,'J':28,'K':30,'M':36,'N':30,'AB':48,'AC':52,'AD':65}.items(): ws.column_dimensions[letter].width=width
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.alignment=Alignment(vertical='top',wrap_text=True)
REPORT=OUT/'全表层级核查报告.xlsx'; R.save(REPORT)
# Cross-check report enumeration against JSON and required category counts.
RR=openpyxl.load_workbook(REPORT,read_only=True,data_only=False)
expected={'全表逐行':s.max_row,'父项空G':stats['blank_G_parents'],'旧策略单子项漏项':stats['old_plugin_single_child_skips'],
 '层级跳跃人工复核':stats['jump_region_rows']+stats['jump_affected_ancestors'],
 'J空白叶子与父项':stats['J_blank_leaves']+stats['J_blank_parents'],'副本G补全清单':len(filled)}
assert all(RR[title].max_row-1==count for title,count in expected.items())
assert len(json.loads(JSON.read_text(encoding='utf-8'))['rows'])==s.max_row
print(json.dumps({'stats':stats,'jump_edges':jumps,'jump_regions':jump_regions,'blocked_ancestors':sorted(blocked_ancestors),
 'single_rows':[n['row'] for n in parents if n['single_child']],
 'differences':[{'row':n['row'],'name':n['name'],'J':n['J_value'],'strict_leaf':n['strict_recursive_leaf_sum_G'],
  'directJ':n['direct_children_J_sum'],'diagnostic_leaf':n['leaf_sum_G'],'delta':n['diagnostic_span_minus_original_J']}
  for n in nodes.values() if n['recursive_J_difference_ge_005'] or n['direct_J_J_difference_ge_005'] or n['diagnostic_span_J_difference_ge_005']],
 'verification':{k:v for k,v in verification.items() if k!='G_changed_cells'},'files':[str(JSON),str(REPORT),str(COPY)]},ensure_ascii=False,indent=2,default=encode))
