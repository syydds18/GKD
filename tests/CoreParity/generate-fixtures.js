'use strict';
// Always execute the user's original JavaScript: expected results are never hand-written.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const source = process.argv[2] || path.join(__dirname, 'ExcelOutlineCore.original.js');
const core = require(path.resolve(source));
const assert = require('assert');
const sheets = [];
function sheet(name, rows, options = {}) {
  const values = {};
  rows.forEach((row, i) => { if (row !== null) { values[i + 1] = {}; row.forEach((v, c) => { if (v !== null && v !== undefined) values[i + 1][c + 1] = v; }); } });
  const input = Object.assign({values, formulas: {}, maxRow: rows.length, maxCol: Math.max(0, ...rows.filter(Boolean).map(r => r.length)), outlineLevels: {}, outlineMax: 0}, options);
  const detection = core.detectSheetValues(input.values, input.maxRow, input.maxCol);
  const outline = {max: input.outlineMax, rows: {}};
  Object.entries(input.outlineLevels).forEach(([r, level]) => outline.rows[r] = {level});
  const validation = core.validateSheetSums({...input, detection, outline});
  sheets.push({name, input, header: core.findHeader(input.values, input.maxRow, input.maxCol), rawDetection: core.detectLevels(input.values, input.maxRow, input.maxCol), detection, validation});
}
const H = ['编号', '项目名称', '金额', '数量', '单价'];
sheet('all-nine-families', [H, ['第1部分','主体',100], ['Ⅰ','工程',100], ['一','分项',100], ['（一）','细分',100], ['1','条目',100], ['（1）','子目',100], ['1）','更细',100], ['1.1','点号',100], ['①','末级',100], ['②','末级乙',50], ['第2部分','第二部分',20]]);
sheet('floating-codes-and-sibling-depth', [H, ['一','章',100], [2,'父',100], [2.2000000000000002,'浮点甲',50], [2.3000000000000003,'浮点乙',50], [1.1200000000000001,'浮点丙',3], ['13.011','不是编号',4], ['2.2.1','三级',4]]);
sheet('section-local-dotted-parents', [H, ['一','甲',30], ['（一）','甲子',30], ['3','编号3',30], ['3.1','甲点1',10], ['3.2','甲点2',20], ['二','乙',15], ['3.1','乙点1',5], ['3.2','乙点2',10], ['（二）','丙',15], ['9.1','缺失序号父',15]]);
sheet('no-code-low-confidence', [['指标','指标名称','金额'], [1,'指标甲',20], [2,'指标乙',30], [3,'指标丙',50]]);
sheet('headerless-text-code-fallback', [['报表'], ['说明'], ['其他'], ['一','甲',10], ['二','乙',20], ['三','丙',30], ['1','子甲',10], ['2','子乙',20]]);
sheet('header-code-only', [['标题'], ['序号','内容','金额'], ['一','父',30], ['1','甲',10], ['2','乙',20]]);
sheet('simple-pass', [H, ['一','父',30,90,5], ['1','甲',10,4,5], ['2','乙',20,6,5]]);
sheet('sum-error-formula-hint', [H, ['一','父',40], ['1','甲',10], ['2','乙',20]], {formulas:{2:{3:'SUM(C3:C4)+C3'}}});
sheet('subtotal-skip', [H, ['一','父',30], ['1','甲',10], ['2','乙',20], ['', '小计',30], ['二','合计',999], ['1','丙',100], ['2','丁',100]]);
sheet('fee-summary-candidate-minimum-difference', [H, ['一','父',35], ['1','甲',10], ['2','乙',20], ['', '增值税',5], ['', '分年度完成投资',999], ['二','父乙',30], ['1','丙',10], ['2','丁',20], ['', '额外明细',7]]);
sheet('fee-candidate-error-basis', [H, ['一','父',38], ['1','甲',10], ['2','乙',20], ['', '增值税',5], ['', '年度完成',999]]);
sheet('all-uncoded-children', [H, ['一','父',30], ['', '明细甲',10], ['', '明细乙',20]]);
sheet('exact-0.05-boundary', [H, ['一','父',0.05], ['1','甲',0], ['2','乙',0]]);
sheet('over-0.05-boundary', [H, ['一','父',0.05000001], ['1','甲',0], ['2','乙',0]]);
sheet('relative-tolerance', [H, ['一','父',1000000001], ['1','甲',500000000], ['2','乙',500000000]]);
sheet('excel-cell-errors', [H, ['一','父',30], ['1','甲','#REF!'], ['2','乙','#DIV/0!'], ['3','丙','#NAME?'], ['4','丁','#N/A'], ['5','戊','#VALUE!'], ['6','己','#UNKNOWN!']], {formulas:{3:{3:'BAD()'},4:{3:''}}});
sheet('outline-fallback', [H, ['一','父',30], ['1','甲',10], ['2','乙',20]], {outlineMax:1, outlineLevels:{2:1,3:1,4:1}});
sheet('outline-primary', [H, ['一','父',30], ['1','甲',10], ['2','乙',20]], {outlineMax:2, outlineLevels:{2:1,3:2,4:2}});
sheet('outline-cannot-bypass-low-confidence', [['指标','项目名称','金额'], [1,'甲',30], [2,'乙',10], [3,'丙',20]], {outlineMax:2, outlineLevels:{2:1,3:2,4:2}});
sheet('no-amount-columns', [['编号','名称','数量','单价','费率'], ['一','父',30,4,0.1], ['1','甲',10,2,0.2], ['2','乙',20,2,0.3]]);
sheet('multi-row-header', [['编号','项目名称','设备费','安装费','第1年'], ['', '', '含税金额','不含税金额','投资'], ['一','父',30,40,50], ['1','甲',10,10,20], ['2','乙',20,30,30]]);
sheet('banner-and-trailing-empty', [H, ['', '第1部分',30], ['一','甲',10], ['二','乙',20], [], null, []]);
sheet('single-child-not-checked', [H, ['一','父',30], ['1','甲',10]]);
sheet('number-parsing', [H, ['一','父','￥1,234.50'], ['1','甲','1，000'], ['2','乙','234.50']]);
sheet('negative-number-parsing', [H, ['一','父','(30)'], ['1','甲','(10)'], ['2','乙','-2e1']]);
sheet('empty', []);
sheet('empty-rows-and-name-fallback', [H, ['', '', '单位'], [], ['一','父',30], ['1','',10,'后备名'], [], ['2','乙',20], []]);
const many = [H];
for(let i=0;i<510;i++){many.push(['一','父'+i,40],['1','甲',10],['2','乙',20]);}
sheet('error-truncation-500',many);
// Deterministic fuzz: mixed numbering, names, amount columns, source outline and formulas.
let seed=0x5eeda11;
function rand(n){seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed%n;}
const codes=['一','二','（一）','（二）','1','2','3','（1）','1）','1.1','1.2.1','①','Ⅰ','第1部分','',null,2.2000000000000002,'13.011'];
const names=['主体','明细','合计','基本预备费','年度完成投资','增值税','总投资','第2部分',''];
const amounts=[0,0.1,0.2,0.3,0.05,-0.1,10,20,30,40,50,'1,000','(20)','#REF!','#NAME?','',null,'￥30',1.05];
for(let i=0;i<160;i++){
 const rows=[H], outlineLevels={}, formulas={};
 for(let r=2;r<12+rand(25);r++){rows.push([codes[rand(codes.length)],names[rand(names.length)],amounts[rand(amounts.length)],rand(100),rand(10)]); if(i%3===0)outlineLevels[r]=rand(5);if(rand(4)===0)formulas[r]={3:'SUM(C1:C2)'};}
 sheet('seeded-mixed-'+i,rows,{outlineLevels,outlineMax:i%3===0?4:0,formulas});
}
const codeInputs=[1e-7,1e-6,1e20,1e21,-0,1.2345678901234567,1000000000000000100,'第1部分','Ⅰ','（一）','一、','1','（1）','1）','3.1.2','①',2.2000000000000002,1.1200000000000001,'13.011','0.397','3．1．',' 第十二章说明 ','1.2.3.4.5','1.2.3.4.5.6','',null,true,'\u00a0（一）\u3000','Ⅼ','99.99000000000001'];
const numberInputs=[0,1,-1,0.05,'1,234.50','￥1，000','(20)','(-20)',' 1 2 ','1e3','.5','1.','+2','NaN','Infinity','#REF!','#NAME?','#N/A','#DIV/0!','#SPILL!','#GETTING_DATA','10%','￥¥20','',null,true];
const fixture={sourceSha256:crypto.createHash('sha256').update(fs.readFileSync(source)).digest('hex'),codeCases:codeInputs.map(input=>({input,expected:core.codeFamily(input)})),numberCases:numberInputs.map(input=>({input,expected:core.toNumberValue(input)})),sheets};
assert(sheets.find(x=>x.name==='outline-fallback').validation.fallback===true);
assert(sheets.find(x=>x.name==='exact-0.05-boundary').validation.status==='pass');
assert(sheets.find(x=>x.name==='over-0.05-boundary').validation.status==='diff');
assert(sheets.find(x=>x.name==='error-truncation-500').validation.hiddenErrors===10);
const out=path.join(__dirname,'fixtures.json');
fs.writeFileSync(out,JSON.stringify(fixture,null,2)+'\n');
console.log(`Reference generated by original JS: ${sheets.length} sheets, ${fixture.codeCases.length} code cases, ${fixture.numberCases.length} number cases; SHA-256 ${fixture.sourceSha256}`);
if(process.argv.includes('--require-core')){
 assert(fs.existsSync(path.join(__dirname,'../../src/Outline.Core/Engine.cs')), 'RED: C# Engine implementation does not exist yet');
}
