import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
const runtimeModules=process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES || process.env.ARTIFACT_TOOL_NODE_MODULES;
const resolved=runtimeModules ? createRequire(path.join(runtimeModules,'anchor.cjs')).resolve('@oai/artifact-tool') : '@oai/artifact-tool';
const { Workbook, SpreadsheetFile } = await import(runtimeModules ? pathToFileURL(resolved).href : resolved);
const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('Usage: node export-excel.mjs input.json output.xlsx');
const a = JSON.parse((await fs.readFile(input,'utf8')).replace(/^\uFEFF/, ''));
for (const k of ['shares','buy_price','bond_quantity','bond_buy_price']) {
  if (!Number.isFinite(a[k])) throw new Error('Missing/invalid '+k);
}
if(a.shares<=0 || !Number.isInteger(a.shares) || a.buy_price<=0 || a.bond_buy_price<=0
 || a.bond_quantity<0 || !Number.isInteger(a.bond_quantity)) throw new Error('Invalid quantity or price');
for (const k of ['buy_fees','total_fees','allocation_per_share']) {
  if(a[k]!=null && (!Number.isFinite(a[k]) || a[k]<0)) throw new Error('Invalid '+k);
}
if(a.total_fees!=null && (a.buy_fees??0)>a.total_fees) throw new Error('Total fees must include buy fees');
const prices = a.bond_prices??[95,100,110,120,145];
if(!Array.isArray(prices)||!prices.length||prices.some(p=>!Number.isFinite(p)||p<0)) throw new Error('Invalid scenarios');
if((a.tick??.01)<=0 || (a.offset??.1)<0) throw new Error('Invalid tick or offset');
const wb=Workbook.create();
const names=['持仓与参数','配债测算','盈亏情景','计算说明'];
const sheets=names.map(n=>wb.worksheets.add(n));
const [i,c,s,n]=sheets;
const rows=[
['股票名称',a.stock_name??'未提供'],['股票代码',a.stock_code??'未提供'],
['股票买入价',a.buy_price],['股票数量',a.shares],['转债认购张数',a.bond_quantity],
['转债认购价',a.bond_buy_price],['总费用（未知留空）',a.total_fees??null],
['股票买入费用',a.buy_fees??0],['股价最小报价单位',a.tick??.01],
['情景偏移（元/股）',a.offset??.1],['每股可配面值（估算）',a.allocation_per_share??null],
['配售状态',a.allocation_status??'未确认'],['参数来源',a.source??'用户输入，待确认'],
];
i.getRange('A4:B16').values=rows;
i.getRange('A2').values=[['持仓与参数']];
i.getRange('D4').values=[['黄色为可修改输入；认购张数是本次情景参数。']];
i.getRange('D6').values=[['总费用留空表示未计费用，填写0表示明确采用零费用假设。']];
i.getRange('D8').values=[['实际配售需核对发行公告与账户，估算额度不保证获配。']];
i.getRange('B6:B14').setNumberFormat('0.00');
i.getRange('B7:B8').setNumberFormat('0');
i.getRange('B14').setNumberFormat('0.0000');
i.getRange('B4:B16').format.fill='#FFF2CC';
i.getRange('B4:B16').format.font.color='#0000FF';
const I=(cell)=>"'持仓与参数'!"+cell;
c.getRange('A2').values=[['配债测算']];
c.getRange('A4:A11').values=[['股票成交金额'],['理论可配面值'],['认购资金'],['名义总投入'],['费用口径'],['认购张数'],['配售状态'],['总投入不是峰值资金占用']];
c.getRange('B4').formulas=[['='+I('B6')+'*'+I('B7')]];
c.getRange('B5').formulas=[['=IF('+I('B14')+'="","",'+I('B7')+'*'+I('B14')+')']];
c.getRange('B6').formulas=[['='+I('B8')+'*'+I('B9')]];
c.getRange('B7').formulas=[['=B4+B6+'+I('B11')]];
c.getRange('B8').formulas=[['=IF('+I('B10')+'="","毛盈亏：未计费用","净盈亏：固定费用假设")']];
c.getRange('B9').formulas=[['='+I('B8')]];
c.getRange('B10').formulas=[['='+I('B15')]];
c.getRange('B4:B7').setNumberFormat('#,##0.00');
c.getRange('D4').values=[['理论可配面值与实际认购张数分别列示，不自动保证一手。']];
s.getRange('A2').values=[['盈亏情景（人民币；收益率非年化）']];
s.getRange('A4:K4').values=[['转债价','行类型','股价','股票盈亏','转债盈亏','费用','合计盈亏','收益率','理论平衡价','费用口径','可行状态']];
let r=5;
const fee='IF('+I('B10')+'="",0,'+I('B10')+')';
for(const bp of prices){
  const anchor=r;
  for(const kind of ['理论平衡','上方0.10元','下方0.10元','相邻报价下','相邻报价上']){
    s.getRange('A'+r).values=[[bp]];
    s.getRange('B'+r).values=[[kind.replace('0.10','偏移')]];
    const be=I('B6')+'-('+I('B8')+'*(A'+r+'-'+I('B9')+')-'+fee+')/'+I('B7');
    s.getRange('I'+r).formulas=[['='+be]];
    let ps='I'+r;
    if(kind.startsWith('上方'))ps+='+'+I('B13');
    if(kind.startsWith('下方'))ps+='-'+I('B13');
    if(kind==='相邻报价下')ps='ROUNDDOWN(I'+r+'/'+I('B12')+',0)*'+I('B12');
    if(kind==='相邻报价上')ps='ROUNDUP(I'+r+'/'+I('B12')+',0)*'+I('B12');
    s.getRange('C'+r).formulas=[['=IF(I'+r+'<0,"",IF(('+ps+')<0,"",'+ps+'))']];
    s.getRange('D'+r).formulas=[['=IF(C'+r+'="","",'+I('B7')+'*(C'+r+'-'+I('B6')+'))']];
    s.getRange('E'+r).formulas=[['='+I('B8')+'*(A'+r+'-'+I('B9')+')']];
    s.getRange('F'+r).formulas=[['='+fee]];
    s.getRange('G'+r).formulas=[['=IF(C'+r+'="","",ROUND(D'+r+'+E'+r+'-F'+r+',2))']];
    s.getRange('H'+r).formulas=[['=IF(G'+r+'="","",G'+r+"/'配债测算'!B7)"]];
    s.getRange('J'+r).formulas=[["='配债测算'!B8"]];
    s.getRange('K'+r).formulas=[['=IF(C'+r+'="","非负股价范围无此情景","情景假设")']];
    s.getRange('C'+r).setNumberFormat(kind==='理论平衡'?'0.0000':'0.00');
    r++;
  }
}
const zeroStart=r+3;
s.getRange('A'+zeroStart+':E'+zeroStart).values=[['未获配情景','股价','股票盈亏','合计盈亏','股票投入收益率']];
for(const offset of [.1,0,-.1,-.3]){
  r=zeroStart+1+[.1,0,-.1,-.3].indexOf(offset);
  s.getRange('A'+r).values=[['Q=0']];
  s.getRange('B'+r).formulas=[['='+I('B6')+(offset>=0?'+':'')+offset]];
  s.getRange('C'+r).formulas=[['='+I('B7')+'*(B'+r+'-'+I('B6')+')']];
  s.getRange('D'+r).formulas=[['=C'+r+'-'+fee]];
  s.getRange('E'+r).formulas=[['=D'+r+'/('+I('B7')+'*'+I('B6')+'+'+I('B11')+')']];
  s.getRange('E'+r).setNumberFormat('0.00%');
}
n.getRange('A2').values=[['计算说明']];
n.getRange('A4:B12').values=[
['版本','1.0.1 / 2026-10-28'],
['总盈亏','N*(Ps-Pb)+Q*(Bs-Bb)-F'],
['股票平衡价','Pb-[Q*(Bs-Bb)-F]/N；固定费用假设'],
['收益率','合计盈亏/名义投入；非年化；未获配只用股票投入'],
['理论与实际','理论价4位；实际报价遵循最小报价单位；原值计算不强制改成0'],
['费用','总费用包括买入费用；未知总费用时仅算毛盈亏'],
['实现状态','两部分卖出才是实现盈亏；持有部分是浮动盈亏'],
['资金','名义投入不代表峰值占用；可能先卖股后缴款'],
['配售','实际配售张数以公告和账户为准；本表不保证发行或价格']
];
for(const sh of sheets){
 sh.showGridLines=false;
 const used=sh.getUsedRange();
 used.format.font.name='Arial';
 used.format.font.size=11;
 used.format.rowHeight=25;
 sh.getRange('A:A').format.columnWidth=30;
 sh.getRange('B:B').format.columnWidth=36;
 sh.getRange('A2').format.font.size=15;
}
for(const sh of [i,c]){sh.getRange('D:D').format.columnWidth=95;}
n.getRange('B:B').format.columnWidth=92;
s.getRange('A:K').format.columnWidth=18;
s.getRange('B:B').format.columnWidth=22;
s.getRange('J:K').format.columnWidth=30;
s.getRange('A4:K4').format.fill='#243B53';
s.getRange('A4:K4').format.font.color='#FFFFFF';
s.getRange('D5:G'+(zeroStart-3)).setNumberFormat('#,##0.00;[Red]-#,##0.00;0.00');
s.getRange('H5:H'+(zeroStart-3)).setNumberFormat('0.00%');
s.getRange('I5:I'+(zeroStart-3)).setNumberFormat('0.0000');
s.freezePanes.freezeRows(4);
const pnl=s.getRange('G5:G'+(zeroStart-3));
pnl.conditionalFormats.add('cellIs',{operator:'greaterThan',formula:0,format:{fill:'#E2F0D9',font:{color:'#216E39'}}});
pnl.conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:'#FCE4D6',font:{color:'#9C0006'}}});
wb.recalculate();
// Exercise dependencies in memory, then restore source inputs before exporting.
const originalQ=a.bond_quantity, originalF=a.total_fees??null;
i.getRange('B8').values=[[0]];
wb.recalculate();
assert.ok(Math.abs(c.getRange('B7').values[0][0]-(a.shares*a.buy_price+(a.buy_fees??0)))<1e-7);
assert.ok(Math.abs(s.getRange('I5').values[0][0]-(a.buy_price+(a.total_fees??0)/a.shares))<1e-7);
i.getRange('B8').values=[[20]];
i.getRange('B10').values=[[12]];
wb.recalculate();
assert.ok(Math.abs(s.getRange('I5').values[0][0]-(a.buy_price-(20*(prices[0]-a.bond_buy_price)-12)/a.shares))<1e-7);
i.getRange('B8').values=[[originalQ]];
i.getRange('B10').values=[[originalF]];
wb.recalculate();
const expected=a.shares*a.buy_price+a.bond_quantity*a.bond_buy_price+(a.buy_fees??0);
assert.ok(Math.abs(c.getRange('B7').values[0][0]-expected)<1e-7,'Capital calculation');
const err=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:10},maxChars:1000});
await fs.mkdir(path.dirname(output),{recursive:true});
const excel=await SpreadsheetFile.exportXlsx(wb);
await excel.save(output);
console.log(JSON.stringify({output,capital:c.getRange('B7').values,errors:err.ndjson}));
if(process.env.PREVIEW_DIR){
 await fs.mkdir(process.env.PREVIEW_DIR,{recursive:true});
 for(let k=0;k<sheets.length;k++){
  const preview=await wb.render({sheetName:names[k],range:k===2?'A1:K12':k===0?'A1:E17':'A1:D13',scale:1,format:'png'});
  await fs.writeFile(path.join(process.env.PREVIEW_DIR,k+'.png'),new Uint8Array(await preview.arrayBuffer()));
 }
}
