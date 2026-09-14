import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const outDir = 'outputs/claim_dashboard';
await fs.mkdir(outDir, { recursive: true });
const wb = Workbook.create();
const dash = wb.worksheets.add('Dashboard');
const raw = wb.worksheets.add('Raw_Data');
const lists = wb.worksheets.add('Lists');
dash.showGridLines = false; raw.showGridLines = true; lists.showGridLines = false;

const navy = '#17365D', blue = '#D9EAF7', pale = '#F7FBFE', red = '#C00000', gray = '#666666';
const font = 'Arial';

// Raw data input area
raw.getRange('A1:H1').values = [['월','차종','부품명','클레임건수','현상','사용기간구간','주행거리구간','국가']];
const sample = [
 ['2026-01','MQ4 20 HEV','GASKET',1,'누수','~6개월','~1만','한국'],
 ['2026-01','KA4','SDPF',4,'경고등 점등','12~18개월','~2만','한국'],
 ['2026-02','MQ4 20','MUFFLER ASSY-CTR',3,'소음','18~24개월','~3만','터키'],
 ['2026-02','JA 17','MUFFLER ASSY-RR',2,'출력부족','24~30개월','~4만','한국'],
 ['2026-03','MQ4 20 HEV','GASKET',2,'누수','6~12개월','~2만','대만'],
 ['2026-03','KA4','SDPF',5,'경고등 점등','24~30개월','~5만','한국'],
 ['2026-04','MQ4 20','MUFFLER ASSY-CTR',4,'소음','30~36개월','~6만','브라질'],
 ['2026-04','JA 17','MUFFLER ASSY-FR',2,'출력부족','36~42개월','~7만','한국'],
 ['2026-05','KA4','SDPF',6,'경고등 점등','12~18개월','~3만','한국'],
 ['2026-05','MQ4 20 HEV','GASKET',2,'누수','18~24개월','~4만','터키'],
 ['2026-06','MQ4 20','MUFFLER ASSY-CTR',3,'소음','24~30개월','~5만','한국'],
 ['2026-06','JA 17','MUFFLER ASSY-RR',1,'부식','42~48개월','~8만','팔레스타인'],
];
raw.getRange(`A2:H${sample.length+1}`).values = sample;
raw.getRange(`A1:H${sample.length+1}`).format.font = { name: font, size: 10 };
raw.getRange('A1:H1').format = { fill: navy, font: { name: font, bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center', verticalAlignment: 'center' };
raw.getRange(`A1:H${sample.length+1}`).format.borders = { preset: 'all', style: 'thin', color: '#D9E2F3' };
raw.getRange(`D2:D${sample.length+1}`).format.numberFormat = '#,##0';
raw.getRange('A:A').format.columnWidth = 13; raw.getRange('B:B').format.columnWidth = 18; raw.getRange('C:C').format.columnWidth = 24;
raw.getRange('D:D').format.columnWidth = 12; raw.getRange('E:E').format.columnWidth = 18; raw.getRange('F:G').format.columnWidth = 15; raw.getRange('H:H').format.columnWidth = 14;
raw.freezePanes.freezeRows(1);
const table = raw.tables.add(`A1:H${sample.length+1}`, true, 'ClaimsInput');
table.style = 'TableStyleMedium2'; table.showFilterButton = true;

// Lists for user-editable analysis categories
lists.getRange('A1:F1').values = [['월','현상','사용기간구간','주행거리구간','국가','차종']];
lists.getRange('A2:A13').values = [['2026-01'],['2026-02'],['2026-03'],['2026-04'],['2026-05'],['2026-06'],['2026-07'],['2026-08'],['2026-09'],['2026-10'],['2026-11'],['2026-12']];
lists.getRange('B2:B8').values = [['경고등 점등'],['매연과다'],['출력부족'],['DPF 불량'],['오일유입'],['소음'],['부식']];
lists.getRange('C2:C10').values = [['~6개월'],['6~12개월'],['12~18개월'],['18~24개월'],['24~30개월'],['30~36개월'],['36~42개월'],['42~48개월'],['48~54개월']];
lists.getRange('D2:D10').values = [['~1만'],['~2만'],['~3만'],['~4만'],['~5만'],['~6만'],['~7만'],['~8만'],['~9만']];
lists.getRange('E2:E6').values = [['한국'],['터키'],['대만'],['브라질'],['팔레스타인']];
lists.getRange('F2:F6').values = [['MQ4 20 HEV'],['KA4'],['MQ4 20'],['JA 17'],['기타']];
lists.getRange('A1:F1').format = { fill: navy, font: { name: font, bold: true, color: '#FFFFFF' } };
lists.getRange('A:F').format.columnWidth = 16;

// Dashboard title and notes
dash.getRange('A1:N1').merge(); dash.getRange('A1').values = [['클레임 데이터 자동 분석']];
dash.getRange('A1:N1').format = { font: { name: font, size: 18, bold: true, color: navy }, verticalAlignment: 'center' };
dash.getRange('A2:N2').merge(); dash.getRange('A2').values = [['Raw_Data 시트의 A:H 열에 데이터를 붙여넣으면 아래 표와 그래프가 자동 갱신됩니다. 분석 기준값은 Lists 시트에서 수정할 수 있습니다.']];
dash.getRange('A2:N2').format = { font: { name: font, size: 10, italic: true, color: gray } };
dash.getRange('A4:B4').values = [['핵심 지표','값']];
dash.getRange('A5:A7').values = [['총 클레임 건수'],['분석 월 수'],['최다 발생 현상']];
dash.getRange('B5:B7').formulas = [['=SUM(Raw_Data!$D$2:$D$1000)'],['=COUNTA(Lists!$A$2:$A$13)'],['=INDEX(Lists!$B$2:$B$8,MATCH(MAX(E12:E18),E12:E18,0))']];
dash.getRange('A4:B7').format.borders = { preset: 'all', style: 'thin', color: '#D9E2F3' };
dash.getRange('A4:B4').format = { fill: navy, font: { name: font, bold: true, color: '#FFFFFF' } };
dash.getRange('A5:A7').format = { fill: blue, font: { name: font, bold: true, color: navy } };
dash.getRange('B5:B7').format = { fill: pale, font: { name: font, bold: true, color: navy } };
dash.getRange('B5').format.numberFormat = '#,##0';

function summaryBlock(startCol, title, listRange, formulaCol, dataStart, dataEnd) {
  const col = String.fromCharCode(65 + startCol);
  const next = String.fromCharCode(65 + startCol + 1);
  dash.getRange(`${col}10:${next}10`).merge(); dash.getRange(`${col}10`).values = [[title]];
  dash.getRange(`${col}10:${next}10`).format = { fill: navy, font: { name: font, bold: true, color: '#FFFFFF' } };
  dash.getRange(`${col}11:${next}11`).values = [['구분','건수']];
  dash.getRange(`${col}11:${next}11`).format = { fill: blue, font: { name: font, bold: true, color: navy } };
  dash.getRange(`${col}12:${col}${11+dataEnd-dataStart+1}`).formulas = Array.from({length:dataEnd-dataStart+1},(_,i)=>[`=Lists!${listRange.split(':')[0].replace(/[0-9]+/,'')}${dataStart+i}`]);
  const formulas = Array.from({length:dataEnd-dataStart+1},(_,i)=>[`=SUMIF(Raw_Data!$${formulaCol}$2:$${formulaCol}$1000,${col}${12+i},Raw_Data!$D$2:$D$1000)`]);
  dash.getRange(`${next}12:${next}${11+dataEnd-dataStart+1}`).formulas = formulas;
  dash.getRange(`${col}11:${next}${11+dataEnd-dataStart+1}`).format.borders = { preset: 'all', style: 'thin', color: '#D9E2F3' };
  dash.getRange(`${next}12:${next}${11+dataEnd-dataStart+1}`).format.numberFormat = '#,##0';
  return {cat:`${col}11:${col}${11+dataEnd-dataStart+1}`, val:`${next}11:${next}${11+dataEnd-dataStart+1}`};
}
// Monthly trend (visible helper table)
dash.getRange('A10:B10').merge(); dash.getRange('A10').values = [['월별 발생 추이']]; dash.getRange('A10:B10').format = { fill: navy, font: { name: font, bold: true, color: '#FFFFFF' } };
dash.getRange('A11:B11').values = [['월','건수']]; dash.getRange('A11:B11').format = { fill: blue, font: { name: font, bold: true, color: navy } };
dash.getRange('A12:A23').formulas = Array.from({length:12},(_,i)=>[`=Lists!A${2+i}`]);
dash.getRange('B12:B23').formulas = Array.from({length:12},(_,i)=>[`=SUMIF(Raw_Data!$A$2:$A$1000,A${12+i},Raw_Data!$D$2:$D$1000)`]);
dash.getRange('A11:B23').format.borders = { preset: 'all', style: 'thin', color: '#D9E2F3' }; dash.getRange('B12:B23').format.numberFormat = '#,##0';
const b1 = summaryBlock(3,'현상별 분석','B2:B8','E',2,8);
const b2 = summaryBlock(6,'사용기간 분석','C2:C10','F',2,10);
const b3 = summaryBlock(9,'주행거리 분석','D2:D10','G',2,10);
const b4 = summaryBlock(12,'국가별 분석','E2:E6','H',2,6);

// Charts
function addChart(type, range, title, pos, color){ const c = dash.charts.add(type, dash.getRange(range)); c.title = title; c.titleTextStyle.typeface = font; c.titleTextStyle.fontSize = 12; c.hasLegend = false; c.xAxis = {axisType:'textAxis', textStyle:{typeface:font, fontSize:9}}; c.yAxis = {numberFormatCode:'#,##0', numberFormatSourceLinked:false, textStyle:{typeface:font, fontSize:9}}; c.setPosition(pos[0],pos[1]); return c; }
addChart('line','A11:B23','월별 클레임 발생 추이',['D4','J18'],'#4472C4');
addChart('bar','D11:E18','현상별 클레임',['K4','Q18'],'#5B9BD5');
addChart('bar','G11:H20','사용기간별 클레임',['D19','J33'],'#70AD47');
addChart('bar','J11:K20','주행거리별 클레임',['K19','Q33'],'#ED7D31');
addChart('bar','M11:N16','국가별 클레임',['D34','J48'],'#A5A5A5');

dash.getRange('A:A').format.columnWidth = 16; dash.getRange('B:B').format.columnWidth = 12; dash.getRange('D:D').format.columnWidth = 18; dash.getRange('E:E').format.columnWidth = 12; dash.getRange('G:G').format.columnWidth = 16; dash.getRange('H:H').format.columnWidth = 12; dash.getRange('J:J').format.columnWidth = 14; dash.getRange('K:K').format.columnWidth = 12; dash.getRange('M:M').format.columnWidth = 14; dash.getRange('N:N').format.columnWidth = 12;
dash.getRange('A1:N50').format.font = { name: font, size: 10 };
dash.getRange('A1:N1').format.font = { name: font, size: 18, bold: true, color: navy };
dash.freezePanes.freezeRows(3);

const errors = await wb.inspect({kind:'match', searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options:{useRegex:true,maxResults:100}, summary:'formula errors'});
console.log(errors.ndjson);
const preview = await wb.render({sheetName:'Dashboard', range:'A1:Q48', scale:1, format:'png'});
await fs.writeFile(`${outDir}/dashboard_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outDir}/클레임_자동분석_템플릿.xlsx`);
console.log('saved');
