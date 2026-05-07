const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumberElement, TabStopType, TabStopPosition,
  LevelFormat, UnderlineType, PageBreak
} = require('docx');
const fs = require('fs');

// ── 食材資料 ────────────────────────────────────────────────
const foods = [
  { name: '馬鈴薯',        grams: 150, prot:  3.87, fat:  0.32, carb: 23.68, kcal_human: 115.4, kcal_pet:  99.2, fiber: 1.95, water: 120.8 },
  { name: '黃秋葵',        grams: 100, prot:  2.06, fat:  0.13, carb:  7.48, kcal_human:  35.6, kcal_pet:  34.4, fiber: 3.71, water:  89.8 },
  { name: '里肌肉（肉雞）', grams: 110, prot: 26.02, fat:  0.67, carb:  0.00, kcal_human: 117.1, kcal_pet:  96.8, fiber: 0.00, water:  83.6 },
  { name: '橄欖油',        grams:  10, prot:  0.00, fat: 10.00, carb:  0.00, kcal_human:  88.4, kcal_pet:  85.0, fiber: 0.00, water:   0.0 },
];
const total = {
  grams: 370,
  prot:  31.95, fat:   11.12, carb:  31.16,
  kcal_human: 352.5, kcal_pet: 315.4,
  fiber: 5.65, water: 294.2, na: 62, sat_fat: 1.91, sugar: 2.58,
};
const per100g = { kcal_pet: 85.2, prot: 8.6, fat: 3.0, carb: 8.4, fiber: 1.5, na: 17 };
const DER = 630;
const formula = 'FMLA-N-D10-10';

// ── 顏色與樣式 ──────────────────────────────────────────────
const DARK_GREEN = '1F4E3D';
const MED_GREEN  = '375623';
const LIGHT_GREEN = 'E2EFDA';
const BLUE = '2E75B6';
const GRAY = 'D9D9D9';
const ORANGE = 'C55A11';
const WHITE = 'FFFFFF';

const border = (color='AAAAAA') => ({
  top: { style: BorderStyle.SINGLE, size: 4, color },
  bottom: { style: BorderStyle.SINGLE, size: 4, color },
  left: { style: BorderStyle.SINGLE, size: 4, color },
  right: { style: BorderStyle.SINGLE, size: 4, color },
});
const noBorder = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};
const cellMargin = { top: 80, bottom: 80, left: 120, right: 120 };

function hdr(text, opts={}) {
  return new TextRun({ text, font:'Arial', size: opts.size||24, bold: opts.bold||false,
    color: opts.color||'000000', underline: opts.ul?{type:UnderlineType.SINGLE}:undefined });
}

function para(children, opts={}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before||0, after: opts.after||120 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: Array.isArray(children) ? children : [children],
  });
}

function spacer(n=1) {
  return Array.from({length:n}, ()=> new Paragraph({ children:[] }));
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font:'Arial', size:28, bold:true, color: DARK_GREEN })],
    spacing: { before: 240, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DARK_GREEN, space: 1 } },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font:'Arial', size:24, bold:true, color: MED_GREEN })],
    spacing: { before: 200, after: 80 },
    indent: { left: 360 },
  });
}

function bullet(text, indent=360) {
  return new Paragraph({
    numbering: { reference:'bullets', level:0 },
    indent: { left: indent+360, hanging: 360 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font:'Arial', size:22, color:'000000' })],
  });
}

// ── 表格 ────────────────────────────────────────────────────
// 表1：營養標示
function makeNutritionTable() {
  const W = 9000;
  const cols = [2800, 2200, 2000, 2000];
  
  function row(label, perServing, per100, bold=false, bg=null) {
    const sh = bg ? { fill: bg, type: ShadingType.CLEAR } : undefined;
    return new TableRow({ children: [
      new TableCell({ borders: border('AAAAAA'), width:{ size: cols[0], type: WidthType.DXA }, margins: cellMargin,
        shading: sh || { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
        children: [para([hdr(label, {bold, size:20})]) ] }),
      new TableCell({ borders: border('AAAAAA'), width:{ size: cols[1], type: WidthType.DXA }, margins: cellMargin,
        children: [para([hdr('', {size:20})]) ] }),
      new TableCell({ borders: border('AAAAAA'), width:{ size: cols[2], type: WidthType.DXA }, margins: cellMargin,
        verticalAlign: VerticalAlign.CENTER,
        children: [para([hdr(perServing, {bold, size:20})], {align: AlignmentType.RIGHT}) ] }),
      new TableCell({ borders: border('AAAAAA'), width:{ size: cols[3], type: WidthType.DXA }, margins: cellMargin,
        verticalAlign: VerticalAlign.CENTER,
        children: [para([hdr(per100, {bold, size:20})], {align: AlignmentType.RIGHT}) ] }),
    ]});
  }
  
  function rowHdr(label, col2, col3, col4, bg=DARK_GREEN, textColor=WHITE) {
    return new TableRow({ children: [
      new TableCell({ borders: border(DARK_GREEN), width:{ size: cols[0], type: WidthType.DXA }, margins: cellMargin,
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [para([hdr(label, {bold:true, size:20, color:textColor})]) ] }),
      new TableCell({ borders: border(DARK_GREEN), width:{ size: cols[1], type: WidthType.DXA }, margins: cellMargin,
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [para([hdr(col2, {bold:true, size:20, color:textColor})]) ] }),
      new TableCell({ borders: border(DARK_GREEN), width:{ size: cols[2], type: WidthType.DXA }, margins: cellMargin,
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [para([hdr(col3, {bold:true, size:20, color:textColor})], {align:AlignmentType.CENTER}) ] }),
      new TableCell({ borders: border(DARK_GREEN), width:{ size: cols[3], type: WidthType.DXA }, margins: cellMargin,
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [para([hdr(col4, {bold:true, size:20, color:textColor})], {align:AlignmentType.CENTER}) ] }),
    ]});
  }

  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      rowHdr('營養標示', '', '', ''),
      rowHdr('每一份量', `${total.grams} 公克`, '', '', MED_GREEN),
      rowHdr('本包裝含', '1 份', '', '', MED_GREEN),
      rowHdr('', '', '每份', '每 100 公克', MED_GREEN),
      row('熱量', `${total.kcal_pet.toFixed(1)} 大卡`, `${per100g.kcal_pet} 大卡`, true),
      row('蛋白質', `${total.prot.toFixed(1)} 公克`, `${per100g.prot} 公克`),
      row('脂肪', `${total.fat.toFixed(1)} 公克`, `${per100g.fat} 公克`),
      row('  飽和脂肪', `${total.sat_fat.toFixed(1)} 公克`, `${(total.sat_fat/total.grams*100).toFixed(1)} 公克`),
      row('  反式脂肪', '0.0 公克', '0.0 公克'),
      row('碳水化合物', `${total.carb.toFixed(1)} 公克`, `${per100g.carb} 公克`),
      row('  糖', `${total.sugar.toFixed(1)} 公克`, `${(total.sugar/total.grams*100).toFixed(1)} 公克`),
      row('膳食纖維', `${total.fiber.toFixed(1)} 公克`, `${per100g.fiber} 公克`),
      row('鈉', `${Math.round(total.na)} 毫克`, `${Math.round(per100g.na)} 毫克`),
    ],
  });
}

// 表2：食材配比
function makeFoodTable() {
  const W = 9000;
  const cols = [2000, 1400, 1400, 1400, 1400, 1400];
  
  function hdrRow(labels, bg=DARK_GREEN) {
    return new TableRow({ children: labels.map((l, i) => new TableCell({
      borders: border(DARK_GREEN),
      width: { size: cols[i], type: WidthType.DXA },
      margins: cellMargin,
      shading: { fill: bg, type: ShadingType.CLEAR },
      children: [para([hdr(l, {bold:true, size:20, color:WHITE})], {align: AlignmentType.CENTER})],
    }))});
  }

  function dataRow(cells, bold=false, bg=null) {
    return new TableRow({ children: cells.map((t, i) => new TableCell({
      borders: border('AAAAAA'),
      width: { size: cols[i], type: WidthType.DXA },
      margins: cellMargin,
      shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
      children: [para([hdr(t, {bold, size:20, color: bold?DARK_GREEN:'000000'})],
        { align: i===0 ? AlignmentType.LEFT : AlignmentType.RIGHT })],
    }))});
  }

  const totalBg = LIGHT_GREEN;
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      hdrRow(['核心食材', '生重 (g)', '蛋白質 (g)', '脂肪 (g)', '碳水化合物 (g)', '熱量 (kcal)']),
      ...foods.map(f => dataRow([
        f.name, String(f.grams), f.prot.toFixed(1), f.fat.toFixed(1),
        f.carb.toFixed(1), f.kcal_pet.toFixed(1),
      ])),
      dataRow([
        '單餐總計', String(total.grams), total.prot.toFixed(1), total.fat.toFixed(1),
        total.carb.toFixed(1), total.kcal_pet.toFixed(1),
      ], true, totalBg),
    ],
  });
}

// ── 頁首頁尾 ─────────────────────────────────────────────────
const header = new Header({
  children: [
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DARK_GREEN, space: 1 } },
      spacing: { after: 80 },
      children: [
        new TextRun({ text: `寵物餐食研發專案  │  精算報告  │  ${formula}`, font:'Arial', size:18, color: MED_GREEN }),
      ],
    }),
  ],
});

const footer = new Footer({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: DARK_GREEN, space: 1 } },
      spacing: { before: 60 },
      children: [
        new TextRun({ text: '中華民國 115年  │  研發配方：', font:'Arial', size:18, color: MED_GREEN }),
        new TextRun({ text: formula, font:'Arial', size:18, bold:true, color: DARK_GREEN }),
        new TextRun({ text: '  │  第 ', font:'Arial', size:18, color: MED_GREEN }),
        new PageNumberElement({ font:'Arial', size:18, color: MED_GREEN }),
        new TextRun({ text: ' 頁', font:'Arial', size:18, color: MED_GREEN }),
      ],
    }),
  ],
});

// ── 封面 ─────────────────────────────────────────────────────
const coverChildren = [
  ...spacer(4),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: '寵物餐食研發專案', font:'Arial', size:56, bold:true, color: DARK_GREEN })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new TextRun({ text: '精算報告', font:'Arial', size:48, bold:true, color: MED_GREEN })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [new TextRun({ text: 'Pet Food Research and Development Project', font:'Arial', size:24, color:'666666' })],
  }),
  ...spacer(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top:{ style:BorderStyle.SINGLE, size:4, color:DARK_GREEN }, bottom:{ style:BorderStyle.SINGLE, size:4, color:DARK_GREEN } },
    spacing: { before:120, after:120 },
    children: [new TextRun({ text: '已絕育成年犬消化之原型食材模組研發', font:'Arial', size:28, bold:true, color: DARK_GREEN })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 240 },
    children: [new TextRun({ text: 'The Effects of Whole Food Modules on Neutered Adult Dogs', font:'Arial', size:22, color:'555555' })],
  }),
  ...spacer(2),
  ...[
    ['專案編號：', 'PHDN-SKD-2026-RD07010'],
    ['研發配方：', `${formula}（黃肉甘藷、菠菜、橄欖油、里肌肉模組）`],
    ['目標對象：', '10 kg 已絕育成年犬（中型犬生理模組）'],
    ['研發負責人：', '侯沂錚 博士（Yi-Cheng Hou, Ph.D.）'],
  ].map(([label, value]) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:60, after:60 },
    children: [
      new TextRun({ text: label, font:'Arial', size:22, bold:true, color: MED_GREEN }),
      new TextRun({ text: value, font:'Arial', size:22, color:'333333' }),
    ],
  })),
  ...spacer(3),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: '中華民國 115年 05月', font:'Arial', size:22, color:'555555' })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0 },
    children: [new TextRun({ text: 'May 2026', font:'Arial', size:22, color:'999999' })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 執行摘要 ─────────────────────────────────────────────────
const execSummary = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DARK_GREEN, space: 1 } },
    children: [new TextRun({ text: '執行摘要', font:'Arial', size:36, bold:true, color: DARK_GREEN })],
  }),
  para([
    hdr('本專案針對市場對於「精準體態管理」與「原型食物機能化」的高度需求，為 10 kg、已絕育中型犬開發出專屬之營養模組 '),
    hdr(formula, {bold:true}),
    hdr('。'),
  ], {before:120, after:120}),
  para([
    hdr(`本配方整合南瓜（140g）、紫色花椰菜（100g）、橄欖油（10g）、里肌肉肉雞（110g）四種原型食材，`),
    hdr('達成「低飽和脂肪、高蛋白、精準補水」三大目標，實現產品價值優勢。'),
  ], {before:80, after:80}),
  para([
    hdr(`單餐提供寵物 ME ${total.kcal_pet.toFixed(0)} kcal（佔 DER ${(total.kcal_pet/DER*100).toFixed(0)}%），`),
    hdr(`含水量 ${Math.round(total.water)}g，蛋白質佔 ME ${(total.prot*3.5/total.kcal_pet*100).toFixed(0)}%，`),
    hdr(`脂肪佔 ME ${(total.fat*8.5/total.kcal_pet*100).toFixed(0)}%，碳水化合物佔 ME ${(total.carb*3.5/total.kcal_pet*100).toFixed(0)}%。`),
  ], {before:80, after:120}),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 目錄 ─────────────────────────────────────────────────────
const tocSection = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:0, after:200 },
    border: { bottom: { style:BorderStyle.SINGLE, size:8, color:DARK_GREEN, space:1 } },
    children: [new TextRun({ text:'目錄', font:'Arial', size:36, bold:true, color:DARK_GREEN })],
  }),
  ...[
    ['執行摘要', 'II'],
    ['目錄', 'III'],
    ['第一章  消化與生理設計重點', '1'],
    ['　第一節  消化效率與營養吸收優化', '1'],
    ['　第二節  水分代謝與泌尿系統支持', '2'],
    ['　第三節  免疫屏障與抗氧化防禦', '2'],
    ['第二章  配方營養組成與精算數據', '3'],
    ['　第一節  精確投料營養標示', '3'],
    ['　第二節  精確投料規格食材配比', '4'],
    ['第三章  技術控管與後續合規建議', '5'],
  ].map(([item, pg]) => new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 8000, leader: TabStopPosition.DOT }],
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: item, font:'Arial', size:22, color: item.startsWith('第') && !item.startsWith('　') ? DARK_GREEN : '333333', bold: item.startsWith('第') && !item.startsWith('　') }),
      new TextRun({ text: '\t' + pg, font:'Arial', size:22, color:'555555' }),
    ],
  })),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 第一章 ───────────────────────────────────────────────────
const ch1 = [
  heading1('第一章  消化與生理設計重點'),
  ...spacer(1),
  heading2('第一節  消化效率與營養吸收優化'),
  para([hdr(
    `此模組設計核心在於確保能量被有效利用，並減輕腸道負擔和胃排空速率調控。透過南瓜（140g）提供 ${total.fiber.toFixed(1)}g 膳食纖維，能有效增加食糜黏稠度，穩定血糖並延長營養素在小腸的停留時間，使消化酵素充分分解食物，顯著提升吸收效率。`
  )], {before:80, after:80}),
  para([hdr(
    `里肌肉（肉雞 110g）作為高生物價蛋白質基質，提供 ${foods[3].prot.toFixed(1)}g 優質蛋白，肌纖維細短，與南瓜纖維結合後形成極易消化的食糜塊，特別適合消化功能較弱的犬隻。橄欖油（10g）以單元不飽和脂肪酸為主，能促進脂溶性維生素（A、D、E、K）的吸收，同時提供高濃縮能量密度。`
  )], {before:80, after:160}),
  heading2('第二節  水分代謝與泌尿系統支持'),
  para([hdr(
    `此模組旨在透過飲食達到自然補水，主動預防因飲水不足引起的結石與泌尿道疾病。經由天然食物組合的代謝水分，單餐可提供約 ${Math.round(total.water)}g 的高含水量，這對於天生口渴調節機制較低的犬隻至關重要。`
  )], {before:80, after:80}),
  para([hdr(
    `高水分攝取能直接稀釋尿液中的結晶物質（如草酸鈣或磷酸銨鎂），並增加排尿量以維持下泌尿道的沖刷機制，藉此降低發炎風險，達到泌尿道健康防禦的效果。南瓜與紫色花椰菜合計約佔總水分的 ${((foods[0].water+foods[1].water)/total.water*100).toFixed(0)}%。`
  )], {before:80, after:160}),
  heading2('第三節  免疫屏障與抗氧化防禦'),
  para([hdr(
    `此模組關注長期健康與減緩細胞老化。紫色花椰菜（100g）的花青素是強效抗氧化劑，能中和消化過程中產生的氧化壓力，達到保護腸道黏膜細胞的作用。橄欖油富含多酚類抗氧化物，能協同強化抗氧化防禦機制。`
  )], {before:80, after:80}),
  para([hdr(
    `南瓜與紫色花椰菜的複合纖維（共 ${total.fiber.toFixed(1)}g）作為「益生元」，為腸道內的益生菌提供養份，可維持腸道微生態平衡，進而強化犬隻整體的免疫屏障。`
  )], {before:80, after:80}),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 第二章 ───────────────────────────────────────────────────
const ch2 = [
  heading1('第二章  配方營養組成與精算數據'),
  ...spacer(1),
  heading2('第一節  精確配方營養標示'),
  para([hdr(
    `本配方依據實際配方與生化精算結果，建立完整之營養標示資訊，以提供產品標準化與消費者清楚透明的營養資訊之依據。同時提供每 100 公克之標準化數值（寵物 ME ${per100g.kcal_pet} 大卡），以利不同餵食情境下之換算與應用。`
  )], {before:80, after:100}),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before:80, after:80 },
    children: [new TextRun({ text:`表1、10kg 已絕育成年犬單餐精確投料營養標示`, font:'Arial', size:20, bold:true, color:DARK_GREEN })],
  }),
  makeNutritionTable(),
  ...spacer(2),
  heading2('第二節  精確配方規格食材比例分配'),
  para([hdr(
    `本配方單餐供應寵物 ME ${total.kcal_pet.toFixed(0)} kcal（佔每日 DER ${DER} kcal 之 ${(total.kcal_pet/DER*100).toFixed(0)}%），經精確計量如下表 2。`
  )], {before:80, after:100}),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before:80, after:80 },
    children: [new TextRun({ text:`表2、10kg 已絕育成年犬單餐精確配方規格食材比例分配`, font:'Arial', size:20, bold:true, color:DARK_GREEN })],
  }),
  makeFoodTable(),
  ...spacer(1),
  // 熱量比例說明
  new Paragraph({
    spacing: { before:80, after:60 },
    children: [new TextRun({ text:`熱量提供比例（寵物 Modified Atwater，×3.5/×8.5/×3.5）：`, font:'Arial', size:20, bold:true, color:MED_GREEN })],
  }),
  ...[
    [`蛋白質：${(total.prot*3.5/total.kcal_pet*100).toFixed(1)}%（${total.prot.toFixed(1)}g × 3.5 = ${(total.prot*3.5).toFixed(1)} kcal）`],
    [`脂肪：${(total.fat*8.5/total.kcal_pet*100).toFixed(1)}%（${total.fat.toFixed(1)}g × 8.5 = ${(total.fat*8.5).toFixed(1)} kcal）`],
    [`碳水化合物：${(total.carb*3.5/total.kcal_pet*100).toFixed(1)}%（${total.carb.toFixed(1)}g × 3.5 = ${(total.carb*3.5).toFixed(1)} kcal）`],
    [`寵物 ME 合計：${total.kcal_pet.toFixed(1)} kcal / DER ${DER} kcal = 達成率 ${(total.kcal_pet/DER*100).toFixed(0)}%`],
  ].map(([t]) => bullet(t, 360)),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 第三章 ───────────────────────────────────────────────────
const ch3 = [
  heading1('第三章  技術控管與後續合規建議'),
  ...spacer(1),
  ...[
    ['一、數據適用性說明：',
     '本報告之配方克數考量天然食材含水量差異，以及製程條件（如加熱壓力、熟化程度）對營養數據的影響，若生產條件有所變動，建議重新進行數據校正與調整。橄欖油建議在烹調後加入，以保留多酚類活性成分。'],
    ['二、第一階段成果確認：',
     `本配方已達成精準控制熱量（寵物 ME ${total.kcal_pet.toFixed(0)} kcal）與三大營養素之合理分配（蛋白 ${(total.prot*3.5/total.kcal_pet*100).toFixed(0)}%、脂肪 ${(total.fat*8.5/total.kcal_pet*100).toFixed(0)}%、碳水 ${(total.carb*3.5/total.kcal_pet*100).toFixed(0)}%）。以寵物生理代謝而言，此組合能在控制熱量的同時滿足高蛋白需求。`],
    ['三、AAFCO 合規提醒：',
     '天然原型食材在微量營養素（如鈣、磷、維生素及礦物質）上可能存在差異與不足。本報告尚未納入 AAFCO 營養標準之全面校正，建議搭配鈣磷補充劑及綜合維生素礦物質預混料。'],
    ['四、後續發展建議：',
     '若未來規劃產品商業化並符合完整營養標準，建議進入第二階段：「AAFCO 合規校正與機能配方開發」，以補足微量營養素配置，並確保產品符合法規與安全要求。'],
  ].map(([title, body]) => [
    new Paragraph({
      spacing: { before:160, after:60 },
      children: [new TextRun({ text: title, font:'Arial', size:22, bold:true, color:DARK_GREEN })],
    }),
    para([hdr(body)], {before:40, after:80, indent:360}),
  ]).flat(),
];

// ── 組合文件 ─────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  styles: {
    default: { document: { run: { font:'Arial', size:22 } } },
    paragraphStyles: [
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{ size:28, bold:true, font:'Arial', color:DARK_GREEN },
        paragraph:{ spacing:{ before:240, after:120 }, outlineLevel:0 } },
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{ size:24, bold:true, font:'Arial', color:MED_GREEN },
        paragraph:{ spacing:{ before:200, after:80 }, outlineLevel:1 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: { default: header },
    footers: { default: footer },
    children: [
      ...coverChildren,
      ...execSummary,
      ...tocSection,
      ...ch1,
      ...ch2,
      ...ch3,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/report_02.docx', buf);
  console.log('Done: /home/claude/report_02.docx');
});