const ExcelJS = require('exceljs');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
      body: ''
    };
  }

  try {
    const { date, site, person, sections, special } = JSON.parse(event.body || '{}');

    const parts = (date || '').split('-');
    const ym = parts.length >= 2 ? `${parts[0]}年 ${parseInt(parts[1])}月度` : '';

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet(site ? site.slice(0, 31) : '点検表');

    // 列幅
    ws.columns = [
      { width: 12.9 }, // A
      { width: 21.6 }, // B
      { width: 40.0 }, // C
      { width: 7.6  }, // D
      { width: 7.6  }, // E
      { width: 7.6  }, // F
      { width: 12.0 }, // G
      { width: 3.3  }, // H
      { width: 12.0 }, // I
      { width: 8.1  }, // J
      { width: 4.6  }, // K
    ];

    const FONT = { name: '游ゴシック' };

    function borderSide(style) { return style ? { style } : undefined; }
    function makeBorder(top, bottom, left, right) {
      const b = {};
      if (top)    b.top    = borderSide(top);
      if (bottom) b.bottom = borderSide(bottom);
      if (left)   b.left   = borderSide(left);
      if (right)  b.right  = borderSide(right);
      return b;
    }

    // ── 行1: タイトル ──
    ws.getRow(1).height = 30;
    ws.mergeCells('A1:K1');
    const t1 = ws.getCell('A1');
    t1.value = `${ym}清掃点検表`;
    t1.font = { ...FONT, size: 22, bold: true };
    t1.alignment = { horizontal: 'center', vertical: 'middle' };

    // ── 行2: 点検日 ──
    ws.getRow(2).height = 21.75;
    const g2 = ws.getCell('G2');
    g2.value = '点検日　：';
    g2.font = { ...FONT, size: 11 };
    g2.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.mergeCells('H2:K2');
    const h2 = ws.getCell('H2');
    h2.value = date || '';
    h2.font = { ...FONT, size: 12 };
    h2.alignment = { horizontal: 'left', vertical: 'middle' };

    // ── 行3: 現場名・点検者 ──
    ws.getRow(3).height = 27;
    ws.mergeCells('B3:C3');
    const b3 = ws.getCell('B3');
    b3.value = `現場名 ： ${site || ''}`;
    b3.font = { ...FONT, size: 14, bold: true };
    b3.alignment = { horizontal: 'left', vertical: 'middle' };
    const g3 = ws.getCell('G3');
    g3.value = '点検者　：';
    g3.font = { ...FONT, size: 11 };
    g3.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.mergeCells('H3:K3');
    const h3 = ws.getCell('H3');
    h3.value = person || '';
    h3.font = { ...FONT, size: 12 };
    h3.alignment = { horizontal: 'left', vertical: 'middle' };

    // ── 行4: 空白 ──
    ws.getRow(4).height = 9.75;

    // ── 行5: ヘッダー ──
    ws.getRow(5).height = 24.75;
    const headers = [
      { col: 'A', label: '場所',        border: makeBorder('medium','double','medium','thin') },
      { col: 'B', label: '点検項目',    border: makeBorder('medium','double','thin','thin') },
      { col: 'C', label: '評価ポイント',border: makeBorder('medium','double','thin','thin') },
      { col: 'D', label: '良',          border: makeBorder('medium','double','thin','thin') },
      { col: 'E', label: '可',          border: makeBorder('medium','double','thin','thin') },
      { col: 'F', label: '不可',        border: makeBorder('medium','double','thin','thin') },
      { col: 'G', label: '備　　　考',  border: makeBorder('medium','double','thin','medium') },
    ];
    headers.forEach(({ col, label, border }) => {
      const cell = ws.getCell(`${col}5`);
      cell.value = label;
      cell.font = { ...FONT, size: 11 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = border;
    });
    ws.mergeCells('G5:K5');
    ['H','I','J'].forEach(c => {
      ws.getCell(`${c}5`).border = makeBorder('medium','double',null,null);
    });
    ws.getCell('K5').border = makeBorder('medium','double',null,'medium');

    // ── データ行 ──
    let currentRow = 6;
    (sections || []).forEach(sec => {
      const items = sec.items || [];
      const n = items.length;
      if (n > 1) {
        ws.mergeCells(`A${currentRow}:A${currentRow + n - 1}`);
      }
      items.forEach((item, i) => {
        const r = currentRow + i;
        ws.getRow(r).height = 23.25;
        const isFirst = i === 0;
        const isLast  = i === n - 1;
        const botB  = isLast ? 'medium' : 'thin';
        const topB  = isFirst ? null : 'thin';
        const topBG = isFirst ? null : 'thin';

        // A: 場所
        const aCell = ws.getCell(`A${r}`);
        if (isFirst) {
          aCell.value = sec.name || '';
          aCell.font = { ...FONT, size: 16 };
          aCell.alignment = { horizontal: 'center', vertical: 'middle' };
          aCell.border = makeBorder('double', isLast ? 'medium' : null, 'medium', 'thin');
        } else {
          aCell.border = makeBorder(null, isLast ? 'medium' : null, 'medium', 'thin');
        }

        // B: 点検項目
        const bCell = ws.getCell(`B${r}`);
        bCell.value = item.name || '';
        bCell.font = { ...FONT, size: 11 };
        bCell.alignment = { vertical: 'middle' };
        bCell.border = makeBorder(topB, botB, 'thin', 'thin');

        // C: 評価ポイント
        const cCell = ws.getCell(`C${r}`);
        cCell.value = item.points || '';
        cCell.font = { ...FONT, size: 11 };
        cCell.alignment = { vertical: 'middle' };
        cCell.border = makeBorder(topB, botB, 'thin', 'thin');

        // D/E/F: 良/可/不可
        const evalVal = item.eval || '';
        [['D','良'],['E','可'],['F','不可']].forEach(([col, key]) => {
          const cell = ws.getCell(`${col}${r}`);
          cell.value = evalVal === key ? '○' : '　';
          cell.font = { ...FONT, size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = makeBorder(topB, botB, 'thin', 'thin');
        });

        // G〜K: 備考（マージ）
        ws.mergeCells(`G${r}:K${r}`);
        const gCell = ws.getCell(`G${r}`);
        gCell.value = item.note || '';
        gCell.font = { ...FONT, size: 11 };
        gCell.alignment = { horizontal: 'left', vertical: 'middle' };
        gCell.border = makeBorder(topBG, null, 'thin', 'medium');
        ['H','I','J'].forEach(c => {
          ws.getCell(`${c}${r}`).border = makeBorder(topBG, null, null, null);
        });
        ws.getCell(`K${r}`).border = makeBorder(topBG, null, null, 'medium');
      });
      currentRow += n;
    });

    // ── 空白行 ──
    ws.getRow(currentRow).height = 11.25;
    currentRow++;

    // ── 特記事項ヘッダー ──
    const shr = currentRow;
    ws.getRow(shr).height = 19.5;
    const shrCell = ws.getCell(`A${shr}`);
    shrCell.value = '特記事項';
    shrCell.font = { ...FONT, size: 11 };
    shrCell.alignment = { horizontal: 'center', vertical: 'middle' };
    shrCell.border = makeBorder('medium', 'thin', 'medium', 'thin');
    for (let c = 2; c <= 11; c++) {
      const cell = ws.getCell(shr, c);
      cell.border = makeBorder('medium', null, null, c === 11 ? 'medium' : null);
    }
    currentRow++;

    // ── 特記事項本文 ──
    const sbr = currentRow;
    ws.getRow(sbr).height   = 40.5;
    ws.getRow(sbr+1).height = 20.25;
    ws.getRow(sbr+2).height = 29.25;
    ws.mergeCells(`A${sbr}:K${sbr+2}`);
    const sbrCell = ws.getCell(`A${sbr}`);
    sbrCell.value = special || '';
    sbrCell.font = { ...FONT, size: 11 };
    sbrCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    sbrCell.border = makeBorder(null, 'medium', 'medium', 'medium');

    // ── バッファ出力 ──
    const buffer = await workbook.xlsx.writeBuffer();
    const base64 = Buffer.from(buffer).toString('base64');
    const filename = encodeURIComponent(`点検表_${site}_${date}.xlsx`);

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${filename}`,
        'Access-Control-Allow-Origin': '*',
      },
      body: base64,
      isBase64Encoded: true,
    };

  } catch (err) {
    return {
      statusCode: 500,
      headers: { 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
