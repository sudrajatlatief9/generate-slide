(() => {
  const get = id => document.getElementById(id);
  const names = ['tren', 'Tabel 1', 'Tabel 2', 'Tabel 3', 'Oretan'];
  const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const tableGraphic = 'http://schemas.openxmlformats.org/drawingml/2006/table';
  const chartNS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const sheetNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const packageRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const serialize = doc => new XMLSerializer().serializeToString(doc);
  const read = async (zip, path) => xml(await zip.file(path).async('string'));
  const child = (node, name) => Array.from(node.children).find(item => item.localName === name);
  const isNumber = value => typeof value === 'number' && Number.isFinite(value);
  const format = (value, digits = 2) => isNumber(value)
    ? value.toLocaleString('en-US', {minimumFractionDigits: digits, maximumFractionDigits: digits})
    : String(value ?? '');
  const percent = (value, digits = 2) => isNumber(value) ? format(value * 100, digits) + '%' : format(value);
  const growth = value => (value > 0 ? '+' : '') + percent(value);
  const arrow = value => value > 0 ? '↑' : value < 0 ? '↓' : '→';
  const trendColor = value => value > 0 ? '008000' : value < 0 ? 'C00000' : '806000';
  const escape = value => String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  const span = (start, end) => Array.from({length: end - start + 1}, (_, index) => start + index);
  const getCell = (table, row, col) => table[row]?.[col];

  function columnName(index) {
    let name = '';
    while (index > 0) {
      index -= 1;
      name = String.fromCharCode(65 + index % 26) + name;
      index = Math.floor(index / 26);
    }
    return name;
  }

  function trend(table, header, rows, limit) {
    const categories = table[header] ?? [];
    const last = Math.max(categories.length, ...rows.map(row => table[row]?.length ?? 0));
    const columns = span(3, last - 1).filter(col => categories[col] != null
      && rows.some(row => isNumber(getCell(table, row, col))));
    const selected = limit ? columns.slice(-limit) : columns;
    if (!selected.length) throw Error('Data tren belum tersedia pada baris ' + header + '.');
    return {
      labels: selected.map(col => categories[col]),
      series: rows.map(row => ({
        name: table[row]?.[2] ?? '',
        values: selected.map(col => {
          const value = getCell(table, row, col);
          return isNumber(value) ? value : null;
        })
      }))
    };
  }

  async function source(file) {
    const data = await ttrWorkbook(file, names);
    const zip = await JSZip.loadAsync(file);
    const book = await read(zip, 'xl/workbook.xml');
    const relations = nodes(await read(zip, 'xl/_rels/workbook.xml.rels'), 'Relationship');
    async function sheet(name) {
      const entry = nodes(book, 'sheet').find(item => norm(item.getAttribute('name')) === norm(name));
      const relation = relations.find(item => item.getAttribute('Id') === entry.getAttributeNS(officeRel, 'id'));
      return read(zip, new URL(relation.getAttribute('Target'), 'https://local/xl/workbook.xml').pathname.slice(1));
    }
    const [areaSheet, comparisonSheet, styles] = await Promise.all([
      sheet('Tabel 2'), sheet('Tabel 3'), read(zip, 'xl/styles.xml')
    ]);
    const [tr, cards, area, comparison, rank] = names.map(name => data[name]);
    const colNumber = address => excelColumn(address.match(/[A-Z]+/)[0]);
    function rules(doc, row, column) {
      return nodes(doc, 'conditionalFormatting').filter(group =>
        group.getAttribute('sqref').split(/\s+/).some(range => {
          const [start, end = start] = range.split(':');
          return row >= Number(start.match(/\d+/)[0]) && row <= Number(end.match(/\d+/)[0])
            && column >= colNumber(start) && column <= colNumber(end);
        })
      ).flatMap(group => nodes(group, 'cfRule')).sort((a, b) => a.getAttribute('priority') - b.getAttribute('priority'));
    }
    function operand(expr, table) {
      const ref = expr.replaceAll('$', '').match(/^([A-Z]+)(\d+)$/);
      return ref ? getCell(table, Number(ref[2]), colNumber(ref[1])) : Number(expr);
    }
    function direction(row, column) {
      const value = getCell(area, row, column);
      const icon = rules(areaSheet, row, column).find(rule => rule.getAttribute('type') === 'iconSet');
      if (!icon || !isNumber(value)) return null;
      const thresholds = nodes(icon, 'cfvo').slice(1);
      if (thresholds.some(item => item.getAttribute('type') === 'percent')) return null;
      const passes = item => {
        const threshold = operand(item.getAttribute('val'), area);
        return item.getAttribute('gte') === '0' ? value > threshold : value >= threshold;
      };
      return passes(thresholds[1]) ? 1 : passes(thresholds[0]) ? 0 : -1;
    }
    const differentials = nodes(styles, 'dxf');
    function comparisonColor(row, column) {
      const value = getCell(comparison, row, column);
      if (!isNumber(value)) return null;
      for (const rule of rules(comparisonSheet, row, column)) {
        if (rule.getAttribute('type') !== 'cellIs' || !rule.hasAttribute('dxfId')) continue;
        const target = operand(first(rule, 'formula').textContent, comparison);
        const op = rule.getAttribute('operator');
        const applies = {lessThan: value < target, greaterThan: value > target, equal: value === target,
          lessThanOrEqual: value <= target, greaterThanOrEqual: value >= target, notEqual: value !== target}[op];
        if (applies) {
          const color = first(first(differentials[Number(rule.getAttribute('dxfId'))], 'font'), 'color');
          if (color?.hasAttribute('rgb')) return color.getAttribute('rgb').slice(-6);
        }
      }
      return null;
    }
    const areaTrend = trend(tr, 17, span(18, 22), 7);
    const imjas = trend(tr, 5, [12, 13], null);
    imjas.labels = imjas.labels.map(label => String(label));
    const improved = span(5, 21).filter(column => isNumber(getCell(area, 11, column)) && direction(11, column) === 1).length;
    const branch = start => span(start, start + 4).map((row, index) => [index + 1, getCell(rank, row, 5), format(getCell(rank, row, 4))]);
    return {
      charts: [
        imjas, trend(tr, 5, [10]), trend(tr, 5, span(6, 9)),
        areaTrend, trend(tr, 25, span(26, 30)), {labels: areaTrend.labels, series: [areaTrend.series[4]]}
      ],
      comparison, comparisonColor, area, direction, cards, improved,
      previousComply: tr[36]?.[3], comply: tr[36]?.[5], nonComply: tr[36]?.[6],
      best: branch(11), worst: branch(3)
    };
  }

  function writeText(node, value) {
    const texts = nodes(node, 't');
    if (texts.length) {
      texts[0].textContent = String(value ?? '');
      texts.slice(1).forEach(item => { item.textContent = ''; });
    }
  }

  function setColor(node, color) {
    for (const props of nodes(node, 'rPr')) {
      const old = child(props, 'solidFill');
      if (old) old.remove();
      if (!color) continue;
      const fill = props.ownerDocument.createElementNS(drawing, 'a:solidFill');
      const rgb = props.ownerDocument.createElementNS(drawing, 'a:srgbClr');
      rgb.setAttribute('val', color);
      fill.append(rgb);
      const next = Array.from(props.children).find(item => !['ln', 'noFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(item.localName));
      props.insertBefore(fill, next ?? null);
    }
  }

  function shape(slide, id) {
    const node = nodes(slide, 'cNvPr').find(item => item.getAttribute('id') === String(id));
    if (!node) throw Error('Objek template tidak ditemukan: ' + id);
    return node.parentNode.parentNode;
  }

  function chartWorkbook(labels, series) {
    const zip = new JSZip();
    const spreadsheet = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + spreadsheet + '.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + spreadsheet + '.worksheet+xml"/></Types>');
    const rel = (type, target) => '<Relationships xmlns="' + packageRel + '"><Relationship Id="rId1" Type="' + officeRel + '/' + type + '" Target="' + target + '"/></Relationships>';
    zip.file('_rels/.rels', rel('officeDocument', 'xl/workbook.xml'));
    zip.file('xl/_rels/workbook.xml.rels', rel('worksheet', 'worksheets/sheet1.xml'));
    zip.file('xl/workbook.xml', '<workbook xmlns="' + sheetNS + '" xmlns:r="' + officeRel + '"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    const rows = [['Periode', ...series.map(item => item.name)], ...labels.map((label, index) => [label, ...series.map(item => item.values[index])])];
    const body = rows.map((values, index) => {
      const row = index + 1;
      return '<row r="' + row + '">' + values.map((value, column) => {
        const ref = columnName(column + 1) + row;
        if (value == null) return '';
        return isNumber(value) ? '<c r="' + ref + '"><v>' + value + '</v></c>'
          : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + escape(value) + '</t></is></c>';
      }).join('') + '</row>';
    }).join('');
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="' + sheetNS + '"><sheetData>' + body + '</sheetData></worksheet>');
    return zip.generateAsync({type: 'uint8array'});
  }

  function updateReference(doc, target, formula, values, numeric) {
    if (!target) return;
    const oldCache = first(target, numeric ? 'numCache' : 'strCache') || first(target, 'numCache');
    const formatCode = oldCache && first(oldCache, 'formatCode')?.textContent;
    const make = (name, value) => {
      const element = doc.createElementNS(chartNS, 'c:' + name);
      if (value != null) element.textContent = value;
      return element;
    };
    const ref = make(numeric ? 'numRef' : 'strRef');
    const cache = make(numeric ? 'numCache' : 'strCache');
    ref.append(make('f', 'Sheet1!' + formula), cache);
    if (numeric) cache.append(make('formatCode', formatCode || 'General'));
    const count = make('ptCount');
    count.setAttribute('val', values.length);
    cache.append(count);
    values.forEach((value, index) => {
      if (value == null) return;
      const point = make('pt');
      point.setAttribute('idx', index);
      point.append(make('v', value));
      cache.append(point);
    });
    target.replaceChildren(ref);
  }

  async function updateChart(zip, index, chart) {
    const path = 'ppt/charts/chart' + index + '.xml';
    const doc = await read(zip, path);
    const entries = nodes(doc, 'ser');
    const {labels, series} = chart;
    if (index === 1) {
      const targetValues = [...nodes(first(entries[1], 'val'), 'pt')].map(point => Number(first(point, 'v').textContent));
      const target = targetValues[0];
      series.splice(1, 0, {name: 'Nilai Imjas Servo', values: labels.map(() => target)});
      for (const axis of nodes(doc, 'dateAx')) {
        const categoryAxis = doc.createElementNS(chartNS, 'c:catAx');
        for (const item of Array.from(axis.children)) {
          if (!['baseTimeUnit', 'majorTimeUnit', 'minorTimeUnit', 'majorUnit', 'minorUnit', 'numFmt'].includes(item.localName)) categoryAxis.append(item);
        }
        const scaling = first(categoryAxis, 'scaling');
        for (const tag of ['min', 'max']) first(scaling, tag)?.remove();
        axis.replaceWith(categoryAxis);
      }
    }
    if (entries.length !== series.length) throw Error('Jumlah seri grafik berbeda dari template.');
    const last = labels.length + 1;
    entries.forEach((entry, idx) => {
      const letter = columnName(idx + 2);
      updateReference(doc, child(entry, 'tx'), '$' + letter + '$1', [series[idx].name], false);
      updateReference(doc, child(entry, 'cat'), '$A$2:$A$' + last, labels, false);
      updateReference(doc, child(entry, 'val'), '$' + letter + '$2:$' + letter + '$' + last, series[idx].values, true);
      for (const tag of ['dLbl', 'dPt']) {
        for (const point of nodes(entry, tag)) {
          if (Number(first(point, 'idx')?.getAttribute('val')) >= labels.length) point.remove();
        }
      }
    });
    const rels = await read(zip, 'ppt/charts/_rels/chart' + index + '.xml.rels');
    const embed = nodes(rels, 'Relationship').find(rel => rel.getAttribute('Type').endsWith('/package'));
    const embedPath = new URL(embed.getAttribute('Target'), 'https://local/' + path).pathname.slice(1);
    zip.file(embedPath, await chartWorkbook(labels, series));
    zip.file(path, serialize(doc));
  }

  function compareTable(slide, model) {
    const rows = nodes(shape(slide, 3), 'tr');
    rows.forEach((row, index) => {
      const sourceRow = index + 5;
      nodes(row, 'tc').slice(0, 6).forEach((cell, i) => {
        const sourceColumn = i + 4;
        const value = getCell(model.comparison, sourceRow, sourceColumn);
        if (value == null || (sourceRow === 25 && sourceColumn === 9)) return;
        const asPercent = sourceRow >= 7 && (sourceColumn === 6 || sourceColumn === 9 || sourceRow === 25 && sourceColumn === 7);
        writeText(cell, asPercent ? percent(value, sourceColumn === 6 ? 0 : 2)
          : typeof value === 'number' && sourceColumn !== 4 ? format(value) : value);
        if (asPercent && (sourceColumn === 9 || sourceRow === 25)) setColor(cell, model.comparisonColor(sourceRow, sourceColumn));
      });
    });
  }

  function insight(model) {
    const current = model.area[11];
    const trend = model.charts[3];
    const previous = trend.series[4].values.at(-2);
    const areas = span(7, 10).map(row => model.area[row]);
    const up = areas.filter(row => row[23] > 0).length;
    const best = areas.reduce((a, b) => a[22] >= b[22] ? a : b);
    return [
      'KPI nasional ' + format(current[22]) + '% (' + growth(current[23]) + ' MoM)'
        + (isNumber(previous) ? ' dari ' + format(previous) + '% pada ' + trend.labels.at(-2) : '') + '.',
      up + ' Area naik. ' + best[4] + ' memimpin dengan KPI ' + format(best[22]) + '%.',
      'KPI comply ' + model.previousComply + ' → ' + model.comply + '. Dari ' + model.nonComply
        + ' KPI not comply, ' + model.improved + ' membaik.',
      'Branch tertinggi ' + model.best[0][1] + ' (' + model.best[0][2] + '%). Terendah '
        + model.worst[0][1] + ' (' + model.worst[0][2] + '%).'
    ];
  }

  function fillSlideTwo(slide, model, period, cutoff) {
    const actual = model.area[11][22], delta = model.area[11][23];
    const previous = model.charts[3].labels.at(-2) ?? '';
    const latest = model.charts[3].labels.at(-1) || period;
    const previousMonth = previous && !/^FM\s/i.test(previous) ? 'FM ' + previous : previous;
    const fields = {
      5: period, 7: 'Cut Off: ' + cutoff, 26: ' (' + period + ')', 27: format(actual) + '%',
      29: 'Growth ' + growth(delta), 30: '(vs ' + previousMonth + ')',
      682: [previousMonth + ' ' + model.previousComply + ' KPI', latest + ' ' + model.comply + ' KPI'],
      685: [model.nonComply + ' KPI NOT COMPLY', 'Namun ' + model.improved + ' KPI Improved dibanding ' + previous],
      699: insight(model)
    };
    for (const [id, content] of Object.entries(fields)) {
      const item = shape(slide, id);
      if (!Array.isArray(content)) writeText(item, content);
      else {
        const paragraphs = nodes(item, 'p').filter(paragraph => nodes(paragraph, 't')
          .some(run => run.textContent.trim()));
        if (paragraphs.length < content.length) throw Error('Paragraf template tidak lengkap: ' + id);
        paragraphs.forEach((paragraph, index) => writeText(paragraph, content[index] ?? ''));
      }
    }
    setColor(shape(slide, 29), trendColor(delta));
    const triangle = shape(slide, 28);
    first(triangle, 'xfrm').setAttribute('rot', delta < 0 ? '10800000' : '0');
    const triangleFill = first(child(triangle, 'spPr'), 'solidFill');
    if (triangleFill) {
      const rgb = slide.createElementNS(drawing, 'a:srgbClr');
      rgb.setAttribute('val', trendColor(delta));
      triangleFill.replaceChildren(rgb);
    }
    for (const [id, ranks] of [[58, model.worst], [59, model.best]]) {
      nodes(shape(slide, id), 'tr').forEach((row, index) =>
        nodes(row, 'tc').forEach((cell, column) => writeText(cell, ranks[index][column])));
    }
    for (const [id, title] of [[414, '5 BRANCH KPI TERBAIK'], [442, '5 BRANCH KPI TERENDAH']]) {
      const paragraphs = nodes(shape(slide, id), 'p');
      writeText(paragraphs[0], title);
      if (paragraphs[1]) writeText(paragraphs[1], '(' + period + ')');
    }
  }

  function editableAreaTable(slide, image, model) {
    const xfrm = first(image, 'xfrm');
    const off = first(xfrm, 'off'), ext = first(xfrm, 'ext');
    const w = Number(ext.getAttribute('cx')), h = Number(ext.getAttribute('cy'));
    const widths = span(3, 23).map(column => column === 3 ? 0.47 : column === 4 ? 1.07 : column === 22 ? 0.8 : 0.59);
    const total = widths.reduce((sum, value) => sum + value, 0);
    const grid = widths.map(width => '<a:gridCol w="' + Math.round(w * width / total) + '"/>').join('');
    const rowHeights = [0.27, 0.04, 0.09, 0.12, 0.12, 0.12, 0.12, 0.12];
    const edges = ['lnL', 'lnR', 'lnT', 'lnB'].map(edge =>
      '<a:' + edge + ' w="12000"><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></a:' + edge + '>').join('');
    const rows = span(4, 11).map(row => {
      const cells = span(3, 23).map(column => {
        const raw = getCell(model.area, row, column);
        const value = column === 17 && row === 4 ? getCell(model.area, 5, 17) + ' ' + raw : raw;
        const header = row <= 6, national = row === 11;
        const move = model.direction(row, column);
        const text = row === 5 ? '' : column === 3 ? format(value, 0) : column === 23 && !header
          ? percent(value) : format(value);
        const background = header || national ? '3E69B7' : 'FFFFFF';
        const color = header || national && (column < 5 || column > 21 || norm(value) === 'OK') ? 'FFFFFF'
          : column < 5 || column > 21 ? '000000' : norm(value) === 'OK' ? '008000' : 'FF0000';
        const merge = row === 11 && column === 3 ? ' gridSpan="2"' : row === 11 && column === 4 ? ' hMerge="1"' : '';
        const vertical = [3, 4, 22, 23].includes(column) ? 3 : 2;
        const rowMerge = row === 4 && vertical > 1 ? ' rowSpan="' + vertical + '"' : row > 4 && row < 4 + vertical ? ' vMerge="1"' : '';
        const size = row === 4 && column >= 5 && column <= 21 ? 645 : 770;
        const run = (label, rgb, font = 'Arial', fontSize = size) => '<a:r><a:rPr lang="id-ID" sz="' + fontSize
          + '" b="1"><a:solidFill><a:srgbClr val="' + rgb + '"/></a:solidFill><a:latin typeface="' + font
          + '"/></a:rPr><a:t>' + escape(label) + '</a:t></a:r>';
        const icon = move > 0 ? '⬆' : move < 0 ? '⬇' : '➡';
        const iconColor = move > 0 ? '529781' : move < 0 ? 'D25531' : trendColor(move);
        const rotation = row === 4 && column >= 5 && column <= 21 ? ' vert="vert270"' : '';
        return '<a:tc' + merge + rowMerge + '><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
          + (move == null || header ? '' : run(icon + '\u00A0', iconColor, 'Segoe UI Symbol', 1000))
          + run(text, color) + '</a:p></a:txBody><a:tcPr' + rotation + ' marL="12000" marR="12000" marT="0" marB="0" anchor="ctr">'
          + edges + '<a:solidFill><a:srgbClr val="' + background + '"/></a:solidFill></a:tcPr></a:tc>';
      }).join('');
      return '<a:tr h="' + Math.round(h * rowHeights[row - 4]) + '">' + cells + '</a:tr>';
    }).join('');
    const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const frame = xml('<p:graphicFrame xmlns:p="' + p + '" xmlns:a="' + drawing + '">'
      + '<p:nvGraphicFramePr><p:cNvPr id="2" name="Tabel KPI Nasional dan Per Area"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
      + '<p:xfrm><a:off x="' + off.getAttribute('x') + '" y="' + off.getAttribute('y') + '"/><a:ext cx="' + w + '" cy="' + h + '"/></p:xfrm>'
      + '<a:graphic><a:graphicData uri="' + tableGraphic + '"><a:tbl><a:tblPr/><a:tblGrid>' + grid
      + '</a:tblGrid>' + rows + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>');
    image.replaceWith(slide.importNode(frame.documentElement, true));
  }

  function fillSlideThree(slide, model) {
    const groups = [105, 104, 106, 107, 109, 110, 111, 112, 113, 114];
    groups.forEach((id, index) => {
      const column = index + 2;
      const group = shape(slide, id);
      const cardShapes = Array.from(group.children).filter(item => item.localName === 'sp');
      const combined = item => nodes(item, 't').map(text => text.textContent).join('');
      const value = cardShapes.find(item => /%$/.test(combined(item)));
      const title = cardShapes[cardShapes.indexOf(value) - 1];
      const delta = getCell(model.cards, 10, column);
      writeText(title, getCell(model.cards, 4, column));
      const valueRuns = nodes(value, 't');
      valueRuns[0].textContent = arrow(delta);
      if (valueRuns.length > 1) {
        valueRuns[1].textContent = percent(getCell(model.cards, 5, column), [6, 8, 9, 10].includes(column) ? 0 : 1);
        valueRuns.slice(2).forEach(run => { run.textContent = ''; });
        setColor(valueRuns[0].parentNode, trendColor(delta));
        const font = first(valueRuns[0].parentNode, 'latin');
        if (font) font.setAttribute('typeface', 'Arial');
      }
      const details = (id === 109 ? Array.from(shape(slide, 108).children) : cardShapes)
        .find(item => /Target|Real\s*:/.test(combined(item)));
      if (!details) return;
      const target = getCell(model.cards, 6, column), actual = getCell(model.cards, 7, column);
      const isCount = column === 11;
      const detail = typeof target === 'string' && target.includes('/')
        ? 'Real : ' + target + ' indikator ' + actual + '%'
        : 'Target : ' + format(target, isCount ? 0 : 2) + '\nReal : ' + format(actual, isCount ? 0 : 2);
      const paragraphs = nodes(details, 'p');
      if (paragraphs.length > 1 && detail.includes('\n')) {
        const lines = detail.split('\n');
        paragraphs.forEach((item, index) => writeText(item, lines[index] ?? ''));
      } else {
        const texts = nodes(details, 't');
        if (detail.includes('\n') && texts.length >= 2) {
          const [targetLine, actualLine] = detail.split('\n');
          const actualRun = texts.find(run => /^Real/.test(run.textContent));
          if (actualRun) {
            texts.forEach(run => { run.textContent = ''; });
            texts[0].textContent = targetLine;
            actualRun.textContent = actualLine;
          } else writeText(details, detail.replace('\n', '   '));
        } else writeText(details, detail);
      }
    });
    editableAreaTable(slide, shape(slide, 2), model);
  }

  async function makeDeck(model, period, cutoff) {
    const zip = await JSZip.loadAsync(await assetBytes('assets/summary-kpi-ioan.pptx'));
    const slides = await Promise.all(span(1, 3).map(i => read(zip, 'ppt/slides/slide' + i + '.xml')));
    writeText(shape(slides[0], 1048635), 'Tren Pencapaian Performance IOAN - ' + period);
    writeText(shape(slides[2], 1048635), 'Assurance Performance IOAN - ' + period);
    for (const slide of [slides[0], slides[2]]) writeText(shape(slide, 1048685), 'Cut off : ' + cutoff);
    compareTable(slides[0], model);
    fillSlideTwo(slides[1], model, period, cutoff);
    fillSlideThree(slides[2], model);
    slides.forEach((slide, index) => zip.file('ppt/slides/slide' + (index + 1) + '.xml', serialize(slide)));
    await Promise.all(model.charts.map((chart, i) =>
      updateChart(zip, i + 1, {labels: chart.labels, series: [...chart.series]})));
    return zip.generateAsync({type: 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
  }

  const menu = get('menuSummary');
  const view = get('summaryView');
  const file = get('summaryFile');
  const period = get('summaryPeriod');
  const date = get('summaryDate');
  const download = get('summaryExport');
  const message = get('summaryMsg');
  const preview = get('summaryPreview');
  let current, loading = false, exporting = false, sequence = 0;

  function render() {
    download.disabled = !current || loading || exporting || !period.value.trim() || !date.value;
    preview.replaceChildren();
    if (!current) return;
    const list = el('ul', preview);
    insight(current).forEach(line => el('li', list, line));
  }

  menu.onclick = () => {
    for (const other of document.querySelectorAll('.reportMenu button')) other.classList.toggle('active', other === menu);
    for (const section of document.querySelectorAll('main')) section.hidden = section !== view;
  };
  for (const other of document.querySelectorAll('.reportMenu button:not(#menuSummary)')) {
    other.addEventListener('click', () => {
      view.hidden = true;
      menu.classList.remove('active');
    });
  }
  file.onchange = async () => {
    const request = ++sequence;
    current = null;
    loading = Boolean(file.files[0]);
    render();
    if (!loading) {
      message.textContent = 'Unggah workbook Formula Radir IOAN.';
      return;
    }
    message.textContent = 'Memuat Excel…';
    try {
      const model = await source(file.files[0]);
      if (request !== sequence) return;
      current = model;
      period.value = periodFromName(file.files[0].name) || String(current.charts[3].labels.at(-1));
      message.textContent = 'Excel termuat. Periksa periode dan isi tanggal cut-off.';
    } catch (error) {
      if (request === sequence) message.textContent = 'Gagal memuat Excel. ' + error.message;
    } finally {
      if (request === sequence) {
        loading = false;
        render();
      }
    }
  };
  period.oninput = date.oninput = render;
  download.onclick = async () => {
    if (download.disabled) return;
    exporting = true;
    message.textContent = 'Menyiapkan PowerPoint…';
    render();
    try {
      const title = period.value.trim();
      const cutoff = new Date(date.value + 'T00:00:00Z').toLocaleDateString('id-ID', {
        day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC'
      });
      const blob = await makeDeck(current, title, cutoff);
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'Summary_KPI_IOAN_B2C_' + title.replace(/[^\w-]+/g, '_') + '.pptx';
      link.click();
      message.textContent = 'PowerPoint siap diunduh.';
      setTimeout(() => URL.revokeObjectURL(link.href), 10000);
    } catch (error) {
      message.textContent = 'Ekspor gagal. ' + error.message;
    } finally {
      exporting = false;
      render();
    }
  };
})();
