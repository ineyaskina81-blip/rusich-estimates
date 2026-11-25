'use strict';

const priceCatalog = (() => {
  const storageKey = 'rusich-price-catalog-v1';
  const xmlDecoder = new TextDecoder();
  const one = (node, name) => node.getElementsByTagNameNS('*', name)[0];
  const all = (node, name) => [...node.getElementsByTagNameNS('*', name)];
  const clean = value => String(value ?? '').trim();

  function xml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (one(doc, 'parsererror')) throw new Error('Не удалось прочитать содержимое книги Excel.');
    return doc;
  }

  // Use the ZIP central directory so archives with deferred local sizes also work.
  function archive(buffer) {
    const view = new DataView(buffer);
    let end = buffer.byteLength - 22;
    const limit = Math.max(0, end - 65535);
    while (end >= limit && (view.getUint32(end, true) !== 0x06054b50 ||
      end + 22 + view.getUint16(end + 20, true) !== buffer.byteLength)) end--;
    if (end < limit) throw new Error('Выберите книгу Excel в формате .xlsx.');
    let offset = view.getUint32(end + 16, true);
    const count = view.getUint16(end + 10, true);
    const entries = new Map();
    for (let i = 0; i < count; i++) {
      if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Повреждённая книга Excel.');
      const nameLength = view.getUint16(offset + 28, true);
      const name = xmlDecoder.decode(new Uint8Array(buffer, offset + 46, nameLength));
      entries.set(name, {
        method: view.getUint16(offset + 10, true),
        size: view.getUint32(offset + 20, true),
        expandedSize: view.getUint32(offset + 24, true),
        local: view.getUint32(offset + 42, true)
      });
      offset += 46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    }
    return async name => {
      const entry = entries.get(name);
      if (!entry) return null;
      const { local, method, size, expandedSize } = entry;
      if (expandedSize > 30_000_000 || view.getUint32(local, true) !== 0x04034b50) {
        throw new Error('Не удалось открыть лист Excel.');
      }
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      const bytes = new Uint8Array(buffer, start, size);
      if (method === 0) return xmlDecoder.decode(bytes);
      if (method !== 8) throw new Error('Формат сжатия книги Excel не поддерживается.');
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Response(stream).text();
    };
  }

  function records(sheet, strings, expectedHeader) {
    const rows = all(sheet, 'row').map(row => {
      const cells = {};
      for (const cell of row.children) {
        const column = cell.getAttribute('r').match(/^[A-Z]+/)?.[0];
        const type = cell.getAttribute('t');
        const value = one(cell, 'v')?.textContent ?? '';
        if (type === 's') cells[column] = strings[Number(value)] ?? '';
        else if (type === 'inlineStr') cells[column] = all(cell, 't').map(t => t.textContent).join('');
        else if ((!type || type === 'n') && value !== '') cells[column] = Number(value);
        else if (type !== 'e') cells[column] = value;
      }
      return cells;
    });
    const headerIndex = rows.findIndex(row => row.A === expectedHeader);
    if (headerIndex < 0) throw new Error('В книге отсутствуют заголовки справочника.');
    const headers = Object.entries(rows[headerIndex]);
    return rows.slice(headerIndex + 1).map(row =>
      Object.fromEntries(headers.map(([column, header]) => [header, row[column]])));
  }

  async function read(file) {
    if (file.size > 20_000_000) throw new Error('Размер справочника превышает 20 МБ.');
    const part = archive(await file.arrayBuffer());
    const bookText = await part('xl/workbook.xml');
    const linksText = await part('xl/_rels/workbook.xml.rels');
    if (!bookText || !linksText) throw new Error('Выберите книгу Excel в формате .xlsx.');
    const book = xml(bookText);
    const links = new Map(all(xml(linksText), 'Relationship').map(link => [
      link.getAttribute('Id'), link.getAttribute('Target')
    ]));
    const sheets = new Map(all(book, 'sheet').map(sheet => [
      sheet.getAttribute('name'), links.get(sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'))
    ]));
    const sharedText = await part('xl/sharedStrings.xml');
    const strings = sharedText
      ? all(xml(sharedText), 'si').map(item => all(item, 't').map(t => t.textContent).join(''))
      : [];
    async function sheet(name, header) {
      const target = sheets.get(name);
      if (!target) throw new Error('В книге отсутствует лист «' + name + '».');
      const text = await part(new URL(target, 'https://xlsx.invalid/xl/').pathname.slice(1));
      if (!text) throw new Error('Не удалось прочитать лист «' + name + '».');
      return records(xml(text), strings, header);
    }

    const positions = (await sheet('Позиции', 'Код позиции')).map(row => ({
      code: clean(row['Код позиции']), category: clean(row['Категория']),
      name: clean(row['Наименование']), unit: clean(row['Ед. изм.']), prices: []
    })).filter(item => item.code && item.category && item.name && item.unit);
    if (!positions.length) throw new Error('В справочнике нет позиций.');
    const byCode = new Map(positions.map(item => [item.code, item]));
    const uses1904 = ['1', 'true'].includes(one(book, 'workbookPr')?.getAttribute('date1904'));
    const epoch = uses1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    const iso = value => {
      if (typeof value === 'number' && value > 0 && value < 2_958_000) {
        return new Date(epoch + Math.floor(value) * 86400000).toISOString().slice(0, 10);
      }
      const s = clean(value);
      return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
    };
    function add(code, row, kind, value, date, expires = '', confirmed = '') {
      const position = byCode.get(clean(code));
      const priceCode = clean(row['Код цены']);
      if (!position || !priceCode || typeof value !== 'number' || value < 0.01 || value > 100000000) return;
      const unit = clean(row['Ед. изм.']);
      if (unit && unit !== position.unit) return;
      if (position.prices.some(price => price.code === priceCode)) return;
      position.prices.push({
        code: priceCode, kind, value, date, expires, confirmed,
        vat: clean(row['Режим НДС']) || 'НДС по источнику',
        source: clean(row['Файл сметы'] || row['Источник цены (ссылка или №)'] || row['Счёт №']),
        counterparty: clean(row['Поставщик'] || row['Заказчик или объект'])
      });
    }
    for (const row of await sheet('История', 'Код цены')) {
      add(row['Код позиции'], row, 'history', row['Цена, руб.'], iso(row['Дата по источнику']));
    }
    for (const [name, kind, counterparty] of [
      ['Закупка ОМТС', 'supply', 'Поставщик'], ['Для заказчика', 'customer', 'Заказчик или объект']
    ]) {
      for (const row of await sheet(name, 'Код цены')) {
        const date = iso(row['Дата цены']);
        const expires = iso(row['Действует по']);
        const confirmed = iso(row['Дата подтверждения']);
        const status = clean(row['Статус записи']);
        if (status && status !== 'Подтверждена') continue;
        if (!date || !confirmed || !clean(row[counterparty]) ||
          !clean(row['Источник цены (ссылка или №)']) || !clean(row['Подтвердил']) ||
          !['Без НДС', 'С НДС'].includes(row['Режим НДС']) ||
          (row['Действует по'] && (!expires || expires < date))) continue;
        add(row['Код позиции'], row, kind, row['Цена, руб.'], date, expires, confirmed);
      }
    }
    if (sheets.has('Счета ОМТС')) {
      for (const row of await sheet('Счета ОМТС', 'Код цены')) {
        if (row['Статус'] !== 'В справочнике' || row['Решение'] !== 'Подтвердить') continue;
        const confirmed = iso(row['Дата проверки']);
        const date = iso(row['Дата счёта']);
        if (!date || !confirmed) continue;
        add(row['Подтверждённый код'], {
          ...row, 'Ед. изм.': row['Единица для сравнения'], 'Режим НДС': row['Режим НДС цены']
        }, 'supply', row['Цена за ед., руб.'], date, iso(row['Действует по']), confirmed);
      }
    }
    return { name: file.name, positions };
  }

  function load() {
    try {
      return JSON.parse(localStorage.getItem(storageKey));
    } catch {
      return null;
    }
  }

  function save(data) {
    try {
      localStorage.setItem(storageKey, JSON.stringify(data));
      return true;
    } catch {
      return false;
    }
  }

  return { read, load, save };
})();
