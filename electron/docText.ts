/**
 * Извлечение текста из .doc (Word 97-2003).
 *
 * Формат — OLE2-контейнер с потоком WordDocument и таблицей кусков (CLX),
 * которая указывает, где и в какой кодировке лежат фрагменты текста.
 * Библиотека не нужна: разбор занимает ровно столько, сколько требует
 * спецификация MS-DOC, и не тянет зависимости в portable-сборку.
 */
import * as iconv from 'iconv-lite';

const CFB_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const MAX_CHAIN = 1_000_000;

interface CfbEntry {
  name: string;
  type: number;
  start: number;
  size: number;
}

interface Cfb {
  streams: Map<string, Buffer>;
}

function readUInt32(buf: Buffer, offset: number): number {
  return offset + 4 <= buf.length ? buf.readUInt32LE(offset) : ENDOFCHAIN;
}

function parseCfb(buf: Buffer): Cfb {
  if (buf.length < 512 || !buf.subarray(0, 8).equals(CFB_SIGNATURE)) {
    throw new Error('файл не является документом Word (OLE2)');
  }

  const sectorSize = 1 << buf.readUInt16LE(0x1e);
  const miniSectorSize = 1 << buf.readUInt16LE(0x20);
  const firstDirSector = readUInt32(buf, 0x30);
  const miniCutoff = readUInt32(buf, 0x38);
  const firstMiniFat = readUInt32(buf, 0x3c);
  const numMiniFat = readUInt32(buf, 0x40);
  const firstDifat = readUInt32(buf, 0x44);
  const numDifat = readUInt32(buf, 0x48);

  const sectorOffset = (sector: number) => (sector + 1) * sectorSize;
  const readSector = (sector: number) => {
    const start = sectorOffset(sector);
    if (start + sectorSize > buf.length) return Buffer.alloc(0);
    return buf.subarray(start, start + sectorSize);
  };

  /* --- Список секторов FAT: 109 записей в заголовке + цепочка DIFAT --- */
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) {
    const sector = readUInt32(buf, 0x4c + i * 4);
    if (sector !== FREESECT && sector !== ENDOFCHAIN) fatSectors.push(sector);
  }
  let difat = firstDifat;
  for (let guard = 0; difat !== ENDOFCHAIN && guard < MAX_CHAIN && guard < numDifat + 8; guard++) {
    const sector = readSector(difat);
    if (!sector.length) break;
    const perSector = sectorSize / 4 - 1;
    for (let i = 0; i < perSector; i++) {
      const value = sector.readUInt32LE(i * 4);
      if (value !== FREESECT && value !== ENDOFCHAIN) fatSectors.push(value);
    }
    difat = sector.readUInt32LE(perSector * 4);
  }

  const fat: number[] = [];
  for (const sector of fatSectors) {
    const data = readSector(sector);
    for (let i = 0; i + 4 <= data.length; i += 4) fat.push(data.readUInt32LE(i));
  }

  const chain = (start: number): number[] => {
    const result: number[] = [];
    let sector = start;
    for (let guard = 0; sector !== ENDOFCHAIN && guard < MAX_CHAIN; guard++) {
      if (sector >= fat.length) break;
      result.push(sector);
      sector = fat[sector];
    }
    return result;
  };

  const readChain = (start: number, size?: number): Buffer => {
    const parts = chain(start).map(readSector);
    const joined = Buffer.concat(parts);
    return size !== undefined && size < joined.length ? joined.subarray(0, size) : joined;
  };

  /* --- Каталог: имя, тип, стартовый сектор и размер потока --- */
  const directory = readChain(firstDirSector);
  const entries: CfbEntry[] = [];
  for (let offset = 0; offset + 128 <= directory.length; offset += 128) {
    const nameLength = directory.readUInt16LE(offset + 0x40);
    const type = directory.readUInt8(offset + 0x42);
    if (!type) continue;
    const rawName = directory.subarray(offset, offset + Math.max(0, nameLength - 2));
    entries.push({
      name: rawName.toString('utf16le'),
      type,
      start: readUInt32(directory, offset + 0x74),
      size: readUInt32(directory, offset + 0x78),
    });
  }

  const root = entries.find((entry) => entry.type === 5);
  const miniStream = root ? readChain(root.start, root.size) : Buffer.alloc(0);

  const miniFat: number[] = [];
  if (numMiniFat > 0) {
    for (const sector of chain(firstMiniFat)) {
      const data = readSector(sector);
      for (let i = 0; i + 4 <= data.length; i += 4) miniFat.push(data.readUInt32LE(i));
    }
  }

  const readMini = (start: number, size: number): Buffer => {
    const parts: Buffer[] = [];
    let sector = start;
    for (let guard = 0; sector !== ENDOFCHAIN && guard < MAX_CHAIN; guard++) {
      const begin = sector * miniSectorSize;
      if (begin + miniSectorSize > miniStream.length) break;
      parts.push(miniStream.subarray(begin, begin + miniSectorSize));
      if (sector >= miniFat.length) break;
      sector = miniFat[sector];
    }
    const joined = Buffer.concat(parts);
    return size < joined.length ? joined.subarray(0, size) : joined;
  };

  const streams = new Map<string, Buffer>();
  for (const entry of entries) {
    if (entry.type !== 2 || !entry.name) continue;
    const data = entry.size < miniCutoff ? readMini(entry.start, entry.size) : readChain(entry.start, entry.size);
    if (data.length) streams.set(entry.name, data);
  }

  return { streams };
}

/** Убирает служебные символы Word и нормализует переводы строк. */
function cleanText(raw: string): string {
  return raw
    .replace(/\r/g, '\n')
    .replace(/[\u0000-\u0006\u0008\u0013-\u0015]/g, '')
    .replace(/\u0007/g, '\t')
    .replace(/[\u000b\u000c]/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** Текст из потока WordDocument по таблице кусков (CLX). */
function extractFromFib(word: Buffer, table: Buffer | undefined): string {
  const fcClx = readUInt32(word, 0x01a2);
  const lcbClx = readUInt32(word, 0x01a6);

  if (table && lcbClx > 0 && fcClx + lcbClx <= table.length) {
    let position = fcClx;
    const end = fcClx + lcbClx;
    while (position < end) {
      const kind = table.readUInt8(position);
      if (kind === 0x01) {
        const size = table.readUInt16LE(position + 1);
        position += 3 + size;
        continue;
      }
      if (kind !== 0x02) break;

      const size = table.readUInt32LE(position + 1);
      const plc = table.subarray(position + 5, position + 5 + size);
      const pieces = Math.floor((size - 4) / 12);
      if (pieces <= 0) break;

      const pcdStart = (pieces + 1) * 4;
      let text = '';
      for (let i = 0; i < pieces; i++) {
        const cpStart = plc.readUInt32LE(i * 4);
        const cpEnd = plc.readUInt32LE((i + 1) * 4);
        const chars = cpEnd - cpStart;
        if (chars <= 0) continue;

        let fc = plc.readUInt32LE(pcdStart + i * 8 + 2);
        const compressed = (fc & 0x40000000) !== 0;
        fc = compressed ? (fc & 0x3fffffff) >>> 1 : fc;
        if (fc >= word.length) continue;

        const bytes = compressed ? chars : chars * 2;
        const slice = word.subarray(fc, Math.min(word.length, fc + bytes));
        text += compressed ? iconv.decode(slice, 'windows-1251') : slice.toString('utf16le');
      }
      return cleanText(text);
    }
  }

  // Word 6.0/95: единый непрерывный кусок между fcMin и fcMac.
  const fcMin = readUInt32(word, 0x18);
  const fcMac = readUInt32(word, 0x1c);
  if (fcMin > 0 && fcMac > fcMin && fcMac <= word.length) {
    return cleanText(iconv.decode(word.subarray(fcMin, fcMac), 'windows-1251'));
  }

  return '';
}

/** Текст документа .doc. Бросает ошибку, если формат не Word. */
export function docBufferToText(buf: Buffer): string {
  const { streams } = parseCfb(buf);
  const word = streams.get('WordDocument');
  if (!word || word.length < 0x200) throw new Error('в файле нет потока WordDocument');

  const ident = word.readUInt16LE(0);
  if (ident !== 0xa5ec) throw new Error('недопустимая сигнатура документа Word');

  const flags = word.readUInt16LE(0x0a);
  const tableName = flags & 0x0200 ? '1Table' : '0Table';
  const table = streams.get(tableName) ?? streams.get(tableName === '1Table' ? '0Table' : '1Table');

  const text = extractFromFib(word, table);
  if (!text) throw new Error('текст документа не найден');
  return text;
}

/** Собирает HTML-абзацы из простого текста для редактора Word. */
export function textToParagraphHtml(text: string): string {
  const escape = (value: string) =>
    value.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[char] ?? char));
  const blocks = text
    .split(/\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => `<p>${escape(line)}</p>`);
  return blocks.join('') || '<p></p>';
}
