// The inventory workbook: categories, dropdowns and category tables, plus the
// add / edit / move / tidy operations the form performs on them.
import * as g from './graph';
import type { Cell, TableData } from './graph';
import { bookFields, cleanIsbn, lookupIsbn } from './isbn';

export interface Category {
  name: string;
  tab: string;
  table: string;
  prefix: string;
}

export interface Item {
  category: Category;
  /** 0-based data row within the category table. */
  index: number;
  row: Cell[];
}

export type Fields = Record<string, string>;

export const AUTO_COLUMNS = ['Entered Via', 'Last Updated'];
export const DATE_COLUMNS = ['Acquired Date', 'Value Date'];
export const NUMBER_COLUMNS = ['Purchase Price', 'Estimated Value', 'Pages'];
export const LONG_COLUMNS = ['Description', 'Notes'];
const SEARCH_COLUMNS = ['Item No', 'Name', 'Creator', 'Description', 'Acquired From Name', 'Location', 'Tags', 'Notes'];
const BOOK_COLUMNS = ['Name', 'Creator', 'Author', 'Publisher', 'Publication Year', 'Pages'];

export interface DuplicateGroup {
  /** Why the items were grouped, e.g. "Same ISBN 9780140268867". */
  reason: string;
  /** "likely" for identical identifiers or names, "possible" for similar names. */
  strength: 'likely' | 'possible';
  items: Item[];
}

/** Lower-case letters and digits only, without a leading "the", "a" or "an". */
export function normalizeName(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/^(the|a|an) /, '');
}

function bigrams(s: string): Set<string> {
  const t = s.replace(/ /g, '');
  const out = new Set<string>();
  for (let i = 0; i < t.length - 1; i++) out.add(t.slice(i, i + 2));
  return out;
}

/** Dice similarity of two letter-pair sets: 1 = same letters in the same pairs, 0 = nothing shared. */
function similarity(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 || b.size === 0) return 0;
  let shared = 0;
  for (const x of a) if (b.has(x)) shared++;
  return (2 * shared) / (a.size + b.size);
}

const SIMILAR_NAME = 0.85;

/**
 * Groups items that look like the same thing entered twice: identical item number,
 * ISBN or serial number, identical names (ignoring case, punctuation and a leading
 * "The"), and very similar names. Pure function so it can be tested without Excel.
 */
export function findDuplicates(items: Item[], valueOf: (item: Item, column: string) => string): DuplicateGroup[] {
  const groups: DuplicateGroup[] = [];
  const byKey = (column: string, label: string, clean: (v: string) => string) => {
    const map = new Map<string, Item[]>();
    for (const it of items) {
      const key = clean(valueOf(it, column));
      if (key) map.set(key, [...(map.get(key) ?? []), it]);
    }
    for (const [key, list] of map) {
      if (list.length > 1) groups.push({ reason: `${label} ${key}`, strength: 'likely', items: list });
    }
  };
  byKey('Item No', 'Same item number', (v) => v.trim().toUpperCase());
  byKey('ISBN', 'Same ISBN', cleanIsbn);
  byKey('Serial Number', 'Same serial number', (v) => v.trim().toUpperCase());

  // Names: identical after normalizing, then very similar.
  const named = items
    .map((it) => ({ it, name: normalizeName(valueOf(it, 'Name')) }))
    .filter((x) => x.name.length >= 3);
  const sameName = new Map<string, Item[]>();
  for (const x of named) sameName.set(x.name, [...(sameName.get(x.name) ?? []), x.it]);
  for (const list of sameName.values()) {
    if (list.length > 1) {
      groups.push({ reason: `Same name "${valueOf(list[0], 'Name')}"`, strength: 'likely', items: list });
    }
  }

  const names = [...sameName.keys()].filter((n) => n.length >= 5);
  const pairs = names.map((n) => bigrams(n));
  for (let i = 0; i < names.length; i++) {
    for (let j = i + 1; j < names.length; j++) {
      if (similarity(pairs[i], pairs[j]) >= SIMILAR_NAME) {
        const a = sameName.get(names[i]) ?? [];
        const b = sameName.get(names[j]) ?? [];
        groups.push({
          reason: `Similar names "${valueOf(a[0], 'Name')}" and "${valueOf(b[0], 'Name')}"`,
          strength: 'possible',
          items: [...a, ...b],
        });
      }
    }
  }
  return groups;
}

/** Thrown when a row was changed in Excel after the form loaded it. */
export class ChangedError extends Error {
  constructor() {
    super('This item was changed in Excel after you opened it.');
  }
}

const isEmpty = (v: Cell | undefined) => v === null || v === undefined || v === '';
const isBlankRow = (row: Cell[]) => row.every(isEmpty);
const sameRow = (a: Cell[], b: Cell[]) => JSON.stringify(a) === JSON.stringify(b);
const pad = (n: number) => String(n).padStart(2, '0');

function nowStamp(): string {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Excel date serial -> "yyyy-mm-dd" (and time when present). */
function fromSerial(serial: number, withTime: boolean): string {
  const iso = new Date(Math.round((serial - 25569) * 86400000)).toISOString();
  return withTime ? `${iso.slice(0, 10)} ${iso.slice(11, 16)}` : iso.slice(0, 10);
}

export function displayValue(column: string, v: Cell | undefined): string {
  if (isEmpty(v)) return '';
  if (typeof v === 'number' && DATE_COLUMNS.includes(column)) return fromSerial(v, false);
  if (typeof v === 'number' && column === 'Last Updated') return fromSerial(v, true);
  return String(v);
}

/** Form text -> cell value. Dates go in as "yyyy-mm-dd", which Excel converts to dates. */
function toCell(column: string, text: string): Cell {
  const t = text.trim();
  if (t === '') return '';
  if (NUMBER_COLUMNS.includes(column)) {
    const n = Number(t.replace(/[$,]/g, ''));
    return Number.isFinite(n) ? n : t;
  }
  return t;
}

async function readList(ref: string): Promise<string[]> {
  const m = /^(\w+)(?:\[(.+)\])?$/.exec(ref.trim());
  if (!m) return [];
  const t = await g.readTable(m[1]);
  const col = m[2] ? t.headers.indexOf(m[2]) : 0;
  if (col < 0) return [];
  return t.rows.map((r) => r[col]).filter((v) => !isEmpty(v)).map(String);
}

export class Workbook {
  categories: Category[] = [];
  dropdowns = new Map<string, string[]>();
  private tables = new Map<string, TableData>();

  async load(): Promise<void> {
    const cats = await g.readTable('tblCategories');
    const col = (name: string) => cats.headers.indexOf(name);
    this.categories = cats.rows
      .filter((r) => !isEmpty(r[col('Name')]) && !isEmpty(r[col('Table')]))
      .map((r) => ({
        name: String(r[col('Name')]),
        tab: String(r[col('Tab')]),
        table: String(r[col('Table')]),
        prefix: String(r[col('Prefix')]),
      }));

    const dd = await g.readTable('tblDropdowns');
    const lists = new Map<string, string[]>();
    this.dropdowns.clear();
    for (const [column, ref] of dd.rows.map((r) => [String(r[0] ?? ''), String(r[1] ?? '')])) {
      if (!column || !ref) continue;
      if (!lists.has(ref)) lists.set(ref, await readList(ref));
      this.dropdowns.set(column, lists.get(ref) ?? []);
    }
    await Promise.all(this.categories.map((c) => this.reload(c)));
  }

  async reload(c: Category): Promise<void> {
    this.tables.set(c.table, await g.readTable(c.table));
  }

  table(c: Category): TableData {
    const t = this.tables.get(c.table);
    if (!t) throw new Error(`Table ${c.table} is not loaded.`);
    return t;
  }

  category(name: string): Category | undefined {
    return this.categories.find((c) => c.name === name);
  }

  items(c?: Category): Item[] {
    return (c ? [c] : this.categories).flatMap((cat) =>
      this.table(cat).rows
        .map((row, index) => ({ category: cat, index, row }))
        .filter((it) => !isBlankRow(it.row)));
  }

  value(item: Item, column: string): string {
    const t = this.table(item.category);
    return displayValue(column, item.row[t.headers.indexOf(column)]);
  }

  find(itemNo: string): Item | undefined {
    return this.items().find((it) => this.value(it, 'Item No') === itemNo);
  }

  search(text: string, category?: Category): Item[] {
    const q = text.trim().toLowerCase();
    return this.items(category)
      .filter((it) => !q || SEARCH_COLUMNS.some((col) => this.value(it, col).toLowerCase().includes(q)))
      .sort((a, b) => this.value(a, 'Name').localeCompare(this.value(b, 'Name')));
  }

  private nextNumberValue(c: Category): number {
    const t = this.table(c);
    const col = t.headers.indexOf('Item No');
    const re = new RegExp(`^${c.prefix}-(\\d+)$`, 'i');
    const max = Math.max(0, ...t.rows.map((r) => Number(re.exec(String(r[col] ?? ''))?.[1] ?? 0)));
    return max + 1;
  }

  private formatNumber(c: Category, n: number): string {
    return `${c.prefix}-${String(n).padStart(4, '0')}`;
  }

  /** Appends a row, reusing the blank placeholder row of an empty table. */
  private async append(c: Category, row: Cell[]): Promise<void> {
    const t = this.table(c);
    if (t.rows.length === 1 && isBlankRow(t.rows[0])) await g.updateRow(t, 0, row);
    else await g.addRow(c.table, row);
  }

  async addItem(c: Category, fields: Fields): Promise<string> {
    await this.reload(c);
    const t = this.table(c);
    const itemNo = this.formatNumber(c, this.nextNumberValue(c));
    const row = t.headers.map((h): Cell => {
      if (h === 'Item No') return itemNo;
      if (h === 'Entered Via') return 'Form';
      if (h === 'Last Updated') return nowStamp();
      return toCell(h, fields[h] ?? '');
    });
    try {
      await this.append(c, row);
    } catch (e) {
      // A timed-out add can still have gone through; look before reporting a failure.
      await this.reload(c);
      if (this.find(itemNo)) return itemNo;
      throw e;
    }
    await this.reload(c);
    return itemNo;
  }

  async updateItem(item: Item, fields: Fields, force = false): Promise<void> {
    const t = this.table(item.category);
    const current = await g.readRow(t, item.index);
    if (!force && !sameRow(current, item.row)) throw new ChangedError();
    const row = t.headers.map((h, i): Cell => {
      if (h === 'Last Updated') return nowStamp();
      if (h in fields) return toCell(h, fields[h]);
      return current[i];
    });
    await g.updateRow(t, item.index, row);
    await this.reload(item.category);
  }

  /** Moves an item to another category tab with a new number; Notes keep the old one. */
  async moveItem(item: Item, target: Category): Promise<string> {
    const src = this.table(item.category);
    const current = await g.readRow(src, item.index);
    if (!sameRow(current, item.row)) throw new ChangedError();
    await this.reload(target);
    const t = this.table(target);
    const itemNo = this.formatNumber(target, this.nextNumberValue(target));
    const oldNo = String(current[src.headers.indexOf('Item No')] ?? '') || item.category.name;
    const today = nowStamp().slice(0, 10);
    const row = t.headers.map((h): Cell => {
      const si = src.headers.indexOf(h);
      const v: Cell = si >= 0 ? current[si] : '';
      if (h === 'Item No') return itemNo;
      if (h === 'Last Updated') return nowStamp();
      if (h === 'Notes') return [v, `Moved from ${oldNo} on ${today}`].filter((x) => !isEmpty(x)).join('\n');
      return v ?? '';
    });
    await this.append(target, row);
    await g.deleteRow(src, item.index);
    await Promise.all([this.reload(target), this.reload(item.category)]);
    return itemNo;
  }

  /** Re-reads every category table, e.g. to pick up rows typed in Excel. */
  async refresh(): Promise<void> {
    await Promise.all(this.categories.map((c) => this.reload(c)));
  }

  duplicates(): DuplicateGroup[] {
    return findDuplicates(this.items(), (it, col) => this.value(it, col));
  }

  /** Removes an item's row. The last row of a table is cleared instead, because a table keeps one row. */
  async deleteItem(item: Item): Promise<void> {
    const t = this.table(item.category);
    const current = await g.readRow(t, item.index);
    if (!sameRow(current, item.row)) throw new ChangedError();
    if (t.rows.length === 1) await g.updateRow(t, item.index, t.headers.map(() => ''));
    else await g.deleteRow(t, item.index);
    await this.reload(item.category);
  }

  /** Numbers rows typed in Excel without an Item No and fills blank book details from ISBN. */
  async tidy(log: (line: string) => void): Promise<void> {
    for (const c of this.categories) {
      await this.reload(c);
      const t = this.table(c);
      const col = (name: string) => t.headers.indexOf(name);
      let next = this.nextNumberValue(c);
      let changes = 0;
      for (let i = 0; i < t.rows.length; i++) {
        const row = [...t.rows[i]];
        if (isBlankRow(row)) continue;
        let changed = false;

        if (isEmpty(row[col('Item No')])) {
          row[col('Item No')] = this.formatNumber(c, next++);
          log(`${c.name}: numbered "${row[col('Name')] ?? ''}" as ${row[col('Item No')]}`);
          changed = true;
        }

        const isbn = col('ISBN') >= 0 ? row[col('ISBN')] : null;
        const status = col('Lookup Status');
        const blanks = BOOK_COLUMNS.filter((f) => col(f) >= 0 && isEmpty(row[col(f)]));
        if (!isEmpty(isbn) && blanks.length > 0 && !(status >= 0 && row[status] === 'Not found')) {
          const info = await lookupIsbn(String(isbn));
          if (info) {
            const found = bookFields(info);
            for (const f of blanks) {
              const v = found[f];
              if (v !== undefined && v !== '') row[col(f)] = v;
            }
            if (status >= 0) row[status] = 'Found';
            log(`${c.name}: filled ${row[col('Item No')]} from ISBN ${isbn}`);
          } else {
            if (status >= 0) row[status] = 'Not found';
            log(`${c.name}: ISBN ${isbn} not found in Open Library`);
          }
          changed = true;
        }

        if (changed) {
          if (col('Last Updated') >= 0) row[col('Last Updated')] = nowStamp();
          await g.updateRow(t, i, row);
          changes++;
        }
      }
      log(`${c.name}: ${changes === 0 ? 'nothing to tidy' : `${changes} row(s) updated`}`);
    }
  }
}
