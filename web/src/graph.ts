// Thin wrapper over the Microsoft Graph Excel API for one workbook in the user's OneDrive.
import { getToken } from './auth';

export type Cell = string | number | boolean | null;

export interface TableData {
  name: string;
  headers: string[];
  rows: Cell[][];
  sheet: string;
  firstCol: string;
  lastCol: string;
  /** Worksheet row number of the first data row. */
  firstDataRow: number;
}

let base = '';

export function setWorkbookPath(path: string): void {
  const p = path.startsWith('/') ? path : `/${path}`;
  base = `https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(p)}:/workbook`;
}

async function call<T>(method: string, url: string, body?: unknown): Promise<T> {
  const token = await getToken();
  const res = await fetch(base + url, {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excel request failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const type = res.headers.get('content-type') ?? '';
  return (type.includes('json') ? await res.json() : undefined) as T;
}

/** "Art!A1:AA3" or "'ISBN Lookup'!A1:G2" -> parts. */
function parseAddress(address: string) {
  const bang = address.lastIndexOf('!');
  const sheet = address.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
  const m = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(address.slice(bang + 1));
  if (!m) throw new Error(`Unexpected range address: ${address}`);
  return { sheet, firstCol: m[1], firstRow: Number(m[2]), lastCol: m[3] };
}

const sheetUrl = (sheet: string) => `/worksheets/${encodeURIComponent(sheet)}`;

function rowAddress(t: TableData, index: number): string {
  const r = t.firstDataRow + index;
  return `${t.firstCol}${r}:${t.lastCol}${r}`;
}

/** Headers and data rows of an Excel table (an empty table has one blank row). */
export async function readTable(name: string): Promise<TableData> {
  const range = await call<{ address: string; values: Cell[][] }>('GET', `/tables/${name}/range`);
  const a = parseAddress(range.address);
  return {
    name,
    headers: range.values[0].map((h) => String(h)),
    rows: range.values.slice(1),
    sheet: a.sheet,
    firstCol: a.firstCol,
    lastCol: a.lastCol,
    firstDataRow: a.firstRow + 1,
  };
}

export async function addRow(table: string, values: Cell[]): Promise<void> {
  await call('POST', `/tables/${table}/rows/add`, { values: [values] });
}

export async function readRow(t: TableData, index: number): Promise<Cell[]> {
  const r = await call<{ values: Cell[][] }>(
    'GET', `${sheetUrl(t.sheet)}/range(address='${rowAddress(t, index)}')`);
  return r.values[0];
}

export async function updateRow(t: TableData, index: number, values: Cell[]): Promise<void> {
  await call('PATCH', `${sheetUrl(t.sheet)}/range(address='${rowAddress(t, index)}')`, { values: [values] });
}

export async function deleteRow(t: TableData, index: number): Promise<void> {
  await call('POST', `${sheetUrl(t.sheet)}/range(address='${rowAddress(t, index)}')/delete`, { shift: 'Up' });
}
