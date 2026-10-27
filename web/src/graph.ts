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

/** A Graph request that failed; `status` is the HTTP status (404 = not found). */
export class GraphError extends Error {
  constructor(readonly status: number, detail: string) {
    super(`Excel request failed (${status}): ${detail}`);
  }
}

export function setWorkbookPath(path: string): void {
  const p = path.startsWith('/') ? path : `/${path}`;
  base = `https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(p)}:/workbook`;
}

// One Excel session for all requests: the workbook is opened once instead of on every call,
// which is faster and avoids most "workbook cannot be opened" timeouts.
let session: string | null = null;

const RETRY_STATUSES = [429, 502, 503, 504];
const MAX_ATTEMPTS = 4;
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function send(method: string, url: string, body: unknown, withSession: boolean): Promise<Response> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${await getToken()}`,
    'Content-Type': 'application/json',
  };
  if (withSession && session) headers['workbook-session-id'] = session;
  return fetch(base + url, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
}

/**
 * Sends a request, retrying timeouts and throttling. `retry` must be false for requests
 * that are not safe to repeat (adding or deleting rows): those fail instead, and the
 * caller checks what happened.
 */
async function call<T>(method: string, url: string, body?: unknown, retry = method !== 'POST'): Promise<T> {
  for (let attempt = 1; ; attempt++) {
    if (!session) await openSession();
    const res = await send(method, url, body, true);
    if (res.ok) {
      const type = res.headers.get('content-type') ?? '';
      return (type.includes('json') ? await res.json() : undefined) as T;
    }
    const text = await res.text();
    if (/session/i.test(text) && [400, 404].includes(res.status) && attempt < MAX_ATTEMPTS) {
      session = null;   // expired after a few idle minutes; open a new one and try again
      continue;
    }
    if (retry && RETRY_STATUSES.includes(res.status) && attempt < MAX_ATTEMPTS) {
      const after = Number(res.headers.get('retry-after'));
      await sleep(after > 0 ? after * 1000 : 2000 * attempt);
      continue;
    }
    throw new GraphError(res.status, text.slice(0, 300));
  }
}

async function openSession(): Promise<void> {
  for (let attempt = 1; ; attempt++) {
    const res = await send('POST', '/createSession', { persistChanges: true }, false);
    if (res.ok) {
      session = ((await res.json()) as { id: string }).id;
      return;
    }
    if (!RETRY_STATUSES.includes(res.status) || attempt >= MAX_ATTEMPTS) {
      throw new GraphError(res.status, (await res.text()).slice(0, 300));
    }
    await sleep(3000 * attempt);
  }
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
