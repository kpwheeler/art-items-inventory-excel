import './style.css';
import { initAuth, signIn, signOut } from './auth';
import { closeSession, GraphError, setWorkbookPath } from './graph';
import { bookFields, lookupIsbn } from './isbn';
import {
  AUTO_COLUMNS, ChangedError, DATE_COLUMNS, LONG_COLUMNS, NUMBER_COLUMNS, Workbook,
  type Category, type Fields, type Item,
} from './workbook';

interface Config {
  clientId: string;
  authority: string;
  workbookPath: string;
}

const main = document.getElementById('main') as HTMLElement;
const accountEl = document.getElementById('account') as HTMLElement;
const nav = document.getElementById('nav') as HTMLElement;
const scanner = document.getElementById('scanner') as HTMLDialogElement;
const scanVideo = document.getElementById('scan-video') as HTMLVideoElement;
const scanCancel = document.getElementById('scan-cancel') as HTMLButtonElement;

const wb = new Workbook();

// Fields kept by "Save & add another".
const KEEP_FIELDS = ['Acquired From Type', 'Acquired From Name', 'Acquired Date', 'Location'];

type Child = Node | string | null | false | undefined;

function h<K extends keyof HTMLElementTagNameMap>(
  tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Child[]
): HTMLElementTagNameMap[K] {
  const el = Object.assign(document.createElement(tag), props);
  el.append(...children.filter((c): c is Node | string => c !== null && c !== false && c !== undefined));
  return el;
}

function show(...children: Child[]): void {
  main.replaceChildren(...children.filter((c): c is Node | string => !!c));
  window.scrollTo(0, 0);
}

function message(text: string, kind: 'ok' | 'error' | 'info' = 'info'): HTMLElement {
  return h('p', { className: `msg ${kind}` }, text);
}

function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

// ---------------------------------------------------------------- the item form

interface ItemForm {
  element: HTMLFormElement;
  fields(): Fields;
}

function buildForm(
  c: Category, values: Fields, onSubmit: (action: string) => Promise<void>, actions: [string, string][],
  lead?: HTMLElement,
): ItemForm {
  const headers = wb.table(c).headers.filter((col) => !AUTO_COLUMNS.includes(col));
  const inputs = new Map<string, HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>();
  const status = h('span', { className: 'muted' });

  const control = (col: string) => {
    const value = values[col] ?? '';
    const options = wb.dropdowns.get(col);
    if (col === 'Item No') {
      return h('input', { name: col, value: value || 'Assigned when saved', readOnly: true });
    }
    if (options) {
      const select = h('select', { name: col }, h('option', { value: '' }, ''),
        ...options.map((o) => h('option', { value: o, selected: o === value }, o)));
      if (value && !options.includes(value)) select.append(h('option', { value, selected: true }, value));
      return select;
    }
    if (LONG_COLUMNS.includes(col)) return h('textarea', { name: col, value, rows: 3 });
    if (DATE_COLUMNS.includes(col)) return h('input', { name: col, type: 'date', value });
    if (NUMBER_COLUMNS.includes(col)) return h('input', { name: col, type: 'number', step: 'any', inputMode: 'decimal', value });
    return h('input', { name: col, value });
  };

  const fillBook = async (isbn: string) => {
    status.textContent = 'Looking up…';
    try {
      const info = await lookupIsbn(isbn);
      const lookupStatus = inputs.get('Lookup Status');
      if (!info) {
        status.textContent = 'Not found — enter details by hand.';
        if (lookupStatus) lookupStatus.value = 'Not found';
        return;
      }
      let filled = 0;
      for (const [col, v] of Object.entries(bookFields(info))) {
        const input = inputs.get(col);
        if (input && !input.value && v !== undefined && v !== '') { input.value = String(v); filled++; }
      }
      if (lookupStatus) lookupStatus.value = 'Found';
      status.textContent = `Found "${info.title ?? isbn}" — filled ${filled} empty field(s).`;
    } catch (e) {
      status.textContent = `Lookup failed: ${errorText(e)}`;
    }
  };

  const rows = headers.map((col) => {
    const input = control(col);
    inputs.set(col, input);
    let extra: Child = null;
    if (col === 'ISBN') {
      const isbnInput = input as HTMLInputElement;
      extra = h('div', { className: 'inline' },
        h('button', { type: 'button', onclick: () => void fillBook(isbnInput.value) }, 'Look up'),
        h('button', { type: 'button', onclick: () => void scan(isbnInput, fillBook) }, 'Scan'),
        status);
    }
    return h('label', {}, h('span', {}, col), input, extra);
  });

  const units = inputs.get('Units');
  if (units && !values['Item No'] && !units.value) units.value = 'in';

  const form = h('form', { className: 'item' }, lead, ...rows,
    h('div', { className: 'actions' },
      ...actions.map(([action, label], i) =>
        h('button', { type: 'submit', value: action, className: i === 0 ? 'primary' : '' }, label))));
  let saving = false;
  form.addEventListener('submit', (ev) => {
    ev.preventDefault();
    if (saving) return;   // one save at a time, however often the button is pressed
    saving = true;
    const submitter = (ev as SubmitEvent).submitter as HTMLButtonElement | null;
    const buttons = form.querySelectorAll<HTMLButtonElement>('.actions button');
    const label = submitter?.textContent ?? '';
    buttons.forEach((b) => { b.disabled = true; });
    if (submitter) submitter.textContent = 'Saving…';
    void onSubmit(submitter?.value ?? actions[0][0]).finally(() => {
      saving = false;
      buttons.forEach((b) => { b.disabled = false; });
      if (submitter) submitter.textContent = label;
    });
  });

  return {
    element: form,
    fields: () => {
      const out: Fields = {};
      for (const [col, input] of inputs) if (col !== 'Item No') out[col] = input.value;
      return out;
    },
  };
}

async function scan(target: HTMLInputElement, then: (isbn: string) => Promise<void>): Promise<void> {
  scanner.showModal();
  try {
    // Loaded on first use: the barcode library is most of the bundle.
    const { startIsbnScan } = await import('./scan');
    const controls = await startIsbnScan(scanVideo, (isbn) => {
      scanner.close();
      target.value = isbn;
      void then(isbn);
    });
    scanCancel.onclick = () => { controls.stop(); scanner.close(); };
    scanner.onclose = () => controls.stop();
  } catch (e) {
    scanner.close();
    alert(`Camera unavailable: ${errorText(e)}`);
  }
}

// ---------------------------------------------------------------- views

function categoryPicker(selected: Category | undefined, onChange: (c: Category) => void, allLabel?: string): HTMLSelectElement {
  const select = h('select', {},
    allLabel ? h('option', { value: '' }, allLabel) : null,
    ...wb.categories.map((c) => h('option', { value: c.name, selected: c === selected }, c.name)));
  select.onchange = () => {
    const c = wb.category(select.value);
    if (c || allLabel) onChange(c as Category);
  };
  return select;
}

/** The Category pull-down shown as the first field of the item form. */
function categoryField(selected: Category, onChange: (c: Category, select: HTMLSelectElement) => void): HTMLElement {
  const select = h('select', { name: 'Category' },
    ...wb.categories.map((c) => h('option', { value: c.name, selected: c === selected }, c.name)));
  select.onchange = () => {
    const c = wb.category(select.value);
    if (c) onChange(c, select);
  };
  return h('label', { className: 'category' }, h('span', {}, 'Category'), select);
}

function showAdd(c: Category = wb.categories[0], keep: Fields = {}, note?: HTMLElement): void {
  // Switching category keeps what was already typed in the fields the categories share.
  const lead = categoryField(c, (next) => showAdd(next, form.fields()));
  const form = buildForm(c, keep, (action) => save(action), [
    ['save', 'Save'], ['again', 'Save & add another'],
  ], lead);
  const save = async (action: string) => {
    const fields = form.fields();
    if (Object.entries(fields).every(([col, v]) => !v || col === 'Units')) {
      show(note ?? null, message('Fill in at least one field before saving.', 'error'), header, form.element);
      return;
    }
    try {
      const itemNo = await wb.addItem(c, fields);
      const done = message(`Saved ${itemNo} on the ${c.tab} tab.`, 'ok');
      if (action === 'again') {
        const kept: Fields = {};
        for (const col of KEEP_FIELDS) if (fields[col]) kept[col] = fields[col];
        showAdd(c, kept, done);
      } else {
        const item = wb.find(itemNo);
        if (item) showEdit(item, done); else showAdd(c, {}, done);
      }
    } catch (e) {
      show(message(`Not saved: ${errorText(e)}`, 'error'), header, form.element);
    }
  };
  const header = h('h2', {}, 'Add item');
  show(note ?? null, header, form.element);
}

function showEdit(item: Item, note?: HTMLElement): void {
  const t = wb.table(item.category);
  const values: Fields = {};
  for (const col of t.headers) values[col] = wb.value(item, col);

  // Choosing another category moves the item to that tab (after confirming).
  const lead = categoryField(item.category, (dest, select) => void move(dest, select));
  const form = buildForm(item.category, values, (action) => save(action, false), [['save', 'Save changes']], lead);
  const move = async (dest: Category, select: HTMLSelectElement) => {
    if (!confirm(`Move ${values['Item No']} to ${dest.name}? It gets a new ${dest.prefix} item number; `
      + 'unsaved changes on this page are not kept.')) {
      select.value = item.category.name;
      return;
    }
    select.disabled = true;
    try {
      const itemNo = await wb.moveItem(item, dest);
      const moved = wb.find(itemNo);
      const done = message(`Moved to ${dest.name} as ${itemNo}. Fill in its ${dest.name} fields below.`, 'ok');
      if (moved) showEdit(moved, done); else showSearch();
    } catch (e) {
      show(message(`Not moved: ${errorText(e)}`, 'error'), header, form.element);
    }
  };
  const save = async (_action: string, force: boolean) => {
    try {
      await wb.updateItem(item, form.fields(), force);
      const updated = wb.find(values['Item No']);
      if (updated) showEdit(updated, message('Changes saved.', 'ok'));
    } catch (e) {
      if (e instanceof ChangedError && confirm(`${e.message} Save your version anyway?`)) {
        await save(_action, true);
        return;
      }
      show(message(`Not saved: ${errorText(e)}`, 'error'), header, form.element);
    }
  };

  const header = h('div', { className: 'bar' },
    h('h2', {}, `${values['Item No'] || 'Item'} · ${item.category.name}`),
    h('span', { className: 'muted' }, values['Last Updated'] ? `Updated ${values['Last Updated']}` : ''));
  show(note ?? null, header, form.element);
}

function showSearch(text = '', c?: Category): void {
  const box = h('input', { type: 'search', placeholder: 'Title, creator, location, tag…', value: text });
  const picker = categoryPicker(c, (next) => showSearch(box.value, next), 'All categories');
  const list = h('ul', { className: 'results' });

  const render = () => {
    const items = wb.search(box.value, wb.category(picker.value));
    const total = items.reduce((sum, it) => sum + (Number(wb.value(it, 'Estimated Value')) || 0), 0);
    list.replaceChildren(
      h('li', { className: 'muted' }, `${items.length} item(s) · estimated $${total.toLocaleString()}`),
      ...items.map((it) => h('li', {},
        h('button', { className: 'result', onclick: () => showEdit(it) },
          h('strong', {}, `${wb.value(it, 'Name') || '(no name)'}`),
          h('span', { className: 'muted' },
            [wb.value(it, 'Item No'), it.category.name, wb.value(it, 'Creator'), wb.value(it, 'Location')]
              .filter(Boolean).join(' · '))))));
  };
  box.oninput = render;
  render();
  show(h('div', { className: 'bar' }, h('h2', {}, 'Search'), picker), box, list);
  box.focus();
}

function showTidy(): void {
  const log = h('ul', { className: 'log' });
  const run = h('button', {
    className: 'primary',
    onclick: async () => {
      run.disabled = true;
      log.replaceChildren();
      try {
        await wb.tidy((line) => log.append(h('li', {}, line)));
        log.append(h('li', { className: 'ok' }, 'Done.'));
      } catch (e) {
        log.append(h('li', { className: 'error' }, errorText(e)));
      }
      run.disabled = false;
    },
  }, 'Tidy workbook');
  show(h('h2', {}, 'Tidy workbook'),
    h('p', {}, 'Numbers rows typed in Excel without an Item No, and fills blank book details from their ISBN. Your own entries are never overwritten.'),
    run, log);
}

async function closeApp(): Promise<void> {
  if (main.querySelector('form.item') && !confirm('Close the inventory? Anything not saved on this page is lost.')) return;
  nav.hidden = true;
  show(message('Closing…'));
  try {
    await closeSession();
  } catch {
    // Nothing to lose: every save is already in the workbook.
  }
  show(message('Finished. Everything you saved is in the workbook. You can close this tab.', 'ok'),
    h('button', { className: 'primary', onclick: () => window.location.reload() }, 'Open again'));
  window.close();   // works only where Chrome allows a page to close itself
}

// ---------------------------------------------------------------- start-up

async function start(): Promise<void> {
  let config: Config;
  try {
    const res = await fetch(`${import.meta.env.BASE_URL}config.json`);
    if (!res.ok) throw new Error(String(res.status));
    config = (await res.json()) as Config;
  } catch {
    show(message('Missing config.json. Copy config.example.json to config.json and fill in the app (client) ID and workbook path.', 'error'));
    return;
  }

  setWorkbookPath(config.workbookPath);
  const account = await initAuth(config.clientId, config.authority);
  if (!account) {
    show(h('p', {}, 'Sign in with the Microsoft account whose OneDrive holds the workbook.'),
      h('button', { className: 'primary', onclick: () => void signIn() }, 'Sign in with Microsoft'));
    return;
  }
  accountEl.replaceChildren(h('span', { className: 'muted' }, account.username),
    h('button', { className: 'link', onclick: () => void signOut() }, 'Sign out'));

  show(message('Opening the workbook…'));
  try {
    await wb.load();
  } catch (e) {
    if (e instanceof GraphError && e.status === 404) {
      show(message(`Workbook not found in your OneDrive at "${config.workbookPath}".`, 'error'),
        h('p', {}, 'Upload Collection Inventory.xlsx to that folder with that exact name, then reload this page. '
          + 'If it is somewhere else, set the WORKBOOK_PATH repository variable to its path and redeploy.'));
      return;
    }
    throw e;
  }
  nav.hidden = false;
  nav.querySelectorAll<HTMLButtonElement>('button[data-view]').forEach((b) => {
    b.onclick = () => {
      if (b.dataset.view === 'add') showAdd();
      else if (b.dataset.view === 'search') showSearch();
      else if (b.dataset.view === 'close') void closeApp();
      else showTidy();
    };
  });
  showAdd();
}

start().catch((e) => show(message(`Could not start: ${errorText(e)}`, 'error')));
