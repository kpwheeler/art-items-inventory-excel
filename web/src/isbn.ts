// Book details from Open Library (free, no key): https://openlibrary.org/dev/docs/api/books

export interface BookInfo {
  isbn: string;
  title?: string;
  authors?: string;
  publisher?: string;
  year?: string;
  pages?: number;
}

interface OpenLibraryBook {
  title?: string;
  authors?: { name: string }[];
  publishers?: { name: string }[];
  publish_date?: string;
  number_of_pages?: number;
}

export function cleanIsbn(raw: string): string {
  return raw.replace(/[^0-9Xx]/g, '').toUpperCase();
}

export function isIsbn(isbn: string): boolean {
  return /^\d{9}[\dX]$/.test(isbn) || /^97[89]\d{10}$/.test(isbn);
}

/** Returns null when the ISBN is malformed or Open Library has no record of it. */
export async function lookupIsbn(raw: string): Promise<BookInfo | null> {
  const isbn = cleanIsbn(raw);
  if (!isIsbn(isbn)) return null;
  const url = `https://openlibrary.org/api/books?bibkeys=ISBN:${isbn}&format=json&jscmd=data`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Open Library error ${res.status}`);
  const data = (await res.json()) as Record<string, OpenLibraryBook | undefined>;
  const book = data[`ISBN:${isbn}`];
  if (!book) return null;
  const names = (xs?: { name: string }[]) => xs?.map((x) => x.name).filter(Boolean).join(', ') || undefined;
  return {
    isbn,
    title: book.title,
    authors: names(book.authors),
    publisher: names(book.publishers),
    year: book.publish_date?.match(/\d{4}/)?.[0] ?? book.publish_date,
    pages: book.number_of_pages,
  };
}

/** Field values to offer for a Books row, keyed by workbook column name. */
export function bookFields(info: BookInfo): Record<string, string | number | undefined> {
  return {
    Name: info.title,
    Creator: info.authors,
    Author: info.authors,
    Publisher: info.publisher,
    'Publication Year': info.year,
    Pages: info.pages,
  };
}
