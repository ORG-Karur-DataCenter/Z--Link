/**
 * Port of zotprep/zotero/client.py — build Zotero items and create/reuse them.
 *
 * Metadata quality matters more than it looks: the Scannable Cite marker only
 * carries author+year, and Zotero renders the real citation from the stored
 * item. So a match that is *correct* but stored with initials-only creators
 * still produces a bad bibliography. Where a DOI exists the canonical Crossref
 * record is re-fetched and the item built from that, rather than from whichever
 * search result happened to win.
 *
 * pyzotero is replaced by direct calls to api.zotero.org, which sends
 * `Access-Control-Allow-Origin: *`. The API key is sent only to that host, in
 * the `Zotero-API-Key` header rather than in a query string, so it never lands
 * in a URL that could be logged or cached.
 */
import { Candidate } from './models.js';
import { normDoi, normText } from './utils.js';
import * as crossref from './search/crossref.js';

const API = 'https://api.zotero.org';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * The per-attempt deadline, plus the run's own cancel signal where the browser
 * can combine them. Without AbortSignal.any the deadline still applies; Cancel
 * then takes effect between attempts rather than during one.
 */
function deadline(ms, signal) {
  const timer = AbortSignal.timeout(ms);
  if (!signal) return timer;
  if (typeof AbortSignal.any === 'function') return AbortSignal.any([signal, timer]);
  return timer;
}

/**
 * A Zotero write token: 32 hex characters identifying one logical write.
 *
 * Zotero ignores a repeat of a write it has already seen under the same token,
 * so a POST whose reply was lost can be sent again without the chunk landing in
 * the library twice.
 */
function writeToken() {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  return [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
}

/** Full names where the provider record has them; surnames otherwise. */
function creatorsFromRaw(cand) {
  const raw = (cand.raw && typeof cand.raw === 'object' && !Array.isArray(cand.raw)) ? cand.raw : {};

  // Crossref: given/family
  if (raw.author) {
    const out = [];
    for (const a of raw.author) {
      if (a.family) {
        out.push({ creatorType: 'author', firstName: a.given || '', lastName: a.family });
      } else if (a.name) {
        out.push({ creatorType: 'author', name: a.name });
      }
    }
    if (out.length) return out;
  }

  // Europe PMC: authorList.author[].firstName/lastName/collectiveName
  const epmc = (raw.authorList || {}).author;
  if (epmc) {
    const out = [];
    for (const a of epmc) {
      if (a.collectiveName) {
        out.push({ creatorType: 'author', name: a.collectiveName });
      } else if (a.lastName) {
        out.push({
          creatorType: 'author',
          firstName: a.firstName || a.initials || '',
          lastName: a.lastName,
        });
      }
    }
    if (out.length) return out;
  }

  // PubMed esummary: "Murray CJL" / CollectiveName
  if (raw.authors) {
    const out = [];
    for (const a of raw.authors) {
      const nm = a.name || '';
      if (!nm) continue;
      if (a.authtype === 'CollectiveName') {
        out.push({ creatorType: 'author', name: nm });
        continue;
      }
      const idx = nm.lastIndexOf(' ');
      const parts = idx < 0 ? [nm] : [nm.slice(0, idx), nm.slice(idx + 1)];
      if (parts.length === 2 && parts[1] === parts[1].toUpperCase()
          && /[A-Z]/.test(parts[1]) && parts[1].length <= 4) {
        out.push({ creatorType: 'author', firstName: parts[1], lastName: parts[0] });
      } else {
        out.push({ creatorType: 'author', name: nm });
      }
    }
    if (out.length) return out;
  }

  // Fall back to whatever surnames the scorer used.
  const out = (cand.authors || []).filter(Boolean)
    .map((s) => ({ creatorType: 'author', firstName: '', lastName: s }));
  if (cand.corporate) out.unshift({ creatorType: 'author', name: cand.corporate });
  return out;
}

/**
 * Zotero item object. Reference-derived fields win for books, since for those
 * the manuscript is the authoritative source.
 */
export function toItem(cand, ref) {
  const item = {
    itemType: cand.item_type || 'journalArticle',
    title: cand.title || ref.title,
    creators: creatorsFromRaw(cand),
    date: String(cand.year || ref.year || ''),
  };
  if (cand.doi) item.DOI = normDoi(cand.doi) || '';
  const extra = [];
  if (cand.pmid) extra.push(`PMID: ${cand.pmid}`);
  if (extra.length) item.extra = extra.join('\n');

  if (item.itemType === 'journalArticle') {
    item.publicationTitle = cand.journal || ref.journal;
    item.journalAbbreviation = cand.journal_abbrev || ref.journal;
    item.volume = cand.volume || ref.volume || '';
    item.issue = cand.issue || ref.issue || '';
    const fp = cand.first_page || ref.first_page;
    const lp = cand.last_page || ref.last_page;
    item.pages = (fp && lp) ? `${fp}-${lp}` : (fp || '');
  } else if (['book', 'bookSection'].includes(item.itemType)) {
    item.publisher = cand.publisher || ref.publisher || '';
    item.place = cand.place || ref.place || '';
    if (ref.edition) item.edition = ref.edition;
    if (item.itemType === 'bookSection') {
      // Zotero needs the containing volume's title, distinct from the chapter
      // title already in item.title.
      item.bookTitle = cand.book_title || ref.book_title;
      const fp = cand.first_page || ref.first_page;
      const lp = cand.last_page || ref.last_page;
      item.pages = (fp && lp) ? `${fp}-${lp}` : (fp || '');
    }
  }
  return item;
}

/** Replace a search hit with the canonical Crossref record for its DOI. */
export async function enrich(cand, mailto, opts = {}) {
  const raw = cand.raw;
  if (!cand.doi || (raw && typeof raw === 'object' && raw.author)) return cand;
  const better = await crossref.byDoi(normDoi(cand.doi), mailto, opts);
  if (better && better.title) {
    better.providers = new Set([...cand.providers, 'crossref:canonical']);
    // keep the item type we decided on; Crossref's is occasionally wrong
    if (['book', 'bookSection'].includes(cand.item_type)) better.item_type = cand.item_type;
    return better;
  }
  return cand;
}

/**
 * Reuses existing items instead of duplicating them.
 *
 * There is no dry-run mode: every run adds its resolved references to the
 * library. Two things carry the safety that a preview pass used to:
 * `existingKey()` matches on DOI and on normalised title+year, so re-running a
 * document reuses items rather than creating second copies; and only references
 * that actually passed the accept gate are ever built into items, so an
 * unresolved reference stays a `{NEEDS REVIEW}` marker rather than becoming a
 * guess in someone's library.
 */
export class ZoteroWriter {
  constructor(userid, apiKey, { onLog = null, signal = null } = {}) {
    this.userid = userid;
    this.apiKey = apiKey;
    this.byDoi = new Map();
    this.byTitle = new Map();
    this.log = onLog || (() => {});
    this.signal = signal;
  }

  get headers() {
    return {
      'Zotero-API-Key': this.apiKey,
      'Zotero-API-Version': '3',
      'Content-Type': 'application/json',
    };
  }

  /**
   * One request to api.zotero.org, retried on the failures that pass.
   *
   * Every call here lands at the end of a run that has already spent minutes
   * resolving references, and on a link the provider searches have usually
   * shown can drop requests. A single `AbortSignal.timeout` with nothing behind
   * it therefore throws the whole run away for one slow response — and throws
   * it away as `signal timed out`, which is only Chrome's wording for that
   * DOMException and says nothing about which step stopped or whether anything
   * reached the library.
   *
   * So: a slow or dropped response is retried with growing backoff, Zotero's
   * own `Backoff` and `Retry-After` headers are obeyed rather than ignored, and
   * what finally escapes names the step it came from.
   */
  async request(url, {
    what, method = 'GET', body = null, headers = null, timeout = 30000, attempts = 4,
  }) {
    let delay = 2000;
    let reason = '';

    for (let attempt = 1; attempt <= attempts; attempt++) {
      let res = null;
      try {
        res = await fetch(url, {
          method,
          headers: { ...this.headers, ...(headers || {}) },
          body,
          signal: deadline(timeout, this.signal),
        });
      } catch (e) {
        if (this.signal?.aborted) throw new DOMException('Cancelled', 'AbortError');
        reason = e?.name === 'TimeoutError'
          ? `no reply within ${Math.round(timeout / 1000)}s`
          : `the connection failed (${e?.message || e})`;
      }

      if (res) {
        // 429 and 5xx are Zotero asking for time, not refusing the work.
        const retryable = res.status === 429 || res.status >= 500;
        if (!retryable) {
          const backoff = Number(res.headers.get('Backoff') || 0);
          if (backoff > 0) await sleep(Math.min(backoff, 30) * 1000);
          return res;
        }
        reason = `Zotero answered HTTP ${res.status}`;
        const after = Number(res.headers.get('Retry-After') || res.headers.get('Backoff') || 0);
        if (after > 0) delay = Math.min(after, 60) * 1000;
      }

      if (attempt === attempts) break;
      this.log('warn', `Zotero: ${reason} while ${what} — retrying in `
        + `${Math.round(delay / 1000)}s (attempt ${attempt + 1} of ${attempts}).`);
      await sleep(delay + Math.random() * 400);
      delay = Math.min(delay * 2, 30000);
      if (this.signal?.aborted) throw new DOMException('Cancelled', 'AbortError');
    }

    throw new Error(`Zotero stopped responding while ${what} — ${reason}, after `
      + `${attempts} attempts. This is the connection to api.zotero.org rather than anything `
      + 'wrong with your document or your key; your references are still resolved, so press '
      + 'the button again and the run will reuse what is already in your library.');
  }

  /**
   * Verify the credentials, and specifically the *write* permission, before
   * anything is resolved.
   *
   * This asks `/keys/current` rather than reading the library, because a read
   * request cannot tell the two most common misconfigurations apart:
   *
   *   - a key created without "Allow write access". It reads the library
   *     perfectly, passes any GET-based check, and then fails at the very end
   *     of a run when the items are posted.
   *   - a userID that belongs to a different account than the key. Reads may
   *     still succeed against a public library, and the write lands nowhere the
   *     reader is looking.
   *
   * `/keys/current` returns the key's owning userID and its exact permissions,
   * so both are caught up front with a message that says which one it is.
   */
  async verify() {
    const res = await this.request(`${API}/keys/current`, {
      what: 'checking your API key', timeout: 20000, attempts: 3,
    });

    if (res.status === 403 || res.status === 401) {
      throw new Error('Zotero does not recognise that API key. Copy it again from '
        + 'zotero.org/settings/keys — the key is shown only once when created, so if you '
        + 'no longer have it, delete that key and make a new one.');
    }
    if (!res.ok) throw new Error(`Zotero API error ${res.status} while checking the key.`);

    const info = await res.json();
    const user = (info.access || {}).user || {};

    if (String(info.userID) !== String(this.userid)) {
      throw new Error(`That API key belongs to Zotero user ${info.userID}`
        + `${info.username ? ` (${info.username})` : ''}, but the userID entered is `
        + `${this.userid}. Use ${info.userID} — it is the number shown on `
        + 'zotero.org/settings/keys, above the key list.');
    }
    if (!user.library) {
      throw new Error('That API key has no access to your personal library. Edit it at '
        + 'zotero.org/settings/keys and tick "Allow library access".');
    }
    if (!user.write) {
      throw new Error('That API key is read-only, so nothing can be added to your library. '
        + 'Edit it at zotero.org/settings/keys, tick "Allow write access" under Personal '
        + 'Library, and save. You can reuse the same key — no need to create a new one.');
    }

    this.log('info', `Zotero key verified for ${info.username || 'user ' + info.userID}`
      + ' — library write access confirmed.');
    return true;
  }

  /** Index the existing library once, so duplicate checks are local. */
  async loadLibrary(onProgress = null) {
    const LIMIT = 100;
    this.byDoi.clear();
    this.byTitle.clear();
    let start = 0;
    for (;;) {
      const res = await this.request(
        `${API}/users/${encodeURIComponent(this.userid)}/items?start=${start}&limit=${LIMIT}`,
        { what: 'indexing your library' },
      );
      if (!res.ok) throw new Error(`Zotero API error ${res.status} while reading the library.`);
      const batch = await res.json();
      if (!batch.length) break;
      for (const it of batch) {
        const data = it.data || {};
        const key = data.key;
        const doi = normDoi(data.DOI);
        if (doi && !this.byDoi.has(doi)) this.byDoi.set(doi, key);
        const title = normText(data.title || '');
        if (title) {
          const tk = `${title}|${(data.date || '').slice(0, 4)}`;
          if (!this.byTitle.has(tk)) this.byTitle.set(tk, key);
        }
      }
      start += batch.length;
      if (onProgress) onProgress(start);
      // A short page is the last page. Asking for the empty one after it is a
      // whole extra round trip to api.zotero.org that can only fail.
      if (batch.length < LIMIT) break;
    }
  }

  existingKey(item) {
    const doi = normDoi(item.DOI);
    if (doi && this.byDoi.has(doi)) return this.byDoi.get(doi);
    const tk = `${normText(item.title || '')}|${(item.date || '').slice(0, 4)}`;
    return this.byTitle.get(tk) ?? null;
  }

  /**
   * Returns { keys, created, reused, failed }.
   *
   * The counts are separated because "8 items ready" reads identically whether
   * eight were created or eight already existed — and a reader who cannot see
   * the difference has no way to tell a working run from one that quietly added
   * nothing.
   */
  async create(itemsByN, onProgress = null) {
    const keys = new Map();
    const pending = [];
    let reused = 0;
    let created = 0;
    const failed = [];

    for (const [n, item] of itemsByN) {
      const hit = this.existingKey(item);
      if (hit) { keys.set(n, hit); reused++; } else pending.push([n, item]);
    }

    // 50 items is Zotero's maximum for a single write.
    const CHUNK = 50;
    for (let start = 0; start < pending.length; start += CHUNK) {
      const chunk = pending.slice(start, start + CHUNK);
      const last = Math.min(start + CHUNK, pending.length);
      // Generated once per chunk, so the retries inside request() are the same
      // write rather than a second one.
      const token = writeToken();
      const res = await this.request(`${API}/users/${encodeURIComponent(this.userid)}/items`, {
        what: `adding items ${start + 1}-${last} of ${pending.length} to your library`,
        method: 'POST',
        headers: { 'Zotero-Write-Token': token },
        body: JSON.stringify(chunk.map(([, it]) => it)),
        timeout: 90000,
      });

      // Zotero rejects a write token it has already seen. That means the reply
      // to an earlier attempt was lost rather than the write itself: the items
      // are in the library, and only their keys are missing. Re-index and look
      // them up, instead of sending them again and leaving duplicates behind.
      if (res.status === 412) {
        this.log('warn', 'The reply to that write was lost, but Zotero had already stored the '
          + 'items — re-reading the library to find them rather than adding them twice.');
        await this.loadLibrary();
        for (const [n, item] of chunk) {
          const hit = this.existingKey(item);
          if (hit) { keys.set(n, hit); created++; } else {
            failed.push([n, 'stored by Zotero, but the item could not be found again']);
            this.log('warn', `Reference ${n} could not be matched back to a library item.`);
          }
        }
        if (onProgress) onProgress(last, pending.length);
        continue;
      }

      if (!res.ok) {
        const detail = (await res.text()).slice(0, 300);
        if (res.status === 403) {
          throw new Error('Zotero refused the write (403). The API key can read your '
            + 'library but not add to it — edit it at zotero.org/settings/keys and tick '
            + '"Allow write access".');
        }
        throw new Error(`Zotero rejected the write: HTTP ${res.status} ${detail}`);
      }
      const body = await res.json();
      for (const [idx, key] of Object.entries(body.success || {})) {
        keys.set(chunk[Number(idx)][0], key);
        created++;
      }
      // Zotero reports per-item problems here rather than failing the request,
      // so without this a rejected item silently never appears in the library.
      for (const [idx, info] of Object.entries(body.failed || {})) {
        const n = chunk[Number(idx)][0];
        const msg = info && info.message ? info.message : JSON.stringify(info);
        failed.push([n, msg]);
        this.log('warn', `Reference ${n} was rejected by Zotero: ${msg}`);
      }
      if (onProgress) onProgress(last, pending.length);
    }
    return { keys, created, reused, failed };
  }
}

/** Rebuild a Candidate from a plain object (used after enrich round-trips). */
export function asCandidate(obj) {
  return obj instanceof Candidate ? obj : new Candidate(obj);
}
