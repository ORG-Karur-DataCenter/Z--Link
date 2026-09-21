/**
 * Port of the `run()` body in zotprep/cli.py — the whole job, start to finish.
 *
 * Split into two halves so the review step can sit between them, which is where
 * the CLI puts it too (`--review` runs after resolve_all and before the accept
 * list is computed):
 *
 *     resolve()  document -> parsed references -> Resolutions
 *     [review]   the reader decides the ambiguous ones
 *     finish()   enrich -> Zotero items -> rewritten .docx + report
 *
 * finish() is idempotent: it re-loads the document from the original bytes each
 * time rather than reusing the one resolve() read. That matters because
 * markBody() and removeRange() mutate the document in place, so running finish()
 * a second time on the same object would mark up already-marked text and delete
 * from an already-truncated body. Re-reading costs a few milliseconds and makes
 * "review, then rebuild" safe to repeat as many times as the reader likes.
 *
 * Unlike the CLI, there is no dry run: every run adds its resolved references to
 * the reader's library. Three things do the work the preview pass used to do —
 * credentials are verified before any resolution starts, existing items are
 * matched and reused rather than duplicated, and only references that passed
 * the accept gate become items at all. An unresolved reference stays a
 * `{NEEDS REVIEW}` marker in the document instead of becoming a guess in
 * someone's library.
 */
import {
  Document, biblioEndIndex, countCitations, findBiblioIndex, hasImages, makeRenderer, markBody,
  markBodyFields, parseBibliography, removeRange,
} from './docx.js';
import { parseReference } from './extractor.js';
import { newCache, resolveAll } from './resolver.js';
import { resetProviders, setLogSink } from './search/base.js';
import { ZoteroWriter, enrich, toItem } from './zotero.js';

function csvCell(v) {
  const s = v === null || v === undefined ? '' : String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function toCsv(rows) {
  // Excel opens UTF-8 CSV as the system codepage unless it sees a BOM, which
  // turns every accented author name into mojibake in the report.
  return '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
}

/**
 * Read the document, parse the bibliography, resolve every reference.
 *
 * @returns state to hand to finish(), plus `refs` and `results` for review.
 */
export async function resolve(file, opts, cb = {}) {
  const {
    mailto = 'you@example.com', userid = '', apiKey = '',
    bibliographyText = '', workers = 12, s2Key = '', signal = null,
  } = opts;
  const onStage = cb.onStage || (() => {});
  const onProgress = cb.onProgress || (() => {});
  const onLog = cb.onLog || (() => {});

  if (!(userid && apiKey)) {
    throw new Error('Adding items to Zotero needs a userID and an API key. '
      + 'Get both at zotero.org/settings/keys — the key needs write access.');
  }

  resetProviders();
  setLogSink(onLog);

  onStage('reading the document');
  const buffer = await file.arrayBuffer();
  const doc = await Document.load(buffer);

  let biblioIdx;
  let biblioText;
  if (bibliographyText.trim()) {
    biblioText = bibliographyText;
    biblioIdx = doc.paragraphs.length;
  } else {
    biblioIdx = findBiblioIndex(doc);
    if (biblioIdx === null) {
      throw new Error('No "References" heading found in this document. '
        + 'Z-Link finds the bibliography by looking for a paragraph that reads exactly '
        + '"References" (or "Bibliography", "Works Cited", or "Reference List"). '
        + 'Add one on its own line above your reference list and try again.');
    }
    // Bounded, so that trailing figures and tables are neither read as
    // references nor removed with them.
    const biblioEnd = biblioEndIndex(doc, biblioIdx);
    biblioText = doc.paragraphs.slice(biblioIdx + 1, biblioEnd).map((p) => p.text).join('\n');
  }

  const rawRefs = parseBibliography(biblioText);
  if (!rawRefs.size) {
    throw new Error('Found the bibliography but could not read any entries from it.');
  }

  // A bibliography with nothing citing it means the in-text notation was not
  // recognised. Stop here rather than after several minutes of resolution: the
  // run would end by removing the reference list and linking nothing.
  const nFound = countCitations(doc, biblioIdx);
  if (!nFound) {
    throw new Error(`Read ${rawRefs.size} references from the bibliography, but found no `
      + 'citations in the body of the document. Z-Link recognises [1] and (2,3), '
      + 'superscript ¹², and plain digits such as "…knee OA 2." or "…analgesia 6,7." '
      + 'Nothing was changed and nothing was added to your library.');
  }

  const refs = new Map();
  for (const [n, t] of rawRefs) refs.set(n, parseReference(n, t));
  onLog('info', `Parsed ${refs.size} references, ${nFound} in-text citations.`);

  // Both checks happen before any resolution work: a wrong key is worth knowing
  // in two seconds rather than after several minutes of searching. The document
  // is read first only because reading it is local and instant.
  onStage('verifying Zotero credentials');
  await new ZoteroWriter(userid, apiKey, { onLog, signal }).verify();

  onStage(`resolving ${refs.size} references`);
  const cache = newCache();
  const results = await resolveAll(refs, mailto, cache, {
    workers, signal, s2Key, onProgress: (_r, done, total) => onProgress(done, total),
  });
  if (signal?.aborted) throw new DOMException('Cancelled', 'AbortError');

  return { buffer, refs, results, usedPastedBibliography: Boolean(bibliographyText.trim()) };
}

/**
 * Turn Resolutions into a rewritten .docx and a report. Safe to call repeatedly
 * on the same state — see the note at the top of this file.
 */
export async function finish(state, opts, cb = {}) {
  const {
    mailto = 'you@example.com', userid = '', apiKey = '', signal = null,
  } = opts;
  const onStage = cb.onStage || (() => {});
  const onProgress = cb.onProgress || (() => {});
  const onLog = cb.onLog || (() => {});
  const { buffer, refs, results, usedPastedBibliography } = state;

  setLogSink(onLog);

  const ordered = [...results.keys()].sort((a, b) => a - b);
  const accepted = ordered.filter((n) => ['ACCEPTED', 'FROM_TEXT'].includes(results.get(n).status));
  // A reference deleted at review is decided, not outstanding — counting it as
  // unresolved would keep reporting a problem the reader has already settled.
  const dropped = ordered.filter((n) => results.get(n).status === 'DROPPED');
  const unresolved = ordered.filter((n) => !accepted.includes(n) && !dropped.includes(n));

  // upgrade winning candidates to canonical Crossref records for item quality
  onStage('fetching canonical metadata');
  let done = 0;
  for (const n of accepted) {
    const res = results.get(n);
    if (res.candidate) res.candidate = await enrich(res.candidate, mailto, signal ? { signal } : {});
    onProgress(++done, accepted.length);
  }

  const writer = new ZoteroWriter(userid, apiKey, { onLog, signal });
  onStage('indexing your Zotero library for duplicates');
  await writer.loadLibrary((n) => onLog('info', `  indexed ${n} items`));

  onStage('adding items to Zotero');
  const items = accepted
    .filter((n) => results.get(n).candidate)
    .map((n) => [n, toItem(results.get(n).candidate, refs.get(n))]);
  const { keys, created, reused, failed } = await writer.create(items, (a, b) => onProgress(a, b));
  for (const [n, k] of keys) results.get(n).zotero_key = k;

  onLog('info', created
    ? `Added ${created} new item${created === 1 ? '' : 's'} to your Zotero library`
      + `${reused ? `, reused ${reused} already there` : ''}.`
    : (reused
      ? `Nothing new to add — all ${reused} items were already in your library.`
      : 'No items were added.'));
  if (failed.length) {
    onLog('warn', `${failed.length} item${failed.length === 1 ? '' : 's'} rejected by Zotero `
      + `(references ${failed.map(([n]) => n).join(', ')}) — see above for each reason.`);
  }
  if (created) {
    onLog('info', 'If Zotero on your desktop does not show them yet, press its Sync button — '
      + 'items added through the API arrive in the web library first.');
  }

  // Two documents, from the same resolutions. `fields` is the one to use: Word
  // field codes Zotero reads directly, finished on download. `scannable` is the
  // old Scannable Cite output, kept for anyone who needs the ODF Scan route.
  //
  // Each is built from a fresh read of the original bytes, because markBody and
  // removeRange mutate the document in place — see the note at the top of this
  // file.
  onStage('rewriting citations');
  const warnings = [];
  const built = {};
  let nMarked = 0;
  for (const style of ['fields', 'scannable']) {
    const doc = await Document.load(buffer);
    const biblioIdx = usedPastedBibliography ? doc.paragraphs.length : findBiblioIndex(doc);
    // Only one pass may collect warnings, or every one is reported twice.
    // The item records travel with the renderer so each field can embed the
    // reference itself, not just a link into this library — a document whose
    // citations only resolve on the machine that made them is not much use to a
    // co-author.
    const render = makeRenderer(results, keys, userid || '0',
      { refs, warnings: style === 'fields' ? warnings : null, style, items: new Map(items) });
    const n = style === 'fields'
      ? markBodyFields(doc, biblioIdx, render)
      : markBody(doc, biblioIdx, render);
    // The reference list goes only when something has taken its place. Removing
    // it after marking nothing would hand back a manuscript with no citations
    // and no bibliography either.
    if (!usedPastedBibliography && n) {
      removeRange(doc, biblioIdx, biblioEndIndex(doc, biblioIdx));
    } else if (!usedPastedBibliography && style === 'fields') {
      warnings.push('No citations were rewritten, so the bibliography has been left in place.');
    }
    built[style] = await doc.save();
    if (style === 'fields') {
      nMarked = n;
      if (hasImages(doc)) {
        warnings.push('This document contains images. The ODF Scan copy below is the one '
          + 'affected: that plugin can write a citation into a picture\'s XML and produce a '
          + 'file Word refuses to open. The Word copy needs no conversion and is unaffected.');
      }
    }
  }
  const docxBlob = built.fields;

  const rows = [[
    'n', 'status', 'tier', 'confidence', 'doi', 'pmid', 'zotero_key',
    'resolved_title', 'reason', 'advisory', 'raw_reference',
  ]];
  for (const n of ordered) {
    const r = results.get(n);
    const c = r.candidate;
    rows.push([n, r.status, r.tier, r.confidence, c ? c.doi : '', c ? c.pmid : '',
      r.zotero_key || '', c ? c.title : '', r.reason, (r.notes || []).join(' | '),
      refs.get(n).raw]);
  }

  const advisories = [];
  for (const n of ordered) {
    for (const note of results.get(n).notes || []) advisories.push([n, note]);
  }

  return {
    docxBlob,
    scannableBlob: built.scannable,
    reportCsv: new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' }),
    results,
    refs,
    keys,
    stats: {
      total: refs.size,
      accepted: accepted.length,
      unresolved,
      dropped,
      nMarked,
      created,
      reused,
      failed,
      inLibrary: keys.size,
    },
    warnings: [...new Set(warnings)],
    advisories,
  };
}

/** resolve() then finish(), for callers that do not review in between. */
export async function run(file, opts, cb = {}) {
  const state = await resolve(file, opts, cb);
  const out = await finish(state, opts, cb);
  return { ...out, state };
}
