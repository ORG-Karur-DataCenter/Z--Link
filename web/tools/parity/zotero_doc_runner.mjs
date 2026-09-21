/**
 * Checks on the parts of a Zotero document that are not citations.
 *
 * A citation that is wrong is visible — the wrong name appears in the text. The
 * things checked here are invisible when they break: preferences that Word
 * cannot read, a relationship that points nowhere, a custom property one
 * character over Word's limit. Nothing in the run fails, the file opens, and the
 * damage only shows when the manuscript is opened by someone who is not the
 * author — which is the case this was written for and the one hardest to notice.
 *
 * No DOM, and so no dependency: everything asserted here is string-level, which
 * is where the format rules actually live.
 */
import {
  DEFAULT_STYLE, customPropertiesXml, prefChunks, withCustomContentType,
  withCustomRelationship, zoteroPrefs,
} from '../../src/docx.js';

const checks = [];
const check = (name, ok, detail = '') => checks.push([name, Boolean(ok), String(detail)]);

// --- the preferences payload --------------------------------------------------
const prefs = zoteroPrefs(DEFAULT_STYLE, 'SESSION1');
check('prefs declare data-version 3', prefs.includes('data-version="3"'),
  'Zotero parses anything else as a pre-5.0 document');
check('prefs name the style', prefs.includes(`id="${DEFAULT_STYLE}"`));
check('prefs say the field type is Field', prefs.includes('name="fieldType" value="Field"'),
  'Bookmark is the LibreOffice storage, and Word will not find the citations');
check('prefs say the references are stored', prefs.includes('name="storeReferences" value="true"'),
  'this is what tells Zotero to trust the embedded itemData on another machine');
check('prefs keep the resolver\'s journal abbreviations',
  prefs.includes('name="automaticJournalAbbreviations" value="false"'),
  'true makes Zotero derive its own and ignore the field the resolver filled');
check('prefs carry a session id', /<session id="SESSION1"\/>/.test(prefs));
check('prefs make no claim about a Zotero version', !prefs.includes('zotero-version'),
  'no Zotero release wrote this file, and saying one did would be untrue');

// --- Word's 255-character limit on a custom property --------------------------
const chunks = prefChunks(prefs);
check('preferences are split across properties', chunks.length >= 1);
check('no chunk exceeds 255 characters', chunks.every((c) => c.length <= 255),
  `longest chunk is ${Math.max(...chunks.map((c) => c.length))}`);
check('chunks reassemble to the original', chunks.join('') === prefs);

// --- the custom-properties part ----------------------------------------------
const fresh = customPropertiesXml(null, ['one', 'two']);
check('properties are numbered from 1', fresh.includes('name="ZOTERO_PREF_1"')
  && fresh.includes('name="ZOTERO_PREF_2"'));
check('pids start at 2', /pid="2"/.test(fresh), 'pid 1 is reserved');
check('pids are unique', new Set([...fresh.matchAll(/pid="(\d+)"/g)].map((m) => m[1])).size === 2);

const foreign = '<?xml version="1.0"?><Properties xmlns="x" xmlns:vt="y">'
  + '<property fmtid="{F}" pid="2" name="JournalRef"><vt:lpwstr>BJJ-2024-0417</vt:lpwstr></property>'
  + '<property fmtid="{F}" pid="7" name="Confidential"><vt:bool>true</vt:bool></property>'
  + '</Properties>';
const merged = customPropertiesXml(foreign, ['one']);
check('a document\'s own properties survive', merged.includes('BJJ-2024-0417'));
check('a typed property is not rewritten', merged.includes('<vt:bool>true</vt:bool>'),
  'parsing and reserialising other people\'s metadata loses type and namespace');
check('new pids clear the highest in use', /pid="8"/.test(merged),
  'a repeated pid makes Word reject the part');

const rerun = customPropertiesXml(merged, ['only-one']);
check('a re-run replaces its own preferences',
  [...rerun.matchAll(/ZOTERO_PREF_\d+/g)].length === 1,
  'otherwise two generations of preferences accumulate');
check('a re-run still keeps foreign properties', rerun.includes('BJJ-2024-0417'));

check('values are XML-escaped', customPropertiesXml(null, ['a & b < c'])
  .includes('a &amp; b &lt; c'), 'an unescaped value is a malformed part');

// --- content types and the package relationship -------------------------------
const types = withCustomContentType('<Types xmlns="x"><Default Extension="xml"/></Types>');
check('the part is given a content type', types.includes('custom-properties+xml'));
check('the content type is not added twice',
  [...withCustomContentType(types).matchAll(/docProps\/custom\.xml/g)].length === 1);

const rels = withCustomRelationship('<Relationships xmlns="x">'
  + '<Relationship Id="rId1" Type="t" Target="docProps/core.xml"/>'
  + '<Relationship Id="rId9" Type="t" Target="word/document.xml"/></Relationships>');
check('the part is related to the package', rels.includes('custom-properties'));
check('the relationship id does not collide', /Id="rId10"/.test(rels),
  'a duplicate id makes Word refuse the file');
check('the relationship is not added twice',
  [...withCustomRelationship(rels).matchAll(/custom-properties/g)].length === 1);

process.stdout.write(JSON.stringify(checks));
