<!--
  The centred header is written as HTML rather than Markdown-inside-HTML on
  purpose. GitHub parses Markdown nested in a block-level tag; most other
  renderers — package registries, editor previews, anything built on plain
  CommonMark — treat everything to the closing tag as literal text, and the
  heading and badges come out as raw "###" and "![live](…)". Image sources are
  absolute for the same reason: a relative path only resolves when the file is
  read inside the repository.
-->
<div align="center">
  <img src="https://raw.githubusercontent.com/ORG-Karur-DataCenter/Z--Link/main/.github/assets/banner.svg" alt="Z-Link — plain-text citations become real, linked Zotero citations" width="880">
  <br><br>
  <h3><a href="https://org-karur-datacenter.github.io/Z--Link/"><b>→ Open Z-Link</b></a></h3>
  <sub>No install. No account here. Your manuscript never leaves the tab.</sub>
  <br><br>
  <img src="https://img.shields.io/badge/app-live-1f7a3f?style=for-the-badge" alt="live">
  <img src="https://img.shields.io/badge/measured-82%2F82-1f7a3f?style=for-the-badge" alt="82 of 82 references resolved">
  <img src="https://img.shields.io/badge/engine%20parity-55%2F55-1f7a3f?style=for-the-badge" alt="55 of 55 injected port mistakes caught">
  <img src="https://img.shields.io/badge/dependencies-zero-8a6a20?style=for-the-badge" alt="zero dependencies">
</div>

<br>

---

Drop in a `.docx` with numbered citations and a plain bibliography. Z-Link finds each
reference across five indexes, verifies it is genuinely the right paper, adds the items to
your Zotero library, and returns the document with **live Zotero citations** in it — open
it in Word, nothing else to install or run.

It does **not** decide where a citation belongs — that is your judgment while writing.
Anything it cannot confirm becomes a visible `{NEEDS REVIEW: n}` rather than a live
citation quietly pointing at the wrong paper.

<br>

## Start

<table>
<tr>
<td width="50%" valign="top">
  <p><b>In the browser</b> — nothing to install</p>
  <ol>
    <li><a href="https://org-karur-datacenter.github.io/Z--Link/">Open Z-Link</a></li>
    <li>Paste your Zotero userID and API key
      <br><sub>the <b>First time here</b> button walks you through it</sub></li>
    <li>Drop in your <code>.docx</code></li>
  </ol>
</td>
<td width="50%" valign="top">
  <p><b>On the command line</b></p>
  <pre><code>pip install -r requirements.txt

python -m zotprep --zotero-userid 1234567 \
  --zotero-key KEY --save-credentials

python -m zotprep --manuscript paper.docx</code></pre>
  <sub>Dry run is the default. Add <code>--live</code> to write.</sub>
</td>
</tr>
</table>

Then open the `.docx` in Word and press **Refresh** on the Zotero tab. The citations are
Zotero field codes, the document already carries the citation style you chose — Vancouver
(superscript) unless you changed it — and the reference list rebuilds itself in place, so
there is nothing to set up and no style to pick.

That holds on anyone's machine, not only yours. Each citation embeds the full reference
rather than a link into the library it came from, and the style and the bibliography travel
with the document, so a co-author or a supervisor who opens the file gets the same
manuscript you do. What they need is Word on the desktop with Zotero running: Word Online,
Google Docs and Pages have no Zotero plugin, and Word installed from the Microsoft Store
blocks it.

<sub>A Scannable Cite copy for the <b>ODF Scan</b> plugin is produced alongside, for
LibreOffice or for checking the markers before they become citations — the plugin is
<a href="web/vendor/">bundled here</a>. It is no longer the main route: ODF Scan finds its
markers by scanning the file as text, and in a manuscript with images that can put a
citation inside a picture's XML and produce a file Word refuses to open.</sub>

<br>

## Why trust it

|  |  |
|:--|:--|
| **82 / 82 references** resolved across two real manuscripts, no manual pass | [details](docs/HOW-IT-WORKS.md#advisories) |
| **Acceptance is a conjunction**, never a score crossing a line. A wrong paper can fake one signal; it cannot fake the title, year, volume *and* first page together | [the accept gate](docs/HOW-IT-WORKS.md#the-accept-gate) |
| **It refuses rather than guesses.** Unconfirmed references stay visibly flagged | [advisories](docs/HOW-IT-WORKS.md#advisories) |
| **The browser engine is a verified port**, not a rewrite — compared against the Python original to the exact float, with 47/47 injected bugs caught | [verification](web/README.md) |
| **Nothing leaves your machine.** No server, no upload, no third-party script on the page | — |

<br>

## Two front ends, one engine

|  | Z-Link — browser | zotprep — CLI |
|:--|:--|:--|
| Install | none | `pip install -r requirements.txt` |
| Writes to Zotero | every run | only with `--live` |
| Preview without writing | — | `--dry-run`, the default |
| Remembers your decisions | for the session | forever, in SQLite |
| Separate bibliography file | — | `--bibliography` |

Use the CLI when you want a preview pass or decisions that persist.

<br>

## Documentation

| | |
|:--|:--|
| [**How it works**](docs/HOW-IT-WORKS.md) | Citation notation, the accept gate, advisories, providers, every flag |
| [**Verification**](web/README.md) | How the browser port is proven equivalent, and how to run it |
| [**Plugin notice**](web/vendor/NOTICE.md) | The bundled ODF Scan `.xpi`, its licence and provenance |

<br>

## Getting a Zotero key

At [zotero.org/settings/keys](https://www.zotero.org/settings/keys):

1. Your **userID** is the number in *"Your userID for use in API calls is …"* — not your
   username.
2. **Create new private key** → tick **Allow library access** *and* **Allow write
   access**. Missing the second one is the usual cause of an opaque `403`.
3. The key is shown once. Copy it immediately.

The key is stored in your browser only, and sent only to `api.zotero.org`.

<br>

## Layout

```
zotprep/     the engine and the CLI
web/         Z-Link — the browser build, its parity harnesses, the bundled plugin
docs/        how it works
```

<br>

---

<div align="center">
<sub>

Bundles [ODF Scan for Zotero](https://github.com/Juris-M/zotero-odf-scan-plugin) by Sebastian
Karcher and Frank Bennett · AGPL-3.0-or-later · [notice](web/vendor/NOTICE.md)

</sub>
</div>
