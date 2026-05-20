# repex

Snapshot a local repository into a single document for an LLM — works whether
the LLM has file-access tools or not, and also produces human-readable Word /
Excel / LibreOffice reports.

A single-file Python script (`repex.py`). No install.

## Get it

```sh
curl -O https://raw.githubusercontent.com/jakubtoczek/repex/main/repex.py
# or clone the repo and run repex.py directly
```

## Quick start

```sh
py repex.py . -o map.md           # markdown map of the current folder
py repex.py . --sections agent    # compact orientation for an LLM agent
py repex.py . -o report.docx      # human-readable Word report
```

## Why

When you start a session on an unfamiliar repo, you usually burn 10–15 rounds
of `grep` / `read` figuring out what calls what, which file is the entry
point, and how things are wired. repex builds that map up front:

- a 2-level static call sketch from detected entry points,
- a `used_by` inverse index across all files,
- ranked "core files" by inbound references and size,
- an auto-grouped architecture view of directories,
- a unified table of contents.

## Two main modes

### `--sections agent` — LLM **with** file tools (Claude Code, Cursor, Aider, …)

A compact orientation document, no inline file contents. The agent reads
files on demand; repex provides the cross-file intelligence that would
otherwise take many tool calls to assemble.

```sh
py repex.py "C:\Projects\MyRepo" --sections agent -o map.md
```

Give the agent `map.md` at session start; it then works normally with
Read / Grep on the live tree.

### `--sections llm` — LLM **without** tools (paste into chat)

Same map as above, plus the full content of every file inline, so the model
sees everything in one shot.

```sh
py repex.py "C:\Projects\MyRepo" --sections llm -o context.md
```

## Other presets

| Preset    | Audience                                  | What's included |
|-----------|-------------------------------------------|-----------------|
| `default` / `all` | Everything (hand-off snapshot)    | every section |
| `agent`   | LLM with file-access tools                | glance, architecture, entry_points, trace, core, toc |
| `llm`     | LLM without tools (paste into chat)       | glance, architecture, entry_points, trace, core, **entries** |
| `human`   | Human reader skim                         | glance, recent, architecture, toc |

Mix and match with `+name` / `-name`, or list sections explicitly:

```sh
--sections agent,+recent          # agent preset plus recent files
--sections all,-entries           # everything except bulky content
--sections glance,toc,entries     # explicit list
```

## Sections

Canonical render order (applies to `docx`, `md`, `odt`):

- **glance** — file count, lines by kind/language, largest file, README excerpt
- **recent** — top 5 recently modified files
- **architecture** — directory groups with auto-derived role labels
- **entry_points** — language-tagged signals (`__main__` guards, `main()`, default exports, …)
- **trace** — two-level static call sketch from entry points
- **core** — files ranked by `used_by` + size
- **toc** — compact unified table of contents (entries flagged T=tracked, U=untracked)
- **entries** — per-file structured headers + full content (the bulky one)

## Output formats

Format is inferred from the `--output` extension; `-f` / `--format` overrides it.

| Format | Extension | Dependency       | Respects `--sections`? |
|--------|-----------|------------------|------------------------|
| Word              | `.docx` | `python-docx` | yes |
| Markdown          | `.md`   | none (stdlib) | yes |
| LibreOffice text  | `.odt`  | `odfpy`       | yes |
| Excel             | `.xlsx` | `openpyxl`    | no (always full) |
| LibreOffice sheet | `.ods`  | `odfpy`       | no (always full) |
| JSON              | `.json` | none (stdlib) | no (always full) |

The "no (always full)" formats produce one row per file (`xlsx` / `ods`) or
one record per file (`json`) — `--sections` doesn't apply. They're useful for
triage (sort by `used_by`, filter by language, scan `entry_signals`) and for
handing a non-developer something they can open and explore.

Every export stamps `repex.py (VERSION)` as the document author/creator
(`docx`/`xlsx`/`odt`/`ods`) or a top-level `generator` field (`json`/`md`),
so a teammate can tell which version produced the file.

## Workflow flags

| Flag | Effect |
|------|--------|
| `--no-gitignore` | Disable the default `.gitignore` filtering. |
| `--since <ref>` | Restrict the export to files changed since the given git revision (committed diff + working tree + untracked-not-ignored). Requires a git repository. |
| `--strip-comments` | Remove line and block comments from code content before embedding. Strings are preserved. Typically saves 15–30 % of tokens on the `llm` preset. |
| `--token-budget N` | Markdown / JSON only. Drop content of low-ranked files (by `used_by` + size) until the rendered output fits ~N tokens. Uses `tiktoken` if installed, else a 4-char/token estimate. |
| `--token-model M` | Tokenizer model name passed to `tiktoken`. Default `gpt-4o`. cl100k is a reasonable proxy for Claude. |
| `--remote OWNER/REPO` | Shallow-clone a remote into a tempdir and export it instead of a local path (also accepts a full clone URL). Tempdir is removed when the run finishes. |
| `--clipboard` | Markdown / JSON only. Also copy the rendered output to the system clipboard. |

In `md` and `json` outputs every file record carries a per-file token
estimate (`~N tokens` in the markdown TOC and entry bullets,
`tokens_estimate` field in JSON). The total content tokens are embedded
near the top of the file (HTML comment + bullet for md, top-level `tokens`
object for json) so the count survives the run.

`--sections` accepts a `-s` short alias, and argparse prefix matching
already handles `--section`, `--sect`, `--sec`, etc.

## Supported languages

Function counts and the call sketch work for:

Python · C · C++ · C# · Java · JavaScript · TypeScript · Rust · Go · Ruby ·
PHP · Kotlin · Scala · Swift · R · Bash

Many more are recognized and embedded as-is (HTML/CSS, JSON/YAML/TOML/XML,
Markdown, SQL, CMake, batch, PowerShell, …). Add more per-run with `--ext`.

## Dependencies

repex itself runs on stdlib Python. Install extras only for the
formats / flags you actually use:

```sh
py -m pip install python-docx     # for .docx
py -m pip install openpyxl        # for .xlsx
py -m pip install odfpy           # for .odt / .ods
py -m pip install tiktoken        # exact token counting for --token-budget
py -m pip install pyperclip       # reliable cross-platform --clipboard
```

All are optional — repex either falls back gracefully (`tiktoken`,
`pyperclip`) or prints a clear "install with: ..." hint when missing
(`python-docx`, `openpyxl`, `odfpy`).

## Examples

```sh
# Agent orientation map (no inline content; for Claude Code, Cursor, ...)
py repex.py "C:\Projects\MyRepo" --sections agent -o map.md

# Paste-into-chat LLM context, comment-stripped, budgeted to ~80k tokens
py repex.py "C:\Projects\MyRepo" --sections llm --strip-comments \
            --token-budget 80000 -o context.md

# Just what changed since main (diff scope)
py repex.py "C:\Projects\MyRepo" --since main --sections llm -o changes.md

# Map a remote repo without cloning manually
py repex.py --remote yamadashy/repomix --sections agent -o repomix-map.md

# Word export, everything (human-friendly)
py repex.py "C:\Projects\MyRepo" -o myrepo.docx

# Excel triage sheet (one row per file)
py repex.py "C:\Projects\MyRepo" -o myrepo.xlsx

# Force format regardless of extension
py repex.py "C:\Projects\MyRepo" -f json -o report.bin

# Send markdown straight to the clipboard
py repex.py "C:\Projects\MyRepo" --sections llm --clipboard -o context.md
```

Run `py repex.py --help` for the full option list.

## When not to use repex

- **You need accurate symbol resolution / per-call references.** Use
  [aider](https://aider.chat)'s repo map — it parses with tree-sitter and
  builds a real symbol graph, not regex-best-effort.
- **You need a large ecosystem, secret scanning, or model-specific output
  templates.** Use [repomix](https://github.com/yamadashy/repomix) — much
  larger community, more configuration knobs, npm-distributed.
- **You need Handlebars-style template-driven prompt assembly.** Use
  [code2prompt](https://github.com/mufeedvh/code2prompt) — designed
  specifically as a customizable prompt builder.

## License

MIT. See [LICENSE](LICENSE).
