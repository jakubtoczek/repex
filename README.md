# repex

Snapshot a local repository into a single document for an LLM — works whether
the LLM has file-access tools or not, and also produces human-readable Word /
Excel / LibreOffice reports.

A single-file Python script (`repex.py`). No install.

## Get it

```sh
# bash / zsh / macOS / WSL / Git Bash
curl -O https://raw.githubusercontent.com/jakubtoczek/repex/main/repex.py

# PowerShell
Invoke-WebRequest https://raw.githubusercontent.com/jakubtoczek/repex/main/repex.py -OutFile repex.py
```

Or clone the repo and run `repex.py` directly.

## Quick start

```sh
py repex.py . -o map.md           # markdown map of the current folder
py repex.py . -s agent            # compact orientation for an LLM agent
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

## Sections

The building blocks the section-aware formats (`docx`, `md`, `odt`) compose
into a document. Canonical render order:

- **glance** — file count, lines by kind/language, largest file, README excerpt
- **recent** — top 5 recently modified files
- **architecture** — directory groups with auto-derived role labels
- **entry_points** — language-tagged signals (`__main__` guards, `main()`,
  default exports, …)
- **trace** — two-level static call sketch from entry points
- **core** — files ranked by `used_by` + size
- **toc** — compact unified table of contents (entries flagged T=tracked,
  U=untracked)
- **entries** — per-file structured headers + full content (the bulky one)

## Presets

`-s` / `--sections` selects which sections to render. The preset table is the
core choice — pick the one matching your audience:

| Preset    | Audience                            | Sections |
|-----------|-------------------------------------|----------|
| `default` / `all` | Everything (hand-off snapshot)        | every section |
| `agent`   | LLM with file tools (Claude Code, Cursor, Aider) | glance, architecture, entry_points, trace, core, toc |
| `llm`     | LLM without tools (paste into chat) | glance, architecture, entry_points, trace, core, **entries** |
| `human`   | Human reader skim                   | glance, recent, architecture, toc |

The two LLM presets are the main differentiator:

- **`agent`** is a compact orientation document, no inline file contents.
  Give the agent `map.md` at session start; it then works normally with
  Read / Grep on the live tree.
- **`llm`** is the same map plus full content inline, so a tool-less model
  sees everything in one shot.

Mix and match with `+name` / `-name`, or list sections explicitly:

```sh
-s agent,+recent          # agent preset plus recent files
-s all,-entries           # everything except bulky content
-s glance,toc,entries     # explicit list
```

## Output formats

Format is inferred from the `-o` / `--output` extension; `-f` / `--format`
overrides it.

| Format            | Extension | Dependency       | Respects `--sections`?     |
|-------------------|-----------|------------------|----------------------------|
| Word              | `.docx`   | `python-docx`    | yes                        |
| Markdown          | `.md`     | none (stdlib)    | yes                        |
| LibreOffice text  | `.odt`    | `odfpy`          | yes                        |
| Excel             | `.xlsx`   | `openpyxl`       | no (one row per file)      |
| LibreOffice sheet | `.ods`    | `odfpy`          | no (one row per file)      |
| JSON              | `.json`   | none (stdlib)    | no (one record per file)   |

Spreadsheet outputs are useful for triage (sort by `used_by`, filter by
language, scan `entry_signals`) and for handing a non-developer something
they can open and explore.

Every format carries a per-file **token estimate** (heading / TOC for text
formats, column / field for spreadsheets and JSON) and embeds a
content-tokens total in the document header. Tokens are not proportional
to size (a 10 kB Python file is ~2× a 10 kB Java file in tokens), so the
count adds real information for cost estimation and LLM-fit decisions
even on the human path.

Every export stamps `repex.py (VERSION)` as the document author/creator
(`docx`/`xlsx`/`odt`/`ods`) or a top-level `generator` field (`json`/`md`),
so a teammate can tell which version produced the file.

## Workflow flags

| Flag                    | Effect |
|-------------------------|--------|
| `-h` / `--help`         | Standard argparse help screen. |
| `--no-gitignore`        | Disable the default `.gitignore` filtering. |
| `--since <ref>`         | Restrict the export to files changed since the given git revision (committed diff + working tree + untracked-not-ignored). Requires a git repository. |
| `--strip-comments`      | Remove line and block comments from code content before embedding. Strings preserved. Saves 15–30 % tokens on the `llm` preset. |
| `--token-budget N`      | Markdown / JSON only. Trim content of low-ranked files (by `used_by` + size) until the rendered output fits ~N tokens. Uses `tiktoken` if installed, else a 4-char/token estimate. For non-LLM formats use `--strip-comments` and let size shrink naturally. |
| `--token-model M`       | Tokenizer model name passed to `tiktoken`. Default `gpt-4o`. cl100k is a reasonable proxy for Claude. |
| `--remote OWNER/REPO`   | Shallow-clone a remote into a tempdir and export it instead of a local path (also accepts a full clone URL). Tempdir is removed when the run finishes. |
| `--clipboard`           | Markdown / JSON only. Also copy the rendered output to the system clipboard. |

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
`pyperclip`) or prints a clear "install with: …" hint when missing
(`python-docx`, `openpyxl`, `odfpy`).

## Examples

```sh
# Paste-into-chat LLM context, comment-stripped, budgeted to ~80k tokens
py repex.py "C:\Projects\MyRepo" -s llm --strip-comments \
            --token-budget 80000 -o context.md

# Just what changed since main (diff scope, full content for review)
py repex.py "C:\Projects\MyRepo" --since main -s llm -o changes.md

# Map a remote repo without cloning manually
py repex.py --remote yamadashy/repomix -s agent -o repomix-map.md

# Send markdown straight to the clipboard
py repex.py "C:\Projects\MyRepo" -s llm --clipboard -o context.md
```

Run `py repex.py -h` for the full option list.

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
