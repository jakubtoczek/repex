# repex

Snapshot a local repository for an LLM in one command — works whether the LLM
has file-access tools or not.

A single-file Python script (`repex.py`). Drop it anywhere, no install.

## Why

When you start a session on an unfamiliar repo, you usually burn 10–15 rounds
of `grep`/`read` just figuring out what calls what, which file is the entry
point, and how things are wired. repex builds that map up front:

- a 2-level static call sketch starting from detected entry points,
- a `used_by` inverse index across all files,
- ranked "core files" by inbound references and size,
- an auto-grouped architecture view of directories,
- a unified table of contents.

There are two main modes:

### `--sections agent` — for an LLM **with** file tools (Claude Code, Cursor, …)

A compact orientation document (no inline file contents). The agent reads
files on demand; repex provides the cross-file intelligence that would
otherwise take many tool calls to assemble.

```sh
py repex.py "C:\Projects\MyRepo" --sections agent -o map.md
```

Read the resulting `map.md` once at session start, then work normally with
Read/Grep on the live tree.

### `--sections llm` — for an LLM **without** tools (paste into chat)

Same map as above, plus the full content of every file inline, so the model
sees everything in one shot.

```sh
py repex.py "C:\Projects\MyRepo" --sections llm -o context.md
```

### Other presets

| Preset    | Audience                                  | What's included |
|-----------|-------------------------------------------|-----------------|
| `default` / `all` | Everything (hand-off snapshot)    | every section |
| `agent`   | LLM with file-access tools                | glance, architecture, entry_points, trace, core, toc |
| `llm`     | LLM without tools (paste into chat)       | glance, architecture, entry_points, trace, core, **entries** |
| `human`   | Human reader skim                         | glance, recent, architecture, toc |

Mix and match: `--sections agent,+recent` or `--sections all,-entries`.

## Output formats

Format is inferred from the `--output` extension; `-f`/`--format` overrides it.

| Format | Extension | Dependency       | Section-aware |
|--------|-----------|------------------|---------------|
| Word   | `.docx`   | `python-docx`    | yes           |
| Markdown | `.md`   | none (stdlib)    | yes           |
| LibreOffice text | `.odt` | `odfpy`    | yes           |
| Excel  | `.xlsx`   | `openpyxl`       | full (one row per file) |
| LibreOffice sheet | `.ods` | `odfpy`   | full (one row per file) |
| JSON   | `.json`   | none (stdlib)    | full          |

The spreadsheet formats are great for triage — sort by `used_by`, filter by
language, scan `entry_signals` — and for handing a non-developer something
they can open and explore.

Every export stamps `repex.py (VERSION)` as the document author/creator
(docx/xlsx/odt/ods) or a top-level generator field (json/md), so a teammate
can tell which version produced the file.

## Sections

- **glance** — file count, lines by kind/language, largest file, README excerpt
- **architecture** — auto-grouped directories with role labels
- **entry_points** — language-tagged signals (`__main__` guards, `main()`, default exports, …)
- **trace** — two-level static call sketch from entry points
- **core** — files ranked by `used_by` + size
- **toc** — compact unified table of contents
- **entries** — per-file structured headers + full content (the bulky one)
- **recent** — top 5 recently modified files

## Supported languages

Function counts and the call sketch work for:

Python · C · C++ · C# · Java · JavaScript · TypeScript · Rust · Go · Ruby ·
PHP · Kotlin · Scala · Swift · R · Bash

Many more are recognized and embedded as-is (HTML/CSS, JSON/YAML/TOML/XML,
Markdown, SQL, CMake, batch, PowerShell, …).

## Usage

```sh
# Agent orientation map (markdown, no inline content)
py repex.py "C:\Projects\MyRepo" --sections agent -o map.md

# Full LLM context for paste-into-chat
py repex.py "C:\Projects\MyRepo" --sections llm -o context.md

# Word export, everything (human-friendly)
py repex.py "C:\Projects\MyRepo" -o myrepo.docx

# Excel triage sheet (one row per file)
py repex.py "C:\Projects\MyRepo" -o myrepo.xlsx

# Force format regardless of extension
py repex.py "C:\Projects\MyRepo" -f json -o report.bin
```

Run `py repex.py --help` for the full option list.

## License

MIT. See [LICENSE](LICENSE).
