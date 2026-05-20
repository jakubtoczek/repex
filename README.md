# repex

Export a local repository or folder to one of six formats for LLM context
ingestion or human overview.

A single-file Python script (`repex.py`) — drop it anywhere, no install.

## Output formats

| Format | Extension | Dependency       | Section-aware |
|--------|-----------|------------------|---------------|
| Word   | `.docx`   | `python-docx`    | yes           |
| Excel  | `.xlsx`   | `openpyxl`       | no (always full) |
| Markdown | `.md`   | none (stdlib)    | yes           |
| JSON   | `.json`   | none (stdlib)    | no (always full) |
| LibreOffice text | `.odt` | `odfpy`    | yes           |
| LibreOffice sheet | `.ods` | `odfpy`   | no (always full) |

The format is inferred from `--output`'s extension; an explicit `-f`/`--format`
always wins.

## What's in the export

- **glance** — file count, lines by kind/language, largest file, README excerpt
- **architecture** — auto-grouped directories with role labels
- **entry_points** — language-tagged entry-point signals (Python `__main__`,
  C `main`, Rust `fn main`, Node default exports, ...)
- **trace** — two-level static call sketch from entry points
- **core** — files ranked by `used_by` and size
- **toc** — compact unified table of contents
- **entries** — per-file structured headers + full content
- **recent** — top 5 recently modified files

## Section presets

| Preset    | Audience                                | Sections |
|-----------|-----------------------------------------|----------|
| `default` / `all` | Everything (hand-off snapshot) | glance, recent, architecture, entry_points, trace, core, toc, entries |
| `llm`     | LLM without tools (paste into chat)     | glance, architecture, entry_points, trace, core, entries |
| `agent`   | LLM with file tools (Claude Code, ...)  | glance, architecture, entry_points, trace, core, toc |
| `human`   | Human reader skim                        | glance, recent, architecture, toc |

Mix and match with `+name` / `-name`:

```sh
py repex.py myrepo --sections llm,+toc       # llm preset plus toc
py repex.py myrepo --sections all,-entries   # everything except contents
```

## Supported languages

Function counts and call-sketch traces work for:

Python, C, C++, C#, Java, JavaScript, TypeScript, Rust, Go, Ruby, PHP, Kotlin,
Scala, Swift, R, Bash.

Many more are recognized and embedded as-is (HTML/CSS, JSON/YAML/TOML/XML,
Markdown, SQL, CMake, batch, PowerShell, ...).

## Usage

```sh
# Word export (default)
py repex.py "C:\Projects\MyRepo" -o myrepo.docx

# LLM-with-tools index (markdown, no inline content)
py repex.py "C:\Projects\MyRepo" --sections agent -o myrepo.md

# Excel sheet (one row per file)
py repex.py "C:\Projects\MyRepo" -o myrepo.xlsx

# Force format regardless of extension
py repex.py "C:\Projects\MyRepo" -f json -o report.bin
```

Run `py repex.py --help` for the full option list.

## Output metadata

Every export stamps `repex.py (VERSION)` as the document author/creator
(`docx`, `xlsx`, `odt`, `ods`) or as a top-level generator field (`json`,
`md`), so a teammate opening the file can tell which version produced it.
