# repex

Snapshot a local repository into a single document for an LLM — works whether
the LLM has file-access tools or not, and also produces human-readable Word /
Excel / LibreOffice reports.

A single-file Python script (`repex.py`). No install.

## Get it

Only `repex.py` is needed — drop it anywhere on disk and run it. The
`LICENSE` and `README.md` live in the repo for documentation; they don't
need to be installed alongside the script.

```sh
# bash / zsh / macOS / WSL / Git Bash
curl -O https://raw.githubusercontent.com/jakubtoczek/repex/main/repex.py

# PowerShell
Invoke-WebRequest https://raw.githubusercontent.com/jakubtoczek/repex/main/repex.py -OutFile repex.py
```

Or clone the repo and run `repex.py` directly.

## Quick start

```sh
py repex.py . -s agent -o map.md       # LLM-agent orientation, no inline content
py repex.py . -s llm -o context.md     # paste-into-chat LLM context, full inline
py repex.py . -o report.docx           # human-readable Word report
```

## Why

repex serves two LLM modes (and one human mode):

- **Agents with file tools** (Claude Code, Cursor, aider) want a *map* —
  enough cross-file intelligence to skip 10–15 rounds of `grep` / `read`
  just figuring out what calls what. Use `--sections agent`.
- **LLMs without tools** (paste into a chat) want a *map plus the full
  content* in one shot. Use `--sections llm`.
- **Humans** want a polished overview document (Word / Excel / LibreOffice)
  to skim or hand to a non-developer. Use `--sections human` plus the
  matching `-o report.docx` / `.xlsx` / `.odt`.

Every mode reuses the same analysis pass:

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
  U=untracked, N=no-repo; see "Git discovery" below)
- **entries** — per-file structured headers + full content (the bulky one)

## Presets

`-s` / `--sections` selects which sections to render. The preset table is the
core choice — pick the one matching your audience:

| Preset    | Audience                            | Sections |
|-----------|-------------------------------------|----------|
| `default` / `all` | Everything (hand-off snapshot)        | every section |
| `agent`   | LLM with file tools (Claude Code, Cursor, aider) | glance, architecture, entry_points, trace, core, toc |
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

## Git discovery

When you pass a path, repex looks for git repositories in three places:

1. **Self** — `path/.git` exists. Standard repo (the usual case).
2. **Ancestor** — walk up the parents (matches `git status` from any
   subfolder).
3. **Children** — if neither of the above, scan immediate subfolders. Useful
   for **workspace layouts** where one parent folder groups several
   independent projects, each with its own `.git`.

Every file is then tagged in the export with a marker:

- `T` — tracked in some discovered repo
- `U` — untracked in some discovered repo
- `N` — outside every discovered repo (no-repo)

In workspace mode you get a mix of `T` / `U` / `N` files in the same export,
and the document header lists every discovered repo with its tracked /
untracked counts plus the no-repo file total. `.gitignore` filtering and
`--since` loop over every discovered root.

## Output formats

Format is inferred from the `-o` / `--output` extension; `-f` / `--format`
overrides it. When neither is set, the default is **md** — stdlib-only and
opens anywhere, matching the LLM-first use case. Pass `-o report.docx`
(or any other extension) for the human-readable formats.

| Format            | Extension | Dependency       | Respects `--sections`?     |
|-------------------|-----------|------------------|----------------------------|
| Markdown          | `.md`     | none (stdlib)    | yes — **default**          |
| Word              | `.docx`   | `python-docx`    | yes                        |
| LibreOffice text  | `.odt`    | `odfpy`          | yes                        |
| JSON              | `.json`   | none (stdlib)    | no (one record per file)   |
| Excel             | `.xlsx`   | `openpyxl`       | no (one row per file)      |
| LibreOffice sheet | `.ods`    | `odfpy`          | no (one row per file)      |

Spreadsheet outputs are useful for triage (sort by `used_by`, filter by
language, scan `entry_signals`, and in workspace mode sort by `git_origin`
to group files by subrepo) and for handing a non-developer something they
can open and explore.

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
| `--since <ref>`         | Restrict the export to files changed since the given git revision (committed diff + working tree + untracked-not-ignored). Runs against every discovered git root, so it works in workspace mode and from a subfolder of a repo, not only on a self-repo path. |
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

For very large folders (workspace layouts mixing a small project with
thousands of unrelated archive / docs / build artifacts), the `llm`
preset can balloon past any reasonable context window. Combine
`--token-budget N` with `--exclude-dir` and `--strip-comments`, or point
repex at the actual project subfolder instead of the workspace root.

## License

0BSD (BSD Zero Clause). No attribution required. See [LICENSE](LICENSE).
