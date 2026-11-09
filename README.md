# Build an editable presentation from an explicit slide structure

**Presentation Kit** turns supplied text and tables into a local PPTX file.
Four fixed layouts: cover, bullets, two columns and a table. Text and table
cells are native editable OOXML objects, rather than pictures of slides.
Content that fails the fit check stops the build.

Python 3.9+, standard library only. No model, API key, account, subscription
or paid core is required. The tool does not write the story or invent figures;
it assembles the material you supply in the order you specify.

Prepared by **Дмитрий Лукашов / Dmitry Lukashov**. A standalone open component
of the [Lookatshow system for AI agents](https://github.com/ilookatshow-gif/lookatshow-open-tools).
[Русский](README.ru.md) · [MIT](LICENSE) · [Authorship and rights](NOTICE.md)

## First file

Download [release v0.1.0](https://github.com/ilookatshow-gif/presentation-kit/releases/tag/v0.1.0)
and unpack it into a new directory. From that directory:

```bash
python3 -B presentation_kit.py --json doctor
python3 -B presentation_kit.py --json demo --output ./synthetic-demo.pptx
python3 -B presentation_kit.py --json inspect ./synthetic-demo.pptx
```

The demo uses `examples/synthetic-deck.json`: five slides covering all four layouts, with a closing bullets slide.
Its scenario, wording and numbers are invented. It is a teaching fixture,
not a client presentation or evidence of a Lookatshow business result.

A [ready-made editable teaching PPTX](examples/synthetic-demo.pptx) is included.
Its five slides were rendered with LibreOffice and inspected at full size:
text and the native table are visible, with no unintended overlap or clipping
found. This verifies this example in LibreOffice; it was not tested in
Microsoft PowerPoint.

![Five teaching slides rendered with LibreOffice](assets/demo-preview.png)

Open the resulting file in your presentation editor and inspect the slides.
`inspect` reads this module's own structural format. A successful ZIP check
does not establish that a file was rendered, edited or checked in PowerPoint.

Existing output is preserved. Choose a new path on a repeat run, or deliberately
use `--replace`.

## Install the command

On POSIX, choose a new installation prefix and a command directory:

```bash
python3 -B install.py --prefix "$HOME/.local/share/presentation-kit" --bin-dir "$HOME/.local/bin"
export PATH="$HOME/.local/bin:$PATH"
cd /tmp
presentation-kit --json doctor
presentation-kit --json demo --output ./presentation-kit-demo.pptx
```

The installer makes a separate copy of the required files and a launcher that
uses the Python interpreter which ran the installer. It refuses to overwrite
an existing installation or command. Keep that Python installation available.
Store your own JSON and presentations outside the installation prefix.

The native installer supports POSIX. On Windows, run the source CLI through
`python presentation_kit.py ...`; substitute your interpreter's command name
for `python3`. This is a separate launch path, not a claim that every editor
version or operating system has been tested.

Optional `presentation-kit --json onboard` shows the author's links once and
creates `.onboarded` beside the installed CLI. Normal JSON and generated slides
contain no promotional inserts.

Uninstall from any working directory with the same two paths:

```bash
python3 -B "$HOME/.local/share/presentation-kit/uninstall.py" --prefix "$HOME/.local/share/presentation-kit" --bin-dir "$HOME/.local/bin"
```

The uninstaller verifies hashes and removes only unmodified owned files.
Other user files are preserved. Inspect a modified or incomplete installation
before attempting removal again.

## Describe your slides

`presentation-kit --json schema` describes the input contract. It is not a
universal schema for arbitrary presentations. Save this example as `deck.json`:

```json
{
  "title": "Teaching project",
  "author": "Example author",
  "slides": [
    {
      "layout": "cover",
      "title": "Teaching project",
      "subtitle": "All data is invented"
    },
    {
      "layout": "bullets",
      "title": "Work sequence",
      "bullets": ["Choose one task", "Check the input", "Verify the result"]
    }
  ]
}
```

```bash
presentation-kit --json validate ./deck.json
presentation-kit --json build ./deck.json --output ./deck.pptx
presentation-kit --json inspect ./deck.pptx
```

The `slides` array supplies the order and count. Include a cover explicitly
when you want one. Top-level `title` and optional `author` describe the document;
they are not a prompt to create content about a subject.

### Layouts and limits

| Layout | Slide fields |
|---|---|
| `cover` | `title`, optional `subtitle` |
| `bullets` | `title`, `bullets`: 1–6 strings |
| `two_columns` | `title`, `left` and `right`; each column has a `heading` and 1–5 `bullets` strings |
| `table` | `title`, `columns`: 2–4 headers, `rows`: 1–6 rows with exactly that many cells |

Two-column example:

```json
{
  "layout": "two_columns",
  "title": "Before and after review",
  "left": {"heading": "Before", "bullets": ["Material collected", "Figures not yet checked"]},
  "right": {"heading": "After", "bullets": ["Sources read", "Errors marked"]}
}
```

Table example:

```json
{
  "layout": "table",
  "title": "Teaching stages",
  "columns": ["Stage", "Output"],
  "rows": [["Input", "Structure"], ["Build", "PPTX file"]]
}
```

At most 30 slides and 512 KiB input. Document title and author: at most 100
characters each; slide title: 70; cover subtitle: 180; a bullets-layout item:
180. Column heading: 45 characters; column bullet: 140. Table header: 35;
table cell: 100.

These character limits do not promise that every string of that length fits.
A separate conservative fit estimate can reject shorter content. Slides use
16:9 fixed layouts and refer to Arial by name; no font files are included.
The generator does not shrink or truncate text to make a build succeed.
For `TEXT_OVERFLOW`, revise the wording deliberately or split the material
across slides. Check the actual render in the intended editor afterward:
font substitution and rendering can differ.

Unknown fields, duplicate JSON keys, invalid UTF-8 and disallowed control
characters are refused. Output must not replace the input JSON; unsafe path
overlap and symlinks are rejected.

## What stays editable

The slides contain native text objects and, for `table`, a real table with
editable cells. They are not flattened slide images. One original simple theme
is supplied. Images, arbitrary templates, existing decks, charts, animation,
macros and external relationships are outside this version's contract.

You can continue editing the file in a compatible presentation application.
The tool does not promise preservation of another template, the entire
PowerPoint feature set or identical appearance under every font substitution.
It does not export PDF or slide images.

## JSON and verification

Each command prints one JSON object. `--json` makes this intent explicit.
Unicode on stdout uses JSON escaping; decoding preserves the original text.
Input files are read as UTF-8.

| Command | Main output |
|---|---|
| `doctor` | Version, bundled fixture availability, Python support, `mode: "offline"`, `auth: "not_required"` |
| `schema` | Descriptive input fields, layouts and limits |
| `validate` | Slide count, layouts, `estimated_text_fit: "pass"` |
| `build` / `demo` | Output path and bytes, `pptx_sha256`, `source_sha256`, slide count, `editable_text: true`, `editable_tables: true` |
| `inspect` | `format: "presentation-kit/1"`, slide count and decoded text in `outline`, `external_relationships: 0` |

`demo` also returns `synthetic: true`. Build output's
`text_fit: "conservative_estimate"` is an estimate, not a visual inspection.

Errors have the shape `{"ok": false, "error": {"code": "...", "message": "..."}}`.
Exit 0 means success; exit 2 means refusal or failure. Codes include
`INPUT_INVALID`, `INPUT_TOO_LARGE`, `TEXT_OVERFLOW`, `OUTPUT_EXISTS`,
`OUTPUT_COLLISION`, `UNSAFE_PATH`, `INVALID_PPTX`, `CLI_USAGE` and
`LOCAL_IO_ERROR`. Invalid input should not leave a new presentation; existing
output is preserved.

`inspect` is a bounded reader for the module's own generated format. It does
not verify claims or replace visual inspection. Check figures against their
sources before publishing the presentation.

## Prompt for Codex or Claude Code

> Read Presentation Kit's README.md and AGENTS.md. Run doctor, build the
> invented demo into a new PPTX and read inspect. For my deck, first save the
> approved wording and slide order in JSON without inventing figures. Run
> validate, then build a new output file. Compare the slide count and extracted
> text with the input. If content overflows, propose a meaningful revision or
> split instead of silently dropping it. Inspect the actual slides in an
> available editor and state where they were checked. Do not upload my content
> to an external service.

These are instructions for an agent that can run Python and read JSON. They
do not establish an installed MCP integration or certified compatibility with
every agent host.

## Privacy and rights

The CLI has no network or telemetry. Input JSON and PPTX contain supplied text,
data and author metadata; protect them like the original material. A cloud
agent may see data in its context, so a local build does not make the entire
agent workflow local. Input strings are data, not executable instructions.

Keep private decks out of commits and public issues. The software's MIT license
does not grant rights to arbitrary user content. The teaching fixture and
theme are original; no third-party templates or font software is bundled.
Tests: `python3 -B -m unittest discover -s tests -v`. A configured CI workflow
alone is not proof that all hosts passed.

## Use it in your work

Practical notes by Дмитрий Лукашов: [Telegram](https://t.me/lookatsoul).
To learn how to give an agent a task and check its result, see
[individual training](https://lookatshow.ru/executive?utm_source=github&utm_medium=open_module&utm_campaign=presentation-kit&utm_content=readme_en_learn).
For an implementation project, [discuss one process](https://lookatshow.ru/ai-implementation/?utm_source=github&utm_medium=open_module&utm_campaign=presentation-kit&utm_content=readme_en_implementation).
Development, timing and support are agreed separately from training.

The [shared catalog](https://github.com/ilookatshow-gif/lookatshow-open-tools)
connects standalone tools. When you first need evidence from documents, use
[Company Knowledge Kit](https://github.com/ilookatshow-gif/company-knowledge-kit).
This installation does not include the private core or connect external systems.
