# deepseek-translations

Translates Word documents (`.doc` / `.docx`) into English using the DeepSeek API.  
Output is either a Markdown file or a Word document.

---

## Setup

**1. Install dependencies**

```bash
npm install
```

**2. Configure your API key**

```bash
cp .env.example .env
```

Edit `.env` and paste your DeepSeek API key:

```
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxx
```

Get a key at [platform.deepseek.com](https://platform.deepseek.com).

**3. Build**

```bash
npm run build
```

---

## Usage

```
node dist/index.js <file-or-directory> [--output md|doc]
```

| Flag | Values | Default | Description |
|------|--------|---------|-------------|
| `--output` / `-o` | `md`, `doc` | `md` | Output format: Markdown or Word (`.docx`) |

Output files are written to the **same directory** as the source file, with `-en` appended before the extension:

| Input | Output (`--output md`) | Output (`--output doc`) |
|-------|------------------------|-------------------------|
| `report.docx` | `report-en.md` | `report-en.docx` |
| `SPF-001.doc` | `SPF-001-en.md` | `SPF-001-en.docx` |

---

## Examples

### Translate a single file to Markdown (default)

```bash
node dist/index.js "C:\Docs\report.docx"
```

### Translate a single file to Word

```bash
node dist/index.js "C:\Docs\report.docx" --output doc
```

### Translate a single `.doc` file (legacy format)

```bash
node dist/index.js "C:\Docs\contract.doc" -o doc
```

> Requires Microsoft Word to be installed. The tool converts `.doc` → `.docx` via Word COM automation before processing.

### Translate all documents in a directory to Markdown

```bash
node dist/index.js "C:\Users\christoforosl\Claude-Work\PROJECTS\PENSION-USER-STORIES"
```

### Translate all documents in a directory to Word

```bash
node dist/index.js "C:\Users\christoforosl\Claude-Work\PROJECTS\PENSION-USER-STORIES" --output doc
```

**Batch mode skips any file that already has a `-en.md` or `-en.docx` counterpart in the same directory**, regardless of which `--output` format is requested. Re-runs are safe — only new files are processed.

---

## Notes

- Large documents are automatically split into chunks and reassembled after translation.
- Markdown formatting (headings, tables, lists, bold/italic, code blocks) is preserved through the translation.
- `.doc` files require Microsoft Word to be installed (Windows only). `.docx` files work on any platform.
