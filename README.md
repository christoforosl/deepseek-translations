# deepseek-translations

Translates Markdown (`.md`) files into English using the DeepSeek API.  
Input and output are always Markdown.

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
node dist/index.js <file-or-directory>
```

Output files are written to the **same directory** as the source file, with `-en` appended before the extension:

| Input | Output |
|-------|--------|
| `report.md` | `report-en.md` |

---

## Examples

### Translate a single file

```bash
node dist/index.js "C:\Docs\report.md"
```

### Translate all documents in a directory

```bash
node dist/index.js "C:\Users\christofjkl\Claude-Work\PROJECTS\PENSION-USER-STORIES"
```

**Batch mode skips any file that already has a `-en.md` counterpart in the same directory**, and skips any file already ending in `-en.md` (so translated outputs are never re-translated). Re-runs are safe — only new files are processed.

---

## Notes
- Markdown formatting (headings, tables, lists, bold/italic, code blocks) is preserved through the translation.
