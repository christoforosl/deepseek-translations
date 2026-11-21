import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { spawnSync } from 'child_process';
import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { marked } from 'marked';
import HTMLtoDOCX from 'html-to-docx';
import OpenAI from 'openai';
import * as dotenv from 'dotenv';

dotenv.config();

// DeepSeek's context window is large (128K tokens), but we chunk defensively
// to handle very large documents and avoid single-request timeouts.
const MAX_CHUNK_CHARS = 12_000;

/**
 * Convert a legacy .doc file to .docx using Microsoft Word via PowerShell COM.
 * Paths are passed as environment variables to avoid encoding issues with
 * non-ASCII filenames. The PS1 script itself uses only ASCII.
 * Returns the path to a temporary .docx file that the caller must delete when done.
 */
function convertDocToDocx(docPath: string): string {
  const tempDocx = path.join(os.tmpdir(), `deepseek_trans_${Date.now()}.docx`);

  // Paths come from env vars — no Unicode in the script source itself.
  const psScript = `
$ErrorActionPreference = 'Stop'
$inputPath  = $env:DEEPSEEK_DOC_INPUT
$outputPath = $env:DEEPSEEK_DOC_OUTPUT
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($inputPath)
    $doc.SaveAs2($outputPath, 16)
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
`;

  const psFile = path.join(os.tmpdir(), `deepseek_conv_${Date.now()}.ps1`);
  // Write as UTF-16 LE with BOM so PowerShell 5.1 reads it correctly.
  fs.writeFileSync(psFile, Buffer.from('\uFEFF' + psScript, 'utf16le'));

  try {
    const result = spawnSync(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', psFile],
      {
        timeout: 60_000,
        encoding: 'utf-8',
        env: { ...process.env, DEEPSEEK_DOC_INPUT: docPath, DEEPSEEK_DOC_OUTPUT: tempDocx },
      }
    );

    if (result.status !== 0) {
      const detail = (result.stderr ?? result.stdout ?? '').trim();
      throw new Error(
        `Word COM conversion failed.\n${detail}\n\nTip: open the file in Word, save as .docx, then re-run.`
      );
    }
  } finally {
    if (fs.existsSync(psFile)) fs.unlinkSync(psFile);
  }

  if (!fs.existsSync(tempDocx)) {
    throw new Error('Word did not produce a .docx output. Is Microsoft Word installed?');
  }

  return tempDocx;
}

async function extractMarkdownFromDocx(filePath: string): Promise<string> {
  const result = await mammoth.convertToHtml({ path: filePath });

  if (result.messages.length > 0) {
    for (const msg of result.messages) {
      if (msg.type === 'warning') {
        console.warn(`  [docx warning] ${msg.message}`);
      }
    }
  }

  const turndown = new TurndownService({
    headingStyle: 'atx',
    bulletListMarker: '-',
    codeBlockStyle: 'fenced',
  });

  return turndown.turndown(result.value);
}

/**
 * Split markdown into chunks at paragraph boundaries so each chunk stays
 * below MAX_CHUNK_CHARS while never cutting mid-paragraph.
 */
function chunkMarkdown(text: string): string[] {
  const paragraphs = text.split(/\n{2,}/);
  const chunks: string[] = [];
  let current = '';

  for (const para of paragraphs) {
    const candidate = current ? `${current}\n\n${para}` : para;
    if (candidate.length > MAX_CHUNK_CHARS && current.length > 0) {
      chunks.push(current.trim());
      current = para;
    } else {
      current = candidate;
    }
  }

  if (current.trim()) {
    chunks.push(current.trim());
  }

  return chunks;
}

async function translateChunk(client: OpenAI, chunk: string, chunkIndex: number, total: number): Promise<string> {
  process.stdout.write(`  Translating chunk ${chunkIndex + 1}/${total}...`);

  const response = await client.chat.completions.create({
    model: 'deepseek-chat',
    messages: [
      {
        role: 'system',
        content: [
          'You are a professional translator.',
          'Translate the user-provided Markdown text into English.',
          'Rules:',
          '- Preserve ALL Markdown formatting exactly: headings (#, ##, …), bold (**), italic (*), lists (-, 1.), tables, code blocks (```), inline code (`), blockquotes (>), links, and horizontal rules.',
          '- Translate only the human-readable text; do not alter URLs, code, or Markdown syntax tokens.',
          '- Return only the translated Markdown. No preamble, no explanation, no triple-backtick wrapper around the entire response.',
        ].join('\n'),
      },
      {
        role: 'user',
        content: chunk,
      },
    ],
    temperature: 0.2,
  });

  const translated = response.choices[0].message.content ?? '';
  console.log(` done (${translated.length} chars)`);
  return translated;
}

type OutputFormat = 'md' | 'doc';

function parseArgs(argv: string[]): { input: string; outputFormat: OutputFormat } {
  const positional: string[] = [];
  let outputFormat: OutputFormat = 'md';

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--output' || arg === '-o') {
      const val = argv[++i];
      if (val !== 'md' && val !== 'doc') {
        console.error(`Error: --output must be 'md' or 'doc', got '${val}'`);
        process.exit(1);
      }
      outputFormat = val;
    } else if (arg.startsWith('-o=') || arg.startsWith('--output=')) {
      const val = arg.split('=')[1];
      if (val !== 'md' && val !== 'doc') {
        console.error(`Error: --output must be 'md' or 'doc', got '${val}'`);
        process.exit(1);
      }
      outputFormat = val as OutputFormat;
    } else {
      positional.push(arg);
    }
  }

  if (positional.length === 0) {
    console.error(
      'Usage:\n' +
      '  node dist/index.js <file.doc|docx|md> [--output md|doc]\n' +
      '  node dist/index.js <directory>        [--output md|doc]\n' +
      '\n' +
      '  --output md   Produce a Markdown file: <name>-en.md  (default)\n' +
      '  --output doc  Produce a Word document: <name>-en.docx\n' +
      '\n' +
      'When a directory is given, all .doc/.docx/.md files are translated.\n' +
      'Files whose output already exists in the same directory are skipped.'
    );
    process.exit(1);
  }

  return { input: positional[0], outputFormat };
}

async function writeWordDocument(markdown: string, outputPath: string): Promise<void> {
  const html = await marked(markdown);
  // html-to-docx returns a Buffer
  const result = await HTMLtoDOCX(html, undefined, {
    table: { row: { cantSplit: true } },
    title: path.basename(outputPath, '.docx'),
  }) as unknown as Buffer;
  fs.writeFileSync(outputPath, result);
}

function outputPathFor(inputPath: string, outputFormat: OutputFormat): string {
  const ext = path.extname(inputPath).toLowerCase();
  const dir = path.dirname(inputPath);
  const baseName = path.basename(inputPath, ext);
  const outputExt = outputFormat === 'doc' ? '.docx' : '.md';
  return path.join(dir, `${baseName}-en${outputExt}`);
}

async function translateFile(inputPath: string, outputFormat: OutputFormat, client: OpenAI): Promise<void> {
  const ext = path.extname(inputPath).toLowerCase();
  const outputPath = outputPathFor(inputPath, outputFormat);

  console.log(`\n  Input : ${inputPath}`);
  console.log(`  Output: ${outputPath}`);

  let workingPath = inputPath;
  let tempDocx: string | null = null;
  if (ext === '.doc') {
    console.log('  Converting .doc → .docx via Microsoft Word...');
    tempDocx = convertDocToDocx(inputPath);
    workingPath = tempDocx;
  }

  try {
    console.log('  Extracting content...');
    const markdown = ext === '.md'
      ? fs.readFileSync(inputPath, 'utf-8')
      : await extractMarkdownFromDocx(workingPath);
    console.log(`  Extracted ${markdown.length.toLocaleString()} characters.`);

    const chunks = chunkMarkdown(markdown);
    console.log(`  Translating ${chunks.length} chunk(s)...`);

    const translatedChunks: string[] = [];
    for (let i = 0; i < chunks.length; i++) {
      const translated = await translateChunk(client, chunks[i], i, chunks.length);
      translatedChunks.push(translated);
    }

    const output = translatedChunks.join('\n\n');

    if (outputFormat === 'doc') {
      await writeWordDocument(output, outputPath);
    } else {
      fs.writeFileSync(outputPath, output, 'utf-8');
    }

    console.log(`  Written → ${outputPath}`);
  } finally {
    if (tempDocx && fs.existsSync(tempDocx)) fs.unlinkSync(tempDocx);
  }
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const { input, outputFormat } = parseArgs(args);

  const inputPath = path.resolve(input);

  if (!fs.existsSync(inputPath)) {
    console.error(`Error: path not found — ${inputPath}`);
    process.exit(1);
  }

  const apiKey = process.env.DEEPSEEK_API_KEY;
  if (!apiKey || apiKey === 'your_api_key_here') {
    console.error('Error: DEEPSEEK_API_KEY is not set. Copy .env.example to .env and add your key.');
    process.exit(1);
  }

  const client = new OpenAI({ apiKey, baseURL: 'https://api.deepseek.com' });

  const stat = fs.statSync(inputPath);

  // ── Directory mode ──────────────────────────────────────────────────────────
  if (stat.isDirectory()) {
    const allFiles = fs.readdirSync(inputPath);
    const docFiles = allFiles.filter(f => /\.(doc|docx|md)$/i.test(f));

    if (docFiles.length === 0) {
      console.log('No .doc/.docx/.md files found in directory.');
      return;
    }

    const toProcess = docFiles.filter(f => {
      const base = path.basename(f, path.extname(f));
      const alreadyDone = ['.md', '.docx'].some(outExt =>
        fs.existsSync(path.join(inputPath, `${base}-en${outExt}`))
      );
      if (alreadyDone) console.log(`  Skipping (already translated): ${f}`);
      return !alreadyDone;
    });

    const skipped = docFiles.length - toProcess.length;
    console.log(`Found ${docFiles.length} document(s). ${skipped} already translated, ${toProcess.length} to process.`);

    if (toProcess.length === 0) {
      console.log('All documents already translated. Nothing to do.');
      return;
    }

    for (let i = 0; i < toProcess.length; i++) {
      const filePath = path.join(inputPath, toProcess[i]);
      console.log(`\n[${i + 1}/${toProcess.length}] ${toProcess[i]}`);
      await translateFile(filePath, outputFormat, client);
    }

    console.log(`\nBatch complete. ${toProcess.length} file(s) translated.`);
    return;
  }

  // ── Single file mode ────────────────────────────────────────────────────────
  const ext = path.extname(inputPath).toLowerCase();
  if (ext !== '.docx' && ext !== '.doc' && ext !== '.md') {
    console.error('Error: input file must have a .doc, .docx, or .md extension');
    process.exit(1);
  }

  await translateFile(inputPath, outputFormat, client);
}

main().catch((err: Error) => {
  console.error('Fatal error:', err.message ?? err);
  process.exit(1);
});
