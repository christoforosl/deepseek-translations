import * as fs from 'fs';
import * as path from 'path';
import OpenAI from 'openai';
import * as dotenv from 'dotenv';

dotenv.config();

// DeepSeek's context window is large (128K tokens), but we chunk defensively
// to handle very large documents and avoid single-request timeouts.
const MAX_CHUNK_CHARS = 12_000;

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

function parseArgs(argv: string[]): { input: string } {
  if (argv.length === 0) {
    console.error(
      'Usage:\n' +
      '  node dist/index.js <file.md>\n' +
      '  node dist/index.js <directory>\n' +
      '\n' +
      '  Produces a Markdown file: <name>-en.md\n' +
      '\n' +
      'When a directory is given, all .md files are translated.\n' +
      'Files whose output already exists in the same directory are skipped.'
    );
    process.exit(1);
  }

  return { input: argv[0] };
}

function outputPathFor(inputPath: string): string {
  const ext = path.extname(inputPath).toLowerCase();
  const dir = path.dirname(inputPath);
  const baseName = path.basename(inputPath, ext);
  return path.join(dir, `${baseName}-en.md`);
}

async function translateFile(inputPath: string, client: OpenAI): Promise<void> {
  const outputPath = outputPathFor(inputPath);

  console.log(`\n  Input : ${inputPath}`);
  console.log(`  Output: ${outputPath}`);

  console.log('  Reading content...');
  const markdown = fs.readFileSync(inputPath, 'utf-8');
  console.log(`  Read ${markdown.length.toLocaleString()} characters.`);

  const chunks = chunkMarkdown(markdown);
  console.log(`  Translating ${chunks.length} chunk(s)...`);

  const translatedChunks: string[] = [];
  for (let i = 0; i < chunks.length; i++) {
    const translated = await translateChunk(client, chunks[i], i, chunks.length);
    translatedChunks.push(translated);
  }

  const output = translatedChunks.join('\n\n');

  fs.writeFileSync(outputPath, output, 'utf-8');

  console.log(`  Written → ${outputPath}`);
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const { input } = parseArgs(args);

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
    const docFiles = allFiles.filter(f => /\.md$/i.test(f) && !/-en\.md$/i.test(f));

    if (docFiles.length === 0) {
      console.log('No .md files found in directory.');
      return;
    }

    const toProcess = docFiles.filter(f => {
      const base = path.basename(f, path.extname(f));
      const alreadyDone = fs.existsSync(path.join(inputPath, `${base}-en.md`));
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
      await translateFile(filePath, client);
    }

    console.log(`\nBatch complete. ${toProcess.length} file(s) translated.`);
    return;
  }

  // ── Single file mode ────────────────────────────────────────────────────────
  const ext = path.extname(inputPath).toLowerCase();
  if (ext !== '.md') {
    console.error('Error: input file must have a .md extension');
    process.exit(1);
  }

  await translateFile(inputPath, client);
}

main().catch((err: Error) => {
  console.error('Fatal error:', err.message ?? err);
  process.exit(1);
});
