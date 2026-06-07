require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'taskpane')));

app.use((req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    const ms = Date.now() - start;
    const status = res.statusCode;
    const color = status >= 500 ? '\x1b[31m' : status >= 400 ? '\x1b[33m' : '\x1b[32m';
    console.log(`${color}${req.method} ${req.path} → ${status} (${ms}ms)\x1b[0m`);
  });
  next();
});

function getClient() {
  return new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
}

function validateEnv() {
  if (!process.env.ANTHROPIC_API_KEY) {
    console.error('\n✗ Missing ANTHROPIC_API_KEY in .env\n');
    process.exit(1);
  }
}

// Fast model (Haiku) for triage/classification; quality model (Sonnet) for suggestions
const FAST_MODEL    = () => process.env.CLAUDE_FAST_MODEL || 'claude-haiku-4-5-20251001';
const QUALITY_MODEL = () => process.env.CLAUDE_MODEL      || 'claude-sonnet-4-6';

app.post('/api/analyze', async (req, res) => {
  const { systemPrompt, userContent, maxTokens } = req.body;
  try {
    const client = getClient();
    const model  = FAST_MODEL();
    console.log(`  → /api/analyze [${model}] (max_tokens: ${maxTokens || 8192})`);
    const t = Date.now();
    const response = await client.messages.create({
      model,
      max_tokens: maxTokens || 8192,
      temperature: 0,
      system:     systemPrompt,
      messages:   [{ role: 'user', content: userContent }]
    });
    const text  = response.content[0].text;
    const usage = response.usage;
    console.log(`  ✓ Done in ${Date.now() - t}ms | tokens: ${usage.input_tokens} in / ${usage.output_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('Anthropic error (/api/analyze):', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/chat', async (req, res) => {
  const { systemPrompt, messages, maxTokens } = req.body;
  try {
    const client = getClient();
    const model  = QUALITY_MODEL();
    console.log(`  → /api/chat [${model}] (${messages.length} msgs, max_tokens: ${maxTokens || 8192})`);
    const t = Date.now();
    const response = await client.messages.create({
      model,
      max_tokens: maxTokens || 8192,
      temperature: 0,
      system:     systemPrompt,
      messages:   messages
    });
    const text  = response.content[0].text;
    const usage = response.usage;
    console.log(`  ✓ Done in ${Date.now() - t}ms | tokens: ${usage.input_tokens} in / ${usage.output_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('Anthropic error (/api/chat):', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', fast: FAST_MODEL(), quality: QUALITY_MODEL() });
});

validateEnv();

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n✓ Contract Reviewer server running on http://localhost:${PORT}`);
  console.log(`  Fast (triage)   : ${FAST_MODEL()}`);
  console.log(`  Quality (suggest): ${QUALITY_MODEL()}`);
  console.log(`\n  Health: http://localhost:${PORT}/health\n`);
});
