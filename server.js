require('dotenv').config();
const express = require('express');
const cors    = require('cors');
const path    = require('path');
const OpenAI  = require('openai');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'taskpane')));

app.use((req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    const ms = Date.now() - start;
    const color = res.statusCode >= 500 ? '\x1b[31m' : res.statusCode >= 400 ? '\x1b[33m' : '\x1b[32m';
    console.log(`${color}${req.method} ${req.path} → ${res.statusCode} (${ms}ms)\x1b[0m`);
  });
  next();
});

function getClient() {
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

function validateEnv() {
  if (!process.env.OPENAI_API_KEY) {
    console.error('\n✗ Missing OPENAI_API_KEY in .env\n');
    process.exit(1);
  }
}

const MODEL = () => process.env.OPENAI_MODEL || 'gpt-4.1';

// /api/analyze — Call 1 (classify) and Call 2 (triage)
app.post('/api/analyze', async (req, res) => {
  const { systemPrompt, userContent, maxTokens } = req.body;
  try {
    const client = getClient();
    const tokens = maxTokens || 8192;
    console.log(`  → /api/analyze [${MODEL()}] (max_tokens: ${tokens})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model:           MODEL(),
      temperature:     0,
      max_tokens:      tokens,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user',   content: userContent  }
      ]
    });
    const text  = response.choices[0].message.content;
    const usage = response.usage;
    console.log(`  ✓ ${Date.now() - t}ms | ${usage.prompt_tokens} in / ${usage.completion_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('OpenAI error (/api/analyze):', error.message);
    res.status(500).json({ error: error.message });
  }
});

// /api/chat — Call 3 (suggestion) and Ask AI
app.post('/api/chat', async (req, res) => {
  const { systemPrompt, messages, maxTokens } = req.body;
  try {
    const client = getClient();
    const tokens = maxTokens || 8192;
    console.log(`  → /api/chat [${MODEL()}] (${messages.length} msgs, max_tokens: ${tokens})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model:           MODEL(),
      temperature:     0,
      max_tokens:      tokens,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        ...messages
      ]
    });
    const text  = response.choices[0].message.content;
    const usage = response.usage;
    console.log(`  ✓ ${Date.now() - t}ms | ${usage.prompt_tokens} in / ${usage.completion_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('OpenAI error (/api/chat):', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', model: MODEL() });
});

validateEnv();

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n✓ Contract Reviewer server running on http://localhost:${PORT}`);
  console.log(`  Model: ${MODEL()}`);
  console.log(`\n  Health: http://localhost:${PORT}/health\n`);
});
