require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const { OpenAI } = require('openai');

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
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

function validateEnv() {
  if (!process.env.OPENAI_API_KEY) {
    console.error('\n✗ Missing required env var: OPENAI_API_KEY');
    console.error('  Copy .env.example to .env and fill in your OpenAI API key.\n');
    process.exit(1);
  }
}

const MODEL = process.env.OPENAI_MODEL || 'gpt-4.1';

app.post('/api/analyze', async (req, res) => {
  const { systemPrompt, userContent, maxTokens } = req.body;
  try {
    const client = getClient();
    console.log(`  → OpenAI call /api/analyze (max_tokens: ${maxTokens || 4000})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model: MODEL,
      temperature: 0,
      max_tokens: maxTokens || 4000,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user',   content: userContent  }
      ]
    });
    const text = response.choices[0].message.content;
    const usage = response.usage;
    console.log(`  ✓ Done in ${Date.now() - t}ms | tokens: ${usage.prompt_tokens} in / ${usage.completion_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('OpenAI error (/api/analyze):', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/chat', async (req, res) => {
  const { systemPrompt, messages, maxTokens } = req.body;
  try {
    const client = getClient();
    console.log(`  → OpenAI call /api/chat (${messages.length} messages, max_tokens: ${maxTokens || 4000})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model: MODEL,
      temperature: 0,
      max_tokens: maxTokens || 4000,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        ...messages
      ]
    });
    const text = response.choices[0].message.content;
    const usage = response.usage;
    console.log(`  ✓ Done in ${Date.now() - t}ms | tokens: ${usage.prompt_tokens} in / ${usage.completion_tokens} out`);
    res.json({ text });
  } catch (error) {
    console.error('OpenAI error (/api/chat):', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', model: MODEL });
});

validateEnv();

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n✓ Contract Reviewer server running on http://localhost:${PORT}`);
  console.log(`  Model          : ${MODEL}`);
  console.log(`\n  Health check   : http://localhost:${PORT}/health\n`);
});
