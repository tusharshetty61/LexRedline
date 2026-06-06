require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const { AzureOpenAI } = require('openai');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'taskpane')));

// Request logger
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
  return new AzureOpenAI({
    endpoint: process.env.AZURE_OPENAI_ENDPOINT,
    apiKey: process.env.AZURE_OPENAI_API_KEY,
    apiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-12-01-preview',
    deployment: process.env.AZURE_OPENAI_DEPLOYMENT
  });
}

function validateEnv() {
  const required = [
    'AZURE_OPENAI_ENDPOINT',
    'AZURE_OPENAI_API_KEY',
    'AZURE_OPENAI_DEPLOYMENT'
  ];
  const missing = required.filter(k => !process.env[k]);
  if (missing.length > 0) {
    console.error(`\n✗ Missing required env vars: ${missing.join(', ')}`);
    console.error('  Copy .env.example to .env and fill in your Azure OpenAI details.\n');
    process.exit(1);
  }
}

// Single-turn endpoint (Call 1 and Call 2)
app.post('/api/analyze', async (req, res) => {
  const { systemPrompt, userContent, maxTokens } = req.body;

  try {
    const client = getClient();
    console.log(`  → Azure call /api/analyze (max_tokens: ${maxTokens || 3000})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model: process.env.AZURE_OPENAI_DEPLOYMENT,
      temperature: 0,
      max_tokens: maxTokens || 3000,
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
    console.error('Azure OpenAI error (/api/analyze):', error.message);
    res.status(500).json({ error: error.message });
  }
});

// Multi-turn endpoint (Call 3 — Ask AI conversation)
app.post('/api/chat', async (req, res) => {
  const { systemPrompt, messages, maxTokens } = req.body;

  try {
    const client = getClient();
    console.log(`  → Azure call /api/chat (${messages.length} messages, max_tokens: ${maxTokens || 3000})`);
    const t = Date.now();
    const response = await client.chat.completions.create({
      model: process.env.AZURE_OPENAI_DEPLOYMENT,
      temperature: 0,
      max_tokens: maxTokens || 3000,
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
    console.error('Azure OpenAI error (/api/chat):', error.message);
    res.status(500).json({ error: error.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    endpoint: process.env.AZURE_OPENAI_ENDPOINT,
    deployment: process.env.AZURE_OPENAI_DEPLOYMENT
  });
});

validateEnv();

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n✓ Contract Reviewer server running on http://localhost:${PORT}`);
  console.log(`  Azure endpoint : ${process.env.AZURE_OPENAI_ENDPOINT}`);
  console.log(`  Deployment     : ${process.env.AZURE_OPENAI_DEPLOYMENT}`);
  console.log(`  API version    : ${process.env.AZURE_OPENAI_API_VERSION || '2024-12-01-preview'}`);
  console.log(`\n  Health check   : http://localhost:${PORT}/health\n`);
});
