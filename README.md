# AI Contract Reviewer — Indian Law Edition

A Microsoft Word task-pane add-in that performs multi-pass AI review of contracts under Indian law and writes accepted suggestions directly into the document as **Word tracked changes**.

---

## What it does

Open any contract in Word, click **Review Contract** in the ribbon, and the add-in:

1. **Classifies** the agreement type (sale, service, employment, NDA, lease, etc.) and determines the applicable Indian statutes.
2. **Runs a 5-pass triage** across the full document:
   - Statutory conflicts (ICA, SOGA, IT Act, RERA, labour law, etc.)
   - Structural inconsistencies (cross-references, definitions, schedules)
   - Missing mandatory clauses
   - Material commercial risk
   - Drafting defects (ambiguity, unenforceability, vague language)
3. **Presents each issue** as a clause card with:
   - Current text (highlighted in the document)
   - Suggested replacement with a negotiation note and fallback position
   - Confidence score and the statute / legal basis
   - Inline "Ask AI" for follow-up questions on that specific clause
4. **Accept** a suggestion → the change is written to the document immediately as a Word tracked change (red-line), visible in the Review tab.
5. **Reject** → the highlight is cleared and you move on.
6. At the end a **summary screen** lists every clause, its priority, and your decision.

All analysis is done by GPT-4.1 at temperature 0 via the OpenAI API.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Microsoft Word | Desktop (Windows or Mac). Word Online is not supported. |
| OpenAI API key | [platform.openai.com](https://platform.openai.com) |
| Node.js 18+ | Only needed if you self-host the backend |

---

## Option A — Use the hosted version (quickest)

The backend is already deployed. All you need to do is sideload `manifest.xml` into Word — no code, no server, no API key required on your machine.

### Step 1 — Download the manifest

Download `manifest.xml` from this repository.

### Step 2 — Sideload into Word

**Windows**

1. Open Word and open any document.
2. Go to **Insert** → **Add-ins** → **My Add-ins**.
3. Click **Upload My Add-in** (top-right of the dialog).
4. Browse to and select `manifest.xml`.
5. Click **Upload**.
6. The **"Contract Review"** group appears in the **Home** tab ribbon.
7. Click **Review Contract** to open the task pane.

**Mac**

1. Open Word and open any document.
2. Go to **Insert** → **Add-ins**.
3. Click **Upload My Add-in**.
4. Select `manifest.xml` and click **Open**.
5. Click **Review Contract** in the Home ribbon.

> The hosted backend handles the OpenAI calls. The add-in reads and writes to your Word document locally — document content is sent to the AI for analysis.

---

## Option B — Self-host your own backend

Use this if you want your own OpenAI key, want to modify the prompts, or want full control over the deployment.

### 1. Clone and install

```bash
git clone https://github.com/your-username/contract-reviewer.git
cd contract-reviewer
npm install
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env`:

```
OPENAI_API_KEY=sk-proj-...
OPENAI_MODEL=gpt-4.1
PORT=3001
```

### 3. Run locally (for testing)

```bash
npm start
# Server runs at http://localhost:3001
```

To use locally you also need to update `manifest.xml` — replace every occurrence of the Railway URL with `http://localhost:3001` — and then sideload that edited manifest into Word.

### 4. Deploy to Railway (recommended)

1. Push to GitHub.
2. Create a new project at [railway.app](https://railway.app), connect your repo.
3. Set the `OPENAI_API_KEY` and `OPENAI_MODEL` environment variables in the Railway dashboard.
4. Railway provides a public HTTPS URL, e.g. `https://your-app.up.railway.app`.

### 5. Update the manifest with your URL

In `manifest.xml`, replace every instance of:

```
https://contract-reviewer-production.up.railway.app
```

with your own deployment URL, e.g.:

```
https://your-app.up.railway.app
```

There are **8 places** in the file where the URL appears.

### 6. Sideload your updated manifest into Word

Follow the same sideloading steps in Option A above, but use your edited `manifest.xml`.

---

## Project structure

```
contract-reviewer/
├── server.js          # Express backend — three-call OpenAI pipeline
├── manifest.xml       # Word add-in manifest (points to deployed backend)
├── .env.example       # Environment variable template
└── taskpane/
    ├── index.html     # Task pane UI
    ├── app.js         # All client-side logic (Word JS API, session state)
    └── styles.css     # Orange / purple / white theme
```

---

## Environment variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `OPENAI_API_KEY` | Yes | — | Your OpenAI secret key |
| `OPENAI_MODEL` | No | `gpt-4.1` | Any OpenAI chat model |
| `PORT` | No | `3001` | Port the Express server listens on |

---

## Re-running after accepting changes

The AI runs against the document text at the time of the review. If you accept several suggestions and want to check the revised contract:

1. In Word, go to the **Review** tab and **Accept All Changes** to commit the tracked changes.
2. Click **Review Contract** again to start a fresh 5-pass review on the updated text.

---

## Notes

- The add-in requires `ReadWriteDocument` permission to highlight text and write tracked changes.
- Document text is sent to OpenAI for analysis. Do not use this on documents that must not leave your network without hosting the backend in a private environment.
- The defined manifest ID (`a1b2c3d4-e5f6-7890-abcd-ef1234567890`) is a placeholder GUID. If you publish multiple versions or submit to AppSource, generate a unique GUID for each.
