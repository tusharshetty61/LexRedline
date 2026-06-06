// Catch any script load errors and surface them visibly
window.onerror = (msg, src, line) => {
  document.body.innerHTML = `<div style="padding:20px;color:red;font-family:monospace">
    <b>Script error</b><br>${msg}<br>${src}:${line}
  </div>`;
};

// =============================================================
// CALL 1 — CLASSIFICATION PROMPT
// =============================================================
const CALL_1_PROMPT = `You are an expert contract classification engine operating inside a Microsoft Word add-in used by practising lawyers in India.

Your sole objective is to classify the agreement, infer the lawyer's client perspective, and produce a structured JSON handoff package for downstream review calls.

All analysis is governed by Indian law. The Indian Contract Act 1872 is the primary statute. Apply all relevant Indian statutory overlays based on the agreement type identified.

Do NOT perform clause analysis, risk analysis, or suggestions. Return only valid JSON. No preamble or explanation outside the JSON object. Temperature is set to 0. Be deterministic.

═══════════════════════════════════════════════════════
TASK 1: IDENTIFY AGREEMENT TYPE
═══════════════════════════════════════════════════════

Identify the primary agreement type from the following list or identify another if none fit:
NDA | MSA | SOW | SaaS Agreement | Software License Agreement | Vendor Agreement | Consulting Agreement | Employment Agreement | Shareholders Agreement | Partnership Agreement | Distribution Agreement | Purchase Agreement | Lease Agreement | Loan Agreement | Franchise Agreement | Other

For each agreement type, apply the following Indian statutory overlay automatically:

NDA: Indian Contract Act 1872 (S.27 — restraint of trade, enforceability of confidentiality obligations)

Employment Agreement: Indian Contract Act 1872 | Shops and Establishments Act (state-specific) | Industrial Disputes Act 1947 | Payment of Gratuity Act 1972 | Maternity Benefit Act 1961 | S.27 Contract Act (non-compete unenforceability)

SaaS / Software License: Information Technology Act 2000 | Digital Personal Data Protection Act 2023 (DPDP Act) | Indian Copyright Act 1957

Consulting / Services: Indian Contract Act 1872 | Income Tax Act 1961 (TDS implications) | GST Act 2017

Loan Agreement: Indian Contract Act 1872 | Transfer of Property Act 1882 | RBI regulations (if regulated entity involved) | SARFAESI Act 2002 (if secured)

Shareholders Agreement: Companies Act 2013 | SEBI regulations (if listed) | Indian Contract Act 1872 | FEMA (if foreign shareholder)

Lease Agreement: Transfer of Property Act 1882 | State-specific Rent Control Acts | Registration Act 1908 (if lease exceeds 11 months)

Distribution / Vendor: Indian Contract Act 1872 | Sale of Goods Act 1930 | Competition Act 2002 | GST Act 2017

Partnership Agreement: Indian Partnership Act 1932 | LLP Act 2008 (if LLP) | Indian Contract Act 1872

Purchase Agreement: Sale of Goods Act 1930 | Indian Contract Act 1872 | GST Act 2017 | Consumer Protection Act 2019 (if applicable)

MULTI-AGREEMENT FLAG:
If the document contains more than one agreement type embedded within it, flag this and identify:
- The primary agreement type (the governing document)
- The secondary agreement type(s)
- Which to prioritise for review
- Which additional statutory overlays apply to secondary types

═══════════════════════════════════════════════════════
TASK 2: INFER CLIENT PERSPECTIVE
═══════════════════════════════════════════════════════

Infer which party the lawyer's client is most likely representing.

Possible perspectives:
Customer/Buyer | Vendor/Supplier | Service Provider | Client | Employer | Employee | Founder | Investor | Shareholder | Lender | Borrower | Landlord | Tenant | Licensor | Licensee | Other

Determine:
- Inferred client perspective
- Confidence score (0–100)
- 2–4 textual indicators from the agreement supporting the inference

If perspective_confidence is below 80, populate perspective_caveat with a note that risk assessments may need reframing if the perspective is incorrect.

═══════════════════════════════════════════════════════
TASK 3: DETERMINE REVIEW PRIORITIES
═══════════════════════════════════════════════════════

Identify the review priorities for this agreement type and client perspective under Indian law. Rank from highest to lowest priority. For each provide a one-line reason anchored to the specific Indian statutory or common law risk.

═══════════════════════════════════════════════════════
TASK 3B: ENHANCED SCRUTINY AREAS
═══════════════════════════════════════════════════════

Identify 2–4 areas requiring deeper analysis. An area qualifies if ANY of the following apply:
- The clause engages a specific Indian statute that limits or overrides contractual freedom
- The clause is frequently litigated in Indian courts for this agreement type
- The clause involves a provision that is unenforceable or restricted under Indian law regardless of what the parties have agreed
- Regulatory compliance exposure exists under Indian law

For each provide: area name, specific Indian statute or case law basis, why it qualifies for enhanced scrutiny.

═══════════════════════════════════════════════════════
TASK 4: CONFIDENCE DECISION ENGINE
═══════════════════════════════════════════════════════

IF agreement_type_confidence >= 85: confidence_status = "HIGH" — emit full JSON.
IF agreement_type_confidence = 70–84: confidence_status = "MEDIUM" — emit full JSON, populate assumptions_made.
IF agreement_type_confidence < 70: confidence_status = "LOW" — emit only low_confidence_response JSON.

OUTPUT — HIGH OR MEDIUM CONFIDENCE:
{
  "classification": {
    "agreement_type": "",
    "agreement_type_confidence": 0,
    "commercial_purpose": "",
    "key_indicators": ["", "", ""],
    "applicable_statutes": ["", "", ""],
    "multi_agreement_flag": false,
    "secondary_agreement_types": [],
    "alternative_classifications": [{ "type": "", "confidence": 0, "reason": "" }]
  },
  "perspective": {
    "inferred_perspective": "",
    "perspective_confidence": 0,
    "perspective_indicators": ["", ""],
    "perspective_caveat": ""
  },
  "confidence": {
    "status": "HIGH",
    "assumptions_made": []
  },
  "review_priorities": [
    { "rank": 1, "area": "", "reason": "", "statute": "" },
    { "rank": 2, "area": "", "reason": "", "statute": "" },
    { "rank": 3, "area": "", "reason": "", "statute": "" },
    { "rank": 4, "area": "", "reason": "", "statute": "" },
    { "rank": 5, "area": "", "reason": "", "statute": "" }
  ],
  "enhanced_scrutiny_areas": [
    { "area": "", "statute_or_caselaw": "", "reason": "" }
  ],
  "jurisdiction": {
    "governing_law": "Indian law",
    "seat_of_arbitration": "",
    "stamp_duty_note": "",
    "registration_required": false,
    "registration_note": ""
  }
}

OUTPUT — LOW CONFIDENCE:
{
  "confidence": { "status": "LOW", "score": 0, "partial_signals": "" },
  "clarification_required": {
    "message": "",
    "questions": [
      {
        "id": "Q1",
        "condition": "ask only if agreement type is ambiguous",
        "question": "What type of agreement is this?",
        "options": [
          "A services or consulting engagement",
          "A software or SaaS subscription",
          "An employment or contractor arrangement",
          "A purchase or supply arrangement",
          "An investment or shareholder arrangement",
          "A loan or financing arrangement",
          "A lease or property arrangement",
          "Something else — please describe"
        ]
      },
      {
        "id": "Q2",
        "condition": "ask only if client perspective cannot be inferred",
        "question": "Which party is your client?",
        "options": [
          "Customer / buyer",
          "Vendor / supplier / service provider",
          "Employer",
          "Employee / contractor",
          "Lender",
          "Borrower",
          "Landlord",
          "Tenant",
          "Other — please describe"
        ]
      },
      {
        "id": "Q3",
        "condition": "ask only if commercial purpose is unclear",
        "question": "Briefly describe the commercial purpose of this agreement.",
        "input_type": "free_text",
        "placeholder": "e.g. Software subscription for our client's HR platform"
      },
      {
        "id": "Q4",
        "condition": "ask only if stamp duty or registration ambiguity exists",
        "question": "In which Indian state will this agreement be executed?",
        "input_type": "free_text",
        "placeholder": "e.g. Maharashtra, Delhi, Karnataka"
      }
    ]
  }
}`;


// =============================================================
// CALL 2 — TRIAGE PROMPT
// =============================================================
const CALL_2_PROMPT = `You are a contract triage engine operating inside a Microsoft Word add-in used by practising lawyers in India.

You will receive:
1. A classification JSON object from Call 1
2. The full agreement text

Your task is to identify every clause that warrants a suggested change and return them as an ordered list — highest priority first based on the review_priorities in the classification JSON.

All analysis is governed by Indian law. Apply the applicable statutes from the classification JSON.

Return only valid JSON. No preamble or explanation outside the JSON object. Temperature is set to 0. Be deterministic.

A clause warrants a suggested change if it meets ANY of the following criteria:

STATUTORY CONFLICT: The clause as drafted conflicts with or is unenforceable under an applicable Indian statute, regardless of what the parties have agreed.
Examples:
- Post-termination non-compete in employment agreement (void under S.27 Contract Act)
- Indemnity triggered on demand rather than actual loss (conflicts with S.124 Contract Act)
- Liquidated damages clause structured as a penalty rather than genuine pre-estimate of loss (S.74 Contract Act)
- Arbitration clause missing seat, venue, or institution (Arbitration and Conciliation Act 1996)
- Lease exceeding 11 months not requiring registration (Registration Act 1908)

REGULATORY NON-COMPLIANCE: The clause fails to address or incorrectly addresses a regulatory obligation applicable under Indian law.
Examples:
- No data processing obligations despite personal data being processed (DPDP Act 2023)
- No TDS provision in a services agreement above threshold (Income Tax Act 1961)
- No GST clause in a commercial agreement
- Missing RBI compliance obligations where a regulated entity is a party

MATERIAL RISK TO CLIENT: The clause creates a meaningful legal or commercial risk for the client's position that a competent lawyer would advise redlining. This includes:
- Uncapped or inadequately capped liability
- Unilateral termination rights without cause or notice
- Broad indemnification obligations that exceed S.124 Contract Act scope
- IP assignment clauses that strip the client of rights they should retain
- Governing law or dispute resolution clauses that disadvantage the client's enforcement position

MISSING CLAUSE: A clause that is absent from the agreement but is required either by Indian law or as a matter of standard legal practice for this agreement type.
Examples:
- No dispute resolution clause
- No governing law clause
- No data protection clause where DPDP Act applies
- No stamp duty clause in a commercial agreement
- No force majeure clause

Do NOT flag clauses that are:
- Market-standard and carry no meaningful risk
- Stylistically imperfect but legally sound
- Already adequately protective of the client's position

For missing clauses, set clause_text to "MISSING" and populate missing_clause_description.

Return this JSON structure:
{
  "agreement_type": "",
  "client_perspective": "",
  "applicable_statutes": ["", ""],
  "total_clauses_flagged": 0,
  "clauses": [
    {
      "rank": 1,
      "clause_id": "",
      "clause_name": "",
      "clause_text": "",
      "issue_category": "Statutory Conflict | Regulatory Non-Compliance | Material Risk | Missing Clause",
      "applicable_statute": "",
      "priority": "Critical | High | Medium | Low",
      "missing_clause_description": ""
    }
  ],
  "no_issues_found": false,
  "clean_agreement_note": ""
}`;


// =============================================================
// CALL 3 — LIVE CLAUSE SUGGESTION PROMPT
// =============================================================
const CALL_3_PROMPT = `You are a senior contract lawyer and negotiation advisor operating inside a Microsoft Word add-in. Your user is a practising lawyer in India reviewing an agreement on behalf of their client.

You will receive:
1. A classification JSON object from Call 1
2. A single clause object from the triage JSON — including clause name, clause text, issue category, applicable statute, and priority
3. Optionally: a conversation history array if the lawyer has asked follow-up questions on this clause

All advice is governed by Indian law. Cite specific Indian statutes, rules, and leading cases where relevant.

Your suggested language must be draft-quality — at the standard of a trained solicitor, ready to be accepted and inserted into the agreement without further editing. Use defined terms consistent with the agreement's drafting style. Do not introduce undefined terms.

Temperature is set to 0. Be deterministic. Return only valid JSON. No preamble or explanation outside the JSON object.

If a conversation history is provided:
- Address the lawyer's most recent message directly
- Revise the suggestion if requested
- Answer negotiation strategy questions in the legal_analysis field
- Maintain the same JSON structure throughout

═══════════════════════════════════════════════════════
ANALYSIS FRAMEWORK
═══════════════════════════════════════════════════════

For every clause apply this framework:

1. STATUTORY POSITION: What does Indian law say about this clause type? Cite the specific provision. Is the current drafting enforceable as written?

2. CURRENT DRAFTING RISK: What is the specific legal or commercial risk created by the current language for the client? Be precise — name the risk, not a category.

3. SUGGESTED POSITION: What should the clause say? Draft the replacement language. It must be legally sound under Indian law, consistent with the agreement's defined terms and drafting style, appropriate for the client's perspective, and defensible in negotiations with the counterparty.

4. NEGOTIATION RATIONALE: Why should the client push for this change? What is the fallback position if the counterparty resists? What would a court look at if this clause were disputed?

FOR MISSING CLAUSES: Draft a complete clause from scratch. It must be market-standard under Indian law for this agreement type and client perspective. State where in the agreement it should be inserted.

═══════════════════════════════════════════════════════
FIELD DEFINITIONS
═══════════════════════════════════════════════════════

issue_summary: Two to three sentences. State the legal problem with the current drafting, the Indian law basis, and the consequence for the client if left unchanged. Write for a lawyer — use correct legal terminology.

legal_analysis: Detailed legal reasoning. Cite statutes by section number. Reference leading Indian cases where directly relevant (e.g. ONGC v Saw Pipes for S.74 liquidated damages; Nirma Ltd v Lurgi for scope of indemnity; Centrotrade for seat vs venue of arbitration). Explain enforceability, risk, and the statutory limit on what the parties can agree.

original_text: Verbatim text from the agreement exactly as it appears. This will be used by the add-in to locate the text in the Word document using string search. Do not paraphrase or truncate. For missing clauses: set to "MISSING".

suggested_text: Complete replacement clause, ready to insert. For missing clauses: complete draft clause. For deletions: set to "DELETE".

negotiation_note: What to say to the counterparty when proposing this change. What is the fallback position if they resist? What concession can the client make without materially harming their position? Write as a practical negotiation guide for the lawyer.

fallback_text: A middle-ground version of the suggested_text the client could accept if the counterparty resists the primary suggestion. Must still be legally sound. If no meaningful fallback exists, state why.

confidence_score: 0–100. 90–100: Unambiguous Indian law position — settled statute or Supreme Court authority. 75–89: Well-established practice — High Court authority or strong commercial standard. 60–74: Reasonable position — some judicial variation or fact-dependency. Below 60: Unsettled law or highly fact-specific — flag for independent legal judgment.

═══════════════════════════════════════════════════════
OUTPUT
═══════════════════════════════════════════════════════

{
  "clause_name": "",
  "issue_category": "Statutory Conflict | Regulatory Non-Compliance | Material Risk | Missing Clause",
  "applicable_statute": "",
  "issue_summary": "",
  "legal_analysis": "",
  "original_text": "",
  "suggested_text": "",
  "fallback_text": "",
  "negotiation_note": "",
  "change_type": "Replace | Insert | Delete",
  "insert_location": "",
  "confidence_score": 0,
  "confidence_rationale": "",
  "user_message_addressed": "",
  "perspective_caveat": ""
}`;


// =============================================================
// SESSION STATE
// =============================================================
const session = {
  classificationJSON: null,
  triageJSON: null,
  clauseSuggestions: [],
  conversationHistory: [],
  currentClauseIndex: 0,
  acceptedChanges: [],
  skippedChanges: [],
  documentText: null
};


// =============================================================
// API HELPERS
// =============================================================

async function callAPI(systemPrompt, userContent, maxTokens = 3000) {
  const response = await fetch('/api/analyze', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ systemPrompt, userContent, maxTokens })
  });
  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error || 'API call failed');
  }
  const data = await response.json();
  if (data.error) throw new Error(data.error);
  return data.text;
}

async function callChat(systemPrompt, messages, maxTokens = 3000) {
  const response = await fetch('/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ systemPrompt, messages, maxTokens })
  });
  const data = await response.json();
  if (data.error) throw new Error(data.error);
  return data.text;
}

function parseJSON(text) {
  const clean = text.replace(/^```json\s*/i, '').replace(/```\s*$/i, '').trim();
  return JSON.parse(clean);
}


// =============================================================
// SCREEN MANAGEMENT
// =============================================================
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  const screen = document.getElementById(id);
  if (screen) screen.classList.add('active');
}

function setLoading(text, sub = '') {
  document.getElementById('loading-text').textContent = text;
  document.getElementById('loading-sub').textContent = sub;
  showScreen('screen-loading');
}


// =============================================================
// OFFICE.JS — DOCUMENT EXTRACTION
// =============================================================
async function extractDocumentText() {
  return await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text;
  });
}


// =============================================================
// MAIN ORCHESTRATION
// =============================================================
async function runReview(clarificationAnswers = null) {
  try {
    setLoading('Reading document…');
    session.documentText = await extractDocumentText();

    if (!session.documentText || session.documentText.trim().length < 50) {
      alert('The document appears to be empty or too short to review. Please ensure a contract is open in Word.');
      showScreen('screen-landing');
      return;
    }

    // Call 1 — Classification
    setLoading('Classifying agreement…', 'Identifying agreement type and applicable Indian statutes');

    let userContent = session.documentText;
    if (clarificationAnswers) {
      userContent = `CLARIFICATION ANSWERS FROM LAWYER:\n${JSON.stringify(clarificationAnswers, null, 2)}\n\nAGREEMENT TEXT:\n${session.documentText}`;
    }

    const call1Text = await callAPI(CALL_1_PROMPT, userContent, 2000);
    session.classificationJSON = parseJSON(call1Text);

    // Confidence gate
    if (session.classificationJSON.confidence && session.classificationJSON.confidence.status === 'LOW') {
      showConfidenceModal(session.classificationJSON.clarification_required);
      return;
    }

    // Call 2 — Triage
    setLoading(
      'Identifying issues…',
      `Reviewing clauses under ${session.classificationJSON.classification.agreement_type} framework`
    );

    const triageInput = JSON.stringify({
      classification: session.classificationJSON,
      agreement_text: session.documentText
    });

    const call2Text = await callAPI(CALL_2_PROMPT, triageInput, 4000);
    session.triageJSON = parseJSON(call2Text);

    // Handle clean agreement
    if (session.triageJSON.no_issues_found) {
      showCleanScreen();
      return;
    }

    // Begin clause-by-clause review
    session.currentClauseIndex = 0;
    session.acceptedChanges = [];
    session.skippedChanges = [];
    session.clauseSuggestions = new Array(session.triageJSON.clauses.length).fill(null);
    await loadClauseCard(0);

  } catch (err) {
    console.error('Review error:', err);
    alert('Error during review: ' + err.message + '\n\nMake sure your Azure credentials are set in .env and the server is running.');
    showScreen('screen-landing');
  }
}


// =============================================================
// CLAUSE CARD — LOAD AND RENDER (Call 3)
// =============================================================
async function loadClauseCard(index) {
  const clause = session.triageJSON.clauses[index];
  const total = session.triageJSON.clauses.length;
  session.conversationHistory = [];

  setLoading(
    `Analysing clause ${index + 1} of ${total}…`,
    clause.clause_name
  );

  try {
    const messages = [{
      role: 'user',
      content: JSON.stringify({
        classification: session.classificationJSON,
        clause: clause
      })
    }];

    const responseText = await callChat(CALL_3_PROMPT, messages, 3000);
    const suggestion = parseJSON(responseText);
    session.clauseSuggestions[index] = suggestion;

    session.conversationHistory.push({ role: 'user', content: messages[0].content });
    session.conversationHistory.push({ role: 'assistant', content: responseText });

    renderClauseCard(clause, suggestion, index, total);
    showScreen('screen-clause');

    if (suggestion.original_text && suggestion.original_text !== 'MISSING') {
      await highlightInDocument(suggestion.original_text, clause.priority);
    }

  } catch (err) {
    console.error('Clause card error:', err);
    alert('Error loading clause suggestion: ' + err.message);
  }
}

function renderClauseCard(clause, suggestion, index, total) {
  document.getElementById('progress-text').textContent = `Clause ${index + 1} of ${total}`;
  document.getElementById('progress-fill').style.width = `${((index + 1) / total) * 100}%`;

  document.getElementById('clause-name').textContent =
    suggestion.clause_name || clause.clause_name;

  const badge = document.getElementById('priority-badge');
  const priority = clause.priority.toUpperCase();
  badge.textContent = priority;
  badge.className = `priority-badge ${priority}`;

  document.getElementById('statute-tag').textContent =
    suggestion.applicable_statute || clause.applicable_statute || '';

  document.getElementById('issue-summary').textContent =
    suggestion.issue_summary || '';
  document.getElementById('original-text').textContent =
    suggestion.original_text === 'MISSING'
      ? '— Clause is missing from the agreement —'
      : (suggestion.original_text || clause.clause_text || '');
  document.getElementById('suggested-text').textContent =
    suggestion.suggested_text || '';
  document.getElementById('negotiation-note').textContent =
    suggestion.negotiation_note || '';
  document.getElementById('fallback-text').textContent =
    suggestion.fallback_text || '';

  const score = suggestion.confidence_score || 0;
  const confFill = document.getElementById('confidence-bar-fill');
  const confLabel = document.getElementById('confidence-label');
  confFill.style.width = `${score}%`;
  confFill.style.background = score >= 90 ? 'var(--success)' : score >= 75 ? 'var(--low)' : 'var(--medium)';
  confLabel.textContent = `${score}% — ${suggestion.confidence_rationale || ''}`;

  document.getElementById('ask-ai-input').value = '';
  const responseEl = document.getElementById('ask-ai-response');
  responseEl.classList.add('hidden');
  responseEl.textContent = '';
}


// =============================================================
// OFFICE.JS — HIGHLIGHT CURRENT CLAUSE
// =============================================================
const PRIORITY_HIGHLIGHT = {
  CRITICAL: 'Red',
  HIGH: 'Orange',
  MEDIUM: 'Yellow',
  LOW: 'Turquoise'
};

async function highlightInDocument(originalText, priority) {
  if (!originalText || originalText === 'MISSING') return;
  try {
    await Word.run(async (context) => {
      const range = await findRange(context, originalText);
      if (range) {
        const p = (priority || 'HIGH').toUpperCase();
        range.font.highlightColor = PRIORITY_HIGHLIGHT[p] || 'Yellow';
        range.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.warn('Could not highlight in document:', e.message);
  }
}

async function clearHighlight(originalText) {
  if (!originalText || originalText === 'MISSING') return;
  try {
    await Word.run(async (context) => {
      const range = await findRange(context, originalText);
      if (range) {
        range.font.highlightColor = 'None';
        await context.sync();
      }
    });
  } catch (e) {
    // non-fatal
  }
}


// =============================================================
// OFFICE.JS — APPLY ACCEPTED CHANGES AS WORD TRACK CHANGES
// =============================================================
async function findRange(context, originalText) {
  // Try progressively shorter substrings until we get a match.
  // Needed because the AI sometimes paraphrases instead of quoting verbatim.
  const lengths = [
    originalText.length,
    200, 150, 100, 60
  ];
  for (const len of lengths) {
    if (len > originalText.length) continue;
    const candidate = originalText.substring(0, len).trim();
    if (candidate.length < 10) continue;
    const results = context.document.body.search(candidate, {
      matchCase: false,
      matchWholeWord: false
    });
    results.load('items');
    await context.sync();
    if (results.items.length > 0) return results.items[0];
  }
  return null;
}

// =============================================================
// ACCEPT / SKIP — changes applied immediately on Accept
// =============================================================
// ACCEPT / SKIP — changes applied immediately on Accept
// =============================================================
async function onAccept() {
  const suggestion = session.clauseSuggestions[session.currentClauseIndex];
  if (!suggestion) { moveToNext(); return; }

  const btnAccept = document.getElementById('btn-accept');
  const btnSkip   = document.getElementById('btn-skip');
  btnAccept.disabled = true;
  btnSkip.disabled   = true;
  btnAccept.textContent = 'Applying…';

  try {
    await clearHighlight(suggestion.original_text);
    await applyOneChange(suggestion);
    session.acceptedChanges.push(suggestion);
  } catch (e) {
    console.warn('Could not apply change:', e.message);
    session.acceptedChanges.push(suggestion);
  } finally {
    btnAccept.disabled = false;
    btnSkip.disabled   = false;
    btnAccept.textContent = 'Accept change ✓';
  }

  moveToNext();
}

async function onSkip() {
  const suggestion = session.clauseSuggestions[session.currentClauseIndex];
  session.skippedChanges.push(session.triageJSON.clauses[session.currentClauseIndex]);
  await clearHighlight(suggestion?.original_text);
  moveToNext();
}

async function moveToNext() {
  session.currentClauseIndex++;
  const total = session.triageJSON.clauses.length;

  if (session.currentClauseIndex < total) {
    await loadClauseCard(session.currentClauseIndex);
  } else {
    showCompletionScreen();
  }
}

// Applies a single suggestion to the Word document as a tracked change
async function applyOneChange(change) {
  await Word.run(async (context) => {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    // ── INSERT (missing clause) ──────────────────────────────
    if (change.change_type === 'Insert' || change.original_text === 'MISSING') {
      const sigSearch = context.document.body.search('IN WITNESS WHEREOF', { matchCase: false });
      sigSearch.load('items');
      await context.sync();

      if (sigSearch.items.length > 0) {
        sigSearch.items[0].insertParagraph(change.suggested_text, 'Before');
        sigSearch.items[0].insertComment(`[AI REVIEWER — NEW CLAUSE | ${change.applicable_statute || ''}] ${change.issue_summary}`);
      } else {
        context.document.body.paragraphs.getLast()
          .insertParagraph(change.suggested_text, 'After');
      }
      await context.sync();

    } else {
      // ── FIND + REPLACE / DELETE ────────────────────────────
      const range = await findRange(context, change.original_text);

      if (range) {
        if (change.change_type === 'Delete') {
          range.insertText('', 'Replace');
        } else {
          range.insertText(change.suggested_text, 'Replace');
        }
        range.insertComment(`[AI REVIEWER | ${change.applicable_statute || ''}] ${change.issue_summary}`);
        await context.sync();
      } else {
        // Can't locate original — leave a comment so nothing is lost
        context.document.body.paragraphs.getLast().insertComment(
          `[AI REVIEWER — APPLY MANUALLY]\nClause: ${change.clause_name}\n${change.issue_summary}\n\nSUGGESTED:\n${change.suggested_text}`
        );
        await context.sync();
      }
    }

    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();
  });
}


// =============================================================
// ASK AI (multi-turn on current clause)
// =============================================================
async function onAskAI() {
  const input = document.getElementById('ask-ai-input');
  const userMessage = input.value.trim();
  if (!userMessage) return;

  const btn = document.getElementById('btn-ask-ai');
  btn.disabled = true;
  btn.textContent = '…';

  try {
    const clause = session.triageJSON.clauses[session.currentClauseIndex];

    session.conversationHistory.push({ role: 'user', content: userMessage });

    const messages = [
      {
        role: 'user',
        content: JSON.stringify({
          classification: session.classificationJSON,
          clause: clause
        })
      },
      ...session.conversationHistory
    ];

    const responseText = await callChat(CALL_3_PROMPT, messages, 3000);
    const updated = parseJSON(responseText);
    session.conversationHistory.push({ role: 'assistant', content: responseText });

    session.clauseSuggestions[session.currentClauseIndex] = updated;

    renderClauseCard(
      clause, updated,
      session.currentClauseIndex,
      session.triageJSON.clauses.length
    );

    const responseEl = document.getElementById('ask-ai-response');
    responseEl.textContent = updated.user_message_addressed || updated.legal_analysis || '';
    responseEl.classList.remove('hidden');
    input.value = '';

  } catch (err) {
    alert('Ask AI error: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Ask';
  }
}


// =============================================================
// CONFIDENCE MODAL
// =============================================================
function showConfidenceModal(clarificationRequired) {
  const container = document.getElementById('modal-questions');
  container.innerHTML = '';

  if (clarificationRequired.message) {
    const msg = document.createElement('p');
    msg.style.cssText = 'font-size:12px;color:#6c757d;margin-bottom:16px;';
    msg.textContent = clarificationRequired.message;
    container.appendChild(msg);
  }

  const questions = clarificationRequired.questions || [];
  questions.forEach(q => {
    const wrapper = document.createElement('div');
    wrapper.className = 'modal-question';

    const label = document.createElement('p');
    label.textContent = q.question;
    wrapper.appendChild(label);

    if (q.options) {
      q.options.forEach(opt => {
        const row = document.createElement('label');
        row.className = 'modal-option';
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.name = q.id;
        radio.value = opt;
        row.appendChild(radio);
        row.appendChild(document.createTextNode(opt));
        wrapper.appendChild(row);
      });
    } else if (q.input_type === 'free_text') {
      const textarea = document.createElement('textarea');
      textarea.id = `modal-${q.id}`;
      textarea.className = 'modal-textarea';
      textarea.placeholder = q.placeholder || '';
      wrapper.appendChild(textarea);
    }

    container.appendChild(wrapper);
  });

  document.getElementById('modal-confidence').classList.remove('hidden');
  document.getElementById('screen-loading').classList.remove('active');
  showScreen('screen-landing');
}

function collectModalAnswers() {
  const answers = {};
  const questions = document.querySelectorAll('.modal-question');
  questions.forEach((qDiv, i) => {
    const radios = qDiv.querySelectorAll('input[type=radio]');
    if (radios.length > 0) {
      const checked = qDiv.querySelector('input[type=radio]:checked');
      if (checked) answers[`Q${i + 1}`] = checked.value;
    } else {
      const textarea = qDiv.querySelector('textarea');
      if (textarea && textarea.value.trim()) {
        answers[`Q${i + 1}`] = textarea.value.trim();
      }
    }
  });
  return answers;
}


// =============================================================
// CLEAN AND COMPLETION SCREENS
// =============================================================
function showCleanScreen() {
  const c = session.classificationJSON;
  document.getElementById('clean-meta').innerHTML = `
    <strong>${c.classification.agreement_type}</strong><br>
    Client: ${c.perspective.inferred_perspective}<br>
    Statutes: ${(c.classification.applicable_statutes || []).slice(0, 3).join(', ')}
  `;
  document.getElementById('clean-note').textContent =
    session.triageJSON.clean_agreement_note ||
    'No issues identified under Indian law for this agreement type and client position.';
  showScreen('screen-clean');
}

function showCompletionScreen() {
  const accepted = session.acceptedChanges.length;
  const skipped = session.skippedChanges.length;
  document.getElementById('completion-stats').innerHTML = `
    <div class="stat"><span>Changes applied</span><span>${accepted}</span></div>
    <div class="stat"><span>Clauses skipped</span><span>${skipped}</span></div>
    <div class="stat"><span>Total reviewed</span><span>${session.triageJSON.total_clauses_flagged}</span></div>
  `;
  showScreen('screen-complete');
}


// =============================================================
// INIT AND EVENT LISTENERS
// =============================================================
Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById('btn-start').addEventListener('click', () => runReview());
    document.getElementById('btn-accept').addEventListener('click', onAccept);
    document.getElementById('btn-skip').addEventListener('click', onSkip);
    document.getElementById('btn-ask-ai').addEventListener('click', onAskAI);
    document.getElementById('ask-ai-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') onAskAI();
    });
    document.getElementById('btn-modal-submit').addEventListener('click', () => {
      const answers = collectModalAnswers();
      document.getElementById('modal-confidence').classList.add('hidden');
      runReview(answers);
    });
    document.getElementById('btn-clean-done').addEventListener('click', () => showScreen('screen-landing'));
    document.getElementById('btn-restart').addEventListener('click', () => {
      Object.assign(session, {
        classificationJSON: null,
        triageJSON: null,
        clauseSuggestions: [],
        conversationHistory: [],
        currentClauseIndex: 0,
        acceptedChanges: [],
        skippedChanges: [],
        documentText: null
      });
      showScreen('screen-landing');
    });
  }
});
