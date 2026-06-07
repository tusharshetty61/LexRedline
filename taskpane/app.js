window.onerror = (msg, src, line) => {
  document.body.innerHTML = `<div style="padding:20px;color:red;font-family:monospace">
    <b>Script error</b><br>${msg}<br>${src}:${line}
  </div>`;
};

// =============================================================
// BASELINE STRUCTURE — 22 sections, five zones
// =============================================================
const BASELINE_SECTIONS = [
  '1. Parties', '2. Background / Recitals', '3. Definitions', '4. Scope of Agreement',
  '5. Services / Obligations', '6. Fees and Payment', '7. Term', '8. Milestones / Deliverables',
  '9. Representations and Warranties', '10. Indemnification', '11. Limitation of Liability',
  '12. Confidentiality', '13. Data Protection / Privacy', '14. IP Ownership / Licence',
  '15. Termination', '16. Force Majeure', '17. Assignment',
  '18. Governing Law', '19. Dispute Resolution / Arbitration', '20. Notices',
  '21. Miscellaneous / Boilerplate', '22. Execution / Signature Block'
];

const BASELINE_VARIANTS = {
  'NDA': [1,2,3,4,10,11,18,19,20,21],
  'Employment Agreement': [1,2,3,4,5,6,7,8,9,10,15,16,17,18,19,20,21],
  'MSA': null,
  'SaaS Agreement': null,
  'Service Agreement': null,
  'Consulting Agreement': [1,2,3,4,5,6,7,8,9,10,11,12,15,16,18,19,20,21],
  'SOW': [1,2,3,4,5,6,7,8,15,20,21],
  'Shareholders Agreement': [1,2,3,4,6,7,10,11,15,17,18,19,20,21],
  'Loan Agreement': [1,2,3,4,6,7,10,11,15,18,19,20,21],
  'Lease Agreement': [1,2,3,4,5,6,7,15,18,19,20,21],
};

function selectBaselineVariant(agreementType, multiAgreementFlag, allTypes) {
  if (multiAgreementFlag && allTypes && allTypes.length > 1) {
    const indices = new Set();
    allTypes.forEach(t => {
      const v = BASELINE_VARIANTS[t];
      if (v) v.forEach(i => indices.add(i));
      else BASELINE_SECTIONS.forEach((_, i) => indices.add(i + 1));
    });
    return [...indices].sort((a, b) => a - b).map(i => BASELINE_SECTIONS[i - 1]);
  }
  const variant = BASELINE_VARIANTS[agreementType];
  if (!variant) return BASELINE_SECTIONS;
  return variant.map(i => BASELINE_SECTIONS[i - 1]);
}


// =============================================================
// CALL 1 — CLASSIFICATION PROMPT
// =============================================================
const CALL_1_PROMPT = `You are an expert contract classification engine operating inside a Microsoft Word add-in used by practising lawyers in India.

Your sole objective is to classify the agreement, infer the lawyer's client perspective, determine risk posture from the draft origin, and produce a structured JSON handoff package for downstream review calls.

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
- The secondary agreement type(s) in all_agreement_types
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
TASK 3: RISK POSTURE FROM DRAFT ORIGIN
═══════════════════════════════════════════════════════

The user has provided a draft_origin value indicating who prepared this draft.

Determine risk_posture as follows:
- draft_origin = "counterparty": risk_posture = "defensive" — the draft is likely skewed against the client. Flag everything improvable. Frame negotiation notes as counterparty is likely to push back.
- draft_origin = "client": risk_posture = "aggressive" — the draft favours the client. Do not flag clauses that merely favour the client; flag only statutory problems.
- draft_origin = "third party" or "unknown": risk_posture = "neutral" — balanced analysis.

═══════════════════════════════════════════════════════
TASK 4: DETERMINE REVIEW PRIORITIES
═══════════════════════════════════════════════════════

Identify the review priorities for this agreement type and client perspective under Indian law. Rank from highest to lowest priority. For each provide a one-line reason anchored to the specific Indian statutory or common law risk.

═══════════════════════════════════════════════════════
TASK 4B: ENHANCED SCRUTINY AREAS
═══════════════════════════════════════════════════════

Identify 2–4 areas requiring deeper analysis. An area qualifies if ANY of the following apply:
- The clause engages a specific Indian statute that limits or overrides contractual freedom
- The clause is frequently litigated in Indian courts for this agreement type
- The clause involves a provision that is unenforceable or restricted under Indian law regardless of what the parties have agreed
- Regulatory compliance exposure exists under Indian law

═══════════════════════════════════════════════════════
TASK 5: CONFIDENCE DECISION ENGINE
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
    "all_agreement_types": [],
    "secondary_agreement_types": [],
    "alternative_classifications": [{ "type": "", "confidence": 0, "reason": "" }]
  },
  "perspective": {
    "inferred_perspective": "",
    "perspective_confidence": 0,
    "perspective_indicators": ["", ""],
    "perspective_caveat": ""
  },
  "risk_posture": "aggressive | defensive | neutral",
  "confidence": {
    "status": "HIGH",
    "assumptions_made": []
  },
  "review_priorities": [
    { "rank": 1, "area": "", "reason": "", "statute": "" }
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
1. A classification JSON object from Call 1 (including risk_posture and draft_origin)
2. A section manifest — array of section headings extracted from the actual document
3. The full segmented agreement text (sections array, including any tables)
4. An active_baseline — the subset of the 22-section standard baseline applicable to this agreement type

Your task is to identify every clause that warrants a suggested change and return them as an ordered list — highest priority first.

All analysis is governed by Indian law. Apply the applicable statutes from the classification JSON.

RISK POSTURE INSTRUCTIONS:
- risk_posture = "defensive" (counterparty draft): flag everything improvable for the client.
- risk_posture = "aggressive" (client draft): flag ONLY statutory conflicts, regulatory non-compliance, and missing mandatory clauses. Do NOT flag clauses that merely favour the client.
- risk_posture = "neutral": balanced analysis — flag all meaningful risks.

MISSING CLAUSE INSTRUCTION: Only flag a section as missing if it is present in the active_baseline AND absent from the section_manifest. Do NOT flag sections that are not in the active_baseline for this agreement type.

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

MATERIAL RISK TO CLIENT: The clause creates a meaningful legal or commercial risk for the client that a competent lawyer would advise redlining.
- Uncapped or inadequately capped liability
- Unilateral termination rights without cause or notice
- Broad indemnification obligations exceeding S.124 Contract Act scope
- IP assignment clauses stripping rights the client should retain
- Governing law or dispute resolution clauses that disadvantage enforcement

MISSING CLAUSE: A clause absent from the agreement that is required by Indian law or as a matter of standard legal practice for this agreement type — AND present in the active_baseline provided.

STRUCTURAL INCONSISTENCY: Two or more clauses in this agreement contradict each other on the same legal point such that both cannot be simultaneously true or enforceable. Flag the weaker or less client-protective clause. Set conflicting_clause_id to the clause_id of the other clause involved.
Examples:
- Clause 3.1 implies contract ends on delivery milestone with no notice; Clause 14.2 requires 30-day notice for termination.
- A term is defined one way in Clause 1 but used differently in Clause 9.
- Payment terms are 30 days in Clause 5 but 45 days in Schedule 2.
Do NOT flag clauses that use different language to say the same thing.

Do NOT flag clauses that are:
- Market-standard and carry no meaningful risk
- Stylistically imperfect but legally sound
- Already adequately protective of the client's position
- Absent from the agreement but also absent from the active_baseline

For missing clauses, set clause_text to "MISSING" and populate missing_clause_description.

SECTION MANIFEST INSTRUCTION: Evaluate ALL sections listed in section_manifest. Your output must include sections_reviewed listing every section_id you examined.

Return this JSON structure:
{
  "agreement_type": "",
  "client_perspective": "",
  "applicable_statutes": ["", ""],
  "total_clauses_flagged": 0,
  "sections_reviewed": ["0", "1", "2"],
  "clauses": [
    {
      "rank": 1,
      "clause_id": "",
      "clause_name": "",
      "clause_text": "",
      "issue_category": "Statutory Conflict | Regulatory Non-Compliance | Material Risk | Missing Clause | Structural Inconsistency",
      "conflicting_clause_id": "",
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
1. A classification JSON object from Call 1 (including risk_posture and draft_origin)
2. A single clause object from the triage JSON — including clause_text (verbatim from document), char_start and char_end (character positions of this segment in the full document text)
3. Optionally: conversation_history if the lawyer has asked follow-up questions
4. pending_changes: a list of changes already accepted in this review session — check if any affect the risk profile of this clause
5. section_manifest: the actual section headings from the document (use verbatim headings for insert_anchor)
6. baseline_sections: the standard baseline for this agreement type
7. defined_terms: list of defined terms extracted from the Definitions clause

All advice is governed by Indian law. Cite specific Indian statutes, rules, and leading cases where relevant.

Your suggested language must be draft-quality — at the standard of a trained solicitor, ready to be accepted and inserted into the agreement without further editing. Use ONLY defined terms from the defined_terms list provided. If you need a term not in the list, either write around it in plain language OR include it in undefined_terms_used with a suggested definition.

DRAFT ORIGIN INSTRUCTIONS:
- risk_posture = "defensive": frame all suggestions strongly in the client's favour.
- risk_posture = "aggressive": flag only statutory problems; suggestions should be more balanced.
- risk_posture = "neutral": balanced approach.

PENDING CHANGES: Review the pending_changes list. If any accepted change affects the risk profile of the current clause (e.g. a deleted non-compete shifts post-term restraint weight to the confidentiality clause), set downstream_impact to a plain-English explanation. Set to null if no dependency exists.

MISSING CLAUSE INSERTION: For change_type "Insert", set insert_anchor to a verbatim heading string from section_manifest — find where this clause type appears in baseline_sections, identify its predecessor section, then map to the nearest matching heading in section_manifest. Set insert_mode to "after_section" or "sub_clause". Set insert_position_note to a plain-English sentence explaining the placement.

Temperature is set to 0. Be deterministic. Return only valid JSON. No preamble or explanation outside the JSON object.

If conversation_history is provided:
- Address the lawyer's most recent message directly
- Revise the suggestion if requested
- Answer negotiation strategy questions in the legal_analysis field
- Maintain the same JSON structure throughout

═══════════════════════════════════════════════════════
ANALYSIS FRAMEWORK
═══════════════════════════════════════════════════════

1. STATUTORY POSITION: What does Indian law say about this clause type? Cite the specific provision. Is the current drafting enforceable as written?

2. CURRENT DRAFTING RISK: What is the specific legal or commercial risk created by the current language for the client?

3. SUGGESTED POSITION: Draft the replacement language. Legally sound under Indian law, consistent with the agreement's defined terms and drafting style, appropriate for the client's perspective.

4. NEGOTIATION RATIONALE: Why should the client push for this change? What is the fallback position? What would a court look at if this clause were disputed?

FOR MISSING CLAUSES: Draft a complete clause from scratch. Market-standard under Indian law for this agreement type and client perspective.

═══════════════════════════════════════════════════════
FIELD DEFINITIONS
═══════════════════════════════════════════════════════

issue_summary: Two to three sentences. State the legal problem, Indian law basis, and consequence for the client if unchanged. Write for a lawyer — use correct legal terminology.

legal_analysis: Detailed legal reasoning. Cite statutes by section number. Reference leading Indian cases where relevant (e.g. ONGC v Saw Pipes for S.74; Nirma Ltd v Lurgi for indemnity scope; Centrotrade for seat vs venue of arbitration).

original_text: Verbatim text from the agreement exactly as it appears. Used for string search in Word. Do not paraphrase or truncate. For missing clauses: set to "MISSING".

char_start: For Replace or Delete only — the character index in the full document text where the specific text to be changed begins. This must be within the clause's char_start..char_end range provided. Use this to pinpoint a sub-clause or sentence rather than the whole section if only part of the section needs changing. Omit (set to null) for Insert changes.

char_end: For Replace or Delete only — the character index where the specific text to be changed ends. Omit (set to null) for Insert changes.

suggested_text: Complete replacement clause, ready to insert. For missing clauses: complete draft clause. For deletions: set to "DELETE".

fallback_text: A middle-ground version the client could accept if the counterparty resists. Must still be legally sound. If no meaningful fallback exists, state why.

negotiation_note: What to say to the counterparty. Fallback position. What concession can the client make without material harm. Practical negotiation guide for the lawyer.

confidence_score: 0–100. 90–100: Unambiguous Indian law — settled statute or Supreme Court authority. 75–89: Well-established practice — High Court authority. 60–74: Reasonable position — some variation. Below 60: Unsettled law — flag for independent judgment.

downstream_impact: Non-null string if any accepted pending_change affects the risk profile of this clause. Null otherwise.

undefined_terms_used: Array of objects { term, suggested_definition } for any term in suggested_text or fallback_text not found in defined_terms. Empty array [] if all terms are defined.

insert_anchor: For change_type "Insert" only — verbatim heading string from section_manifest after which to insert. Null if no suitable anchor exists.

insert_mode: "after_section" or "sub_clause".

insert_position_note: Plain-English one sentence shown in the task pane explaining why this position was chosen.

═══════════════════════════════════════════════════════
OUTPUT
═══════════════════════════════════════════════════════

{
  "clause_name": "",
  "issue_category": "Statutory Conflict | Regulatory Non-Compliance | Material Risk | Missing Clause | Structural Inconsistency",
  "applicable_statute": "",
  "issue_summary": "",
  "legal_analysis": "",
  "original_text": "",
  "suggested_text": "",
  "fallback_text": "",
  "negotiation_note": "",
  "change_type": "Replace | Insert | Delete",
  "char_start": null,
  "char_end": null,
  "insert_anchor": null,
  "insert_mode": "after_section",
  "insert_position_note": "",
  "downstream_impact": null,
  "undefined_terms_used": [],
  "confidence_score": 0,
  "confidence_rationale": "",
  "user_message_addressed": "",
  "perspective_caveat": ""
}`;


// =============================================================
// SESSION STATE
// =============================================================
const sessionState = {
  classificationJSON: null,
  triageJSON: null,
  currentClauseIndex: 0,
  conversationHistory: [],
  conversationHistoryByClause: {},
  pendingChanges: {},          // { clauseIndex: suggestionObj } — deferred apply
  rejectedChanges: [],
  suggestionCache: {},         // { clauseIndex: suggestionObj }
  decisionMap: {},             // { clauseIndex: 'accepted'|'accepted_fallback'|'rejected'|null }
  documentText: null,
  clauseMap: [],
  tableMap: [],
  definedTerms: [],
  draftOrigin: 'counterparty',
  sectionManifest: [],
  activeBaseline: [],
};


// =============================================================
// API HELPERS
// =============================================================
async function callAPI(systemPrompt, userContent, maxTokens = 4000) {
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

async function callChat(systemPrompt, messages, maxTokens = 4000) {
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
// PRE-PROCESSING
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

function extractDefinedTerms(rawText) {
  const terms = new Set();
  const quotedPattern = /[“”""]([A-Z][A-Za-z\s]+)["”“"]\s+(?:means|shall mean)/g;
  const capPattern = /([A-Z][A-Za-z]+(?:\s[A-Z][A-Za-z]+)*)\s+(?:means|shall mean)/g;
  let m;
  while ((m = quotedPattern.exec(rawText)) !== null) terms.add(m[1].trim());
  while ((m = capPattern.exec(rawText)) !== null) terms.add(m[1].trim());
  return [...terms];
}

function segmentDocument(rawText) {
  const headingRegex = /^(\d+\.(?:\d+\.?)*\s+[A-Z][A-Za-z]|[A-Z]{4,}[\s\w]*$)/gm;
  const segments = [];
  let lastIndex = 0;
  let lastHeading = 'Preamble';
  let segId = 0;
  let match;

  while ((match = headingRegex.exec(rawText)) !== null) {
    if (match.index > lastIndex) {
      segments.push({
        section_id: String(segId++),
        heading: lastHeading,
        text: rawText.slice(lastIndex, match.index).trim(),
        charStart: lastIndex,    // position in full document text
        charEnd: match.index
      });
    }
    lastHeading = match[0].trim();
    lastIndex = match.index;
  }
  segments.push({
    section_id: String(segId),
    heading: lastHeading,
    text: rawText.slice(lastIndex).trim(),
    charStart: lastIndex,
    charEnd: rawText.length
  });
  return segments;
}

async function extractTables() {
  return await Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load('items');
    await ctx.sync();

    const result = [];
    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      table.load('rowCount,columnCount');
      table.rows.load('items');
      await ctx.sync();

      for (const row of table.rows.items) {
        row.cells.load('items');
      }
      await ctx.sync();

      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          cell.body.load('text');
        }
      }
      await ctx.sync();

      let md = '';
      table.rows.items.forEach((row, ri) => {
        const cells = row.cells.items.map(c => c.body.text.trim());
        md += '| ' + cells.join(' | ') + ' |\n';
        if (ri === 0) md += '| ' + cells.map(() => '---').join(' | ') + ' |\n';
      });

      result.push({
        table_id: `table_${i}`,
        source: 'table',
        markdown: md,
        row_count: table.rowCount
      });
    }
    return result;
  }).catch(() => []);
}

function mergeTablesIntoSegments(segments, tables) {
  tables.forEach((t, i) => {
    segments.push({
      section_id: `table_${i}`,
      heading: `[TABLE ${i + 1}]`,
      text: t.markdown,
      source: 'table'
    });
  });
  return segments;
}


// =============================================================
// MAIN ORCHESTRATION
// =============================================================
async function runReview(clarificationAnswers = null) {
  try {
    // Read draft origin from radio
    const originRadio = document.querySelector('input[name="draft-origin"]:checked');
    sessionState.draftOrigin = originRadio ? originRadio.value : 'unknown';

    setLoading('Reading document…');
    sessionState.documentText = await extractDocumentText();

    if (!sessionState.documentText || sessionState.documentText.trim().length < 50) {
      alert('The document appears to be empty or too short to review. Please ensure a contract is open in Word.');
      showScreen('screen-landing');
      return;
    }

    // Pre-processing
    setLoading('Pre-processing document…', 'Extracting sections, tables, and defined terms');
    sessionState.definedTerms = extractDefinedTerms(sessionState.documentText);
    const segments = segmentDocument(sessionState.documentText);
    sessionState.sectionManifest = segments.map(s => s.heading);
    const tables = await extractTables();
    sessionState.tableMap = tables;
    const mergedSegments = mergeTablesIntoSegments([...segments], tables);
    sessionState.clauseMap = mergedSegments;

    // Call 1 — Classification
    setLoading('Classifying agreement…', 'Identifying agreement type and applicable Indian statutes');

    let userContent = `DRAFT ORIGIN: ${sessionState.draftOrigin}\n\nAGREEMENT TEXT:\n${sessionState.documentText}`;
    if (clarificationAnswers) {
      userContent = `DRAFT ORIGIN: ${sessionState.draftOrigin}\n\nCLARIFICATION ANSWERS FROM LAWYER:\n${JSON.stringify(clarificationAnswers, null, 2)}\n\nAGREEMENT TEXT:\n${sessionState.documentText}`;
    }

    const call1Text = await callAPI(CALL_1_PROMPT, userContent, 2000);
    sessionState.classificationJSON = parseJSON(call1Text);

    if (sessionState.classificationJSON.confidence && sessionState.classificationJSON.confidence.status === 'LOW') {
      showConfidenceModal(sessionState.classificationJSON.clarification_required);
      return;
    }

    // Select baseline variant for this agreement type
    const cls = sessionState.classificationJSON.classification;
    sessionState.activeBaseline = selectBaselineVariant(
      cls.agreement_type,
      cls.multi_agreement_flag,
      cls.all_agreement_types || [cls.agreement_type]
    );

    // Call 2 — Triage
    setLoading(
      'Identifying issues…',
      `Reviewing clauses under ${cls.agreement_type} framework`
    );

    const call2Payload = JSON.stringify({
      classification: sessionState.classificationJSON,
      draft_origin: sessionState.draftOrigin,
      risk_posture: sessionState.classificationJSON.risk_posture || 'neutral',
      section_manifest: sessionState.sectionManifest,
      active_baseline: sessionState.activeBaseline,
      sections: mergedSegments,
      instruction: 'Evaluate ALL sections listed in section_manifest. Your output must include sections_reviewed listing every section_id you examined.'
    });

    const call2Text = await callAPI(CALL_2_PROMPT, call2Payload, 6000);
    sessionState.triageJSON = parseJSON(call2Text);

    if (sessionState.triageJSON.no_issues_found) {
      showCleanScreen();
      return;
    }

    // Reset navigation state
    sessionState.currentClauseIndex = 0;
    sessionState.pendingChanges = {};
    sessionState.rejectedChanges = [];
    sessionState.suggestionCache = {};
    sessionState.decisionMap = {};
    sessionState.conversationHistoryByClause = {};

    await navigateTo(0);

  } catch (err) {
    console.error('Review error:', err);
    alert('Error during review: ' + err.message + '\n\nMake sure your OPENAI_API_KEY is set in .env and the server is running.');
    showScreen('screen-landing');
  }
}


// =============================================================
// SEGMENT LOOKUP — verbatim text from clauseMap for reliable searching
// =============================================================
function resolveSegment(clause) {
  const map = sessionState.clauseMap;
  if (!map || map.length === 0) return null;

  const id   = (clause.clause_id   || '').toLowerCase().trim();
  const name = (clause.clause_name || '').toLowerCase().trim();

  // 1. Match by clause_id prefix in heading (e.g. "14.2" → "14.2 Termination")
  if (id) {
    const seg = map.find(s => {
      const h = s.heading.toLowerCase();
      return h.startsWith(id + ' ') || h.startsWith(id + '.') || h === id;
    });
    if (seg && seg.text && seg.text.length > 10) return seg;
  }

  // 2. Match by clause_name substring in heading
  if (name) {
    const seg = map.find(s => s.heading.toLowerCase().includes(name));
    if (seg && seg.text && seg.text.length > 10) return seg;
  }

  return null;
}

// Keep backward-compat alias used in navigateTo / onAskAI
function resolveSegmentText(clause) {
  const seg = resolveSegment(clause);
  return seg ? seg.text : null;
}

// Find the Word range for a clause body by searching for its heading, then
// selecting from END of heading to START of next heading.
// The heading is preserved; only the body text is in the returned range.
// Headings are short, verbatim, and unique — far more reliable than
// searching for long body text.
async function findRangeByHeading(clause, ctx) {
  const seg = resolveSegment(clause);
  if (!seg || !seg.heading || seg.heading === 'Preamble') return null;

  const headingResults = ctx.document.body.search(seg.heading.trim(), { matchCase: false });
  headingResults.load('items');
  await ctx.sync();
  if (headingResults.items.length === 0) return null;

  // Body starts at the END of the heading range
  const bodyStart = headingResults.items[0].getRange('End');

  const segIndex = sessionState.clauseMap.indexOf(seg);
  const nextSeg = sessionState.clauseMap
    .slice(segIndex + 1)
    .find(s => s.heading && s.heading !== 'Preamble' && !s.source);

  if (!nextSeg) {
    const bodyEnd = ctx.document.body.getRange('End');
    return bodyStart.expandTo(bodyEnd);
  }

  const nextResults = ctx.document.body.search(nextSeg.heading.trim(), { matchCase: false });
  nextResults.load('items');
  await ctx.sync();

  if (nextResults.items.length === 0) {
    const bodyEnd = ctx.document.body.getRange('End');
    return bodyStart.expandTo(bodyEnd);
  }

  // Body ends at the START of the next heading
  const nextHeadingStart = nextResults.items[0].getRange('Start');
  return bodyStart.expandTo(nextHeadingStart);
}


// =============================================================
// NAVIGATION
// =============================================================
async function navigateTo(index) {
  const total = sessionState.triageJSON.clauses.length;
  if (index < 0 || index >= total) return;

  // Save conversation history for current clause
  sessionState.conversationHistoryByClause[sessionState.currentClauseIndex] =
    [...sessionState.conversationHistory];

  sessionState.currentClauseIndex = index;

  // Restore conversation history for target clause
  sessionState.conversationHistory =
    sessionState.conversationHistoryByClause[index] || [];

  const clause = sessionState.triageJSON.clauses[index];

  // Serve from cache if available
  if (sessionState.suggestionCache[index]) {
    renderClauseCard(clause, sessionState.suggestionCache[index], index, total);
    updateNavButtons(index, total);
    showScreen('screen-clause');
    return;
  }

  // Fire Call 3
  showCardLoadingState(clause, index, total);

  try {
    // Look up verbatim segment (heading + text + char positions) before sending to model
    const seg = resolveSegment(clause);
    const verbatimText = seg ? seg.text : null;

    const call3Payload = JSON.stringify({
      classification: sessionState.classificationJSON,
      clause: seg
        ? { ...clause, clause_text: seg.text, char_start: seg.charStart, char_end: seg.charEnd }
        : clause,
      draft_origin: sessionState.draftOrigin,
      risk_posture: sessionState.classificationJSON.risk_posture || 'neutral',
      conversation_history: sessionState.conversationHistory,
      pending_changes: Object.values(sessionState.pendingChanges),
      section_manifest: sessionState.sectionManifest,
      baseline_sections: sessionState.activeBaseline,
      defined_terms: sessionState.definedTerms
    });

    const messages = [{ role: 'user', content: call3Payload }];
    const responseText = await callChat(CALL_3_PROMPT, messages, 4000);
    const suggestion = parseJSON(responseText);

    // Resolve original_text in priority order:
    // 1. AI-returned char indices — most precise (sub-clause level)
    // 2. Full segment text — whole section fallback
    // Either way, text comes from sessionState.documentText, never from AI quoting.
    if (suggestion.original_text !== 'MISSING' && suggestion.change_type !== 'Insert') {
      const aiStart = suggestion.char_start;
      const aiEnd   = suggestion.char_end;
      const docText = sessionState.documentText;
      if (aiStart != null && aiEnd != null && aiEnd > aiStart
          && seg && aiStart >= seg.charStart && aiEnd <= seg.charEnd + 50) {
        // Indices are within the expected segment bounds — use them
        suggestion.original_text = docText.substring(aiStart, aiEnd).trim();
      } else if (verbatimText) {
        // Fall back to whole segment verbatim text
        suggestion.original_text = verbatimText;
      }
    }

    sessionState.suggestionCache[index] = suggestion;
    sessionState.conversationHistory = [
      { role: 'user', content: call3Payload },
      { role: 'assistant', content: responseText }
    ];
    sessionState.conversationHistoryByClause[index] = [...sessionState.conversationHistory];

    renderClauseCard(clause, suggestion, index, total);
    updateNavButtons(index, total);
    showScreen('screen-clause');

    if (suggestion.original_text && suggestion.original_text !== 'MISSING') {
      await highlightInDocument(suggestion.original_text, clause.priority);
    }

  } catch (err) {
    console.error('Clause card error:', err);
    alert('Error loading clause suggestion: ' + err.message);
    showScreen('screen-landing');
  }
}

function showCardLoadingState(clause, index, total) {
  setLoading(
    `Analysing clause ${index + 1} of ${total}…`,
    clause.clause_name
  );
}

function onPrev() { navigateTo(sessionState.currentClauseIndex - 1); }
function onNext() { navigateTo(sessionState.currentClauseIndex + 1); }


// =============================================================
// RENDER CLAUSE CARD
// =============================================================
function renderClauseCard(clause, suggestion, index, total) {
  document.getElementById('progress-text').textContent = `Clause ${index + 1} of ${total}`;
  document.getElementById('progress-fill').style.width = `${((index + 1) / total) * 100}%`;

  document.getElementById('clause-name').textContent =
    suggestion.clause_name || clause.clause_name;

  const badge = document.getElementById('priority-badge');
  const priority = (clause.priority || 'HIGH').toUpperCase();
  badge.textContent = priority;
  badge.className = `priority-badge ${priority}`;

  document.getElementById('statute-tag').textContent =
    suggestion.applicable_statute || clause.applicable_statute || '';

  document.getElementById('issue-summary').textContent = suggestion.issue_summary || '';
  document.getElementById('original-text').textContent =
    suggestion.original_text === 'MISSING'
      ? '— Clause is missing from the agreement —'
      : (suggestion.original_text || clause.clause_text || '');
  document.getElementById('suggested-text').textContent = suggestion.suggested_text || '';
  document.getElementById('negotiation-note').textContent = suggestion.negotiation_note || '';
  document.getElementById('fallback-text').textContent = suggestion.fallback_text || '';

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

  // Dependency warning
  const depWarn = document.getElementById('dependency-warning');
  const depText = document.getElementById('dependency-warning-text');
  if (suggestion.downstream_impact) {
    depText.textContent = suggestion.downstream_impact;
    depWarn.classList.remove('hidden');
  } else {
    depWarn.classList.add('hidden');
  }

  // Undefined terms warning
  const termWarn = document.getElementById('undefined-terms-warning');
  const termText = document.getElementById('undefined-terms-text');
  const undefinedTerms = suggestion.undefined_terms_used || [];
  if (undefinedTerms.length > 0) {
    const termNames = undefinedTerms.map(u => u.term).join(', ');
    termText.textContent = `Undefined terms in suggestion: ${termNames}. `;
    termWarn.classList.remove('hidden');
    // Store for the accept+def button
    termWarn._undefinedTerms = undefinedTerms;
  } else {
    termWarn.classList.add('hidden');
    termWarn._undefinedTerms = [];
  }

  // Decision badge & Undo button
  updateDecisionBadge(index);

  // Update status counts
  updateStatusCounts();
}

function updateDecisionBadge(index) {
  const badge = document.getElementById('decision-badge');
  const undoBtn = document.getElementById('btn-undo');
  const decision = sessionState.decisionMap[index];

  if (decision === 'accepted') {
    badge.textContent = '✔ Accepted';
    badge.className = 'decision-badge accepted';
    undoBtn.classList.remove('hidden');
  } else if (decision === 'accepted_fallback') {
    badge.textContent = '✔ Accepted Fallback';
    badge.className = 'decision-badge accepted-fallback';
    undoBtn.classList.remove('hidden');
  } else if (decision === 'rejected') {
    badge.textContent = '✗ Rejected';
    badge.className = 'decision-badge rejected';
    undoBtn.classList.remove('hidden');
  } else {
    badge.textContent = '';
    badge.className = 'decision-badge hidden';
    undoBtn.classList.add('hidden');
  }
}

function updateStatusCounts() {
  const map = sessionState.decisionMap;
  const total = sessionState.triageJSON ? sessionState.triageJSON.clauses.length : 0;
  let accepted = 0, fallback = 0, rejected = 0;
  Object.values(map).forEach(d => {
    if (d === 'accepted') accepted++;
    else if (d === 'accepted_fallback') fallback++;
    else if (d === 'rejected') rejected++;
  });
  const pending = total - accepted - fallback - rejected;

  const el = document.getElementById('status-counts');
  el.innerHTML =
    `<span class="s-accepted">✔ ${accepted + fallback} accepted${fallback ? ` (${fallback} fallback)` : ''}</span> ` +
    `<span class="s-rejected">✗ ${rejected} rejected</span> ` +
    `<span class="s-pending">● ${pending} pending</span>`;
}

function updateNavButtons(index, total) {
  document.getElementById('btn-prev').disabled = (index === 0);
  document.getElementById('btn-next').disabled = (index === total - 1);

  const allVisited = Object.keys(sessionState.suggestionCache).length === total;
  const finaliseBtn = document.getElementById('btn-finalise');
  if (index === total - 1 || allVisited) {
    finaliseBtn.classList.remove('hidden');
  } else {
    finaliseBtn.classList.add('hidden');
  }
}


// =============================================================
// DECISIONS — deferred until Finalise
// =============================================================
function onAccept() {
  const idx = sessionState.currentClauseIndex;
  const suggestion = sessionState.suggestionCache[idx];
  if (!suggestion) return;

  sessionState.pendingChanges[idx] = { ...suggestion, accepted_as: 'full' };
  sessionState.decisionMap[idx] = 'accepted';
  updateDecisionBadge(idx);
  updateStatusCounts();
  onNext();
}

function onAcceptFallback() {
  const idx = sessionState.currentClauseIndex;
  const suggestion = sessionState.suggestionCache[idx];
  if (!suggestion || !suggestion.fallback_text) return;

  sessionState.pendingChanges[idx] = {
    ...suggestion,
    suggested_text: suggestion.fallback_text,
    accepted_as: 'fallback'
  };
  sessionState.decisionMap[idx] = 'accepted_fallback';
  updateDecisionBadge(idx);
  updateStatusCounts();
  onNext();
}

function onReject() {
  const idx = sessionState.currentClauseIndex;
  const suggestion = sessionState.suggestionCache[idx];
  if (suggestion) {
    sessionState.rejectedChanges.push(sessionState.triageJSON.clauses[idx]);
  }
  delete sessionState.pendingChanges[idx];
  sessionState.decisionMap[idx] = 'rejected';
  updateDecisionBadge(idx);
  updateStatusCounts();
  onNext();
}

function onUndoDecision() {
  const idx = sessionState.currentClauseIndex;
  delete sessionState.pendingChanges[idx];
  sessionState.decisionMap[idx] = null;
  updateDecisionBadge(idx);
  updateStatusCounts();
}

function onAcceptAddDefinition() {
  const idx = sessionState.currentClauseIndex;
  const suggestion = sessionState.suggestionCache[idx];
  if (!suggestion) return;

  // Queue the main change
  sessionState.pendingChanges[idx] = { ...suggestion, accepted_as: 'full' };
  sessionState.decisionMap[idx] = 'accepted';

  // Queue definition insertions for each undefined term
  const termWarn = document.getElementById('undefined-terms-warning');
  const undefinedTerms = termWarn._undefinedTerms || [];
  undefinedTerms.forEach((u, i) => {
    const defIdx = `def_${idx}_${i}`;
    sessionState.pendingChanges[defIdx] = {
      clause_name: `Definition: ${u.term}`,
      change_type: 'Insert',
      original_text: 'MISSING',
      suggested_text: `"${u.term}" means ${u.suggested_definition}`,
      insert_anchor: null,
      insert_position_note: `Definition of "${u.term}" to be inserted in the Definitions clause`,
      applicable_statute: '',
      issue_summary: `Adding definition for term "${u.term}" used in suggested redline.`,
      accepted_as: 'full'
    };
  });

  updateDecisionBadge(idx);
  updateStatusCounts();
  onNext();
}


// =============================================================
// FINALISE — batch apply
// =============================================================
async function onFinalise() {
  const accepted = Object.values(sessionState.pendingChanges);

  if (accepted.length === 0) {
    showSummaryScreen();
    return;
  }

  setLoading('Applying changes to document…', 'Writing tracked changes');

  try {
    await applyAcceptedChanges(accepted);
  } catch (err) {
    console.error('Apply error:', err);
    alert('Error applying some changes: ' + err.message + '\nPartial changes may have been applied. Check the document.');
  }

  showSummaryScreen();
}

async function applyAcceptedChanges(accepted) {
  await Word.run(async (ctx) => {
    ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await ctx.sync();

    // Apply order: Deletes → Replaces → Inserts
    const ordered = [
      ...accepted.filter(c => c.change_type === 'Delete'),
      ...accepted.filter(c => c.change_type === 'Replace' || (!c.change_type && c.original_text !== 'MISSING')),
      ...accepted.filter(c => c.change_type === 'Insert' || c.original_text === 'MISSING'),
    ];

    for (const c of ordered) {
      if (c.change_type === 'Insert' || c.original_text === 'MISSING') {
        const target = await resolveInsertionAnchor(c, ctx);
        if (target) {
          target.insertParagraph(c.suggested_text, Word.InsertLocation.before);
          target.insertComment(
            `[AI REVIEWER — NEW CLAUSE | ${c.applicable_statute || ''}] ${c.insert_position_note || c.issue_summary}`
          );
        }
        continue;
      }

      // 1. Try heading-based range (verbatim heading from clauseMap — most reliable)
      let target = await findRangeByHeading(c, ctx);

      // 2. Fall back to heuristic prefix search on original_text
      if (!target) {
        target = await findRange(ctx, c.original_text);
      }

      if (target) {
        if (c.change_type === 'Delete') {
          target.insertText('', Word.InsertLocation.replace);
        } else {
          target.insertText(c.suggested_text, Word.InsertLocation.replace);
        }
        target.insertComment(`[AI REVIEWER | ${c.applicable_statute || ''}] ${c.issue_summary}`);
      } else {
        // Can't locate — leave a comment so nothing is silently lost
        ctx.document.body.paragraphs.getLast().insertComment(
          `[AI REVIEWER — APPLY MANUALLY]\nClause: ${c.clause_name}\n${c.issue_summary}\n\nSUGGESTED:\n${c.suggested_text}`
        );
      }
    }

    await ctx.sync();
    ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await ctx.sync();
  });
}

async function resolveInsertionAnchor(change, ctx) {
  const anchorsToTry = [];
  if (change.insert_anchor) anchorsToTry.push(change.insert_anchor);
  anchorsToTry.push('IN WITNESS WHEREOF');

  for (const anchor of anchorsToTry) {
    const results = ctx.document.body.search(anchor, { matchCase: false });
    results.load('items');
    await ctx.sync();
    if (results.items.length > 0) return results.items[0];
  }

  const paras = ctx.document.body.paragraphs;
  paras.load('items');
  await ctx.sync();
  return paras.getLast();
}


// =============================================================
// HIGHLIGHTING
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

async function findRange(context, originalText) {
  if (!originalText || originalText.trim().length < 5) return null;

  async function trySearch(candidate) {
    if (!candidate || candidate.trim().length < 10) return null;
    const results = context.document.body.search(candidate.trim(), {
      matchCase: false, matchWholeWord: false
    });
    results.load('items');
    await context.sync();
    return results.items.length > 0 ? results.items[0] : null;
  }

  for (const len of [originalText.length, 200, 150, 100, 60, 40]) {
    if (len > originalText.length) continue;
    const hit = await trySearch(originalText.substring(0, len));
    if (hit) return hit;
  }

  const sentences = originalText
    .split(/(?<=[.;])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 20);

  for (const sentence of sentences.slice(0, 5)) {
    const hit = await trySearch(sentence.substring(0, 120));
    if (hit) return hit;
  }

  return null;
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
    const idx = sessionState.currentClauseIndex;
    const clause = sessionState.triageJSON.clauses[idx];

    sessionState.conversationHistory.push({ role: 'user', content: userMessage });

    const messages = [
      {
        role: 'user',
        content: JSON.stringify({
          classification: sessionState.classificationJSON,
          clause: clause,
          draft_origin: sessionState.draftOrigin,
          risk_posture: sessionState.classificationJSON.risk_posture || 'neutral',
          pending_changes: Object.values(sessionState.pendingChanges),
          section_manifest: sessionState.sectionManifest,
          baseline_sections: sessionState.activeBaseline,
          defined_terms: sessionState.definedTerms
        })
      },
      ...sessionState.conversationHistory
    ];

    const responseText = await callChat(CALL_3_PROMPT, messages, 4000);
    const updated = parseJSON(responseText);

    sessionState.conversationHistory.push({ role: 'assistant', content: responseText });

    // Re-apply index/verbatim override after Ask AI update
    const seg = resolveSegment(clause);
    if (updated.original_text !== 'MISSING' && updated.change_type !== 'Insert') {
      const aiStart = updated.char_start;
      const aiEnd   = updated.char_end;
      const docText = sessionState.documentText;
      if (aiStart != null && aiEnd != null && aiEnd > aiStart
          && seg && aiStart >= seg.charStart && aiEnd <= seg.charEnd + 50) {
        updated.original_text = docText.substring(aiStart, aiEnd).trim();
      } else if (seg && seg.text) {
        updated.original_text = seg.text;
      }
    }

    // Reset cache entry so back-navigation shows updated suggestion
    sessionState.suggestionCache[idx] = updated;

    renderClauseCard(clause, updated, idx, sessionState.triageJSON.clauses.length);

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
  showScreen('screen-landing');
}

function collectModalAnswers() {
  const answers = {};
  document.querySelectorAll('.modal-question').forEach((qDiv, i) => {
    const radios = qDiv.querySelectorAll('input[type=radio]');
    if (radios.length > 0) {
      const checked = qDiv.querySelector('input[type=radio]:checked');
      if (checked) answers[`Q${i + 1}`] = checked.value;
    } else {
      const textarea = qDiv.querySelector('textarea');
      if (textarea && textarea.value.trim()) answers[`Q${i + 1}`] = textarea.value.trim();
    }
  });
  return answers;
}


// =============================================================
// CLEAN SCREEN
// =============================================================
function showCleanScreen() {
  const c = sessionState.classificationJSON;
  document.getElementById('clean-meta').innerHTML = `
    <strong>${c.classification.agreement_type}</strong><br>
    Client: ${c.perspective.inferred_perspective}<br>
    Statutes: ${(c.classification.applicable_statutes || []).slice(0, 3).join(', ')}
  `;
  document.getElementById('clean-note').textContent =
    sessionState.triageJSON.clean_agreement_note ||
    'No issues identified under Indian law for this agreement type and client position.';
  showScreen('screen-clean');
}


// =============================================================
// SUMMARY SCREEN
// =============================================================
function showSummaryScreen() {
  const clauses = sessionState.triageJSON.clauses;
  const map = sessionState.decisionMap;

  let fullAccepted = 0, fallbackAccepted = 0, rejected = 0, pending = 0, criticalResolved = 0;

  clauses.forEach((clause, idx) => {
    const d = map[idx];
    if (d === 'accepted') {
      fullAccepted++;
      if ((clause.priority || '').toUpperCase() === 'CRITICAL') criticalResolved++;
    } else if (d === 'accepted_fallback') {
      fallbackAccepted++;
      if ((clause.priority || '').toUpperCase() === 'CRITICAL') criticalResolved++;
    } else if (d === 'rejected') {
      rejected++;
    } else {
      pending++;
    }
  });

  document.getElementById('summary-split').innerHTML = `
    <div class="stat"><span>Changes accepted (full)</span><span>${fullAccepted}</span></div>
    <div class="stat"><span>Changes accepted (fallback)</span><span>${fallbackAccepted}</span></div>
    <div class="stat"><span>Rejected</span><span>${rejected}</span></div>
    <div class="stat"><span>Not reviewed</span><span>${pending}</span></div>
    <div class="stat"><span>Critical issues resolved</span><span>${criticalResolved}</span></div>
    <div class="stat"><span>Total flagged</span><span>${clauses.length}</span></div>
  `;

  const tbody = document.getElementById('summary-tbody');
  tbody.innerHTML = '';
  clauses.forEach((clause, idx) => {
    const d = map[idx];
    let decisionLabel = '—';
    let decisionClass = 'td-pending';
    if (d === 'accepted')          { decisionLabel = 'Accepted';          decisionClass = 'td-accepted'; }
    else if (d === 'accepted_fallback') { decisionLabel = 'Fallback';     decisionClass = 'td-fallback'; }
    else if (d === 'rejected')     { decisionLabel = 'Rejected';          decisionClass = 'td-rejected'; }

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${clause.clause_name || clause.clause_id || `Clause ${idx + 1}`}</td>
      <td>${clause.priority || ''}</td>
      <td class="${decisionClass}">${decisionLabel}</td>
    `;
    tbody.appendChild(tr);
  });

  showScreen('screen-summary');
}

function downloadReport() {
  const clauses = sessionState.triageJSON.clauses;
  const map = sessionState.decisionMap;
  const lines = [
    'AI CONTRACT REVIEWER — REVIEW REPORT',
    `Agreement: ${sessionState.classificationJSON.classification.agreement_type}`,
    `Client perspective: ${sessionState.classificationJSON.perspective.inferred_perspective}`,
    `Draft origin: ${sessionState.draftOrigin}`,
    `Risk posture: ${sessionState.classificationJSON.risk_posture || 'neutral'}`,
    '',
    'CLAUSE DECISIONS',
    '─'.repeat(60)
  ];

  clauses.forEach((clause, idx) => {
    const d = map[idx] || 'not reviewed';
    const suggestion = sessionState.suggestionCache[idx];
    lines.push('');
    lines.push(`${idx + 1}. ${clause.clause_name || clause.clause_id} [${clause.priority}] — ${d.toUpperCase()}`);
    if (suggestion) {
      lines.push(`   Issue: ${suggestion.issue_summary || ''}`);
      if (d === 'accepted' || d === 'accepted_fallback') {
        const appliedText = d === 'accepted_fallback' ? suggestion.fallback_text : suggestion.suggested_text;
        lines.push(`   Applied: ${(appliedText || '').substring(0, 200)}…`);
      }
    }
  });

  const blob = new Blob([lines.join('\n')], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'contract-review-report.txt';
  a.click();
}


// =============================================================
// INIT AND EVENT LISTENERS
// =============================================================
Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById('btn-start').addEventListener('click', () => runReview());

    // Clause card decisions
    document.getElementById('btn-accept').addEventListener('click', onAccept);
    document.getElementById('btn-accept-fallback').addEventListener('click', onAcceptFallback);
    document.getElementById('btn-reject').addEventListener('click', onReject);
    document.getElementById('btn-undo').addEventListener('click', onUndoDecision);
    document.getElementById('btn-accept-add-def').addEventListener('click', onAcceptAddDefinition);

    // Navigation
    document.getElementById('btn-prev').addEventListener('click', onPrev);
    document.getElementById('btn-next').addEventListener('click', onNext);

    // Finalise
    document.getElementById('btn-finalise').addEventListener('click', onFinalise);

    // Ask AI
    document.getElementById('btn-ask-ai').addEventListener('click', onAskAI);
    document.getElementById('ask-ai-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') onAskAI();
    });

    // Confidence modal
    document.getElementById('btn-modal-submit').addEventListener('click', () => {
      const answers = collectModalAnswers();
      document.getElementById('modal-confidence').classList.add('hidden');
      runReview(answers);
    });

    // Clean screen
    document.getElementById('btn-clean-done').addEventListener('click', () => showScreen('screen-landing'));

    // Summary screen
    document.getElementById('btn-download-report').addEventListener('click', downloadReport);
    document.getElementById('btn-restart').addEventListener('click', () => {
      Object.assign(sessionState, {
        classificationJSON: null,
        triageJSON: null,
        currentClauseIndex: 0,
        conversationHistory: [],
        conversationHistoryByClause: {},
        pendingChanges: {},
        rejectedChanges: [],
        suggestionCache: {},
        decisionMap: {},
        documentText: null,
        clauseMap: [],
        tableMap: [],
        definedTerms: [],
        draftOrigin: 'counterparty',
        sectionManifest: [],
        activeBaseline: [],
      });
      showScreen('screen-landing');
    });
  }
});
