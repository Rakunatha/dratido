"""
Dratido — "Draft Till Done"
────────────────────────────────────────────────
A minimal, chat-first AI drafting assistant.

There is no login, no dashboard, no case-file system — just a conversation.
The assistant walks the user through:

  1. Choosing how to start — name a document type + details, OR paste a
     template + details.
  2. Which side of the matter the draft should be written for.
  3. Free-form brainstorming / refinement of the draft with the AI.
  4. Generating the final draft (viewable in a side panel) and downloading
     it as a formatted, watermarked .docx.

AI Provider:
  Groq (free tier) — https://console.groq.com
  set GROQ_API_KEY=your_key_here

Usage:
  python dratido_app.py
"""

import os, re, time, uuid, json
import xml.sax.saxutils as _sax
from flask import Flask, request, jsonify, send_file, Response
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from pypdf import PdfReader

app = Flask(__name__)

APP_NAME    = 'Dratido'
APP_TAGLINE = 'Draft Till Done'

# In-memory conversation store: conv_id -> conversation state (no DB, no login)
CONVS = {}


# ═══════════════════════════════════════════════════════════════════════════════
#  AI CLIENT  (Groq — fast free inference)
# ═══════════════════════════════════════════════════════════════════════════════

_GROQ_PREFERRED_MODELS = [
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
    "openai/gpt-oss-120b",
    "openai/gpt-oss-20b",
]


def _get_groq_models(api_key, requests_module):
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        resp = requests_module.get(
            "https://api.groq.com/openai/v1/models",
            headers=headers,
            timeout=20,
        )
        if resp.status_code != 200:
            return [], f"HTTP {resp.status_code} from Groq /models: {resp.text[:300]}"
        data = resp.json()
        models = data.get("data", []) if isinstance(data, dict) else []
        ids = []
        for item in models:
            if not isinstance(item, dict):
                continue
            model_id = item.get("id")
            if model_id and item.get("active", True):
                ids.append(model_id)
        return ids, None
    except Exception as e:
        return [], f"Could not query Groq /models: {e}"


def _select_groq_models(api_key, requests_module):
    preferred_override = os.environ.get("GROQ_MODEL", "").strip()
    available, discovery_error = _get_groq_models(api_key, requests_module)
    available_set = set(available)

    if preferred_override:
        selected = [preferred_override]
        selected.extend(m for m in _GROQ_PREFERRED_MODELS
                        if m != preferred_override and m in available_set)
    elif available:
        selected = [m for m in _GROQ_PREFERRED_MODELS if m in available_set]
        excluded = ("whisper", "guard", "safeguard", "compound")
        selected.extend(
            m for m in available
            if m not in selected and not any(x in m.lower() for x in excluded)
        )
    else:
        selected = [preferred_override] if preferred_override else list(_GROQ_PREFERRED_MODELS)

    return list(dict.fromkeys(selected)), discovery_error


def ai_chat(messages: list, temperature: float = 0.6, max_tokens: int = 4096, attempts: int = 3) -> str:
    """Call Groq's chat-completions API with a full message list (multi-turn),
    with model fallback + exponential backoff on 429."""
    import requests as _req

    api_key = os.environ.get("GROQ_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("GROQ_API_KEY not set. Get a free key at https://console.groq.com")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json",
    }

    _GROQ_MODELS, discovery_error = _select_groq_models(api_key, _req)
    print(f"[Groq] Models selected for this key/project: {_GROQ_MODELS}")
    if discovery_error:
        print(f"[Groq] Model discovery warning: {discovery_error}")
    if not _GROQ_MODELS:
        raise RuntimeError(
            "No active Groq text models are available to this API key/project. "
            "Check Groq Project > Settings > Limits/Model Permissions and API key."
        )

    last_error = None

    for model in _GROQ_MODELS:
        payload = {
            "model":       model,
            "messages":    messages,
            "temperature": temperature,
            "max_completion_tokens": max_tokens,
            "stream":      False,
        }

        for attempt in range(attempts):
            try:
                resp = _req.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=90,
                )
            except _req.exceptions.Timeout:
                last_error = f"Timeout on {model}"
                print(f"[Groq] Timeout on {model}, trying next...")
                break
            except _req.exceptions.RequestException as e:
                last_error = f"Request error on {model}: {e}"
                print(f"[Groq] {last_error}")
                break

            status = resp.status_code

            if status == 429:
                try:
                    wait = min(float(resp.headers.get("retry-after", "")) + 1, 60)
                except (TypeError, ValueError):
                    wait = min(2 ** (attempt + 2), 45)
                last_error = f"429 rate-limited on {model} (attempt {attempt+1})"
                print(f"[Groq] 429 on {model}, waiting {wait}s...")
                time.sleep(wait)
                continue

            if status in (400, 402, 404, 503):
                body = resp.text[:300]
                last_error = f"HTTP {status} on {model}: {body}"
                print(f"[Groq] {status} on {model} (skipping): {body[:120]}")
                break

            if status != 200:
                last_error = f"HTTP {status} on {model}: {resp.text[:300]}"
                print(f"[Groq] Unexpected {status} on {model}: {resp.text[:120]}")
                break

            try:
                data = resp.json()
            except Exception as e:
                last_error = f"JSON parse error on {model}: {e}"
                print(f"[Groq] {last_error}")
                break

            if "error" in data:
                err = data["error"]
                last_error = f"API error on {model}: {err}"
                print(f"[Groq] {last_error}")
                err_str = str(err).lower()
                if "rate" in err_str or "quota" in err_str or "limit" in err_str:
                    wait = 2 ** (attempt + 2)
                    print(f"[Groq] Quota error, waiting {wait}s...")
                    time.sleep(wait)
                    continue
                break

            try:
                text = (data["choices"][0]["message"]["content"] or "").strip()
            except (KeyError, IndexError, TypeError) as e:
                last_error = f"Unexpected shape from {model}: {e}"
                print(f"[Groq] {last_error}")
                break

            if not text:
                last_error = f"Empty content from {model}"
                print(f"[Groq] {last_error}")
                break

            print(f"[Groq] \u2713 {model} ({len(text)} chars)")
            return text

        time.sleep(1)

    raise RuntimeError(
        f"All accessible Groq models failed. Last error: {last_error}. "
        "The application queried Groq /models first, so this error now reflects "
        "models visible to your current API key/project. Check Groq Project "
        "model permissions and API key at https://console.groq.com"
    )


def ai_generate(prompt: str, system: str = "", temperature: float = 0.6) -> str:
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    return ai_chat(messages, temperature=temperature)


# ═══════════════════════════════════════════════════════════════════════════════
#  TEMPLATE FILE EXTRACTION  (.docx / .pdf uploads)
# ═══════════════════════════════════════════════════════════════════════════════

ALLOWED_TEMPLATE_EXTENSIONS = {'.docx', '.pdf'}
MAX_TEMPLATE_UPLOAD_BYTES = 15 * 1024 * 1024  # 15 MB


def extract_text_from_docx(file_stream) -> str:
    """Pull readable text (paragraphs + table cells, in document order) out of an
    uploaded .docx reference template."""
    doc = Document(file_stream)
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append('\t'.join(cells))
    return '\n'.join(parts).strip()


def extract_text_from_pdf(file_stream) -> str:
    """Pull readable text out of an uploaded .pdf reference template, page by page.
    Scanned/image-only PDFs will yield little or no text — callers should treat an
    empty result as a failure and ask the user for another file."""
    reader = PdfReader(file_stream)
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt('')
        except Exception:
            pass
    parts = []
    for page in reader.pages:
        text = (page.extract_text() or '').strip()
        if text:
            parts.append(text)
    return '\n\n'.join(parts).strip()


# ═══════════════════════════════════════════════════════════════════════════════
#  MOOT MEMORIAL TEMPLATE LIBRARY  (compressed + encrypted bundled asset)
# ═══════════════════════════════════════════════════════════════════════════════
#
# A curated set of real, past moot-court memorials was analysed offline and reduced
# to short "style/structure" excerpts per section (list of abbreviations, index of
# authorities, statement of jurisdiction, statement of issues, arguments-advanced
# heading style, prayer for relief) for both the petitioner and respondent side.
# Case-specific facts and arguments were deliberately NOT retained — only formatting
# and phrasing conventions — so the library can safely inform drafting for a brand
# new, unrelated moot problem without leaking someone else's facts.
#
# The resulting JSON is compressed and encrypted at rest as an external asset file
# (moot_templates.enc) rather than shipped as a folder of plaintext PDFs/JSON in the
# repo. This is asset obfuscation, not a security boundary — the app ships with the
# key it needs to read its own asset, and anyone with the source can derive it too.
# Set MOOT_TEMPLATE_PASSPHRASE in the environment to use a different key if desired.

MOOT_TEMPLATES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'moot_templates.enc')
_MOOT_DEFAULT_PASSPHRASE = "dratido-moot-template-library-v1"
_MOOT_FIXED_SALT = b"dratido-moot-salt-2026-v1-static"

_moot_templates_cache = None


def _moot_fernet():
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    passphrase = os.environ.get("MOOT_TEMPLATE_PASSPHRASE", "").strip() or _MOOT_DEFAULT_PASSPHRASE
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=_MOOT_FIXED_SALT, iterations=200_000)
    key = __import__('base64').urlsafe_b64encode(kdf.derive(passphrase.encode('utf-8')))
    return Fernet(key)


def load_moot_templates():
    """Decrypt + decompress the bundled moot-memorial style library once per process,
    then cache it in memory. Returns [] (without raising) if the asset is missing or
    unreadable, so the moot-memorial flow degrades gracefully to a plain AI draft."""
    global _moot_templates_cache
    if _moot_templates_cache is not None:
        return _moot_templates_cache

    if not os.path.exists(MOOT_TEMPLATES_FILE):
        print(f"[Moot] Template asset not found at {MOOT_TEMPLATES_FILE}")
        _moot_templates_cache = []
        return _moot_templates_cache

    try:
        with open(MOOT_TEMPLATES_FILE, 'rb') as f:
            token = f.read()
        compressed = _moot_fernet().decrypt(token)
        import zlib
        raw = zlib.decompress(compressed)
        _moot_templates_cache = json.loads(raw.decode('utf-8'))
        print(f"[Moot] Loaded {len(_moot_templates_cache)} moot-memorial style templates")
    except Exception as e:
        print(f"[Moot] Failed to load template asset: {e}")
        _moot_templates_cache = []

    return _moot_templates_cache


def _format_moot_reference(skel: dict) -> str:
    """Render one stored skeleton's excerpts into a single labelled reference block
    for the drafting prompt."""
    labelled = [
        ("LIST OF ABBREVIATIONS (style example)", skel.get("abbreviations_sample")),
        ("INDEX OF AUTHORITIES (style example)", skel.get("authorities_style")),
        ("STATEMENT OF JURISDICTION (style example)", skel.get("jurisdiction_example")),
        ("STATEMENT OF ISSUES (style example)", skel.get("issues_style")),
        ("ARGUMENTS ADVANCED (heading/style example)", skel.get("arguments_style")),
        ("PRAYER FOR RELIEF (style example)", skel.get("prayer_example")),
    ]
    parts = [f"{label}:\n{text.strip()}" for label, text in labelled if text and text.strip()]
    return "\n\n".join(parts)


def select_moot_reference(side_key: str, user_text: str, profile: dict = None) -> str:
    """Pick the stored template that best matches the user's side and (loosely, by
    keyword overlap) their moot problem, and return it formatted as a style-reference
    block. Returns '' if the library is unavailable."""
    templates = load_moot_templates()
    if not templates:
        return ""

    matching = [t for t in templates if t.get("side") == side_key] or templates

    user_text = ((user_text or "") + " " + " ".join(((profile or {}).get("court", ""),
                 (profile or {}).get("subject", ""), MOOT_CASE_TYPES.get((profile or {}).get("case_type"), {}).get("label", "")))
                 ).strip()
    user_words = set(re.findall(r'[a-zA-Z]{4,}', user_text.lower()))
    if not user_words:
        chosen = matching[0]
    else:
        def _score(t):
            blob = ' '.join([t.get('court', ''), t.get('parties', ''),
                              t.get('issues_style', ''), t.get('arguments_style', '')]).lower()
            blob_words = set(re.findall(r'[a-zA-Z]{4,}', blob))
            return len(user_words & blob_words)
        chosen = max(matching, key=_score)

    return _format_moot_reference(chosen)


# ═══════════════════════════════════════════════════════════════════════════════
#  DOCX BUILDING
# ═══════════════════════════════════════════════════════════════════════════════




def build_ai_legal_docx(doc_type: str, ai_text: str) -> str:
    """Convert the AI-drafted plain-text legal document into a formatted,
    watermarked .docx file resembling a formal court filing."""
    doc = Document()
    for sec in doc.sections:
        sec.page_width    = Inches(8.5)
        sec.page_height   = Inches(11)
        sec.top_margin    = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin   = Inches(1.25)
        sec.right_margin  = Inches(1.25)

    TNR = 'Times New Roman'
    lines = [ln.rstrip() for ln in ai_text.strip().split('\n')]

    numbered_re   = re.compile(r'^\s*(\d{1,3})[\.\)]\s+(.*)$')
    title_written = False

    for ln in lines:
        stripped = ln.strip()
        if not stripped:
            continue
        clean = stripped.strip('#').strip()
        clean = re.sub(r'^\*\*(.*)\*\*$', r'\1', clean).strip()
        clean = clean.lstrip('*').strip()
        if not clean:
            continue

        m = numbered_re.match(clean)
        if not title_written and not m:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(16)
            r = p.add_run(clean.upper())
            r.bold = True; r.font.size = Pt(16); r.font.name = TNR
            title_written = True
            continue

        if m:
            num, body = m.group(1), m.group(2)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(4)
            p.paragraph_format.left_indent  = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            r_num = p.add_run(f'{num}.  ')
            r_num.bold = True; r_num.font.size = Pt(12); r_num.font.name = TNR
            r_body = p.add_run(body)
            r_body.font.size = Pt(12); r_body.font.name = TNR
        elif clean.isupper() and len(clean) < 80:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(8)
            r = p.add_run(clean)
            r.bold = True; r.font.size = Pt(13); r.font.name = TNR
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after  = Pt(6)
            r = p.add_run(clean)
            r.font.size = Pt(12); r.font.name = TNR

    add_watermark(doc, DRATIDO_WATERMARK_TEXT)

    os.makedirs('generated', exist_ok=True)
    safe = re.sub(r'[^\w\-]', '_', (doc_type or 'Legal_Draft')[:40]) or 'Legal_Draft'
    out  = os.path.abspath(f'generated/{safe}_{uuid.uuid4().hex[:8]}.docx')
    doc.save(out)
    return out


# ═══════════════════════════════════════════════════════════════════════════════
#  CONVERSATION / DRAFTING WORKFLOW
# ═══════════════════════════════════════════════════════════════════════════════
#
# Stages:
#   start                 -> choose "type" or "template"
#   ask_type              -> waiting for the document type (via searchable list/modal)
#   ask_facts             -> waiting for facts/details, entered via popup (type path)
#   ask_template          -> waiting for the reference template — uploaded as a .docx/.pdf
#                            file (drag & drop, browse, or device paste) or pasted as text
#   ask_template_details  -> waiting for facts/details, entered via popup (template path)
#   ask_side              -> waiting for which side the draft favours (only asked when
#                            the AI pipeline determines the document type is adversarial)
#   brainstorm            -> free-form chat with the AI; draft can be generated
#                            at any point from here on

START_BUTTONS = [
    {"label": "Specify Document Type",
     "value": "I'd like to specify the type of document and enter the details myself."},
    {"label": "Use a Reference Template",
     "value": "I'd like to provide a reference template and enter the data to fill into it."},
]

DOCUMENT_TYPES = [
    {"group": "Notices & Replies", "items": [
        "Legal Notice", "Reply to Legal Notice", "Notice under Section 80 CPC",
        "Notice under Section 138 NI Act (Cheque Bounce)", "Demand Notice",
        "Cease and Desist Letter", "Termination Notice",
    ]},
    {"group": "Civil Pleadings", "items": [
        "Plaint", "Written Statement", "Counter Claim", "Rejoinder",
        "Interlocutory Application", "Application for Interim Injunction",
        "Written Arguments", "Memo of Appeal (Civil)", "Revision Petition",
        "Review Petition", "Execution Petition",
    ]},
    {"group": "Criminal Matters (BNS, BNSS & BSA)", "items": [
        "Complaint under Section 223 BNSS (formerly Sec. 200 CrPC)",
        "FIR / Complaint to Police (Section 173 BNSS)",
        "Bail Application — Regular (Section 480 BNSS)",
        "Anticipatory Bail Application (Section 482 BNSS)",
        "Quashing Petition (Section 528 BNSS, formerly Sec. 482 CrPC)",
        "Application under Section 63 BSA (Electronic Evidence Certificate)",
        "Criminal Appeal", "Criminal Revision", "Protest Petition",
    ]},
    {"group": "Affidavits & Declarations", "items": [
        "Affidavit of Facts", "Affidavit of Income", "Affidavit of Address Proof",
        "Affidavit for Name Change", "Declaration of Marriage",
        "Solvency Affidavit", "General Declaration",
    ]},
    {"group": "Agreements & Contracts", "items": [
        "Rental / Lease Agreement", "Sale Agreement", "Partnership Deed",
        "Employment Agreement", "Non-Disclosure Agreement (NDA)",
        "Loan Agreement", "Franchise Agreement",
        "Memorandum of Understanding (MoU)", "Service Agreement",
        "Joint Venture Agreement", "Vendor Agreement",
    ]},
    {"group": "Family Law", "items": [
        "Divorce Petition (Mutual Consent)", "Divorce Petition (Contested)",
        "Maintenance Petition (Section 144 BNSS, formerly Sec. 125 CrPC)",
        "Domestic Violence Complaint", "Child Custody Petition",
        "Adoption Deed", "Will / Testament",
        "Succession Certificate Petition",
    ]},
    {"group": "Property & Real Estate", "items": [
        "Sale Deed", "Gift Deed", "Mortgage Deed",
        "Power of Attorney (General)", "Power of Attorney (Special)",
        "No Objection Certificate (NOC)", "Relinquishment Deed",
        "Partition Deed", "Lease Deed",
    ]},
    {"group": "Corporate & Commercial", "items": [
        "Memorandum of Association (MoA)", "Articles of Association (AoA)",
        "Board Resolution", "Shareholders Agreement",
        "Indemnity Bond", "Consultancy Agreement",
    ]},
    {"group": "Writs & Constitutional", "items": [
        "Writ Petition (Habeas Corpus)", "Writ Petition (Mandamus)",
        "Writ Petition (Certiorari)", "Public Interest Litigation (PIL)",
    ]},
    {"group": "Consumer & Labour", "items": [
        "Consumer Complaint", "Labour / Industrial Dispute Complaint",
        "Application under RTI Act",
    ]},
    {"group": "For Law Students", "items": [
        "Drafting Memorial for Moot Court",
    ]},
    {"group": "Other", "items": [
        "Undertaking", "Indemnity Letter", "Authorization Letter",
        "Deed of Assignment", "Statement of Case",
    ]},
]

SIDE_BUTTONS = [
    {"label": "Petitioner / Plaintiff", "value": "Petitioner / Plaintiff side"},
    {"label": "Respondent / Defendant", "value": "Respondent / Defendant side"},
    {"label": "Other — I'll specify", "value": "Other side — let me specify who this favours"},
]

MOOT_MEMORIAL_DOC_TYPE = "Drafting Memorial for Moot Court"

MOOT_SIDE_BUTTONS = [
    {"label": "Petitioner / Appellant / Plaintiff / Prosecution",
     "value": "Petitioner / Appellant / Plaintiff / Prosecution side"},
    {"label": "Respondent / Defendant / Accused",
     "value": "Respondent / Defendant / Accused side"},
]

WELCOME_MSG = (
    "Hi, I'm Dratido — your drafting assistant. I'll help you brainstorm and put together a "
    "draft, then hand you a clean Word document at the end.\n\n"
    "How would you like to start?"
)


def new_conversation() -> dict:
    conv_id = uuid.uuid4().hex
    conv = {
        "id": conv_id,
        "stage": "start",
        "mode": None,          # "type" | "template"
        "doc_type": "",
        "template_text": "",
        "template_source": "",  # human-readable note on where the template came from
        "details": "",
        "side": "",
        "messages": [],        # full transcript, for display
        "brainstorm": [],      # {role, content} sent to the LLM during brainstorming
        "draft_text": "",
        "docx_path": "",
    }
    CONVS[conv_id] = conv
    return conv


def get_conversation(conv_id: str):
    return CONVS.get(conv_id)


def push(conv, role, content, buttons=None, modal=None):
    entry = {"role": role, "content": content}
    if buttons:
        entry["buttons"] = buttons
    if modal:
        entry["modal"] = modal
    conv["messages"].append(entry)
    return entry


BRAINSTORM_SYSTEM_TMPL = (
    "You are Dratido, a collaborative AI drafting assistant. You are helping the user "
    "brainstorm and refine a legal draft before it is generated as a final document.\n\n"
    "Context for this draft:\n"
    "- Document type: {doc_type}\n"
    "- Reference template supplied by user: {has_template}\n"
    "- Facts / details supplied: {details}\n"
    "- Side this draft must favour / be enforced in favour of: {side}\n\n"
    "Your job in this chat:\n"
    "- Refer to the relevant Acts, Rules, Sections, or procedural codes where they are relevant "
    "to the document — including the current codes (BNS, BNSS, BSA) in place of any superseded "
    "ones (IPC, CrPC, Evidence Act) for criminal-law matters, alongside CPC, Contract Act, "
    "Transfer of Property Act, Companies Act, etc. as applicable — based on the user's own "
    "details.\n"
    "- If a side is specified, think and respond from that standpoint, so the draft ends up "
    "strongly and correctly serving that side's interests. If no side is specified, treat this "
    "as a neutral or bilateral document and keep your suggestions balanced.\n"
    "- Suggest structure, clauses, arguments, or missing facts that would strengthen the draft.\n"
    "- Ask short, targeted clarifying questions when something important is missing or ambiguous "
    "(e.g. which state/court, applicable limitation period, stamp duty considerations).\n"
    "- Keep replies conversational and concise (a few sentences or a short list) — this is a "
    "brainstorm, not the final document.\n"
    "- When the discussion has enough to work with, tell the user they can hit 'Generate Draft' "
    "whenever they're ready.\n"
    "- Never say you cannot help with legal matters — you are a drafting tool for the user's own "
    "professional or personal use; give substantive, practical drafting help.\n\n"
    "RESPONSE FORMAT — reply with ONLY a single valid JSON object, nothing before or after it, "
    "no markdown code fences, shaped exactly like this:\n"
    '{{"reply": "your conversational message as plain text — never use double asterisks for '
    'emphasis", "quick_replies": [{{"label": "Short button text", "value": "Full text sent if the '
    'user clicks it"}}]}}\n'
    "Populate quick_replies with 2-4 short items ONLY when your reply ends in a closed-set "
    "clarifying question that has a small number of obvious discrete answers (e.g. yes/no, or a "
    "choice between a few named options). Otherwise return an empty array for quick_replies."
)

DRAFT_SYSTEM = (
    "You are an expert legal drafter trained in formal court-filing and legal-drafting "
    "conventions. Draft a complete, professional, ready-to-use legal document in plain text "
    "(no markdown, no asterisks, no code fences).\n"
    "Draft strictly in accordance with the applicable procedure and formatting conventions — "
    "citing the relevant Acts, Sections, or procedural codes where appropriate to the document "
    "type, including the current codes (BNS, BNSS, BSA) in place of any superseded ones (IPC, "
    "CrPC, Evidence Act) for criminal-law matters, alongside CPC, Contract Act, Transfer of "
    "Property Act, Companies Act, or other applicable statute — based on the user's own "
    "details.\n"
    "Structure: a centred ALL-CAPS title on the first line (naming the document and, where "
    "appropriate, a case-number / court placeholder), then the cause-title / parties / preamble "
    "as plain paragraphs, then the operative clauses or averments as a numbered list (\"1. \", "
    "\"2. \", ...), then a prayer/relief clause where applicable, and finally a verification and "
    "signature block in the form used in formal pleadings and deeds.\n"
    "If a side is specified below, the document must be written squarely from the standpoint "
    "of, and in the interest of, that side — its framing, emphasis and relief sought should "
    "serve that side. If no side is specified, draft the document in neutral, standard form "
    "appropriate to its type (e.g. a mutual agreement, affidavit, undertaking, or declaration).\n"
    "Use precise, formal legal language appropriate to standard legal drafting practice. Output "
    "ONLY the document text — no commentary, notes, or explanations outside it."
)

DRAFT_SYSTEM_MOOT = (
    "You are an expert moot-court coach and legal drafter. Draft a complete, professional "
    "MEMORIAL (written submission) for a moot court competition, in plain text (no markdown, "
    "no asterisks, no code fences).\n"
    "Produce ALL of the following sections, in this order, each on its own ALL-CAPS heading:\n"
    "1. A cover block naming the court/forum, the parties (with '...PETITIONER'/'...APPELLANT' "
    "and '...RESPONDENT' style annotations) and 'MEMORANDUM ON BEHALF OF THE <SIDE>'.\n"
    "2. TABLE OF CONTENTS — the section names only, no page numbers.\n"
    "3. LIST OF ABBREVIATIONS — a short list of the abbreviations actually used in this memorial.\n"
    "4. INDEX OF AUTHORITIES — cases, statutes, books and web sources actually relevant to the "
    "facts and issues given (invent no fake citations; use well-known, plausible authorities for "
    "the subject matter, or generic statutory references where a specific case isn't certain).\n"
    "5. STATEMENT OF JURISDICTION — the statutory provision(s) under which this court/forum has "
    "jurisdiction, and a short formal submission sentence.\n"
    "6. STATEMENT OF FACTS — a clear, numbered, chronological account of the facts as given by "
    "the user.\n"
    "7. STATEMENT OF ISSUES — the legal issues, phrased as 'Whether ...' questions, numbered as "
    "Issue I, Issue II, etc.\n"
    "8. SUMMARY OF ARGUMENTS — a short paragraph per issue summarising the position taken.\n"
    "9. ARGUMENTS ADVANCED — the substantive legal arguments, organised issue-by-issue (I., II., "
    "...) with sub-points (A., B., ... and 1., 2., ... where useful), applying relevant statutes, "
    "sections and case law to the facts given.\n"
    "10. PRAYER FOR RELIEF — the specific relief sought, in the formal 'it is most humbly prayed "
    "...' style.\n\n"
    "The memorial must be written squarely from the standpoint of, and in the interest of, the "
    "side specified below — its framing, emphasis and relief sought should serve that side. Use "
    "precise, formal legal language in standard moot-memorial drafting conventions. Output ONLY "
    "the memorial text — no commentary, notes, or explanations outside it."
)


TEMPLATE_SWITCH_VALUE = "I'd like to provide a reference template and enter the data to fill into it."


def start_template_mode(conv):
    conv["mode"] = "template"
    conv["stage"] = "ask_template"
    push(conv, "assistant",
         "Understood — you'd like to work from a reference template. Upload a .docx or .pdf "
         "file below (drag & drop, tap to browse, or paste a copied document) — or paste the "
         "template text directly if you'd rather (placeholders like [NAME], [DATE], etc. are fine).",
         modal={"type": "upload",
                "title": "Reference Template",
                "hint": "Drop a .docx or .pdf file, tap to browse, or paste a copied document.",
                "accept": ".docx,.pdf",
                "placeholder": "Paste your template text here...",
                "submit_label": "Save Template"})


def ask_for_template_details(conv, intro_text):
    """Shared transition into the 'enter data to fill into the template' step, used by both
    the paste-your-own-template flow and the search-the-web flow."""
    conv["stage"] = "ask_template_details"
    push(conv, "assistant", intro_text,
         modal={"title": "Data for Template",
                "placeholder": "Names, dates, amounts, and other specifics...",
                "submit_label": "Save Data"})


def stage_start(conv, text):
    lower = text.lower()
    if 'template' in lower:
        start_template_mode(conv)
    else:
        conv["mode"] = "type"
        conv["stage"] = "ask_type"
        push(conv, "assistant",
             "What type of document would you like to draft? Click below to "
             "search or scroll through the list — or switch to a reference template instead.",
             modal={"type": "list",
                    "title": "Select Document Type",
                    "placeholder": "Search document types...",
                    "groups": DOCUMENT_TYPES,
                    "submit_label": "Choose Document Type"})


def stage_ask_type(conv, text):
    stripped = text.strip()
    if stripped == TEMPLATE_SWITCH_VALUE:
        start_template_mode(conv)
        return
    conv["doc_type"] = stripped

    if stripped == MOOT_MEMORIAL_DOC_TYPE:
        conv["stage"] = "ask_moot_side"
        push(conv, "assistant",
             "Great — let's put together a moot court memorial. Which side will you be "
             "arguing?", buttons=MOOT_SIDE_BUTTONS)
        return

    conv["stage"] = "ask_facts"
    push(conv, "assistant",
         f"Got it — a {conv['doc_type']}. Click below to enter the facts and details "
         f"(parties, dates, key events, amounts, relief sought — whatever you have; you can "
         f"add more later).",
         modal={"title": f"Facts & Details — {conv['doc_type']}",
                "placeholder": "Parties, dates, key events, amounts, relief sought...",
                "submit_label": "Save Details"})


def stage_ask_facts(conv, text):
    conv["details"] = text.strip()
    decide_side_stage(conv)


def stage_ask_template(conv, text):
    conv["template_text"] = text.strip()
    conv["template_source"] = "pasted by you"
    ask_for_template_details(
        conv,
        "Template received. Click below to enter the data to fill into it (names, dates, "
        "amounts, and any other specifics)."
    )


def stage_ask_template_details(conv, text):
    conv["details"] = text.strip()
    decide_side_stage(conv)


def needs_side_question(conv) -> bool:
    """Ask the AI pipeline whether this document type is inherently adversarial (so a
    favoured side needs to be picked) or neutral/bilateral (so the question can be skipped)."""
    doc_type = conv["doc_type"] or "(unspecified — inferred from the reference template)"
    context = (conv["template_text"] or conv["details"])[:1500]
    prompt = (
        f'Document type: "{doc_type}"\n'
        f'Details / template excerpt:\n"""{context}"""\n\n'
        'Does drafting this document require picking one contesting party whose interest the '
        'document should favour or enforce — as with a legal notice, plaint, written statement, '
        'reply to notice, or complaint? Or is it a neutral, bilateral, or administrative document '
        '— such as a mutual agreement, affidavit of facts, NOC, power of attorney, undertaking, or '
        'declaration — where no single side needs to be favoured?\n'
        'Reply with exactly one word: YES or NO.'
    )
    answer = ai_generate(prompt, temperature=0).strip().upper()
    return answer.startswith("Y")


def decide_side_stage(conv):
    """After facts/data are collected, decide — via the AI pipeline — whether asking which
    side the draft should favour is actually relevant, and either ask it or skip straight
    to the brainstorm."""
    try:
        needs_side = needs_side_question(conv)
    except Exception:
        needs_side = True  # safest default if the classification call fails

    if needs_side:
        conv["stage"] = "ask_side"
        push(conv, "assistant",
             "Which side is this draft for — whose interest should it be written to favour "
             "or enforce?", buttons=SIDE_BUTTONS)
    else:
        conv["side"] = ""
        conv["stage"] = "brainstorm"
        try:
            reply_text, quick_replies = run_brainstorm_turn(conv, opening=True)
        except Exception as e:
            reply_text, quick_replies = (
                f"(AI is temporarily unavailable: {e}) You can still describe what you'd "
                f"like in the draft, or click Generate Draft when ready.", [])
        push(conv, "assistant", reply_text, buttons=quick_replies)


def stage_ask_side(conv, text):
    conv["side"] = text.strip()
    conv["stage"] = "brainstorm"
    try:
        reply_text, quick_replies = run_brainstorm_turn(conv, opening=True)
    except Exception as e:
        reply_text, quick_replies = (
            f"(AI is temporarily unavailable: {e}) You can still describe what you'd "
            f"like in the draft, or click Generate Draft when ready.", [])
    push(conv, "assistant", reply_text, buttons=quick_replies)


def _enter_brainstorm(conv):
    """Shared tail used once side + facts are both known: kick off the opening
    brainstorm turn and push the assistant's reply."""
    conv["stage"] = "brainstorm"
    try:
        reply_text, quick_replies = run_brainstorm_turn(conv, opening=True)
    except Exception as e:
        reply_text, quick_replies = (
            f"(AI is temporarily unavailable: {e}) You can still describe what you'd "
            f"like in the draft, or click Generate Draft when ready.", [])
    push(conv, "assistant", reply_text, buttons=quick_replies)


def stage_ask_moot_side(conv, text):
    conv["side"] = text.strip()
    conv["stage"] = "ask_moot_facts"
    push(conv, "assistant",
         "Got it. Click below to paste the moot proposition / problem — the court or "
         "forum, the parties, the key facts, and the issues as framed (or however much "
         "of it you already have).",
         modal={"title": "Moot Problem & Facts",
                "placeholder": "Paste the moot proposition, parties, court/forum, facts, "
                                "and issues...",
                "submit_label": "Save Moot Problem"})


def stage_ask_moot_facts(conv, text):
    conv["details"] = text.strip()
    side_key = moot_side_key(conv)
    conv["moot_profile"] = moot_profile_from_text(conv["details"])
    try:
        conv["template_text"] = select_moot_reference(side_key, conv["details"], conv["moot_profile"])
    except Exception as e:
        print(f"[Moot] Template selection failed: {e}")
        conv["template_text"] = ""
    conv["template_source"] = ("Dratido moot-memorial structure library (internal, "
                                "style reference only)" if conv["template_text"] else "")
    conv["mode"] = "template"
    _enter_brainstorm(conv)


def stage_brainstorm(conv, text):
    conv["brainstorm"].append({"role": "user", "content": text})
    try:
        reply_text, quick_replies = run_brainstorm_turn(conv, opening=False)
    except Exception as e:
        reply_text, quick_replies = (
            f"(AI is temporarily unavailable: {e}) Feel free to try again, or click "
            f"Generate Draft.", [])
    push(conv, "assistant", reply_text, buttons=quick_replies)


STAGE_HANDLERS = {
    "start":                stage_start,
    "ask_type":              stage_ask_type,
    "ask_facts":             stage_ask_facts,
    "ask_template":          stage_ask_template,
    "ask_template_details":  stage_ask_template_details,
    "ask_side":              stage_ask_side,
    "ask_moot_side":         stage_ask_moot_side,
    "ask_moot_facts":        stage_ask_moot_facts,
    "brainstorm":            stage_brainstorm,
}


def run_brainstorm_turn(conv, opening=False):
    """Run one brainstorm turn. Returns (reply_text, quick_replies) where quick_replies
    is a list of {label, value} dicts suitable for rendering as clickable buttons."""
    system = BRAINSTORM_SYSTEM_TMPL.format(
        doc_type=conv["doc_type"] or "(based on the supplied template)",
        has_template="yes" if conv["template_text"] else "no",
        details=conv["details"][:3000] or "(none yet)",
        side=conv["side"] or "(not specified — draft neutrally)",
    )
    if conv["doc_type"] == MOOT_MEMORIAL_DOC_TYPE and conv.get("moot_profile"):
        system += (
            "\n\nMOOT MEMORIAL MODE. Detected profile of the problem:\n" + moot_profile_summary(conv["moot_profile"]) +
            "\nStatute policy: " + moot_statute_policy(conv["moot_profile"]) +
            "\nIGNORE the instruction above about replacing IPC/CrPC/Evidence Act with BNS/BNSS/BSA: for a moot, "
            "follow the statutes the problem itself uses. The final memorial is generated section by section "
            "(jurisdiction, facts, each sub-submission, prayer), then assembled with cover page, index of authorities "
            "and footnotes. Useful things to learn from the user: competition name, team code, case number, "
            "page/word limits, authorities they have already researched (you must use these), and any "
            "clarifications issued by the organisers. Never invent case citations in chat; if you suggest a case, "
            "say it must be verified."
        )
    messages = [{"role": "system", "content": system}]
    messages.extend(conv["brainstorm"][-16:])
    if opening:
        opener = ("Kick off the brainstorm: briefly note how you'll approach this draft, and ask "
                  "1-2 short questions if anything important is still missing.")
        if conv["doc_type"] == MOOT_MEMORIAL_DOC_TYPE and conv.get("moot_profile"):
            opener = ("Kick off: in 2-3 short sentences confirm what you detected (case type, forum, side, statute "
                      "regime, issues) and ask the user to correct anything wrong. Then ask for the competition name, "
                      "team code and any authorities they already want used (one question, short).")
        messages.append({"role": "user", "content": opener})
    raw = ai_chat(messages, temperature=0.6)
    reply_text, quick_replies = _parse_brainstorm_json(raw)
    conv["brainstorm"].append({"role": "assistant", "content": reply_text})
    return reply_text, quick_replies


def _parse_brainstorm_json(raw: str):
    """Best-effort parse of the model's structured {reply, quick_replies} JSON. Falls back
    to treating the raw text as the reply (with no quick-reply buttons) if parsing fails."""
    text = raw.strip()
    text = re.sub(r'^```(?:json)?\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
    try:
        data = json.loads(text)
        reply = str(data.get("reply", "")).strip() or raw.strip()
        raw_quick = data.get("quick_replies") or []
        quick_replies = []
        for item in raw_quick[:4]:
            if isinstance(item, dict) and item.get("label") and item.get("value"):
                quick_replies.append({
                    "label": str(item["label"])[:40],
                    "value": str(item["value"]),
                })
        return reply, quick_replies
    except Exception:
        return raw.strip(), []


def generate_draft(conv) -> str:
    side_line = conv["side"] or "(none specified — draft in neutral, standard form for this document type)"
    is_moot = False   # moot memorials use generate_moot_memorial() (sectioned pipeline)

    if is_moot:
        system = DRAFT_SYSTEM_MOOT
        if conv["template_text"]:
            prompt = (
                f'Below is a STYLE & STRUCTURE reference drawn from past moot memorials for this '
                f'side — use it only to match section order, heading conventions, phrasing style '
                f'and formal tone. Do NOT reuse any facts, party names, case citations, statutes '
                f'or numbers from it — this is a DIFFERENT case with its own facts and law.\n\n'
                f'--- STYLE & STRUCTURE REFERENCE ---\n{conv["template_text"][:6000]}\n\n'
                f'--- THE ACTUAL MOOT PROBLEM / CASE FACTS FOR THIS MEMORIAL ---\n{conv["details"]}\n\n'
                f'--- SIDE THIS MEMORIAL MUST ARGUE FOR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{_digest(conv)}\n\n'
                f'Now produce the complete final memorial text, following the required section '
                f'structure exactly.'
            )
        else:
            prompt = (
                f'--- THE MOOT PROBLEM / CASE FACTS FOR THIS MEMORIAL ---\n{conv["details"]}\n\n'
                f'--- SIDE THIS MEMORIAL MUST ARGUE FOR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{_digest(conv)}\n\n'
                f'Now produce the complete final memorial text, following the required section '
                f'structure exactly.'
            )
    else:
        system = DRAFT_SYSTEM
        if conv["template_text"]:
            prompt = (
                f'Use the following as the FORMAT/STRUCTURE reference — follow its layout, clause '
                f'structure and drafting style closely, but replace names, dates, amounts and other '
                f'details with the DATA and brainstorm notes below. Fill in any gaps sensibly.\n\n'
                f'--- FORMAT REFERENCE ---\n{conv["template_text"][:6000]}\n\n'
                f'--- DATA TO USE ---\n{conv["details"]}\n\n'
                f'--- SIDE THIS MUST FAVOUR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{_digest(conv)}\n\n'
                f'Now produce the complete final document text.'
            )
        else:
            prompt = (
                f'Draft a "{conv["doc_type"]}" document using the following details and data:\n\n'
                f'{conv["details"]}\n\n'
                f'--- SIDE THIS MUST FAVOUR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{_digest(conv)}\n\n'
                f'Produce the complete, professional, ready-to-use document text.'
            )

    draft_text = ai_generate(prompt, system=system, temperature=0.4)
    conv["draft_text"] = draft_text
    conv["docx_path"] = build_ai_legal_docx(conv["doc_type"] or "Legal_Draft", draft_text)
    return draft_text


def _digest(conv, limit_chars=3000):
    parts = []
    for m in conv["brainstorm"][-16:]:
        parts.append(f'{m["role"].upper()}: {m["content"]}')
    text = "\n".join(parts)
    return text[-limit_chars:] if text else "(no additional notes)"


# ═══════════════════════════════════════════════════════════════════════════════
#  MOOT MEMORIAL ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
#
#  Memorials are NOT written in one shot. The pipeline is:
#
#    1. profile  (at setup)   - detect case type, forum, parties, statute regime, issues
#    2. plan                  - issues -> sub-submissions (A, B, C ...) + cover details
#    3. write                 - jurisdiction, facts, summary, one call PER sub-submission, prayer
#    4. assemble              - cover page, TOC field, index of authorities (built from the
#                               real footnotes), abbreviations (built from the real text),
#                               real Word footnotes, roman/arabic page numbering
#
#  Each AI call stays small, so it fits free-tier token limits and keeps depth per section.

import threading
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as _RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part as _OpcPart
from docx.oxml.ns import qn

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ───────────────────────────── Case-type playbooks ─────────────────────────────
# The two sample memorials (a civil/family plaint and a criminal appeal) show that the
# skeleton changes with the forum. Each playbook tells the model how arguments for that
# kind of matter are normally structured.

MOOT_CASE_TYPES = {
    "civil_suit": {
        "label": "Civil suit / plaint (trial court)",
        "roles": ("PLAINTIFF", "DEFENDANT"),
        "jurisdiction": "Section 9 and Section 20 CPC (subject-matter and territorial jurisdiction), "
                        "plus the special statute the suit is brought under (e.g. s 19 Hindu Marriage Act, "
                        "s 31 Special Marriage Act, Family Courts Act 1984).",
        "skeleton": "Maintainability first (jurisdiction, cause of action, limitation, locus, bar of "
                    "suit), then merits element by element, then relief (decree, injunction, damages, "
                    "costs). Use the governing statute's own ingredients as sub-headings.",
        "prayer": "declare maintainability; adjudge/declare each issue in favour; grant the decree or "
                  "injunction sought; costs; any other order in the interest of justice, equity and good conscience.",
    },
    "criminal_appeal": {
        "label": "Criminal appeal against conviction/acquittal",
        "roles": ("APPELLANT", "RESPONDENT"),
        "jurisdiction": "The appeal provision (s 374 CrPC / s 415 BNSS for appeals against conviction, "
                        "s 378 CrPC / s 419 BNSS for appeals against acquittal), or Art 134/136 where in the "
                        "Supreme Court; state the court's appellate powers.",
        "skeleton": "(a) Appreciation of evidence and standard of proof (beyond reasonable doubt; the five "
                    "golden principles of circumstantial evidence from Hanumant / Sharad Birdhichand Sarda "
                    "where the case is circumstantial), (b) each circumstance in the chain (motive, last "
                    "seen, recovery/disclosure, identification, medical evidence), (c) ingredients of the "
                    "offence and the General Exceptions or Exceptions to the murder-definition section, "
                    "(d) supplementary charges (e.g. destruction of evidence), (e) sentence if relevant. "
                    "The appellant attacks each link; the respondent defends the trial court's findings.",
        "prayer": "appellant: allow the appeal, set aside conviction and sentence, acquit. respondent: "
                  "dismiss the appeal, uphold conviction and sentence, declare guilt established beyond reasonable doubt.",
    },
    "criminal_trial": {
        "label": "Criminal trial (Sessions / Magistrate court)",
        "roles": ("PROSECUTION", "DEFENDANT"),
        "jurisdiction": "Court of Session / Magistrate competence (s 177-209 CrPC / ss 197-232 BNSS), "
                        "committal, and the schedule classifying the offence.",
        "skeleton": "Prosecution: framing of charge, each ingredient of each offence proved by evidence, "
                    "chain of circumstances, rebuttal of defences. Defence: discharge/framing of charge "
                    "challenges, failure to prove ingredients, benefit of doubt, general exceptions, "
                    "defects in investigation, inadmissible evidence.",
        "prayer": "prosecution: frame charges / convict under the sections and sentence. defence: discharge or "
                  "acquit, benefit of doubt.",
    },
    "constitutional_writ": {
        "label": "Writ petition / PIL (Article 32 or 226)",
        "roles": ("PETITIONER", "RESPONDENT"),
        "jurisdiction": "Article 32 (Supreme Court) or Article 226 (High Court) of the Constitution of India; "
                        "mention Art 12 'State', locus standi / PIL standing if relevant.",
        "skeleton": "I. Maintainability (Art 32/226, locus standi, Art 12, alternative remedy, delay). "
                    "II.-III. Merits: violation of Arts 14, 15, 19, 21 etc. applying the exact tests "
                    "(arbitrariness, reasonable classification, proportionality per Puttaswamy, "
                    "Maneka Gandhi). Last: reliefs and the court's remedial powers.",
        "prayer": "issue a writ of certiorari / mandamus / declaration; strike down or read down the "
                  "impugned provision/action; directions; costs.",
    },
    "slp_appeal": {
        "label": "SLP / civil appeal before the Supreme Court or High Court",
        "roles": ("APPELLANT", "RESPONDENT"),
        "jurisdiction": "Article 136 (special leave) / Article 133 / s 100 or s 96 CPC; state the substantial "
                        "question of law.",
        "skeleton": "Maintainability and substantial question of law, then error(s) in the impugned "
                    "judgment issue by issue, then relief. Respondent supports the reasoning below.",
        "prayer": "appellant: set aside the impugned judgment; respondent: dismiss the appeal and uphold it.",
    },
    "family": {
        "label": "Family / personal law matter",
        "roles": ("PETITIONER", "RESPONDENT"),
        "jurisdiction": "Family Courts Act 1984 s 7; the personal-law statute's forum section (e.g. s 19 HMA); "
                        "Guardians and Wards Act 1890 for custody.",
        "skeleton": "Maintainability/jurisdiction, validity of marriage or status, the ground for relief "
                    "(cruelty, desertion, restitution), maintenance/custody on welfare-of-child principle, "
                    "constitutional challenges if raised.",
        "prayer": "decree sought (divorce / restitution / custody / maintenance), costs, any other order.",
    },
    "arbitration_commercial": {
        "label": "Arbitration / commercial dispute",
        "roles": ("PETITIONER", "RESPONDENT"),
        "jurisdiction": "Arbitration and Conciliation Act 1996 (ss 8, 9, 11, 34, 37) or Commercial Courts Act 2015.",
        "skeleton": "Existence/validity of arbitration agreement, arbitrability, seat and jurisdiction, "
                    "limited scope of judicial interference, merits on contract terms, interim relief.",
        "prayer": "refer to arbitration / set aside or uphold the award / interim relief / costs.",
    },
    "other": {
        "label": "Other (corporate, IPR, tax, service, environment, etc.)",
        "roles": ("PETITIONER", "RESPONDENT"),
        "jurisdiction": "The statute or constitutional provision that confers jurisdiction on the forum.",
        "skeleton": "Maintainability first, then merits issue by issue using the governing statute's "
                    "ingredients as sub-headings, then relief.",
        "prayer": "declare each issue in favour of the side; grant the relief; costs; any other order.",
    },
}

COMMON_ABBREVIATIONS = [
    # (abbr, full form, regex that must match somewhere in the memorial text)
    ("&", "And", r"&"),
    ("¶", "Paragraph", r"¶"),
    ("AIR", "All India Reporter", r"\bAIR\b"),
    ("All", "Allahabad", r"\bAll\b"),
    ("Art.", "Article", r"\b[Aa]rts?\.? ?\d"),
    ("BNS", "Bharatiya Nyaya Sanhita, 2023", r"\bBNS\b"),
    ("BNSS", "Bharatiya Nagarik Suraksha Sanhita, 2023", r"\bBNSS\b"),
    ("BSA", "Bharatiya Sakshya Adhiniyam, 2023", r"\bBSA\b"),
    ("Bom", "Bombay", r"\bBom\b"),
    ("Cal", "Calcutta", r"\bCal\b"),
    ("CPC", "Code of Civil Procedure, 1908", r"\bCPC\b"),
    ("CrPC", "Code of Criminal Procedure, 1973", r"\bCrPC\b"),
    ("Del", "Delhi", r"\bDel\b"),
    ("Ed.", "Edition", r"\bedn\b|\bEdn\b"),
    ("HC", "High Court", r"\bHC\b"),
    ("HMA", "Hindu Marriage Act, 1955", r"\bHMA\b"),
    ("Hon'ble", "Honourable", r"Hon[’']ble"),
    ("i.e.", "That is", r"\bi\.e\."),
    ("IEA", "Indian Evidence Act, 1872", r"\bIEA\b"),
    ("IPC", "Indian Penal Code, 1860", r"\bIPC\b"),
    ("J&K", "Jammu and Kashmir", r"\bJ&K\b"),
    ("Mad", "Madras", r"\bMad\b"),
    ("MP", "Madhya Pradesh", r"\bMP\b"),
    ("No.", "Number", r"\bNo\.\s?\d"),
    ("Ors.", "Others", r"\bOrs\.?"),
    ("P&H", "Punjab and Haryana", r"\bP&H\b"),
    ("Raj", "Rajasthan", r"\bRaj\b"),
    ("s / ss", "Section / Sections", r"\bss?\s\d"),
    ("SC", "Supreme Court", r"\bSC\b"),
    ("SCC", "Supreme Court Cases", r"\bSCC\b"),
    ("SMA", "Special Marriage Act, 1954", r"\bSMA\b"),
    ("UOI", "Union of India", r"\bUOI\b"),
    ("UP", "Uttar Pradesh", r"\bUP\b"),
    ("v", "Versus", r" v\.? "),
]


# ───────────────────────────── small helpers ─────────────────────────────

def _extract_json(raw: str):
    """Tolerant JSON extraction: strips code fences, finds the outermost {...}."""
    if not raw:
        return None
    text = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw.strip())
    for candidate in (text, text[text.find('{'):text.rfind('}') + 1] if '{' in text else ''):
        if not candidate:
            continue
        try:
            return json.loads(candidate)
        except Exception:
            # trailing commas are the most common model slip
            try:
                return json.loads(re.sub(r',\s*([}\]])', r'\1', candidate))
            except Exception:
                continue
    return None


def _clean_prose(text: str) -> str:
    """Strip markdown the model sometimes emits despite instructions."""
    t = text.replace('\r', '')
    t = re.sub(r'```[a-z]*', '', t)
    t = re.sub(r'\*\*(.+?)\*\*', r'\1', t)
    t = re.sub(r'(?m)^\s*#{1,6}\s*', '', t)
    t = re.sub(r'(?m)^\s*[-•*]\s+', '', t)
    return t.strip()


def _split_paragraphs(text: str):
    """Blank-line separated paragraphs; leading '1.' / '(a)' numbering removed because the
    builder numbers paragraphs itself."""
    out = []
    for block in re.split(r'\n\s*\n', _clean_prose(text)):
        block = ' '.join(block.split())
        block = re.sub(r'^\(?\d{1,3}[\.\)]\s+', '', block)
        if len(block) > 3:
            out.append(block)
    return out


def moot_side_key(conv) -> str:
    return "petitioner" if (conv.get("side") or "").lower().startswith("petitioner") else "respondent"


def moot_roles(profile: dict):
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    r1 = (profile.get("party1_role") or ct["roles"][0]).upper()
    r2 = (profile.get("party2_role") or ct["roles"][1]).upper()
    return r1, r2


# ───────────────────────────── 1. profile ─────────────────────────────

def moot_profile_from_text(details: str) -> dict:
    """Read the moot proposition and extract a structured case profile. Never raises."""
    keys = ", ".join(MOOT_CASE_TYPES)
    prompt = (
        "Read this moot court problem and extract a profile. Return ONLY a JSON object with exactly "
        "these keys (use \"\" or [] when the problem does not say):\n"
        f'{{"case_type": one of [{keys}], "competition": "", "team_code": "", "court": "forum in capitals, '
        'e.g. HON\'BLE HIGH COURT OF DELHI", "case_number": "", "filed_under": "provision the matter is filed under, if stated", '
        '"subject": "one-line IN THE CASE CONCERNING ... description", "party1": "first-named party", '
        '"party1_role": "PETITIONER|APPELLANT|PLAINTIFF|PROSECUTION|...", "party2": "second-named party", '
        '"party2_role": "RESPONDENT|DEFENDANT|...", "statute_regime": "old (IPC/CrPC/Evidence Act) | new '
        '(BNS/BNSS/BSA) | fictional (the problem invents its own country/codes) | unspecified", '
        '"statutes": ["statutes/codes the problem itself names"], "issues": ["issues exactly as framed in the problem, '
        'phrased Whether ..."]}}\n\n'
        f"MOOT PROBLEM:\n\"\"\"{details[:9000]}\"\"\""
    )
    data = None
    try:
        data = _extract_json(ai_generate(prompt, temperature=0))
    except Exception as e:
        print(f"[Moot] profile extraction failed: {e}")
    data = data if isinstance(data, dict) else {}
    prof = {
        "case_type": data.get("case_type") if data.get("case_type") in MOOT_CASE_TYPES else "other",
        "competition": str(data.get("competition") or ""),
        "team_code": str(data.get("team_code") or ""),
        "court": str(data.get("court") or ""),
        "case_number": str(data.get("case_number") or ""),
        "filed_under": str(data.get("filed_under") or ""),
        "subject": str(data.get("subject") or ""),
        "party1": str(data.get("party1") or ""),
        "party1_role": str(data.get("party1_role") or ""),
        "party2": str(data.get("party2") or ""),
        "party2_role": str(data.get("party2_role") or ""),
        "statute_regime": str(data.get("statute_regime") or "unspecified"),
        "statutes": [str(s) for s in (data.get("statutes") or [])][:12],
        "issues": [str(s) for s in (data.get("issues") or [])][:6],
    }
    return prof


def moot_profile_summary(profile: dict) -> str:
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    r1, r2 = moot_roles(profile)
    lines = [
        f"Case type: {ct['label']}",
        f"Forum: {profile.get('court') or '(not stated)'}",
        f"Parties: {profile.get('party1') or '?'} ({r1}) v {profile.get('party2') or '?'} ({r2})",
        f"Statute regime: {profile.get('statute_regime')}"
        + (f" — {', '.join(profile['statutes'])}" if profile.get("statutes") else ""),
    ]
    if profile.get("issues"):
        lines.append("Issues as framed: " + " | ".join(profile["issues"]))
    return "\n".join(lines)


def moot_statute_policy(profile: dict) -> str:
    regime = (profile.get("statute_regime") or "").lower()
    named = ", ".join(profile.get("statutes") or []) or "the codes named in the problem"
    if regime.startswith("fictional"):
        return ("The problem is set in a fictional jurisdiction with its own codes. Use the problem's own "
                f"names and section numbers exactly as given ({named}); treat real Indian/Commonwealth case "
                "law as persuasive authority and say so. Never rename the fictional statutes.")
    if regime.startswith("old"):
        return (f"The problem uses the older codes ({named}). Cite those sections exactly as the problem does; "
                "optionally add the new-code equivalent in square brackets on first mention only.")
    if regime.startswith("new"):
        return (f"The problem uses the new codes ({named}). Cite BNS/BNSS/BSA sections; add the old "
                "IPC/CrPC/Evidence Act equivalent in square brackets on first mention only if you are certain of it.")
    return ("Follow whichever statutes the problem names. If it names none and the matter is a real Indian "
            "criminal matter after 1 July 2024, cite BNS/BNSS/BSA with the IPC/CrPC/Evidence Act equivalent in "
            "square brackets on first mention.")


# ───────────────────────────── 2/3. writing prompts ─────────────────────────────

MOOT_WRITER_SYSTEM = (
    "You are a senior Indian moot-court advocate and memorial coach writing ONE part of a competition "
    "memorial. Write in formal, persuasive memorial register ('It is humbly submitted...', 'It is most "
    "respectfully contended...'), always squarely for the assigned side, never conceding the other side's case.\n"
    "ABSOLUTE RULES:\n"
    "1. Plain text only: no markdown, no asterisks, no headings, no numbering, no bullet points.\n"
    "2. Separate paragraphs with one blank line.\n"
    "3. FOOTNOTES: put a footnote marker {{fn: citation}} immediately after the punctuation of any sentence that "
    "relies on a case, statute, book or fact in the problem. Formats: "
    "case — Party v Party (Year) Vol SCC Page   or   Party v Party AIR Year SC Page ; "
    "statute — Code of Civil Procedure 1908, s 20 (statute title first, then the section) ; "
    "constitution — Constitution of India 1950, art 21 ; "
    "fact — Moot Proposition ¶ N (only when the problem has numbered paragraphs, otherwise 'Moot Proposition') ; "
    "book — Author, Title (nth edn, Publisher Year) page. One authority per marker; use ';' only to separate two authorities.\n"
    "4. CITATION HONESTY: cite only decisions you are highly confident exist and whose holding you "
    "remember correctly, with the citation you are confident of. If you remember the case but not the exact "
    "reporter citation, give name, court and year and begin the marker with '(?) '. NEVER invent a case, "
    "citation, quotation or section number. If the user's notes supply authorities, use those. A memorial with "
    "fewer, accurate authorities beats many doubtful ones.\n"
    "5. Describe each authority's holding accurately and briefly ('In X v Y, the Supreme Court held that ...'), then "
    "APPLY it to the facts of THIS problem using the parties' names. Do not copy facts, names or citations from any "
    "style reference supplied."
)


def _moot_context_block(conv, profile, notes_limit=1800, style_chars=0):
    r1, r2 = moot_roles(profile)
    side_key = moot_side_key(conv)
    my_role, other_role = (r1, r2) if side_key == "petitioner" else (r2, r1)
    my_party = profile.get("party1") if side_key == "petitioner" else profile.get("party2")
    other_party = profile.get("party2") if side_key == "petitioner" else profile.get("party1")
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    notes = _digest(conv, limit_chars=notes_limit)
    return (
        f"SIDE YOU ARGUE FOR: {my_role} ({my_party or 'as named in the problem'}). Opposing side: {other_role} "
        f"({other_party or 'as named in the problem'}).\n"
        f"CASE TYPE: {ct['label']}\n"
        f"STATUTE POLICY: {moot_statute_policy(profile)}\n\n"
        f"MOOT PROBLEM (the only source of facts):\n\"\"\"{conv['details'][:8000]}\"\"\"\n\n"
        f"TEAM NOTES / USER-SUPPLIED AUTHORITIES (from the brainstorm chat):\n{notes}"
        + (("\n\nSTYLE REFERENCE from past memorials of this side (match tone, heading and phrasing conventions ONLY; "
            "it is a DIFFERENT case — never reuse its facts, names, cases or statutes):\n"
            + conv["template_text"][:style_chars]) if style_chars and conv.get("template_text") else "")
    )


def _moot_call(prompt, max_tokens=2000, temperature=0.35, system=MOOT_WRITER_SYSTEM):
    messages = [{"role": "system", "content": system}, {"role": "user", "content": prompt}]
    return ai_chat(messages, temperature=temperature, max_tokens=max_tokens, attempts=5)


def moot_make_plan(conv, profile) -> dict:
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    framed = profile.get("issues") or []
    max_sub = int(os.environ.get("MOOT_MAX_SUBHEADS", "10"))
    prompt = (
        _moot_context_block(conv, profile) + "\n\n"
        f"PLAYBOOK FOR THIS KIND OF MATTER:\n- Forum/jurisdiction: {ct['jurisdiction']}\n- Argument skeleton: {ct['skeleton']}\n\n"
        + ("ISSUES AS FRAMED IN THE PROBLEM (keep them, in this order, in the same wording): "
           + " | ".join(framed) + "\n\n" if framed else
           "The problem frames no issues: frame 2-4 issues yourself, each starting 'Whether ...?'.\n\n")
        + "TASK: plan the memorial. Return ONLY JSON of this shape:\n"
        '{"cover": {"competition": "", "team_code": "", "case_number": "", "filed_under": "", "subject": ""},\n'
        ' "jurisdiction_provisions": ["provision that confers jurisdiction, e.g. Article 226 of the Constitution of India"],\n'
        ' "abbreviations": [["short form you will use", "full form"]],\n'
        ' "issues": [{"text": "Whether ...?", "submission": "That <one-sentence thesis for our side, normal sentence case>", '
        '"subs": [{"title": "sub-submission heading in normal sentence case, e.g. Section 20 of CPC is applicable", "points": "2-3 sentences: the rule, the key authorities '
        'you will use, and the facts that apply"}]}],\n'
        ' "prayer": ["declaration/relief 1", "relief 2"]}\n'
        f"Rules: 2-4 issues; 2-4 subs per issue; at most {max_sub} subs in total; if the user's team notes give a "
        "competition name, team code, case number or authorities, use them in the cover/points; cover fields "
        "the problem does not state stay empty strings. Arrange subs in the order a court would take them "
        "(maintainability/preliminary points first, relief last)."
    )
    plan = None
    for attempt in range(2):
        raw = _moot_call(prompt if attempt == 0 else
                         "Your previous reply was not valid JSON. Return the SAME plan again as ONLY valid JSON, "
                         "no commentary, no code fences.\n\n" + prompt,
                         max_tokens=2600, temperature=0.2, system="You output only valid JSON.")
        plan = _extract_json(raw)
        if isinstance(plan, dict) and plan.get("issues"):
            break
        plan = None

    if plan is None:  # deterministic fallback so generation never dies on a bad JSON reply
        issues_src = framed or ["Whether the present matter is maintainable before this Hon'ble Court?",
                                "Whether the " + ("petitioner's" if moot_side_key(conv) == "petitioner" else "respondent's")
                                + " case is made out on merits?"]
        plan = {"cover": {}, "jurisdiction_provisions": [], "abbreviations": [], "prayer": [],
                "issues": [{"text": t, "submission": "That " + t.rstrip("?").replace("Whether", "", 1).strip(),
                            "subs": [{"title": "THE LEGAL POSITION", "points": "Governing rule and authorities."},
                                     {"title": "APPLICATION TO THE FACTS", "points": "Apply the rule to the facts."}]}
                           for t in issues_src[:4]]}

    # normalise + cap
    issues, total = [], 0
    for i, it in enumerate((plan.get("issues") or [])[:4]):
        text = str(it.get("text") or (framed[i] if i < len(framed) else f"Whether issue {i + 1}?")).strip()
        subs = []
        for s in (it.get("subs") or [])[:4]:
            if total >= max_sub:
                break
            subs.append({"title": str(s.get("title") or "SUBMISSION").strip().rstrip('.'),
                         "points": str(s.get("points") or "")})
            total += 1
        if not subs:
            subs = [{"title": "THE SUBMISSION", "points": ""}]
            total += 1
        issues.append({"text": text, "submission": str(it.get("submission") or "").strip(), "subs": subs})
    plan["issues"] = issues
    plan["cover"] = plan.get("cover") if isinstance(plan.get("cover"), dict) else {}
    plan["jurisdiction_provisions"] = [str(p) for p in (plan.get("jurisdiction_provisions") or [])][:6]
    plan["abbreviations"] = [a for a in (plan.get("abbreviations") or []) if isinstance(a, (list, tuple)) and len(a) == 2][:20]
    plan["prayer"] = [str(p) for p in (plan.get("prayer") or [])][:8]
    return plan


def moot_write_jurisdiction(conv, profile, plan) -> list:
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    r1, r2 = moot_roles(profile)
    my_role = r1 if moot_side_key(conv) == "petitioner" else r2
    prov = "; ".join(plan.get("jurisdiction_provisions") or []) or ct["jurisdiction"]
    prompt = (
        _moot_context_block(conv, profile, notes_limit=800, style_chars=1400) + "\n\n"
        "TASK: write the STATEMENT OF JURISDICTION: 1-2 paragraphs. First paragraph: 'The "
        f"{my_role.title()} has invoked / the Respondent submits to the jurisdiction of this Hon'ble Court under ...' naming "
        f"the exact provisions that give THIS forum jurisdiction (likely: {prov}). Put each provision in a footnote marker "
        "in the form {{fn: Statute title, s N: short description of what the section provides}} - quote the section text "
        "ONLY if you are certain of it, otherwise describe it in one line. Final sentence: that the party "
        "'most humbly submits to the jurisdiction of this Hon'ble Court'. No headings."
    )
    return _split_paragraphs(_moot_call(prompt, max_tokens=1100))


def moot_write_facts(conv, profile) -> list:
    prompt = (
        _moot_context_block(conv, profile, notes_limit=600) + "\n\n"
        "TASK: write the STATEMENT OF FACTS: a chronological account of 6-14 short numbered-style paragraphs "
        "(do not number them; separate with blank lines). Use ONLY facts in the problem; do not invent or "
        "embellish; do not argue or use adjectives that characterise motives. Choose and order facts so that "
        "those favourable to your side are prominent, but state adverse facts accurately and neutrally. "
        "Third person, past tense, parties by name. Where the problem's paragraphs are numbered, end the "
        "relevant sentence with {{fn: Moot Proposition ¶ N}}; otherwise use no footnotes."
    )
    return _split_paragraphs(_moot_call(prompt, max_tokens=1800, temperature=0.2))


def moot_write_summary(conv, profile, plan) -> list:
    issues = "\n".join(f"Issue {i + 1}: {it['text']}  | thesis: {it['submission']} | subs: "
                       + "; ".join(s['title'] for s in it['subs']) for i, it in enumerate(plan["issues"]))
    prompt = (
        _moot_context_block(conv, profile, notes_limit=500) + "\n\nPLAN:\n" + issues + "\n\n"
        f"TASK: write the SUMMARY OF ARGUMENTS: exactly {len(plan['issues'])} paragraphs, one per issue in order, "
        "each 80-130 words, each summarising our side's position, its 2-3 strongest authorities by name, and the key "
        "facts. Separate paragraphs with a blank line. Do not repeat the issue text and do not label the paragraphs. "
        "Footnote markers are allowed but keep them few."
    )
    paras = _split_paragraphs(_moot_call(prompt, max_tokens=1600))
    n = len(plan["issues"])
    if len(paras) > n:  # merge extras into the last so no content is dropped
        paras = paras[:n - 1] + [' '.join(paras[n - 1:])]
    while len(paras) < n:
        paras.append("It is humbly submitted that the submissions on this issue are set out in full in the Arguments Advanced.")
    return paras


def moot_write_sub(conv, profile, plan, issue, sub, cited_names) -> list:
    avoid = ("Authorities already used elsewhere in this memorial (do not lean on them again unless essential): "
             + "; ".join(cited_names[-30:]) + "\n") if cited_names else ""
    prompt = (
        _moot_context_block(conv, profile) + "\n\n"
        f"ISSUE: {issue['text']}\nOUR THESIS ON THIS ISSUE: {issue['submission']}\n"
        f"SUB-SUBMISSION YOU ARE WRITING: {sub['title']}\nPLANNED POINTS: {sub['points']}\n{avoid}\n"
        "TASK: write the body of this sub-submission: 5-8 paragraphs of 90-150 words each, in this rhythm: "
        "(1) state the legal rule with the statutory provision; (2) 2-4 authorities with their accurate holdings; "
        "(3) apply the rule to the facts of THIS problem explicitly, citing the facts with {{fn: Moot Proposition ¶ N}}; "
        "(4) pre-empt the opposing side's likely contention and answer it; (5) a closing paragraph concluding this "
        "sub-submission. Every authority or fact reference must carry a footnote marker as per the rules. "
        "Do not write a heading."
    )
    return _split_paragraphs(_moot_call(prompt, max_tokens=2300))


def moot_write_prayer(conv, profile, plan) -> list:
    ct = MOOT_CASE_TYPES.get(profile.get("case_type"), MOOT_CASE_TYPES["other"])
    issues = " | ".join(it['text'] for it in plan["issues"])
    hint = "; ".join(plan.get("prayer") or [])
    prompt = (
        _moot_context_block(conv, profile, notes_limit=400, style_chars=900) + "\n\n"
        f"ISSUES: {issues}\nPRAYER IDEAS FROM THE PLAN: {hint}\nPLAYBOOK PRAYER STYLE: {ct['prayer']}\n\n"
        "TASK: write the operative prayer items as 3-6 short paragraphs separated by blank lines, each a complete "
        "declaration or relief that the court is asked to grant to OUR side (one per issue, then costs/other relief). "
        "Do not number them; do not write the opening 'Wherefore...' or the closing line, only the items."
    )
    return _split_paragraphs(_moot_call(prompt, max_tokens=900, temperature=0.25))


# ───────────────────────────── footnote handling ─────────────────────────────

_FN_RE = re.compile(r'\{\{\s*fn\s*:\s*(.*?)\s*\}\}', re.S)
_SKIP_AUTH_RE = re.compile(r'^(ibid|id\b|supra|moot proposition|statement of facts|facts\b|clarification|n\s?\d|\(n\s?\d)', re.I)


def _parse_marked(text: str):
    """Split 'text {{fn: x}} more' into [('t','text '),('fn','x'),('t',' more')]."""
    parts, pos = [], 0
    for m in _FN_RE.finditer(text):
        if m.start() > pos:
            parts.append(('t', text[pos:m.start()]))
        parts.append(('fn', m.group(1).strip()))
        pos = m.end()
    if pos < len(text):
        parts.append(('t', text[pos:]))
    # a stray unclosed marker must never leak into the document
    return [(k, re.sub(r'\{\{.*$', '', v) if k == 't' else v) for k, v in parts]


class _Footnotes:
    def __init__(self):
        self.items = []   # [(id, text, uncertain, block_key)]

    def add(self, text, block_key=None):
        uncertain = text.startswith('(?)')
        text = text.replace('(?)', '', 1).strip() if uncertain else text.strip()
        fid = len(self.items) + 1
        self.items.append((fid, text, uncertain, block_key))
        return fid

    def part_xml(self) -> bytes:
        def esc(s):
            return _sax.escape(s)
        body = [
            f'<w:footnotes xmlns:w="{W_NS}">',
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:separator/></w:r></w:p></w:footnote>',
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
        ]
        for fid, text, _u, _k in self.items:
            body.append(
                f'<w:footnote w:id="{fid}"><w:p><w:pPr><w:spacing w:after="40" w:line="240" w:lineRule="auto"/>'
                '<w:jc w:val="both"/></w:pPr>'
                '<w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:footnoteRef/></w:r>'
                f'<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> {esc(text)}</w:t></w:r></w:p></w:footnote>')
        body.append('</w:footnotes>')
        return ''.join(body).encode('utf-8')


def _attach_footnotes(doc, fb: _Footnotes):
    part = _OpcPart(PackURI('/word/footnotes.xml'),
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
                    fb.part_xml(), doc.part.package)
    doc.part.relate_to(part, _RT.FOOTNOTES)
    settings = doc.settings.element
    fpr = parse_xml(f'<w:footnotePr xmlns:w="{W_NS}"><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>')
    anchor = next((c for c in settings if c.tag in (qn('w:endnotePr'), qn('w:compat'))), None)
    settings.insert(list(settings).index(anchor), fpr) if anchor is not None else settings.append(fpr)


def _set_update_fields(doc):
    settings = doc.settings.element
    el = parse_xml(f'<w:updateFields xmlns:w="{W_NS}" w:val="true"/>')
    later = (qn('w:hdrShapeDefaults'), qn('w:footnotePr'), qn('w:endnotePr'), qn('w:compat'))
    anchor = next((c for c in settings if c.tag in later), None)
    settings.insert(list(settings).index(anchor), el) if anchor is not None else settings.append(el)


# ───────────────────────────── authority index ─────────────────────────────

_PIN_RE = re.compile(r',?\s*(?:at\s+)?(?:paras?\.?|paragraphs?|¶|pp?\.?)\s*[\d\-–,\s]+$', re.I)
_STAT_RE = re.compile(r'\b(Act|Code|Constitution|Sanhita|Adhiniyam|Rules|Regulations|Order|Ordinance|Bill)\b')
_CASE_RE = re.compile(r'\sv\.?\s|\svs\.?\s|\sversus\s', re.I)
_BOOK_RE = re.compile(r'\bedn\b|\bedition\b|\bcommentary\b|\btreatise\b', re.I)


def _classify_authority(raw: str):
    """-> (kind, display) or None when the footnote is only a cross-reference to the facts."""
    t = raw.strip().rstrip('.').strip()
    if not t or _SKIP_AUTH_RE.match(t):
        return None
    if re.search(r'https?://|www\.', t):
        return ('web', t)
    if _CASE_RE.search(t) and not _BOOK_RE.search(t):
        return ('case', _PIN_RE.sub('', t).strip().rstrip(',').strip())
    if _BOOK_RE.search(t):
        return ('book', re.sub(r',\s*\d+(?:[\-–]\d+)?$', '', t))
    if _STAT_RE.search(t):
        t2 = t.split(':')[0]
        m = re.match(r'^(?:ss?|sections?|arts?|articles?|rr?|rules?|orders?)\.?\s*[\w()\-, ]+?,\s*(.+)$', t2, re.I)
        if m:
            t2 = m.group(1)
        t2 = re.split(r',\s*(?:ss?|sections?|arts?|articles?|rr?|rules?|orders?|schedule|chapter|part)\b', t2, 1, flags=re.I)[0]
        return ('statute', t2.strip().rstrip(',').strip())
    return None


def build_authority_index(fb: _Footnotes):
    """Collect authorities from the real footnotes. Returns (index, verify) where index maps
    kind -> [(display, [block_keys])] sorted alphabetically."""
    found = {'case': {}, 'statute': {}, 'book': {}, 'web': {}}
    verify = {}
    for _fid, text, uncertain, block in fb.items:
        chunks = [text] if '"' in text or '“' in text else [c for c in text.split(';')]
        for chunk in chunks:
            res = _classify_authority(chunk)
            if not res:
                continue
            kind, disp = res
            key = re.sub(r'\W+', ' ', disp.lower()).strip()
            entry = found[kind].setdefault(key, [disp, []])
            if block and block not in entry[1]:
                entry[1].append(block)
            if kind == 'case':
                verify[key] = (disp, uncertain or verify.get(key, (None, False))[1])
    index = {k: sorted(((d, b) for d, b in v.values()), key=lambda x: x[0].lower()) for k, v in found.items()}
    verify_list = sorted(verify.values(), key=lambda x: (not x[1], x[0].lower()))
    return index, verify_list


# ───────────────────────────── low-level docx helpers ─────────────────────────────

_SECTPR_ORDER = ['footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
                 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi',
                 'rtlGutter', 'docGrid', 'printerSettings']


def _sectpr_insert(sectPr, el):
    name = el.tag.split('}')[1]
    for old in sectPr.findall(qn('w:' + name)):
        sectPr.remove(old)
    idx = _SECTPR_ORDER.index(name)
    for child in list(sectPr):
        cname = child.tag.split('}')[1]
        if cname in _SECTPR_ORDER and _SECTPR_ORDER.index(cname) > idx:
            child.addprevious(el)
            return
    sectPr.append(el)


def _set_run_font(run, size=12, bold=None, italic=None, caps=False, small_caps=False, color=None):
    run.font.name = 'Times New Roman'
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = parse_xml(f'<w:rFonts xmlns:w="{W_NS}"/>')
        rpr.insert(0, rfonts)
    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        rfonts.set(qn('w:' + a), 'Times New Roman')
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if caps:
        run.font.all_caps = True
    if small_caps:
        run.font.small_caps = True
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _style_base(doc):
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(12)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = parse_xml(f'<w:rFonts xmlns:w="{W_NS}"/>')
        rpr.insert(0, rf)
    for a in list(rf.attrib):
        del rf.attrib[a]
    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        rf.set(qn('w:' + a), 'Times New Roman')
    for name, size, before, after in (('Heading 1', 14, 0, 12), ('Heading 2', 12, 12, 8), ('Heading 3', 12, 10, 6)):
        h = doc.styles[name]
        h.font.name = 'Times New Roman'
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.italic = False
        h.font.color.rgb = RGBColor(0, 0, 0)
        hr = h.element.get_or_add_rPr()
        hf = hr.find(qn('w:rFonts'))
        if hf is None:
            hf = parse_xml(f'<w:rFonts xmlns:w="{W_NS}"/>')
            hr.insert(0, hf)
        for a in list(hf.attrib):
            del hf.attrib[a]
        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            hf.set(qn('w:' + a), 'Times New Roman')
        h.paragraph_format.space_before = Pt(before)
        h.paragraph_format.space_after = Pt(after)
        h.paragraph_format.keep_with_next = True


def _pborder(p, side='bottom', sz=18, space=4, val='single'):
    ppr = p._p.get_or_add_pPr()
    bd = ppr.find(qn('w:pBdr'))
    if bd is None:
        bd = parse_xml(f'<w:pBdr xmlns:w="{W_NS}"/>')
        ppr.append(bd)
    bd.append(parse_xml(f'<w:{side} xmlns:w="{W_NS}" w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="000000"/>'))


def _fld(p, instr, cached='', size=12, bold=None):
    """Insert a complex field (begin / instr / separate / cached / end) into paragraph p."""
    def mk(xml):
        r = p.add_run()
        r._r.append(parse_xml(xml))
        _set_run_font(r, size=size, bold=bold)
        return r
    mk(f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="begin"/>')
    r = p.add_run()
    it = parse_xml(f'<w:instrText xmlns:w="{W_NS}" xml:space="preserve"> {_sax.escape(instr)} </w:instrText>')
    r._r.append(it)
    mk(f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="separate"/>')
    rc = p.add_run(cached)
    _set_run_font(rc, size=size, bold=bold)
    mk(f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="end"/>')


class _Bm:
    n = 100


def _bookmark_wrap(p, name):
    _Bm.n += 1
    start = parse_xml(f'<w:bookmarkStart xmlns:w="{W_NS}" w:id="{_Bm.n}" w:name="{name}"/>')
    end = parse_xml(f'<w:bookmarkEnd xmlns:w="{W_NS}" w:id="{_Bm.n}"/>')
    ppr = p._p.pPr
    if ppr is not None:
        ppr.addnext(start)
    else:
        p._p.insert(0, start)
    p._p.append(end)


def _rich(p, text, fb: _Footnotes, block_key=None, size=12, bold=False):
    """Add text with {{fn: ...}} markers as runs + real Word footnote references."""
    for kind, val in _parse_marked(text):
        if kind == 't':
            if val:
                r = p.add_run(val)
                _set_run_font(r, size=size, bold=bold)
        else:
            if not val:
                continue
            fid = fb.add(val, block_key)
            r = p.add_run()
            r._r.append(parse_xml(f'<w:footnoteReference xmlns:w="{W_NS}" w:id="{fid}"/>'))
            rpr = r._r.get_or_add_rPr()
            rpr.append(parse_xml(f'<w:vertAlign xmlns:w="{W_NS}" w:val="superscript"/>'))


def _body_para(doc, text, fb, block_key=None, num=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
               spacing=1.5, indent_first=False):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = spacing
    pf.space_after = Pt(6)
    pf.widow_control = True
    if num is not None:
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.5)
        pf.tab_stops.add_tab_stop(Inches(0.5))
        r = p.add_run(f'{num}.\t')
        _set_run_font(r, bold=True)
    elif indent_first:
        pf.first_line_indent = Inches(0.5)
    _rich(p, text, fb, block_key)
    if block_key and block_key.startswith('cite_') and '{{' in text:
        _bookmark_wrap(p, block_key)
    return p


def _heading(doc, text, level=1, align=None, page_break=False):
    h = doc.add_heading('', level=level)
    if page_break:
        h.paragraph_format.page_break_before = True
    h.alignment = align if align is not None else (WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT)
    r = h.add_run(text)
    _set_run_font(r, size={1: 14, 2: 12, 3: 12}[level], bold=True, color='000000')
    if level >= 2:
        h.paragraph_format.left_indent = Inches(0.4)
        h.paragraph_format.first_line_indent = Inches(-0.4)
    return h


def _text_width_in(doc):
    s = doc.sections[0]
    return (s.page_width - s.left_margin - s.right_margin) / 914400


# ───────────────────────────── cover page ─────────────────────────────

def _cover_page(doc, cover: dict, side_title: str, roles):
    width = _text_width_in(doc)
    # team code box (right aligned)
    t = doc.add_table(rows=1, cols=1, style='Table Grid')
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    t.autofit = False
    cell = t.rows[0].cells[0]
    cell.width = Inches(2.0)
    t.columns[0].width = Inches(2.0)
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tblPr.append(parse_xml(f'<w:tblW xmlns:w="{W_NS}" w:w="2880" w:type="dxa"/>'))
    cp = cell.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cp.add_run(f"Team Code: {cover.get('team_code') or '________'}")
    _set_run_font(cr, size=11)

    def gap(pts):
        g = doc.add_paragraph()
        g.paragraph_format.space_after = Pt(pts)
        return g

    def line(text, size=12, bold=False, italic=False, rule_top=False, rule_bottom=False, small_caps=False,
             before=6, after=6, align=WD_ALIGN_PARAGRAPH.CENTER):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        if rule_top:
            _pborder(p, 'top', sz=36, space=8)
        if rule_bottom:
            _pborder(p, 'bottom', sz=36, space=8)
        r = p.add_run(text)
        _set_run_font(r, size=size, bold=bold, italic=italic, small_caps=small_caps)
        return p

    gap(24)
    if cover.get('competition'):
        line(re.sub(r'(\d)(ST|ND|RD|TH)\b', lambda m: m.group(1) + m.group(2).lower(), cover['competition'].upper()), size=13, bold=True, rule_top=True, rule_bottom=True, before=10, after=10)
    gap(6)
    line('Before', italic=True)
    line((cover.get('court') or "THE HON'BLE COURT").upper(), size=12, rule_bottom=True, after=14)
    line(cover.get('case_number') or 'CASE NO. ________ OF 20__', small_caps=True, after=10)
    if cover.get('filed_under'):
        line(cover['filed_under'], size=12, rule_top=True, rule_bottom=True, before=10, after=10)
    if cover.get('subject'):
        line(cover['subject'].upper(), size=11, before=14, after=14)
    gap(6)
    line('IN THE MATTER BETWEEN:', size=11, after=14)

    def party(name, role, top=False, bottom=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if top:
            _pborder(p, 'top', sz=12, space=6, val='double')
        if bottom:
            _pborder(p, 'bottom', sz=12, space=6, val='double')
        r = p.add_run(f"{(name or '________').upper()}\t{role}")
        _set_run_font(r, size=12)

    party(cover.get('party1'), roles[0], top=True)
    line('v', size=11, before=10, after=10)
    party(cover.get('party2'), roles[1], bottom=True)
    gap(30)
    line(f"MEMORIAL ON BEHALF OF {side_title}", size=13, bold=True, rule_top=True, rule_bottom=True, before=16, after=16)


# ───────────────────────────── headers / footers / sections ─────────────────────────────

def _make_header_footer(section, side_title, team_code, width_in):
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    hp = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
    for r in list(hp.runs):
        r._r.getparent().remove(r._r)
    hp.style = 'Normal'
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _pborder(hp, 'bottom', sz=8, space=2)
    r = hp.add_run(f"MEMORIAL ON BEHALF OF {side_title}")
    _set_run_font(r, size=10, bold=True)

    fp = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
    for r in list(fp.runs):
        r._r.getparent().remove(r._r)
    fp.style = 'Normal'
    fp.paragraph_format.tab_stops.add_tab_stop(Inches(width_in), WD_TAB_ALIGNMENT.RIGHT)
    _pborder(fp, 'top', sz=8, space=2)
    r = fp.add_run(f"Team Code: {team_code or '________'}\tPage | ")
    _set_run_font(r, size=10)
    _fld(fp, 'PAGE', '1', size=10)


def _blank_header_footer(section):
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False


def _page_numbering(section, fmt):
    _sectpr_insert(section._sectPr, parse_xml(
        f'<w:pgNumType xmlns:w="{W_NS}" w:fmt="{fmt}" w:start="1"/>'))


def _page_border(section):
    _sectpr_insert(section._sectPr, parse_xml(
        f'<w:pgBorders xmlns:w="{W_NS}" w:offsetFrom="page">'
        '<w:top w:val="thinThickSmallGap" w:sz="24" w:space="24" w:color="000000"/>'
        '<w:left w:val="thinThickSmallGap" w:sz="24" w:space="24" w:color="000000"/>'
        '<w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="24" w:color="000000"/>'
        '<w:right w:val="thickThinSmallGap" w:sz="24" w:space="24" w:color="000000"/></w:pgBorders>'))


# ───────────────────────────── the memorial document ─────────────────────────────

def build_moot_docx(m: dict) -> str:
    """m = {cover, side_title, roles, jurisdiction[], facts[], issues[{text, submission, subs[{title, paras[]}]}],
    summary[], prayer[], abbreviations[]}. Returns path of the saved .docx."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)   # A4
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1.1)
    width = _text_width_in(doc)
    _style_base(doc)
    fb = _Footnotes()
    side_title = m['side_title']
    team = m['cover'].get('team_code')

    # ── cover (own section: page border, no header/footer)
    _cover_page(doc, m['cover'], side_title, m['roles'])
    _page_border(sec)
    _blank_header_footer(sec)

    # ── front matter section (roman numerals)
    front = doc.add_section(WD_SECTION.NEW_PAGE)
    for k in ('pgBorders',):
        for old in front._sectPr.findall(qn('w:' + k)):
            front._sectPr.remove(old)
    _page_numbering(front, 'lowerRoman')
    _make_header_footer(front, side_title, team, width)

    # TABLE OF CONTENTS (real TOC field, cached with the headings so it is never empty)
    toc_entries = []   # (level, text) filled after the body is known; we write placeholders now
    _heading(doc, 'TABLE OF CONTENTS', 1)
    toc_anchor = doc.add_paragraph()   # replaced with the field once we know every heading
    # INDEX OF AUTHORITIES placeholder anchor (filled after footnotes are known)
    h_idx = _heading(doc, 'INDEX OF AUTHORITIES', 1, page_break=True)
    idx_anchor = doc.add_paragraph()
    h_abbr = _heading(doc, 'LIST OF ABBREVIATIONS', 1, page_break=True)
    abbr_anchor = doc.add_paragraph()

    _heading(doc, 'STATEMENT OF JURISDICTION', 1, page_break=True)
    for i, para in enumerate(m['jurisdiction']):
        _body_para(doc, para, fb, block_key=f'cite_j_{i}', indent_first=True)

    _heading(doc, 'STATEMENT OF FACTS', 1, page_break=True)
    for i, para in enumerate(m['facts'], 1):
        _body_para(doc, para, fb, block_key=f'cite_f_{i}', num=i)

    _heading(doc, 'ISSUES RAISED', 1, page_break=True)
    for i, it in enumerate(m['issues'], 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(f'ISSUE {_roman(i)}:')
        _set_run_font(r, bold=True)
        _body_para(doc, it['text'], fb, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    _heading(doc, 'SUMMARY OF ARGUMENTS', 1, page_break=True)
    for i, it in enumerate(m['issues']):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(it['text'])
        _set_run_font(r, bold=True)
        _body_para(doc, m['summary'][i], fb, block_key=f'cite_s_{i}', indent_first=True)

    # ── body section (arabic numerals restart at 1)
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    _page_numbering(body, 'decimal')
    _make_header_footer(body, side_title, team, width)

    _heading(doc, 'ARGUMENTS ADVANCED', 1)
    for i, it in enumerate(m['issues'], 1):
        title = f"{_roman(i)}. " + (it['submission'] or it['text']).upper().rstrip('.')
        _heading(doc, title, 2, page_break=(i > 1))
        subs = it['subs']
        if len(subs) > 1:
            intro = ("It is most humbly submitted before this Hon'ble Court that the above submission is made out "
                     "on the following grounds:")
            _body_para(doc, intro, fb, indent_first=True)
            for j, s in enumerate(subs):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.9)
                p.paragraph_format.first_line_indent = Inches(-0.4)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.tab_stops.add_tab_stop(Inches(0.9))
                t = s['title'].strip().rstrip('.')
                t = (t[0].upper() + t[1:].lower()) if t.isupper() else (t[0].upper() + t[1:])
                r = p.add_run(f"{chr(97 + j)})\t{t}.")
                _set_run_font(r)
        n = 0
        for j, s in enumerate(subs):
            _heading(doc, f"{chr(65 + j)}. {s['title'].upper()}", 3)
            for k, para in enumerate(s['paras']):
                n += 1
                bk = f"cite_{i}_{j}_{k}"
                _body_para(doc, para, fb, block_key=bk, num=n)
    _heading(doc, 'PRAYER FOR RELIEF', 1, page_break=True)
    _body_para(doc, "Wherefore, in light of the issues raised, arguments advanced and authorities cited, it is most "
                    f"humbly prayed that this Hon'ble Court may be pleased to adjudge and declare that:", fb, indent_first=True)
    for i, item in enumerate(m['prayer'], 1):
        _body_para(doc, item, fb, num=i)
    _body_para(doc, "And pass any other order, direction or relief that this Hon'ble Court may deem fit in the "
                    "interest of justice, equity and good conscience.", fb, indent_first=True)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run("All of which is most humbly prayed.")
    _set_run_font(r)
    sig = doc.add_paragraph()
    sig.paragraph_format.tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT)
    sig.paragraph_format.space_before = Pt(20)
    r = sig.add_run("PLACE: ____________\tSD/- ____________")
    _set_run_font(r)
    sig2 = doc.add_paragraph()
    sig2.paragraph_format.tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT)
    r = sig2.add_run(f"DATE: ___/___/____\tCOUNSEL FOR THE {side_title.replace('THE ', '', 1)}")
    _set_run_font(r)

    # ── index of authorities, built from the real footnotes
    index, verify_list = build_authority_index(fb)
    _fill_index(doc, idx_anchor, index, width)
    _fill_abbreviations(doc, abbr_anchor, m, fb)
    _fill_toc(doc, toc_anchor, m, width, index)

    _attach_footnotes(doc, fb)
    _set_update_fields(doc)
    add_watermark(doc, DRATIDO_WATERMARK_TEXT, preserve=True)

    os.makedirs('generated', exist_ok=True)
    out = os.path.abspath(f'generated/Moot_Memorial_{uuid.uuid4().hex[:8]}.docx')
    doc.save(out)
    return out, fb, verify_list


def _roman(n):
    vals = [(10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    out = ''
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _move_after(anchor, paragraphs):
    """Place already-created paragraphs after anchor, preserving order; drop anchor."""
    prev = anchor._p
    for p in paragraphs:
        prev.addnext(p._p)
        prev = p._p
    anchor._p.getparent().remove(anchor._p)


def _fill_index(doc, anchor, index, width):
    made = []

    def sub(title):
        h = doc.add_heading('', level=2)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = h.add_run(title)
        _set_run_font(r, bold=True, color='000000')
        made.append(h)

    def entry(n, disp, blocks):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Inches(0.4)
        pf.first_line_indent = Inches(-0.4)
        pf.right_indent = Inches(0.5)
        pf.space_after = Pt(4)
        pf.tab_stops.add_tab_stop(Inches(0.4))
        pf.tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        lead = f'{n}.\t' if n else '•\t'
        r = p.add_run(f'{lead}{disp}\t')
        _set_run_font(r, size=11)
        # collapse repeats within one sub-submission to one page reference
        seen, refs = set(), []
        for b in blocks:
            grp = b.rsplit('_', 1)[0] if b.startswith('cite_') else b
            if grp in seen:
                continue
            seen.add(grp)
            refs.append(b)
        if len(refs) > 6:
            rr = p.add_run('passim')
            _set_run_font(rr, size=11, italic=True)
            refs = []
        for k, b in enumerate(refs):
            if k:
                rr = p.add_run(', ')
                _set_run_font(rr, size=11)
            if b.startswith('cite_'):
                _fld(p, f'PAGEREF {b} \\h', '–', size=11)
        made.append(p)

    sections = (('case', 'CASES REFERRED', True), ('statute', 'STATUTES REFERRED', False),
                ('book', 'BOOKS & COMMENTARIES REFERRED', False), ('web', 'WEBSITES REFERRED', False))
    any_found = False
    for key, title, numbered in sections:
        items = index.get(key) or []
        if not items:
            continue
        any_found = True
        sub(title)
        for n, (disp, blocks) in enumerate(items, 1):
            entry(n if numbered else None, disp, blocks)
    if not any_found:
        p = doc.add_paragraph()
        r = p.add_run('(No authorities cited.)')
        _set_run_font(r)
        made.append(p)
    _move_after(anchor, made)


def _fill_abbreviations(doc, anchor, m, fb):
    width = _text_width_in(doc)
    corpus = ' '.join([' '.join(m['jurisdiction']), ' '.join(m['facts']), ' '.join(m['summary']),
                       ' '.join(m['prayer'])]
                      + [p for it in m['issues'] for s in it['subs'] for p in s['paras']]
                      + [t for _i, t, _u, _k in fb.items])
    corpus = _FN_RE.sub(' ', corpus) + ' ' + ' '.join(t for _i, t, _u, _k in fb.items)
    rows = [(a, f) for a, f, rx in COMMON_ABBREVIATIONS if re.search(rx, corpus)]
    known = {a.lower() for a, _ in rows}
    for pair in m.get('abbreviations') or []:
        a, f = str(pair[0]).strip(), str(pair[1]).strip()
        if a and f and a.lower() not in known and re.search(r'\b' + re.escape(a) + r'\b', corpus):
            rows.append((a, f))
            known.add(a.lower())
    rows.sort(key=lambda x: x[0].lower())
    tbl = doc.add_table(rows=1, cols=2, style='Table Grid')
    hdr = tbl.rows[0].cells
    for c, txt in zip(hdr, ('ABBREVIATION', 'FULL FORM')):
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.paragraphs[0].add_run(txt)
        _set_run_font(r, size=11, bold=True)
    for a, f in rows:
        cells = tbl.add_row().cells
        for c, txt in zip(cells, (a, f)):
            r = c.paragraphs[0].add_run(txt)
            _set_run_font(r, size=11)
    for row in tbl.rows:
        row.cells[0].width = Inches(1.6)
        row.cells[1].width = Inches(width - 1.6)
    anchor._p.addnext(tbl._tbl)
    anchor._p.getparent().remove(anchor._p)


def _fill_toc(doc, anchor, m, width, index):
    """TOC field. Cached result lists every heading (no page numbers) so it reads correctly even
    before Word refreshes the field; Word fills page numbers on open (updateFields)."""
    idx_names = (('case', 'Cases Referred'), ('statute', 'Statutes Referred'),
                 ('book', 'Books & Commentaries Referred'), ('web', 'Websites Referred'))
    heads = [(1, 'Table of Contents'), (1, 'Index of Authorities')]
    heads += [(2, nm) for k, nm in idx_names if index.get(k)]
    heads += [(1, 'List of Abbreviations'), (1, 'Statement of Jurisdiction'), (1, 'Statement of Facts'), (1, 'Issues Raised'),
             (1, 'Summary of Arguments'), (1, 'Arguments Advanced')]
    for i, it in enumerate(m['issues'], 1):
        heads.append((2, f"{_roman(i)}. " + (it['submission'] or it['text']).rstrip('.')))
        for j, s in enumerate(it['subs']):
            heads.append((3, f"{chr(65 + j)}. {s['title']}"))
    heads.append((1, 'Prayer for Relief'))

    paras = []
    for n, (lvl, txt) in enumerate(heads):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Inches(0.3 * (lvl - 1))
        pf.space_after = Pt(3)
        pf.right_indent = Inches(0.4)
        pf.tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if n == 0:
            for xml in (f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="begin"/>',):
                r = p.add_run(); r._r.append(parse_xml(xml))
            r = p.add_run()
            r._r.append(parse_xml(f'<w:instrText xmlns:w="{W_NS}" xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'))
            r = p.add_run(); r._r.append(parse_xml(f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="separate"/>'))
        r = p.add_run(f'{txt}\t–')
        _set_run_font(r, size=11, bold=(lvl == 1))
        if n == len(heads) - 1:
            r = p.add_run(); r._r.append(parse_xml(f'<w:fldChar xmlns:w="{W_NS}" w:fldCharType="end"/>'))
        paras.append(p)
    _move_after(anchor, paras)


# ───────────────────────────── orchestration ─────────────────────────────

def moot_preview_text(m, fb) -> str:
    """Plain-text rendering for the side panel: footnote markers become [n] and footnotes are listed
    under each block. The cover/TOC/index are assembled in the .docx."""
    counter = [0]
    out = []

    def render(text):
        notes = []

        def sub(match):
            counter[0] += 1
            val = match.group(1).replace('(?)', '', 1).strip()
            notes.append(f'    [{counter[0]}] {val}')
            return f'[{counter[0]}]'
        body = _FN_RE.sub(sub, text)
        return body, notes

    def block(title, paras, numbered=False):
        out.append(title)
        out.append('')
        for i, p in enumerate(paras, 1):
            body, notes = render(p)
            out.append(f'{i}. {body}' if numbered else body)
            out.extend(notes)
            out.append('')

    r1, r2 = m['roles']
    c = m['cover']
    out += [(c.get('competition') or '').upper(), '', f"BEFORE {(c.get('court') or '').upper()}", '',
            f"{(c.get('party1') or '').upper()} ........ {r1}", '   v', f"{(c.get('party2') or '').upper()} ........ {r2}", '',
            f"MEMORIAL ON BEHALF OF {m['side_title']}", '',
            '[Cover page, Table of Contents, Index of Authorities and List of Abbreviations are built in the .docx]', '']
    block('STATEMENT OF JURISDICTION', m['jurisdiction'])
    block('STATEMENT OF FACTS', m['facts'], numbered=True)
    out += ['ISSUES RAISED', '']
    out += [f"ISSUE {_roman(i)}: {it['text']}" for i, it in enumerate(m['issues'], 1)] + ['']
    block('SUMMARY OF ARGUMENTS', m['summary'])
    out += ['ARGUMENTS ADVANCED', '']
    for i, it in enumerate(m['issues'], 1):
        out += [f"{_roman(i)}. {(it['submission'] or it['text']).upper()}", '']
        n = 0
        for j, s in enumerate(it['subs']):
            out += [f"{chr(65 + j)}. {s['title'].upper()}", '']
            for para in s['paras']:
                n += 1
                body, notes = render(para)
                out.append(f'{n}. {body}')
                out.extend(notes)
                out.append('')
    block('PRAYER FOR RELIEF', m['prayer'], numbered=True)
    return '\n'.join(out)


def generate_moot_memorial(conv, progress):
    """Full pipeline. `progress(done, total, label)` is called between steps."""
    profile = conv.get("moot_profile") or moot_profile_from_text(conv["details"])
    conv["moot_profile"] = profile
    side_key = moot_side_key(conv)
    r1, r2 = moot_roles(profile)

    progress(0, 6, "Planning issues and sub-submissions…")
    plan = moot_make_plan(conv, profile)
    total = 4 + sum(len(it['subs']) for it in plan['issues'])
    done = 1
    warnings = []

    def step(label, fn, fallback):
        nonlocal done
        progress(done, total, label)
        try:
            res = fn()
        except Exception as e:
            print(f"[Moot] step failed ({label}): {e}")
            warnings.append(label)
            res = fallback
        done += 1
        return res

    jurisdiction = step("Writing Statement of Jurisdiction…",
                        lambda: moot_write_jurisdiction(conv, profile, plan),
                        ["[Statement of Jurisdiction could not be generated — please write it manually.]"])
    facts = step("Writing Statement of Facts…", lambda: moot_write_facts(conv, profile),
                 ["[Statement of Facts could not be generated — please write it manually.]"])

    cited = []
    for i, it in enumerate(plan['issues'], 1):
        for j, s in enumerate(it['subs']):
            label = f"Arguing Issue {_roman(i)} — {s['title'].title()[:60]}…"
            paras = step(label, lambda it=it, s=s: moot_write_sub(conv, profile, plan, it, s, cited),
                         ["[This sub-submission could not be generated — please regenerate the draft.]"])
            s['paras'] = paras
            for p in paras:
                for kind, val in _parse_marked(p):
                    if kind == 'fn' and _CASE_RE.search(val):
                        cited.append(re.split(r'[,(\[]', val)[0].strip())

    summary = step("Writing Summary of Arguments…", lambda: moot_write_summary(conv, profile, plan),
                   ["[Summary could not be generated.]"] * len(plan['issues']))
    prayer = step("Writing Prayer for Relief…", lambda: moot_write_prayer(conv, profile, plan),
                  [f"The issues be decided in favour of the {r1 if side_key == 'petitioner' else r2}."])

    progress(total, total, "Assembling the Word document…")
    cover = {k: (plan['cover'].get(k) or profile.get(k) or '') for k in
             ('competition', 'team_code', 'case_number', 'filed_under', 'subject')}
    cover.update({"court": profile.get("court"), "party1": profile.get("party1"), "party2": profile.get("party2")})
    my_role = r1 if side_key == 'petitioner' else r2
    side_title = my_role if my_role.upper().startswith('THE ') else f"THE {my_role}"
    memorial = {"cover": cover, "side_title": side_title, "roles": (r1, r2), "jurisdiction": jurisdiction,
                "facts": facts, "issues": plan['issues'], "summary": summary, "prayer": prayer,
                "abbreviations": plan.get('abbreviations')}
    path, fb, verify_list = build_moot_docx(memorial)
    preview = moot_preview_text(memorial, fb)
    return {"path": path, "preview": preview, "verify": verify_list, "warnings": warnings, "profile": profile}


# ═══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/api/start', methods=['POST'])
def api_start():
    conv = new_conversation()
    push(conv, "assistant", WELCOME_MSG, buttons=START_BUTTONS)
    return jsonify({"success": True, "conv_id": conv["id"], "messages": conv["messages"]})


@app.route('/api/message', methods=['POST'])
def api_message():
    data = request.get_json(silent=True) or {}
    conv_id = data.get('conv_id', '')
    text = (data.get('text') or '').strip()
    conv = get_conversation(conv_id)
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found. Start a new draft."}), 404
    if not text:
        return jsonify({"success": False, "message": "Please enter a message."}), 400

    push(conv, "user", text)

    handler = STAGE_HANDLERS.get(conv["stage"])
    if not handler:
        return jsonify({"success": False, "message": "Unknown stage."}), 400
    handler(conv, text)

    return jsonify({
        "success": True,
        "messages": conv["messages"],
        "stage": conv["stage"],
        "can_generate": conv["stage"] == "brainstorm",
    })


@app.route('/api/upload_template', methods=['POST'])
def api_upload_template():
    conv_id = request.form.get('conv_id', '')
    conv = get_conversation(conv_id)
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found. Start a new draft."}), 404
    if conv["stage"] != "ask_template":
        return jsonify({"success": False, "message": "Not expecting a template upload right now."}), 400

    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({"success": False, "message": "No file received."}), 400

    filename = f.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_TEMPLATE_EXTENSIONS:
        return jsonify({"success": False, "message": "Please upload a .docx or .pdf file."}), 400

    f.stream.seek(0, os.SEEK_END)
    size = f.stream.tell()
    f.stream.seek(0)
    if size == 0:
        return jsonify({"success": False, "message": "That file appears to be empty."}), 400
    if size > MAX_TEMPLATE_UPLOAD_BYTES:
        return jsonify({"success": False, "message": "That file is too large (max 15 MB)."}), 400

    try:
        if ext == '.docx':
            extracted = extract_text_from_docx(f.stream)
        else:
            extracted = extract_text_from_pdf(f.stream)
    except Exception as e:
        print(f"[Upload] Extraction failed for {filename}: {e}")
        return jsonify({"success": False,
                        "message": "Couldn't read that file — it may be corrupted, password-protected, "
                                   "or an unsupported format. Try another file or paste the template "
                                   "text instead."}), 400

    if not extracted.strip():
        return jsonify({"success": False,
                        "message": "No readable text was found in that file (it may be a scanned "
                                   "image rather than real text). Try another file or paste the "
                                   "template text instead."}), 400

    push(conv, "user", f"📎 Uploaded template: {filename}")

    conv["template_text"] = extracted
    conv["template_source"] = f"uploaded file: {filename}"
    conv["mode"] = "template"
    ask_for_template_details(
        conv,
        f'Got it — I\'ve read "{filename}". Click below to enter the data to fill into it '
        f'(names, dates, amounts, and any other specifics).'
    )

    return jsonify({
        "success": True,
        "messages": conv["messages"],
        "stage": conv["stage"],
        "can_generate": conv["stage"] == "brainstorm",
    })


def _run_generation_job(conv):
    """Runs in a background thread so no HTTP request outlives the web server's worker timeout."""
    job = conv["job"]

    def progress(done, total, label):
        job.update(done=done, total=total, label=label)

    try:
        if conv["doc_type"] == MOOT_MEMORIAL_DOC_TYPE:
            res = generate_moot_memorial(conv, progress)
            conv["draft_text"] = res["preview"]
            conv["docx_path"] = res["path"]
            verify = res["verify"]
            note = ("Your memorial is ready — review it in the side panel and download the Word file. "
                    "The .docx has a cover page, real footnotes, a table of contents and an index of authorities. "
                    "When Word asks to update fields on opening, click Yes so page numbers fill in.")
            if verify:
                lines = [f"  • {'(?) ' if unsure else ''}{name}" for name, unsure in verify[:40]]
                note += ("\n\nIMPORTANT — verify every case before you file. AI can mis-remember citations and "
                         "holdings. Check these on SCC Online / Manupatra" +
                         (" (the ones marked (?) are the least certain):\n" if any(u for _n, u in verify) else ":\n")
                         + "\n".join(lines))
            if res["warnings"]:
                note += ("\n\nThese parts could not be generated (usually an API rate limit) and are marked in the "
                         "document: " + "; ".join(res["warnings"]) + ". Click Generate Draft again to retry.")
            push(conv, "assistant", note)
        else:
            conv["draft_text"] = generate_draft(conv)
            push(conv, "assistant",
                 "Here's a draft based on everything we've discussed. Review it in the side panel, and keep "
                 "chatting if you'd like changes — you can regenerate any time.")
        job["state"] = "done"
    except Exception as e:
        print(f"[Generate] failed: {e}")
        job.update(state="error", error=str(e))


@app.route('/api/generate', methods=['POST'])
def api_generate():
    data = request.get_json(silent=True) or {}
    conv_id = data.get('conv_id', '')
    conv = get_conversation(conv_id)
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found. Start a new draft."}), 404
    if conv["stage"] != "brainstorm":
        return jsonify({"success": False, "message": "Finish the setup questions before generating a draft."}), 400
    if not os.environ.get('GROQ_API_KEY', '').strip():
        return jsonify({"success": False,
                        "message": "GROQ_API_KEY not set. Get a free key at https://console.groq.com"}), 400
    if (conv.get("job") or {}).get("state") == "running":
        return jsonify({"success": True, "started": False, "message": "Already generating."})

    conv["job"] = {"state": "running", "done": 0, "total": 1, "label": "Starting…", "error": None}
    threading.Thread(target=_run_generation_job, args=(conv,), daemon=True).start()
    return jsonify({"success": True, "started": True})


@app.route('/api/generate_status/<conv_id>')
def api_generate_status(conv_id):
    conv = get_conversation(conv_id)
    if not conv or not conv.get("job"):
        return jsonify({"success": False, "message": "No generation in progress."}), 404
    job = conv["job"]
    out = {"success": True, "state": job["state"], "done": job.get("done", 0),
           "total": job.get("total", 1), "label": job.get("label", "")}
    if job["state"] == "done":
        out.update(draft_text=conv["draft_text"], messages=conv["messages"])
    elif job["state"] == "error":
        out.update(success=False, state="error", message=job.get("error") or "Generation failed.")
    return jsonify(out)


@app.route('/api/download/<conv_id>')
def api_download(conv_id):
    conv = get_conversation(conv_id)
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found."}), 404
    fp = conv.get("docx_path")
    if not fp or not os.path.exists(fp):
        return jsonify({"success": False, "message": "No draft generated yet."}), 404

    slug = re.sub(r'[^\w\-]', '_', (conv.get("doc_type") or "draft")[:40]) or "draft"
    return send_file(fp, as_attachment=True,
                     download_name=f'dratido_{slug}.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/')
def index():
    return Response(HTML, mimetype='text/html')


# ═══════════════════════════════════════════════════════════════════════════════
#  FRONTEND (single-page chat app)
# ═══════════════════════════════════════════════════════════════════════════════

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Dratido — Draft Till Done</title>
<style>
  :root{
    --maroon:#8B1E2D; --maroon-dark:#6e1723; --ink:#1c1a19; --paper:#faf7f2;
    --panel:#ffffff; --line:#e7e0d6; --muted:#7a7268; --bubble-user:#8B1E2D;
    --bubble-ai:#ffffff;
  }
  *{box-sizing:border-box;}
  html,body{height:100%;}
  body{
    margin:0; font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
    background:var(--paper); color:var(--ink); display:flex; flex-direction:column;
  }
  header{
    display:flex; align-items:center; justify-content:space-between;
    padding:14px 20px; background:var(--panel); border-bottom:1px solid var(--line);
    flex-shrink:0; z-index:5;
  }
  .brand{display:flex; align-items:baseline; gap:10px;}
  .brand .name{font-size:22px; font-weight:700; color:var(--maroon); letter-spacing:.3px;}
  .brand .tagline{font-size:12px; color:var(--muted); font-style:italic;}
  .header-actions{display:flex; gap:10px;}
  .btn{
    border:1px solid var(--line); background:var(--panel); color:var(--ink);
    padding:8px 14px; border-radius:8px; font-size:13px; cursor:pointer;
    transition:.15s; white-space:nowrap;
  }
  .btn:hover{border-color:var(--maroon); color:var(--maroon);}
  .btn.primary{background:var(--maroon); border-color:var(--maroon); color:#fff;}
  .btn.primary:hover{background:var(--maroon-dark);}
  .btn:disabled{opacity:.45; cursor:not-allowed;}

  main{flex:1; display:flex; min-height:0; position:relative;}

  #chat-col{flex:1; display:flex; flex-direction:column; min-width:0;}
  #messages{flex:1; overflow-y:auto; padding:24px 16px 8px; display:flex; flex-direction:column; gap:14px;}
  .row{display:flex; width:100%;}
  .row.user{justify-content:flex-end;}
  .row.assistant{justify-content:flex-start;}
  .bubble{
    max-width:min(640px,86%); padding:12px 16px; border-radius:14px; line-height:1.5;
    font-size:14.5px; white-space:pre-wrap; word-wrap:break-word; box-shadow:0 1px 2px rgba(0,0,0,.05);
  }
  .row.user .bubble{background:var(--bubble-user); color:#fff; border-bottom-right-radius:4px;}
  .row.assistant .bubble{background:var(--bubble-ai); border:1px solid var(--line); border-bottom-left-radius:4px;}
  .quick-replies{display:flex; flex-wrap:wrap; gap:8px; margin-top:8px; max-width:640px;}
  .qr-btn{
    border:1px solid var(--maroon); color:var(--maroon); background:#fff;
    padding:7px 13px; border-radius:20px; font-size:13px; cursor:pointer; transition:.15s;
  }
  .qr-btn:hover{background:var(--maroon); color:#fff;}
  .qr-btn:disabled{opacity:.4; cursor:not-allowed;}
  .modal-trigger-btn{
    margin-top:8px; border:1px solid var(--maroon); background:var(--maroon); color:#fff;
    padding:8px 14px; border-radius:8px; font-size:13px; cursor:pointer; transition:.15s;
    align-self:flex-start;
  }
  .modal-trigger-btn:hover{background:var(--maroon-dark);}
  .typing{font-size:13px; color:var(--muted); padding:0 16px 8px; font-style:italic;}

  .modal-overlay{
    position:fixed; inset:0; background:rgba(28,26,25,.45); display:none;
    align-items:center; justify-content:center; z-index:50; padding:20px;
  }
  .modal-overlay.open{display:flex;}
  .modal-box{
    background:var(--panel); border-radius:12px; width:100%; max-width:560px;
    padding:22px 22px 18px; box-shadow:0 12px 40px rgba(0,0,0,.25);
  }
  .modal-box h3{margin:0 0 4px; color:var(--maroon); font-size:16px;}
  .modal-box .modal-hint{margin:0 0 14px; font-size:12.5px; color:var(--muted);}
  .modal-box textarea{
    width:100%; min-height:200px; resize:vertical; border:1px solid var(--line);
    border-radius:8px; padding:12px 14px; font-size:14px; font-family:inherit;
    outline:none; box-sizing:border-box;
  }
  .modal-box textarea:focus{border-color:var(--maroon);}
  .modal-actions{display:flex; justify-content:flex-end; gap:10px; margin-top:14px;}

  #modal-search{
    width:100%; border:1px solid var(--line); border-radius:8px; padding:10px 14px;
    font-size:14px; font-family:inherit; outline:none; box-sizing:border-box; margin-bottom:10px;
  }
  #modal-search:focus{border-color:var(--maroon);}
  #modal-list-results{
    max-height:320px; overflow-y:auto; border:1px solid var(--line); border-radius:8px;
    padding:6px; background:var(--paper);
  }
  .modal-group-heading{
    font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:var(--muted);
    padding:8px 8px 4px; font-weight:600;
  }
  .modal-list-item{
    display:block; width:100%; text-align:left; background:none; border:none;
    padding:9px 10px; border-radius:6px; font-size:14px; color:var(--ink);
    cursor:pointer; transition:.12s;
  }
  .modal-list-item:hover{background:#f1e9dd; color:var(--maroon);}
  .modal-list-empty{padding:16px 10px; color:var(--muted); font-size:13.5px; text-align:center;}
  .modal-template-switch-btn{
    display:block; width:100%; margin-top:10px; border:1px dashed var(--maroon);
    background:#fff; color:var(--maroon); border-radius:8px; padding:10px 14px;
    font-size:13.5px; font-weight:600; cursor:pointer; transition:.15s;
  }
  .modal-template-switch-btn:hover{background:var(--maroon); color:#fff; border-style:solid;}

  .modal-upload-zone{
    border:2px dashed var(--line); border-radius:10px; padding:30px 16px; text-align:center;
    cursor:pointer; transition:.15s; background:var(--paper);
  }
  .modal-upload-zone.drag{border-color:var(--maroon); background:#f6ece7;}
  .modal-upload-zone .icon{font-size:30px; margin-bottom:8px;}
  .modal-upload-zone .main-text{font-size:14px; color:var(--ink); font-weight:600;}
  .modal-upload-zone .sub-text{font-size:12px; color:var(--muted); margin-top:5px; line-height:1.5;}
  .modal-paste-input{
    width:100%; margin-top:12px; border:1px dashed var(--line); border-radius:8px;
    padding:9px 12px; font-size:12.5px; font-family:inherit; color:var(--muted);
    resize:none; height:38px; outline:none; box-sizing:border-box;
  }
  .modal-paste-input:focus{border-color:var(--maroon); color:var(--ink);}
  #modal-upload-status{
    font-size:12.5px; color:var(--maroon); margin-top:10px; min-height:16px; text-align:center;
  }

  #composer{
    display:flex; gap:10px; padding:14px 16px; border-top:1px solid var(--line);
    background:var(--panel); flex-shrink:0; align-items:flex-end;
  }
  #composer textarea{
    flex:1; resize:none; border:1px solid var(--line); border-radius:10px;
    padding:11px 14px; font-size:14.5px; font-family:inherit; max-height:140px; min-height:44px;
    outline:none;
  }
  #composer textarea:focus{border-color:var(--maroon);}
  #send-btn{
    background:var(--maroon); color:#fff; border:none; border-radius:10px;
    width:44px; height:44px; font-size:18px; cursor:pointer; flex-shrink:0;
  }
  #send-btn:hover{background:var(--maroon-dark);}
  #send-btn:disabled{opacity:.4; cursor:not-allowed;}
  #generate-btn{flex-shrink:0;}

  #panel{
    width:0; overflow:hidden; border-left:1px solid var(--line); background:var(--panel);
    transition:width .22s ease; flex-shrink:0; display:flex; flex-direction:column;
  }
  #panel.open{width:420px;}
  #panel-inner{width:420px; display:flex; flex-direction:column; height:100%;}
  #panel-header{
    padding:16px 20px; border-bottom:1px solid var(--line); display:flex;
    align-items:center; justify-content:space-between; flex-shrink:0;
  }
  #panel-header h3{margin:0; font-size:15px; color:var(--maroon);}
  #panel-body{flex:1; overflow-y:auto; padding:20px;}
  #panel-body .placeholder{color:var(--muted); font-size:13.5px; line-height:1.6;}
  #panel-body .setup-item{margin-bottom:14px; font-size:13px;}
  #panel-body .setup-item .k{color:var(--muted); text-transform:uppercase; font-size:11px; letter-spacing:.5px; margin-bottom:3px;}
  #panel-body .setup-item .v{color:var(--ink);}
  #draft-text{
    font-family:'Georgia','Times New Roman',serif; font-size:13.5px; line-height:1.7;
    white-space:pre-wrap; word-wrap:break-word; color:var(--ink);
  }
  #panel-footer{padding:14px 20px; border-top:1px solid var(--line); flex-shrink:0;}
  #panel-footer .btn{width:100%;}

  @media (max-width:820px){
    #panel.open{position:fixed; top:0; right:0; bottom:0; width:100%; z-index:20;}
    #panel-inner{width:100%;}
    .brand .tagline{display:none;}
  }
</style>
</head>
<body>

<header>
  <div class="brand">
    <span class="name">Dratido</span>
    <span class="tagline">Draft Till Done</span>
  </div>
  <div class="header-actions">
    <button class="btn" id="panel-toggle">Draft ▤</button>
    <button class="btn" id="new-draft-btn">＋ New Draft</button>
  </div>
</header>

<main>
  <div id="chat-col">
    <div id="messages"></div>
    <div class="typing" id="typing" style="display:none;">Dratido is thinking…</div>
    <div id="composer">
      <textarea id="input" placeholder="Type your message…" rows="1"></textarea>
      <button class="btn primary" id="generate-btn" style="display:none;">Generate Draft</button>
      <button id="send-btn" title="Send">➤</button>
    </div>
  </div>

  <div id="panel">
    <div id="panel-inner">
      <div id="panel-header">
        <h3>Draft</h3>
        <button class="btn" id="panel-close">✕</button>
      </div>
      <div id="panel-body">
        <div class="placeholder">Your draft will appear here once we've brainstormed enough to generate it.</div>
      </div>
      <div id="panel-footer" style="display:none;">
        <button class="btn primary" id="download-btn">⬇ Download as Word (.docx)</button>
      </div>
    </div>
  </div>
</main>

<div class="modal-overlay" id="modal-overlay">
  <div class="modal-box">
    <h3 id="modal-title">Enter Details</h3>
    <p class="modal-hint" id="modal-hint">This opens in its own window so you can enter everything comfortably before it's added to the chat.</p>

    <div id="modal-text-mode">
      <textarea id="modal-textarea" placeholder=""></textarea>
      <div class="modal-actions">
        <button class="btn" id="modal-cancel">Cancel</button>
        <button class="btn primary" id="modal-submit">Submit</button>
      </div>
    </div>

    <div id="modal-list-mode" style="display:none;">
      <input type="text" id="modal-search" placeholder="Search document types..." autocomplete="off">
      <div id="modal-list-results"></div>
      <button class="modal-template-switch-btn" id="modal-template-switch">⇄ Use a Reference Template Instead</button>
      <div class="modal-actions">
        <button class="btn" id="modal-list-cancel">Cancel</button>
      </div>
    </div>

    <div id="modal-upload-mode" style="display:none;">
      <div class="modal-upload-zone" id="modal-upload-zone">
        <div class="icon">📎</div>
        <div class="main-text">Drop your template here, or tap to browse</div>
        <div class="sub-text">.docx or .pdf files only</div>
      </div>
      <input type="file" id="modal-file-input" accept=".docx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document" style="display:none;">
      <textarea id="modal-paste-catcher" class="modal-paste-input" rows="1" placeholder="On iPhone/iPad: tap here, then Paste a copied document"></textarea>
      <div id="modal-upload-status"></div>
      <button class="modal-template-switch-btn" id="modal-upload-toggle-text">✎ Or Paste the Template Text Instead</button>
      <div class="modal-actions">
        <button class="btn" id="modal-upload-cancel">Cancel</button>
      </div>
    </div>
  </div>
</div>

<script>
let convId = sessionStorage.getItem('dratido_conv_id') || null;
let canGenerate = false;
let hasDraft = false;

const messagesEl = document.getElementById('messages');
const inputEl = document.getElementById('input');
const sendBtn = document.getElementById('send-btn');
const generateBtn = document.getElementById('generate-btn');
const typingEl = document.getElementById('typing');
const panelEl = document.getElementById('panel');
const panelBody = document.getElementById('panel-body');
const panelFooter = document.getElementById('panel-footer');

function scrollBottom(){ messagesEl.scrollTop = messagesEl.scrollHeight; }

// Renders plain text as real bold (no literal ** markup) and preserves line breaks.
function richText(text){
  const esc = (text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  return esc
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*(?!\*)([^*\n]+?)\*(?!\*)/g, '$1<strong>$2</strong>')
    .replace(/\n/g, '<br>');
}

function renderMessages(msgs){
  messagesEl.innerHTML = '';
  let modalToAutoOpen = null;

  msgs.forEach((m, idx) => {
    const row = document.createElement('div');
    row.className = 'row ' + (m.role === 'user' ? 'user' : 'assistant');
    const bubbleWrap = document.createElement('div');
    bubbleWrap.style.display = 'flex';
    bubbleWrap.style.flexDirection = 'column';
    bubbleWrap.style.alignItems = m.role === 'user' ? 'flex-end' : 'flex-start';

    const bubble = document.createElement('div');
    bubble.className = 'bubble';
    bubble.innerHTML = richText(m.content);
    bubbleWrap.appendChild(bubble);

    if (m.modal){
      const openBtn = document.createElement('button');
      openBtn.className = 'modal-trigger-btn';
      openBtn.textContent = '✎ ' + (m.modal.submit_label || 'Enter Details');
      openBtn.onclick = () => openModal(m.modal);
      bubbleWrap.appendChild(openBtn);
      if (idx === msgs.length - 1) modalToAutoOpen = m.modal;
    }

    if (m.buttons && m.buttons.length){
      const qr = document.createElement('div');
      qr.className = 'quick-replies';
      m.buttons.forEach(b => {
        const btn = document.createElement('button');
        btn.className = 'qr-btn';
        btn.textContent = b.label;
        btn.onclick = () => sendMessage(b.value);
        qr.appendChild(btn);
      });
      bubbleWrap.appendChild(qr);
    }
    row.appendChild(bubbleWrap);
    messagesEl.appendChild(row);
  });
  scrollBottom();
  if (modalToAutoOpen) setTimeout(() => openModal(modalToAutoOpen), 300);
}

const textModeEl = document.getElementById('modal-text-mode');
const listModeEl = document.getElementById('modal-list-mode');
const uploadModeEl = document.getElementById('modal-upload-mode');
const modalHintEl = document.getElementById('modal-hint');
const modalSearchEl = document.getElementById('modal-search');
const modalListResultsEl = document.getElementById('modal-list-results');
const TEMPLATE_SWITCH_VALUE = "I'd like to provide a reference template and enter the data to fill into it.";
let currentModalGroups = [];
let currentUploadCfg = null;

function openModal(cfg){
  document.getElementById('modal-title').textContent = cfg.title || 'Enter Details';
  textModeEl.style.display = 'none';
  listModeEl.style.display = 'none';
  uploadModeEl.style.display = 'none';

  if (cfg.type === 'list'){
    modalHintEl.textContent = 'Search or scroll to find your document type.';
    listModeEl.style.display = 'block';
    currentModalGroups = cfg.groups || [];
    modalSearchEl.value = '';
    renderModalList('');
    document.getElementById('modal-overlay').classList.add('open');
    setTimeout(() => modalSearchEl.focus(), 50);
  } else if (cfg.type === 'upload'){
    modalHintEl.textContent = cfg.hint || 'Drop a file, tap to browse, or paste a copied document.';
    uploadModeEl.style.display = 'block';
    currentUploadCfg = cfg;
    resetUploadZone();
    document.getElementById('modal-overlay').classList.add('open');
  } else {
    modalHintEl.textContent = "This opens in its own window so you can enter everything comfortably before it's added to the chat.";
    textModeEl.style.display = 'block';
    const ta = document.getElementById('modal-textarea');
    ta.placeholder = cfg.placeholder || '';
    ta.value = '';
    document.getElementById('modal-submit').textContent = cfg.submit_label || 'Submit';
    document.getElementById('modal-overlay').classList.add('open');
    setTimeout(() => ta.focus(), 50);
  }
}
function closeModal(){
  document.getElementById('modal-overlay').classList.remove('open');
}

function switchToTextMode(prefill){
  uploadModeEl.style.display = 'none';
  textModeEl.style.display = 'block';
  const ta = document.getElementById('modal-textarea');
  ta.placeholder = (currentUploadCfg && currentUploadCfg.placeholder) || 'Paste your template text here...';
  ta.value = prefill || '';
  document.getElementById('modal-submit').textContent = (currentUploadCfg && currentUploadCfg.submit_label) || 'Save Template';
  setTimeout(() => ta.focus(), 50);
}

const uploadZoneEl = document.getElementById('modal-upload-zone');
const fileInputEl = document.getElementById('modal-file-input');
const pasteCatcherEl = document.getElementById('modal-paste-catcher');
const uploadStatusEl = document.getElementById('modal-upload-status');

function resetUploadZone(){
  uploadZoneEl.classList.remove('drag');
  uploadStatusEl.textContent = '';
  fileInputEl.value = '';
  pasteCatcherEl.value = '';
}

function isValidTemplateFile(file){
  const name = (file.name || '').toLowerCase();
  return name.endsWith('.docx') || name.endsWith('.pdf');
}

async function uploadTemplateFile(file){
  if (!isValidTemplateFile(file)){
    uploadStatusEl.textContent = 'Please choose a .docx or .pdf file.';
    return;
  }
  if (file.size > 15 * 1024 * 1024){
    uploadStatusEl.textContent = 'That file is too large (max 15 MB).';
    return;
  }
  uploadStatusEl.textContent = 'Reading ' + (file.name || 'your file') + '…';
  setBusy(true);
  try{
    const fd = new FormData();
    fd.append('conv_id', convId);
    fd.append('file', file, file.name || 'template');
    const res = await fetch('/api/upload_template', {method:'POST', body: fd});
    const data = await res.json();
    setBusy(false);
    if (data.success){
      closeModal();
      renderMessages(data.messages);
      canGenerate = !!data.can_generate;
      generateBtn.style.display = canGenerate ? 'inline-block' : 'none';
    } else {
      uploadStatusEl.textContent = data.message || 'Could not read that file.';
    }
  } catch (err){
    setBusy(false);
    uploadStatusEl.textContent = 'Upload failed. Please try again.';
  }
}

uploadZoneEl.addEventListener('click', () => fileInputEl.click());
uploadZoneEl.addEventListener('dragover', (e) => { e.preventDefault(); uploadZoneEl.classList.add('drag'); });
uploadZoneEl.addEventListener('dragleave', () => uploadZoneEl.classList.remove('drag'));
uploadZoneEl.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadZoneEl.classList.remove('drag');
  const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) uploadTemplateFile(file);
});
fileInputEl.addEventListener('change', () => {
  const file = fileInputEl.files && fileInputEl.files[0];
  if (file) uploadTemplateFile(file);
});
// Handles the "document paste" gesture on iOS/iPadOS (copy a file in Files/Share sheet,
// then long-press > Paste here) as well as desktop Ctrl+V of a copied file. Falls back to
// treating a plain-text paste as pasted template text.
pasteCatcherEl.addEventListener('paste', (e) => {
  const files = (e.clipboardData && e.clipboardData.files) ? Array.from(e.clipboardData.files) : [];
  if (files.length){
    e.preventDefault();
    uploadTemplateFile(files[0]);
    pasteCatcherEl.value = '';
    return;
  }
  setTimeout(() => {
    const pasted = pasteCatcherEl.value.trim();
    if (pasted) switchToTextMode(pasted);
    pasteCatcherEl.value = '';
  }, 0);
});
document.getElementById('modal-upload-toggle-text').onclick = () => switchToTextMode('');
document.getElementById('modal-upload-cancel').onclick = closeModal;

function renderModalList(filterRaw){
  const filter = (filterRaw || '').trim().toLowerCase();
  modalListResultsEl.innerHTML = '';
  let anyMatch = false;

  currentModalGroups.forEach(group => {
    const matches = (group.items || []).filter(item => item.toLowerCase().includes(filter));
    if (!matches.length) return;
    anyMatch = true;
    const heading = document.createElement('div');
    heading.className = 'modal-group-heading';
    heading.textContent = group.group;
    modalListResultsEl.appendChild(heading);
    matches.forEach(item => {
      const btn = document.createElement('button');
      btn.className = 'modal-list-item';
      btn.textContent = item;
      btn.onclick = () => {
        closeModal();
        sendMessage(item);
      };
      modalListResultsEl.appendChild(btn);
    });
  });

  if (!anyMatch){
    const empty = document.createElement('div');
    empty.className = 'modal-list-empty';
    empty.textContent = 'No matching document type found in the list.';
    modalListResultsEl.appendChild(empty);
  }
}

modalSearchEl.addEventListener('input', () => renderModalList(modalSearchEl.value));
document.getElementById('modal-list-cancel').onclick = closeModal;
document.getElementById('modal-template-switch').onclick = () => {
  closeModal();
  sendMessage(TEMPLATE_SWITCH_VALUE);
};

document.getElementById('modal-cancel').onclick = closeModal;
document.getElementById('modal-submit').onclick = () => {
  const ta = document.getElementById('modal-textarea');
  const val = ta.value.trim();
  if (!val){ ta.focus(); return; }
  closeModal();
  sendMessage(val);
};

function setBusy(busy){
  typingEl.style.display = busy ? 'block' : 'none';
  sendBtn.disabled = busy;
  generateBtn.disabled = busy;
  if (busy) scrollBottom();
}

async function startConversation(){
  setBusy(true);
  const res = await fetch('/api/start', {method:'POST'});
  const data = await res.json();
  setBusy(false);
  if (data.success){
    convId = data.conv_id;
    sessionStorage.setItem('dratido_conv_id', convId);
    canGenerate = false; hasDraft = false;
    generateBtn.style.display = 'none';
    panelFooter.style.display = 'none';
    panelBody.innerHTML = '<div class="placeholder">Your draft will appear here once we\'ve brainstormed enough to generate it.</div>';
    renderMessages(data.messages);
  }
}

async function sendMessage(text){
  if (!text || !text.trim() || !convId) return;
  inputEl.value = '';
  autoGrow();
  setBusy(true);
  const res = await fetch('/api/message', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conv_id: convId, text: text})
  });
  const data = await res.json();
  setBusy(false);
  if (data.success){
    renderMessages(data.messages);
    canGenerate = !!data.can_generate;
    generateBtn.style.display = canGenerate ? 'inline-block' : 'none';
  } else {
    alert(data.message || 'Something went wrong.');
  }
}

async function generateDraft(){
  if (!convId) return;
  setBusy(true);
  typingEl.textContent = 'Starting…';
  let res, data;
  try {
    res = await fetch('/api/generate', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({conv_id: convId})
    });
    data = await res.json();
  } catch (e) { setBusy(false); typingEl.textContent = 'Dratido is thinking…'; alert('Could not reach the server.'); return; }
  if (!data.success){
    setBusy(false); typingEl.textContent = 'Dratido is thinking…';
    alert(data.message || 'Could not generate the draft.');
    return;
  }
  // poll until the background job finishes (a full memorial takes a few minutes)
  while (true){
    await new Promise(r => setTimeout(r, 2500));
    let st;
    try {
      const r = await fetch('/api/generate_status/' + convId);
      st = await r.json();
    } catch (e) { continue; }          // transient network blip: keep polling
    if (st.state === 'running'){
      const pct = st.total ? Math.min(99, Math.round(100 * st.done / st.total)) : 0;
      typingEl.textContent = (st.label || 'Working…') + '  (' + pct + '%)';
      continue;
    }
    setBusy(false);
    typingEl.textContent = 'Dratido is thinking…';
    if (st.state === 'done'){
      renderMessages(st.messages);
      showDraft(st.draft_text);
      openPanel();
    } else {
      alert(st.message || 'Could not generate the draft.');
    }
    break;
  }
}

function showDraft(text){
  hasDraft = true;
  panelBody.innerHTML = '';
  const pre = document.createElement('div');
  pre.id = 'draft-text';
  pre.textContent = text;
  panelBody.appendChild(pre);
  panelFooter.style.display = 'block';
}

function openPanel(){ panelEl.classList.add('open'); }
function closePanel(){ panelEl.classList.remove('open'); }

document.getElementById('panel-toggle').onclick = () => {
  panelEl.classList.contains('open') ? closePanel() : openPanel();
};
document.getElementById('panel-close').onclick = closePanel;
document.getElementById('new-draft-btn').onclick = () => {
  closePanel();
  startConversation();
};
document.getElementById('download-btn').onclick = () => {
  if (convId) window.location.href = '/api/download/' + convId;
};
sendBtn.onclick = () => sendMessage(inputEl.value);
generateBtn.onclick = generateDraft;
inputEl.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey){
    e.preventDefault();
    sendMessage(inputEl.value);
  }
});
function autoGrow(){
  inputEl.style.height = 'auto';
  inputEl.style.height = Math.min(inputEl.scrollHeight, 140) + 'px';
}
inputEl.addEventListener('input', autoGrow);

startConversation();
</script>
</body>
</html>
"""


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    os.makedirs('generated', exist_ok=True)

    groq_key = os.environ.get('GROQ_API_KEY', '').strip()
    key_str = '\u2713 Groq \u2014 ready!' if groq_key else '\u2717 NOT SET \u2014 see below'
    print('\n' + '=' * 60)
    print(f'  {APP_NAME} \u2014 {APP_TAGLINE}')
    print('  AI drafting assistant — chat-first, no login')
    print('  Powered by Groq (free tier)')
    print('  Open browser:  http://127.0.0.1:8081')
    print(f'  GROQ_API_KEY: {key_str}')
    print('=' * 60 + '\n')
    if not groq_key:
        print('  Get your free Groq API key at https://console.groq.com')
        print('  then:  export GROQ_API_KEY=your_key_here   (Mac/Linux)')
        print('         set GROQ_API_KEY=your_key_here       (Windows)\n')

    port = int(os.environ.get('PORT', 8081))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
