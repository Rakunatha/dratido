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

import os, re, time, uuid, json, base64, zlib, threading, traceback
import requests
from requests.adapters import HTTPAdapter
from flask import Flask, request, jsonify, send_file, Response
from werkzeug.exceptions import HTTPException
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from pypdf import PdfReader

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
GENERATED_DIR = os.path.join(BASE_DIR, 'generated')

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024   # hard cap on any request body

APP_NAME    = 'Dratido'
APP_TAGLINE = 'Draft Till Done'

# In-memory conversation store: conv_id -> conversation state (no DB, no login).
# Idle conversations are evicted (see _evict_stale_conversations) so memory and the
# generated/ folder don't grow forever on a long-running free-tier instance.
CONVS = {}
CONVS_LOCK = threading.Lock()
CONV_TTL_SECONDS = 6 * 3600
MAX_CONVS = 500
MAX_MESSAGE_CHARS = 30000


# ══════════════════════════════════════════════════════════════════[...]
#  AI CLIENT  (Groq — fast free inference)
# ══════════════════════════════════════════════════════════════════[...]
#
# Speed design:
#   * one pooled HTTPS session (no TLS handshake per call)
#   * the model list is discovered once and cached (was: an extra HTTP call per AI call)
#   * on a 429 we fail over to the next model immediately (each Groq model has its own
#     rate-limit bucket) instead of sleeping 4/8/16 s on the same one
#   * small/fast model for the one-word classification call
#   * per-call max_tokens sized to the job (a 4096 cap on a one-word answer, or on a
#     short chat reply, wastes rate-limit budget; a 4096 cap on a full memorial truncates it)
#   * the final draft is streamed so the user sees text within ~1 s, and is
#     auto-continued if the model stops because it hit the token limit

GROQ_BASE_URL = os.environ.get("GROQ_BASE_URL", "https://api.groq.com/openai/v1").rstrip("/")

_GROQ_PREFERRED_MODELS = [
    "llama-3.3-70b-versatile",
    "openai/gpt-oss-120b",
    "llama-3.1-8b-instant",
    "openai/gpt-oss-20b",
]
_GROQ_FAST_MODELS = ["llama-3.1-8b-instant", "openai/gpt-oss-20b"]
_GROQ_EXCLUDED = ("whisper", "guard", "safeguard", "compound", "tts", "orpheus",
                  "playai", "embed", "moderation")
MAX_MODELS_TRIED = 4
RATE_LIMIT_MAX_WAIT = 12          # seconds; longest we'll ever sleep on a 429

def _int_env(name, default):
    try:
        return int(os.environ.get(name, "").strip() or default)
    except ValueError:
        return default

BRAINSTORM_MAX_TOKENS = _int_env("BRAINSTORM_MAX_TOKENS", 900)
DRAFT_MAX_TOKENS      = _int_env("DRAFT_MAX_TOKENS", 5000)
MOOT_MAX_TOKENS       = _int_env("MOOT_MAX_TOKENS", 6000)

_HTTP = requests.Session()
_HTTP.mount("https://", HTTPAdapter(pool_connections=4, pool_maxsize=16))
_HTTP.mount("http://",  HTTPAdapter(pool_connections=4, pool_maxsize=16))

_MODEL_CACHE = {"all": None, "available": set(), "expires": 0.0}
_MODEL_LOCK = threading.Lock()


def _discover_groq_models(api_key):
    try:
        resp = _HTTP.get(f"{GROQ_BASE_URL}/models",
                         headers={"Authorization": f"Bearer {api_key}"}, timeout=(4, 6))
        if resp.status_code != 200:
            return [], f"HTTP {resp.status_code} from Groq /models: {resp.text[:200]}"
        data = resp.json()
        items = data.get("data", []) if isinstance(data, dict) else []
        ids = [i["id"] for i in items
               if isinstance(i, dict) and i.get("id") and i.get("active", True)]
        return ids, None
    except Exception as e:
        return [], f"Could not query Groq /models: {e}"


def get_groq_models(api_key, fast=False):
    """Ordered list of models to try. Discovery runs at most once per hour (30 s after a
    failed discovery) instead of on every AI call."""
    now = time.time()
    with _MODEL_LOCK:
        if _MODEL_CACHE["all"] is None or now > _MODEL_CACHE["expires"]:
            available, err = _discover_groq_models(api_key)
            if err:
                print(f"[Groq] Model discovery warning: {err}")
            avail_set = set(available)
            override = os.environ.get("GROQ_MODEL", "").strip()
            if available:
                selected = [m for m in _GROQ_PREFERRED_MODELS if m in avail_set]
                selected += [m for m in available
                             if m not in selected
                             and not any(x in m.lower() for x in _GROQ_EXCLUDED)]
            else:
                selected = list(_GROQ_PREFERRED_MODELS)
            if override:
                selected = [override] + [m for m in selected if m != override]
            _MODEL_CACHE["all"] = list(dict.fromkeys(selected))
            _MODEL_CACHE["available"] = avail_set
            _MODEL_CACHE["expires"] = now + (3600 if available else 30)
            print(f"[Groq] Models: {_MODEL_CACHE['all'][:MAX_MODELS_TRIED]}")
        models = list(_MODEL_CACHE["all"])
        avail = _MODEL_CACHE["available"]

    if fast:
        fast_override = os.environ.get("GROQ_FAST_MODEL", "").strip()
        head = [fast_override] if fast_override else []
        head += [m for m in _GROQ_FAST_MODELS if not avail or m in avail]
        models = list(dict.fromkeys(head + models))
    return models[:MAX_MODELS_TRIED]


class _Skip(Exception):
    """This model can't serve the request — try the next one."""

class _RateLimited(Exception):
    def __init__(self, wait):
        super().__init__(f"rate limited, retry in {wait}s")
        self.wait = wait


def _retry_after(resp):
    try:
        return max(0.5, min(float(resp.headers.get("retry-after", "")), 60.0))
    except (TypeError, ValueError):
        return 5.0


def _chat_once(api_key, model, messages, temperature, max_tokens, on_text=None):
    """One attempt against one model. Returns (text, finish_reason).
    If on_text is given the response is streamed and on_text(accumulated_text) is called
    as tokens arrive."""
    stream = on_text is not None
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_completion_tokens": max_tokens,
        "stream": stream,
    }
    if model.startswith("openai/gpt-oss"):
        payload["reasoning_effort"] = "low"     # reasoning tokens are pure latency here

    try:
        resp = _HTTP.post(f"{GROQ_BASE_URL}/chat/completions",
                          headers={"Authorization": f"Bearer {api_key}",
                                   "Content-Type": "application/json"},
                          json=payload, timeout=(5, 60), stream=stream)
    except requests.exceptions.Timeout:
        raise _Skip(f"Timeout on {model}")
    except requests.exceptions.RequestException as e:
        raise _Skip(f"Request error on {model}: {e}")

    try:
        if resp.status_code == 429:
            raise _RateLimited(_retry_after(resp))
        if resp.status_code != 200:
            raise _Skip(f"HTTP {resp.status_code} on {model}: {resp.text[:200]}")

        finish = None
        if not stream:
            try:
                data = resp.json()
                if "error" in data:
                    raise _Skip(f"API error on {model}: {data['error']}")
                choice = data["choices"][0]
                text = choice["message"]["content"] or ""
                finish = choice.get("finish_reason")
            except _Skip:
                raise
            except Exception as e:
                raise _Skip(f"Unexpected response from {model}: {e}")
        else:
            acc = ""
            for raw in resp.iter_lines():
                if not raw:
                    continue
                line = raw.decode("utf-8", "replace")
                if not line.startswith("data:"):
                    continue
                body = line[5:].strip()
                if body == "[DONE]":
                    break
                try:
                    evt = json.loads(body)
                except ValueError:
                    continue
                if "error" in evt:
                    raise _Skip(f"Stream error on {model}: {evt['error']}")
                ch = (evt.get("choices") or [{}])[0]
                delta = (ch.get("delta") or {}).get("content")
                if delta:
                    acc += delta
                    on_text(acc)
                if ch.get("finish_reason"):
                    finish = ch["finish_reason"]
            text = acc
    except requests.exceptions.RequestException as e:
        raise _Skip(f"Connection dropped on {model}: {e}")
    finally:
        resp.close()

    if not text.strip():
        raise _Skip(f"Empty content from {model}")
    return text, finish


def _chat_with_failover(api_key, models, messages, temperature, max_tokens, on_text=None):
    last_error = None
    for rnd in range(2):
        min_wait = None
        for model in models:
            try:
                text, finish = _chat_once(api_key, model, messages, temperature,
                                          max_tokens, on_text)
                print(f"[Groq] ✓ {model} ({len(text)} chars, finish={finish})")
                return text, finish, model
            except _RateLimited as e:
                last_error = f"429 rate-limited on {model}"
                min_wait = e.wait if min_wait is None else min(min_wait, e.wait)
                print(f"[Groq] {last_error} — trying next model")
            except _Skip as e:
                last_error = str(e)
                print(f"[Groq] {last_error[:160]} — trying next model")
        # Every model failed. Only worth a second pass if the cause was rate limiting.
        if min_wait is None or rnd == 1:
            break
        time.sleep(min(min_wait, RATE_LIMIT_MAX_WAIT))
    raise RuntimeError(f"All Groq models failed. Last error: {last_error}")


def ai_chat(messages: list, temperature: float = 0.6, max_tokens: int = BRAINSTORM_MAX_TOKENS,
            fast: bool = False, on_text=None, max_continuations: int = 0) -> str:
    """Call Groq's chat-completions API (multi-turn) with model failover.

    fast=True              prefers the small/fast models (classification, short answers)
    on_text(text)          streams the response; called with the text so far
    max_continuations=N    if the model stops because it hit max_tokens, ask it to carry
                           on (up to N times) so long documents are never silently cut off
    """
    api_key = os.environ.get("GROQ_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("GROQ_API_KEY not set. Get a free key at https://console.groq.com")

    models = get_groq_models(api_key, fast=fast)
    if not models:
        raise RuntimeError("No active Groq text models are available to this API key/project.")

    text, finish, used = _chat_with_failover(api_key, models, messages, temperature,
                                             max_tokens, on_text)
    for _ in range(max_continuations):
        if finish != "length":
            break
        print("[Groq] Output hit the token limit — requesting continuation")
        order = [used] + [m for m in models if m != used]
        cont_msgs = messages + [
            {"role": "assistant", "content": text},
            {"role": "user", "content":
                "Continue the document exactly where you stopped. Do not repeat anything "
                "already written, and do not add commentary — output only the continuation."},
        ]
        prefix = text
        cb = (lambda t, _p=prefix: on_text(_p + t)) if on_text else None
        try:
            more, finish, used = _chat_with_failover(api_key, order, cont_msgs, temperature,
                                                     max_tokens, cb)
        except RuntimeError as e:
            print(f"[Groq] Continuation failed, returning partial text: {e}")
            break
        text += more
    return text.strip()


def ai_generate(prompt: str, system: str = "", temperature: float = 0.6, **kw) -> str:
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    return ai_chat(messages, temperature=temperature, **kw)


# ══════════════════════════════════════════════════════════════════[...]
#  TEMPLATE FILE EXTRACTION  (.docx / .pdf uploads)
# ══════════════════════════════════════════════════════════════════[...]

ALLOWED_TEMPLATE_EXTENSIONS = {'.docx', '.pdf'}
MAX_TEMPLATE_UPLOAD_BYTES = 15 * 1024 * 1024  # 15 MB


MAX_TEMPLATE_CHARS = 20000   # the draft prompt only uses the first ~6000; no point reading 200 pages


def _iter_block_items(doc):
    """Yield paragraphs and tables in true document order."""
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)


def extract_text_from_docx(file_stream) -> str:
    """Pull readable text (paragraphs + table cells, in document order) out of an
    uploaded .docx reference template."""
    doc = Document(file_stream)
    parts, total = [], 0
    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                parts.append(t)
                total += len(t)
        else:
            for row in block.rows:
                seen, cells = set(), []
                for c in row.cells:
                    if c._tc in seen:          # merged cells are returned once per grid column
                        continue
                    seen.add(c._tc)
                    if c.text.strip():
                        cells.append(c.text.strip())
                if cells:
                    line = '\t'.join(cells)
                    parts.append(line)
                    total += len(line)
        if total >= MAX_TEMPLATE_CHARS:
            break
    return '\n'.join(parts).strip()


def extract_text_from_pdf(file_stream) -> str:
    """Pull readable text out of an uploaded .pdf reference template, page by page.
    Scanned/image-only PDFs will yield little or no text — callers should treat an
    empty result as a failure and ask the user for another file."""
    reader = PdfReader(file_stream)
    if getattr(reader, "is_encrypted", False):
        try:
            ok = reader.decrypt('')
        except Exception:
            ok = 0
        if not ok:
            raise ValueError("PDF is password-protected")
    parts, total = [], 0
    for page in reader.pages[:60]:
        text = (page.extract_text() or '').strip()
        if text:
            parts.append(text)
            total += len(text)
            if total >= MAX_TEMPLATE_CHARS:
                break
    return '\n\n'.join(parts).strip()


# ══════════════════════════════════════════════════════════════════[...]
#  MOOT MEMORIAL TEMPLATE LIBRARY  (compressed + encrypted bundled asset)
# ══════════════════════════════════════════════════════════════════[...]
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
    key = base64.urlsafe_b64encode(kdf.derive(passphrase.encode('utf-8')))
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


def select_moot_reference(side_key: str, user_text: str) -> str:
    """Pick the stored template that best matches the user's side and (loosely, by
    keyword overlap) their moot problem, and return it formatted as a style-reference
    block. Returns '' if the library is unavailable."""
    templates = load_moot_templates()
    if not templates:
        return ""

    matching = [t for t in templates if t.get("side") == side_key] or templates

    def _completeness(t):
        return sum(1 for k in ("abbreviations_sample", "authorities_style", "jurisdiction_example",
                               "issues_style", "arguments_style", "prayer_example")
                   if (t.get(k) or "").strip())

    user_words = set(re.findall(r'[a-zA-Z]{4,}', (user_text or "").lower()))
    if not user_words:
        chosen = max(matching, key=_completeness)
    else:
        def _score(t):
            blob = ' '.join([t.get('court', ''), t.get('parties', ''),
                              t.get('issues_style', ''), t.get('arguments_style', '')]).lower()
            blob_words = set(re.findall(r'[a-zA-Z]{4,}', blob))
            return len(user_words & blob_words)
        chosen = max(matching, key=lambda t: (_score(t), _completeness(t)))

    return _format_moot_reference(chosen)


# ══════════════════════════════════════════════════════════════════[...]
#  DOCX BUILDING
# ══════════════════════════════════════════════════════════════════[...]


_TNR = 'Times New Roman'
_NUMBERED_RE = re.compile(r'^\s*(\d{1,3})[\.\)]\s+(.*)$')
_BULLET_RE   = re.compile(r'^(?:[-•\u2022]|\*(?!\*))\s+(.*)$')
_SUBITEM_RE  = re.compile(r'^\(?([a-zA-Z]|[ivxIVX]{1,5}|\d{1,2})\)\s+(.*)$')
_RULE_RE     = re.compile(r'^(?:-{3,}|\*{3,}|={3,}|`{3,}.*)$')   # markdown rules / code fences
_INLINE_RE   = re.compile(r'(\*\*[^*\n]+?\*\*|(?<![*\w])\*[^*\s][^*\n]*?\*(?![*\w]))')


def _style_run(run, size=12, bold=None, italic=None):
    run.font.name = _TNR
    run.font.size = Pt(size)
    if bold:
        run.bold = True
    if italic:
        run.italic = True


def _add_runs(p, text, size=12, bold=False):
    """Add text to a paragraph, turning **bold** / *italic* markdown into real formatting
    (models emit it even when told not to) and dropping any stray asterisks."""
    for part in _INLINE_RE.split(text):
        if not part:
            continue
        b, it = bold, False
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            part, b = part[2:-2], True
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            part, it = part[1:-1], True
        part = part.replace('**', '')
        if part:
            _style_run(p.add_run(part), size=size, bold=b, italic=it)


def build_ai_legal_docx(doc_type: str, ai_text: str) -> str:
    """Convert the AI-drafted plain-text legal document into a formatted,
    .docx file resembling a formal court filing."""
    doc = Document()
    for sec in doc.sections:
        sec.page_width    = Inches(8.5)
        sec.page_height   = Inches(11)
        sec.top_margin    = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin   = Inches(1.25)
        sec.right_margin  = Inches(1.25)

    normal = doc.styles['Normal']
    normal.font.name = _TNR
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), _TNR)

    title_written = False

    for ln in ai_text.strip().split('\n'):
        stripped = ln.strip()
        if not stripped or _RULE_RE.match(stripped):
            continue
        clean = re.sub(r'^#{1,6}\s*', '', stripped)
        clean = re.sub(r'\s+#{2,}$', '', clean).strip()
        whole_bold = re.fullmatch(r'\*\*(.+?)\*\*', clean)
        if whole_bold:
            clean = whole_bold.group(1).strip()
        if not clean:
            continue

        m_num = _NUMBERED_RE.match(clean)
        m_bul = None if m_num else _BULLET_RE.match(clean)
        m_sub = None if (m_num or m_bul) else _SUBITEM_RE.match(clean)

        if not title_written and not m_num and not m_bul:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(16)
            _add_runs(p, clean.upper().replace('**', ''), size=16, bold=True)
            title_written = True

        elif m_num:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(4)
            p.paragraph_format.left_indent  = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            _style_run(p.add_run(f'{m_num.group(1)}.  '), bold=True)
            _add_runs(p, m_num.group(2))

        elif m_bul:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            p.paragraph_format.left_indent  = Inches(0.75)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            _style_run(p.add_run('•  '))
            _add_runs(p, m_bul.group(1))

        elif m_sub:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after  = Pt(3)
            p.paragraph_format.left_indent  = Inches(0.9)
            p.paragraph_format.first_line_indent = Inches(-0.4)
            _style_run(p.add_run(f'({m_sub.group(1)})  '))
            _add_runs(p, m_sub.group(2))

        elif clean.replace('**', '').isupper() and len(clean) < 80:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(8)
            _add_runs(p, clean.replace('**', ''), size=13, bold=True)

        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after  = Pt(6)
            _add_runs(p, clean)

    os.makedirs(GENERATED_DIR, exist_ok=True)
    safe = re.sub(r'[^\w\-]', '_', (doc_type or 'Legal_Draft')[:40]) or 'Legal_Draft'
    out  = os.path.join(GENERATED_DIR, f'{safe}_{uuid.uuid4().hex[:8]}.docx')
    doc.save(out)
    return out


# ══════════════════════════════════════════════════════════════════[...]
#  MOOT MEMORIAL DOCX BUILDER  (mirrors the format of real competition memorials)
# ══════════════════════════════════════════════════════════════════[...]
# The memorial is drafted by the AI in a tagged plain-text format (see DRAFT_SYSTEM_MOOT)
# and laid out here: bordered cover page, running header/footer ("P a g e | n"), table of
# contents and index of authorities with live page-number fields, 2-column abbreviations
# table, real Word footnotes, issue-wise numbered arguments, prayer and signature block.
from xml.sax.saxutils import escape as _xesc
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.shared import RGBColor
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as _RT

_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_FN_RE = re.compile(r'\{\{\s*fn\s*:(.*?)\}\}', re.S)
_MOOT_SECTIONS = ('COVER', 'CASES', 'STATUTES', 'BOOKS', 'WEBSITES', 'ABBREVIATIONS',
                  'JURISDICTION', 'FACTS', 'ISSUES', 'SUMMARY', 'ARGUMENTS', 'PRAYER', 'SIGNATURE')


def parse_moot_text(text: str) -> dict:
    """Split '@@SECTION' tagged text into {SECTION: [lines]}. Empty dict if untagged."""
    secs, cur = {}, None
    for ln in text.splitlines():
        m = re.match(r'^\s*@@\s*([A-Za-z_ ]+?)\s*:?\s*$', ln)
        if m:
            name = m.group(1).strip().upper().replace(' ', '_')
            cur = name if name in _MOOT_SECTIONS else None
            if cur:
                secs.setdefault(cur, [])
            continue
        if cur is not None and ln.strip():
            secs[cur].append(ln.strip())
    return secs if ('ARGUMENTS' in secs and 'COVER' in secs) else {}


def _tidy(s):
    s = re.sub(r'^#{1,6}\s*', '', s.strip())
    return s.replace('**', '').replace('__', '')


def _border(p, edge, sz=24, val='single', space=4):
    pPr = p._p.get_or_add_pPr()
    bdr = pPr.find(qn('w:pBdr'))
    if bdr is None:
        bdr = OxmlElement('w:pBdr'); pPr.append(bdr)
    e = OxmlElement(f'w:{edge}')
    for k, v in (('val', val), ('sz', str(sz)), ('space', str(space)), ('color', '000000')):
        e.set(qn(f'w:{k}'), v)
    bdr.append(e)


def _field(p, instr, cached='1', size=12, bold=False):
    def r(child):
        run = p.add_run(); _style_run(run, size=size, bold=bold); run._r.append(child); return run
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin'); r(b)
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = f' {instr} '; r(it)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'), 'separate'); r(s)
    cached_run = p.add_run(cached)
    _style_run(cached_run, size=size, bold=bold)
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end'); r(e)
    return cached_run


def _bookmark(p, name, bid):
    s = OxmlElement('w:bookmarkStart'); s.set(qn('w:id'), str(bid)); s.set(qn('w:name'), name)
    e = OxmlElement('w:bookmarkEnd'); e.set(qn('w:id'), str(bid))
    return s, e


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


class _MootBuilder:
    def __init__(self, secs):
        self.secs = secs
        self.doc = Document()
        self.headings = []      # (level, text, bookmark)
        self.footnotes = []     # (id, text)
        self.fn_bm = {}         # footnote id -> bookmark name
        self._bid = 100
        self.refs = []          # (bookmark, cached-text run) for every PAGEREF field
        self.cover = self._parse_cover()
        self.side_title = self.cover.get('MEMORIAL_TITLE', 'MEMORIAL').upper()

    # ---------- helpers
    def _bm_id(self):
        self._bid += 1
        return self._bid

    def _parse_cover(self):
        d, parties = {}, []
        for ln in self.secs.get('COVER', []):
            k, _, v = ln.partition(':')
            k = k.strip().upper().replace(' ', '_')
            if k.startswith('PARTY'):
                parties.append(v.strip())
            else:
                d[k] = v.strip()
        d['PARTIES'] = parties
        return d

    def para(self, text='', size=12, bold=False, italic=False, align='justify', before=0, after=6,
             line=None, left=None, first=None, keep=False):
        p = self.doc.add_paragraph()
        p.alignment = {'justify': WD_ALIGN_PARAGRAPH.JUSTIFY, 'center': WD_ALIGN_PARAGRAPH.CENTER,
                       'left': WD_ALIGN_PARAGRAPH.LEFT, 'right': WD_ALIGN_PARAGRAPH.RIGHT}[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        if line:
            pf.line_spacing = line
        if left is not None:
            pf.left_indent = Inches(left)
        if first is not None:
            pf.first_line_indent = Inches(first)
        if keep:
            pf.keep_with_next = True
        if text:
            self.rich(p, text, size=size, bold=bold, italic=italic)
        return p

    def rich(self, p, text, size=12, bold=False, italic=False):
        """Text with inline {{fn: ...}} footnotes and tidy markdown."""
        pos = 0
        for m in _FN_RE.finditer(text):
            self._plain(p, text[pos:m.start()], size, bold, italic)
            self._footnote_ref(p, m.group(1).strip())
            pos = m.end()
        self._plain(p, text[pos:], size, bold, italic)

    def _plain(self, p, t, size, bold, italic):
        if not t:
            return
        if italic:
            for part in _INLINE_RE.split(t):
                if part:
                    _style_run(p.add_run(part.replace('**', '').strip('*') if part.startswith('*') else part),
                               size=size, bold=bold, italic=True)
        else:
            _add_runs(p, t, size=size, bold=bold)

    def _footnote_ref(self, p, text):
        fid = len(self.footnotes) + 1
        self.footnotes.append((fid, _tidy(text)))
        name = f'fn_{fid}'
        self.fn_bm[fid] = name
        bs, be = _bookmark(p, name, self._bm_id())
        run = p.add_run()
        rPr = run._r.get_or_add_rPr()
        va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rPr.append(va)
        ref = OxmlElement('w:footnoteReference'); ref.set(qn('w:id'), str(fid))
        p._p.append(bs); p._p.append(run._r); p._p.append(be)
        run._r.append(ref)

    def heading(self, text, level=1, center=None, pbb=None, before=6, after=10):
        text = _tidy(text)
        p = self.doc.add_paragraph(style=f'Heading {level}')
        if center is None:
            center = (level == 1)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.keep_with_next = Pt(before), Pt(after), True
        pf.page_break_before = (level == 1) if pbb is None else pbb
        name = f'h_{len(self.headings) + 1}'
        bs, be = _bookmark(p, name, self._bm_id())
        p._p.append(bs)
        _style_run(p.add_run(text.upper() if level == 1 else text), size=14 if level == 1 else 12, bold=True)
        p._p.append(be)
        self.headings.append((level, text.upper() if level == 1 else text, name))
        return p

    def numbered(self, n, text, **kw):
        p = self.para('', line=1.5, **kw)
        _style_run(p.add_run(f'{n}.  '), bold=False)
        self.rich(p, text)
        return p

    def fill_at(self, marker, fn):
        """Run fn() (which appends to the document end) then move its output before marker."""
        body = self.doc.element.body
        n0 = len(body)
        fn()
        new = list(body)[n0 - 1:len(body) - 1]
        for el in new:
            marker._p.addprevious(el)
        marker._p.getparent().remove(marker._p)

    # ---------- styles / sections
    def setup(self):
        d = self.doc
        for lvl in (1, 2, 3):
            st = d.styles[f'Heading {lvl}']
            st.font.name = _TNR; st.font.size = Pt(12); st.font.bold = True
            st.font.color.rgb = RGBColor(0, 0, 0)
            rpr = st.element.get_or_add_rPr()
            rf = rpr.find(qn('w:rFonts'))
            if rf is None:
                rf = OxmlElement('w:rFonts'); rpr.append(rf)
            for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                rf.set(qn(f'w:{a}'), _TNR)
        n = d.styles['Normal']
        n.font.name = _TNR; n.font.size = Pt(12)
        n.element.get_or_add_rPr().find(qn('w:rFonts')).set(qn('w:eastAsia'), _TNR)
        s = d.sections[0]
        s.page_width, s.page_height = Inches(8.27), Inches(11.69)
        s.top_margin = s.bottom_margin = Inches(1)
        s.left_margin = s.right_margin = Inches(1.25)

    def build_cover(self):
        c, d = self.cover, self.doc
        t = d.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.RIGHT
        t.style = 'Table Grid'
        t.autofit = False
        t.columns[0].width = Inches(1.9)
        cell = t.rows[0].cells[0]
        cell.width = Inches(1.9)
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'FFFFFF')
        tcPr.append(shd)
        cp = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _style_run(cp.add_run(f"Team Code: {c.get('TEAM_CODE', '______')}"), size=11)
        self.para('', after=14)
        p = self.para(c.get('COMPETITION', 'MOOT COURT COMPETITION'), bold=True, align='center', before=8, after=8)
        _border(p, 'top', 36); _border(p, 'bottom', 36)
        self.para('Before', italic=True, align='center', before=14, after=10)
        p = self.para(c.get('COURT', ''), align='center', after=10)
        for r in p.runs: r.font.small_caps = True
        _border(p, 'bottom', 24)
        if c.get('CASE_NO'):
            p = self.para(c['CASE_NO'], align='center', before=10, after=14)
            for r in p.runs: r.font.small_caps = True
        if c.get('FILED_UNDER'):
            p = self.para(c['FILED_UNDER'], align='center', before=8, after=8)
            _border(p, 'top', 36); _border(p, 'bottom', 36)
        if c.get('SUBJECT'):
            p = self.para(c['SUBJECT'], size=10, align='center', before=14, after=14)
            for r in p.runs: r.font.small_caps = True
        p = self.para('IN THE MATTER BETWEEN:', size=10, align='center', before=10, after=14)
        parties = c['PARTIES'] or ['PARTY ONE | PETITIONER', 'PARTY TWO | RESPONDENT']
        for i, pt in enumerate(parties[:2]):
            name, _, role = pt.partition('|')
            p = self.para('', align='left', before=6, after=6)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(5.77), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            _style_run(p.add_run(f'{name.strip().upper()}\t{role.strip().upper()}'))
            if i == 0:
                _border(p, 'top', 12, 'double'); 
                self.para('v', size=10, align='center', before=8, after=8)
            else:
                _border(p, 'bottom', 12, 'double')
        p = self.para(self.side_title, bold=True, align='center', before=40, after=8)
        _border(p, 'bottom', 36)

    def cover_color(self):
        s = (self.side_title + ' ' + ' '.join(self.cover.get('PARTIES', []))[-0:]).lower()
        s = self.side_title.lower()
        return 'E0312B' if re.search(r'respondent|defendant|opposite|non-?appellant', s) else '4472C4'

    def cover_background(self):
        """Full-page colour rectangle, anchored to the page behind the text, in the cover
        section's header (so it shows on the cover page only)."""
        hp = self.doc.sections[0].header.paragraphs[0]
        hp.paragraph_format.space_after = Pt(0)
        xml = (f'<w:r xmlns:w="{_W_NS}" xmlns:v="urn:schemas-microsoft-com:vml"><w:pict>'
               f'<v:rect style="position:absolute;margin-left:0;margin-top:0;width:{8.27*72:.1f}pt;'
               f'height:{11.69*72:.1f}pt;z-index:-251658240;mso-position-horizontal-relative:page;'
               f'mso-position-vertical-relative:page" fillcolor="#{self.cover_color()}" stroked="f"/>'
               f'</w:pict></w:r>')
        hp._p.append(parse_xml(xml))

    def first_section_border(self):
        self.cover_background()
        sect = self.doc.sections[0]._sectPr
        pg = parse_xml(
            f'<w:pgBorders xmlns:w="{_W_NS}" w:offsetFrom="page">' +
            ''.join(f'<w:{e} w:val="thinThickSmallGap" w:sz="24" w:space="24" w:color="000000"/>'
                    for e in ('top', 'left', 'bottom', 'right')) + '</w:pgBorders>')
        pgmar = sect.find(qn('w:pgMar'))
        pgmar.addnext(pg)

    def header_footer(self):
        sec = self.doc.sections[1]
        sec.header.is_linked_to_previous = False
        sec.footer.is_linked_to_previous = False
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _style_run(hp.add_run(self.side_title), size=10, bold=True)
        _border(hp, 'bottom', 8, space=2)
        fp = sec.footer.paragraphs[0]
        fp.style = self.doc.styles['Normal']
        hp.style = self.doc.styles['Normal']
        fp.paragraph_format.tab_stops.add_tab_stop(Inches(5.77), WD_TAB_ALIGNMENT.RIGHT)
        _style_run(fp.add_run(f"Team Code: {self.cover.get('TEAM_CODE', '______')}\tP a g e | "), size=10)
        _field(fp, 'PAGE', size=10)

    # ---------- body sections
    def toc_entries(self):
        for lvl, text, bm in self.headings:
            p = self.para('', align='left', after=4, left=0.3 * (lvl - 1))
            p.paragraph_format.tab_stops.add_tab_stop(Inches(5.77), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            _style_run(p.add_run(text + '\t'), bold=(lvl == 1))
            self.refs.append((bm, _field(p, f'PAGEREF {bm} \\h', bold=(lvl == 1))))

    def authorities(self):
        self.heading('Index of Authorities', 1)
        groups = (('CASES', 'Cases Referred', True), ('STATUTES', 'Statutes Referred', False),
                  ('BOOKS', 'Books Referred', False), ('WEBSITES', 'Websites Referred', False))
        fn_norm = [(fid, _norm(t)) for fid, t in self.footnotes]
        derived = self._cases_from_footnotes()
        for key, title, is_cases in groups:
            items = [re.sub(r'^\s*(?:\d{1,3}[\.\)]|[-•*])\s*', '', _tidy(x)) for x in self.secs.get(key, [])]
            if is_cases:
                seen = {_norm(re.split(r'[\[\(,]', x, 1)[0]) for x in items}
                items += [d for d in derived if _norm(re.split(r'[\[\(,]', d, 1)[0]) not in seen]
            if not items:
                continue
            self.heading(title, 2, center=False, pbb=False, before=10, after=6)
            if is_cases:
                items.sort(key=lambda s: s.lower())
            for i, it in enumerate(items, 1):
                p = self.para('', align='left', after=4, left=0.35, first=-0.35)
                p.paragraph_format.tab_stops.add_tab_stop(Inches(5.77), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                _style_run(p.add_run((f'{i}.  ' if is_cases else '•  ') + it))
                if is_cases:
                    key_s = _norm(re.split(r'[\[\(,]', it, 1)[0])
                    hit = next((fid for fid, t in fn_norm if key_s and key_s in t), None)
                    if hit:
                        _style_run(p.add_run('\t'))
                        self.refs.append((self.fn_bm[hit], _field(p, f'PAGEREF {self.fn_bm[hit]} \\h')))

    def _cases_from_footnotes(self):
        out, seen = [], set()
        for _, t in self.footnotes:
            for seg in re.split(r';', t):
                seg = re.sub(r'^\s*(?:see also|see|also|cf\.?)\s*:?\s*', '', seg.strip(), flags=re.I).rstrip('.').strip()
                if re.search(r'\s[vV]s?\.?\s', seg) and re.search(r'[\[\(]\s*\d{4}|\b(AIR|SCC|SCR)\b', seg):
                    k = _norm(re.split(r'[\[\(,]', seg, 1)[0])
                    if k and k not in seen:
                        seen.add(k); out.append(seg)
        return out

    def abbreviations(self):
        self.heading('Index of Abbreviations', 1)
        rows = [ln.split('|', 1) for ln in self.secs.get('ABBREVIATIONS', []) if '|' in ln]
        rows.sort(key=lambda r: r[0].strip().lower())
        t = self.doc.add_table(rows=1, cols=2)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for c, h in zip(t.rows[0].cells, ('Abbreviation', 'Full Form')):
            _style_run(c.paragraphs[0].add_run(h), bold=True)
        for a, f in rows:
            cells = t.add_row().cells
            _style_run(cells[0].paragraphs[0].add_run(_tidy(a)))
            _style_run(cells[1].paragraphs[0].add_run(_tidy(f)))
        for r in t.rows:
            r.cells[0].width, r.cells[1].width = Inches(1.7), Inches(4.3)

    def plain_section(self, key, title, numbered=True):
        self.heading(title, 1)
        n = 0
        for ln in self.secs.get(key, []):
            m = _NUMBERED_RE.match(ln)
            if numbered and m:
                n += 1
                self.numbered(n, m.group(2), before=0)
            else:
                self.para(_tidy(ln), line=1.5)

    def jurisdiction(self):
        self.plain_section('JURISDICTION', 'Statement of Jurisdiction', numbered=False)

    def issues(self):
        self.heading('Issues Raised', 1)
        for ln in self.secs.get('ISSUES', []):
            ln = _tidy(ln)
            m = re.match(r'^(ISSUE\s*[0-9IVX]+\s*[:.\-]?)\s*(.*)$', ln, re.I)
            if m:
                self.para(m.group(1).upper().rstrip(' .-') + (':' if not m.group(1).strip().endswith(':') else ''),
                          bold=True, align='left', before=10, after=4, keep=True)
                if m.group(2):
                    self.para(m.group(2), line=1.5)
            else:
                self.para(ln, line=1.5)

    def summary(self):
        self.heading('Summary of Arguments', 1)
        for ln in self.secs.get('SUMMARY', []):
            t = _tidy(ln) if not _FN_RE.search(ln) else ln
            if re.match(r'^ISSUE\s*[0-9IVX]+', t, re.I):
                self.para(t, bold=True, align='left', before=10, after=4, keep=True)
            else:
                self.para(t, line=1.5)

    def arguments(self):
        self.heading('Arguments Advanced', 1)
        n = 0
        for raw in self.secs.get('ARGUMENTS', []):
            if raw.startswith('###'):
                self.heading(raw.lstrip('#'), 3, center=False, pbb=False, before=6, after=6)
            elif raw.startswith('##'):
                self.heading(raw.lstrip('#'), 3, center=False, pbb=False, before=8, after=6)
            elif raw.startswith('#'):
                n = 0
                self.heading(raw.lstrip('#'), 2, center=False, pbb=False, before=14, after=8)
            else:
                m = _NUMBERED_RE.match(raw)
                sub = re.match(r'^\(?([a-z]|[ivx]{1,4})\)\s+(.*)$', raw)
                if m:
                    n += 1
                    self.numbered(n, m.group(2))
                elif sub:
                    p = self.para('', line=1.5, left=0.8, first=-0.4, after=3)
                    _style_run(p.add_run(f'{sub.group(1)})  '))
                    self.rich(p, sub.group(2))
                else:
                    self.para(raw.replace('**', ''), line=1.5)

    def prayer(self):
        self.heading('Prayer for Relief', 1)
        n = 0
        for ln in self.secs.get('PRAYER', []):
            m = _NUMBERED_RE.match(ln)
            if m:
                n += 1
                self.numbered(n, m.group(2))
            else:
                self.para(_tidy(ln), line=1.5, before=6)
        sig = [_tidy(x) for x in self.secs.get('SIGNATURE', [])]
        if sig:
            self.para('', after=18)
            t = self.doc.add_table(rows=1, cols=2)
            left = [s for s in sig if re.match(r'^(PLACE|DATE)\b', s, re.I)]
            right = [s for s in sig if s not in left]
            for cell, lines, al in ((t.rows[0].cells[0], left, WD_ALIGN_PARAGRAPH.LEFT),
                                    (t.rows[0].cells[1], right, WD_ALIGN_PARAGRAPH.RIGHT)):
                for i, s in enumerate(lines):
                    cp = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                    cp.alignment = al
                    _style_run(cp.add_run(s), bold=(al == WD_ALIGN_PARAGRAPH.RIGHT))

    # ---------- footnotes part + settings
    def attach_footnotes(self):
        if not self.footnotes:
            return
        rp = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>'
        out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="{_W_NS}">',
               '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
               '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']
        for fid, t in self.footnotes:
            out.append(f'<w:footnote w:id="{fid}"><w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
                       f'<w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:footnoteRef/></w:r>'
                       f'<w:r>{rp}<w:t xml:space="preserve"> {_xesc(t)}</w:t></w:r></w:p></w:footnote>')
        out.append('</w:footnotes>')
        pkg = self.doc.part.package
        part = Part(PackURI('/word/footnotes.xml'),
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
                    ''.join(out).encode('utf-8'), pkg)
        self.doc.part.relate_to(part, _RT.FOOTNOTES)

    def update_fields_on_open(self):
        st = self.doc.settings.element
        u = OxmlElement('w:updateFields'); u.set(qn('w:val'), 'true')
        st.append(u)

    def estimate_pages(self):
        """Word only recalculates PAGEREF fields when the user accepts the 'update fields' prompt;
        otherwise the cached text is shown. So we fill the cached text with an estimate of the real
        page (simulated pagination) — the TOC/index then look right even if the prompt is declined."""
        import math
        fn_len = {fid: len(t) for fid, t in self.footnotes}
        body = self.doc.element.body
        W = lambda t: qn('w:' + t)
        CAP, LH = 640.0, 13.8
        page, cur, pages = 1, 0.0, {}
        for el in list(body):
            if el.tag == W('tbl'):
                rows = len(el.findall(W('tr')))
                h = rows * 22.0
                if cur + h > CAP:
                    page, cur = page + 1, 0.0
                cur += h
                continue
            if el.tag != W('p'):
                continue
            pPr = el.find(W('pPr'))
            sp = pPr.find(W('spacing')) if pPr is not None else None
            line = (int(sp.get(W('line'), 240)) / 240.0) if sp is not None else 1.0
            before = int(sp.get(W('before'), 0)) / 20.0 if sp is not None else 0.0
            after = int(sp.get(W('after'), 0)) / 20.0 if sp is not None else 0.0
            if pPr is not None and pPr.find(W('pageBreakBefore')) is not None \
                    and not (pPr.find(W('pageBreakBefore')).get(W('val')) in ('0', 'false')):
                if cur > 0:
                    page, cur = page + 1, 0.0
            text = ''.join(t.text or '' for t in el.iter(W('t')))
            ind = pPr.find(W('ind')) if pPr is not None else None
            left = int(ind.get(W('left'), 0)) / 1440.0 if ind is not None else 0.0
            cpl = max(40, int((6.0 - left) * 14.5))
            lines = max(1, math.ceil(len(text) / cpl))
            h = lines * LH * line + before + after
            for ref in el.iter(W('footnoteReference')):
                h += math.ceil(fn_len.get(int(ref.get(W('id'))), 60) / 105.0) * 11.5 + 3
            if cur + h > CAP:
                page, cur = page + 1, 0.0
            for bs in el.iter(W('bookmarkStart')):
                pages[bs.get(W('name'))] = page
            cur += h
            if pPr is not None and pPr.find(W('sectPr')) is not None:
                page, cur = page + 1, 0.0
        for bm, run in self.refs:
            if bm in pages:
                run.text = str(pages[bm])

    def build(self):
        self.setup()
        self.build_cover()
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        self.first_section_border()
        self.header_footer()
        self.heading('Table of Contents', 1, pbb=False)
        m_toc = self.para('')
        m_auth = self.para('')
        self.abbreviations()
        self.jurisdiction()
        self.plain_section('FACTS', 'Statement of Facts')
        self.issues()
        self.summary()
        self.arguments()
        self.prayer()
        n0 = len(self.headings)
        self.fill_at(m_auth, self.authorities)
        auth_h = self.headings[n0:]
        self.headings = self.headings[:1] + auth_h + self.headings[1:n0]   # authorities follow the TOC entry
        self.fill_at(m_toc, self.toc_entries)
        self.estimate_pages()
        self.attach_footnotes()
        self.update_fields_on_open()
        return self.doc


def build_moot_docx(doc_type: str, ai_text: str):
    """Returns the saved .docx path, or None if the text isn't in the tagged memorial format."""
    secs = parse_moot_text(ai_text)
    if not secs:
        return None
    doc = _MootBuilder(secs).build()
    os.makedirs(GENERATED_DIR, exist_ok=True)
    safe = re.sub(r'[^\w\-]', '_', (doc_type or 'Moot_Memorial')[:40]) or 'Moot_Memorial'
    out = os.path.join(GENERATED_DIR, f'{safe}_{uuid.uuid4().hex[:8]}.docx')
    doc.save(out)
    return out


# ══════════════════════════════════════════════════════════════════[...]
#  CONVERSATION / DRAFTING WORKFLOW
# ══════════════════════════════════════════════════════════════════[...]
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
    {"label": "Petitioner / Appellant / Plaintiff",
     "value": "Petitioner / Appellant / Plaintiff side"},
    {"label": "Respondent / Defendant",
     "value": "Respondent / Defendant side"},
]

WELCOME_MSG = (
    "Hi, I'm Dratido — your drafting assistant. I'll help you brainstorm and put together a "
    "draft, then hand you a clean Word document at the end.\n\n"
    "How would you like to start?"
)


def _remove_file(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def _evict_stale_conversations():
    """Drop idle conversations (and their .docx files) so the process doesn't leak."""
    now = time.time()
    with CONVS_LOCK:
        def running(c):
            j = c.get("job")
            return bool(j and j.get("status") == "running")
        stale = [cid for cid, c in CONVS.items()
                 if now - c["touched"] > CONV_TTL_SECONDS and not running(c)]
        overflow = len(CONVS) - len(stale) - MAX_CONVS
        if overflow > 0:
            by_age = sorted((c["touched"], cid) for cid, c in CONVS.items()
                            if cid not in stale and not running(c))
            stale += [cid for _, cid in by_age[:overflow]]
        for cid in stale:
            _remove_file(CONVS.pop(cid).get("docx_path"))


def new_conversation() -> dict:
    _evict_stale_conversations()
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
        "job": None,           # background draft-generation job, if any
        "lock": threading.RLock(),   # serialises requests for one conversation
        "touched": time.time(),
    }
    with CONVS_LOCK:
        CONVS[conv_id] = conv
    return conv


def get_conversation(conv_id: str):
    conv = CONVS.get(conv_id)
    if conv:
        conv["touched"] = time.time()
    return conv


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
    "You are an expert moot-court coach and legal drafter. Draft a complete MEMORIAL in the exact "
    "structure and style of a competition-winning Indian moot memorial (Table of Contents, Index of "
    "Authorities, Index of Abbreviations, Statement of Jurisdiction, Statement of Facts, Issues Raised, "
    "Summary of Arguments, Arguments Advanced, Prayer for Relief).\n"
    "OUTPUT FORMAT — plain text only (no markdown, no asterisks, no code fences), using these section "
    "tags, each alone on its own line, in this exact order. A program lays the document out from them, "
    "so follow the tag syntax precisely and do NOT write a table of contents, page numbers, or headings "
    "for sections yourself:\n"
    "@@COVER\n"
    "TEAM_CODE: <code from the problem, else ______>\n"
    "COMPETITION: <full competition name, e.g. 4th X NATIONAL FAMILY LAW MOOT COURT COMPETITION, 2023>\n"
    "COURT: <court/forum before which the matter is argued>\n"
    "CASE_NO: <e.g. PLAINT NUMBER _______ / 2022  or  Criminal Appeal No. ____ of 2021>\n"
    "FILED_UNDER: <statute/section line, only if applicable>\n"
    "SUBJECT: <one-line description of the matter, e.g. IN THE CASE CONCERNING ...>\n"
    "PARTY1: <FIRST PARTY NAME> | <ROLE e.g. PLAINTIFF / PETITIONER / APPELLANT>\n"
    "PARTY2: <SECOND PARTY NAME> | <ROLE e.g. DEFENDANT / RESPONDENT>\n"
    "MEMORIAL_TITLE: MEMORIAL ON BEHALF OF <SIDE IN CAPS, e.g. PLAINTIFF>\n"
    "@@CASES  — one case per line: 'Party A v Party B, [Year] Vol Reporter Page' (no numbering)\n"
    "@@STATUTES  — one per line, e.g. The Hindu Marriage Act 1955.\n"
    "@@BOOKS  — one per line, with author, title, edition, publisher, year\n"
    "@@WEBSITES  — one URL per line (e.g. https://www.scconline.com/)\n"
    "@@ABBREVIATIONS  — one per line as: Abbreviation | Full Form  (every abbreviation used in the memorial)\n"
    "@@JURISDICTION  — paragraph(s) beginning 'The <side> has invoked the jurisdiction of this Hon'ble "
    "Court under ...'; after each provision cited, put a footnote reproducing that provision.\n"
    "@@FACTS  — numbered paragraphs '1. ...', '2. ...' stating the facts chronologically, from the "
    "side's standpoint, using only facts given.\n"
    "@@ISSUES  — for each issue, two lines: 'ISSUE 1:' then the 'Whether ...?' question on the next line.\n"
    "@@SUMMARY  — for each issue: a line 'Whether ...?' (the issue, repeated) then one paragraph "
    "beginning 'It is humbly contended/submitted that ...'; footnotes allowed.\n"
    "@@ARGUMENTS  — per issue: a line '# 1. THE ISSUE STATED AS A POSITIVE PROPOSITION IN CAPS'; then an "
    "introductory line 'It is humbly submitted before this Hon'ble Court that ... because:'; then "
    "sub-points as '(a) ...' lines; then each sub-heading as '## A. SUB-HEADING IN CAPS' followed by "
    "numbered paragraphs '1. ...', '2. ...' (numbering continues across sub-headings within an issue and "
    "restarts at 1 for the next issue). Each paragraph argues the law and applies it to the facts. "
    "Cite authorities in running text as 'In X v Y, the Court held ...'.\n"
    "@@PRAYER  — an opening line 'Wherefore in light of the issues raised, authorities cited and "
    "arguments advanced, the <side> humbly prays that this Hon'ble Court may be pleased to adjudge and "
    "declare that:' then numbered reliefs '1. ...' (one per issue, plus any consequential relief), then "
    "'And pass any other order which this Hon'ble Court may deem fit in the interest of JUSTICE, EQUITY "
    "AND GOOD CONSCIENCE.' and 'All of which is humbly prayed.'\n"
    "@@SIGNATURE  — lines: 'PLACE: <place>', 'DATE: ___/___/____', 'SD/-____________', 'COUNSEL FOR <SIDE>'\n"
    "FOOTNOTES: put every citation in a footnote written inline as {{fn: Case Name, [Year] Vol Reporter "
    "Page.}} immediately after the sentence it supports (statute provisions as {{fn: s 9, Code of Civil "
    "Procedure 1908.}}; facts as {{fn: Moot Proposition ¶ 7.}}). Every case named in a footnote or in the "
    "text MUST also be listed under @@CASES, and every statute under @@STATUTES. Use ONLY real, well-known "
    "authorities you are confident exist with correct citations; never invent a case or citation — if "
    "unsure, argue from the statutory text and principles instead. Where the problem names a fictional "
    "country/statute (e.g. 'Frisk Penal Code'), use that name consistently.\n"
    "STYLE: formal, persuasive moot language ('It is humbly submitted', 'It is most respectfully "
    "contended'); write squarely for the side specified below; address every issue with sub-headings "
    "and 3-6 substantive numbered paragraphs each. Output ONLY the tagged memorial text — no commentary."
)



MOOT_FRONT_NOTE = (
    "\n\nTHIS PASS: output ONLY these tags, in order: @@COVER, @@STATUTES, @@BOOKS, @@WEBSITES, "
    "@@ABBREVIATIONS, @@JURISDICTION, @@FACTS, @@ISSUES, @@SUMMARY, @@PRAYER, @@SIGNATURE. Do NOT output "
    "@@CASES or @@ARGUMENTS (they are produced separately). Be as full as a real competition memorial: "
    "@@STATUTES 6-10 entries; @@BOOKS 5-8 entries; @@WEBSITES 4 entries; @@ABBREVIATIONS 25-40 entries "
    "(every abbreviation used, plus standard ones like AIR, SCC, HC, SC, Art, s, v, Ors, UOI); "
    "@@JURISDICTION 1-2 paragraphs with a footnote reproducing EACH statutory provision relied on; "
    "@@FACTS 10-15 detailed numbered paragraphs; @@ISSUES 3-5 issues; @@SUMMARY one 120-180 word "
    "paragraph per issue with footnotes; @@PRAYER one relief per issue plus consequential relief."
)

DRAFT_SYSTEM_MOOT_ARGS = (
    "You are an expert moot-court drafter writing ONE issue of the ARGUMENTS ADVANCED section of a "
    "memorial, to the depth of a national-level winning memorial. Plain text only (no markdown, no "
    "asterisks). Do NOT output any @@ tag, preamble or commentary. Format exactly:\n"
    "# <n>. <THE ISSUE STATED AS A POSITIVE PROPOSITION IN CAPS>\n"
    "It is humbly submitted before this Hon'ble Court that ... because:\n"
    "(a) <ground one>\n(b) <ground two>\n(c) <ground three>\n"
    "then 3-5 sub-headings, each written as '## A. SUB-HEADING IN CAPS' (A., B., C., ...) with 5-7 "
    "numbered paragraphs under each ('1. ...', numbering continuing across sub-headings within the issue). "
    "In total write AT LEAST 20 numbered paragraphs for the issue. Each paragraph is 70-130 words: state the "
    "rule, cite and explain the authority ('In X v Y, the Hon'ble Supreme Court held that ...'), and then "
    "apply it to the facts of THIS case ('In the present case, ...'). Anticipate and rebut the opposite "
    "side's likely contentions. Put each citation in an inline footnote immediately after the sentence it "
    "supports, written {{fn: Case Name, [Year] Vol Reporter Page.}}; statutes as {{fn: s 9, Code of Civil "
    "Procedure 1908.}}; facts as {{fn: Moot Proposition ¶ 7.}}; several authorities in one footnote are "
    "separated by semicolons. Use ONLY real, well-known authorities with correct citations that you are "
    "confident about — never invent a case; if unsure, rely on the statutory text and general principles. "
    "Cite roughly 15-25 distinct authorities for the issue. Write squarely for the side given."
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


# Whether a document type needs "which side does this favour?" is decided locally for every
# type in the built-in list (instant, free, deterministic). Only free-typed document types and
# template-only drafts fall through to the AI classifier.
_ADVERSARIAL_GROUPS = {"Notices & Replies", "Civil Pleadings", "Criminal Matters (BNS, BNSS & BSA)",
                       "Writs & Constitutional", "Consumer & Labour", "Family Law"}
_NEUTRAL_GROUPS = {"Affidavits & Declarations", "Agreements & Contracts",
                   "Property & Real Estate", "Corporate & Commercial"}
_SIDE_OVERRIDES = {
    "Application under RTI Act": False,
    "Divorce Petition (Mutual Consent)": False, "Adoption Deed": False,
    "Will / Testament": False, "Declaration of Marriage": False,
    "Undertaking": False, "Indemnity Letter": False, "Authorization Letter": False,
    "Deed of Assignment": False, "Statement of Case": True,
}
_SIDE_BY_TYPE = {}
for _g in DOCUMENT_TYPES:
    for _item in _g["items"]:
        if _g["group"] in _ADVERSARIAL_GROUPS:
            _SIDE_BY_TYPE[_item] = True
        elif _g["group"] in _NEUTRAL_GROUPS:
            _SIDE_BY_TYPE[_item] = False
_SIDE_BY_TYPE.update(_SIDE_OVERRIDES)

SIDE_OTHER_VALUE = SIDE_BUTTONS[2]["value"]


def needs_side_question(conv) -> bool:
    """Is this document inherently adversarial (so a favoured side must be picked) or
    neutral/bilateral (so the question can be skipped)?"""
    known = _SIDE_BY_TYPE.get(conv["doc_type"])
    if known is not None:
        return known

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
    answer = ai_generate(prompt, temperature=0, fast=True, max_tokens=16).strip().upper()
    if answer.startswith("N"):
        return False
    return True      # YES, or anything unclear -> ask; asking is the safe default


def _ai_unavailable_note(e, retry_hint):
    msg = str(e).strip().replace("\n", " ")
    if len(msg) > 220:
        msg = msg[:220] + "…"
    return f"(AI is temporarily unavailable: {msg}) {retry_hint}"


def _enter_brainstorm(conv):
    """Shared tail used once side + facts are both known: kick off the opening
    brainstorm turn and push the assistant's reply."""
    conv["stage"] = "brainstorm"
    try:
        reply_text, quick_replies = run_brainstorm_turn(conv, opening=True)
    except Exception as e:
        reply_text, quick_replies = (
            _ai_unavailable_note(e, "You can still describe what you'd like in the draft, "
                                    "or click Generate Draft when ready."), [])
    push(conv, "assistant", reply_text, buttons=quick_replies)


def decide_side_stage(conv):
    """After facts/data are collected, decide whether asking which side the draft should
    favour is actually relevant, and either ask it or skip straight to the brainstorm."""
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
        _enter_brainstorm(conv)


def stage_ask_side(conv, text):
    if text.strip() == SIDE_OTHER_VALUE:
        # "Other — I'll specify": actually ask who, instead of storing the button label as the side.
        push(conv, "assistant",
             "Sure — who should this draft favour? Type the party's name or role "
             "(for example, \"the landlord\" or \"the complainant\").")
        return    # stay in ask_side
    conv["side"] = text.strip()
    _enter_brainstorm(conv)


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
    side_key = ("respondent"
                if re.search(r'respondent|defendant|opposite part', conv["side"], re.I)
                else "petitioner")
    try:
        conv["template_text"] = select_moot_reference(side_key, conv["details"])
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
    messages = [{"role": "system", "content": system}]
    messages.extend(conv["brainstorm"][-16:])
    if opening:
        messages.append({"role": "user", "content":
            "Kick off the brainstorm: briefly note how you'll approach this draft, and ask "
            "1-2 short questions if anything important is still missing."})
    raw = ai_chat(messages, temperature=0.6, max_tokens=BRAINSTORM_MAX_TOKENS)
    reply_text, quick_replies = _parse_brainstorm_json(raw)
    conv["brainstorm"].append({"role": "assistant", "content": reply_text})
    return reply_text, quick_replies


def _clean_quick_replies(raw_quick):
    out = []
    if isinstance(raw_quick, list):
        for item in raw_quick[:4]:
            if isinstance(item, dict) and item.get("label") and item.get("value"):
                out.append({"label": str(item["label"])[:40], "value": str(item["value"])})
    return out


def _parse_brainstorm_json(raw: str):
    """Best-effort parse of the model's structured {reply, quick_replies} JSON. Handles
    code fences, chatter around the object, and replies cut off mid-JSON by the token
    limit; falls back to the raw text (no buttons) so users never see broken JSON."""
    text = raw.strip()
    text = re.sub(r'^```(?:json)?\s*', '', text)
    text = re.sub(r'\s*```$', '', text)

    data = None
    try:
        data = json.loads(text)
    except ValueError:
        s, e = text.find('{'), text.rfind('}')
        if s != -1 and e > s:
            try:
                data = json.loads(text[s:e + 1])
            except ValueError:
                data = None

    if isinstance(data, dict):
        reply = str(data.get("reply", "")).strip()
        if reply:
            return reply, _clean_quick_replies(data.get("quick_replies"))

    m = re.search(r'"reply"\s*:\s*"((?:[^"\\]|\\.)*)', text, re.S)   # truncated JSON
    if m:
        frag = m.group(1)
        try:
            return json.loads('"' + frag + '"').strip(), []
        except ValueError:
            return frag.replace('\\n', '\n').replace('\\"', '"').strip(), []
    return raw.strip(), []


def build_draft_prompt(conv):
    """Return (system, prompt) for the final draft."""
    side_line = conv["side"] or "(none specified — draft in neutral, standard form for this document type)"
    is_moot = conv["doc_type"] == MOOT_MEMORIAL_DOC_TYPE
    details = conv["details"][:12000]
    notes = _digest(conv)

    if is_moot:
        system = DRAFT_SYSTEM_MOOT
        if conv["template_text"]:
            prompt = (
                f'Below is a STYLE & STRUCTURE reference drawn from past moot memorials for this '
                f'side — use it only to match section order, heading conventions, phrasing style '
                f'and formal tone. The required @@ tag output format in the system message ALWAYS takes precedence over the reference. Do NOT reuse any facts, party names, case citations, statutes '
                f'or numbers from it — this is a DIFFERENT case with its own facts and law.\n\n'
                f'--- STYLE & STRUCTURE REFERENCE ---\n{conv["template_text"][:6000]}\n\n'
                f'--- THE ACTUAL MOOT PROBLEM / CASE FACTS FOR THIS MEMORIAL ---\n{details}\n\n'
                f'--- SIDE THIS MEMORIAL MUST ARGUE FOR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{notes}\n\n'
                f'Now produce the complete final memorial text, following the required section '
                f'structure exactly.'
            )
        else:
            prompt = (
                f'--- THE MOOT PROBLEM / CASE FACTS FOR THIS MEMORIAL ---\n{details}\n\n'
                f'--- SIDE THIS MEMORIAL MUST ARGUE FOR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{notes}\n\n'
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
                f'--- DATA TO USE ---\n{details}\n\n'
                f'--- SIDE THIS MUST FAVOUR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{notes}\n\n'
                f'Now produce the complete final document text.'
            )
        else:
            prompt = (
                f'Draft a "{conv["doc_type"]}" document using the following details and data:\n\n'
                f'{details}\n\n'
                f'--- SIDE THIS MUST FAVOUR ---\n{side_line}\n\n'
                f'--- BRAINSTORM NOTES ---\n{notes}\n\n'
                f'Produce the complete, professional, ready-to-use document text.'
            )
    return system, prompt


def _clean_draft(text: str) -> str:
    """Strip code fences a model sometimes wraps around the document."""
    text = re.sub(r'^\s*```[a-zA-Z]*\s*\n', '', text)
    text = re.sub(r'\n\s*```\s*$', '', text)
    return text.strip()


DRAFT_READY_NOTE = ("Here's a draft based on everything we've discussed. Review it in the side "
                    "panel, and keep chatting if you'd like changes — you can regenerate any time.")



def generate_moot_memorial(conv, system, prompt, on_text, max_tokens):
    """Multi-pass generation so the memorial reaches the length of real competition memorials:
    pass 1 = front matter/facts/issues/summary/prayer; then one pass per issue for the arguments
    (each long and heavily footnoted). The Index of Authorities is derived from the footnotes."""
    front = ai_generate(prompt, system=system + MOOT_FRONT_NOTE, temperature=0.4,
                        max_tokens=max_tokens, on_text=on_text, max_continuations=3)
    front = _clean_draft(front)
    issues = []
    m = re.search(r'@@\s*ISSUES\s*:?\s*\n(.*?)(?=\n\s*@@|\Z)', front, re.S)
    if m:
        lines = [l.strip() for l in m.group(1).splitlines() if l.strip()]
        for i, l in enumerate(lines):
            if re.match(r'^ISSUE\s*[0-9IVX]+', l, re.I):
                rest = re.sub(r'^ISSUE\s*[0-9IVX]+\s*[:.\-]?\s*', '', l, flags=re.I)
                if not rest and i + 1 < len(lines):
                    rest = lines[i + 1]
                issues.append(rest)
    issues = [x for x in issues if x][:6] or ["the principal issue in the moot problem"]

    details = conv["details"][:9000]
    side = conv["side"] or "the side specified"
    notes = _digest(conv, 1500)
    facts_ctx = ""
    fm = re.search(r'@@\s*FACTS\s*:?\s*\n(.*?)(?=\n\s*@@|\Z)', front, re.S)
    if fm:
        facts_ctx = fm.group(1).strip()[:3500]
    parts = ["@@ARGUMENTS"]
    for n, issue in enumerate(issues, 1):
        others = "\n".join(f"Issue {k}: {t}" for k, t in enumerate(issues, 1))
        p = (f"MOOT PROBLEM:\n{details}\n\nSTATEMENT OF FACTS (as drafted):\n{facts_ctx}\n\n"
             f"ALL ISSUES:\n{others}\n\nSIDE TO ARGUE FOR: {side}\n\nBRAINSTORM NOTES:\n{notes}\n\n"
             f"Now write the full arguments for ISSUE {n} ONLY: {issue}\n"
             f"Start with the line '# {n}. ' followed by the issue as a positive proposition in caps.")
        base = front + "\n" + "\n".join(parts) + "\n"
        cb = (lambda t, _b=base: on_text(_b + t)) if on_text else None
        out = ai_generate(p, system=DRAFT_SYSTEM_MOOT_ARGS, temperature=0.4, max_tokens=max_tokens,
                          on_text=cb, max_continuations=3)
        out = _clean_draft(out)
        out = re.sub(r'(?m)^\s*@@.*$', '', out)      # strip any stray tags
        parts.append(out.strip())
        if on_text:
            on_text(front + "\n" + "\n".join(parts) + "\n")
    return front + "\n" + "\n".join(parts) + "\n"


def _run_generation_job(conv, job, system, prompt):
    """Background worker: streams the draft into job['text'] so the UI can show it live,
    then builds the .docx. Runs in a thread so no HTTP request is held open for the
    whole generation (Render's proxy would cut a long request)."""
    try:
        def on_text(t):
            job["text"] = t

        is_moot = (conv.get("doc_type") == MOOT_MEMORIAL_DOC_TYPE)
        if is_moot:
            text = generate_moot_memorial(conv, system, prompt, on_text, MOOT_MAX_TOKENS)
        else:
            text = ai_generate(prompt, system=system, temperature=0.4,
                               max_tokens=DRAFT_MAX_TOKENS, on_text=on_text, max_continuations=2)
        text = _clean_draft(text)
        job["text"] = text

        try:
            path = None
            if is_moot:
                try:
                    path = build_moot_docx(conv["doc_type"], text)
                except Exception:
                    traceback.print_exc()
            if not path:
                path = build_ai_legal_docx(conv["doc_type"] or "Legal_Draft", text)
        except Exception:
            traceback.print_exc()
            path = ""

        with conv["lock"]:
            old = conv.get("docx_path")
            conv["draft_text"] = text
            conv["docx_path"] = path
            if old and old != path:
                _remove_file(old)
            push(conv, "assistant", DRAFT_READY_NOTE if path else
                 DRAFT_READY_NOTE + " (The Word export failed this time — regenerate to retry.)")
        job["status"] = "done"
    except Exception as e:
        traceback.print_exc()
        job["error"] = str(e)[:400] or "Draft generation failed."
        job["status"] = "error"


def _digest(conv, limit_chars=3000):
    parts = []
    for m in conv["brainstorm"][-16:]:
        parts.append(f'{m["role"].upper()}: {m["content"]}')
    text = "\n".join(parts)
    return text[-limit_chars:] if text else "(no additional notes)"


def _digest(conv, limit_chars=3000):
    parts = []
    for m in conv["brainstorm"][-16:]:
        parts.append(f'{m["role"].upper()}: {m["content"]}')
    text = "\n".join(parts)
    return text[-limit_chars:] if text else "(no additional notes)"


# ══════════════════════════════════════════════════════════════════[...]
#  ROUTES
# ══════════════════════════════════════════════════════════════════[...]

@app.errorhandler(413)
def _too_large(e):
    return jsonify({"success": False, "message": "That upload is too large (max 15 MB)."}), 413


@app.errorhandler(Exception)
def _handle_error(e):
    """Always answer the JS client with JSON — a bare HTML 500 page used to leave the UI
    stuck on 'Dratido is thinking…'."""
    if isinstance(e, HTTPException):
        return jsonify({"success": False, "message": e.description or e.name}), e.code
    traceback.print_exc()
    return jsonify({"success": False,
                    "message": "Something went wrong on the server. Please try again."}), 500


@app.route('/healthz')
def healthz():
    return "ok", 200


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
    if len(text) > MAX_MESSAGE_CHARS:
        return jsonify({"success": False,
                        "message": f"That message is too long (max {MAX_MESSAGE_CHARS:,} characters)."}), 400

    with conv["lock"]:      # one request at a time per conversation (double-clicks, retries)
        handler = STAGE_HANDLERS.get(conv["stage"])
        if not handler:
            return jsonify({"success": False, "message": "Unknown stage."}), 400
        push(conv, "user", text)
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

    with conv["lock"]:
        if conv["stage"] != "ask_template":
            return jsonify({"success": False, "message": "Not expecting a template upload right now."}), 400

        try:
            extracted = (extract_text_from_docx(f.stream) if ext == '.docx'
                         else extract_text_from_pdf(f.stream))
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


@app.route('/api/generate', methods=['POST'])
def api_generate():
    """Starts draft generation in the background and returns immediately.
    The client then polls /api/generate_status/<conv_id> and sees the draft stream in."""
    data = request.get_json(silent=True) or {}
    conv = get_conversation(data.get('conv_id', ''))
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found. Start a new draft."}), 404
    if not os.environ.get('GROQ_API_KEY', '').strip():
        return jsonify({"success": False,
                        "message": "GROQ_API_KEY not set. Get a free key at https://console.groq.com"}), 400

    with conv["lock"]:
        if conv["stage"] != "brainstorm":
            return jsonify({"success": False,
                            "message": "Finish the setup questions before generating a draft."}), 400
        job = conv.get("job")
        if job and job["status"] == "running":
            return jsonify({"success": True, "status": "running"})
        system, prompt = build_draft_prompt(conv)
        job = {"id": uuid.uuid4().hex, "status": "running", "text": "", "error": "",
               "started": time.time()}
        conv["job"] = job

    threading.Thread(target=_run_generation_job, args=(conv, job, system, prompt),
                     daemon=True).start()
    return jsonify({"success": True, "status": "running"})


@app.route('/api/generate_status/<conv_id>')
def api_generate_status(conv_id):
    conv = get_conversation(conv_id)
    if not conv:
        return jsonify({"success": False, "message": "Conversation not found."}), 404
    job = conv.get("job")
    if not job:
        return jsonify({"success": False, "message": "No draft is being generated."}), 400

    have = request.args.get('have', default=-1, type=int)
    status = job["status"]
    text = job["text"]
    resp = {"success": True, "status": status}
    if len(text) != have:                 # only resend the text when it has changed
        resp["draft_text"] = text
    if status == "done":
        resp["messages"] = conv["messages"]
        resp["has_docx"] = bool(conv.get("docx_path"))
    elif status == "error":
        resp["message"] = job["error"] or "Draft generation failed."
    return jsonify(resp)


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


# ════════════════════��═════════════════════════════════════════════[...]
#  FRONTEND (single-page chat app)
# ══════════════════════════════════════════════════════════════════[...]

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

let busy = false;
function setBusy(b, label){
  busy = b;
  typingEl.textContent = label || 'Dratido is thinking…';
  typingEl.style.display = b ? 'block' : 'none';
  sendBtn.disabled = b;
  generateBtn.disabled = b;
  if (b) scrollBottom();
}

// fetch wrapper: never throws, always resolves to an object with .success / .message.
// (Previously a network error or an HTML 500 page left the UI stuck on "thinking…".)
async function api(url, opts){
  let res;
  try { res = await fetch(url, opts); }
  catch (e){ return {success:false, message:'Network error — please check your connection and try again.'}; }
  let data = null;
  try { data = await res.json(); } catch (e){}
  if (!data) data = {success:false, message:'Unexpected server response (HTTP ' + res.status + '). Please try again.'};
  data.http = res.status;
  return data;
}

function failed(data){
  if (data.http === 404){
    // The server forgot this conversation (restart / idle spin-down on the free tier).
    alert('This session has expired because the server restarted or was idle. Starting a fresh draft.');
    startConversation();
  } else {
    alert(data.message || 'Something went wrong.');
  }
}

async function startConversation(){
  setBusy(true);
  const data = await api('/api/start', {method:'POST'});
  setBusy(false);
  if (data.success){
    convId = data.conv_id;
    sessionStorage.setItem('dratido_conv_id', convId);
    canGenerate = false; hasDraft = false;
    generateBtn.style.display = 'none';
    panelFooter.style.display = 'none';
    panelBody.innerHTML = '<div class="placeholder">Your draft will appear here once we\'ve brainstormed enough to generate it.</div>';
    renderMessages(data.messages);
  } else {
    alert(data.message || 'Could not start a new draft. Please reload the page.');
  }
}

async function sendMessage(text){
  if (busy || !text || !text.trim() || !convId) return;
  const typed = inputEl.value;
  inputEl.value = '';
  autoGrow();
  setBusy(true);
  const data = await api('/api/message', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conv_id: convId, text: text})
  });
  setBusy(false);
  if (data.success){
    renderMessages(data.messages);
    canGenerate = !!data.can_generate;
    generateBtn.style.display = canGenerate ? 'inline-block' : 'none';
  } else {
    if (!inputEl.value) { inputEl.value = typed; autoGrow(); }   // don't lose what they typed
    failed(data);
  }
}

async function generateDraft(){
  if (busy || !convId) return;
  const myConv = convId;
  setBusy(true, 'Drafting your document…');
  const start = await api('/api/generate', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conv_id: myConv})
  });
  if (!start.success){ setBusy(false); failed(start); return; }

  hasDraft = false;
  showDraft('Starting…', false);
  openPanel();

  let have = -1, current = '', errors = 0;
  while (convId === myConv){
    await new Promise(r => setTimeout(r, 600));
    const s = await api('/api/generate_status/' + myConv + '?have=' + have);
    if (convId !== myConv) return;                 // user started a new draft meanwhile
    if (!s.success){
      if (s.http === 404 || ++errors > 5){ setBusy(false); failed(s); return; }
      continue;
    }
    errors = 0;
    if (typeof s.draft_text === 'string'){
      have = s.draft_text.length;
      current = s.draft_text;
      if (s.status === 'running') showDraft(current || 'Starting…', false);
    }
    if (s.status === 'done'){
      setBusy(false);
      renderMessages(s.messages);
      showDraft(current, true);
      return;
    }
    if (s.status === 'error'){
      setBusy(false);
      panelBody.innerHTML = '<div class="placeholder">The draft could not be generated. Close this panel and try again.</div>';
      panelFooter.style.display = 'none';
      alert(s.message || 'Could not generate the draft.');
      return;
    }
  }
}

// While streaming (final=false) the text grows in place; once final the Download button appears.
function showDraft(text, final){
  let pre = document.getElementById('draft-text');
  if (!pre){
    panelBody.innerHTML = '';
    pre = document.createElement('div');
    pre.id = 'draft-text';
    panelBody.appendChild(pre);
  }
  pre.textContent = text;
  if (final){
    hasDraft = true;
    panelFooter.style.display = 'block';
  } else {
    panelFooter.style.display = 'none';
    panelBody.scrollTop = panelBody.scrollHeight;
  }
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


# ══════════════════════════════════════════════════════════[...]
#  ENTRY POINT
# ══════════════════════════════════════════════════════════[...]

def _warmup():
    """Do the one-time slow work (template decrypt, Groq model discovery, TLS handshake)
    in the background at boot so the first real request doesn't pay for it."""
    try:
        load_moot_templates()
        key = os.environ.get('GROQ_API_KEY', '').strip()
        if key:
            get_groq_models(key)
    except Exception as e:
        print(f"[Warmup] {e}")


threading.Thread(target=_warmup, daemon=True).start()


if __name__ == '__main__':
    os.makedirs(GENERATED_DIR, exist_ok=True)

    groq_key = os.environ.get('GROQ_API_KEY', '').strip()
    key_str = '✓ Groq — ready!' if groq_key else '✗ NOT SET — see below'
    print('\n' + '=' * 60)
    print(f'  {APP_NAME} — {APP_TAGLINE}')
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
