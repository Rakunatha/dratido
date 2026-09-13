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
import html as _html_lib
from urllib.parse import urlparse, parse_qs, unquote
import xml.sax.saxutils as _sax
from flask import Flask, request, jsonify, send_file, Response
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

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


def ai_chat(messages: list, temperature: float = 0.6) -> str:
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
            "max_completion_tokens": 4096,
            "stream":      False,
        }

        for attempt in range(3):
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
                wait = 2 ** (attempt + 2)
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
#  WEB SEARCH  (find candidate reference templates, no API key required)
# ═══════════════════════════════════════════════════════════════════════════════

_SEARCH_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

_RESULT_LINK_RE = re.compile(
    r'class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', re.S
)
_TAG_RE = re.compile(r'<[^>]+>')


def _clean_html_text(fragment: str) -> str:
    text = _TAG_RE.sub('', fragment)
    return _html_lib.unescape(text).strip()


def web_search_for_templates(query: str, max_results: int = 5):
    """Search the web (DuckDuckGo's no-JS HTML endpoint — no API key needed) for candidate
    reference-document templates. Returns a list of {title, url, domain} dicts, best-effort;
    returns an empty list on any failure rather than raising, so callers can fall back
    gracefully to manual template entry."""
    import requests as _req

    search_query = f"{query} template format sample"
    try:
        resp = _req.get(
            "https://html.duckduckgo.com/html/",
            params={"q": search_query},
            headers={"User-Agent": _SEARCH_USER_AGENT},
            timeout=15,
        )
        resp.raise_for_status()
    except Exception as e:
        print(f"[WebSearch] Search request failed: {e}")
        return []

    results = []
    seen_urls = set()
    for match in _RESULT_LINK_RE.finditer(resp.text):
        raw_href, raw_title = match.group(1), match.group(2)
        real_url = raw_href
        if 'uddg=' in raw_href:
            qs = parse_qs(urlparse(raw_href).query)
            if 'uddg' in qs:
                real_url = unquote(qs['uddg'][0])
        if not real_url.startswith('http') or real_url in seen_urls:
            continue
        title = _clean_html_text(raw_title) or real_url
        domain = urlparse(real_url).netloc.replace('www.', '')
        results.append({"title": title, "url": real_url, "domain": domain})
        seen_urls.add(real_url)
        if len(results) >= max_results:
            break
    return results


def fetch_page_text(url: str, max_chars: int = 6000) -> str:
    """Best-effort extraction of readable text from a web page, for use as reference-template
    source material. Strips scripts/styles/tags and collapses whitespace."""
    import requests as _req

    resp = _req.get(url, headers={"User-Agent": _SEARCH_USER_AGENT}, timeout=15)
    resp.raise_for_status()
    body = resp.text
    body = re.sub(r'<(script|style|noscript)[^>]*>.*?</\1>', ' ', body, flags=re.S | re.I)
    text = _clean_html_text(body)
    text = re.sub(r'\s+', ' ', text).strip()
    return text[:max_chars]


# ═══════════════════════════════════════════════════════════════════════════════
#  DOCX BUILDING
# ═══════════════════════════════════════════════════════════════════════════════

DRATIDO_WATERMARK_TEXT = 'Dratido - Draft Till Done'


def add_watermark(doc, text: str = DRATIDO_WATERMARK_TEXT):
    """Insert a diagonal, semi-transparent watermark into the header of every
    section (the classic Word VML watermark technique), plus a small text
    credit line in the footer as a reliable fallback for viewers that don't
    render VML shapes."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _ALIGN

    safe_text = _sax.escape(text, {'"': '&quot;'})

    watermark_xml = (
        '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office">'
        '<v:shapetype id="_x0000_t136" coordsize="1600,21600" o:spt="136" adj="10800" '
        'path="m@7,0l@8,0m@5,21600l@6,21600e">'
        '<v:formulas>'
        '<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 #0"/>'
        '<v:f eqn="sum 0 0 #1"/><v:f eqn="prod #1 2 1"/><v:f eqn="sum 21600 0 #1"/>'
        '<v:f eqn="if #0 #3 0"/><v:f eqn="if #0 21600 #1"/><v:f eqn="if #3 21600 #2"/>'
        '<v:f eqn="if #3 #1 21600"/><v:f eqn="mid #4 #5"/><v:f eqn="mid #6 #7"/><v:f eqn="val #0"/>'
        '</v:formulas>'
        '<v:path textpathok="t" o:connecttype="custom" '
        'o:connectlocs="@9,0;@10,10800;@9,21600;@8,10800" o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/>'
        '</v:shapetype>'
        '<v:shape id="DratidoWatermark" o:spid="_x0000_s2049" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:520pt;height:110pt;'
        'rotation:315;z-index:-251654144;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:margin;mso-position-vertical:center;'
        'mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="#8B1E2D" stroked="f">'
        '<v:fill opacity=".18"/>'
        f'<v:textpath style="font-family:\'Calibri\';font-size:1pt" string="{safe_text}"/>'
        '</v:shape>'
        '</w:pict>'
    )

    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        h_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        h_para.text = ''
        h_para.alignment = _ALIGN.CENTER
        run = h_para.add_run()
        r_el = run._r
        pict = parse_xml(watermark_xml)
        r_el.append(pict)

        footer = section.footer
        footer.is_linked_to_previous = False
        f_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        f_para.text = ''
        f_para.alignment = _ALIGN.CENTER
        f_run = f_para.add_run(text)
        f_run.font.size = Pt(8)
        f_run.font.color.rgb = RGBColor(0xA0, 0xA0, 0xA0)
        f_run.italic = True


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
#   start                 -> choose "type", "template", or "web_template"
#   ask_type              -> waiting for the document type (via searchable list/modal)
#   ask_facts             -> waiting for facts/details, entered via popup (type path)
#   ask_template          -> waiting for the pasted template text, entered via popup
#   ask_web_query         -> waiting for a description of the template to search the web for
#   ask_web_pick          -> waiting for the user to pick one of the web search results
#                            (or opt to paste their own instead)
#   ask_template_details  -> waiting for facts/details, entered via popup (template path,
#                            reached from either the paste flow or the web-search flow)
#   ask_side              -> waiting for which side the draft favours (only asked when
#                            the AI pipeline determines the document type is adversarial)
#   brainstorm            -> free-form chat with the AI; draft can be generated
#                            at any point from here on

START_BUTTONS = [
    {"label": "Specify Document Type",
     "value": "I'd like to specify the type of document and enter the details myself."},
    {"label": "Use a Reference Template",
     "value": "I'd like to provide a reference template and enter the data to fill into it."},
    {"label": "Search the Web for a Template",
     "value": "I'd like to search the web for a reference template."},
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
        "mode": None,          # "type" | "template" | "web_template"
        "doc_type": "",
        "template_text": "",
        "template_source": "",  # human-readable note on where the template came from
        "details": "",
        "side": "",
        "messages": [],        # full transcript, for display
        "brainstorm": [],      # {role, content} sent to the LLM during brainstorming
        "draft_text": "",
        "docx_path": "",
        "web_query": "",
        "web_results": {},     # url -> {title, url, domain}, for the "search the web" flow
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


TEMPLATE_SWITCH_VALUE = "I'd like to provide a reference template and enter the data to fill into it."
WEB_SEARCH_SWITCH_VALUE = "I'd like to search the web for a reference template."
WEB_PICK_PASTE_OWN = "__paste_own__"
WEB_PICK_SEARCH_AGAIN = "__search_again__"


def start_template_mode(conv):
    conv["mode"] = "template"
    conv["stage"] = "ask_template"
    push(conv, "assistant",
         "Understood — you'd like to work from a reference template. Click below to paste "
         "it in (placeholders like [NAME], [DATE], etc. are fine).",
         modal={"title": "Reference Template",
                "placeholder": "Paste your template text here...",
                "submit_label": "Save Template"})


def start_web_search_mode(conv):
    conv["mode"] = "web_template"
    conv["stage"] = "ask_web_query"
    push(conv, "assistant",
         "Sure — tell me what kind of template you're looking for (e.g. \"rental agreement "
         "format India\", \"legal notice for cheque bounce\"), and I'll search the web for a "
         "few options you can pick from.",
         modal={"title": "Search the Web for a Template",
                "placeholder": "Describe the document/template you're looking for...",
                "submit_label": "Search"})


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
    if 'search' in lower and 'web' in lower:
        start_web_search_mode(conv)
    elif 'template' in lower:
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
    if stripped == WEB_SEARCH_SWITCH_VALUE:
        start_web_search_mode(conv)
        return
    conv["doc_type"] = stripped
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


def stage_ask_web_query(conv, text):
    conv["web_query"] = text.strip()
    try:
        results = web_search_for_templates(conv["web_query"], max_results=5)
    except Exception as e:
        print(f"[WebSearch] Unexpected error: {e}")
        results = []

    if not results:
        conv["web_results"] = {}
        push(conv, "assistant",
             "I couldn't find any usable results for that search. You can try describing it "
             "differently, or paste your own reference template instead.",
             buttons=[{"label": "🔍 Try a Different Search", "value": WEB_PICK_SEARCH_AGAIN},
                      {"label": "✎ Paste My Own Template", "value": WEB_PICK_PASTE_OWN}])
        conv["stage"] = "ask_web_pick"
        return

    conv["web_results"] = {r["url"]: r for r in results}
    conv["stage"] = "ask_web_pick"
    buttons = []
    for r in results:
        label = r["title"][:55] + ("…" if len(r["title"]) > 55 else "")
        buttons.append({"label": f'{label} — {r["domain"]}', "value": r["url"]})
    buttons.append({"label": "🔍 Search Again", "value": WEB_PICK_SEARCH_AGAIN})
    buttons.append({"label": "✎ Paste My Own Instead", "value": WEB_PICK_PASTE_OWN})
    push(conv, "assistant",
         "Here's what I found — pick one to use as the reference template, search again, or "
         "paste your own.", buttons=buttons)


def stage_ask_web_pick(conv, text):
    choice = text.strip()

    if choice == WEB_PICK_PASTE_OWN:
        start_template_mode(conv)
        return

    if choice == WEB_PICK_SEARCH_AGAIN:
        conv["web_results"] = {}
        push(conv, "assistant",
             "No problem — what should I search for instead?",
             modal={"title": "Search the Web for a Template",
                    "placeholder": "Describe the document/template you're looking for...",
                    "submit_label": "Search"})
        conv["stage"] = "ask_web_query"
        return

    result = conv.get("web_results", {}).get(choice)
    if not result:
        # Unrecognised input (e.g. stale button click) — fall back to a fresh search prompt.
        push(conv, "assistant",
             "I didn't recognise that choice. What should I search for?",
             modal={"title": "Search the Web for a Template",
                    "placeholder": "Describe the document/template you're looking for...",
                    "submit_label": "Search"})
        conv["stage"] = "ask_web_query"
        return

    try:
        page_text = fetch_page_text(result["url"])
    except Exception as e:
        print(f"[WebSearch] Fetch failed for {result['url']}: {e}")
        page_text = ""

    if not page_text:
        push(conv, "assistant",
             f"I couldn't retrieve usable content from \"{result['title']}\" — that page may "
             f"block automated access. Pick another result, search again, or paste your own "
             f"template.",
             buttons=[{"label": lbl["label"], "value": lbl["value"]}
                      for lbl in _rebuild_web_pick_buttons(conv)])
        return

    conv["template_text"] = page_text
    conv["template_source"] = f'{result["title"]} ({result["domain"]})'
    conv["mode"] = "template"
    ask_for_template_details(
        conv,
        f'Got reference material from "{result["title"]}" ({result["domain"]}). Click below '
        f'to enter the data to fill into it (names, dates, amounts, and any other specifics).'
    )


def _rebuild_web_pick_buttons(conv):
    buttons = []
    for r in conv.get("web_results", {}).values():
        label = r["title"][:55] + ("…" if len(r["title"]) > 55 else "")
        buttons.append({"label": f'{label} — {r["domain"]}', "value": r["url"]})
    buttons.append({"label": "🔍 Search Again", "value": WEB_PICK_SEARCH_AGAIN})
    buttons.append({"label": "✎ Paste My Own Instead", "value": WEB_PICK_PASTE_OWN})
    return buttons


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
    "ask_web_query":         stage_ask_web_query,
    "ask_web_pick":          stage_ask_web_pick,
    "ask_template_details":  stage_ask_template_details,
    "ask_side":              stage_ask_side,
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
    draft_text = ai_generate(prompt, system=DRAFT_SYSTEM, temperature=0.4)
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

    try:
        draft_text = generate_draft(conv)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

    note = "Here's a draft based on everything we've discussed. Review it in the side panel, and keep chatting if you'd like changes — you can regenerate any time."
    push(conv, "assistant", note)

    return jsonify({
        "success": True,
        "draft_text": draft_text,
        "messages": conv["messages"],
    })


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
      <button class="modal-template-switch-btn" id="modal-web-search-switch">🔍 Search the Web for a Template Instead</button>
      <div class="modal-actions">
        <button class="btn" id="modal-list-cancel">Cancel</button>
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
const modalHintEl = document.getElementById('modal-hint');
const modalSearchEl = document.getElementById('modal-search');
const modalListResultsEl = document.getElementById('modal-list-results');
const TEMPLATE_SWITCH_VALUE = "I'd like to provide a reference template and enter the data to fill into it.";
const WEB_SEARCH_SWITCH_VALUE = "I'd like to search the web for a reference template.";
let currentModalGroups = [];

function openModal(cfg){
  document.getElementById('modal-title').textContent = cfg.title || 'Enter Details';

  if (cfg.type === 'list'){
    modalHintEl.textContent = 'Search or scroll to find your document type.';
    textModeEl.style.display = 'none';
    listModeEl.style.display = 'block';
    currentModalGroups = cfg.groups || [];
    modalSearchEl.value = '';
    renderModalList('');
    document.getElementById('modal-overlay').classList.add('open');
    setTimeout(() => modalSearchEl.focus(), 50);
  } else {
    modalHintEl.textContent = "This opens in its own window so you can enter everything comfortably before it's added to the chat.";
    textModeEl.style.display = 'block';
    listModeEl.style.display = 'none';
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
document.getElementById('modal-web-search-switch').onclick = () => {
  closeModal();
  sendMessage(WEB_SEARCH_SWITCH_VALUE);
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
  const res = await fetch('/api/generate', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conv_id: convId})
  });
  const data = await res.json();
  setBusy(false);
  if (data.success){
    renderMessages(data.messages);
    showDraft(data.draft_text);
    openPanel();
  } else {
    alert(data.message || 'Could not generate the draft.');
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
