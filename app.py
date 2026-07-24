"""
Auto PPT Maker — Web Server with Multi-User Auth
"""

import os, sys, json, re, threading, uuid, tempfile, base64, io
from pathlib import Path
from functools import wraps
from flask import Flask, request, jsonify, send_file, session, redirect, url_for, render_template_string
from werkzeug.security import generate_password_hash, check_password_hash

SCRIPT_DIR = Path(__file__).parent
USERS_FILE = SCRIPT_DIR / "users.json"
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-production")
JOBS: dict = {}
IMG_CACHE: dict = {}  # keyword → base64 — avoids re-downloading same images

TEXT_LIMIT = 12_000   # chars to extract from any file — enough for full structure
PAGE_LIMIT = 20       # PDF pages — content beyond this rarely adds new concepts

# ── User storage ──────────────────────────────────────────────────────────────
def load_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text())
        except: pass
    return {}

def save_users(u):
    USERS_FILE.write_text(json.dumps(u, indent=2))

def init_admin():
    """Always sync admin password from env var on every startup."""
    pw = os.environ.get("ADMIN_PASSWORD","").strip()
    if not pw: return
    u = load_users()
    u["admin"] = {"password": generate_password_hash(pw), "name": "Administrator", "role": "admin"}
    save_users(u)
    print("[startup] Admin account ready.")

init_admin()

# ── Auth helpers ──────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def d(*a,**k):
        if "username" not in session: return redirect(url_for("login_page"))
        return f(*a,**k)
    return d

def admin_required(f):
    @wraps(f)
    def d(*a,**k):
        if "username" not in session: return redirect(url_for("login_page"))
        if load_users().get(session["username"],{}).get("role") != "admin": return "Access denied",403
        return f(*a,**k)
    return d

def extract_text(path):
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        # ── Try PyMuPDF first (10× faster than pypdf for large files) ──
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(str(path))
            parts, total = [], 0
            for i in range(min(len(doc), PAGE_LIMIT)):
                t = doc[i].get_text("text")
                parts.append(t); total += len(t)
                if total >= TEXT_LIMIT: break
            doc.close()
            return "\n".join(parts)[:TEXT_LIMIT]
        except ImportError:
            pass
        # ── Fallback: pypdf, still capped ──
        from pypdf import PdfReader
        parts, total = [], 0
        for page in list(PdfReader(path).pages)[:PAGE_LIMIT]:
            t = page.extract_text() or ""
            parts.append(t); total += len(t)
            if total >= TEXT_LIMIT: break
        return "\n".join(parts)[:TEXT_LIMIT]
    elif ext == ".docx":
        import docx
        return "\n".join(p.text for p in docx.Document(path).paragraphs)[:TEXT_LIMIT]
    return Path(path).read_text(encoding="utf-8", errors="ignore")[:TEXT_LIMIT]

SLIDE_COUNTS = {"short": "10", "medium": "10-20", "long": "20-25"}
STYLE_HINTS = {
    "summary":     "Use very brief bullet points (max 8 words each). Focus on key facts only. Max 3 bullets per slide.",
    "descriptive": "Use detailed bullets (up to 15 words each). Include explanations and context. Up to 5 bullets per slide.",
    "visual":      "Use minimal text (2-3 short bullets per slide). Always provide a strong image_keyword. Include a key_fact on every slide.",
}

def structure(text, topic, subject, grade, log, length="medium", style="descriptive"):
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        log("📝 No API key — using smart auto-structuring (works great!)")
        return smart_structure(text, topic, subject, grade, length, style)
    client = anthropic.Anthropic(api_key=key)
    trunc = text[:6000] + ("..." if len(text) > 6000 else "")
    n_slides = SLIDE_COUNTS.get(length, "10-20")
    style_hint = STYLE_HINTS.get(style, STYLE_HINTS["descriptive"])
    prompt = f"""You are a teaching assistant creating PowerPoint slides from lesson notes.
INPUT TEXT:\n{trunc}
Topic: {topic or "auto-detect"} | Subject: {subject or "auto-detect"} | Grade: {grade or "general"}
STYLE INSTRUCTIONS: {style_hint}
Return ONLY valid JSON (no markdown) with this exact structure:
{{"topic":"Main topic","subtitle":"One-line description","subject":"Subject","grade":"Grade",
"objectives":["Up to 4 objectives starting with a verb"],
"slides":[{{"title":"Slide title (max 8 words)","content":["Bullets per style above"],"key_term":"Term: definition (optional)","key_fact":"Striking stat or quote (optional)","image_keyword":"2-3 word image search term"}}],
"activity":"Short engaging class activity","summary":["Up to 5 key takeaways"]}}
Create exactly {n_slides} content slides. Follow the style instructions strictly."""
    try:
        r = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=6000,
                                   messages=[{"role":"user","content":prompt}])
        raw = re.sub(r'^```(?:json)?\s*','',r.content[0].text.strip(),flags=re.MULTILINE)
        raw = re.sub(r'\s*```\s*$','',raw,flags=re.MULTILINE)
        return json.loads(raw)
    except Exception as e:
        log(f"⚠️  AI error: {e} — falling back to smart structuring")
        return smart_structure(text, topic, subject, grade, length, style)

# ── Smart free-mode structuring (no API key needed) ───────────────────────────
def smart_structure(text, topic, subject, grade, length="medium", style="descriptive"):
    import re as _re

    # ── 0. Pre-truncate for speed — 10 k chars is plenty for structure ────────
    text = text[:10_000]

    # ── 1. Clean & split into sentences / lines ───────────────────────────────
    lines = [l.strip() for l in text.split("\n") if l.strip()][:200]
    all_words = text.lower().split()[:3_000]   # cap word-freq input
    word_freq = {}
    STOP = {"the","a","an","is","are","was","were","be","been","being","have","has","had",
            "do","does","did","will","would","could","should","may","might","must","shall",
            "to","of","in","on","at","by","for","with","from","and","or","but","not","this",
            "that","it","its","their","they","them","he","she","we","you","i","as","also",
            "which","who","what","how","when","where","why","if","then","than","so","all",
            "can","our","your","his","her","into","about","up","out","more","some","such",
            "each","between","after","before","through","during","these","those","any","no"}
    for w in all_words:
        w = _re.sub(r'[^a-z]','',w)
        if len(w)>3 and w not in STOP:
            word_freq[w] = word_freq.get(w,0)+1

    # Top keywords = potential image keywords
    top_kw = sorted(word_freq, key=lambda w:-word_freq[w])[:20]

    # ── 2. Auto-detect topic ──────────────────────────────────────────────────
    detected_topic = topic
    if not detected_topic:
        for ln in lines[:8]:
            if 4 < len(ln.split()) < 12 and not ln.endswith('.'):
                detected_topic = ln.rstrip(':').strip()
                break
        if not detected_topic:
            detected_topic = " ".join(w.capitalize() for w in top_kw[:3]) if top_kw else "Lesson"

    # ── 3. Detect section headings ────────────────────────────────────────────
    def is_heading(ln):
        words = ln.split()
        return (len(words) <= 8 and not ln.endswith('.') and not ln.endswith(',') and
                (ln.isupper() or ln.istitle() or ln.endswith(':') or
                 _re.match(r'^\d+[\.\)]\s', ln) or
                 _re.match(r'^[A-Z][a-zA-Z\s]{3,40}:?\s*$', ln)))

    def score_sentence(sent):
        """Score a sentence by keyword density and length."""
        words = _re.sub(r'[^a-z ]', '', sent.lower()).split()
        if len(words) < 4: return 0
        kw_hits = sum(1 for w in words if w in word_freq and word_freq[w] > 1)
        length_bonus = min(len(words), 20) / 20
        return kw_hits * length_bonus

    def clean_bullet(s):
        s = s.strip().rstrip('.')
        # Remove leading list markers
        s = _re.sub(r'^[\-\*•\d]+[\.\)]\s*', '', s)
        return s[:120]

    def img_kw(title):
        words = [w.lower() for w in title.split() if w.lower() not in STOP and len(w)>3]
        kws = words[:3] or top_kw[:3]
        return " ".join(kws) if kws else (detected_topic or "education")

    # ── 4. Group lines into sections ──────────────────────────────────────────
    sections = []
    cur_title, cur_lines = "Introduction", []
    for ln in lines:
        if is_heading(ln) and cur_lines:
            sections.append((cur_title.rstrip(':'), cur_lines[:]))
            cur_title, cur_lines = ln.rstrip(':'), []
        elif is_heading(ln):
            cur_title = ln.rstrip(':')
        else:
            cur_lines.append(ln)
    if cur_lines:
        sections.append((cur_title.rstrip(':'), cur_lines))

    # If no sections detected, chunk into groups of ~5 lines
    if len(sections) <= 1:
        content_lines = [l for l in lines if not is_heading(l) and len(l) > 20]
        chunk = max(3, len(content_lines) // 8)
        sections = []
        for i in range(0, min(len(content_lines), 200), chunk):
            title = " ".join(content_lines[i].split()[:6]).rstrip(',.:')
            sections.append((title, content_lines[i:i+chunk]))

    # ── 5. Target slide count based on length ─────────────────────────────────
    target = {"short": 10, "medium": 15, "long": 22}.get(length, 15)

    # ── 6. Build slides ───────────────────────────────────────────────────────
    bullets_per_slide = {"summary": 3, "descriptive": 5, "visual": 2}.get(style, 4)

    slides = []
    for sec_title, sec_lines in sections:
        # Score and pick best sentences
        scored = sorted([(score_sentence(l), l) for l in sec_lines if len(l.split())>4], reverse=True)
        best = [clean_bullet(l) for _,l in scored[:bullets_per_slide] if l]
        if not best:
            best = [clean_bullet(l) for l in sec_lines[:bullets_per_slide] if len(l)>10]
        if not best:
            continue

        # Key term: look for "X: definition" or "X is a..." patterns
        key_term = ""
        for ln in sec_lines:
            m = _re.match(r'^([A-Z][a-zA-Z\s]{2,25}):\s+(.{10,80})', ln)
            if m:
                key_term = f"{m.group(1)}: {m.group(2)[:60]}"; break
            m2 = _re.match(r'^([A-Z][a-z]+(?:\s[a-z]+)?)\s+is\s+(?:a|an|the)?\s*(.{10,60})', ln)
            if m2:
                key_term = f"{m2.group(1)}: {m2.group(2)[:60]}"; break

        # Key fact: longest sentence as a notable point
        key_fact = ""
        long_sents = sorted(sec_lines, key=lambda s:-len(s))
        for ls in long_sents[:3]:
            if 30 < len(ls) < 200 and ls not in [b for b in best]:
                key_fact = ls[:120]; break

        slides.append({
            "title": sec_title[:65],
            "content": best,
            "key_term": key_term,
            "key_fact": key_fact if style != "summary" else "",
            "image_keyword": img_kw(sec_title),
            "image_b64": ""
        })
        if len(slides) >= target:
            break

    # Fallback if no slides
    if not slides:
        content = [clean_bullet(l) for l in lines if len(l.split())>5][:bullets_per_slide]
        slides = [{"title": "Overview", "content": content, "key_term": "",
                   "key_fact": "", "image_keyword": detected_topic or "education", "image_b64": ""}]

    # ── 7. Generate objectives ────────────────────────────────────────────────
    verbs = ["Understand","Explain","Identify","Describe","Analyse","Evaluate","Apply","Compare"]
    objectives = []
    for i, (_, sec_lines) in enumerate(sections[:4]):
        kws = [w for w in sec_lines[0].split()[:5] if w.lower() not in STOP and len(w)>3][:3]
        phrase = " ".join(kws) if kws else detected_topic or "key concepts"
        objectives.append(f"{verbs[i % len(verbs)]} {phrase.lower()}")

    # ── 8. Summary ────────────────────────────────────────────────────────────
    summary_sents = sorted(
        [(score_sentence(l), l) for l in lines if len(l.split())>6],
        reverse=True)[:5]
    summary = [clean_bullet(l)[:100] for _,l in summary_sents] or [f"Key concepts of {detected_topic}"]

    # ── 9. Activity ───────────────────────────────────────────────────────────
    activity_opts = [
        f"In pairs, discuss: What is the most important idea from {detected_topic}?",
        f"Write 3 things you learned about {detected_topic} today.",
        f"Create a mind map connecting the key ideas from {detected_topic}.",
        f"Quiz a partner: take turns asking questions about {detected_topic}.",
        f"Summarise {detected_topic} in exactly 3 sentences.",
    ]
    activity = activity_opts[len(slides) % len(activity_opts)]

    return {
        "topic": detected_topic,
        "subtitle": f"A structured lesson on {detected_topic}",
        "subject": subject or "",
        "grade": grade or "",
        "objectives": objectives or [f"Understand {detected_topic}", "Apply the concepts", "Evaluate key ideas"],
        "slides": slides,
        "activity": activity,
        "summary": summary,
    }

# ── Image sources (open / free — tried in order) ─────────────────────────────
_IMG_TIMEOUT = 2.5  # per-request timeout — matches race window so threads exit cleanly

def _img_unsplash(kw: str) -> bytes:
    import requests
    r = requests.get(f"https://source.unsplash.com/640x400/?{kw.replace(' ',',')}",
                     timeout=_IMG_TIMEOUT, allow_redirects=True, headers={"User-Agent":"Mozilla/5.0"})
    return r.content if r.status_code == 200 and len(r.content) > 3000 else b""

def _img_openverse(kw: str) -> bytes:
    """Openverse — Creative Commons licensed images, no API key needed."""
    import requests
    try:
        q = kw.replace(' ', '%20')
        r = requests.get(f"https://api.openverse.org/v1/images/?q={q}&license_type=all&page_size=3",
                         timeout=_IMG_TIMEOUT, headers={"User-Agent":"PPTex/1.0"})
        if r.status_code == 200:
            for res in r.json().get("results", []):
                url = res.get("url", "")
                if url and any(url.lower().endswith(x) for x in ('.jpg', '.jpeg', '.png')):
                    r2 = requests.get(url, timeout=_IMG_TIMEOUT, headers={"User-Agent":"PPTex/1.0"})
                    if r2.status_code == 200 and len(r2.content) > 3000:
                        return r2.content
    except: pass
    return b""

def _img_wikimedia(kw: str) -> bytes:
    """Wikimedia Commons — search + imageinfo in ONE API call (fast)."""
    import requests
    try:
        q = kw.replace(' ', '%20')
        # generator=search fetches both file list AND imageinfo in a single request
        r = requests.get(
            f"https://commons.wikimedia.org/w/api.php?action=query"
            f"&generator=search&gsrsearch={q}&gsrnamespace=6&gsrlimit=5"
            f"&prop=imageinfo&iiprop=url&iiurlwidth=640&format=json",
            timeout=_IMG_TIMEOUT, headers={"User-Agent":"PPTex/1.0"})
        if r.status_code != 200: return b""
        for page in r.json().get("query", {}).get("pages", {}).values():
            title = page.get("title", "")
            if title.lower().endswith('.svg'): continue
            ii = page.get("imageinfo", [{}])
            img_url = ii[0].get("thumburl", "") or ii[0].get("url", "")
            if not img_url: continue
            r2 = requests.get(img_url, timeout=_IMG_TIMEOUT, headers={"User-Agent":"PPTex/1.0"})
            if r2.status_code == 200 and len(r2.content) > 3000:
                return r2.content
    except: pass
    return b""

def _img_picsum(kw: str) -> bytes:
    """Picsum — instant deterministic fallback (always succeeds)."""
    import requests
    try:
        r = requests.get(f"https://picsum.photos/seed/{kw.replace(' ','-')}/640/400",
                         timeout=_IMG_TIMEOUT, allow_redirects=True)
        return r.content if r.status_code == 200 and len(r.content) > 3000 else b""
    except: return b""

IMG_SOURCES = [
    ("Unsplash",          _img_unsplash),
    ("Openverse",         _img_openverse),
    ("Wikimedia Commons", _img_wikimedia),
    ("Picsum",            _img_picsum),
]

def fetch_img(kw, log):
    """Race all 4 sources simultaneously — quality wins, Picsum is instant backstop.
    Typical time: 0.5–1.5 s.  Worst case: 2.5 s (all quality sources timeout).
    """
    if not kw: return ""
    if kw in IMG_CACHE:
        log(f"  ⚡ {kw[:25]} (cached)"); return IMG_CACHE[kw]
    try:
        from PIL import Image
        from concurrent.futures import ThreadPoolExecutor, wait as _wait, FIRST_COMPLETED

        QUALITY = [("Unsplash", _img_unsplash),
                   ("Openverse", _img_openverse),
                   ("Wikimedia", _img_wikimedia)]

        with ThreadPoolExecutor(max_workers=4) as ex:
            q_futs = {ex.submit(fn, kw): nm for nm, fn in QUALITY}
            p_fut  = ex.submit(_img_picsum, kw)  # always fast backstop

            raw, src = b"", ""

            # Give quality sources 2.3 s head start
            done, _ = _wait(list(q_futs), timeout=2.3, return_when=FIRST_COMPLETED)
            for f in done:
                try:
                    r = f.result(timeout=0)
                    if r: raw, src = r, q_futs[f]; break
                except: pass

            # Nothing quality yet → grab Picsum (should be done in ~0.3 s)
            if not raw:
                try: raw, src = p_fut.result(timeout=1.0), "Picsum"
                except: pass

        if raw:
            img = Image.open(io.BytesIO(raw)).convert("RGB").resize((640, 400))
            buf = io.BytesIO(); img.save(buf, "JPEG", quality=75)
            b64 = base64.b64encode(buf.getvalue()).decode()
            IMG_CACHE[kw] = b64
            log(f"  📸 {kw[:25]} · {src}")
            return b64
    except Exception: pass
    log(f"  ❌ {kw[:25]}")
    return ""

def fetch_images_parallel(slides, log, progress_cb=None):
    """10 outer workers × racing inner fetch — hard 22 s cap on entire batch.
    progress_cb(done, total) called after each image completes.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    try: from concurrent.futures import TimeoutError as FutTimeout
    except ImportError: FutTimeout = Exception

    keywords = [sl.get("image_keyword", "") for sl in slides]
    results  = [""] * len(slides)
    total    = len(slides)
    done_n   = [0]
    log(f"🖼️  Fetching {total} images in parallel…")
    with ThreadPoolExecutor(max_workers=10) as ex:
        futures = {ex.submit(fetch_img, kw, log): i for i, kw in enumerate(keywords)}
        try:
            for f in as_completed(futures, timeout=22):
                try: results[futures[f]] = f.result()
                except: pass
                done_n[0] += 1
                if progress_cb: progress_cb(done_n[0], total)
        except FutTimeout:
            log("  ⏱️  Image batch cap reached — using what we have")
    return results

def run_job(job_id, filepath, topic, subject, grade, fetch_images, length, style, theme):
    import time
    job = JOBS[job_id]
    t0 = time.time()
    def log(msg): job["log"].append(msg); print(msg)
    def pct(p): job["progress"] = int(p)   # thread-safe int write
    def elapsed(): return f"{time.time()-t0:.1f}s"
    try:
        pct(5)
        log("📖 Reading your lesson file…")
        text = extract_text(filepath)
        pct(15)
        log(f"   ✅ {len(text):,} chars extracted ({elapsed()})")

        pct(20)
        log(f"🧠 Analysing content ({length} · {style} · {theme} theme)…")
        data = structure(text, topic, subject, grade, log, length, style)
        pct(40)
        data["theme"] = theme
        slides = data.get("slides", [])
        log(f"   ✅ Planned {len(slides)} slides on \"{data.get('topic')}\" ({elapsed()})")

        if fetch_images and slides:
            n = len(slides)
            # Images go from 40 % → 80 %, one tick per completed image
            def img_progress(done, total):
                pct(40 + int(done / total * 40))
            imgs = fetch_images_parallel(slides, log, progress_cb=img_progress)
            for sl, img in zip(slides, imgs): sl["image_b64"] = img
            got = sum(1 for i in imgs if i)
            log(f"   ✅ {got}/{n} images ready ({elapsed()})")
            pct(82)
        else:
            for sl in slides: sl["image_b64"] = ""
            pct(45)

        pct(85)
        log("🎨 Designing slides with your theme…")
        sys.path.insert(0, str(SCRIPT_DIR))
        from generate_pptx import generate_slides
        pptx_bytes = generate_slides(data)
        pct(97)
        out = SCRIPT_DIR / f"_out_{job_id}.pptx"; out.write_bytes(pptx_bytes)
        job["file"] = str(out)
        job["filename"] = data.get("topic","Lesson").replace(" ","_") + "_Presentation.pptx"
        log(f"✅ Done in {elapsed()}! {job['filename']}")
        pct(100); job["status"] = "done"
    except Exception as e:
        log(f"❌ Error: {e}"); job["status"] = "error"
    finally:
        try: os.remove(filepath)
        except: pass

# ── HTML loaders ──────────────────────────────────────────────────────────────
def _html(name):
    p = SCRIPT_DIR / name
    return p.read_text(encoding="utf-8") if p.exists() else f"<h1>{name} missing</h1>"

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/debug-auth")
def debug_auth():
    """Temporary debug route — shows env var status and forces admin login."""
    pw = os.environ.get("ADMIN_PASSWORD","")
    session["username"] = "admin"
    session["name"] = "Administrator"
    session["role"] = "admin"
    return f"<h2>Debug Auth</h2><p>ADMIN_PASSWORD length: {len(pw)}</p><p>Value: '{pw}'</p><p>Admin session created. <a href='/'>Click here to continue</a></p>"

@app.route("/login", methods=["GET","POST"])
def login_page():
    error = ""
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","")
        # Admin: always check directly against ADMIN_PASSWORD env var
        admin_pw = os.environ.get("ADMIN_PASSWORD","").strip()
        if u == "admin" and admin_pw and p == admin_pw:
            session["username"] = "admin"
            session["name"] = "Administrator"
            session["role"] = "admin"
            return redirect(url_for("index"))
        # Other users: check users.json with hashed passwords
        users = load_users()
        usr = users.get(u)
        if usr and check_password_hash(usr["password"], p):
            session["username"] = u
            session["name"] = usr.get("name", u)
            session["role"] = usr.get("role","teacher")
            return redirect(url_for("index"))
        error = "Invalid username or password"
    return render_template_string(_html("login.html"), error=error)

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login_page"))

@app.route("/")
@login_required
def index():
    html = _html("index.html")
    admin_link = '<a href="/admin" style="color:var(--gold);font-size:12px;margin-left:auto;text-decoration:none">⚙️ Admin</a>' if session.get("role")=="admin" else ""
    logout_link = f'<a href="/logout" style="color:#ccc;font-size:12px;margin-left:12px;text-decoration:none">👋 Logout ({session.get("name","")})</a>'
    html = html.replace("<!--ADMIN_LINK-->", admin_link + logout_link)
    return html

# ── Admin ─────────────────────────────────────────────────────────────────────
@app.route("/admin")
@admin_required
def admin_page():
    users = load_users()
    user_list = [{"username":k,"name":v.get("name",""),"role":v.get("role","teacher")} for k,v in users.items()]
    html = _html("admin.html")
    html = html.replace("{{user_list}}", json.dumps(user_list))
    html = html.replace("{{api_key_set}}", "true" if os.environ.get("ANTHROPIC_API_KEY","").strip() else "false")
    return html

@app.route("/admin/add-user", methods=["POST"])
@admin_required
def add_user():
    d = request.get_json(force=True) or {}
    uname,pw,name,role = d.get("username","").strip(),d.get("password","").strip(),d.get("name","").strip(),d.get("role","teacher")
    if not uname or not pw: return jsonify({"error":"Username and password required"}),400
    users = load_users()
    if uname in users: return jsonify({"error":"Username already exists"}),400
    users[uname] = {"password":generate_password_hash(pw),"name":name or uname,"role":role}
    save_users(users); return jsonify({"ok":True})

@app.route("/admin/delete-user", methods=["POST"])
@admin_required
def delete_user():
    d = request.get_json(force=True) or {}
    uname = d.get("username","").strip()
    if uname == "admin": return jsonify({"error":"Cannot delete admin"}),400
    users = load_users(); users.pop(uname,None); save_users(users)
    return jsonify({"ok":True})

@app.route("/admin/change-password", methods=["POST"])
@admin_required
def change_password():
    d = request.get_json(force=True) or {}
    uname,pw = d.get("username","").strip(),d.get("password","").strip()
    if not uname or not pw: return jsonify({"error":"Missing fields"}),400
    users = load_users()
    if uname not in users: return jsonify({"error":"User not found"}),404
    users[uname]["password"] = generate_password_hash(pw); save_users(users)
    return jsonify({"ok":True})

@app.route("/admin/set-api-key", methods=["POST"])
@admin_required
def set_api_key_admin():
    d = request.get_json(force=True) or {}
    key = d.get("key","").strip()
    if key: os.environ["ANTHROPIC_API_KEY"] = key
    return jsonify({"ok":True})

@app.route("/set_key", methods=["POST"])
@login_required
def set_key():
    key = (request.get_json(force=True) or {}).get("key","").strip()
    if key: os.environ["ANTHROPIC_API_KEY"] = key
    return jsonify({"ok":True})

MAX_CONCURRENT_JOBS = 10  # total server cap — safe ceiling for 512 MB free-tier RAM

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    user = session["username"]
    # 1 per user at a time
    user_running = sum(1 for j in JOBS.values() if j["status"] == "running" and j.get("user") == user)
    if user_running >= 1:
        return jsonify({"error": "You already have a presentation generating — please wait for it to finish."}), 429
    # Total server cap (RAM safety)
    total_running = sum(1 for j in JOBS.values() if j["status"] == "running")
    if total_running >= MAX_CONCURRENT_JOBS:
        return jsonify({"error": f"Server is busy ({total_running}/10 slots used). Please try again in a moment."}), 429
    f = request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"No file uploaded"}),400
    ext = Path(f.filename).suffix.lower()
    if ext not in (".pdf",".docx",".txt",".md"): return jsonify({"error":f"Unsupported format: {ext}"}),400
    tmp = tempfile.NamedTemporaryFile(delete=False,suffix=ext,dir=str(SCRIPT_DIR))
    f.save(tmp.name); tmp.close()
    job_id = uuid.uuid4().hex[:10]
    JOBS[job_id] = {"status":"running","log":[],"file":None,"filename":None,"progress":0,"user":user}
    threading.Thread(target=run_job,args=(job_id,tmp.name,
        request.form.get("topic","").strip(),
        request.form.get("subject","").strip(),
        request.form.get("grade","").strip(),
        request.form.get("images","true").lower()=="true",
        request.form.get("length","medium"),
        request.form.get("style","descriptive"),
        request.form.get("theme","classic")),daemon=True).start()
    return jsonify({"job_id":job_id})

@app.route("/status/<job_id>")
@login_required
def status(job_id):
    job = JOBS.get(job_id)
    if not job: return jsonify({"error":"Not found"}),404
    return jsonify({"status":job["status"],"log":job["log"],"progress":job.get("progress",0)})

@app.route("/download/<job_id>")
@login_required
def download(job_id):
    job = JOBS.get(job_id)
    if not job or not job["file"]: return jsonify({"error":"Not ready"}),404
    resp = send_file(job["file"],as_attachment=True,download_name=job["filename"],
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    @resp.call_on_close
    def cleanup():
        try: os.remove(job["file"])
        except: pass
        JOBS.pop(job_id,None)
    return resp

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5000,debug=False)
