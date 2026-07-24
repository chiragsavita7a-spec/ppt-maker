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
        from pypdf import PdfReader
        return "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)
    elif ext == ".docx":
        import docx
        return "\n".join(p.text for p in docx.Document(path).paragraphs)
    return Path(path).read_text(encoding="utf-8", errors="ignore")

def structure(text, topic, subject, grade, log):
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        log("⚠️  No API key — using basic extraction"); return basic_structure(text, topic, subject, grade)
    client = anthropic.Anthropic(api_key=key)
    trunc = text[:6000] + ("..." if len(text) > 6000 else "")
    prompt = f"""You are a teaching assistant creating PowerPoint slides from lesson notes.
INPUT TEXT:\n{trunc}\nTopic: {topic or "auto-detect"}\nSubject: {subject or "auto-detect"}\nGrade: {grade or "general"}
Return ONLY valid JSON (no markdown) with this structure:
{{"topic":"Main topic","subtitle":"One-line description","subject":"Subject","grade":"Grade",
"objectives":["Up to 4 objectives starting with a verb"],
"slides":[{{"title":"Slide title (max 8 words)","content":["Bullet (max 15 words)","up to 5"],"key_term":"Term: definition (optional)","key_fact":"Striking stat or quote (optional)","image_keyword":"2-3 word image search"}}],
"activity":"Short class activity","summary":["Up to 5 takeaways"]}}
Create 4-7 content slides. Each bullet max 15 words."""
    try:
        r = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=2500,
                                   messages=[{"role":"user","content":prompt}])
        raw = re.sub(r'^```(?:json)?\s*','',r.content[0].text.strip(),flags=re.MULTILINE)
        raw = re.sub(r'\s*```\s*$','',raw,flags=re.MULTILINE)
        return json.loads(raw)
    except Exception as e:
        log(f"⚠️  AI error: {e}"); return basic_structure(text, topic, subject, grade)

def basic_structure(text, topic, subject, grade):
    lines=[l.strip() for l in text.split("\n") if l.strip() and len(l.strip())>10]
    slides,title,bullets=[],("Overview"),[]
    for ln in lines[:80]:
        if len(ln)<60 and (ln.isupper() or ln.endswith(':') or len(ln.split())<=5):
            if bullets: slides.append({"title":title[:60],"content":bullets[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""})
            title=ln.rstrip(":"); bullets=[]
        elif len(bullets)<5: bullets.append(ln[:120])
    if bullets: slides.append({"title":title[:60],"content":bullets[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""})
    slides=slides[:6] or [{"title":"Overview","content":lines[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""}]
    return {"topic":topic or "Lesson","subtitle":f"A lesson on {topic or 'this topic'}",
            "subject":subject or "","grade":grade or "",
            "objectives":[f"Understand {topic or 'the topic'}","Apply the concepts","Evaluate key ideas"],
            "slides":slides,
            "activity":"Discuss with a partner: What is the most important idea from today's lesson?",
            "summary":[f"Covered key concepts of {topic or 'this topic'}","Vocabulary and definitions","Real-world applications"]}

def fetch_img(kw, log):
    if not kw: return ""
    try:
        import requests
        from PIL import Image
        url=f"https://source.unsplash.com/800x500/?{kw.replace(' ',',')}"
        r=requests.get(url,timeout=12,allow_redirects=True,headers={"User-Agent":"Mozilla/5.0"})
        if r.status_code==200 and len(r.content)>5000:
            img=Image.open(io.BytesIO(r.content)).convert("RGB").resize((800,500))
            buf=io.BytesIO(); img.save(buf,"JPEG",quality=80)
            log(f"  📸 {kw[:35]}"); return base64.b64encode(buf.getvalue()).decode()
    except Exception as e: log(f"  ⚠️  Image skipped ({e})")
    return ""

def run_job(job_id, filepath, topic, subject, grade, fetch_images):
    job=JOBS[job_id]
    def log(msg): job["log"].append(msg); print(msg)
    try:
        log("📄 Extracting text from file…")
        text=extract_text(filepath); log(f"   ✅ {len(text):,} characters extracted")
        log("🤖 Structuring content with AI…")
        data=structure(text,topic,subject,grade,log)
        log(f"   ✅ {len(data.get('slides',[]))} slides — topic: {data.get('topic')}")
        if fetch_images:
            log("🖼️  Fetching images…")
            for sl in data.get("slides",[]): sl["image_b64"]=fetch_img(sl.get("image_keyword",""),log)
        else:
            for sl in data.get("slides",[]): sl["image_b64"]=""
        log("⚙️  Building presentation…")
        sys.path.insert(0,str(SCRIPT_DIR))
        from generate_pptx import generate_slides
        pptx_bytes=generate_slides(data)
        out=SCRIPT_DIR/f"_out_{job_id}.pptx"; out.write_bytes(pptx_bytes)
        job["file"]=str(out)
        job["filename"]=data.get("topic","Lesson").replace(" ","_")+"_Presentation.pptx"
        log(f"✅ Done! {job['filename']}"); job["status"]="done"
    except Exception as e:
        log(f"❌ Error: {e}"); job["status"]="error"
    finally:
        try: os.remove(filepath)
        except: pass

# ── HTML loaders ──────────────────────────────────────────────────────────────
def _html(name):
    p = SCRIPT_DIR / name
    return p.read_text(encoding="utf-8") if p.exists() else f"<h1>{name} missing</h1>"

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET","POST"])
def login_page():
    error = ""
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","")
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

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    f = request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"No file uploaded"}),400
    ext = Path(f.filename).suffix.lower()
    if ext not in (".pdf",".docx",".txt",".md"): return jsonify({"error":f"Unsupported format: {ext}"}),400
    tmp = tempfile.NamedTemporaryFile(delete=False,suffix=ext,dir=str(SCRIPT_DIR))
    f.save(tmp.name); tmp.close()
    job_id = uuid.uuid4().hex[:10]
    JOBS[job_id] = {"status":"running","log":[],"file":None,"filename":None}
    threading.Thread(target=run_job,args=(job_id,tmp.name,
        request.form.get("topic","").strip(),request.form.get("subject","").strip(),
        request.form.get("grade","").strip(),request.form.get("images","true").lower()=="true"),daemon=True).start()
    return jsonify({"job_id":job_id})

@app.route("/status/<job_id>")
@login_required
def status(job_id):
    job = JOBS.get(job_id)
    if not job: return jsonify({"error":"Not found"}),404
    return jsonify({"status":job["status"],"log":job["log"]})

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
