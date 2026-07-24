"""
Auto PPT Maker — Multi-user Web App
"""
import os,sys,json,re,threading,uuid,tempfile,base64,io
from pathlib import Path
from functools import wraps
from flask import Flask,request,jsonify,send_file,session,redirect,url_for,render_template_string
from werkzeug.security import generate_password_hash,check_password_hash

app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY",os.urandom(24).hex())
SCRIPT_DIR=Path(__file__).parent
USERS_FILE=SCRIPT_DIR/"users.json"
JOBS:dict={}

def load_users():
    if USERS_FILE.exists(): return json.loads(USERS_FILE.read_text())
    pw=os.environ.get("ADMIN_PASSWORD","admin123")
    u={"admin":{"password":generate_password_hash(pw),"role":"admin","name":"Administrator"}}
    save_users(u); return u
def save_users(u): USERS_FILE.write_text(json.dumps(u,indent=2))
def get_user(u): return load_users().get(u)
def login_required(f):
    @wraps(f)
    def d(*a,**k):
        if "username" not in session: return redirect(url_for("login"))
        return f(*a,**k)
    return d
def admin_required(f):
    @wraps(f)
    def d(*a,**k):
        if "username" not in session: return redirect(url_for("login"))
        u=get_user(session["username"])
        if not u or u.get("role")!="admin": return jsonify({"error":"Admin only"}),403
        return f(*a,**k)
    return d

def extract_text(path):
    ext=Path(path).suffix.lower()
    if ext==".pdf":
        from pypdf import PdfReader
        return "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)
    elif ext==".docx":
        import docx; return "\n".join(p.text for p in docx.Document(path).paragraphs)
    return Path(path).read_text(encoding="utf-8",errors="ignore")

def structure(text,topic,subject,grade,log):
    import anthropic
    key=os.environ.get("ANTHROPIC_API_KEY","")
    if not key: log("⚠️  No API key — using basic extraction"); return basic(text,topic,subject,grade)
    client=anthropic.Anthropic(api_key=key)
    trunc=text[:6000]+("..." if len(text)>6000 else "")
    prompt=f"""Create PowerPoint slides from lesson notes. Return ONLY valid JSON:
INPUT: {trunc}
Topic:{topic or "auto"} Subject:{subject or "auto"} Grade:{grade or "general"}
{{"topic":"Title","subtitle":"One line","subject":"Subject","grade":"Grade",
"objectives":["Up to 4 objectives"],
"slides":[{{"title":"Title (max 8 words)","content":["Bullet (max 15 words)","up to 5"],
"key_term":"Term: def (optional)","key_fact":"Stat (optional)","image_keyword":"2-3 words"}}],
"activity":"Class activity","summary":["Up to 5 takeaways"]}}
Make 4-7 slides. Be concise."""
    try:
        r=client.messages.create(model="claude-haiku-4-5-20251001",max_tokens=2500,messages=[{"role":"user","content":prompt}])
        raw=re.sub(r'^```(?:json)?\s*','',r.content[0].text.strip(),flags=re.MULTILINE)
        raw=re.sub(r'\s*```\s*$','',raw,flags=re.MULTILINE)
        return json.loads(raw)
    except Exception as e:
        log(f"⚠️  AI error: {e}"); return basic(text,topic,subject,grade)

def basic(text,topic,subject,grade):
    lines=[l.strip() for l in text.split("\n") if l.strip() and len(l.strip())>10]
    slides,title,bullets=[],("Overview"),[]
    for ln in lines[:80]:
        if len(ln)<60 and (ln.isupper() or ln.endswith(':') or len(ln.split())<=5):
            if bullets: slides.append({"title":title[:60],"content":bullets[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""})
            title=ln.rstrip(":"); bullets=[]
        elif len(bullets)<5: bullets.append(ln[:120])
    if bullets: slides.append({"title":title[:60],"content":bullets[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""})
    slides=slides[:6] or [{"title":"Overview","content":lines[:5],"key_term":"","key_fact":"","image_keyword":topic or "education","image_b64":""}]
    return {"topic":topic or "Lesson","subtitle":f"A lesson on {topic or 'this topic'}","subject":subject or "","grade":grade or "",
            "objectives":[f"Understand {topic or 'the topic'}","Apply the concepts","Evaluate key ideas"],
            "slides":slides,"activity":"Discuss: What is the most important idea from today's lesson?",
            "summary":[f"Key concepts of {topic or 'this topic'}","Vocabulary and definitions","Real-world applications"]}

def fetch_img(kw,log):
    if not kw: return ""
    try:
        import requests; from PIL import Image
        r=requests.get(f"https://source.unsplash.com/800x500/?{kw.replace(' ',',')}",timeout=12,allow_redirects=True,headers={"User-Agent":"Mozilla/5.0"})
        if r.status_code==200 and len(r.content)>5000:
            img=Image.open(io.BytesIO(r.content)).convert("RGB").resize((800,500))
            buf=io.BytesIO(); img.save(buf,"JPEG",quality=80)
            log(f"  📸 {kw[:35]}"); return base64.b64encode(buf.getvalue()).decode()
    except: log("  ⚠️  Image skipped")
    return ""

def run_job(job_id,filepath,topic,subject,grade,fetch_images):
    job=JOBS[job_id]
    def log(msg): job["log"].append(msg)
    try:
        log("📄 Extracting text…"); text=extract_text(filepath); log(f"   ✅ {len(text):,} characters")
        log("🤖 Structuring with AI…"); data=structure(text,topic,subject,grade,log); log(f"   ✅ {len(data.get('slides',[]))} slides — {data.get('topic')}")
        if fetch_images:
            log("🖼️  Fetching images…")
            for sl in data.get("slides",[]): sl["image_b64"]=fetch_img(sl.get("image_keyword",""),log)
        else:
            for sl in data.get("slides",[]): sl["image_b64"]=""
        log("⚙️  Building slides…")
        sys.path.insert(0,str(SCRIPT_DIR))
        from generate_pptx import generate_slides
        pptx_bytes=generate_slides(data)
        out=SCRIPT_DIR/f"_out_{job_id}.pptx"; out.write_bytes(pptx_bytes)
        job["file"]=str(out); job["filename"]=data.get("topic","Lesson").replace(" ","_")+"_Lesson.pptx"
        log("✅ Done! Ready to download."); job["status"]="done"
    except Exception as e:
        log(f"❌ Error: {e}"); job["status"]="error"
    finally:
        try: os.remove(filepath)
        except: pass

@app.route("/login",methods=["GET","POST"])
def login():
    error=""
    if request.method=="POST":
        u=request.form.get("username","").strip(); p=request.form.get("password","")
        user=get_user(u)
        if user and check_password_hash(user["password"],p):
            session["username"]=u; session["role"]=user.get("role","teacher"); session["name"]=user.get("name",u)
            session.permanent=True; return redirect(url_for("index"))
        error="Invalid username or password"
    return render_template_string(open(SCRIPT_DIR/"login.html").read(),error=error)

@app.route("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    return render_template_string(open(SCRIPT_DIR/"index.html").read(),
        username=session.get("name","Teacher"),is_admin=session.get("role")=="admin")

@app.route("/generate",methods=["POST"])
@login_required
def generate():
    f=request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"No file uploaded"}),400
    ext=Path(f.filename).suffix.lower()
    if ext not in (".pdf",".docx",".txt",".md"): return jsonify({"error":f"Unsupported: {ext}"}),400
    tmp=tempfile.NamedTemporaryFile(delete=False,suffix=ext,dir=str(SCRIPT_DIR))
    f.save(tmp.name); tmp.close()
    job_id=uuid.uuid4().hex[:10]
    JOBS[job_id]={"status":"running","log":[],"file":None,"filename":None}
    user_key=request.form.get("api_key","").strip()
    if user_key: os.environ["ANTHROPIC_API_KEY"]=user_key
    threading.Thread(target=run_job,args=(job_id,tmp.name,
        request.form.get("topic","").strip(),request.form.get("subject","").strip(),
        request.form.get("grade","").strip(),request.form.get("images","true").lower()=="true"),daemon=True).start()
    return jsonify({"job_id":job_id})

@app.route("/status/<job_id>")
@login_required
def status(job_id):
    job=JOBS.get(job_id)
    if not job: return jsonify({"error":"Not found"}),404
    return jsonify({"status":job["status"],"log":job["log"]})

@app.route("/download/<job_id>")
@login_required
def download(job_id):
    job=JOBS.get(job_id)
    if not job or not job["file"]: return jsonify({"error":"Not ready"}),404
    resp=send_file(job["file"],as_attachment=True,download_name=job["filename"],
                   mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    @resp.call_on_close
    def cleanup():
        try: os.remove(job["file"])
        except: pass
        JOBS.pop(job_id,None)
    return resp

@app.route("/admin")
@admin_required
def admin():
    return render_template_string(open(SCRIPT_DIR/"admin.html").read(),
        users=load_users(),username=session.get("name","Admin"),api_key=os.environ.get("ANTHROPIC_API_KEY",""))

@app.route("/admin/add-user",methods=["POST"])
@admin_required
def add_user():
    d=request.get_json(force=True); u=d.get("username","").strip(); p=d.get("password","").strip()
    if not u or not p: return jsonify({"error":"Username and password required"}),400
    users=load_users()
    if u in users: return jsonify({"error":"Username already exists"}),400
    users[u]={"password":generate_password_hash(p),"role":d.get("role","teacher"),"name":d.get("name",u)}
    save_users(users); return jsonify({"ok":True})

@app.route("/admin/delete-user",methods=["POST"])
@admin_required
def delete_user():
    d=request.get_json(force=True); u=d.get("username","")
    if u=="admin": return jsonify({"error":"Cannot delete admin"}),400
    users=load_users(); users.pop(u,None); save_users(users); return jsonify({"ok":True})

@app.route("/admin/change-password",methods=["POST"])
@admin_required
def change_password():
    d=request.get_json(force=True); u=d.get("username",""); p=d.get("password","")
    if not u or not p: return jsonify({"error":"Required"}),400
    users=load_users()
    if u not in users: return jsonify({"error":"User not found"}),404
    users[u]["password"]=generate_password_hash(p); save_users(users); return jsonify({"ok":True})

@app.route("/admin/set-api-key",methods=["POST"])
@admin_required
def set_api_key():
    key=(request.get_json(force=True) or {}).get("key","").strip()
    if key: os.environ["ANTHROPIC_API_KEY"]=key
    return jsonify({"ok":True})

if __name__=="__main__":
    load_users()
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port,debug=False)
