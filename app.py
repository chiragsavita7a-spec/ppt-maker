"""
Auto PPT Maker — Web Server
Do NOT run this directly. Use:  python run.py
"""

import os, sys, json, re, threading, uuid, tempfile, base64, io
from pathlib import Path
from flask import Flask, request, jsonify, send_file

SCRIPT_DIR = Path(__file__).parent
app = Flask(__name__)
JOBS: dict = {}

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

@app.route("/")
def index(): return HTML

@app.route("/set_key",methods=["POST"])
def set_key():
    key=(request.get_json(force=True) or {}).get("key","").strip()
    if key: os.environ["ANTHROPIC_API_KEY"]=key
    return jsonify({"ok":True})

@app.route("/generate",methods=["POST"])
def generate():
    f=request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"No file uploaded"}),400
    ext=Path(f.filename).suffix.lower()
    if ext not in (".pdf",".docx",".txt",".md"): return jsonify({"error":f"Unsupported format: {ext}"}),400
    tmp=tempfile.NamedTemporaryFile(delete=False,suffix=ext,dir=str(SCRIPT_DIR))
    f.save(tmp.name); tmp.close()
    job_id=uuid.uuid4().hex[:10]
    JOBS[job_id]={"status":"running","log":[],"file":None,"filename":None}
    threading.Thread(target=run_job,args=(job_id,tmp.name,
        request.form.get("topic","").strip(),request.form.get("subject","").strip(),
        request.form.get("grade","").strip(),request.form.get("images","true").lower()=="true"),daemon=True).start()
    return jsonify({"job_id":job_id})

@app.route("/status/<job_id>")
def status(job_id):
    job=JOBS.get(job_id)
    if not job: return jsonify({"error":"Not found"}),404
    return jsonify({"status":job["status"],"log":job["log"]})

@app.route("/download/<job_id>")
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

HTML=open(Path(__file__).parent/"index.html").read() if (Path(__file__).parent/"index.html").exists() else "<h1>UI missing</h1>"

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5000,debug=False)
