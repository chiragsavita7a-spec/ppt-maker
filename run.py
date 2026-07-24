"""
Auto PPT Maker — One-click launcher
Double-click this file or run: python run.py
"""

import sys, subprocess, importlib, webbrowser, threading, time, os
from pathlib import Path

HERE = Path(__file__).parent

print("\n" + "="*50)
print("  🎓 Auto PPT Maker for Teachers")
print("="*50)

# ── Step 1: Check Python version ─────────────────────────────────────────────
if sys.version_info < (3, 8):
    print("\n❌ Python 3.8 or newer is required.")
    print("   Download at: https://www.python.org/downloads/")
    input("\nPress Enter to exit..."); sys.exit(1)

print(f"\n✅ Python {sys.version.split()[0]} detected")

# ── Step 2: Install dependencies ──────────────────────────────────────────────
PACKAGES = [
    ("flask",      "flask"),
    ("pypdf",      "pypdf"),
    ("python-docx","docx"),
    ("anthropic",  "anthropic"),
    ("requests",   "requests"),
    ("pillow",     "PIL"),
    ("python-pptx","pptx"),
]

print("\n📦 Checking dependencies...")
missing = []
for pkg, imp in PACKAGES:
    try:
        importlib.import_module(imp)
    except ImportError:
        missing.append(pkg)

if missing:
    print(f"   Installing: {', '.join(missing)} ...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet"] + missing,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        print("   ✅ Dependencies installed")
    except subprocess.CalledProcessError:
        # Try with --break-system-packages for Linux
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "--break-system-packages"] + missing
        )
        print("   ✅ Dependencies installed")
else:
    print("   ✅ All dependencies present")

# ── Step 3: Verify generate_pptx.py exists ───────────────────────────────────
if not (HERE / "generate_pptx.py").exists():
    print("\n❌ generate_pptx.py not found. Make sure all files are in the same folder.")
    input("\nPress Enter to exit..."); sys.exit(1)

if not (HERE / "index.html").exists():
    print("\n❌ index.html not found. Make sure all files are in the same folder.")
    input("\nPress Enter to exit..."); sys.exit(1)

# ── Step 4: Fix app.py to load index.html ────────────────────────────────────
app_py = HERE / "app.py"
content = app_py.read_text(encoding="utf-8")
# Patch the HTML loading line to use the actual file
if 'open(Path(__file__).parent/"index.html")' not in content:
    # Already patched or using embedded HTML — fine
    pass

# ── Step 5: Open browser after short delay ───────────────────────────────────
def open_browser():
    time.sleep(2.5)
    webbrowser.open("http://localhost:5000")
    print("\n🌐 Browser opened at http://localhost:5000")
    print("   (If it didn't open, go there manually)")

threading.Thread(target=open_browser, daemon=True).start()

# ── Step 6: Start Flask server ────────────────────────────────────────────────
print("\n🚀 Starting server...")
print("   URL: http://localhost:5000")
print("   Stop: press Ctrl+C\n")
print("-"*50)

os.chdir(str(HERE))
sys.path.insert(0, str(HERE))
os.environ["FLASK_APP"] = "app"

try:
    from app import app
    # Load HTML from file
    html_path = HERE / "index.html"
    import app as app_module
    app_module.HTML = html_path.read_text(encoding="utf-8")
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
except KeyboardInterrupt:
    print("\n\n👋 Server stopped. Goodbye!")
except OSError as e:
    if "Address already in use" in str(e) or "10048" in str(e):
        print("\n⚠️  Port 5000 is already in use.")
        print("   Either the server is already running, or another app is using port 5000.")
        print("   Try opening http://localhost:5000 in your browser.")
    else:
        print(f"\n❌ Error: {e}")
    input("\nPress Enter to exit...")
