
import streamlit as st
from pathlib import Path
import base64, io, json, os, re, tempfile, requests
from html import unescape
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

ROOT = Path(__file__).parent
BG_PATH = ROOT / "assets" / "slidecraft_background.png"

st.set_page_config(
    page_title="SlideCraft-by-Ajelle",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# INTERFACE
# =========================================================
def bg_uri(path):
    if not path.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()

BG = bg_uri(BG_PATH)

st.markdown(f"""
<style>
:root {{
  --purple:#6e24b5;
  --purple2:#9330d9;
  --dark:#48116e;
  --ink:#25182e;
  --muted:#6d6077;
  --border:#c6a5e8;
  --soft:#f8f3fc;
}}
.stApp {{
  background:
    linear-gradient(rgba(31,9,60,.23),rgba(14,30,83,.22)),
    url("{BG}") center/cover fixed;
}}
.block-container {{
  max-width:1240px;
  padding-top:1.25rem;
  padding-bottom:4rem;
}}
#MainMenu, footer {{visibility:hidden}}
.hero {{
  background:rgba(255,255,255,.96);
  border:1px solid #dfcff0;
  border-radius:26px;
  padding:27px 34px;
  box-shadow:0 18px 50px rgba(31,8,60,.21);
  margin-bottom:18px;
}}
.hero h1 {{
  margin:0;
  color:var(--dark);
  font-size:2.45rem;
}}
.hero p {{
  color:#5c4d67;
  line-height:1.55;
  margin:.65rem 0 0;
}}
.mode-note {{
  display:inline-block;
  margin-top:13px;
  padding:8px 14px;
  border-radius:999px;
  color:white;
  font-weight:750;
  background:linear-gradient(135deg,var(--purple),var(--purple2));
}}
.section {{
  background:#fbf9fe;
  border:1px solid #e8daf5;
  border-left:5px solid var(--purple);
  border-radius:16px;
  padding:14px 18px;
  margin:14px 0 15px;
}}
.section b {{
  color:#52117e;
  font-size:1.18rem;
}}
.section span {{
  display:block;
  color:var(--muted);
  margin-top:3px;
  font-size:.94rem;
}}
div[data-testid="stForm"] {{
  background:rgba(255,255,255,.97);
  border:1px solid #e4d5f3;
  border-radius:25px;
  padding:25px 28px 32px;
  box-shadow:0 18px 55px rgba(31,8,60,.21);
}}
div[data-testid="stWidgetLabel"] p {{
  color:#321943!important;
  font-weight:750!important;
}}
.stTextInput input,.stTextArea textarea,.stNumberInput input {{
  background:#fff!important;
  color:#24182e!important;
  -webkit-text-fill-color:#24182e!important;
  border:1.5px solid var(--border)!important;
  border-radius:11px!important;
}}
.stTextInput input::placeholder,.stTextArea textarea::placeholder {{
  color:#87758f!important;
  opacity:1!important;
}}
div[data-baseweb="select"],
div[data-baseweb="select"] > div,
div[data-baseweb="select"] div,
div[data-baseweb="input"] > div,
div[data-baseweb="base-input"] {{
  background:#fff!important;
  color:#24182e!important;
}}
div[data-baseweb="select"] > div {{
  border:1.5px solid var(--border)!important;
  border-radius:11px!important;
}}
div[data-baseweb="select"] span,
div[data-baseweb="select"] * {{
  color:#24182e!important;
  -webkit-text-fill-color:#24182e!important;
}}
[role="listbox"],[role="option"],
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li {{
  background:#fff!important;
  color:#24182e!important;
}}
[role="option"]:hover {{
  background:#f3e9fb!important;
}}
div[data-testid="stFileUploaderDropzone"] {{
  background:#fff!important;
  border:1.5px dashed #9b5bd2!important;
  border-radius:14px!important;
}}
div[data-testid="stFileUploaderDropzone"] *,
div[data-testid="stFileUploader"] button {{
  color:#4f1c70!important;
  -webkit-text-fill-color:#4f1c70!important;
}}
div[data-testid="stFileUploader"] button {{
  background:#f3e9fb!important;
  border:1px solid #b47cdd!important;
}}
.stRadio label span,.stCheckbox label span {{
  color:#352047!important;
}}
.stButton>button,.stDownloadButton>button {{
  width:100%;
  min-height:53px;
  border:0!important;
  border-radius:13px!important;
  background:linear-gradient(135deg,var(--purple),var(--purple2))!important;
  color:#fff!important;
  font-weight:800!important;
}}
section[data-testid="stSidebar"] {{
  background:rgba(251,248,255,.97);
  border-right:1px solid #e3d2f3;
}}
section[data-testid="stSidebar"] * {{
  color:#351c4d!important;
}}
.tip {{
  background:#f7f0fc;
  border-radius:13px;
  padding:11px 14px;
  color:#543667;
  margin-bottom:12px;
}}
.mode-card {{
  background:rgba(255,255,255,.97);
  border:1px solid #dfcff0;
  border-radius:18px;
  padding:16px 20px 12px;
  box-shadow:0 10px 28px rgba(31,8,60,.16);
  margin:0 0 8px;
}}
.mode-title {{
  color:#4b126f;
  font-weight:800;
  font-size:1.15rem;
  margin-bottom:4px;
}}
.mode-sub {{
  color:#6a5875;
  font-size:.94rem;
}}
div[data-testid="stRadio"] {{
  background:rgba(255,255,255,.97)!important;
  border:1px solid #e3d4f1!important;
  border-radius:16px!important;
  padding:10px 16px!important;
  margin-bottom:18px!important;
  box-shadow:0 8px 22px rgba(31,8,60,.12)!important;
}}
div[data-testid="stRadio"] label {{
  background:#f8f2fc!important;
  border:1px solid #d8bde9!important;
  border-radius:12px!important;
  padding:10px 14px!important;
  margin-right:10px!important;
}}
div[data-testid="stRadio"] label span {{
  color:#351943!important;
  -webkit-text-fill-color:#351943!important;
  font-weight:750!important;
}}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## SlideCraft")
    st.caption("by Ajelle")
    st.markdown("---")
    st.markdown("### Two generator modes")
    st.write("Full Lesson to PPT")
    st.write("Quiz to PPT Only")
    st.markdown("---")
    st.markdown("### Readability")
    st.write("Lesson slides never go below 45 pt.")
    st.write("Quiz slides start at 45 pt and shrink only when one long item must fit on a single slide.")
    st.markdown("---")
    st.caption("Simple inputs. Teacher-ready output.")
    st.markdown("---")
    st.markdown("**Created by @Aurieeeejjjj**")

st.markdown("""
<div class="hero">
<h1>SlideCraft-by-Ajelle</h1>
<p>Create either a complete lesson presentation or a quiz-only PowerPoint.
The generator keeps the structure simple, uses teacher-provided content and references,
and arranges the slides so the text stays readable.</p>
<div class="mode-note">Choose your generator mode below</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="mode-card">
  <div class="mode-title">Choose what you want to create</div>
  <div class="mode-sub">Select one option below. You can create a complete lesson presentation or a quiz-only PowerPoint.</div>
</div>
""", unsafe_allow_html=True)

mode = st.radio(
    "Generator Mode",
    ["Full Lesson to PPT", "Quiz to PPT Only"],
    horizontal=True,
    label_visibility="collapsed"
)

def section(n, title, note):
    st.markdown(
        f'<div class="section"><b>{n}. {title}</b><span>{note}</span></div>',
        unsafe_allow_html=True
    )

# =========================================================
# REFERENCES
# =========================================================
def read_uploaded_file(upload):
    if upload is None:
        return ""
    raw = upload.read()
    name = upload.name.lower()
    try:
        if name.endswith(".txt"):
            return raw.decode("utf-8", errors="ignore")
        if name.endswith(".pdf"):
            from pypdf import PdfReader
            return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(raw)).pages)
        if name.endswith(".docx"):
            from docx import Document
            doc = Document(io.BytesIO(raw))
            parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    vals = [c.text.strip() for c in row.cells if c.text.strip()]
                    if vals:
                        parts.append(" | ".join(vals))
            return "\n".join(parts)
        if name.endswith(".pptx"):
            from pptx import Presentation as PptReader
            prs = PptReader(io.BytesIO(raw))
            parts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        parts.append(shape.text.strip())
            return "\n".join(parts)
    except Exception:
        return ""
    return ""

def fetch_link_text(url):
    try:
        r = requests.get(
            url,
            timeout=20,
            headers={"User-Agent":"Mozilla/5.0"}
        )
        r.raise_for_status()
        ctype = r.headers.get("content-type","").lower()
        if "text/html" not in ctype:
            return f"Reference link: {url}"
        text = re.sub(r"(?is)<script.*?>.*?</script>", " ", r.text)
        text = re.sub(r"(?is)<style.*?>.*?</style>", " ", text)
        text = re.sub(r"(?s)<[^>]+>", " ", text)
        text = unescape(text)
        text = re.sub(r"\s+", " ", text).strip()
        return text[:10000]
    except Exception:
        return f"Reference link: {url}"

def collect_references(files, links_text):
    parts = []
    for f in files or []:
        txt = read_uploaded_file(f)
        if txt:
            parts.append(f"FILE: {f.name}\n{txt[:12000]}")
    for line in links_text.splitlines():
        url = line.strip()
        if url.startswith("http://") or url.startswith("https://"):
            parts.append(f"LINK: {url}\n{fetch_link_text(url)}")
    return "\n\n".join(parts)[:30000]

# =========================================================
# GEMINI
# =========================================================
def api_key():
    try:
        return st.secrets.get("GEMINI_API_KEY","")
    except Exception:
        return ""

def choose_models(key):
    try:
        r = requests.get(
            f"https://generativelanguage.googleapis.com/v1beta/models?key={key}",
            timeout=30
        )
        r.raise_for_status()
        models = []
        for item in r.json().get("models",[]):
            if "generateContent" in item.get("supportedGenerationMethods",[]):
                name = item.get("name","").replace("models/","")
                if name:
                    models.append(name)
        models.sort(key=lambda x:(0 if "flash" in x.lower() else 1,
                                  0 if "2.5" in x.lower() else 1,x))
        return models
    except Exception:
        return ["gemini-2.5-flash","gemini-2.0-flash","gemini-1.5-flash"]

def call_gemini(key, prompt):
    body = {
        "contents":[{"parts":[{"text":prompt}]}],
        "generationConfig":{"temperature":0.28,"responseMimeType":"application/json"}
    }
    errs = []
    for model in choose_models(key)[:8]:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
            r = requests.post(url,json=body,timeout=120)
            if not r.ok:
                errs.append(f"{model}:{r.status_code}")
                continue
            text = r.json()["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(text)
        except Exception as e:
            errs.append(f"{model}:{type(e).__name__}")
    raise RuntimeError("Gemini generation failed: " + " | ".join(errs[:4]))

# =========================================================
# LESSON GENERATION
# =========================================================
def lesson_ai_plan(key, data):
    prompt = f"""
You are an expert teacher and PowerPoint instructional designer.

Create a COMPLETE teacher-ready lesson visual aid.

INPUT PRIORITY
1. Teacher's lesson content/knowledge is mandatory source material when provided.
2. Uploaded teaching resources and reference links should be used for alignment.
3. Learning competency and objectives guide the lesson.
4. Do not omit teacher-provided information.
5. If the preferred slide count is not enough to include all content legibly,
   ADD more slides. Never remove important teacher-provided content.

READABILITY
- Times New Roman only in the final PPT.
- Lesson slide text must NEVER be below 45 pt.
- Keep 2–3 concise content blocks/bullets per slide.
- Split dense explanations across continuation slides.
- Do not create long paragraphs.
- Every slide must fit completely inside the presentation.
- Use descriptive slide titles.
- No icons.

LESSON STRUCTURE
Build a logical sequence such as:
Lesson Title
Learning Objectives
Learning Competency
Review / Recall
Motivation / Engage
Lesson Introduction
Key Concepts and explanations
Examples / Applications
Guided Practice
Student Activity
Processing Questions
Generalization
Real-Life / Values Connection when appropriate
Assessment
Closing / Summary

CONTENT QUALITY
- Key concept slides must contain ACTUAL explanations a teacher can use as a visual aid.
- Include important terms and examples.
- Guided practice must contain actual tasks/instructions.
- Student activity must contain actual instructions.
- Assessment must contain ACTUAL items.

ASSESSMENT
For each assessment item:
- Put ONE item on ONE slide.
- Left-align everything.
- For multiple choice, include number, complete question, and all choices.
- Put the correct answer ONLY in "answer_note", not in visible slide content.
- Do not make a visible answer-key slide unless specifically requested.
- Keep assessment item content self-contained.

PREFERRED SLIDE COUNT: {data["preferred_slide_count"]}
This is a preferred count, not a hard maximum. Add slides if necessary for readability or completeness.

Return ONLY valid JSON:
{{
  "slides":[
    {{
      "title":"string",
      "kind":"cover|content|activity|assessment|closing",
      "bullets":["actual visible content"],
      "question":"assessment question or empty",
      "choices":["A. ...","B. ...","C. ...","D. ..."],
      "answer_note":"correct answer and short teacher note, or empty",
      "visual_query":"specific useful visual query or empty"
    }}
  ]
}}

LESSON DATA:
{json.dumps(data, ensure_ascii=False)}
"""
    return call_gemini(key, prompt)

def lesson_fallback(data):
    topic = data["title"]
    comp = data["competency"] or f"Demonstrate understanding of the important concepts related to {topic}."
    objectives = data["objectives"] or [
        f"Explain the major ideas of {topic}.",
        f"Apply the lesson through guided and independent tasks.",
        f"Demonstrate understanding through a short assessment."
    ]
    content = data["content"] or f"Teacher-provided or reference-based information about {topic}."

    slides = [
        {"title":topic,"kind":"cover","bullets":[],"question":"","choices":[],"answer_note":"","visual_query":topic},
        {"title":"Learning Objectives","kind":"content","bullets":objectives[:3],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Learning Competency","kind":"content","bullets":[comp],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Review / Recall","kind":"content","bullets":[f"What do you already know about {topic}?","Share one related idea from a previous lesson."],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Motivation / Engage","kind":"content","bullets":[f"Observe or think of a real situation related to {topic}.","What do you notice? Why might it matter?"],"question":"","choices":[],"answer_note":"","visual_query":topic},
        {"title":"Lesson Introduction","kind":"content","bullets":[f"This lesson focuses on {topic}.","We will examine its main ideas, examples, and applications."],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Key Concepts","kind":"content","bullets":[content[:260]],"question":"","choices":[],"answer_note":"","visual_query":topic},
        {"title":"Guided Practice","kind":"activity","bullets":[f"Work through one example about {topic} with the teacher.","Explain each step or idea before moving on."],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Student Activity","kind":"activity","bullets":[f"Apply what you learned about {topic}.","Work individually or with a partner and prepare to explain your answer."],"question":"","choices":[],"answer_note":"","visual_query":""},
        {"title":"Generalization","kind":"content","bullets":[f"What are the most important ideas about {topic}?","State one key takeaway in your own words."],"question":"","choices":[],"answer_note":"","visual_query":""},
    ]
    count = data["assessment_count"]
    atype = data["assessment_type"]
    for i in range(1,count+1):
        if atype == "Multiple Choice":
            q = f"{i}. Which statement best demonstrates correct understanding of {topic}?"
            ch = ["A. First possible response","B. Second possible response","C. Third possible response","D. Fourth possible response"]
            ans = "Teacher guide: replace with the correct option if using fallback mode."
        elif atype == "True or False":
            q = f"{i}. True or False: Evaluate a statement about {topic}."
            ch = []
            ans = "Teacher guide: verify the correct response using the lesson content."
        else:
            q = f"{i}. Give a correct lesson-based response about {topic}."
            ch = []
            ans = "Teacher guide: accept an accurate answer supported by the lesson."
        slides.append({"title":f"Assessment {i}","kind":"assessment","bullets":[],"question":q,"choices":ch,"answer_note":ans,"visual_query":""})
    slides.append({"title":"Lesson Summary","kind":"closing","bullets":[f"Review the most important concepts about {topic}.","Connect the lesson to one practical situation."],"question":"","choices":[],"answer_note":"","visual_query":""})
    return {"slides":slides}

# =========================================================
# QUIZ PARSER
# =========================================================
def quiz_placeholder(assessment_type):
    if assessment_type == "Multiple Choice":
        return """1. Which process changes liquid water into water vapor?
A. Condensation
B. Evaporation
C. Precipitation
D. Collection
Answer: B

2. Which stage forms clouds?
A. Evaporation
B. Collection
C. Condensation
D. Runoff
Answer: C"""
    if assessment_type == "True or False":
        return """1. Evaporation is caused by heat.
Answer: True

2. Condensation changes liquid water into vapor.
Answer: False"""
    return """1. What process changes water vapor into liquid droplets?
Answer: Condensation

2. Name one form of precipitation.
Answer: Rain"""

def parse_quiz(text, assessment_type):
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text.strip()) if b.strip()]
    items = []
    for idx, block in enumerate(blocks,1):
        lines = [x.strip() for x in block.splitlines() if x.strip()]
        answer = ""
        visible = []
        for line in lines:
            if re.match(r"(?i)^answer\s*:", line):
                answer = re.sub(r"(?i)^answer\s*:\s*", "", line).strip()
            else:
                visible.append(line)

        if not visible:
            continue

        question = visible[0]
        if not re.match(r"^\d+[\.\)]", question):
            question = f"{idx}. {question}"

        choices = []
        if assessment_type == "Multiple Choice":
            choices = visible[1:]
        else:
            # Keep any extra prompt lines as part of the question.
            if len(visible) > 1:
                question += "\n" + "\n".join(visible[1:])

        items.append({
            "number": idx,
            "question": question,
            "choices": choices,
            "answer": answer
        })
    return items

# =========================================================
# PPTX HELPERS
# =========================================================
THEMES = {
    "Mathematics":(RGBColor(62,75,155),RGBColor(239,242,255),RGBColor(29,36,73)),
    "Science":(RGBColor(44,113,89),RGBColor(238,248,244),RGBColor(26,65,53)),
    "English":(RGBColor(116,61,141),RGBColor(248,240,250),RGBColor(64,34,76)),
    "Filipino":(RGBColor(121,57,104),RGBColor(250,240,247),RGBColor(70,34,60)),
    "Araling Panlipunan":(RGBColor(137,89,47),RGBColor(250,244,236),RGBColor(77,50,29)),
    "TLE / ICT":(RGBColor(55,87,131),RGBColor(238,244,250),RGBColor(31,48,71)),
    "MAPEH":(RGBColor(130,59,114),RGBColor(249,240,247),RGBColor(73,36,65)),
    "Values / ESP":(RGBColor(98,64,141),RGBColor(245,240,250),RGBColor(55,37,77)),
    "Other":(RGBColor(104,38,132),RGBColor(248,241,250),RGBColor(45,31,53)),
}
WHITE = RGBColor(255,255,255)

def style_run(run, size, bold, color):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

def textbox(slide,text,l,t,w,h,size,bold,color,align=PP_ALIGN.LEFT,valign=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(l,t,w,h)
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    style_run(r,size,bold,color)
    return shape

def base_slide(prs, subject):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    accent, soft, ink = THEMES[subject]
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = soft

    left_band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(.16),Inches(7.5))
    left_band.fill.solid()
    left_band.fill.fore_color.rgb = accent
    left_band.line.fill.background()

    top_band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(.16),0,Inches(13.17),Inches(.12))
    top_band.fill.solid()
    top_band.fill.fore_color.rgb = accent
    top_band.line.fill.background()

    panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(.55),Inches(.45),Inches(12.15),Inches(6.5)
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = WHITE
    panel.line.color.rgb = accent
    panel.line.transparency = 70
    return slide, accent, soft, ink

def add_note(slide, text):
    if not text:
        return
    try:
        slide.notes_slide.notes_text_frame.text = text
    except Exception:
        pass

def split_lesson_slide(slide_data):
    if slide_data.get("kind") == "assessment":
        return [slide_data]
    bullets = []
    for b in slide_data.get("bullets",[]) or []:
        b = re.sub(r"\s+"," ",str(b)).strip()
        if not b:
            continue
        if len(b) <= 125:
            bullets.append(b)
        else:
            parts = [x.strip() for x in re.split(r"(?<=[.!?;:])\s+",b) if x.strip()]
            bullets.extend(parts or [b])

    chunks = [bullets[i:i+3] for i in range(0,len(bullets),3)] or [[]]
    out = []
    for i,ch in enumerate(chunks):
        d = dict(slide_data)
        d["bullets"] = ch
        if i:
            d["title"] = slide_data.get("title","Lesson") + " (continued)"
            d["visual_query"] = ""
        out.append(d)
    return out

def commons_image(query):
    if not query:
        return None
    try:
        params={"action":"query","generator":"search","gsrsearch":query,"gsrnamespace":6,"gsrlimit":5,
                "prop":"imageinfo","iiprop":"url","iiurlwidth":900,"format":"json","origin":"*"}
        pages=requests.get("https://commons.wikimedia.org/w/api.php",params=params,timeout=15).json().get("query",{}).get("pages",{})
        for p in pages.values():
            info=(p.get("imageinfo") or [{}])[0]
            url=info.get("thumburl") or info.get("url")
            if url:
                r=requests.get(url,timeout=15)
                if r.ok:
                    return io.BytesIO(r.content)
    except Exception:
        pass
    return None

def add_visual(slide, image_bytes, accent):
    try:
        im = Image.open(image_bytes).convert("RGB")
        max_w_px, max_h_px = 1000, 800
        im.thumbnail((max_w_px, max_h_px))
        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        im.save(tmp.name, "JPEG", quality=90)

        frame_left, frame_top = Inches(8.68), Inches(1.92)
        frame_w, frame_h = Inches(3.62), Inches(3.65)

        frame = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            frame_left, frame_top, frame_w, frame_h
        )
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = accent

        img_w, img_h = im.size
        box_w, box_h = 3.30, 3.32
        ratio = min(box_w / img_w, box_h / img_h)
        draw_w = img_w * ratio
        draw_h = img_h * ratio

        left = 8.84 + (box_w - draw_w) / 2
        top = 2.08 + (box_h - draw_h) / 2

        slide.shapes.add_picture(
            tmp.name,
            Inches(left), Inches(top),
            width=Inches(draw_w), height=Inches(draw_h)
        )
        os.unlink(tmp.name)
    except Exception:
        pass

def quiz_font_size(question, choices):
    total = len(question) + sum(len(x) for x in choices)
    if total <= 260:
        return 45
    if total <= 360:
        return 40
    if total <= 500:
        return 36
    if total <= 680:
        return 32
    return 28


def add_answer_key_slides(prs, subject, answers, title="Answer Key"):
    """Create readable visible answer-key slides. Every item is included."""
    if not answers:
        return

    # 5 short answers per slide keeps 45 pt readable.
    # Longer answers are grouped more conservatively.
    chunks = []
    current = []
    current_chars = 0

    for number, answer in answers:
        line = f"{number}. {answer or 'Answer not provided'}"
        # Start a new slide if we already have 5 items or too much text.
        if current and (len(current) >= 5 or current_chars + len(line) > 260):
            chunks.append(current)
            current = []
            current_chars = 0
        current.append(line)
        current_chars += len(line)

    if current:
        chunks.append(current)

    for idx, chunk in enumerate(chunks):
        slide, accent, soft, ink = base_slide(prs, subject)
        heading = title if idx == 0 else f"{title} (continued)"
        textbox(
            slide, heading,
            Inches(.9), Inches(.68), Inches(11.3), Inches(.78),
            48, True, accent
        )

        box = slide.shapes.add_textbox(
            Inches(1.0), Inches(1.65), Inches(10.9), Inches(4.95)
        )
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True

        for i, line in enumerate(chunk):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(10)
            for r in p.runs:
                style_run(r, 45, False, ink)

def build_lesson_ppt(plan, data, include_visuals):
    prs=Presentation()
    prs.slide_width=Inches(13.333)
    prs.slide_height=Inches(7.5)

    slides=[]
    for s in plan.get("slides",[]):
        slides.extend(split_lesson_slide(s))

    assessment_answers = []
    assessment_number = 0

    for s in slides:
        slide,accent,soft,ink=base_slide(prs,data["subject"])

        if s.get("kind")=="cover":
            textbox(slide,data["title"],Inches(1),Inches(1.35),Inches(11.2),Inches(1.55),58,True,accent,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
            textbox(
                slide,
                f'{data["subject"]} {data["year_level"]}',
                Inches(1.3), Inches(3.55), Inches(10.6), Inches(1.15),
                45, False, ink, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE
            )
            continue

        if s.get("kind")=="assessment":
            assessment_number += 1
            textbox(
                slide, f"Assessment {assessment_number}",
                Inches(.9), Inches(.68), Inches(11.3), Inches(.75),
                48, True, accent
            )
            q=s.get("question","").strip()
            choices=s.get("choices",[]) or []
            size=45
            textbox(
                slide, q,
                Inches(.95), Inches(1.55), Inches(11.0), Inches(1.65),
                size, False, ink, PP_ALIGN.LEFT
            )
            if choices:
                ctext="\n".join(choices)
                textbox(
                    slide, ctext,
                    Inches(1.05), Inches(3.15), Inches(10.8), Inches(2.95),
                    45, False, ink, PP_ALIGN.LEFT
                )

            answer_text = (s.get("answer_note","") or "").strip()
            add_note(slide, answer_text)
            assessment_answers.append((assessment_number, answer_text))
            continue

        textbox(slide,s.get("title","Lesson"),Inches(.9),Inches(.68),Inches(11.3),Inches(.78),48,True,accent)

        visual_query = (s.get("visual_query") or "").strip()
        if include_visuals and not visual_query and s.get("kind") in ("content", "activity"):
            visual_query = f'{data["title"]} {s.get("title","lesson concept")}'
        img = commons_image(visual_query) if include_visuals and visual_query else None
        text_w=Inches(7.1) if img else Inches(11.0)
        box=slide.shapes.add_textbox(Inches(.95),Inches(1.65),text_w,Inches(4.95))
        tf=box.text_frame
        tf.clear()
        tf.word_wrap=True
        for i,b in enumerate(s.get("bullets",[])[:3]):
            p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
            p.text=b
            p.space_after=Pt(8)
            for r in p.runs:
                style_run(r,45,False,ink)
        if img:
            add_visual(slide,img,accent)
        add_note(slide,s.get("answer_note",""))

    # Visible answer key is placed at the end and includes every assessment item.
    add_answer_key_slides(prs, data["subject"], assessment_answers, "Answer Key")

    out=io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def build_quiz_ppt(title, year, subject, assessment_type, items):
    prs=Presentation()
    prs.slide_width=Inches(13.333)
    prs.slide_height=Inches(7.5)

    slide,accent,soft,ink=base_slide(prs,subject)
    textbox(slide,title,Inches(1),Inches(1.35),Inches(11.2),Inches(1.6),58,True,accent,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
    textbox(
        slide,
        f"{subject} {year}",
        Inches(1.2), Inches(3.55), Inches(10.8), Inches(1.15),
        45, False, ink, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE
    )

    answer_entries = []

    for item in items:
        slide,accent,soft,ink=base_slide(prs,subject)
        textbox(
            slide, f"Item {item['number']}",
            Inches(.9), Inches(.68), Inches(11.2), Inches(.72),
            48, True, accent
        )
        size=quiz_font_size(item["question"],item["choices"])
        textbox(
            slide, item["question"],
            Inches(.95), Inches(1.48), Inches(11.1), Inches(1.75),
            size, False, ink, PP_ALIGN.LEFT
        )
        if item["choices"]:
            choice_size=min(45,size)
            textbox(
                slide, "\n".join(item["choices"]),
                Inches(1.05), Inches(3.10), Inches(10.8), Inches(3.0),
                choice_size, False, ink, PP_ALIGN.LEFT
            )

        answer = item["answer"] or "Not provided"
        add_note(slide, "Answer: " + answer)
        answer_entries.append((item["number"], answer))

    # Add a complete answer key at the end.
    add_answer_key_slides(prs, subject, answer_entries, "Answer Key")

    out=io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# =========================================================
# FULL LESSON FORM
# =========================================================
if mode == "Full Lesson to PPT":
    with st.form("lesson_form"):
        st.markdown('<div class="tip"><b>Required:</b> Lesson Title, Year Level, and Subject only. Everything else helps SlideCraft align and improve the presentation.</div>',unsafe_allow_html=True)

        section(1,"Required Lesson Information","Only three fields are required.")
        c1,c2,c3=st.columns(3)
        with c1:
            lesson_title=st.text_input("Lesson Title *",placeholder="Example: The Water Cycle")
        with c2:
            year_level=st.text_input("Year Level *",placeholder="Example: Grade 7")
        with c3:
            subject=st.selectbox("Subject *",list(THEMES.keys()))

        section(2,"Learning Competency and Objectives","Enter these when available. They can be left blank.")
        competency=st.text_area("Learning Competency (Optional)",placeholder="Paste the curriculum competency here.",height=95)
        objectives_text=st.text_area(
            "Learning Objectives (Optional)",
            placeholder="One objective per line.\nExample:\nIdentify the stages of the water cycle.\nExplain each stage.\nRelate the water cycle to daily weather.",
            height=135
        )

        section(3,"Lesson Content / Knowledge","Put the actual information that must appear in the PowerPoint.")
        lesson_content=st.text_area(
            "Lesson Content / Knowledge (Optional but recommended)",
            placeholder="Paste the explanations, terms, facts, examples, formulas, Bible verses, procedures, or other lesson information that students need to see.",
            height=260,
            help="SlideCraft will arrange this content across enough slides so it remains readable."
        )

        section(4,"Teaching Resources / Lesson Links","Upload files or paste links that SlideCraft can use to align the lesson.")
        resources=st.file_uploader(
            "Teaching Resources",
            type=["pdf","docx","pptx","txt"],
            accept_multiple_files=True
        )
        links=st.text_area(
            "Lesson / Reference Links",
            placeholder="Paste one link per line.\nhttps://example.com/resource-1\nhttps://example.com/resource-2",
            height=100
        )

        section(5,"Assessment and Presentation Length","Assessment items are placed one per slide. Answers are stored in speaker notes and repeated in a complete Answer Key at the end.")
        c1,c2,c3=st.columns(3)
        with c1:
            assessment_type=st.selectbox(
                "Assessment Type",
                ["Multiple Choice","True or False","Identification","Short Response","Problem Solving"]
            )
        with c2:
            assessment_count=st.number_input("Number of Assessment Items",min_value=0,max_value=20,value=5)
        with c3:
            preferred_slides=st.number_input(
                "Preferred Number of Slides",
                min_value=6,max_value=50,value=15,
                help="SlideCraft may add more slides if needed so no teacher-provided information is omitted or overcrowded."
            )

        include_visuals=st.checkbox("Include relevant topic visuals when available",value=True)
        extra=st.text_area(
            "Additional Instructions (Optional)",
            placeholder="Example: Use simple Grade 7 language and include a group activity.",
            height=90
        )

        submit_lesson=st.form_submit_button("Generate Full Lesson PowerPoint")

    if submit_lesson:
        if not lesson_title.strip() or not year_level.strip():
            st.error("Please enter the Lesson Title and Year Level.")
            st.stop()

        refs=collect_references(resources,links)
        objectives=[x.strip(" •-\t") for x in objectives_text.splitlines() if x.strip()][:5]
        data={
            "title":lesson_title.strip(),
            "year_level":year_level.strip(),
            "subject":subject,
            "competency":competency.strip(),
            "objectives":objectives,
            "content":lesson_content.strip(),
            "references":refs,
            "assessment_type":assessment_type,
            "assessment_count":int(assessment_count),
            "preferred_slide_count":int(preferred_slides),
            "additional_instructions":extra.strip(),
        }

        with st.spinner("Creating your complete lesson presentation..."):
            plan=None
            key=api_key()
            if key:
                try:
                    plan=lesson_ai_plan(key,data)
                except Exception:
                    st.warning("Gemini could not generate this presentation, so the built-in generator was used.")
            if plan is None:
                plan=lesson_fallback(data)

            ppt=build_lesson_ppt(plan,data,include_visuals)

        safe=re.sub(r"[^A-Za-z0-9_-]+","_",lesson_title).strip("_") or "lesson"
        st.success("Your full lesson PowerPoint is ready.")
        st.download_button(
            "Download Full Lesson PowerPoint",
            ppt,
            f"{safe}_SlideCraft_Lesson.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        with st.expander("Generated slide structure"):
            for i,s in enumerate(plan.get("slides",[]),1):
                st.write(f"{i}. {s.get('title','Slide')}")

# =========================================================
# QUIZ-ONLY FORM
# =========================================================
else:
    with st.form("quiz_form"):
        st.markdown('<div class="tip"><b>Quiz mode:</b> one item per slide. Multiple-choice slides show the item number, complete question, and choices. The correct answer is saved in the speaker notes, and a complete Answer Key is also added at the end of the PowerPoint.</div>',unsafe_allow_html=True)

        section(1,"Quiz Information","Only the quiz title, year level, and subject are needed.")
        c1,c2,c3=st.columns(3)
        with c1:
            quiz_title=st.text_input("Quiz Title *",placeholder="Example: Water Cycle Quiz")
        with c2:
            quiz_year=st.text_input("Year Level *",placeholder="Example: Grade 7")
        with c3:
            quiz_subject=st.selectbox("Subject *",list(THEMES.keys()),key="quiz_subject")

        section(2,"Assessment Type","Choose the format of the questions you will paste.")
        quiz_type=st.selectbox(
            "Assessment Type",
            ["Multiple Choice","True or False","Identification / Short Answer"]
        )

        section(3,"Paste Quiz Questions","Separate each item with a blank line. Include an Answer line so the teacher answer is saved in the slide notes.")
        quiz_text=st.text_area(
            "Quiz Questions *",
            placeholder=quiz_placeholder(quiz_type),
            height=390
        )

        submit_quiz=st.form_submit_button("Generate Quiz PowerPoint")

    if submit_quiz:
        if not quiz_title.strip() or not quiz_year.strip() or not quiz_text.strip():
            st.error("Please complete the Quiz Title, Year Level, and Quiz Questions.")
            st.stop()

        items=parse_quiz(quiz_text,quiz_type)
        if not items:
            st.error("No quiz items could be read. Please follow the sample format shown in the question box.")
            st.stop()

        ppt=build_quiz_ppt(
            quiz_title.strip(),
            quiz_year.strip(),
            quiz_subject,
            quiz_type,
            items
        )
        safe=re.sub(r"[^A-Za-z0-9_-]+","_",quiz_title).strip("_") or "quiz"
        st.success(f"Your quiz PowerPoint is ready with {len(items)} item slide(s).")
        st.download_button(
            "Download Quiz PowerPoint",
            ppt,
            f"{safe}_SlideCraft_Quiz.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )


st.markdown("""
<div style="
    margin-top:24px;
    padding:14px 18px;
    border-radius:14px;
    background:rgba(255,255,255,.94);
    border:1px solid #e2d2f0;
    text-align:center;
    color:#4d2863;
    font-weight:700;">
    SlideCraft-by-Ajelle • Created by @Aurieeeejjjj
</div>
""", unsafe_allow_html=True)

