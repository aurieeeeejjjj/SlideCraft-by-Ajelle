
import streamlit as st
from pathlib import Path
from datetime import date
import base64, io, json, re, tempfile, os, requests
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

ROOT = Path(__file__).parent
BG = ROOT / "assets" / "slidecraft_background.png"

st.set_page_config(
    page_title="SlideCraft-by-Ajelle",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

def data_uri(path):
    if not path.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()

bg = data_uri(BG)

st.markdown(f"""
<style>
:root {{
  --purple:#6f25b5;
  --purple2:#9630df;
  --purpleDark:#47106f;
  --purpleSoft:#f7f1fc;
  --ink:#25192f;
  --muted:#6e6179;
  --border:#c9a8ea;
}}
.stApp {{
  background:
    linear-gradient(rgba(35,10,66,.22), rgba(19,29,82,.22)),
    url("{bg}") center/cover fixed;
}}
.block-container {{
  max-width: 1250px;
  padding-top: 1.2rem;
  padding-bottom: 4rem;
}}
#MainMenu, footer {{visibility:hidden}}

.hero {{
  background: rgba(255,255,255,.96);
  border: 1px solid #dfcff0;
  border-radius: 26px;
  padding: 28px 34px;
  box-shadow: 0 18px 50px rgba(30,8,60,.22);
  margin-bottom: 18px;
}}
.hero h1 {{
  margin:0;
  color:var(--purpleDark);
  font-size:2.55rem;
  letter-spacing:-.03em;
}}
.hero p {{
  color:#594c64;
  font-size:1.04rem;
  line-height:1.55;
  margin:.65rem 0 0;
}}
.pill {{
  display:inline-block;
  margin-top:14px;
  padding:8px 14px;
  border-radius:999px;
  background:linear-gradient(135deg,var(--purple),var(--purple2));
  color:white;
  font-weight:800;
  font-size:.8rem;
  letter-spacing:.08em;
}}

div[data-testid="stForm"] {{
  background:rgba(255,255,255,.97);
  border:1px solid #e4d5f3;
  border-radius:26px;
  padding:26px 28px 32px;
  box-shadow:0 18px 55px rgba(31,8,60,.22);
}}
.section {{
  background:#fbf8ff;
  border:1px solid #e8daf6;
  border-left:5px solid var(--purple);
  border-radius:16px;
  padding:15px 18px;
  margin:13px 0 15px;
}}
.section b {{
  color:#541484;
  font-size:1.20rem;
}}
.section span {{
  display:block;
  color:var(--muted);
  margin-top:3px;
  font-size:.94rem;
}}
.note {{
  background:#f6effd;
  border-radius:12px;
  padding:10px 14px;
  color:#513664;
  margin-bottom:12px;
}}

div[data-testid="stWidgetLabel"] p {{
  color:#321943!important;
  font-weight:750!important;
  font-size:.97rem!important;
}}

.stTextInput input,
.stTextArea textarea,
.stNumberInput input,
.stDateInput input {{
  background:#fff!important;
  color:#24182e!important;
  -webkit-text-fill-color:#24182e!important;
  border:1.5px solid var(--border)!important;
  border-radius:12px!important;
}}
.stTextInput input::placeholder,
.stTextArea textarea::placeholder,
.stNumberInput input::placeholder {{
  color:#85748f!important;
  opacity:1!important;
}}

div[data-baseweb="select"],
div[data-baseweb="select"] > div,
div[data-baseweb="select"] div,
div[data-baseweb="base-input"],
div[data-baseweb="input"] > div {{
  background-color:#ffffff!important;
  color:#24182e!important;
}}
div[data-baseweb="select"] > div {{
  border:1.5px solid var(--border)!important;
  border-radius:12px!important;
}}
div[data-baseweb="select"] span,
div[data-baseweb="select"] svg {{
  color:#24182e!important;
  fill:#5f2b7b!important;
}}
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li,
[role="listbox"],
[role="option"] {{
  background:#ffffff!important;
  color:#24182e!important;
}}
[role="option"]:hover {{
  background:#f3e9fb!important;
}}

div[data-testid="stFileUploader"] {{
  background:#fbf8ff;
  border:1.5px dashed #ab79d7;
  border-radius:15px;
  padding:5px;
}}
div[data-testid="stFileUploader"] * {{
  color:#342044!important;
}}
.stCheckbox label span,
.stRadio label span {{
  color:#352047!important;
}}

.stButton>button,
.stDownloadButton>button {{
  width:100%;
  min-height:54px;
  border:0!important;
  border-radius:14px!important;
  background:linear-gradient(135deg,var(--purple),var(--purple2))!important;
  color:#fff!important;
  font-weight:800!important;
  font-size:1rem!important;
  box-shadow:0 10px 24px rgba(101,28,173,.23);
}}

section[data-testid="stSidebar"] {{
  background:rgba(251,248,255,.97);
  border-right:1px solid #e3d2f3;
}}
section[data-testid="stSidebar"] * {{
  color:#351c4d!important;
}}
@media(max-width:900px) {{
  .hero h1 {{font-size:2rem}}
  div[data-testid="stForm"] {{padding:18px}}
}}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ✨ SlideCraft")
    st.caption("by Ajelle")
    st.markdown("---")
    st.markdown("### Generate PPT")
    st.write("Build an editable teacher visual aid from a few lesson details.")
    st.markdown("---")
    st.markdown("### How it works")
    st.write("Required fields give the lesson context.")
    st.write("Optional blanks are completed automatically by AI.")
    st.write("Uploaded curriculum materials become the strongest alignment reference.")
    st.markdown("---")
    st.caption("Teach • Create • Inspire")

st.markdown("""
<div class="hero">
<h1>SlideCraft-by-Ajelle</h1>
<p>Create a complete teacher-ready visual aid—not just an outline. SlideCraft can generate
explanations, key concepts, examples, guided practice, student activities, processing questions,
assessment, answer key, assignment, summary, and closing slides.</p>
<div class="pill">TEACH ✦ CREATE ✦ INSPIRE</div>
</div>
""", unsafe_allow_html=True)

def section(n, title, helptext):
    st.markdown(
        f'<div class="section"><b>{n}. {title}</b><span>{helptext}</span></div>',
        unsafe_allow_html=True
    )

def read_upload(upload):
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
            d = Document(io.BytesIO(raw))
            parts = [p.text.strip() for p in d.paragraphs if p.text.strip()]
            for table in d.tables:
                for row in table.rows:
                    vals = [c.text.strip() for c in row.cells if c.text.strip()]
                    if vals:
                        parts.append(" | ".join(vals))
            return "\n".join(parts)
        if name.endswith(".pptx"):
            from pptx import Presentation as PPTReader
            prs = PPTReader(io.BytesIO(raw))
            parts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        parts.append(shape.text.strip())
            return "\n".join(parts)
    except Exception:
        return ""
    return ""

def get_api_key():
    try:
        return st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        return ""

SLIDE_GUIDE = [
    ("Lesson Title", "Topic/title, subject, grade level, teacher name, date"),
    ("Learning Objectives", "3 clear objectives"),
    ("Learning Competency", "Required curriculum competency or generated equivalent"),
    ("Review / Recall", "Previous lesson review or 2–3 recall questions"),
    ("Motivation / Engage", "Opening activity, picture prompt, scenario, question, quotation, or challenge"),
    ("Lesson Introduction", "What the lesson is about and why it matters"),
    ("Key Concept 1", "Major idea with explanation, terms, and example"),
    ("Key Concept 2", "Second major idea with explanation and example"),
    ("Key Concept 3", "Third major idea or supporting information"),
    ("Example / Application", "Real-life or worked example / demonstration"),
    ("Guided Practice", "Teacher-guided activity with clear instructions"),
    ("Student Activity", "Individual, pair, or group application task"),
    ("Processing Questions", "2–4 reflection/explanation questions"),
    ("Generalization", "Main takeaway / what have we learned"),
    ("Values / Real-Life Connection", "Everyday life, values, responsibility, practical use"),
    ("Assessment", "Actual quiz/items"),
    ("Answer Key", "Correct answers"),
    ("Assignment / Enrichment", "Homework or enrichment"),
    ("Lesson Summary", "Very short recap"),
    ("Closing Slide", "Closing message or reflection prompt"),
]

def target_sections(slide_count):
    names = [x[0] for x in SLIDE_GUIDE]
    if slide_count >= 20:
        return names[:20]
    if slide_count == 15:
        return [
            "Lesson Title", "Learning Objectives + Competency", "Review / Recall",
            "Motivation / Engage", "Lesson Introduction", "Key Concept 1",
            "Key Concept 2", "Key Concept 3 + Example / Application",
            "Guided Practice", "Student Activity", "Processing Questions",
            "Generalization + Values / Real-Life Connection",
            "Assessment", "Answer Key", "Assignment + Lesson Summary + Closing"
        ]
    if slide_count == 10:
        return [
            "Lesson Title", "Objectives + Competency", "Review + Motivation",
            "Lesson Introduction", "Key Concepts 1–2", "Key Concept 3 + Application",
            "Guided Practice + Student Activity", "Processing + Generalization + Values",
            "Assessment + Answer Key", "Assignment + Summary + Closing"
        ]
    # custom
    return names[:min(slide_count, 20)] + [f"Extended Learning {i+1}" for i in range(max(0, slide_count-20))]

def ai_plan(key, payload):
    sections = target_sections(payload["slide_count"])
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"

    prompt = f"""
You are an expert classroom teacher, curriculum specialist, assessment writer,
and instructional PowerPoint designer.

Your task is to create a COMPLETE TEACHER VISUAL AID, not a bare outline.

PRIORITY OF SOURCES
1. Uploaded Curriculum Map / Unit Plan / Lesson Plan / reference materials are the strongest alignment sources.
2. Teacher-entered lesson topic and content must be accurately taught.
3. If optional fields are blank, generate them intelligently from the topic, grade, subject, duration, and uploads.
4. Never invent claims about an uploaded source. Only derive what is supported.

TEACHER-READY CONTENT RULES
- Explain concepts in content that a teacher can actually display and discuss.
- Key concept slides must contain clear explanations, important terms, and examples.
- Avoid vague bullets such as "Discuss the topic" or "Explain the concept."
- Write the actual explanation, example, question, instruction, or problem.
- Include teacher-useful examples and age-appropriate language.
- Include actual guided practice tasks and actual student activity instructions.
- Include actual processing questions.
- Include an actual assessment with the requested number/type of items.
- Include a separate answer key.
- Add an assignment/enrichment task.
- Add a short lesson summary.
- Include a values/real-life connection only when meaningful.
- Follow the chosen strategy while still covering the slide guide.
- Respect the lesson duration; activities should be realistic within the available time.

POWERPOINT READABILITY
- Times New Roman only.
- Minimum 48 pt.
- Keep each slide concise enough to remain readable.
- If content is too much, use more slides or shorter bullets rather than reducing font.
- Prefer 2–4 substantial but concise bullets per slide.
- Use short paragraphs only when an explanation truly needs them.
- Visual suggestions should support the topic, not be decorative.

TARGET NUMBER OF SLIDES: {payload["slide_count"]}
TARGET SLIDE STRUCTURE:
{json.dumps(sections, ensure_ascii=False)}

FULL 20-SLIDE GUIDE FOR REFERENCE:
{json.dumps(SLIDE_GUIDE, ensure_ascii=False)}

Return ONLY valid JSON:
{{
  "alignment_summary": {{
    "source": "Uploaded Materials|Teacher Inputs|Combined",
    "competency": "string",
    "objectives": ["string","string","string"],
    "strategy": "string"
  }},
  "slides": [
    {{
      "title": "string",
      "purpose": "string",
      "bullets": ["actual teacher-ready content"],
      "teacher_note": "brief optional teaching cue",
      "visual_query": "specific useful visual search phrase or empty string",
      "is_cover": false
    }}
  ]
}}

Teacher inputs:
{json.dumps(payload, ensure_ascii=False)}
"""

    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.30, "responseMimeType": "application/json"},
    }
    r = requests.post(url, json=body, timeout=120)
    r.raise_for_status()
    return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])

def fallback_plan(payload):
    t = payload["topic"]
    competency = payload["competency"] or f"Demonstrate understanding of the major concepts and applications of {t}."
    objectives = payload["objectives"] or [
        f"Explain the main concepts of {t}.",
        f"Apply the lesson through a guided or collaborative task.",
        f"Demonstrate understanding through a short assessment.",
    ]
    previous = payload["previous_lesson"] or "Recall the most important idea from the previous lesson and connect it to today's topic."
    content = payload["lesson_content"] or f"Core concepts and examples related to {t}."

    slides = [
        {"title": t, "purpose":"Lesson Title", "bullets":[], "teacher_note":"","visual_query":t, "is_cover":True},
        {"title":"Learning Objectives", "purpose":"Objectives", "bullets":objectives[:3], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Learning Competency", "purpose":"Competency", "bullets":[competency], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Review / Recall", "purpose":"Review", "bullets":[previous, f"What prior idea can help us understand {t}?", "Share one connection with a partner."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Motivation / Engage", "purpose":"Engage", "bullets":[f"Observe or imagine a real situation related to {t}.", "What do you notice?", "Why might this matter in real life?"], "teacher_note":"","visual_query":t, "is_cover":False},
        {"title":"Lesson Introduction", "purpose":"Introduction", "bullets":[f"Today's lesson focuses on {t}.", "We will connect important ideas, examples, and applications.", "The goal is to understand the concept well enough to use it."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Key Concept 1", "purpose":"Concept", "bullets":[content[:220]], "teacher_note":"Explain the first major idea clearly.", "visual_query":t, "is_cover":False},
        {"title":"Key Concept 2", "purpose":"Concept", "bullets":[f"Identify another important idea related to {t}.", "Define the key term in your own words.", "Give one clear example."], "teacher_note":"","visual_query":t, "is_cover":False},
        {"title":"Key Concept 3", "purpose":"Concept", "bullets":[f"Connect the earlier ideas to a third important part of {t}.", "Compare it with the previous concept.", "Give one additional example."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Example / Application", "purpose":"Application", "bullets":[f"Apply {t} to a realistic situation.", "Identify the important information.", "Explain how the lesson concept helps solve or understand the situation."], "teacher_note":"","visual_query":t, "is_cover":False},
        {"title":"Guided Practice", "purpose":"Guided Practice", "bullets":["Complete one example together with the teacher.", "Explain each step or idea before moving on.", "Check the class answer and correct misconceptions."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Student Activity", "purpose":"Activity", "bullets":[payload["activity_preference"] or "Work with a partner or group.", f"Create or solve a task that applies {t}.", "Prepare to explain your answer or output."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Processing Questions", "purpose":"Processing", "bullets":[f"What did the activity help you understand about {t}?", "Which part was easiest or most challenging?", "What evidence supports your answer?"], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Generalization", "purpose":"Generalization", "bullets":[f"What have we learned about {t}?", "State the most important idea in one sentence.", "Explain how the key concepts are connected."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Values / Real-Life Connection", "purpose":"Values", "bullets":[f"How can understanding {t} help in real life?", "What responsible action, value, or habit connects to this lesson?"], "teacher_note":"","visual_query":"","is_cover":False},
    ]

    questions, answers = [], []
    for i in range(1, payload["assessment_count"] + 1):
        if payload["assessment_type"] == "Multiple Choice":
            questions.append(f"{i}. Which statement best shows correct understanding of {t}?")
            answers.append(f"{i}. Correct option should match the lesson explanation.")
        elif payload["assessment_type"] == "True or False":
            questions.append(f"{i}. True or False: Evaluate a lesson-based statement about {t}.")
            answers.append(f"{i}. Determine from the lesson content.")
        elif payload["assessment_type"] == "Problem Solving":
            questions.append(f"{i}. Solve a lesson-based problem involving {t}.")
            answers.append(f"{i}. Accept the correct process and final answer.")
        else:
            questions.append(f"{i}. Give a correct response based on the lesson about {t}.")
            answers.append(f"{i}. Accept an accurate response supported by the lesson.")

    slides += [
        {"title":"Assessment", "purpose":"Assessment", "bullets":questions, "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Answer Key", "purpose":"Answer Key", "bullets":answers, "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Assignment / Enrichment", "purpose":"Assignment", "bullets":[f"Find or create one example showing {t} outside the classroom.", "Explain the example in 3–5 sentences or through a simple output."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Lesson Summary", "purpose":"Summary", "bullets":[f"{t} can be understood through its key ideas, examples, and applications.", "Remember the main terms and how they connect.", "Use the concept accurately in new situations."], "teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Closing Reflection", "purpose":"Closing", "bullets":[f"Complete the sentence: One important thing I learned about {t} is...", "What question do you still have?"], "teacher_note":"","visual_query":"","is_cover":False},
    ]

    target = payload["slide_count"]
    if target < len(slides):
        # Keep essential beginning and ending slides; combine concept/activity content.
        essential = slides[:6]
        middle = slides[6:15]
        ending = slides[15:]
        combined = []
        for s in middle:
            combined.append(s)
        slides = (essential + combined + ending)[:target]
        if target >= 3:
            slides[-2] = ending[-2]
            slides[-1] = ending[-1]
    while len(slides) < target:
        slides.insert(-2, {
            "title":"Extended Learning",
            "purpose":"Enrichment",
            "bullets":[f"Explore another example or application of {t}.", "Connect it to the lesson's key concepts.", "Share your explanation with the class."],
            "teacher_note":"","visual_query":t,"is_cover":False
        })

    return {
        "alignment_summary":{"source":"Teacher Inputs","competency":competency,"objectives":objectives[:3],"strategy":payload["strategy"]},
        "slides":slides[:target]
    }

def commons_image(query):
    if not query:
        return None
    try:
        params = {
            "action":"query","generator":"search","gsrsearch":query,
            "gsrnamespace":6,"gsrlimit":6,"prop":"imageinfo",
            "iiprop":"url","iiurlwidth":900,"format":"json","origin":"*"
        }
        r = requests.get("https://commons.wikimedia.org/w/api.php", params=params, timeout=15)
        pages = r.json().get("query",{}).get("pages",{})
        for page in pages.values():
            info = (page.get("imageinfo") or [{}])[0]
            url = info.get("thumburl") or info.get("url")
            if url:
                rr = requests.get(url, timeout=15)
                if rr.ok:
                    return io.BytesIO(rr.content)
    except Exception:
        pass
    return None

THEMES = {
    "Mathematics":(RGBColor(63,75,157),RGBColor(239,242,255),RGBColor(29,36,73)),
    "Science":(RGBColor(44,113,89),RGBColor(238,248,244),RGBColor(26,65,53)),
    "English":(RGBColor(116,61,141),RGBColor(248,240,250),RGBColor(64,34,76)),
    "Filipino":(RGBColor(121,57,104),RGBColor(250,240,247),RGBColor(70,34,60)),
    "Araling Panlipunan":(RGBColor(137,89,47),RGBColor(250,244,236),RGBColor(77,50,29)),
    "TLE / ICT":(RGBColor(55,87,131),RGBColor(238,244,250),RGBColor(31,48,71)),
    "MAPEH":(RGBColor(130,59,114),RGBColor(249,240,247),RGBColor(73,36,65)),
    "Values / ESP":(RGBColor(98,64,141),RGBColor(245,240,250),RGBColor(55,37,77)),
    "Other":(RGBColor(104,38,132),RGBColor(248,241,250),RGBColor(45,31,53))
}
WHITE = RGBColor(255,255,255)

def set_run(run, size, bold, color):
    run.font.name = "Times New Roman"
    run.font.size = Pt(max(48, size))
    run.font.bold = bold
    run.font.color.rgb = color

def add_textbox(slide, text, left, top, width, height, size, bold, color,
                align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color)
    return box

def split_for_readability(slide):
    bullets = []
    for item in slide.get("bullets",[]) or []:
        x = re.sub(r"\s+"," ",str(item)).strip(" •-\n\t")
        if x:
            bullets.append(x)

    expanded = []
    for b in bullets:
        if len(b) <= 135:
            expanded.append(b)
        else:
            parts = [x.strip() for x in re.split(r"(?<=[.!?;:])\s+", b) if x.strip()]
            expanded.extend(parts or [b])

    per = 3
    chunks = [expanded[i:i+per] for i in range(0,len(expanded),per)] or [[]]
    out = []
    for i, c in enumerate(chunks):
        d = dict(slide)
        d["bullets"] = c
        if i:
            d["title"] = slide.get("title","Lesson") + " (continued)"
            d["visual_query"] = ""
        out.append(d)
    return out

def add_visual(slide, image_bytes, accent):
    try:
        im = Image.open(image_bytes).convert("RGB")
        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        im.save(tmp.name, "JPEG", quality=90)
        frame = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(8.65), Inches(1.92), Inches(3.75), Inches(3.72)
        )
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = accent
        slide.shapes.add_picture(
            tmp.name, Inches(8.82), Inches(2.08),
            width=Inches(3.4), height=Inches(3.38)
        )
        os.unlink(tmp.name)
    except Exception:
        pass

def build_ppt(plan, meta, include_visuals):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    accent, soft, ink = THEMES[meta["subject"]]

    output_slides = []
    for s in plan.get("slides",[]):
        output_slides.extend(split_for_readability(s))

    for s in output_slides:
        slide = prs.slides.add_slide(blank)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = soft

        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(.18), Inches(7.5))
        strip.fill.solid()
        strip.fill.fore_color.rgb = accent
        strip.line.fill.background()

        panel = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(.55), Inches(.45), Inches(12.1), Inches(6.5)
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = WHITE
        panel.line.color.rgb = accent
        panel.line.transparency = 68

        if s.get("is_cover"):
            add_textbox(
                slide, meta["topic"], Inches(1.0), Inches(1.20), Inches(11.2), Inches(1.55),
                60, True, accent, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE
            )
            details = f'{meta["subject"]} • {meta["grade"]}\nTeacher: {meta["teacher"]}\n{meta["lesson_date"]} • {meta["duration"]}'
            add_textbox(
                slide, details, Inches(1.2), Inches(3.20), Inches(10.8), Inches(2.0),
                48, False, ink, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE
            )
            continue

        add_textbox(
            slide, s.get("title","Lesson"), Inches(.92), Inches(.68), Inches(11.25), Inches(.85),
            52, True, accent
        )

        img = commons_image(s.get("visual_query","")) if include_visuals and s.get("visual_query") else None
        text_width = Inches(7.05) if img else Inches(11.0)

        box = slide.shapes.add_textbox(Inches(.95), Inches(1.72), text_width, Inches(4.95))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        for i, bullet in enumerate(s.get("bullets",[])[:3]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = bullet
            p.space_after = Pt(9)
            for r in p.runs:
                set_run(r, 48, False, ink)

        if img:
            add_visual(slide, img, accent)

        if s.get("teacher_note"):
            try:
                slide.notes_slide.notes_text_frame.text = "Teacher cue: " + s["teacher_note"]
            except Exception:
                pass

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

with st.form("slidecraft"):
    st.markdown('<div class="note"><b>Only the essential lesson details are required.</b> Leave optional fields blank and SlideCraft will generate them from the topic and uploaded materials.</div>', unsafe_allow_html=True)

    section(1, "Required Lesson Information", "These details define the class context and appear on the lesson title slide.")
    c1, c2 = st.columns(2)
    with c1:
        teacher = st.text_input("Teacher's Name *", placeholder="Example: Maria D. Santos")
        topic = st.text_input("Lesson Topic / Title *", placeholder="Example: The Water Cycle")
        grade = st.text_input("Grade Level / Section *", placeholder="Example: Grade 7 - Rizal")
    with c2:
        subject = st.selectbox("Subject / Learning Area *", list(THEMES.keys()))
        duration = st.text_input("Lesson Duration *", placeholder="Example: 60 minutes")
        lesson_date = st.date_input("Lesson Date", value=date.today())

    section(2, "Optional Alignment Details", "Enter these if you already have them. If blank, AI will generate them.")
    competency = st.text_area(
        "Learning Competency (Optional)",
        placeholder="Example: Describe the processes involved in the water cycle.",
        height=100,
        help="If blank, AI derives one from your topic and uploaded curriculum/reference materials."
    )
    objectives_text = st.text_area(
        "3 Learning Objectives (Optional)",
        placeholder="One objective per line.\nExample:\nIdentify the stages of the water cycle.\nExplain how each stage works.\nRelate the water cycle to daily weather.",
        height=145,
    )
    previous_lesson = st.text_input(
        "Previous Lesson (Optional)",
        placeholder="Example: States of Matter"
    )

    c3, c4 = st.columns(2)
    with c3:
        strategy = st.selectbox(
            "Teaching Strategy (Optional)",
            ["Auto-select", "I Do, We Do, You Do", "4E: Engage, Explore, Explain, Evaluate",
             "Gamification", "Real-World Connections", "Inquiry-Based Learning",
             "Collaborative Learning", "Problem-Based Learning", "Discussion-Based Learning"]
        )
    with c4:
        activity_preference = st.selectbox(
            "Activity Preference (Optional)",
            ["Auto-generate", "Individual", "Pair Work", "Small Group", "Whole Class", "Hands-on / Performance"]
        )

    section(3, "Lesson Content / Knowledge", "Paste any facts, explanations, examples, formulas, references, or notes you want the presentation to actually teach.")
    lesson_content = st.text_area(
        "Content / Key Points (Optional)",
        placeholder="Example:\n- Evaporation: liquid water changes to water vapor because of heat.\n- Condensation: water vapor cools and forms droplets.\n- Precipitation: water falls as rain, snow, sleet, or hail.\n- Collection: water gathers in rivers, lakes, and oceans.",
        height=220,
        help="If blank, AI creates the lesson content from the topic and uploaded materials."
    )
    extra = st.text_area(
        "Additional Instructions (Optional)",
        placeholder="Example: Use simple Grade 7 language, include a local weather example, and make the activity suitable for a 60-minute class.",
        height=110
    )

    section(4, "Upload Teaching Resources (Optional)", "Upload curriculum or lesson materials. SlideCraft will use them as the strongest alignment references.")
    curriculum = st.file_uploader("Curriculum Map", type=["pdf","docx","pptx","txt"])
    unit_plan = st.file_uploader("Unit Plan", type=["pdf","docx","pptx","txt"])
    lesson_plan = st.file_uploader("Lesson Plan / Reference Material", type=["pdf","docx","pptx","txt"])

    section(5, "Presentation Length & Assessment", "Choose how detailed the presentation should be.")
    p1, p2, p3 = st.columns(3)
    with p1:
        slide_option = st.selectbox("Number of Slides", ["10", "15", "20", "Custom"])
    with p2:
        assessment_type = st.selectbox(
            "Assessment Type (Optional)",
            ["Auto-select", "Mixed", "Multiple Choice", "True or False",
             "Identification", "Short Response", "Problem Solving", "Performance Task"]
        )
    with p3:
        assessment_count = st.number_input("Assessment Questions", min_value=3, max_value=10, value=5)

    custom_count = None
    if slide_option == "Custom":
        custom_count = st.number_input("Custom Slide Count", min_value=8, max_value=30, value=20)

    include_visuals = st.checkbox("Include relevant topic visuals / graphics when available", value=True)

    submit = st.form_submit_button("✨ Generate Teacher Lesson PowerPoint")

if submit:
    if not all([teacher.strip(), topic.strip(), grade.strip(), duration.strip()]):
        st.error("Please complete Teacher's Name, Lesson Topic / Title, Grade Level / Section, and Lesson Duration.")
        st.stop()

    objectives = [x.strip(" •-\t") for x in objectives_text.splitlines() if x.strip()][:3]
    slide_count = int(custom_count if slide_option == "Custom" else slide_option)

    resources = {
        "curriculum_map": read_upload(curriculum)[:12000],
        "unit_plan": read_upload(unit_plan)[:12000],
        "lesson_plan_or_reference": read_upload(lesson_plan)[:12000],
    }

    payload = {
        "teacher_name": teacher.strip(),
        "topic": topic.strip(),
        "subject": subject,
        "grade_level": grade.strip(),
        "duration": duration.strip(),
        "lesson_date": str(lesson_date),
        "competency": competency.strip(),
        "objectives": objectives,
        "previous_lesson": previous_lesson.strip(),
        "strategy": strategy,
        "activity_preference": activity_preference,
        "lesson_content": lesson_content.strip(),
        "additional_instructions": extra.strip(),
        "assessment_type": assessment_type,
        "assessment_count": int(assessment_count),
        "slide_count": slide_count,
        "uploaded_resources": resources,
    }

    with st.spinner("SlideCraft is creating a complete teacher visual aid..."):
        plan = None
        key = get_api_key()
        if key:
            try:
                plan = ai_plan(key, payload)
            except Exception as e:
                st.warning("AI generation was temporarily unavailable, so the built-in generator was used.")
        else:
            st.info("Gemini is not configured in Streamlit Secrets, so the built-in generator is being used.")

        if plan is None:
            plan = fallback_plan(payload)

        meta = {
            "teacher": teacher.strip(),
            "topic": topic.strip(),
            "subject": subject,
            "grade": grade.strip(),
            "duration": duration.strip(),
            "lesson_date": lesson_date.strftime("%B %d, %Y"),
        }
        ppt = build_ppt(plan, meta, include_visuals)

    safe = re.sub(r"[^A-Za-z0-9_-]+","_",topic).strip("_") or "lesson"
    st.success("Your teacher-ready lesson PowerPoint is ready.")
    st.download_button(
        "Download Lesson PowerPoint",
        data=ppt,
        file_name=f"{safe}_SlideCraft.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

    with st.expander("View generated slide structure"):
        for i, s in enumerate(plan.get("slides",[]), 1):
            st.write(f"{i}. {s.get('title','Slide')} — {s.get('purpose','')}")
