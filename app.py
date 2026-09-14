
import streamlit as st
from pathlib import Path
import json
import os
import io
import re
import tempfile
import requests
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

APP_DIR = Path(__file__).parent
BG_PATH = APP_DIR / "assets" / "purple_background.png"

st.set_page_config(page_title="Purple Lesson PPT Generator", page_icon="📚", layout="wide")

# ---------- UI ----------
def apply_ui():
    bg_css = ""
    if BG_PATH.exists():
        import base64
        b64 = base64.b64encode(BG_PATH.read_bytes()).decode()
        bg_css = f"""
        .stApp {{
            background:
              linear-gradient(rgba(255,255,255,.82), rgba(255,255,255,.90)),
              url("data:image/png;base64,{b64}") center/cover fixed;
        }}
        """
    st.markdown(f"""
    <style>
    {bg_css}
    :root {{
        --purple:#6f2a8e;
        --purple2:#9d58b6;
        --lav:#f4e9f8;
        --ink:#2e1c36;
    }}
    .block-container {{ max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem; }}
    h1,h2,h3 {{ color:var(--purple)!important; }}
    .hero {{
        padding: 28px 30px;
        border-radius: 24px;
        background: rgba(255,255,255,.88);
        border: 1px solid rgba(111,42,142,.18);
        box-shadow: 0 12px 36px rgba(91,31,120,.12);
        margin-bottom: 18px;
    }}
    .hero h1 {{ margin:0; font-size:2.4rem; }}
    .hero p {{ color:#5f4b67; font-size:1.05rem; margin:.5rem 0 0 0; }}
    .stButton>button, .stDownloadButton>button {{
        background: linear-gradient(135deg,var(--purple),var(--purple2));
        color:white; border:none; border-radius:14px; font-weight:700;
        min-height:48px;
    }}
    div[data-testid="stForm"] {{
        background:rgba(255,255,255,.88); padding:22px; border-radius:20px;
        border:1px solid rgba(111,42,142,.16);
    }}
    </style>
    """, unsafe_allow_html=True)

apply_ui()

st.markdown("""
<div class="hero">
<h1>Purple Lesson PPT Generator</h1>
<p>Create learner-centered, editable PowerPoint lessons aligned to your learning plan, lesson content, and teaching strategy.</p>
</div>
""", unsafe_allow_html=True)

# ---------- File reading ----------
def read_learning_plan(uploaded):
    if uploaded is None:
        return ""
    name = uploaded.name.lower()
    data = uploaded.read()
    try:
        if name.endswith(".txt"):
            return data.decode("utf-8", errors="ignore")
        if name.endswith(".docx"):
            from docx import Document
            doc = Document(io.BytesIO(data))
            parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
            return "\n".join(parts)
        if name.endswith(".pdf"):
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(data))
            return "\n".join((p.extract_text() or "") for p in reader.pages)
    except Exception:
        return ""
    return ""

# ---------- AI ----------
def generate_with_gemini(api_key, model, payload):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    prompt = f"""
You are an expert classroom teacher, curriculum planner, and PowerPoint instructional designer.

Create a LEARNER-CENTERED lesson PowerPoint plan.

ALIGNMENT PRIORITY:
- If a learning plan is supplied, use it as the PRIMARY source. Extract and follow its competency, objectives, lesson sequence, activities, assessment, and content.
- Teacher-entered lesson content may expand the learning plan but must not contradict it.
- If no learning plan is supplied, build a complete age-appropriate lesson from the teacher inputs.

NON-NEGOTIABLE POWERPOINT RULES:
- Times New Roman only.
- Minimum font size is 48 pt.
- Keep each slide concise: usually 2 to 4 short bullets.
- Split content into more slides instead of shrinking text.
- Include actual quiz/assessment questions, not merely an assessment heading.
- Include a separate answer key.
- Make activities learner-centered: learners should discuss, solve, observe, create, compare, practice, investigate, or reflect.
- Add real-world connections when appropriate.
- Use the selected teaching strategy as the organizing lesson flow.
- For I Do-We Do-You Do: model, guided practice, independent practice.
- For 4E: Engage, Explore, Explain, Evaluate.
- For Gamification: use missions/challenges/rounds while keeping learning goals central.
- For Inquiry: question, investigate, evidence, explain.
- For Collaborative Learning: include a clear pair/group output.
- For Problem-Based Learning: real problem, investigation, solution, reflection.
- Suggest one relevant visual search phrase per suitable content slide.
- Cover slide must show topic, subject, and teacher name.
- Do not place teacher-only instructions as learner-facing content.

Return ONLY valid JSON. No markdown fences.

JSON schema:
{{
  "alignment_summary": {{
    "source": "Learning Plan|Teacher Inputs|Combined",
    "competency": "string",
    "objectives": ["string", "string", "string"],
    "strategy": "string"
  }},
  "presentation_title": "string",
  "theme_keywords": ["string"],
  "slides": [
    {{
      "title": "string",
      "section": "Cover|Objectives|Review|Engage|Explore|Explain|Practice|Application|Assessment|Answer Key|Closing|Other",
      "bullets": ["short bullet", "short bullet"],
      "teacher_note": "optional concise teaching cue",
      "visual_query": "short search phrase or empty",
      "is_cover": false
    }}
  ]
}}

Teacher inputs:
{json.dumps(payload, ensure_ascii=False)}
"""
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.45,
            "responseMimeType": "application/json"
        }
    }
    r = requests.post(url, json=body, timeout=120)
    r.raise_for_status()
    out = r.json()
    text = out["candidates"][0]["content"]["parts"][0]["text"]
    return json.loads(text)

def fallback_plan(payload):
    topic = payload["topic"]
    subject = payload["subject"]
    teacher = payload["teacher_name"]
    strategy = payload["strategy"]
    content = payload["lesson_content"] or f"Key ideas and examples about {topic}."
    objectives = payload["objectives"] or [
        f"Explain the key idea of {topic}.",
        f"Apply the lesson through a guided activity.",
        f"Show understanding through a short assessment."
    ]
    if isinstance(objectives, str):
        objectives = [x.strip("-• \n") for x in objectives.splitlines() if x.strip()][:3]
    slides = [
        {"title": topic, "section":"Cover","bullets":[subject, f"Teacher: {teacher}"],"teacher_note":"","visual_query":topic,"is_cover":True},
        {"title":"Learning Objectives","section":"Objectives","bullets":objectives[:3],"teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Connect to What You Know","section":"Review","bullets":[f"What do you already know about {topic}?","Share one example with a partner."],"teacher_note":"Invite 2–3 responses.","visual_query":"","is_cover":False},
    ]
    if strategy.startswith("I Do"):
        slides += [
            {"title":"I Do: Teacher Model","section":"Explain","bullets":[content[:220]],"teacher_note":"Model one clear example.","visual_query":topic,"is_cover":False},
            {"title":"We Do: Guided Practice","section":"Practice","bullets":["Solve or analyze one example together.","Ask learners to explain each step."],"teacher_note":"","visual_query":"","is_cover":False},
            {"title":"You Do: Independent Practice","section":"Application","bullets":["Complete a similar task independently.","Compare answers with a partner after finishing."],"teacher_note":"","visual_query":"","is_cover":False},
        ]
    elif strategy.startswith("4E"):
        slides += [
            {"title":"Engage","section":"Engage","bullets":[f"Observe a real-life example connected to {topic}.","What do you notice? What do you wonder?"],"teacher_note":"","visual_query":topic,"is_cover":False},
            {"title":"Explore","section":"Explore","bullets":["Work in pairs or groups on a short task.","Record observations, patterns, or possible answers."],"teacher_note":"","visual_query":"","is_cover":False},
            {"title":"Explain","section":"Explain","bullets":[content[:220]],"teacher_note":"Connect student ideas to the formal concept.","visual_query":topic,"is_cover":False},
            {"title":"Evaluate","section":"Assessment","bullets":["Answer the assessment individually."],"teacher_note":"","visual_query":"","is_cover":False},
        ]
    else:
        slides += [
            {"title":"Discover the Lesson","section":"Engage","bullets":[f"Look at an example related to {topic}.","Discuss what it might mean or show."],"teacher_note":"","visual_query":topic,"is_cover":False},
            {"title":"Key Ideas","section":"Explain","bullets":[content[:220]],"teacher_note":"","visual_query":topic,"is_cover":False},
            {"title":"Learner Activity","section":"Practice","bullets":["Work with a partner or group.","Use the lesson ideas to complete the task.","Explain your answer or output."],"teacher_note":"","visual_query":"","is_cover":False},
            {"title":"Real-World Connection","section":"Application","bullets":[f"Where can we see or use {topic} in everyday life?","Give one practical example."],"teacher_note":"","visual_query":topic,"is_cover":False},
        ]
    qn = max(3, min(int(payload["quiz_count"]), 10))
    qs = [f"{i}. Write one correct idea, answer, or example about {topic}." for i in range(1, qn+1)]
    ans = [f"{i}. Accept a correct response based on the lesson." for i in range(1, qn+1)]
    slides += [
        {"title":"Quick Assessment","section":"Assessment","bullets":qs,"teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Answer Key","section":"Answer Key","bullets":ans,"teacher_note":"","visual_query":"","is_cover":False},
        {"title":"Wrap-Up","section":"Closing","bullets":[f"Complete this sentence: Today I learned that {topic}...","Ask one final question if anything is unclear."],"teacher_note":"","visual_query":"","is_cover":False}
    ]
    return {"presentation_title": topic, "theme_keywords":[subject, topic], "slides":slides}

# ---------- Visuals from Wikimedia Commons ----------
def wikimedia_image(query):
    if not query:
        return None
    try:
        params = {
            "action":"query","generator":"search","gsrsearch":query,
            "gsrnamespace":6,"gsrlimit":6,"prop":"imageinfo",
            "iiprop":"url","iiurlwidth":900,"format":"json","origin":"*"
        }
        r = requests.get("https://commons.wikimedia.org/w/api.php", params=params, timeout=20)
        r.raise_for_status()
        pages = (r.json().get("query", {}) or {}).get("pages", {}) or {}
        for p in pages.values():
            info = (p.get("imageinfo") or [{}])[0]
            url = info.get("thumburl") or info.get("url")
            if url:
                rr = requests.get(url, timeout=20)
                rr.raise_for_status()
                return io.BytesIO(rr.content)
    except Exception:
        return None
    return None

# ---------- PowerPoint helpers ----------
PURPLE = RGBColor(104, 38, 132)
PURPLE2 = RGBColor(152, 86, 177)
WHITE = RGBColor(255,255,255)
INK = RGBColor(44, 31, 51)
SOFT = RGBColor(246, 239, 249)

THEMES = {
    "math": (RGBColor(65,89,150), RGBColor(33,44,74), RGBColor(228,236,250)),
    "science": (RGBColor(50,112,88), RGBColor(29,63,51), RGBColor(231,245,239)),
    "english": (RGBColor(125,70,145), RGBColor(66,39,76), RGBColor(245,235,248)),
    "history": (RGBColor(135,91,50), RGBColor(75,50,29), RGBColor(248,240,228)),
    "ict": (RGBColor(67,92,129), RGBColor(35,49,68), RGBColor(232,239,247)),
    "default": (PURPLE, INK, SOFT),
}

def choose_theme(subject, topic):
    text = f"{subject} {topic}".lower()
    if any(k in text for k in ["math","mathematics","algebra","geometry","calculus","statistics"]): return THEMES["math"]
    if any(k in text for k in ["science","biology","chemistry","physics","earth","environment"]): return THEMES["science"]
    if any(k in text for k in ["english","literature","reading","grammar","language"]): return THEMES["english"]
    if any(k in text for k in ["history","araling panlipunan","rizal","social science"]): return THEMES["history"]
    if any(k in text for k in ["ict","computer","technology","programming","tle"]): return THEMES["ict"]
    return THEMES["default"]

def add_full_bg(slide, prs):
    if BG_PATH.exists():
        slide.shapes.add_picture(str(BG_PATH), 0, 0, width=prs.slide_width, height=prs.slide_height)

def add_overlay(slide, left, top, width, height, transparency=12, line_color=PURPLE2):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.fill.transparency = transparency
    shape.line.color.rgb = line_color
    shape.line.transparency = 55
    return shape

def set_text_style(run, size, bold=False, color=INK):
    run.font.name = "Times New Roman"
    run.font.size = Pt(max(48, size))
    run.font.bold = bold
    run.font.color.rgb = color

def add_textbox(slide, text, left, top, width, height, size=48, bold=False,
                color=INK, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_text_style(r, size, bold, color)
    return box

def split_slide_content(slide_data, max_bullets=4, max_chars=150):
    bullets = slide_data.get("bullets", []) or []
    expanded = []
    for b in bullets:
        b = re.sub(r"\s+", " ", str(b)).strip()
        if len(b) <= max_chars:
            expanded.append(b)
        else:
            parts = re.split(r"(?<=[.!?])\s+", b)
            expanded.extend([p for p in parts if p])
    if slide_data.get("section") in ["Assessment", "Answer Key"]:
        max_bullets = min(max_bullets, 3)
    chunks = [expanded[i:i+max_bullets] for i in range(0, len(expanded), max_bullets)] or [[]]
    result = []
    for i, ch in enumerate(chunks):
        d = dict(slide_data)
        d["bullets"] = ch
        if i > 0:
            d["title"] = slide_data.get("title","Lesson") + " (continued)"
            d["visual_query"] = ""
        result.append(d)
    return result

def build_ppt(plan, teacher_name, subject, topic, grade_level="", include_visuals=True):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    accent, theme_ink, theme_soft = choose_theme(subject, topic)

    normalized = []
    for s in plan["slides"]:
        normalized.extend(split_slide_content(s))

    for idx, s in enumerate(normalized):
        slide = prs.slides.add_slide(blank)
        add_full_bg(slide, prs)

        is_cover = bool(s.get("is_cover")) or s.get("section") == "Cover"
        if is_cover:
            add_overlay(slide, Inches(0.8), Inches(1.15), Inches(11.75), Inches(5.1), 15, accent)
            add_textbox(slide, s.get("title", topic), Inches(1.15), Inches(1.55), Inches(11), Inches(1.7),
                        size=58, bold=True, color=accent, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            sub = subject + (f" • {grade_level}" if grade_level else "") + f"\nTeacher: {teacher_name}"
            add_textbox(slide, sub, Inches(1.3), Inches(3.55), Inches(10.7), Inches(1.65),
                        size=48, color=theme_ink, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            continue

        add_overlay(slide, Inches(0.55), Inches(0.45), Inches(12.2), Inches(6.55), 8, accent)
        add_textbox(slide, s.get("title","Lesson"), Inches(0.9), Inches(0.72), Inches(11.6), Inches(0.85),
                    size=52, bold=True, color=accent)

        bullets = s.get("bullets", []) or []
        visual = wikimedia_image(s.get("visual_query","")) if include_visuals and s.get("visual_query") else None
        text_w = Inches(7.2) if visual else Inches(11.3)
        box = slide.shapes.add_textbox(Inches(0.95), Inches(1.75), text_w, Inches(4.75))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        for j, b in enumerate(bullets[:4]):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = b
            p.level = 0
            p.space_after = Pt(12)
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                set_text_style(run, 48, False, theme_ink)

        if visual:
            try:
                img = Image.open(visual)
                img.thumbnail((1200, 800))
                tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                img.convert("RGB").save(tmp.name, "PNG")
                slide.shapes.add_picture(tmp.name, Inches(8.55), Inches(1.9), width=Inches(3.6), height=Inches(3.6))
                os.unlink(tmp.name)
            except Exception:
                pass


    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# ---------- Form ----------
with st.form("lesson_form"):
    st.subheader("1. Class Information")
    c1, c2 = st.columns(2)
    with c1:
        teacher_name = st.text_input("Teacher's name")
        subject = st.text_input("Subject")
        grade_level = st.text_input("Grade level / section")
    with c2:
        topic = st.text_input("Lesson topic / title")
        strategy = st.selectbox("Teaching strategy", [
            "I Do, We Do, You Do",
            "4E: Engage, Explore, Explain, Evaluate",
            "Gamification",
            "Real-World Connections",
            "Inquiry-Based Learning",
            "Collaborative Learning",
            "Problem-Based Learning",
            "Discussion-Based Learning",
            "Custom / AI chooses the best approach"
        ])
        custom_strategy = st.text_input("Custom strategy (optional)")

    st.subheader("2. Lesson Alignment")
    learning_plan = st.file_uploader("Upload learning plan (optional)", type=["pdf","docx","txt"])
    lesson_content = st.text_area(
        "What should be taught?",
        placeholder="Paste the necessary lesson, key concepts, examples, facts, formulas, or explanations."
    )
    competency = st.text_area("Learning competency (optional)")
    objectives = st.text_area("Learning objectives (optional)", placeholder="One objective per line. Leave blank for AI-generated objectives.")

    st.subheader("3. Assessment & Presentation")
    c3, c4, c5 = st.columns(3)
    with c3:
        assessment_type = st.selectbox("Assessment type", [
            "Mixed", "Multiple Choice", "True or False", "Identification",
            "Short Response", "Problem Solving", "Performance Task"
        ])
    with c4:
        quiz_count = st.number_input("Number of questions", 3, 10, 5)
    with c5:
        difficulty = st.selectbox("Difficulty", ["Easy", "Moderate", "Challenging", "Mixed"])

    include_visuals = st.checkbox("Include related pictures/graphics", value=True)
    extra = st.text_area("Additional teacher instructions (optional)")

    st.subheader("4. AI")
    use_ai = st.checkbox("Use Gemini AI for smarter lesson generation", value=True)
    try:
        saved_api_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        saved_api_key = ""
    if saved_api_key:
        st.success("Gemini API key is saved privately in Streamlit Secrets.")
        api_key = saved_api_key
    else:
        api_key = st.text_input("Gemini API key", type="password",
                                help="For a public app, save your key in Streamlit Secrets so users cannot see it.")
    model = st.selectbox("Gemini model", ["gemini-2.5-flash", "gemini-2.0-flash"])

    submitted = st.form_submit_button("Generate Lesson PowerPoint", use_container_width=True)

if submitted:
    if not teacher_name or not subject or not topic:
        st.error("Please enter the teacher name, subject, and lesson topic.")
        st.stop()

    with st.spinner("Creating your learner-centered lesson PowerPoint..."):
        lp_text = read_learning_plan(learning_plan)
        obj_list = [x.strip() for x in objectives.splitlines() if x.strip()] if objectives else []
        effective_strategy = custom_strategy.strip() if custom_strategy.strip() else strategy

        payload = {
            "teacher_name": teacher_name,
            "subject": subject,
            "grade_level": grade_level,
            "topic": topic,
            "strategy": effective_strategy,
            "learning_plan": lp_text[:18000],
            "lesson_content": lesson_content,
            "competency": competency,
            "objectives": obj_list,
            "assessment_type": assessment_type,
            "quiz_count": int(quiz_count),
            "difficulty": difficulty,
            "additional_instructions": extra,
            "font_rule": "Times New Roman, minimum 48 pt; split slides instead of shrinking text."
        }

        plan = None
        if use_ai and api_key:
            try:
                plan = generate_with_gemini(api_key.strip(), model, payload)
            except Exception as e:
                st.warning("AI generation failed, so the app used the built-in lesson generator instead.")
        if plan is None:
            plan = fallback_plan(payload)

        ppt = build_ppt(plan, teacher_name, subject, topic, grade_level, include_visuals)

    safe_topic = re.sub(r"[^A-Za-z0-9_-]+", "_", topic).strip("_") or "lesson"
    st.success("Your PowerPoint is ready.")
    st.download_button(
        "Download editable PowerPoint",
        data=ppt,
        file_name=f"{safe_topic}_lesson.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
    alignment = plan.get("alignment_summary", {})
    if alignment:
        with st.expander("Alignment summary"):
            st.write("**Source:**", alignment.get("source", "Not specified"))
            st.write("**Competency:**", alignment.get("competency", competency or "Generated from lesson"))
            st.write("**Strategy:**", alignment.get("strategy", effective_strategy))
            for obj in alignment.get("objectives", []):
                st.write("•", obj)

    with st.expander("Preview generated lesson structure"):
        for i, s in enumerate(plan["slides"], 1):
            st.markdown(f"**{i}. {s.get('title','Slide')}** — {s.get('section','')}")
