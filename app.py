import streamlit as st
from pathlib import Path
import base64, io, json, re, tempfile, os, requests
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

ROOT = Path(__file__).parent
BG = ROOT / "assets" / "slidecraft_background.png"

st.set_page_config(page_title="SlideCraft-by-Ajelle", page_icon="✨", layout="wide",
                   initial_sidebar_state="expanded")

def data_uri(path):
    if not path.exists(): return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()

bg = data_uri(BG)

st.markdown(f"""
<style>
:root {{
 --p900:#421067; --p800:#57158a; --p700:#7120b4; --p600:#9136d9;
 --p300:#c6a5ee; --p100:#f7f1fc; --ink:#251b2e; --muted:#6e6079;
}}
.stApp {{
 background:
 linear-gradient(rgba(31,10,59,.25),rgba(17,27,78,.27)),
 url("{bg}") center/cover fixed;
}}
.block-container {{max-width:1240px;padding-top:1.2rem;padding-bottom:4rem}}
#MainMenu, footer {{visibility:hidden}}
.hero {{
 background:rgba(255,255,255,.95); border:1px solid #dfcdf2; border-radius:26px;
 padding:28px 34px; box-shadow:0 18px 50px rgba(34,10,64,.22); margin-bottom:18px
}}
.hero h1 {{margin:0;color:var(--p800);font-size:2.55rem;letter-spacing:-.03em}}
.hero p {{color:#594b65;font-size:1.05rem;line-height:1.55;margin:.6rem 0 0}}
.pill {{display:inline-block;margin-top:14px;padding:8px 14px;border-radius:999px;
 background:linear-gradient(135deg,#641bad,#9c31dd);color:white;font-weight:800;
 font-size:.8rem;letter-spacing:.08em}}
div[data-testid="stForm"] {{
 background:rgba(255,255,255,.97); border:1px solid #e5d6f4; border-radius:26px;
 padding:26px 28px 32px; box-shadow:0 18px 55px rgba(32,9,61,.22)
}}
.section {{
 background:#fbf8ff;border:1px solid #e8daf6;border-left:5px solid #7a25bb;
 border-radius:16px;padding:15px 18px;margin:12px 0 15px
}}
.section b {{color:#51147f;font-size:1.22rem}}
.section span {{display:block;color:#6e6079;margin-top:3px;font-size:.94rem}}
div[data-testid="stWidgetLabel"] p {{
 color:#321943!important;font-weight:750!important;font-size:.97rem!important
}}
.stTextInput input,.stTextArea textarea {{
 background:white!important;color:#24182e!important;border:1.5px solid #c7a5e9!important;
 border-radius:12px!important
}}
.stTextInput input::placeholder,.stTextArea textarea::placeholder {{
 color:#82718e!important;opacity:1!important
}}
div[data-baseweb="select"]>div {{
 background:white!important;color:#24182e!important;border:1.5px solid #c7a5e9!important;
 border-radius:12px!important
}}
div[data-baseweb="select"] span {{color:#24182e!important}}
div[data-testid="stFileUploader"] {{
 background:#fbf8ff;border:1.5px dashed #ae7cdb;border-radius:15px;padding:5px
}}
div[data-testid="stFileUploader"] * {{color:#342044!important}}
.stCheckbox label span {{color:#352047!important}}
.stButton>button,.stDownloadButton>button {{
 width:100%;min-height:54px;border:0!important;border-radius:14px!important;
 background:linear-gradient(135deg,#651cad,#982dd9)!important;color:white!important;
 font-weight:800!important;font-size:1rem!important;box-shadow:0 10px 24px rgba(101,28,173,.23)
}}
section[data-testid="stSidebar"] {{background:rgba(251,248,255,.97);border-right:1px solid #e3d2f3}}
section[data-testid="stSidebar"] * {{color:#351c4d!important}}
@media(max-width:900px){{.hero h1{{font-size:2rem}}div[data-testid="stForm"]{{padding:18px}}}}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ✨ SlideCraft")
    st.caption("by Ajelle")
    st.markdown("---")
    st.markdown("### Generate PPT")
    st.write("Create an editable, learner-centered lesson presentation.")
    st.markdown("---")
    st.markdown("### Quick Tips")
    st.write("Upload a Learning Plan when available.")
    st.write("Paste the exact lesson content you want taught.")
    st.write("Optional fields may be left blank for AI.")
    st.markdown("---")
    st.caption("Teach • Create • Inspire")

st.markdown("""<div class="hero"><h1>SlideCraft-by-Ajelle</h1>
<p>Turn your learning plan, lesson notes, and teaching ideas into an editable,
learner-centered PowerPoint presentation—aligned, readable, and ready for class.</p>
<div class="pill">TEACH ✦ CREATE ✦ INSPIRE</div></div>""", unsafe_allow_html=True)

def read_plan(upload):
    if upload is None: return ""
    raw, name = upload.read(), upload.name.lower()
    try:
        if name.endswith(".txt"): return raw.decode("utf-8", errors="ignore")
        if name.endswith(".pdf"):
            from pypdf import PdfReader
            return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(raw)).pages)
        if name.endswith(".docx"):
            from docx import Document
            d=Document(io.BytesIO(raw)); parts=[p.text.strip() for p in d.paragraphs if p.text.strip()]
            for table in d.tables:
                for row in table.rows:
                    vals=[c.text.strip() for c in row.cells if c.text.strip()]
                    if vals: parts.append(" | ".join(vals))
            return "\n".join(parts)
    except Exception: return ""
    return ""

def api_key():
    try: return st.secrets.get("GEMINI_API_KEY","")
    except Exception: return ""

def ai_plan(key, payload):
    url=f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"
    prompt="""You are an expert teacher, curriculum planner, assessment designer, and PowerPoint designer.
Create a learner-centered lesson presentation.

ALIGNMENT:
- If a Learning Plan is supplied, it is the PRIMARY alignment source.
- Respect its competency, objectives, sequence, content, activities, and assessment.
- Integrate teacher-entered lesson content accurately.
- If optional competency/objectives are blank, generate suitable aligned ones.

POWERPOINT:
- Cover: topic, subject, teacher.
- Times New Roman only; minimum 48 pt.
- 2-4 concise bullets per content slide. Split content instead of shrinking text.
- Learners should actively discuss, solve, investigate, observe, create, compare, practice, apply, or reflect.
- Follow the selected strategy.
- Include actual assessment questions and a separate answer key.
- Add real-world connections where appropriate.
- Suggest useful visual search phrases.

STRATEGIES:
I Do/We Do/You Do = model, guided practice, independent practice.
4E = Engage, Explore, Explain, Evaluate.
Gamification = learning missions/challenges/rounds.
Inquiry = question, investigate, evidence, explain.
Collaborative = structured pair/group work with output.
Problem-Based = authentic problem, investigate, solve, reflect.
Discussion-Based = prompt, think/pair/share, evidence, synthesis.

Return ONLY valid JSON:
{"alignment_summary":{"source":"Learning Plan|Teacher Inputs|Combined","competency":"string",
"objectives":["string","string","string"],"strategy":"string"},
"slides":[{"title":"string","section":"Cover|Objectives|Review|Engage|Explore|Explain|Practice|Application|Assessment|Answer Key|Closing|Other",
"bullets":["short bullet"],"teacher_note":"short cue","visual_query":"specific visual search phrase or empty","is_cover":false}]}

TEACHER INPUTS:
""" + json.dumps(payload, ensure_ascii=False)
    body={"contents":[{"parts":[{"text":prompt}]}],
          "generationConfig":{"temperature":.35,"responseMimeType":"application/json"}}
    r=requests.post(url,json=body,timeout=120); r.raise_for_status()
    return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])

def fallback(payload):
    t=payload["topic"]; c=payload["competency"] or f"Demonstrate understanding of the key concepts in {t}."
    objs=payload["objectives"] or [f"Explain the main ideas of {t}.",f"Apply {t} through guided practice.",
                                  "Demonstrate understanding through an assessment."]
    slides=[
      {"title":t,"section":"Cover","bullets":[],"teacher_note":"","visual_query":t,"is_cover":True},
      {"title":"Learning Competency","section":"Objectives","bullets":[c],"teacher_note":"","visual_query":"","is_cover":False},
      {"title":"Learning Objectives","section":"Objectives","bullets":objs[:3],"teacher_note":"","visual_query":"","is_cover":False},
      {"title":"Activate Prior Knowledge","section":"Review","bullets":[f"What do you already know about {t}?","Share one idea with a partner.","Connect it to today's lesson."],"teacher_note":"","visual_query":"","is_cover":False},
    ]
    s=payload["strategy"].lower()
    if "i do" in s:
        slides += [
          {"title":"I Do: Teacher Model","section":"Explain","bullets":[payload["lesson_content"][:220]],"teacher_note":"Model clearly and think aloud.","visual_query":t,"is_cover":False},
          {"title":"We Do: Guided Practice","section":"Practice","bullets":["Work through an example together.","Explain each step or idea.","Correct misconceptions together."],"teacher_note":"","visual_query":"","is_cover":False},
          {"title":"You Do: Independent Practice","section":"Application","bullets":["Complete a similar task independently.","Show your process or reasoning.","Check your work before submitting."],"teacher_note":"","visual_query":"","is_cover":False}]
    elif "4e" in s:
        slides += [
          {"title":"Engage","section":"Engage","bullets":[f"Observe an example related to {t}.","What do you notice?","What do you wonder?"],"teacher_note":"","visual_query":t,"is_cover":False},
          {"title":"Explore","section":"Explore","bullets":["Work with a partner or group.","Investigate the task.","Record evidence or observations."],"teacher_note":"","visual_query":"","is_cover":False},
          {"title":"Explain","section":"Explain","bullets":[payload["lesson_content"][:220]],"teacher_note":"Connect learner ideas to the concept.","visual_query":t,"is_cover":False}]
    elif "gamification" in s:
        slides += [
          {"title":"Mission Brief","section":"Engage","bullets":[f"Mission: master the key ideas in {t}.","Complete each learning challenge."],"teacher_note":"","visual_query":t,"is_cover":False},
          {"title":"Challenge 1","section":"Explore","bullets":["Study the clue or example.","Work with your team.","Explain your discovery."],"teacher_note":"","visual_query":"","is_cover":False},
          {"title":"Challenge 2","section":"Practice","bullets":["Apply the lesson idea.","Show evidence for your answer."],"teacher_note":"","visual_query":"","is_cover":False}]
    else:
        slides += [
          {"title":"Discover","section":"Engage","bullets":[f"Observe an example related to {t}.","What do you notice?","What questions arise?"],"teacher_note":"","visual_query":t,"is_cover":False},
          {"title":"Key Ideas","section":"Explain","bullets":[payload["lesson_content"][:220]],"teacher_note":"","visual_query":t,"is_cover":False},
          {"title":"Learner Activity","section":"Practice","bullets":["Work with a partner or group.","Apply the lesson idea.","Explain your answer or output."],"teacher_note":"","visual_query":"","is_cover":False},
          {"title":"Real-World Connection","section":"Application","bullets":[f"Where can we use {t} in daily life?","Give one practical example.","Explain why it matters."],"teacher_note":"","visual_query":t,"is_cover":False}]
    qs=[]; ans=[]
    for i in range(1,int(payload["quiz_count"])+1):
        qs.append(f"{i}. Answer one lesson-based question about {t}.")
        ans.append(f"{i}. Accept the correct answer supported by the lesson.")
    slides += [
      {"title":"Assessment","section":"Assessment","bullets":qs,"teacher_note":"","visual_query":"","is_cover":False},
      {"title":"Answer Key","section":"Answer Key","bullets":ans,"teacher_note":"","visual_query":"","is_cover":False},
      {"title":"Wrap-Up","section":"Closing","bullets":[f"Today I learned that {t}...","Share one real-world connection.","Ask one remaining question."],"teacher_note":"","visual_query":"","is_cover":False}]
    return {"alignment_summary":{"source":"Teacher Inputs","competency":c,"objectives":objs[:3],"strategy":payload["strategy"]},"slides":slides}

def commons_image(q):
    if not q:return None
    try:
        params={"action":"query","generator":"search","gsrsearch":q,"gsrnamespace":6,"gsrlimit":6,
                "prop":"imageinfo","iiprop":"url","iiurlwidth":900,"format":"json","origin":"*"}
        pages=requests.get("https://commons.wikimedia.org/w/api.php",params=params,timeout=15).json().get("query",{}).get("pages",{})
        for p in pages.values():
            info=(p.get("imageinfo") or [{}])[0]; url=info.get("thumburl") or info.get("url")
            if url:
                r=requests.get(url,timeout=15)
                if r.ok:return io.BytesIO(r.content)
    except Exception:pass
    return None

THEMES={
"Mathematics":(RGBColor(62,75,155),RGBColor(239,242,255),RGBColor(30,37,75)),
"Science":(RGBColor(44,113,89),RGBColor(238,248,244),RGBColor(26,65,53)),
"English":(RGBColor(116,61,141),RGBColor(248,240,250),RGBColor(64,34,76)),
"Filipino":(RGBColor(121,57,104),RGBColor(250,240,247),RGBColor(70,34,60)),
"Araling Panlipunan":(RGBColor(137,89,47),RGBColor(250,244,236),RGBColor(77,50,29)),
"TLE / ICT":(RGBColor(55,87,131),RGBColor(238,244,250),RGBColor(31,48,71)),
"MAPEH":(RGBColor(130,59,114),RGBColor(249,240,247),RGBColor(73,36,65)),
"Values / ESP":(RGBColor(98,64,141),RGBColor(245,240,250),RGBColor(55,37,77)),
"Other":(RGBColor(104,38,132),RGBColor(248,241,250),RGBColor(45,31,53))}
WHITE=RGBColor(255,255,255)

def runstyle(run,size,bold,color):
    run.font.name="Times New Roman";run.font.size=Pt(max(48,size));run.font.bold=bold;run.font.color.rgb=color

def textbox(slide,text,l,t,w,h,size,bold,color,align=PP_ALIGN.LEFT,valign=MSO_ANCHOR.TOP):
    b=slide.shapes.add_textbox(l,t,w,h);tf=b.text_frame;tf.clear();tf.word_wrap=True;tf.vertical_anchor=valign
    p=tf.paragraphs[0];p.alignment=align;r=p.add_run();r.text=text;runstyle(r,size,bold,color);return b

def chunks(s):
    bs=[re.sub(r"\s+"," ",str(x)).strip(" •-\n\t") for x in s.get("bullets",[]) if str(x).strip()]
    expanded=[]
    for b in bs:
        expanded += [b] if len(b)<=145 else ([x.strip() for x in re.split(r"(?<=[.!?;:])\s+",b) if x.strip()] or [b])
    n=3 if s.get("section") in ("Assessment","Answer Key") else 4
    out=[]
    for i in range(0,max(len(expanded),1),n):
        d=dict(s);d["bullets"]=expanded[i:i+n]
        if i:d["title"]=s.get("title","Lesson")+" (continued)";d["visual_query"]=""
        out.append(d)
    return out

def make_ppt(plan,teacher,subject,topic,grade,visuals):
    prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5);blank=prs.slide_layouts[6]
    accent,soft,ink=THEMES[subject]
    allslides=[]
    for s in plan.get("slides",[]):allslides.extend(chunks(s))
    for s in allslides:
        slide=prs.slides.add_slide(blank);slide.background.fill.solid();slide.background.fill.fore_color.rgb=soft
        strip=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(.18),Inches(7.5));strip.fill.solid();strip.fill.fore_color.rgb=accent;strip.line.fill.background()
        panel=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(.55),Inches(.45),Inches(12.1),Inches(6.5));panel.fill.solid();panel.fill.fore_color.rgb=WHITE;panel.line.color.rgb=accent;panel.line.transparency=65
        cover=s.get("is_cover") or s.get("section")=="Cover"
        if cover:
            textbox(slide,s.get("title",topic),Inches(1.1),Inches(1.45),Inches(11),Inches(1.6),60,True,accent,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
            sub=subject+(f" • {grade}" if grade else "")+f"\nTeacher: {teacher}"
            textbox(slide,sub,Inches(1.2),Inches(3.45),Inches(10.8),Inches(1.55),48,False,ink,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE);continue
        textbox(slide,s.get("title","Lesson"),Inches(.92),Inches(.68),Inches(11.3),Inches(.8),52,True,accent)
        img=commons_image(s.get("visual_query","")) if visuals and s.get("visual_query") else None
        tw=Inches(7.15) if img else Inches(11.05)
        box=slide.shapes.add_textbox(Inches(.95),Inches(1.72),tw,Inches(4.95));tf=box.text_frame;tf.clear();tf.word_wrap=True
        for i,b in enumerate(s.get("bullets",[])[:4]):
            p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=b;p.space_after=Pt(10)
            for r in p.runs:runstyle(r,48,False,ink)
        if img:
            try:
                im=Image.open(img).convert("RGB");tmp=tempfile.NamedTemporaryFile(suffix=".jpg",delete=False);im.save(tmp.name,"JPEG",quality=90)
                slide.shapes.add_picture(tmp.name,Inches(8.75),Inches(1.9),width=Inches(3.45),height=Inches(3.65));os.unlink(tmp.name)
            except Exception:pass
        if s.get("teacher_note"):
            try:slide.notes_slide.notes_text_frame.text="Teacher cue: "+s["teacher_note"]
            except Exception:pass
    out=io.BytesIO();prs.save(out);out.seek(0);return out

def section(num,title,helptext):
    st.markdown(f'<div class="section"><b>{num}. {title}</b><span>{helptext}</span></div>',unsafe_allow_html=True)

with st.form("form"):
    section(1,"Teacher Information","Add the class details that should appear in your presentation.")
    a,b=st.columns(2)
    with a:teacher=st.text_input("Teacher's Name *",placeholder="Example: Juan D. dela Cruz",help="This appears on the cover slide.")
    with b:grade=st.text_input("Grade Level / Section",placeholder="Example: Grade 8 - Rose",help="Optional.")

    section(2,"Lesson Details","Tell SlideCraft what you are teaching and how you want the lesson delivered.")
    a,b=st.columns(2)
    with a:
        topic=st.text_input("Lesson Topic / Title *",placeholder="Example: Solving Linear Equations in One Variable")
        competency=st.text_area("Learning Competency (Optional)",placeholder="Example: Illustrates and solves linear equations in one variable.",height=110,
                                help="Leave blank if your Learning Plan contains it or you want AI to generate one.")
    with b:
        subject=st.selectbox("Subject *",list(THEMES.keys()))
        strategy=st.selectbox("Teaching Strategy *",["I Do, We Do, You Do","4E: Engage, Explore, Explain, Evaluate","Gamification",
          "Real-World Connections","Inquiry-Based Learning","Collaborative Learning","Problem-Based Learning","Discussion-Based Learning"])
    objectives=st.text_area("Learning Objectives (Optional)",placeholder="One objective per line.\nExample:\nExplain the key concept.\nApply it correctly.\nConnect it to real life.",height=145,
                            help="Leave blank for three AI-generated aligned objectives.")

    section(3,"Lesson Content / Knowledge","Paste the concepts, formulas, facts, explanations, examples, or notes learners need to understand.")
    content=st.text_area("Content / Key Points *",placeholder="Example:\n- Definition\n- Important concepts\n- Formula or process\n- Worked examples\n- Common mistakes\n- Real-world application",height=220)
    extra=st.text_area("Additional Instructions (Optional)",placeholder="Example: Use simple language, include local examples, add a group activity, and provide an exit ticket.",height=115)

    section(4,"Upload Learning Plan (Optional)","If uploaded, this becomes SlideCraft's primary basis for lesson alignment.")
    upload=st.file_uploader("Learning Plan / Lesson Plan",type=["pdf","docx","txt"],help="Supported: PDF, DOCX, TXT.")

    section(5,"Assessment & Design","Choose the assessment and presentation preferences.")
    a,b,c=st.columns(3)
    with a:atype=st.selectbox("Assessment Type",["Mixed","Multiple Choice","True or False","Identification","Short Response","Problem Solving","Performance Task"])
    with b:qcount=st.number_input("Number of Questions",3,10,5)
    with c:difficulty=st.selectbox("Difficulty Level",["Easy","Moderate","Challenging","Mixed"])
    a,b=st.columns(2)
    with a:design=st.selectbox("Presentation Style",["Modern - Clean and Professional","Student-Friendly","Minimal Academic","Bright Classroom"])
    with b:colors=st.selectbox("Color Theme",["Auto-generate based on subject","Purple Academic","Soft Neutral"])
    visuals=st.checkbox("Include relevant lesson pictures / graphics",value=True)
    activities=st.checkbox("Include learner activity slides when appropriate",value=True)
    submit=st.form_submit_button("✨ Generate Lesson PowerPoint")

if submit:
    if not teacher.strip() or not topic.strip() or not content.strip():
        st.error("Please complete Teacher's Name, Lesson Topic / Title, and Content / Key Points.");st.stop()
    objs=[x.strip(" •-\t") for x in objectives.splitlines() if x.strip()][:5]
    payload={"teacher_name":teacher.strip(),"grade_level":grade.strip(),"topic":topic.strip(),"subject":subject,
      "strategy":strategy,"competency":competency.strip(),"objectives":objs,"lesson_content":content.strip(),
      "learning_plan":read_plan(upload)[:20000],"assessment_type":atype,"quiz_count":int(qcount),"difficulty":difficulty,
      "design_style":design,"color_theme":colors,"include_activity_slides":activities,"additional_instructions":extra.strip()}
    with st.spinner("SlideCraft is crafting your lesson presentation..."):
        plan=None;k=api_key()
        if k:
            try:plan=ai_plan(k,payload)
            except Exception:st.warning("AI was temporarily unavailable, so SlideCraft used its built-in generator.")
        else:st.info("Gemini is not configured in Streamlit Secrets, so the built-in generator is being used.")
        if plan is None:plan=fallback(payload)
        ppt=make_ppt(plan,teacher.strip(),subject,topic.strip(),grade.strip(),visuals)
    safe=re.sub(r"[^A-Za-z0-9_-]+","_",topic).strip("_") or "lesson"
    st.success("Your editable lesson PowerPoint is ready.")
    st.download_button("Download Lesson PowerPoint",ppt,f"{safe}_SlideCraft.pptx",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    with st.expander("View alignment summary"):
        al=plan.get("alignment_summary",{});st.write("**Source:**",al.get("source","Teacher Inputs"))
        st.write("**Competency:**",al.get("competency",competency or "Generated"));st.write("**Strategy:**",al.get("strategy",strategy))
    with st.expander("View generated slide structure"):
        for i,s in enumerate(plan.get("slides",[]),1):st.write(f"{i}. {s.get('title','Slide')} — {s.get('section','')}")
