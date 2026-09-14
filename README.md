# Purple Lesson PPT Generator

A free-to-host Streamlit app that creates editable, learner-centered PowerPoint lessons.

## Main features
- Optional learning-plan upload: PDF, DOCX, or TXT
- Teacher can paste the lesson content to be taught
- Optional competency and objectives
- Teaching strategy choices:
  - I Do, We Do, You Do
  - 4E
  - Gamification
  - Real-World Connections
  - Inquiry-Based Learning
  - Collaborative Learning
  - Problem-Based Learning
  - Discussion-Based Learning
- Actual assessment/quiz content and answer key
- Times New Roman throughout
- 48 pt minimum text
- Automatically splits content rather than shrinking text
- Topic-related visuals from Wikimedia Commons
- Editable `.pptx` output
- Purple/white interface using your uploaded butterfly background
- Works in built-in mode even without an AI key
- Optional Gemini AI mode for stronger lesson generation

## Run on your computer
1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Run:
   pip install -r requirements.txt
4. Run:
   streamlit run app.py

## Free shareable deployment with Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload all files in this folder.
3. Go to Streamlit Community Cloud and create an app from the repository.
4. Set the main file to `app.py`.
5. Deploy.
6. Share the generated Streamlit link.

## Keeping the Gemini key private
For your personal/public deployment, do not place the key directly in the code.

Create `.streamlit/secrets.toml` in Streamlit Cloud only if you later modify the app to read:
GEMINI_API_KEY = "your-key"

This starter version also allows a teacher to paste a key privately in the interface. Streamlit password inputs are not displayed as plain text.

## Important note about "free"
The Streamlit app itself can be hosted on a free Community Cloud plan subject to Streamlit's current limits. Gemini and other AI providers can change their free quotas. The built-in non-AI generator remains usable without a paid API.


## v2 upgrades
- Learning-plan-first AI alignment
- DOCX table reading
- Alignment summary after generation
- Subject-aware PowerPoint accent colors
- Stricter slide splitting for the 48 pt minimum rule
- Assessment slides limited to fewer questions per slide for visibility
- Streamlit Secrets support for a private Gemini API key
- Grade/section shown on the cover when provided
