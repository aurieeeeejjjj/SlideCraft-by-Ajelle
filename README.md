[README.md](https://github.com/user-attachments/files/32177112/README.md)
# SlideCraft-by-Ajelle v6

This revision adds two modes:

1. **Full Lesson to PPT**
   - Required: Lesson Title, Year Level, Subject
   - Optional: Competency, objectives, lesson content, teaching resources, links
   - Teacher-provided content is preserved and arranged across enough slides to remain readable
   - Lesson text never goes below 45 pt
   - Assessment uses one item per slide
   - Assessment answers are stored in speaker notes

2. **Quiz to PPT Only**
   - Teacher pastes the quiz questions
   - One item per slide
   - Multiple-choice questions show the number, question, and choices, all left-aligned
   - Times New Roman
   - Starts at 45 pt and reduces only when a very long item must fit on one slide
   - Correct answers are stored in speaker notes

The presentation design is clean and creative without icons.

## Streamlit Secret

```toml
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```
