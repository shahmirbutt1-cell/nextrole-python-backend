from fastapi import FastAPI, UploadFile, File, Body
from docx import Document
import shutil
import os
import re
from openai import OpenAI

app = FastAPI()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


@app.get("/")
def root():
    return {"status": "NextRole Python backend running"}

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    full_text = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            full_text.append(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    full_text.append(cell.text)

    return "\n".join(full_text)

def get_section_paragraphs_universal(doc, section_aliases):
    """
    Detect section paragraphs in BOTH:
    - Table-based resumes
    - Standard paragraph resumes
    """

    # Normalize aliases
    section_aliases = [alias.upper() for alias in section_aliases]

    # --------------------------
    # 1️⃣ TABLE-BASED DETECTION
    # --------------------------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs = cell.paragraphs

                for i, paragraph in enumerate(paragraphs):
                    if paragraph.text.strip().upper() in section_aliases:

                        section_paragraphs = []
                        j = i + 1
                        while j < len(paragraphs) and paragraphs[j].text.strip():
                            section_paragraphs.append(paragraphs[j])
                            j += 1

                        if section_paragraphs:
                            return section_paragraphs

    # --------------------------
    # 2️⃣ PARAGRAPH-BASED DETECTION
    # --------------------------
    paragraphs = doc.paragraphs

    for i, paragraph in enumerate(paragraphs):
        if paragraph.text.strip().upper() in section_aliases:

            section_paragraphs = []
            j = i + 1
            while j < len(paragraphs) and paragraphs[j].text.strip():
                section_paragraphs.append(paragraphs[j])
                j += 1

            if section_paragraphs:
                return section_paragraphs

    return []

def replace_paragraph_text_preserve_style(paragraph, new_text):
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return

    first_run = paragraph.runs[0]
    font_name = first_run.font.name
    font_size = first_run.font.size
    bold = first_run.bold
    italic = first_run.italic

    for run in paragraph.runs:
        run.text = ""

    new_run = paragraph.add_run(new_text)
    new_run.font.name = font_name
    new_run.font.size = font_size
    new_run.bold = bold
    new_run.italic = italic

@app.post("/analyze-resume")
async def analyze_resume(file: UploadFile = File(...)):
    temp_file = f"temp_{file.filename}"

    with open(temp_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    extracted_text = extract_text_from_docx(temp_file)

    return {
        "file_name": file.filename,
        "text_length": len(extracted_text),
        "preview": extracted_text[:300],
        "full_text": extracted_text
    }


@app.post("/match-job")
async def match_job(data: dict = Body(...)):
    resume_text = data.get("resume_text", "").lower()
    job_description = data.get("job_description", "").lower()

    resume_words = set(re.findall(r'\w+', resume_text))
    job_words = set(re.findall(r'\w+', job_description))

    matching = resume_words.intersection(job_words)
    missing = job_words.difference(resume_words)

    if len(job_words) == 0:
        score = 0
    else:
        score = int(len(matching) / len(job_words) * 100)

    return {
        "match_score": score,
        "matching_skills": list(matching)[:10],
        "missing_skills": list(missing)[:10]
    }

    
from fastapi.responses import FileResponse
import uuid

@app.post("/tailor-resume-docx-preserve")
async def tailor_resume_docx_preserve(
    file: UploadFile = File(...),
    job_description: str = Body(...)
):
    input_filename = f"input_{uuid.uuid4()}.docx"
    output_filename = f"tailored_{uuid.uuid4()}.docx"

    # Save file
    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    doc = Document(input_filename)

    # --- PROFILE SECTION ---
    profile_paragraphs = get_section_paragraphs_universal(
    doc,
    ["PROFILE", "PROFESSIONAL SUMMARY", "SUMMARY", "OBJECTIVE"]
    )
    original_profile_lines = [p.text for p in profile_paragraphs]

    if original_profile_lines:
        profile_prompt = f"""
Rewrite each line below individually to better match the job description.

IMPORTANT:
- Return EXACTLY {len(original_profile_lines)} lines.
- Do not combine lines.
- Do not add extra lines.
- Keep similar length.

Lines:
{chr(10).join(original_profile_lines)}

Job Description:
{job_description}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert resume editor."},
                {"role": "user", "content": profile_prompt}
            ]
        )

        new_profile_lines = response.choices[0].message.content.strip().split("\n")

        for i in range(len(profile_paragraphs)):
            if i < len(new_profile_lines):
                replace_paragraph_text_preserve_style(
                    profile_paragraphs[i],
                    new_profile_lines[i]
                )

    # --- CORE COMPETENCIES SECTION ---
    skills_paragraphs = get_section_paragraphs_universal(
    doc,
    ["CORE COMPETENCIES", "SKILLS", "TECHNICAL SKILLS", "KEY SKILLS"]
    )
    original_skill_lines = [p.text for p in skills_paragraphs]

    if original_skill_lines:
        skills_prompt = f"""
Rewrite each skill line individually.

IMPORTANT:
- Return EXACTLY {len(original_skill_lines)} lines.
- Keep bullet-style structure.
- Do not combine lines.
- Do not add extra lines.

Lines:
{chr(10).join(original_skill_lines)}

Job Description:
{job_description}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert resume editor."},
                {"role": "user", "content": skills_prompt}
            ]
        )

        new_skill_lines = response.choices[0].message.content.strip().split("\n")

        for i in range(len(skills_paragraphs)):
            if i < len(new_skill_lines):
                replace_paragraph_text_preserve_style(
                    skills_paragraphs[i],
                    new_skill_lines[i]
                )

    doc.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Tailored_Resume.docx"
    )
