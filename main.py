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

def replace_section_in_tables(doc, section_title, new_content):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs = cell.paragraphs

                for i, paragraph in enumerate(paragraphs):
                    if section_title.upper() in paragraph.text.upper():
                        
                        # Clear existing content below header
                        j = i + 1
                        while j < len(paragraphs) and paragraphs[j].text.strip():
                            paragraphs[j].text = ""
                            j += 1
                        
                        # Insert new content
                        lines = new_content.split("\n")
                        for k, line in enumerate(lines):
                            if i + 1 + k < len(paragraphs):
                                paragraphs[i + 1 + k].text = line
                        
                        return
def extract_section_from_tables(doc, section_title):
    """
    Extract text under a section header inside table cells.
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs = cell.paragraphs

                for i, paragraph in enumerate(paragraphs):
                    if section_title.upper() in paragraph.text.upper():
                        section_lines = []
                        j = i + 1
                        while j < len(paragraphs) and paragraphs[j].text.strip():
                            section_lines.append(paragraphs[j].text)
                            j += 1
                        return "\n".join(section_lines)
    return ""


def replace_section_preserve_format(doc, section_title, new_text):
    """
    Replace section content while preserving formatting & layout.
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs = cell.paragraphs

                for i, paragraph in enumerate(paragraphs):
                    if section_title.upper() in paragraph.text.upper():

                        # Identify existing section paragraphs
                        j = i + 1
                        section_paragraphs = []
                        while j < len(paragraphs) and paragraphs[j].text.strip():
                            section_paragraphs.append(paragraphs[j])
                            j += 1

                        new_lines = new_text.split("\n")

                        for idx, p in enumerate(section_paragraphs):
                            if idx < len(new_lines):
                                replace_paragraph_preserve_runs(p, new_lines[idx])
                            else:
                                p.text = ""

                        return


def replace_paragraph_preserve_runs(paragraph, new_text):
    """
    Replace paragraph text but preserve run formatting.
    """
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return

    first_run = paragraph.runs[0]
    font_name = first_run.font.name
    font_size = first_run.font.size
    bold = first_run.bold
    italic = first_run.italic

    # Clear existing runs
    for run in paragraph.runs:
        run.text = ""

    # Insert new run with preserved formatting
    new_run = paragraph.add_run(new_text)
    new_run.font.name = font_name
    new_run.font.size = font_size
    new_run.bold = bold
    new_run.italic = italic
    
from fastapi.responses import FileResponse
import uuid

@app.post("/tailor-resume-docx-preserve")
async def tailor_resume_docx_preserve(
    file: UploadFile = File(...),
    job_description: str = Body(...)
):
    input_filename = f"input_{uuid.uuid4()}.docx"
    output_filename = f"tailored_{uuid.uuid4()}.docx"

    # Save uploaded file
    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    doc = Document(input_filename)

    # Extract only relevant sections
    profile_text = extract_section_from_tables(doc, "PROFILE")
    skills_text = extract_section_from_tables(doc, "CORE COMPETENCIES")

    # --- AI CALLS PER SECTION (Controlled & Precise) ---

    # Rewrite Profile
    profile_prompt = f"""
Rewrite this professional summary to better match the job description.
Keep similar length and tone. Do not invent experience.

Summary:
{profile_text}

Job Description:
{job_description}
"""

    profile_response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert resume editor."},
            {"role": "user", "content": profile_prompt}
        ]
    )

    improved_profile = profile_response.choices[0].message.content.strip()

    # Rewrite Skills
    skills_prompt = f"""
Improve these skills to better match the job description.
Keep formatting as bullet-style lines.
Do not invent skills not relevant.

Skills:
{skills_text}

Job Description:
{job_description}
"""

    skills_response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert resume editor."},
            {"role": "user", "content": skills_prompt}
        ]
    )

    improved_skills = skills_response.choices[0].message.content.strip()

    # --- Inject Back While Preserving Formatting ---

    replace_section_preserve_format(doc, "PROFILE", improved_profile)
    replace_section_preserve_format(doc, "CORE COMPETENCIES", improved_skills)

    doc.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Tailored_Resume.docx"
    )
