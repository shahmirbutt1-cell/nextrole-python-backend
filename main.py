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

    # Extract full text for AI
    full_text = extract_text_from_docx(input_filename)

    prompt = f"""
You are a professional resume editor.

Return your response in this EXACT format:

PROFILE:
<improved summary>

CORE COMPETENCIES:
<improved skills list>

Only rewrite those sections.
Do not change layout or structure.

Resume:
{full_text}

Job Description:
{job_description}
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert resume editor."},
            {"role": "user", "content": prompt}
        ]
    )

    tailored_text = response.choices[0].message.content

    # --- Parse Sections ---
    profile_text = ""
    skills_text = ""

    if "PROFILE:" in tailored_text:
        profile_text = tailored_text.split("PROFILE:")[1].split("CORE COMPETENCIES:")[0].strip()

    if "CORE COMPETENCIES:" in tailored_text:
        skills_text = tailored_text.split("CORE COMPETENCIES:")[1].strip()

    # --- Replace Sections in Tables ---
    replace_section_in_tables(doc, "PROFILE", profile_text)
    replace_section_in_tables(doc, "CORE COMPETENCIES", skills_text)

    doc.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Tailored_Resume.docx"
    )
