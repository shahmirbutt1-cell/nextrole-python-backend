from fastapi import FastAPI, UploadFile, File
from docx import Document
import shutil
import os
import openai

app = FastAPI()

openai.api_key = os.getenv('OPENAI_API_KEY')

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
from fastapi import Body
import re

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


@app.post("/tailor-resume")
async def tailor_resume(data: dict = Body(...)):
    resume_text = data.get("resume_text", "")
    job_description = data.get("job_description", "")

    prompt = f"""
    You are an expert resume editor.

    Keep formatting structure intact.
    Only improve:
    - Professional Summary
    - Skills section
    - Relevant bullet points

    Resume:
    {resume_text}

    Job Description:
    {job_description}

    Return the full improved resume text.
    """

    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert resume editor."},
            {"role": "user", "content": prompt}
        ]
    )

    tailored_text = response.choices[0].message["content"]

    return {
        "tailored_resume": tailored_text
    }
