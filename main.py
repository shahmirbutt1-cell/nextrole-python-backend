import os
import re
import shutil

from docx import Document
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse
from openai import OpenAI

from resume_tailor import ResumeTailorEngine
from semantic_parser import SemanticResumeParser

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


def extract_keywords(text):
    words = re.findall(r"\b\w+\b", text.lower())
    stopwords = {
        "the", "and", "with", "for", "from", "that", "this", "have", "has",
        "are", "was", "were", "will", "shall", "your", "their", "about",
        "into", "within", "across", "using", "use", "used"
    }
    return set(w for w in words if len(w) > 3 and w not in stopwords)


def role_relevance_score(role, job_keywords):
    role_text = " ".join([p.text for p in role["bullets"]]).lower()
    role_words = set(re.findall(r"\b\w+\b", role_text))

    overlap = role_words.intersection(job_keywords)

    if not role_words:
        return 0

    return len(overlap)


def detect_industry_keywords(text):
    finance_terms = {"finance", "financial", "accounting", "budget", "forecast", "p&l", "audit"}
    medical_terms = {"medical", "device", "orthopedic", "trauma", "surgical", "hospital"}
    tech_terms = {"software", "saas", "cloud", "api", "engineering", "data", "ai"}

    text_words = set(re.findall(r"\b\w+\b", text.lower()))

    return {
        "finance": len(text_words.intersection(finance_terms)),
        "medical": len(text_words.intersection(medical_terms)),
        "tech": len(text_words.intersection(tech_terms))
    }


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
async def match_job(data: dict = Form(...)):
    resume_text = data.get("resume_text", "").lower()
    job_description = data.get("job_description", "").lower()

    resume_words = set(re.findall(r"\w+", resume_text))
    job_words = set(re.findall(r"\w+", job_description))

    matching = resume_words.intersection(job_words)
    missing = job_words.difference(resume_words)

    score = 0 if len(job_words) == 0 else int(len(matching) / len(job_words) * 100)

    return {
        "match_score": score,
        "matching_skills": list(matching)[:10],
        "missing_skills": list(missing)[:10]
    }


@app.post("/debug-parse")
async def debug_parse(file: UploadFile = File(...)):
    input_filename = "temp.docx"

    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    doc = Document(input_filename)
    parser = SemanticResumeParser(doc, openai_client=client)
    model = parser.parse()

    print("==== PARSE DEBUG ====")
    print("Summary:", model.get("summary"))
    print("Skills count:", len(model.get("skills", [])))
    print("Experience roles:", len(model.get("experience", [])))
    print("==== PARSE DEBUG END ====")

    return model


@app.post("/tailor-resume-docx-preserve")
async def tailor_resume_docx_preserve(
    file: UploadFile = File(...),
    job_description: str = Form(...),
    mode: str = Form("balanced")
):
    input_filename = "input.docx"
    output_filename = "tailored.docx"

    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Single Document instance for parse -> tailor -> save
    doc = Document(input_filename)

    parser = SemanticResumeParser(doc, openai_client=client)
    resume_model = parser.parse()

    print("==== PARSE DEBUG ====")
    print("Summary:", resume_model.get("summary"))
    print("Skills count:", len(resume_model.get("skills", [])))
    print("Experience roles:", len(resume_model.get("experience", [])))
    print("==== PARSE DEBUG END ====")

    tailor_engine = ResumeTailorEngine(
        resume_model=resume_model,
        job_description=job_description,
        openai_client=client,
        mode=mode
    )
    tailor_engine.tailor()

    doc.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Tailored_Resume.docx"
    )
