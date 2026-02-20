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

def get_experience_roles(doc):
    """
    Detect experience section and group bullet paragraphs under each job role.
    Returns list of roles with bullet paragraph objects.
    """

    roles = []
    in_experience = False
    current_role = None

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        # Detect experience header
        if text.upper() in ["WORK EXPERIENCE", "EXPERIENCE", "PROFESSIONAL EXPERIENCE"]:
            in_experience = True
            continue

        if in_experience:

            # Detect job title line (heuristic: contains company dash or bold text)
            if paragraph.runs and paragraph.runs[0].bold:
                if current_role:
                    roles.append(current_role)

                current_role = {
                    "title": text,
                    "bullets": []
                }

            # Detect bullet paragraph
            elif paragraph.style.name.startswith("List") or text.startswith("•") or text.startswith("-"):
                if current_role:
                    current_role["bullets"].append(paragraph)

    if current_role:
        roles.append(current_role)

    return roles

def extract_keywords(text):
    words = re.findall(r'\b\w+\b', text.lower())
    stopwords = {
        "the","and","with","for","from","that","this","have","has",
        "are","was","were","will","shall","your","their","about",
        "into","within","across","using","use","used"
    }
    return set(w for w in words if len(w) > 3 and w not in stopwords)

def role_relevance_score(role, job_keywords):
    role_text = " ".join([p.text for p in role["bullets"]]).lower()
    role_words = set(re.findall(r'\b\w+\b', role_text))

    overlap = role_words.intersection(job_keywords)

    if not role_words:
        return 0

    return len(overlap)

def detect_industry_keywords(text):
    finance_terms = {"finance","financial","accounting","budget","forecast","p&l","audit"}
    medical_terms = {"medical","device","orthopedic","trauma","surgical","hospital"}
    tech_terms = {"software","saas","cloud","api","engineering","data","ai"}

    text_words = set(re.findall(r'\b\w+\b', text.lower()))

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

@app.post("/debug-parse")
async def debug_parse(file: UploadFile = File(...)):
    input_filename = "temp.docx"
    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    parser = SemanticResumeParser(input_filename, openai_client=client)
    model = parser.parse()

    return model

@app.post("/tailor-resume-docx-preserve")
async def tailor_resume_docx_preserve(
    file: UploadFile = File(...),
    job_description: str = Body(...),
    mode: str = Body("balanced")  # "balanced" or "aggressive"
):
    input_filename = f"input_{uuid.uuid4()}.docx"
    output_filename = f"tailored_{uuid.uuid4()}.docx"

    # Save uploaded file
    with open(input_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    # Parse structure
    parser = SemanticResumeParser(input_filename, openai_client=client)
    resume_model = parser.parse()

    doc = Document(input_filename)

    # =============================
    # INDUSTRY + MODE DETECTION
    # =============================
    resume_text = extract_text_from_docx(input_filename)

    resume_industry = detect_industry_keywords(resume_text)
    job_industry = detect_industry_keywords(job_description)

    if max(resume_industry.values()) == 0 or max(job_industry.values()) == 0:
        industry_mismatch = False
    else:
        resume_primary = max(resume_industry, key=resume_industry.get)
        job_primary = max(job_industry, key=job_industry.get)
        industry_mismatch = resume_primary != job_primary

    if industry_mismatch:
        mode = "conservative"

    # =============================
    # HELPER: SAFE GPT REWRITE
    # =============================
    def safe_rewrite(original_lines, prompt_body):
        if not original_lines:
            return original_lines

        prompt = f"""
STRICT OUTPUT RULES:
- MODE: {mode}
- Do NOT add numbering
- Do NOT add prefixes like "Skill:" or "Improved:"
- Do NOT add section headers
- Do NOT fabricate metrics or achievements
- Do NOT change companies, roles, dates, or industries
- Keep EXACT same number of lines
- Preserve meaning and structure
- Return plain text only
- No extra commentary

Return EXACTLY {len(original_lines)} lines.

CONTENT TO REWRITE:
{chr(10).join(original_lines)}

JOB DESCRIPTION:
{job_description}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a professional resume optimization engine."},
                {"role": "user", "content": prompt}
            ]
        )

        output_lines = response.choices[0].message.content.strip().split("\n")

        # Safety guard
        if len(output_lines) != len(original_lines):
            return original_lines

        return output_lines

    # =============================
    # PROFILE SECTION
    # =============================
    profile_paragraphs = get_section_paragraphs_universal(
        doc,
        ["PROFILE", "PROFESSIONAL SUMMARY", "SUMMARY", "OBJECTIVE"]
    )

    if profile_paragraphs:
        original_lines = [p.text.strip() for p in profile_paragraphs]
        new_lines = safe_rewrite(original_lines, "Rewrite profile section.")

        for i, paragraph in enumerate(profile_paragraphs):
            replace_paragraph_text_preserve_style(paragraph, new_lines[i])

    # =============================
    # CORE COMPETENCIES
    # =============================
    skills_paragraphs = get_section_paragraphs_universal(
        doc,
        ["CORE COMPETENCIES", "SKILLS", "TECHNICAL SKILLS", "KEY SKILLS"]
    )

    if skills_paragraphs:
        original_lines = [p.text.strip() for p in skills_paragraphs]
        new_lines = safe_rewrite(original_lines, "Optimize skills section.")

        for i, paragraph in enumerate(skills_paragraphs):
            replace_paragraph_text_preserve_style(paragraph, new_lines[i])

    # =============================
    # EXPERIENCE RELEVANCE TARGETING
    # =============================
    roles = get_experience_roles(doc)

    if roles:
        job_keywords = extract_keywords(job_description)

        scored_roles = []
        for role in roles:
            score = role_relevance_score(role, job_keywords)
            scored_roles.append((role, score))

        scored_roles.sort(key=lambda x: x[1], reverse=True)

        RELEVANCE_THRESHOLD = 1 if mode == "balanced" else 3

        for role, score in scored_roles:

            if score < RELEVANCE_THRESHOLD:
                continue

            bullet_paragraphs = role["bullets"]
            original_bullets = [p.text.strip() for p in bullet_paragraphs]

            if not original_bullets:
                continue

            new_bullets = safe_rewrite(original_bullets, "Rewrite experience bullets.")

            for i, paragraph in enumerate(bullet_paragraphs):
                replace_paragraph_text_preserve_style(paragraph, new_bullets[i])

    # =============================
    # SAVE & RETURN
    # =============================
    doc.save(output_filename)

    return FileResponse(
        output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Tailored_Resume.docx"
    )
