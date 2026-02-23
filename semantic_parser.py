import json
import logging
import re
from typing import Any, Dict, List, Union

from docx import Document


# Compiled once at module load — not rebuilt on every method call
_HEADER_RULES: Dict[str, List[str]] = {
    "summary": [
        r"^summary$",
        r"^professional summary$",
        r"^profile$",
        r"^career summary$",
        r"^about$",
        r"^objective$",
        r"^career objective$",
    ],
    "skills": [
        r"^skills$",
        r"^technical skills$",
        r"^core competencies$",
        r"^competencies$",
        r"^tools$",
        r"^technologies$",
        r"^areas of expertise$",
    ],
    "experience": [
        r"^experience$",
        r"^work experience$",
        r"^professional experience$",
        r"^employment history$",
        r"^career history$",
        r"^work history$",
    ],
    "education": [
        r"^education$",
        r"^academic background$",
        r"^qualifications$",
        r"^academic qualifications$",
        r"^degrees$",
    ],
}


class SemanticResumeParser:
    """
    Parses a .docx resume into structured data.

    Accepts either a file path string or a pre-loaded python-docx Document
    object. Uses GPT for section header classification when an OpenAI client
    is provided, falling back to regex rules otherwise.

    Usage:
        parser = SemanticResumeParser("resume.docx", openai_client=client)
        model = parser.parse()
        # model keys: summary, skills, experience, education

    Returns a fully JSON-serializable dict. Raw paragraph objects are kept
    separately in self._paragraph_refs for use by ResumeTailorEngine.
    """

    def __init__(self, source: Union[str, Any], openai_client=None):
        self.client = openai_client
        self.paragraphs: List[Dict[str, Any]] = []
        self.sections: Dict[str, List[Dict[str, Any]]] = {}

        # JSON-serializable output model
        self.resume_model: Dict[str, Any] = {
            "summary": None,
            "skills": [],
            "experience": [],
            "education": [],
        }

        # Internal store of raw docx paragraph objects keyed by paragraph id()
        # Used by ResumeTailorEngine to write back changes. Not exposed in
        # resume_model so the model stays JSON-serializable.
        self._paragraph_refs: Dict[int, Any] = {}

        if hasattr(source, "paragraphs") and hasattr(source, "tables"):
            self.file_path = None
            self.doc = source
        else:
            self.file_path = source
            self.doc = Document(source)

    # =========================================
    # PUBLIC ENTRY
    # =========================================

    def parse(self) -> Dict[str, Any]:
        """
        Parse the loaded resume into structured data.

        Returns:
            dict with keys: summary, skills, experience, education.
            All values are JSON-serializable. Raw docx paragraph objects
            are stored in self._paragraph_refs (keyed by id) for the
            tailor engine.
        """
        if self.resume_model.get("_parsed"):
            return self.resume_model

        self._extract_paragraphs()
        self._detect_sections()
        self._parse_summary()
        self._parse_skills()
        self._parse_experience()
        self._parse_education()

        self.resume_model["_parsed"] = True
        return self.resume_model

    # =========================================
    # STEP 1 — EXTRACT PARAGRAPHS
    # =========================================

    def _extract_paragraphs(self):
        """
        Extract all non-empty paragraphs from the document.

        Builds a set of table-cell paragraph objects first to avoid
        double-counting: python-docx includes table cell paragraphs in
        doc.paragraphs, so iterating doc.tables separately would duplicate them.
        """
        def capture_paragraph(p):
            ref_id = id(p)
            self._paragraph_refs[ref_id] = p
            return {
                "text": p.text.strip(),
                "style": p.style.name if p.style else "",
                "bold": any(run.bold for run in p.runs),
                "ref_id": ref_id,   # key into self._paragraph_refs
            }

        # Collect all table-cell paragraph objects to avoid duplicating them
        table_paragraph_objects = {
            p
            for table in self.doc.tables
            for row in table.rows
            for cell in row.cells
            for p in cell.paragraphs
        }

        self.paragraphs = []

        # Main body paragraphs — skip any that live inside tables
        for p in self.doc.paragraphs:
            if p.text.strip() and p not in table_paragraph_objects:
                self.paragraphs.append(capture_paragraph(p))

        # Table cell paragraphs — add these separately in reading order
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text.strip():
                            self.paragraphs.append(capture_paragraph(p))

    # =========================================
    # STEP 2 — SECTION DETECTION
    # =========================================

    def _identify_header_candidates(self):
        candidates = []
        for p in self.paragraphs:
            text = p["text"]
            if len(text.split()) <= 6 and (p["bold"] or text.isupper()):
                candidates.append(p)
        return candidates

    def _validate_headers_with_gpt(self, candidates):
        """
        Ask GPT to classify each candidate header into a section category.
        Returns a dict of {header_text: category} or None on failure.
        """
        if not self.client or not candidates:
            return None

        header_texts = [p["text"] for p in candidates]
        headers_joined = "\n".join(header_texts)

        prompt = f"""
You are a resume structure classification engine.

Classify each line below into ONE of these categories:

- summary
- skills
- experience
- education
- none

Return ONLY valid JSON in this format:

{{
  "Header Text": "category"
}}

Headers:
{headers_joined}
"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You classify resume section headers."},
                    {"role": "user", "content": prompt}
                ],
                timeout=10
            )
        except Exception as e:
            logging.warning(f"GPT header classification request failed: {e}")
            return None

        content = response.choices[0].message.content.strip()

        try:
            parsed = json.loads(content)
        except json.JSONDecodeError as e:
            logging.warning(
                f"GPT returned invalid JSON for header classification: {e}. "
                f"Raw response (first 300 chars): {content[:300]}"
            )
            return None

        if not isinstance(parsed, dict):
            logging.warning(f"GPT classification result was not a dict: {type(parsed)}")
            return None

        allowed = {"summary", "skills", "experience", "education", "none"}
        invalid = [label for label in parsed.values() if label not in allowed]
        if invalid:
            logging.warning(f"GPT returned unexpected category labels: {invalid}. Falling back to regex.")
            return None

        return parsed

    def _classify_header_fallback(self, text: str):
        """Classify a header using regex rules. Returns category string or None."""
        normalized = re.sub(r"[^a-z\s]", "", text.lower()).strip()

        for label, patterns in _HEADER_RULES.items():
            if any(re.match(pattern, normalized) for pattern in patterns):
                return label

        return None

    def _detect_sections(self):
        section_map = {
            "summary": [],
            "skills": [],
            "experience": [],
            "education": [],
        }

        candidates = self._identify_header_candidates()
        gpt_classification = self._validate_headers_with_gpt(candidates)

        if not gpt_classification:
            gpt_classification = {
                p["text"]: self._classify_header_fallback(p["text"]) or "none"
                for p in candidates
            }

        current_section = None

        for p in self.paragraphs:
            text = p["text"]

            if text in gpt_classification:
                label = gpt_classification[text]
                current_section = label if label in section_map else None
                continue

            if current_section:
                section_map[current_section].append(p)

        self.sections = section_map

    # =========================================
    # STEP 3 — SUMMARY
    # =========================================

    def _parse_summary(self):
        summary_paragraphs = self.sections.get("summary", [])
        if not summary_paragraphs:
            return

        combined = " ".join([p["text"] for p in summary_paragraphs])
        self.resume_model["summary"] = {
            "text": combined,
            # Store only serializable data; ref_ids link back to _paragraph_refs
            "paragraphs": [
                {"text": p["text"], "ref_id": p["ref_id"]}
                for p in summary_paragraphs
            ],
        }

    # =========================================
    # STEP 4 — SKILLS
    # =========================================

    def _parse_skills(self):
        skills_paragraphs = self.sections.get("skills", [])
        for p in skills_paragraphs:
            self.resume_model["skills"].append({
                "text": p["text"],
                "ref_id": p["ref_id"],
            })

    # =========================================
    # STEP 5 — EXPERIENCE
    # =========================================

    def _parse_experience(self):
        experience_paragraphs = self.sections.get("experience", [])

        roles = []
        current_role = None

        for p in experience_paragraphs:
            text = p["text"]

            if self._is_role_header(p):
                if current_role:
                    roles.append(current_role)
                current_role = {
                    "title": text,
                    "title_ref_id": p["ref_id"],
                    "bullets": [],
                }
                continue

            if self._is_bullet(p) and current_role:
                current_role["bullets"].append({
                    "text": p["text"],
                    "ref_id": p["ref_id"],
                })

        if current_role:
            roles.append(current_role)

        self.resume_model["experience"] = roles

    def _is_role_header(self, paragraph):
        """
        Determine whether a paragraph is a job role header.

        Bullets are explicitly excluded first — a bold bullet is not a header.
        Date patterns and em/en-dash separators (common in job title lines)
        are the primary positive signals.
        """
        # Bullets are never role headers — check this first
        if self._is_bullet(paragraph):
            return False

        text = paragraph["text"]

        # Year pattern is the strongest signal (e.g. "Acme Corp  2019 – 2022")
        if re.search(r"\b(19|20)\d{2}\b", text):
            return True

        # Dash separator + bold (e.g. "Software Engineer – Acme Corp")
        if (" - " in text or " – " in text) and paragraph["bold"]:
            return True

        return False

    def _is_bullet(self, paragraph):
        text = paragraph["text"]
        style = paragraph["style"]

        if text.startswith(("•", "-", "–", "*")):
            return True

        if style.startswith("List"):
            return True

        return False

    # =========================================
    # STEP 6 — EDUCATION
    # =========================================

    def _parse_education(self):
        education_paragraphs = self.sections.get("education", [])
        for p in education_paragraphs:
            self.resume_model["education"].append({
                "text": p["text"],
                "ref_id": p["ref_id"],
            })
