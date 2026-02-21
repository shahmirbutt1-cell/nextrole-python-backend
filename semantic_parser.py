import json
import re
from typing import Any, Dict, List, Union

from docx import Document


class SemanticResumeParser:
    def __init__(self, source: Union[str, Any], openai_client=None):
        self.client = openai_client
        self.paragraphs: List[Dict[str, Any]] = []
        self.sections: Dict[str, List[Dict[str, Any]]] = {}
        self.resume_model = {
            "summary": None,
            "skills": [],
            "experience": [],
            "education": []
        }

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
        self._extract_paragraphs()
        self._detect_sections()
        self._parse_summary()
        self._parse_skills()
        self._parse_experience()
        self._parse_education()
        return self.resume_model

    # =========================================
    # STEP 1 — EXTRACT PARAGRAPHS
    # =========================================

    def _extract_paragraphs(self):
        def capture_paragraph(p):
            return {
                "text": p.text.strip(),
                "style": p.style.name if p.style else "",
                "bold": any(run.bold for run in p.runs),
                "object": p
            }

        self.paragraphs = []

        for p in self.doc.paragraphs:
            if p.text.strip():
                self.paragraphs.append(capture_paragraph(p))

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

            if (
                len(text.split()) <= 6 and
                (p["bold"] or text.isupper())
            ):
                candidates.append(p)

        return candidates

    def _validate_headers_with_gpt(self, candidates):
        if not self.client or not candidates:
            return None

        header_texts = [p["text"] for p in candidates]

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
{chr(10).join(header_texts)}
"""

        response = self.client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You classify resume section headers."},
                {"role": "user", "content": prompt}
            ]
        )

        content = response.choices[0].message.content.strip()

        try:
            parsed = json.loads(content)
        except Exception:
            return None

        if not isinstance(parsed, dict):
            return None

        allowed = {"summary", "skills", "experience", "education", "none"}
        if any(label not in allowed for label in parsed.values()):
            return None

        return parsed

    def _classify_header_fallback(self, text: str):
        normalized = re.sub(r"[^a-z\s]", "", text.lower()).strip()

        rules = {
            "summary": [
                r"^summary$",
                r"^professional summary$",
                r"^profile$",
                r"^career summary$",
                r"^about$"
            ],
            "skills": [
                r"^skills$",
                r"^technical skills$",
                r"^core competencies$",
                r"^competencies$",
                r"^tools$"
            ],
            "experience": [
                r"^experience$",
                r"^work experience$",
                r"^professional experience$",
                r"^employment history$"
            ],
            "education": [
                r"^education$",
                r"^academic background$",
                r"^qualifications$"
            ]
        }

        for label, patterns in rules.items():
            if any(re.match(pattern, normalized) for pattern in patterns):
                return label

        return None

    def _detect_sections(self):
        section_map = {
            "summary": [],
            "skills": [],
            "experience": [],
            "education": []
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

                if label in section_map:
                    current_section = label
                else:
                    current_section = None

                continue

            if current_section:
                section_map[current_section].append(p)

        self.sections = section_map

    # =========================================
    # STEP 3 — SUMMARY
    # =========================================

    def _parse_summary(self):
        summary_paragraphs = self.sections.get("summary", [])

        if summary_paragraphs:
            combined = " ".join([p["text"] for p in summary_paragraphs])

            self.resume_model["summary"] = {
                "text": combined,
                "paragraphs": summary_paragraphs
            }

    # =========================================
    # STEP 4 — SKILLS
    # =========================================

    def _parse_skills(self):
        skills_paragraphs = self.sections.get("skills", [])

        for p in skills_paragraphs:
            self.resume_model["skills"].append({
                "text": p["text"],
                "paragraph": p
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
                    "bullets": [],
                    "header_paragraph": p
                }

                continue

            if self._is_bullet(p) and current_role:
                current_role["bullets"].append(p)

        if current_role:
            roles.append(current_role)

        self.resume_model["experience"] = roles

    def _is_role_header(self, paragraph):
        text = paragraph["text"]

        if paragraph["bold"]:
            return True

        if re.search(r"\b(19|20)\d{2}\b", text):
            return True

        if " - " in text or " – " in text:
            return True

        return False

    def _is_bullet(self, paragraph):
        text = paragraph["text"]
        style = paragraph["style"]

        if text.startswith(("•", "-", "–")):
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
                "paragraph": p
            })
