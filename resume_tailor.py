rom typing import Dict, Any, List


class ResumeTailorEngine:

    def __init__(self, resume_model: Dict[str, Any], job_description: str, openai_client, mode="balanced"):
        self.model = resume_model
        self.job_description = job_description
        self.client = openai_client
        self.mode = mode

    # =========================================
    # PUBLIC ENTRY
    # =========================================

    def tailor(self):
        self._rewrite_summary()
        self._rewrite_skills()
        self._rewrite_experience()

    # =========================================
    # SAFE GPT CALL
    # =========================================

    def _safe_rewrite(self, original_lines: List[str]) -> List[str]:

        if not original_lines:
            return original_lines

        prompt = f"""
STRICT OUTPUT RULES:
- MODE: {self.mode}
- Do NOT add numbering
- Do NOT add prefixes like "Skill:" or "Revised:"
- Do NOT add section headers
- Do NOT fabricate metrics or achievements
- Do NOT change companies, roles, dates, or industries
- Keep EXACT same number of lines
- Preserve meaning
- Return plain text only
- No commentary

Return EXACTLY {len(original_lines)} lines.

CONTENT:
{chr(10).join(original_lines)}

JOB DESCRIPTION:
{self.job_description}
"""

        response = self.client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a professional resume optimization engine."},
                {"role": "user", "content": prompt}
            ]
        )

        output = response.choices[0].message.content.strip().split("\n")

        if len(output) != len(original_lines):
            return original_lines

        return output

    # =========================================
    # SUMMARY
    # =========================================

    def _rewrite_summary(self):

        summary_section = self.model.get("summary")

        if not summary_section:
            return

        original_text = summary_section["text"]

        new_text = self._safe_rewrite([original_text])[0]

        for p in summary_section["paragraphs"]:
            self._replace_paragraph_text_preserve_style(p["object"], new_text)

    # =========================================
    # SKILLS
    # =========================================

    def _rewrite_skills(self):

        skills = self.model.get("skills", [])

        if not skills:
            return

        original_lines = [s["text"] for s in skills]
        new_lines = self._safe_rewrite(original_lines)

        for i, skill in enumerate(skills):
            paragraph_obj = skill["paragraph"]["object"]
            self._replace_paragraph_text_preserve_style(paragraph_obj, new_lines[i])

    # =========================================
    # EXPERIENCE
    # =========================================

    def _rewrite_experience(self):

        experience_roles = self.model.get("experience", [])

        for role in experience_roles:

            bullets = role.get("bullets", [])

            if not bullets:
                continue

            original_lines = [b["text"] for b in bullets]
            new_lines = self._safe_rewrite(original_lines)

            for i, bullet in enumerate(bullets):
                self._replace_paragraph_text_preserve_style(
                    bullet["object"],
                    new_lines[i]
                )

    # =========================================
    # SAFE DOCX REPLACEMENT
    # =========================================

    def _replace_paragraph_text_preserve_style(self, paragraph, new_text):

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
