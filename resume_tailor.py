import logging
from typing import Any, Dict, List


class ResumeTailorEngine:
    """
    Rewrites resume sections to better match a job description, while
    preserving the original .docx formatting.

    Requires a resume_model produced by SemanticResumeParser and the
    parser's _paragraph_refs dict so it can write back changes to the
    live docx paragraph objects.

    Usage:
        engine = ResumeTailorEngine(
            resume_model=model,
            paragraph_refs=parser._paragraph_refs,
            job_description=jd,
            openai_client=client,
            mode="balanced"
        )
        engine.tailor()
        doc.save("tailored.docx")
    """

    def __init__(
        self,
        resume_model: Dict[str, Any],
        job_description: str,
        openai_client,
        paragraph_refs: Dict[int, Any],
        mode: str = "balanced",
    ):
        self.model = resume_model
        self.job_description = job_description
        self.client = openai_client
        self.refs = paragraph_refs   # ref_id (int) -> live docx paragraph object
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
        """
        Ask GPT to rewrite lines to better match the job description.

        Always returns the same number of lines as the input.
        Falls back to the original lines unchanged if the API call fails.
        """
        if not original_lines:
            return original_lines

        lines_joined = "\n".join(original_lines)

        prompt = f"""
STRICT OUTPUT RULES:
- MODE: {self.mode}
- Do NOT add numbering
- Do NOT add prefixes
- Do NOT add section headers
- Do NOT fabricate metrics
- Keep EXACT same number of lines
- Return plain text only
- No commentary

Return EXACTLY {len(original_lines)} lines.

CONTENT:
{lines_joined}

JOB DESCRIPTION:
{self.job_description}
"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a professional resume optimization engine."},
                    {"role": "user", "content": prompt}
                ],
                timeout=15
            )
        except Exception as e:
            logging.warning(f"GPT rewrite failed, returning original lines unchanged: {e}")
            return original_lines

        raw_output = response.choices[0].message.content.strip()

        output = [
            line.strip()
            for line in raw_output.split("\n")
            if line.strip()
        ]

        # Enforce exact line count
        if len(output) < len(original_lines):
            output += original_lines[len(output):]
        elif len(output) > len(original_lines):
            output = output[:len(original_lines)]

        return output

    # =========================================
    # SUMMARY
    # =========================================

    def _rewrite_summary(self):
        summary_section = self.model.get("summary")
        if not summary_section:
            return

        paragraphs = summary_section.get("paragraphs", [])
        if not paragraphs:
            return

        # Rewrite each paragraph individually — not all with one shared string
        original_lines = [p["text"] for p in paragraphs]
        new_lines = self._safe_rewrite(original_lines)

        for i, p in enumerate(paragraphs):
            doc_paragraph = self.refs.get(p["ref_id"])
            if doc_paragraph:
                self._replace_paragraph_text_preserve_style(doc_paragraph, new_lines[i])

    # =========================================
    # SKILLS
    # =========================================

    def _rewrite_skills(self):
        skills = self.model.get("skills", [])
        if not skills:
            return

        # Deduplicate by ref_id to avoid rewriting the same paragraph twice
        unique_skills = []
        seen_ref_ids = set()

        for skill in skills:
            ref_id = skill["ref_id"]
            if ref_id in seen_ref_ids:
                continue
            seen_ref_ids.add(ref_id)
            unique_skills.append(skill)

        original_lines = [s["text"] for s in unique_skills]
        new_lines = self._safe_rewrite(original_lines)

        for i, skill in enumerate(unique_skills):
            doc_paragraph = self.refs.get(skill["ref_id"])
            if doc_paragraph:
                self._replace_paragraph_text_preserve_style(doc_paragraph, new_lines[i])

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
                doc_paragraph = self.refs.get(bullet["ref_id"])
                if doc_paragraph:
                    self._replace_paragraph_text_preserve_style(doc_paragraph, new_lines[i])

    # =========================================
    # CORE FORMATTING-PRESERVING REPLACEMENT
    # =========================================

    def _replace_paragraph_text_preserve_style(self, paragraph, new_text: str):
        """
        Replace a paragraph's text while preserving original run formatting.

        Three cases:

        1. No runs — add a single new run with no special styling.

        2. Single run — replace .text in place. All run properties
           (bold, italic, font name, size, color) are completely untouched.

        3. Multiple runs — distribute the new text proportionally across
           existing runs based on each run's share of the original total
           character count. Each run keeps its own formatting properties.

           Example:
             Original: ["Acme Corp " (bold, 10 chars), "2019-2022" (regular, 9 chars)]
             Total: 19 chars.
             New text: "Globex Inc 2021-Present" (23 chars)
             Run 0 weight: 10/19 = 53%  -> gets 12 chars -> "Globex Inc  "
             Run 1 weight:  9/19 = 47%  -> gets 11 chars -> "2021-Present"
             (last run absorbs remainder to prevent rounding loss)

        Note: proportional distribution is a heuristic. When GPT changes
        the length significantly, word boundaries shift. But bold/italic/
        color/font are preserved on the right structural segments (title vs
        date, company vs role) which is far better than flattening everything
        to a single style.
        """
        runs = paragraph.runs

        # Case 1: no runs
        if not runs:
            paragraph.add_run(new_text)
            return

        # Case 2: single run — simplest path, swap text in place
        if len(runs) == 1:
            runs[0].text = new_text
            return

        # Case 3: multiple runs — proportional distribution
        original_total = sum(len(r.text) for r in runs)

        if original_total == 0:
            runs[0].text = new_text
            for r in runs[1:]:
                r.text = ""
            return

        new_total = len(new_text)
        chars_assigned = 0

        for i, run in enumerate(runs):
            is_last = (i == len(runs) - 1)

            if is_last:
                # Last run gets everything remaining — no rounding loss
                run.text = new_text[chars_assigned:]
            else:
                weight = len(run.text) / original_total
                chars = round(weight * new_total)
                # Guarantee later runs each get at least 1 char slot
                max_chars = new_total - chars_assigned - (len(runs) - i - 1)
                chars = max(0, min(chars, max_chars))
                run.text = new_text[chars_assigned: chars_assigned + chars]
                chars_assigned += chars
