"""
resume_tailor.py

Mode behaviour (now actually enforced, not just passed as a label):

  conservative  Rewrites bullets for the most recent role only.
                Summary and skills are left unchanged.
                Use when you want minimal, targeted changes.

  balanced      Rewrites bullets for ALL roles.
                Summary is rewritten. Skills are left unchanged.
                Default — good for most use cases.

  aggressive    Rewrites bullets for ALL roles.
                Summary AND skills are both rewritten.
                Use when the resume needs heavy alignment to the JD.

Role order:
  Most resumes list newest job first (reverse-chronological).
  The parser captures roles in document reading order, so roles[0]
  is assumed to be the most recent. A date-based check confirms this
  and reverses the list if needed so roles[0] is always newest.
"""

import logging
import re
from typing import Any, Dict, List, Optional

VALID_MODES = {"conservative", "balanced", "aggressive"}

# Regex to find a year in a role title string
_YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")


def _extract_year(text: str) -> Optional[int]:
    """Return the first 4-digit year found in text, or None."""
    m = _YEAR_RE.search(text or "")
    return int(m.group()) if m else None


class ResumeTailorEngine:
    """
    Rewrite resume sections to better match a job description while
    preserving original .docx formatting.

    Parameters
    ----------
    resume_model : dict
        Output of SemanticResumeParser.parse()
    paragraph_refs : dict
        SemanticResumeParser._paragraph_refs — maps ref_id → live docx paragraph
    job_description : str
        The target job description text
    openai_client :
        Initialised OpenAI client
    mode : str
        One of 'conservative', 'balanced' (default), 'aggressive'
        Controls which sections are rewritten and how many roles are touched.
    newest_role_first : bool
        Set to False if the resume lists oldest job first.
        Default True (reverse-chronological — most common format).
    """

    def __init__(
        self,
        resume_model: Dict[str, Any],
        job_description: str,
        openai_client,
        paragraph_refs: Dict[int, Any],
        mode: str = "balanced",
        newest_role_first: bool = True,
    ):
        if mode not in VALID_MODES:
            logging.warning(
                f"Unknown mode '{mode}'. Defaulting to 'balanced'. "
                f"Valid options: {VALID_MODES}"
            )
            mode = "balanced"

        self.model      = resume_model
        self.jd         = job_description
        self.client     = openai_client
        self.refs       = paragraph_refs
        self.mode       = mode
        self.newest_role_first = newest_role_first

    # ──────────────────────────────────────────────────────────────────────────
    # PUBLIC ENTRY
    # ──────────────────────────────────────────────────────────────────────────

    def tailor(self) -> Dict[str, Any]:
        """
        Rewrite the resume according to self.mode.

        Returns a summary dict describing what was changed:
            {
              "mode": str,
              "roles_rewritten": int,
              "summary_rewritten": bool,
              "skills_rewritten": bool,
            }
        """
        roles = self._get_roles_in_order()
        report = {
            "mode":              self.mode,
            "roles_available":   len(roles),
            "roles_rewritten":   0,
            "summary_rewritten": False,
            "skills_rewritten":  False,
        }

        if self.mode == "conservative":
            # Only the most recent role's bullets
            roles_to_rewrite = roles[:1]
            rewrite_summary  = False
            rewrite_skills   = False

        elif self.mode == "balanced":
            # All roles + summary
            roles_to_rewrite = roles
            rewrite_summary  = True
            rewrite_skills   = False

        else:  # aggressive
            # Everything
            roles_to_rewrite = roles
            rewrite_summary  = True
            rewrite_skills   = True

        for role in roles_to_rewrite:
            changed = self._rewrite_role_bullets(role)
            if changed:
                report["roles_rewritten"] += 1

        if rewrite_summary:
            report["summary_rewritten"] = self._rewrite_summary()

        if rewrite_skills:
            report["skills_rewritten"] = self._rewrite_skills()

        logging.info(f"Tailor report: {report}")
        return report

    # ──────────────────────────────────────────────────────────────────────────
    # ROLE ORDERING
    # ──────────────────────────────────────────────────────────────────────────

    def _get_roles_in_order(self) -> List[Dict[str, Any]]:
        """
        Return experience roles with the most recent role first.

        Detects the actual ordering from dates in role titles so that
        'conservative' mode always hits the most recent job, regardless
        of how the resume was laid out.
        """
        roles = self.model.get("experience", [])
        if len(roles) < 2:
            return roles

        # Try to infer order from years in title strings
        years = [_extract_year(r.get("title", "")) for r in roles]
        years_found = [y for y in years if y is not None]

        if len(years_found) >= 2:
            # If first role has the highest year → already newest-first
            if years[0] and years[-1] and years[0] < years[-1]:
                logging.info(
                    "Role order appears oldest-first — reversing for processing "
                    f"({years[0]} → {years[-1]})"
                )
                return list(reversed(roles))

        # Fall back to the caller's hint
        return roles if self.newest_role_first else list(reversed(roles))

    # ──────────────────────────────────────────────────────────────────────────
    # SAFE GPT CALL
    # ──────────────────────────────────────────────────────────────────────────

    def _safe_rewrite(self, original_lines: List[str], context: str = "") -> List[str]:
        """
        Ask GPT to rewrite lines to better match the job description.

        Always returns the same number of lines as the input.
        Falls back to original lines unchanged on any API failure.

        Parameters
        ----------
        original_lines : list of str
        context : str
            Optional label for the prompt (e.g. 'summary', 'skills',
            'experience bullets — Software Engineer at Acme Corp').
            Helps GPT produce more targeted rewrites.
        """
        if not original_lines:
            return original_lines

        # Mode-specific instruction injected into the prompt
        mode_instructions = {
            "conservative": (
                "Make minimal, targeted changes. Only adjust keywords and phrasing "
                "to better reflect the job description. Preserve the original meaning "
                "and structure as closely as possible."
            ),
            "balanced": (
                "Rewrite to better align with the job description. You may rephrase "
                "and adjust terminology, but do not fabricate new achievements or "
                "drastically change the meaning."
            ),
            "aggressive": (
                "Rewrite aggressively to maximise alignment with the job description. "
                "Use the JD's language and keywords throughout. You may restructure "
                "sentences significantly, but do not fabricate metrics or experiences "
                "that aren't implied by the originals."
            ),
        }

        lines_block = "\n".join(original_lines)
        context_str = f"\nCONTEXT: {context}" if context else ""

        prompt = f"""You are a professional resume optimisation engine.{context_str}

MODE INSTRUCTION: {mode_instructions[self.mode]}

STRICT RULES:
- Return EXACTLY {len(original_lines)} lines — no more, no fewer
- Do NOT add numbering, prefixes, or section headers
- Do NOT fabricate specific metrics, names, or dates
- Return plain text only — no markdown, no commentary

LINES TO REWRITE:
{lines_block}

JOB DESCRIPTION:
{self.jd}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a professional resume optimisation engine. "
                            "Return only the rewritten lines, nothing else."
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
                timeout=20,
            )
        except Exception as e:
            logging.warning(
                f"GPT rewrite failed for [{context or 'unknown'}], "
                f"returning originals: {e}"
            )
            return original_lines

        raw = response.choices[0].message.content.strip()
        output = [line.strip() for line in raw.split("\n") if line.strip()]

        # Enforce exact line count
        if len(output) < len(original_lines):
            logging.warning(
                f"GPT returned {len(output)} lines for {len(original_lines)} inputs "
                f"[{context}] — padding with originals"
            )
            output += original_lines[len(output):]
        elif len(output) > len(original_lines):
            logging.warning(
                f"GPT returned {len(output)} lines for {len(original_lines)} inputs "
                f"[{context}] — truncating"
            )
            output = output[:len(original_lines)]

        return output

    # ──────────────────────────────────────────────────────────────────────────
    # SUMMARY
    # ──────────────────────────────────────────────────────────────────────────

    def _rewrite_summary(self) -> bool:
        """Rewrite summary paragraphs. Returns True if any paragraph was changed."""
        section = self.model.get("summary")
        if not section:
            return False

        paragraphs = section.get("paragraphs", [])
        if not paragraphs:
            return False

        original_lines = [p["text"] for p in paragraphs]
        new_lines = self._safe_rewrite(original_lines, context="professional summary")

        changed = False
        for i, p in enumerate(paragraphs):
            doc_para = self.refs.get(p["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                changed = True

        return changed

    # ──────────────────────────────────────────────────────────────────────────
    # SKILLS
    # ──────────────────────────────────────────────────────────────────────────

    def _rewrite_skills(self) -> bool:
        """Rewrite skills paragraphs. Returns True if any paragraph was changed."""
        skills = self.model.get("skills", [])
        if not skills:
            return False

        # Deduplicate by ref_id
        unique, seen = [], set()
        for s in skills:
            if s["ref_id"] not in seen:
                seen.add(s["ref_id"])
                unique.append(s)

        original_lines = [s["text"] for s in unique]
        new_lines = self._safe_rewrite(original_lines, context="skills and competencies")

        changed = False
        for i, skill in enumerate(unique):
            doc_para = self.refs.get(skill["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                changed = True

        return changed

    # ──────────────────────────────────────────────────────────────────────────
    # EXPERIENCE
    # ──────────────────────────────────────────────────────────────────────────

    def _rewrite_role_bullets(self, role: Dict[str, Any]) -> bool:
        """
        Rewrite the bullet points for a single role.
        Returns True if at least one bullet was changed.
        """
        bullets = role.get("bullets", [])
        if not bullets:
            return False

        title   = role.get("title", "this role")
        context = f"experience bullets — {title}"

        original_lines = [b["text"] for b in bullets]
        new_lines = self._safe_rewrite(original_lines, context=context)

        changed = False
        for i, bullet in enumerate(bullets):
            doc_para = self.refs.get(bullet["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                changed = True

        return changed

    # ──────────────────────────────────────────────────────────────────────────
    # CORE FORMATTING-PRESERVING REPLACEMENT
    # ──────────────────────────────────────────────────────────────────────────

    def _replace_paragraph_text_preserve_style(self, paragraph, new_text: str):
        """
        Replace a paragraph's text while preserving original run formatting.

        Three cases:
          1. No runs      → add a single new run.
          2. Single run   → replace .text in place; all formatting untouched.
          3. Multi-run    → distribute new text proportionally across existing
                            runs so each run keeps its own bold/italic/font/
                            color/size. The last run absorbs any remainder.
        """
        runs = paragraph.runs

        if not runs:
            paragraph.add_run(new_text)
            return

        if len(runs) == 1:
            runs[0].text = new_text
            return

        # Multi-run: proportional distribution
        original_total = sum(len(r.text) for r in runs)

        if original_total == 0:
            runs[0].text = new_text
            for r in runs[1:]:
                r.text = ""
            return

        new_total      = len(new_text)
        chars_assigned = 0

        for i, run in enumerate(runs):
            if i == len(runs) - 1:
                run.text = new_text[chars_assigned:]
            else:
                weight    = len(run.text) / original_total
                chars     = round(weight * new_total)
                max_chars = new_total - chars_assigned - (len(runs) - i - 1)
                chars     = max(0, min(chars, max_chars))
                run.text  = new_text[chars_assigned: chars_assigned + chars]
                chars_assigned += chars
