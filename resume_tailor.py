"""
resume_tailor.py

Mode behaviour:
  conservative  → most recent role bullets only. No summary/skills.
  balanced      → all roles + summary. No skills. (default)
  aggressive    → all roles + summary + skills.

Bullet balancing (new):
  After rewriting, any bullet that is more than 2x the median length
  of its siblings is sent to GPT to be split into multiple concise
  bullets. The overlong paragraph is replaced with the first split
  bullet and new paragraphs are inserted after it for the rest,
  inheriting the original bullet's style and formatting.
"""

import copy
import logging
import re
from typing import Any, Dict, List, Optional

from docx.oxml.ns import qn

VALID_MODES = {"conservative", "balanced", "aggressive"}

# A bullet is considered overlong if it is this many times longer than
# the median length of its siblings in the same role.
_OVERLONG_RATIO   = 2.0

# Hard minimum: always flag bullets longer than this many chars,
# even if the median is also large.
_OVERLONG_MIN_CHARS = 200

# Target character count per bullet (used in the GPT split prompt).
_TARGET_BULLET_CHARS = 100

_YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")


def _extract_year(text: str) -> Optional[int]:
    m = _YEAR_RE.search(text or "")
    return int(m.group()) if m else None


def _median_int(values: List[int]) -> float:
    if not values:
        return 0
    s = sorted(values)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


# ─────────────────────────────────────────────────────────────────────────────

class ResumeTailorEngine:
    """
    Rewrite resume sections to better match a job description while
    preserving original .docx formatting.

    Parameters
    ----------
    resume_model    : dict  — output of SemanticResumeParser.parse()
    paragraph_refs  : dict  — SemanticResumeParser._paragraph_refs
    job_description : str
    openai_client   :       — initialised OpenAI client
    mode            : str   — 'conservative' | 'balanced' | 'aggressive'
    newest_role_first: bool — True if resume lists newest job first (default)
    balance_bullets : bool  — split overlong bullets after rewriting (default True)
    """

    def __init__(
        self,
        resume_model:     Dict[str, Any],
        job_description:  str,
        openai_client,
        paragraph_refs:   Dict[int, Any],
        mode:             str  = "balanced",
        newest_role_first: bool = True,
        balance_bullets:  bool = True,
    ):
        if mode not in VALID_MODES:
            logging.warning(f"Unknown mode '{mode}'. Defaulting to 'balanced'.")
            mode = "balanced"

        self.model             = resume_model
        self.jd                = job_description
        self.client            = openai_client
        self.refs              = paragraph_refs
        self.mode              = mode
        self.newest_role_first = newest_role_first
        self.balance_bullets   = balance_bullets

    # ──────────────────────────────────────────────────────────────────────────
    # PUBLIC ENTRY
    # ──────────────────────────────────────────────────────────────────────────

    def tailor(self) -> Dict[str, Any]:
        """
        Rewrite the resume according to self.mode, then balance bullet lengths.

        Returns a report dict:
          mode, roles_available, roles_rewritten,
          summary_rewritten, skills_rewritten, bullets_split
        """
        roles = self._get_roles_in_order()

        report = {
            "mode":              self.mode,
            "roles_available":   len(roles),
            "roles_rewritten":   0,
            "summary_rewritten": False,
            "skills_rewritten":  False,
            "bullets_split":     0,
        }

        if self.mode == "conservative":
            roles_to_rewrite = roles[:1]
            rewrite_summary  = False
            rewrite_skills   = False
        elif self.mode == "balanced":
            roles_to_rewrite = roles
            rewrite_summary  = True
            rewrite_skills   = False
        else:  # aggressive
            roles_to_rewrite = roles
            rewrite_summary  = True
            rewrite_skills   = True

        for role in roles_to_rewrite:
            if self._rewrite_role_bullets(role):
                report["roles_rewritten"] += 1

        if rewrite_summary:
            report["summary_rewritten"] = self._rewrite_summary()

        if rewrite_skills:
            report["skills_rewritten"] = self._rewrite_skills()

        # ── Bullet balancing pass ─────────────────────────────────────────
        if self.balance_bullets:
            for role in roles_to_rewrite:
                report["bullets_split"] += self._balance_role_bullets(role)

        logging.info(f"Tailor report: {report}")
        return report

    # ──────────────────────────────────────────────────────────────────────────
    # ROLE ORDERING
    # ──────────────────────────────────────────────────────────────────────────

    def _get_roles_in_order(self) -> List[Dict[str, Any]]:
        """Return roles with most recent first, auto-detecting order from dates."""
        roles = self.model.get("experience", [])
        if len(roles) < 2:
            return roles

        years = [_extract_year(r.get("title", "")) for r in roles]
        years_found = [y for y in years if y is not None]

        if len(years_found) >= 2 and years[0] and years[-1]:
            if years[0] < years[-1]:
                logging.info(f"Oldest-first order detected ({years[0]}→{years[-1]}) — reversing.")
                return list(reversed(roles))

        return roles if self.newest_role_first else list(reversed(roles))

    # ──────────────────────────────────────────────────────────────────────────
    # SAFE GPT CALL
    # ──────────────────────────────────────────────────────────────────────────

    def _safe_rewrite(self, original_lines: List[str], context: str = "") -> List[str]:
        """
        Rewrite lines to better match the JD. Always returns same line count.
        Falls back to originals on any API failure.
        """
        if not original_lines:
            return original_lines

        mode_instructions = {
            "conservative": (
                "Make minimal, targeted changes. Only adjust keywords and phrasing "
                "to better reflect the job description. Preserve the original meaning "
                "and structure as closely as possible."
            ),
            "balanced": (
                "Rewrite to better align with the job description. Rephrase and adjust "
                "terminology, but do not fabricate new achievements or drastically "
                "change the meaning."
            ),
            "aggressive": (
                "Rewrite aggressively to maximise alignment with the job description. "
                "Use the JD's language and keywords throughout. Restructure sentences "
                "significantly if needed, but do not fabricate metrics or experiences "
                "that aren't implied by the originals."
            ),
        }

        context_str = f"\nCONTEXT: {context}" if context else ""
        lines_block = "\n".join(original_lines)

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
                    {"role": "system", "content": "You are a professional resume optimisation engine. Return only the rewritten lines."},
                    {"role": "user",   "content": prompt},
                ],
                timeout=20,
            )
        except Exception as e:
            logging.warning(f"GPT rewrite failed [{context}], keeping originals: {e}")
            return original_lines

        raw    = response.choices[0].message.content.strip()
        output = [line.strip() for line in raw.split("\n") if line.strip()]

        if len(output) < len(original_lines):
            logging.warning(f"GPT returned {len(output)}/{len(original_lines)} lines [{context}] — padding")
            output += original_lines[len(output):]
        elif len(output) > len(original_lines):
            logging.warning(f"GPT returned {len(output)}/{len(original_lines)} lines [{context}] — truncating")
            output = output[:len(original_lines)]

        return output

    # ──────────────────────────────────────────────────────────────────────────
    # SUMMARY / SKILLS / EXPERIENCE REWRITES
    # ──────────────────────────────────────────────────────────────────────────

    def _rewrite_summary(self) -> bool:
        section = self.model.get("summary")
        if not section:
            return False
        paragraphs     = section.get("paragraphs", [])
        original_lines = [p["text"] for p in paragraphs]
        new_lines      = self._safe_rewrite(original_lines, context="professional summary")
        changed = False
        for i, p in enumerate(paragraphs):
            doc_para = self.refs.get(p["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                changed = True
        return changed

    def _rewrite_skills(self) -> bool:
        skills = self.model.get("skills", [])
        if not skills:
            return False
        unique, seen = [], set()
        for s in skills:
            if s["ref_id"] not in seen:
                seen.add(s["ref_id"]); unique.append(s)
        original_lines = [s["text"] for s in unique]
        new_lines      = self._safe_rewrite(original_lines, context="skills and competencies")
        changed = False
        for i, skill in enumerate(unique):
            doc_para = self.refs.get(skill["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                changed = True
        return changed

    def _rewrite_role_bullets(self, role: Dict[str, Any]) -> bool:
        bullets = role.get("bullets", [])
        if not bullets:
            return False
        title          = role.get("title", "this role")
        original_lines = [b["text"] for b in bullets]
        new_lines      = self._safe_rewrite(original_lines, context=f"experience bullets — {title}")
        changed = False
        for i, bullet in enumerate(bullets):
            doc_para = self.refs.get(bullet["ref_id"])
            if doc_para and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_para, new_lines[i])
                # Update the model text so the balancing pass sees current text
                bullet["text"] = new_lines[i]
                changed = True
        return changed

    # ──────────────────────────────────────────────────────────────────────────
    # BULLET BALANCING
    # ──────────────────────────────────────────────────────────────────────────

    def _balance_role_bullets(self, role: Dict[str, Any]) -> int:
        """
        Find overlong bullets in a role and split them into concise ones.

        A bullet is overlong if:
          - Its length is > _OVERLONG_RATIO × median sibling length, AND
          - Its length is > _OVERLONG_MIN_CHARS

        The overlong paragraph is replaced in-place with the first split
        bullet. Additional split bullets are inserted as new paragraphs
        immediately after, copying the original paragraph's XML structure
        (so they inherit list indentation, bullet character, and spacing).

        Returns the number of bullets that were split.
        """
        bullets = role.get("bullets", [])
        if len(bullets) < 2:
            return 0

        lengths = [len(b["text"]) for b in bullets]
        med     = _median_int(lengths)

        splits_done = 0

        # Iterate in reverse so that inserting new paragraphs after a bullet
        # doesn't shift the indices of earlier bullets we haven't processed yet.
        for i in reversed(range(len(bullets))):
            bullet = bullets[i]
            text   = bullet["text"]
            length = len(text)

            if length <= _OVERLONG_MIN_CHARS:
                continue
            if med > 0 and length / med <= _OVERLONG_RATIO:
                continue

            # This bullet is overlong — ask GPT to split it
            split_lines = self._gpt_split_bullet(text, context=role.get("title", ""))
            if not split_lines or len(split_lines) <= 1:
                continue  # GPT couldn't improve it — leave as-is

            doc_para = self.refs.get(bullet["ref_id"])
            if not doc_para:
                continue

            # Replace the existing paragraph with the first split line
            self._replace_paragraph_text_preserve_style(doc_para, split_lines[0])

            # Insert remaining split lines as new paragraphs after doc_para
            for extra_text in reversed(split_lines[1:]):
                new_para = self._insert_paragraph_after(doc_para, extra_text)
                if new_para:
                    # Register the new paragraph so future ref lookups work
                    new_ref_id = id(new_para)
                    self.refs[new_ref_id] = new_para

            splits_done += 1
            logging.info(
                f"Split overlong bullet ({length} chars) into "
                f"{len(split_lines)} bullets for role: {role.get('title','')}"
            )

        return splits_done

    def _gpt_split_bullet(self, text: str, context: str = "") -> List[str]:
        """
        Ask GPT to split one overlong bullet into multiple concise ones.

        Returns a list of bullet strings (no bullet characters — plain text).
        Returns an empty list on failure so the caller can skip gracefully.
        """
        estimated_count = max(2, round(len(text) / _TARGET_BULLET_CHARS))
        context_str = f" for the role: {context}" if context else ""

        prompt = f"""You are editing a resume{context_str}.

The following bullet point is too long and needs to be split into {estimated_count} separate, concise bullet points.

RULES:
- Each bullet should be roughly {_TARGET_BULLET_CHARS} characters (1-2 lines when printed)
- Do NOT use bullet characters, numbers, or prefixes — plain text only
- Do NOT fabricate new information — only use what is in the original
- Do NOT add a preamble or commentary
- Return exactly {estimated_count} lines, one bullet per line

ORIGINAL BULLET:
{text}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You split resume bullets into concise lines. Return only the lines."},
                    {"role": "user",   "content": prompt},
                ],
                timeout=15,
            )
        except Exception as e:
            logging.warning(f"GPT bullet split failed: {e}")
            return []

        raw   = response.choices[0].message.content.strip()
        lines = [l.strip().lstrip("•-–*›◦▪▸○→✓✔►0123456789.) ") for l in raw.split("\n") if l.strip()]

        # Sanity: reject if GPT returned just one very long line (didn't split)
        if len(lines) == 1 and len(lines[0]) > _OVERLONG_MIN_CHARS:
            return []

        return lines if lines else []

    # ──────────────────────────────────────────────────────────────────────────
    # DOCX PARAGRAPH INSERTION
    # ──────────────────────────────────────────────────────────────────────────

    def _insert_paragraph_after(self, reference_paragraph, text: str):
        """
        Insert a new paragraph immediately after reference_paragraph in the
        document, copying its XML structure so it inherits bullet formatting,
        list indentation, style, and spacing.

        Returns the new python-docx Paragraph object, or None on failure.
        """
        try:
            from docx.text.paragraph import Paragraph as DocxParagraph

            # Deep-copy the reference paragraph's XML element
            ref_elem = reference_paragraph._element
            new_elem = copy.deepcopy(ref_elem)

            # Clear all run text in the copy, then set the new text in
            # the first run (or add one if there are none)
            runs_in_new = new_elem.findall(qn("w:r"))
            if runs_in_new:
                # Set first run's text, clear the rest
                first_t = runs_in_new[0].find(qn("w:t"))
                if first_t is None:
                    first_t = runs_in_new[0].makeelement(qn("w:t"), {})
                    runs_in_new[0].append(first_t)
                first_t.text = text
                # Preserve whitespace if needed
                if text and (text[0] == " " or text[-1] == " "):
                    first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                # Remove extra runs
                for extra_run in runs_in_new[1:]:
                    new_elem.remove(extra_run)
            else:
                # No runs in the copy — build a minimal run
                r_elem = new_elem.makeelement(qn("w:r"), {})
                t_elem = new_elem.makeelement(qn("w:t"), {})
                t_elem.text = text
                r_elem.append(t_elem)
                new_elem.append(r_elem)

            # Insert into the document XML immediately after the reference
            ref_elem.addnext(new_elem)

            # Wrap in a python-docx Paragraph object and return it
            return DocxParagraph(new_elem, reference_paragraph._p.getparent()
                                 if hasattr(reference_paragraph._p, 'getparent') else None)

        except Exception as e:
            logging.warning(f"Failed to insert paragraph after reference: {e}")
            return None

    # ──────────────────────────────────────────────────────────────────────────
    # CORE FORMATTING-PRESERVING REPLACEMENT
    # ──────────────────────────────────────────────────────────────────────────

    def _replace_paragraph_text_preserve_style(self, paragraph, new_text: str):
        """
        Replace paragraph text while preserving run-level formatting.

        1. No runs      → add a single run.
        2. Single run   → replace .text in place.
        3. Multi-run    → distribute new text proportionally across runs
                          so each run keeps its bold/italic/font/color/size.
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
