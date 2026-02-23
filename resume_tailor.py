"""
resume_tailor.py

Critical fix in this version
─────────────────────────────
_replace_paragraph_text_preserve_style now separates every paragraph's runs
into two buckets BEFORE writing anything:

  drawing_runs  — runs that contain w:drawing, AlternateContent (WPS shapes,
                  grouped objects), or w:object elements. These are NEVER
                  touched. They carry background shapes, sidebar decorations,
                  column dividers, and other layout graphics that are anchored
                  to paragraphs as empty runs.

  text_runs     — all other runs. Only these have their .text replaced.

Proportional distribution (for multi-run paragraphs) is calculated purely
from text-run character counts. Drawing runs are completely invisible to the
distribution logic.

Without this separation the old proportional loop wrote new text directly
into drawing runs, overwriting the AlternateContent XML and destroying every
embedded graphic in any paragraph it touched.

Mode behaviour (unchanged)
───────────────────────────
  conservative  → most recent role bullets only. No summary/skills.
  balanced      → all roles + summary. (default)
  aggressive    → all roles + summary + skills.
"""

import copy
import logging
import re
from typing import Any, Dict, List, Optional

from docx.oxml.ns import qn

VALID_MODES = {"conservative", "balanced", "aggressive"}

_OVERLONG_RATIO      = 2.0
_OVERLONG_MIN_CHARS  = 200
_TARGET_BULLET_CHARS = 100

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

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
# Drawing-run detection
# ─────────────────────────────────────────────────────────────────────────────

def _is_drawing_run(run) -> bool:
    """
    Return True if this run holds a drawing, shape, or embedded object.

    Such runs must NEVER have their text modified. They contain:
      - w:drawing  — inline or anchored images / SmartArt / charts
      - mc:AlternateContent  — WPS shapes, grouped objects, background fills
        (the most common cause of layout corruption in designed resumes)
      - w:object  — OLE embedded objects
    """
    e = run._element
    if e.find(qn("w:drawing")) is not None:
        return True
    if e.find(f"{{{_MC_NS}}}AlternateContent") is not None:
        return True
    if e.find(qn("w:object")) is not None:
        return True
    return False


def _paragraph_has_drawings(paragraph) -> bool:
    """Return True if any run in the paragraph is a drawing run."""
    return any(_is_drawing_run(r) for r in paragraph.runs)


def _paragraph_is_narrow_column(paragraph) -> bool:
    """
    Return True if the paragraph sits in a narrow/sidebar column.
    Signal: center-aligned AND large left indent.
    Used to force keyword-only rewrites that never expand compact content.
    """
    try:
        pPr = paragraph._element.find(qn("w:pPr"))
        if pPr is not None:
            jc = pPr.find(qn("w:jc"))
            if jc is not None and jc.get(qn("w:val"), "").lower() == "center":
                ind = paragraph.paragraph_format.left_indent
                if ind and int(ind) > 100000:
                    return True
    except Exception:
        pass
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Engine
# ─────────────────────────────────────────────────────────────────────────────

class ResumeTailorEngine:
    """
    Rewrite resume sections to better match a job description, preserving
    all original .docx formatting including embedded drawings, floating
    shapes, background fills, and multi-column sidebar layouts.

    Parameters
    ----------
    resume_model      : dict   output of SemanticResumeParser.parse()
    paragraph_refs    : dict   SemanticResumeParser._paragraph_refs
    job_description   : str
    openai_client     :        initialised OpenAI client
    mode              : str    'conservative' | 'balanced' | 'aggressive'
    newest_role_first : bool   True if newest job appears first in the doc
    balance_bullets   : bool   split overlong bullets after rewriting
    """

    def __init__(
        self,
        resume_model:      Dict[str, Any],
        job_description:   str,
        openai_client,
        paragraph_refs:    Dict[int, Any],
        mode:              str  = "balanced",
        newest_role_first: bool = True,
        balance_bullets:   bool = True,
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
        Rewrite the resume and return a report of what changed.
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

        if self.balance_bullets:
            for role in roles_to_rewrite:
                report["bullets_split"] += self._balance_role_bullets(role)

        logging.info(f"Tailor report: {report}")
        return report

    # ──────────────────────────────────────────────────────────────────────────
    # ROLE ORDERING
    # ──────────────────────────────────────────────────────────────────────────

    def _get_roles_in_order(self) -> List[Dict[str, Any]]:
        roles = self.model.get("experience", [])
        if len(roles) < 2:
            return roles
        years = [_extract_year(r.get("title", "")) for r in roles]
        years_found = [y for y in years if y is not None]
        if len(years_found) >= 2 and years[0] and years[-1] and years[0] < years[-1]:
            logging.info("Oldest-first role order detected — reversing.")
            return list(reversed(roles))
        return roles if self.newest_role_first else list(reversed(roles))

    # ──────────────────────────────────────────────────────────────────────────
    # GPT REWRITE
    # ──────────────────────────────────────────────────────────────────────────

    def _safe_rewrite(
        self,
        original_lines: List[str],
        context:        str  = "",
        keyword_only:   bool = False,
    ) -> List[str]:
        """
        Ask GPT to rewrite lines to better match the job description.
        Always returns the same number of lines as input.
        Falls back to originals on any API failure.

        keyword_only=True forces compact rewrites — used for sidebar content
        that must not be expanded into prose.
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
                "terminology freely, but do not fabricate new achievements or drastically "
                "change the meaning."
            ),
            "aggressive": (
                "Rewrite aggressively to maximise alignment with the job description. "
                "Use the JD's language and keywords throughout. Do not fabricate metrics "
                "or experiences not implied by the originals."
            ),
        }

        if keyword_only:
            instruction = (
                "IMPORTANT: This content appears in a narrow column or sidebar. "
                "Keep each output line roughly the same LENGTH as its input line. "
                "Only substitute relevant keywords — do NOT expand into sentences or prose."
            )
        else:
            instruction = mode_instructions[self.mode]

        context_str = f"\nCONTEXT: {context}" if context else ""
        lines_block = "\n".join(original_lines)

        prompt = f"""You are a professional resume optimisation engine.{context_str}

MODE INSTRUCTION: {instruction}

STRICT RULES:
- Return EXACTLY {len(original_lines)} lines — no more, no fewer
- Do NOT add numbering, bullet characters, prefixes, or section headers
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
                        "content": "You are a professional resume optimisation engine. "
                                   "Return only the rewritten lines, nothing else.",
                    },
                    {"role": "user", "content": prompt},
                ],
                timeout=20,
            )
        except Exception as e:
            logging.warning(f"GPT rewrite failed [{context}]: {e}")
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
    # SECTION REWRITES
    # ──────────────────────────────────────────────────────────────────────────

    def _rewrite_summary(self) -> bool:
        section = self.model.get("summary")
        if not section:
            return False
        paragraphs = section.get("paragraphs", [])
        if not paragraphs:
            return False

        keyword_only = any(
            (doc_p := self.refs.get(p["ref_id"])) is not None and
            (_paragraph_is_narrow_column(doc_p) or _paragraph_has_drawings(doc_p))
            for p in paragraphs
        )

        original_lines = [p["text"] for p in paragraphs]
        new_lines = self._safe_rewrite(original_lines, "professional summary", keyword_only)

        changed = False
        for i, p in enumerate(paragraphs):
            doc_p = self.refs.get(p["ref_id"])
            if doc_p and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_p, new_lines[i])
                changed = True
        return changed

    def _rewrite_skills(self) -> bool:
        skills = self.model.get("skills", [])
        if not skills:
            return False

        unique, seen = [], set()
        for s in skills:
            if s["ref_id"] not in seen:
                seen.add(s["ref_id"])
                unique.append(s)

        keyword_only = any(
            (doc_p := self.refs.get(s["ref_id"])) is not None and
            (_paragraph_is_narrow_column(doc_p) or _paragraph_has_drawings(doc_p))
            for s in unique
        )

        original_lines = [s["text"] for s in unique]
        new_lines = self._safe_rewrite(original_lines, "skills and competencies", keyword_only)

        changed = False
        for i, s in enumerate(unique):
            doc_p = self.refs.get(s["ref_id"])
            if doc_p and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_p, new_lines[i])
                changed = True
        return changed

    def _rewrite_role_bullets(self, role: Dict[str, Any]) -> bool:
        bullets = role.get("bullets", [])
        if not bullets:
            return False

        keyword_only = any(
            (doc_p := self.refs.get(b["ref_id"])) is not None and
            (_paragraph_is_narrow_column(doc_p) or _paragraph_has_drawings(doc_p))
            for b in bullets
        )

        original_lines = [b["text"] for b in bullets]
        new_lines = self._safe_rewrite(
            original_lines,
            f"experience bullets — {role.get('title', 'this role')}",
            keyword_only,
        )

        changed = False
        for i, b in enumerate(bullets):
            doc_p = self.refs.get(b["ref_id"])
            if doc_p and new_lines[i] != original_lines[i]:
                self._replace_paragraph_text_preserve_style(doc_p, new_lines[i])
                b["text"] = new_lines[i]
                changed = True
        return changed

    # ──────────────────────────────────────────────────────────────────────────
    # BULLET BALANCING
    # ──────────────────────────────────────────────────────────────────────────

    def _balance_role_bullets(self, role: Dict[str, Any]) -> int:
        bullets = role.get("bullets", [])
        if len(bullets) < 2:
            return 0

        lengths     = [len(b["text"]) for b in bullets]
        med         = _median_int(lengths)
        splits_done = 0

        for i in reversed(range(len(bullets))):
            b      = bullets[i]
            text   = b["text"]
            length = len(text)

            if length <= _OVERLONG_MIN_CHARS:
                continue
            if med > 0 and length / med <= _OVERLONG_RATIO:
                continue

            doc_p = self.refs.get(b["ref_id"])
            if not doc_p:
                continue
            if _paragraph_is_narrow_column(doc_p) or _paragraph_has_drawings(doc_p):
                continue

            split_lines = self._gpt_split_bullet(text, role.get("title", ""))
            if not split_lines or len(split_lines) <= 1:
                continue

            self._replace_paragraph_text_preserve_style(doc_p, split_lines[0])
            for extra in reversed(split_lines[1:]):
                new_p = self._insert_paragraph_after(doc_p, extra)
                if new_p:
                    self.refs[id(new_p)] = new_p

            splits_done += 1
            logging.info(f"Split bullet ({length}ch) → {len(split_lines)} for: {role.get('title','')}")

        return splits_done

    def _gpt_split_bullet(self, text: str, context: str = "") -> List[str]:
        count       = max(2, round(len(text) / _TARGET_BULLET_CHARS))
        context_str = f" for the role: {context}" if context else ""

        prompt = f"""You are editing a resume{context_str}.

Split this overlong bullet into {count} concise bullet points (~{_TARGET_BULLET_CHARS} chars each).

RULES:
- Plain text only — no bullet characters, numbers, or prefixes
- Do NOT fabricate new information
- Return exactly {count} lines

ORIGINAL:
{text}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Split resume bullets. Return only the lines."},
                    {"role": "user",   "content": prompt},
                ],
                timeout=15,
            )
        except Exception as e:
            logging.warning(f"GPT bullet split failed: {e}")
            return []

        raw   = response.choices[0].message.content.strip()
        lines = [
            ln.strip().lstrip("•-–—*›◦▪▸○→✓✔►0123456789.) ")
            for ln in raw.split("\n") if ln.strip()
        ]
        if len(lines) == 1 and len(lines[0]) > _OVERLONG_MIN_CHARS:
            return []
        return lines or []

    # ──────────────────────────────────────────────────────────────────────────
    # PARAGRAPH INSERTION (for bullet splitting)
    # ──────────────────────────────────────────────────────────────────────────

    def _insert_paragraph_after(self, reference_paragraph, text: str):
        try:
            from docx.text.paragraph import Paragraph as DocxParagraph

            ref_elem = reference_paragraph._element
            new_elem = copy.deepcopy(ref_elem)

            # Only operate on text runs in the copy — leave drawing runs alone
            all_runs_in_copy  = new_elem.findall(qn("w:r"))
            text_runs_in_copy = [
                r for r in all_runs_in_copy
                if r.find(f"{{{_MC_NS}}}AlternateContent") is None
                and r.find(qn("w:drawing")) is None
                and r.find(qn("w:object")) is None
            ]

            if text_runs_in_copy:
                first_t = text_runs_in_copy[0].find(qn("w:t"))
                if first_t is None:
                    first_t = text_runs_in_copy[0].makeelement(qn("w:t"), {})
                    text_runs_in_copy[0].append(first_t)
                first_t.text = text
                if text and (text[0] == " " or text[-1] == " "):
                    first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                for extra in text_runs_in_copy[1:]:
                    new_elem.remove(extra)
            else:
                r_e = new_elem.makeelement(qn("w:r"), {})
                t_e = new_elem.makeelement(qn("w:t"), {})
                t_e.text = text
                r_e.append(t_e)
                new_elem.append(r_e)

            ref_elem.addnext(new_elem)
            return DocxParagraph(new_elem, None)

        except Exception as e:
            logging.warning(f"Paragraph insertion failed: {e}")
            return None

    # ──────────────────────────────────────────────────────────────────────────
    # CORE REPLACEMENT — drawing-safe
    # ──────────────────────────────────────────────────────────────────────────

    def _replace_paragraph_text_preserve_style(self, paragraph, new_text: str):
        """
        Replace a paragraph's text while preserving ALL formatting and
        embedded graphics.

        Step 1: Classify every run as either a drawing run or a text run.
                Drawing runs are NEVER modified — they hold background shapes,
                decorative graphics, and layout elements that must survive.

        Step 2: Apply the new text only to text runs:
          - No text runs  → add a new plain run
          - One text run  → replace .text in place (formatting untouched)
          - Many text runs → distribute new text proportionally across text
                            runs only, based on their original character-count
                            weights. Drawing runs are skipped entirely.
        """
        all_runs  = paragraph.runs
        text_runs = [r for r in all_runs if not _is_drawing_run(r)]

        # No writable runs at all
        if not text_runs:
            paragraph.add_run(new_text)
            return

        # Single text run — simplest and safest
        if len(text_runs) == 1:
            text_runs[0].text = new_text
            return

        # Multiple text runs — proportional distribution (text runs only)
        orig_total = sum(len(r.text) for r in text_runs)

        if orig_total == 0:
            text_runs[0].text = new_text
            for r in text_runs[1:]:
                r.text = ""
            return

        new_total = len(new_text)
        assigned  = 0

        for i, run in enumerate(text_runs):
            if i == len(text_runs) - 1:
                run.text = new_text[assigned:]
            else:
                weight   = len(run.text) / orig_total
                chars    = round(weight * new_total)
                max_c    = new_total - assigned - (len(text_runs) - i - 1)
                chars    = max(0, min(chars, max_c))
                run.text = new_text[assigned: assigned + chars]
                assigned += chars
