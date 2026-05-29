"""
resume_parser.py
----------------
Parses a PDF or DOCX resume and extracts:
  Name, Email, Phone, Skills, Education, Experience/Projects

Handles both standard resumes AND heavily designed multi-column PDFs
where section headers get concatenated into one blob.

Usage:
    python resume_parser.py <path_to_resume>
"""

import re
import json
import os
import sys


# ─────────────────────────────────────────────
# SECTION 1: TEXT EXTRACTION
# ─────────────────────────────────────────────

def extract_text_from_pdf(filepath):
    """
    Uses pdfminer.six — much better than pypdf for designed/multi-column PDFs.
    pypdf reads raw character coordinates and produces garbled output on
    styled resumes. pdfminer analyses full page layout first.
    """
    try:
        from pdfminer.high_level import extract_text as pdfminer_extract
        text = pdfminer_extract(filepath)
        if not text or not text.strip():
            raise ValueError("No text extracted. PDF may be image-based (scanned).")
        return text
    except ImportError:
        raise ImportError("pdfminer.six not installed. Run: pip install pdfminer.six")
    except Exception as e:
        raise ValueError(f"Could not read PDF: {e}")


def extract_text_from_docx(filepath):
    """
    Uses python-docx to read every paragraph from a Word document.
    """
    try:
        from docx import Document
        doc = Document(filepath)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        if not full_text.strip():
            raise ValueError("DOCX file appears to be empty.")
        return full_text
    except ImportError:
        raise ImportError("python-docx not installed. Run: pip install python-docx")
    except Exception as e:
        raise ValueError(f"Could not read DOCX: {e}")


def extract_text(filepath):
    """
    Router: checks file exists, reads extension, calls the right extractor.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: '{filepath}'")

    _, ext = os.path.splitext(filepath)
    ext = ext.lower()

    if ext == ".pdf":
        return extract_text_from_pdf(filepath)
    elif ext == ".docx":
        return extract_text_from_docx(filepath)
    else:
        raise ValueError(f"Unsupported file type: '{ext}'. Use .pdf or .docx.")


# ─────────────────────────────────────────────
# SECTION 2: DETECT PDF LAYOUT TYPE
# ─────────────────────────────────────────────

def is_blob_layout(text):
    """
    Detects whether the extracted text is a 'blob' — i.e. a multi-column
    designed PDF where pdfminer collapses everything into one or very few lines.

    WHY this matters: a standard resume has 50–200 lines.
    A designed/multi-column PDF often comes out as 1–5 very long lines.

    If it's a blob, we switch to pattern-based extraction instead of
    line-by-line section detection.
    """
    lines = [l for l in text.split("\n") if l.strip()]
    if not lines:
        return True
    avg_line_length = sum(len(l) for l in lines) / len(lines)
    # If average line is very long, it's a blob
    return avg_line_length > 300


# ─────────────────────────────────────────────
# SECTION 3: STANDARD EXTRACTION (line-by-line)
# Used for normal single-column DOCX / simple PDFs
# ─────────────────────────────────────────────

def extract_name_standard(text):
    """
    For normal resumes: name is the first short clean line at the top.
    Skip lines with @, digits, slashes (email/phone/url).
    """
    for line in text.strip().split("\n")[:6]:
        line = line.strip()
        if 2 < len(line) < 60 and not re.search(r'[@\d/\\|:_]', line):
            return line
    return "Not found"


def extract_section_standard(text, section_keywords):
    """
    For normal resumes: scan line by line, collect content between
    the target heading and the next ALL-CAPS heading.
    """
    ALL_HEADERS = [
        "skills", "education", "experience", "projects", "achievements",
        "about me", "career objective", "hobbies", "certifications",
        "languages", "references", "work experience", "summary",
        "professional experience", "employment"
    ]
    lines = text.split("\n")
    collected = []
    inside = False

    for line in lines:
        clean = line.strip()
        lower = clean.lower()

        if any(kw in lower for kw in section_keywords):
            inside = True
            continue

        if inside:
            if any(lower == h for h in ALL_HEADERS if h not in section_keywords):
                break
            if re.match(r'^[A-Z][A-Z\s]{3,}$', clean):
                break
            if clean:
                collected.append(clean)

    return collected if collected else ["Not found"]


# ─────────────────────────────────────────────
# SECTION 4: BLOB EXTRACTION (pattern-based)
# Used for designed multi-column PDFs
# ─────────────────────────────────────────────

def extract_name_blob(text):
    """
    For blob PDFs: the name is an ALL-CAPS sequence like 'HAMZA MUNIR'.

    The challenge: in multi-column blobs, names are directly concatenated
    with the next section header — e.g. 'HAMZA MUNIRABOUT ME' with no space.
    So normal word-boundary \\b fails after the name.

    Fix: use a LOOKAHEAD — match the all-caps name only when it is immediately
    followed by a known section heading (even without a space).

    re.search returns the FIRST match, which is the name.
    Group 1 captures just the name part, not the header.
    """
    HEADERS = (
        "ABOUT|SKILLS|EDUCATION|EXPERIENCE|PROJECTS|ACHIEVEMENTS|"
        "CAREER|HOBBIES|CERTIFICATIONS|SUMMARY|REFERENCES"
    )
    # Match 1-3 all-caps words immediately before a known section header
    pattern = rf'([A-Z]{{2,}}(?:\s[A-Z]{{2,}}){{0,2}})(?={HEADERS})'
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip()

    # Fallback: any 2 all-caps words separated by a space
    matches = re.findall(r'\b([A-Z]{2,}\s[A-Z]{2,})\b', text)
    SKIP = {
        "ABOUT ME", "SKILLS", "EDUCATION", "EXPERIENCE", "PROJECTS",
        "ACHIEVEMENTS", "CAREER OBJECTIVE", "HOBBIES", "WORK EXPERIENCE"
    }
    for m in matches:
        if m.strip() not in SKIP:
            return m.strip()

    return "Not found"


def extract_skills_blob(text):
    """
    For blob PDFs: skill categories appear as 'Label: content, content'
    e.g. 'Languages: Python, C++, Java'

    The problem: categories are concatenated in the blob — 
    'Languages: Python, C++DataAnalysis: Pandas...' with no separators.

    Fix: find the START position of each category label, then extract
    the text between that label and the NEXT label (or end of a long run).
    This way each category is cleanly isolated.
    """
    # Known category label patterns
    category_labels = [
        "Languages", "Data Analysis", "Machine Learning",
        "Databases?", r"Web\s*(?:&|and)?\s*Tools?",
        "Frameworks?", "Technologies", "Soft Skills?",
        "Cloud", "DevOps", "Testing"
    ]

    # Build a combined pattern to find ALL label positions
    label_pattern = r'(' + '|'.join(category_labels) + r')\s*:\s*'

    # Find all (label, start_position) pairs
    label_matches = list(re.finditer(label_pattern, text, re.IGNORECASE))

    if not label_matches:
        return ["Not found"]

    results = []
    for i, match in enumerate(label_matches):
        label = match.group(1).strip()
        content_start = match.end()

        # Content ends where the next label starts (or 200 chars max)
        if i + 1 < len(label_matches):
            content_end = label_matches[i + 1].start()
        else:
            content_end = content_start + 200

        content = text[content_start:content_end].strip()
        # Clean line breaks and collapse whitespace
        content = re.sub(r'\s+', ' ', content).strip()
        # Remove any trailing all-caps blob leak (e.g. 'GitHubACADEMIC')
        content = re.split(r'[A-Z]{4,}', content)[0].strip()

        if content:
            results.append(f"{label}: {content}")

    return results if results else ["Not found"]


def extract_education_blob(text):
    """
    For blob PDFs: look for education signals — degree keywords,
    university names, CGPA, year ranges.

    Strategy: find segments containing degree/institution keywords
    and extract the surrounding context.
    """
    edu_signals = [
        r'(?:BS|BE|MS|MBA|PhD|Bachelor|Master|Undergraduate|Intermediate|Matriculation)'
        r'.{0,120}(?:University|College|Institute|School)',
        r'CGPA\s*:\s*[\d.]+',
        r'\b(?:20\d{2})\s*[-–]\s*(?:20\d{2}|ongoing|present)',
    ]

    found = []
    for pattern in edu_signals:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for m in matches:
            clean = m.strip()
            if clean and clean not in found:
                found.append(clean)

    return found if found else ["Not found"]


def extract_experience_blob(text):
    """
    For blob PDFs: project/experience bullet points start with action verbs.
    The blob concatenates them without newlines, e.g.:
    'Built X.Developed Y.Designed Z.'

    Strategy:
    1. Find all sentences starting with known action verbs
    2. Split them individually (not as one huge string)
    3. Also find project titles (Title – Category pattern)

    Each action-verb sentence is one bullet point. We cap at 200 chars
    to avoid runaway matches.
    """
    action_verbs = (
        r'(?:Built|Developed|Designed|Implemented|Created|Performed|'
        r'Trained|Managed|Led|Achieved|Deployed|Automated|Analysed|Analyzed|'
        r'Extracted|Optimized|Maintained|Integrated|Contributed)'
    )

    # Each bullet: action verb + content, stop at next capital letter sentence
    # or after 200 chars. The [^.]{10,200} captures up to 200 non-period chars.
    bullet_pattern = rf'{action_verbs}[^.]*\.?'
    bullets = re.findall(bullet_pattern, text)

    # Project title pattern: e.g. "Obesity Level Prediction - ML based Project"
    title_pattern = (
        r'[A-Z][a-zA-Z\s]+(?:–|-)\s*[A-Z][a-zA-Z\s]+'
        r'(?:Project|System|Application|App)\b'
    )
    titles = re.findall(title_pattern, text)

    # Combine: titles first, then bullets
    combined = []
    seen = set()
    for item in titles + bullets:
        clean = re.sub(r'\s+', ' ', item).strip()
        # Trim to 180 chars max for readability
        if len(clean) > 180:
            clean = clean[:180].rstrip() + "..."
        if clean and clean not in seen:
            seen.add(clean)
            combined.append(clean)

    return combined if combined else ["Not found"]


# ─────────────────────────────────────────────
# SECTION 5: SHARED EXTRACTORS (work on any layout)
# ─────────────────────────────────────────────

def extract_email(text):
    """
    Finds the first email address.

    THE BLOB PROBLEM: in designed PDFs, pdfminer concatenates the contact
    line so phone, email, and LinkedIn URL run together with no spaces:
      '+92 3261764672iamhamza1032@gmail.comlinkedin.com/in/user/'

    CHAIN OF FIXES applied in order:
    1. Pre-process: insert a space after '.com/.org/.pk' etc. when immediately
       followed by another letter. This separates 'gmail.comlinkedin' into
       'gmail.com linkedin' BEFORE the regex even runs.
    2. Require local part to start with a letter — skips leading digits
       from a phone number that ran into the email with no separator.
    3. Negative lookahead (?![a-zA-Z]) as a final guard on the TLD.
    """
    # Step 1: separate concatenated domains, e.g. 'gmail.comlinkedin' → 'gmail.com linkedin'
    cleaned = re.sub(
        r'(\.(com|org|net|io|pk|co|edu|gov|info))([a-zA-Z])',
        r'\1 \3',
        text,
        flags=re.IGNORECASE
    )

    # Step 2 & 3: match email with letter-start local part and bounded TLD
    pattern = r'[a-zA-Z][a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}(?![a-zA-Z])'
    match = re.search(pattern, cleaned)
    return match.group() if match else "Not found"


def extract_phone(text):
    """
    Regex: optional +, starts with digit, 7-15 mixed chars, ends with digit.
    Handles: +92 3261764672, +1 (555) 555-5555, 03001234567
    """
    pattern = r'(\+?\d[\d\s\-(). ]{7,15}\d)'
    matches = re.findall(pattern, text)
    return matches[0].strip() if matches else "Not found"


# ─────────────────────────────────────────────
# SECTION 6: ORCHESTRATION
# ─────────────────────────────────────────────

def parse_resume(filepath):
    """
    Full pipeline:
    1. Extract raw text
    2. Detect whether it's a blob (multi-column PDF) or normal layout
    3. Route to the right set of extractors
    4. Return structured dict

    This two-path approach is what makes the parser robust:
    - Normal PDFs / DOCX → line-by-line section detection
    - Designed multi-column PDFs → direct pattern/regex extraction
    """
    print(f"\n Parsing: {filepath}")

    text = extract_text(filepath)
    blob = is_blob_layout(text)

    if blob:
        print("   Layout type: multi-column / designed PDF (using pattern extraction)")
        name       = extract_name_blob(text)
        skills     = extract_skills_blob(text)
        education  = extract_education_blob(text)
        experience = extract_experience_blob(text)
    else:
        print("   Layout type: standard single-column (using line-by-line extraction)")
        name       = extract_name_standard(text)
        skills     = extract_section_standard(text, ["skills", "technical skills", "core competencies"])
        education  = extract_section_standard(text, ["education", "academic background", "qualifications"])
        experience = extract_section_standard(text, ["experience", "work experience", "professional experience", "projects"])

    data = {
        "source_file": os.path.basename(filepath),
        "name":        name,
        "email":       extract_email(text),
        "phone":       extract_phone(text),
        "skills":      skills,
        "education":   education,
        "experience":  experience,
    }

    return data


# ─────────────────────────────────────────────
# SECTION 7: SAVE TO JSON
# ─────────────────────────────────────────────

def save_to_json(data, output_path):
    """
    Writes the result dict to a pretty-printed JSON file.
    Creates output directory if it doesn't exist.
    utf-8 encoding handles non-Latin characters.
    ensure_ascii=False keeps them readable (not \\uXXXX escaped).
    """
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    print(f"\n Saved to: {output_path}")


# ─────────────────────────────────────────────
# SECTION 8: ENTRY POINT
# ─────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("=" * 50)
        print("  Resume Parser")
        print("=" * 50)
        print("Usage:   python resume_parser.py <path_to_resume>")
        print("Example: python resume_parser.py my_resume.pdf")
        print("         python resume_parser.py resume.docx")
        sys.exit(1)

    input_file  = sys.argv[1]
    output_file = "output/parsed_resume.json"

    try:
        data = parse_resume(input_file)
        save_to_json(data, output_file)

        print("\n Extracted Data:")
        print("-" * 40)
        print(json.dumps(data, indent=4, ensure_ascii=False))

    except FileNotFoundError as e:
        print(f"\n File Error: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"\n Parse Error: {e}")
        sys.exit(1)
    except ImportError as e:
        print(f"\n Missing Library: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n Unexpected Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()