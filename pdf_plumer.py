import pdfplumber
import re
import pandas as pd

def extract_text_from_pdf(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text


def extract_cat7_sbs_with_desc(text):
    # Step 1: Find Cat 7 SB section
    pattern = r'Incorporate.*Cat\s*7.*SB[’\'s]*:?(.*?)(?=\n\s*[a-z]\.|$)'
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    
    if not match:
        return []

    section_text = match.group(1)

    # Step 2: Fix broken lines (VERY IMPORTANT for PDFs)
    section_text = re.sub(r'\n+', ' ', section_text)

    # Step 3: Extract SB + Description
    matches = re.findall(r'(\d{2}-\d{4})\s*[–-]\s*(.*?)(?=\d{2}-\d{4}|$)', section_text)

    results = [(m[0], m[1].strip()) for m in matches]

    return results


def process_pdf(pdf_path):
    text = extract_text_from_pdf(pdf_path)
    return extract_cat7_sbs_with_desc(text)


# -------- MAIN --------
pdf_path = "input.pdf"

data = process_pdf(pdf_path)

df = pd.DataFrame(data, columns=["SB Number", "Description"])

df.to_excel("sb_output.xlsx", index=False)

print(df)