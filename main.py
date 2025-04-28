import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from lxml import etree
from zipfile import ZipFile
import re
import io

# === SETTINGS ===
output_path = "Result.docx"

# === CREATE OUTPUT DOCUMENT ===
formatted_doc = Document()

# === 1. Utility Functions ===

def is_grey_textbox(paragraph):
    p = paragraph._element
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        shd = pPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'))
            return fill == 'D9D9D9'
    return False

def set_character_spacing(run, spacing_val):
    rPr = run._element.get_or_add_rPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:val'), str(spacing_val))
    rPr.append(spacing)

def extract_grey_textboxes_from_docx(docx_path):
    with ZipFile(docx_path) as docx_zip:
        xml = docx_zip.read("word/document.xml")
    tree = etree.XML(xml)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    boxes = []
    for p in tree.xpath('//w:p', namespaces=ns):
        pPr = p.find('w:pPr', namespaces=ns)
        if pPr is not None:
            shd = pPr.find('w:shd', namespaces=ns)
            if shd is not None and shd.get(f'{{{ns["w"]}}}fill') == 'D9D9D9':
                texts = p.xpath('.//w:t', namespaces=ns)
                full_text = ''.join(t.text for t in texts if t.text)
                if full_text.strip():
                    boxes.append(full_text.strip())
    return boxes

def set_paragraph_border(paragraph):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')

    for border in ['top', 'left', 'bottom', 'right']:
        side = OxmlElement(f'w:{border}')
        side.set(qn('w:val'), 'single')
        side.set(qn('w:sz'), '2')
        side.set(qn('w:space'), '2')
        side.set(qn('w:color'), 'auto')
        pBdr.append(side)

    pPr.append(pBdr)

    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'D9D9D9')
    pPr.append(shd)

def apply_format(paragraph, font_size, bold, align, spacing_after=6, spacing_before=0, line_spacing_rule=WD_LINE_SPACING.SINGLE):
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    font = run.font
    font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    font.size = Pt(font_size)
    font.bold = bold
    font.color.rgb = RGBColor(0, 0, 0)
    pf = paragraph.paragraph_format
    pf.alignment = align
    pf.space_after = Pt(spacing_after)
    pf.space_before = Pt(spacing_before)
    pf.line_spacing_rule = line_spacing_rule
    pf.line_spacing = Pt(12)

def is_list_item(paragraph):
    if paragraph.style.name.startswith('List Number'):
        return True
    return False

def is_signature_line(text):
    return bool(re.search(r'.{5,}\s{2,}.+', text)) or '\t' in text

def add_signature_line(title, initials):
    table = formatted_doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    table.columns[0].width = Inches(4.7)
    table.columns[1].width = Inches(1.8)

    row = table.rows[0]
    cell_title = row.cells[0]
    cell_initials = row.cells[1]

    p1 = cell_title.paragraphs[0]
    run1 = p1.add_run(title)
    run1.bold = True
    run1.font.name = 'Times New Roman'
    run1.font.size = Pt(12)

    p2 = cell_initials.paragraphs[0]
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run2 = p2.add_run(initials)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

def classify_and_format(paragraph):

    if is_list_item(paragraph):
        p = formatted_doc.add_paragraph(paragraph.text, style='List Number')
        
    text = paragraph.text.strip()
    if not text:
        return

    if text == "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА":
        p = formatted_doc.add_paragraph()
        run = p.add_run(text)
        set_character_spacing(run, 60)
        apply_format(p, 14, True, WD_PARAGRAPH_ALIGNMENT.CENTER)

    elif text.lower().startswith("к вопросу"):
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.CENTER)
        p = formatted_doc.add_paragraph()

    elif is_signature_line(text):
        if '\t' in text:
            parts = text.split('\t')
        else:
            parts = re.split(r'\s{2,}', text)
        title = parts[0].strip()
        initials = parts[1].strip() if len(parts) > 1 else ""
        add_signature_line(title, initials)

    elif text.strip() == "Согласовано":
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.CENTER)

    elif text.strip() == "Информация о страховом случае": 
        p = formatted_doc.add_paragraph(text)
        apply_format(p,11,True,WD_PARAGRAPH_ALIGNMENT.LEFT)

    elif text in ("Управляющий директор", "Исполнитель"):
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.LEFT)

    elif text.startswith("Основание для рассмотрения вопроса Советом директоров"):
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.LEFT)

    elif text.startswith("* * *"):
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, False, WD_PARAGRAPH_ALIGNMENT.CENTER)
    
    elif text.startswith("Приложения") :
        p = formatted_doc.add_paragraph(text)
    
    elif text.startswith("Осуществить страховую выплату"):
        p = formatted_doc.add_paragraph(text)
    else:
        p = formatted_doc.add_paragraph(text)
        apply_format(p, 12, False, WD_PARAGRAPH_ALIGNMENT.LEFT)

# === Streamlit UI ===

# Upload DOCX file
uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    # Save uploaded file to a temporary location
    with open("temp.docx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Load the original DOCX file
    original_doc = Document("temp.docx")
    
    # Process paragraphs from the uploaded DOCX
    for para in original_doc.paragraphs:
        if is_grey_textbox(para):
            # This paragraph is a grey textbox
            p = formatted_doc.add_paragraph()
            run = p.add_run(para.text.strip())
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
            set_paragraph_border(p)
        else: 
            classify_and_format(para)

    # Save the formatted result to a file
    formatted_doc.save(output_path)

    # Provide a download link to the user
    with open(output_path, "rb") as f:
        st.download_button("Download Reformatted DOCX", f, file_name="Reformatted_Result.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
