import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT,WD_LINE_SPACING,WD_COLOR_INDEX,WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import io
#input_path = "Change.docx"
#output_path = "Reformatted_Change.docx"
st.title("DOCX Formatter")
uploaded_file = st.file_uploader("Upload your Word (.docx) file", type=["docx"])

if uploaded_file:
# Load document
    doc = Document(uploaded_file)
    formatted_doc = Document()
    
    sections = doc.sections
    formatted_sections = formatted_doc.sections
    for fsection in formatted_doc.sections:
        fsection.top_margin = Inches(1)
        fsection.bottom_margin = Inches(1)
        fsection.left_margin = Inches(1)
        fsection.right_margin = Inches(1)
        
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    # for i, section in enumerate(doc.sections):
    #     top = section.top_margin.inches
    #     bottom = section.bottom_margin.inches
    #     left = section.left_margin.inches
    #     right = section.right_margin.inches
    #     print(top)
    #     print(bottom)
    #     print(left)
    #     print(right)
    #     section.top_margin = Inches(1)
    #     section.bottom_margin = Inches(1)
    #     section.left_margin = Inches(1)
    #     section.right_margin = Inches(1)
    def insert_page_numbers_except_first(document):
        for section in document.sections:
            # Enable different first page
            section.different_first_page_header_footer = True

            # Get footer for all pages except first
            footer = section.footer
            paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

            # Add only the page number field { PAGE }
            run = paragraph.add_run()

            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')

            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = 'PAGE'

            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'separate')

            fldChar3 = OxmlElement('w:fldChar')
            fldChar3.set(qn('w:fldCharType'), 'end')

            run._r.append(fldChar1)
            run._r.append(instrText)
            run._r.append(fldChar2)
            run._r.append(fldChar3)
    def clean_text(text):
        text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces/newlines with a single space
        return text.strip()
    def clean_text_extended(text):
    # Remove invisible characters and carriage returns
        text = re.sub(r'[\u200B-\u200D\uFEFF]', '', text)
        
        # Replace all newlines, tabs, and multiple spaces with a single space
        text = re.sub(r'[\n\r\t]+', ' ', text)
        text = re.sub(r'\s{2,}', ' ', text)
        # [\n\r\t]+: removes any newlines (\n), carriage returns (\r), and tabs (\t).

        # \s{2,}: collapses double/multiple spaces into one.

        # strip(): trims leading/trailing whitespace.
        text = text.replace('\n','')
        return text.strip()
    def clean_from_invisible_char(text):
        text = re.sub()
        text = text.replace('\r', '')
    def shade_paragraph(paragraph, color="D9D9D9"):
        p_pr = paragraph._element.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)
        p_pr.append(shd)
    def apply_format(paragraph, font_size, bold, align, spacing_after=3, spacing_before=0):
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
        pf.line_spacing = 1.0  
       # pf.line_spacing = Pt(12)

    def apply_format2(paragraph, font_size, bold, align, spacing_after, spacing_before=0):
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
        #pf.line_spacing = Pt(12)

    def set_cell_shading(cell, color):
    # Add shading to a table cell
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)  # Hex color, e.g., 'D9D9D9'
        tcPr.append(shd)
    def move_short_words_to_next_line(text):
        lines = text.split('\n')
        new_lines = []
        
        for line in lines:
            words = line.strip().split()
            if len(words) >= 2 and len(words[-1]) > 2:
                if len(words[-2]) <= 2:
                    # move short word to next line
                    new_line = " ".join(words[:-2])
                    short_word = words[-2]
                    last_word = words[-1]
                    if new_line:
                        new_lines.append(new_line)
                    new_lines.append(f"{short_word}\u00A0{last_word}")
                    continue
            new_lines.append(line)
        
        return '\n'.join(new_lines)

    def apply_typographic_fixes(text):
        # 1. Non-breaking space after №
        text = text.replace("№ ", "№\u00A0")

        # 2. Non-breaking space between day and month (using raw string pattern, regular string replacement)
        text = re.sub(
            r'\b(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\b',
            lambda m: f"{m.group(1)}\u00A0{m.group(2)}",
            text
        )

        # 3. Move 1–2 letter word from end of line to start of next
        text = move_short_words_to_next_line(text)

        return text

    def set_character_spacing(run, spacing_val):
        rPr = run._element.get_or_add_rPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:val'), str(spacing_val))
        rPr.append(spacing)

    def fix_docx_numbering(text):
        """
        Fix duplicated numbering at the beginning of a line: '1. 1. Text' -> '1. Text'
        """
        return re.sub(r'^(\d+\.)\s+\d+\.\s+', r'\1 ', text)

    def set_format(paragraph, size=12, bold=False, align=WD_PARAGRAPH_ALIGNMENT.LEFT):
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        font = run.font
        font.name = 'Times New Roman'
        font.size = Pt(size)
        font.bold = bold
        font.color.rgb = RGBColor(0, 0, 0)
        paragraph.paragraph_format.alignment = align

    def add_signature_table(title, name):
        table = formatted_doc.add_table(rows=1, cols=2)
        table.columns[0].width = Inches(4.7)
        table.columns[1].width = Inches(1.8)
        cell1, cell2 = table.rows[0].cells
        p1 = cell1.paragraphs[0].add_run(title)
        p1.bold = True
        p1.font.size = Pt(12)
        p1.font.name = 'Times New Roman'
        p2 = cell2.paragraphs[0].add_run(name)
        p2.font.size = Pt(12)
        p2.font.name = 'Times New Roman'
        cell2.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    def set_shading(run, fill_color):
        """Applies background shading (highlighting) to a run."""
        rPr = run._element.get_or_add_rPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_color)  # e.g., 'FFFF00' for yellow
        rPr.append(shd)

    def clear_headers_and_footers(doc: Document):
        for section in doc.sections:
            # Очистка верхнего колонтитула
            header = section.header
            for para in header.paragraphs:
                para.clear()

            # Очистка нижнего колонтитула
            footer = section.footer
            for para in footer.paragraphs:
                para.clear()
    def bold_keywords(paragraph, keywords):
        for keyword in keywords:
            if keyword in paragraph.text:
                # Store existing text and clear paragraph
                text = paragraph.text
                paragraph.clear()
                i = 0
                while i < len(text):
                    matched = False
                    for word in keywords:
                        if text[i:].startswith(word):
                            run = paragraph.add_run(word)
                            run.bold = True
                            run.font.name = 'Times New Roman'
                            i += len(word)
                            matched = True
                            break
                    if not matched:
                        run = paragraph.add_run(text[i])
                        i += 1
    
    cleaned_paragraphs = [
        clean_text_extended(para.text)
        for para in doc.paragraphs
        if clean_text_extended(para.text)  # filters out empty ones
    ]
    # for i, para in enumerate(cleaned_paragraphs):
    #     print(f"{i+1}. {para}")
        
    # === Split paragraphs by blocks ===
    blocks = {}
    current_block = None
    
    clear_headers_and_footers(doc)
    for para in doc.paragraphs:
        text = para.text.strip()
        #text = clean_text_extended(para.text.strip())
        text = clean_text_extended(text)
        if not text:
            continue
        #print(repr(para.text))
        #print(para[0].text)
        #print(text)
        block_header = re.match(r"^Блок(\d+)", text)
        #block_header = {}
        if block_header:
            current_block = f"Блок{block_header.group(1)}"
            blocks[current_block] = []
        elif current_block:
            blocks[current_block].append(text)
    #insert_page_numbers_except_first(doc)
    # === Process blocks with specific styles ===
    if doc.paragraphs and doc.paragraphs[0].text.strip() == "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА":
        for block, paragraphs in blocks.items():
            #formatted_doc.add_paragraph(block).runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red block header

            if block == "Блок1":
                text = "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА"
                p = formatted_doc.add_paragraph()
                run = p.add_run(text)
                set_character_spacing(run,60)
                apply_format(p,14,True,WD_PARAGRAPH_ALIGNMENT.CENTER)
                for para in paragraphs:
                    para = clean_text(para)
                   # print(para)
                    p = formatted_doc.add_paragraph(para)
                    set_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.CENTER)
                    apply_typographic_fixes(p.text)
            elif block == "Блок2":
                text = "Условные (сокращенные) обозначения, использованные в пояснительной записке"
                p = formatted_doc.add_paragraph()
                run = p.add_run(text)
                apply_format(p,11,True,WD_PARAGRAPH_ALIGNMENT.LEFT)
                for para in paragraphs:
                    para = clean_text(para)
                    if "–" in para or "-" in para:
                        p = formatted_doc.add_paragraph(para)
                        set_format(p, 11, False)
                        apply_typographic_fixes(p.text)
                        apply_format(p,11,False,WD_PARAGRAPH_ALIGNMENT.LEFT)
                    else:
                        parts = para.split("–", 1)
                        if len(parts) == 2:
                            term, desc = parts
                            p = formatted_doc.add_paragraph()
                            apply_typographic_fixes(p.text)
                            run = p.add_run(term.strip() + " – ")
                            run.bold = True
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(11)
                            p.add_run(desc.strip())
                            set_format(p, 11)
                            apply_format(p,11,False,WD_PARAGRAPH_ALIGNMENT.LEFT)
                formatted_doc.add_paragraph()
                #formatted_doc.add_paragraph()
            # elif block == "Блок3":
            #     #lines_to_merge = []
            #     #for p in doc.paragraphs:
            #     #    if p.text.strip():  # skip empty lines
            #     #        lines_to_merge.append(p.text)

            #     for para in paragraphs:
            #         p = formatted_doc.add_paragraph(para)
            #         run = p.add_run()
            #         run = p.runs[0] if p.runs else p.add_run()
            #         font = run.font
            #         font.name = 'Times New Roman'
            #         run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #         font.size = Pt(11)
            #         font.bold = False
            #         font.color.rgb = RGBColor(0, 0, 0)
            #         set_shading(run, 'D9D9D9')
            #     # Join all lines with a line break
            #     #for idx, line in enumerate(lines_to_merge):
            #     #    run.add_text(line)
            #     #    if idx != len(lines_to_merge) - 1:
            #     #        run.add_break() 

            #         pf = p.paragraph_format
            #         pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            #         pf.space_after = Pt(3)
            #         pf.space_before = Pt(0)
            #         pf.line_spacing = 1.0  
            #         pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            #         if p.text == "Основание выноса вопроса на рассмотрение Советом директоров":
            #             font.bold = True
            elif block == "Блок3":
                
                # table = formatted_doc.add_table(rows=1, cols=1)
                # cell = table.cell(0, 0)
                # set_cell_shading(cell, 'D9D9D9')

                # for para in paragraphs:
                #     p = cell.add_paragraph(para)
                #     run = p.runs[0] if p.runs else p.add_run()
                #     font = run.font
                #     font.name = 'Times New Roman'
                #     run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                #     font.size = Pt(11)
                #     font.bold = False
                #     font.color.rgb = RGBColor(0, 0, 0)

                #     pf = p.paragraph_format
                #     pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                #     pf.space_after = Pt(3)
                #     pf.space_before = Pt(0)
                #     pf.line_spacing = 1.0
                #     pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

                #     if para.strip() == "Основание выноса вопроса на рассмотрение Советом директоров":
                #         font.bold = True
                for para in paragraphs: 
                    #formatted_doc.add_paragraph()
                    para = clean_text(para)
                    p = formatted_doc.add_paragraph(para)
                    run = p.add_run()
                    run = p.runs[0] if p.runs else p.add_run()
                    font = run.font

                    font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    font.size = Pt(11)
                    font.bold = False
                    if para.strip() == "Основание выноса вопроса на рассмотрение Советом директоров":
                        font.bold = True
                    
                    pf = p.paragraph_format
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(2)  # Уменьшает расстояние после строки
                    pf.line_spacing = 1.0 
                    shade_paragraph(p)
                    
            elif block == "Блок4":
                for para in paragraphs: 
                    para = clean_text(para)
                    formatted_doc.add_paragraph()
                    p = formatted_doc.add_paragraph(para)
                    run = p.add_run()
                    run = p.runs[0] if p.runs else p.add_run()
                    font = run.font
                    font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    font.size = Pt(12)
                    font.bold = False
                    font.color.rgb = RGBColor(0, 0, 0)


                    pf = p.paragraph_format
                    pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                    pf.space_after = Pt(6)
                    pf.space_before = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    #set_shading(run, 'D9D9D9')

            # elif block == "Блок5":
            #     for para in paragraphs: 
                        
            #             fixed_text = apply_typographic_fixes(para) 
            #             p = formatted_doc.add_paragraph(fixed_text)
                        
            #             #if para.startswith(""): 

            #             pf = p.paragraph_format
            #             pf.first_line_indent = Inches(0.3)                     # Выступ (отступ первой строки)
            #             pf.space_before = Pt(0)                                # Интервал перед абзацем
            #             pf.space_after = Pt(3)                                 # Интервал после абзаца
            #             pf.line_spacing = 1.0                                  # Одинарный межстрочный
            #             pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            #             pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY       
            #             #set_format(p)
            #             #apply_typographic_fixes(p.text)
            #             run = p.add_run()
            #             run = p.runs[0] if p.runs else p.add_run()
            #             font = run.font
            #             font.name = 'Times New Roman'
            #             run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #             font.size = Pt(11)
            # elif block == "Блок5":
            #     for para in paragraphs:
            #         fixed_text = apply_typographic_fixes(para)
                    
            #         # Проверка наличия тире
            #         if "–" in fixed_text:
            #             parts = fixed_text.split("–", 1)
            #             term, desc = parts[0].strip(), parts[1].strip()

            #             p = formatted_doc.add_paragraph()
            #             pf = p.paragraph_format
            #             pf.first_line_indent = Inches(0.3)
            #             pf.space_before = Pt(0)
            #             pf.space_after = Pt(3)
            #             pf.line_spacing = 1.0
            #             pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            #             pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

            #             # Первая часть (жирная)
            #             run = p.add_run(term + " –\t")
            #             run.font.bold = True
            #             run.font.name = 'Times New Roman'
            #             run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #             run.font.size = Pt(11)

            #             # Вторая часть (обычная)
            #             run2 = p.add_run(desc)
            #             run2.font.name = 'Times New Roman'
            #             run2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #             run2.font.size = Pt(11)

            #         else:
            #             # Обычный абзац
            #             p = formatted_doc.add_paragraph(fixed_text)
            #             pf = p.paragraph_format
            #             pf.first_line_indent = Inches(0.3)
            #             pf.space_before = Pt(0)
            #             pf.space_after = Pt(3)
            #             pf.line_spacing = 1.0
            #             pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            #             pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

            #             run = p.runs[0]
            #             run.font.name = 'Times New Roman'
            #             run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #             run.font.size = Pt(11)
            # elif block == "Блок5":
            #     for para in paragraphs: 
            #         fixed_text = apply_typographic_fixes(para) 

            #         # Создаем абзац с текстом
            #         p = formatted_doc.add_paragraph(fixed_text)

            #         # Форматирование абзаца
            #         pf = p.paragraph_format
            #         pf.first_line_indent = Inches(0.3)                 # Отступ первой строки (выступ)
            #         pf.space_before = Pt(0)                            # Интервал перед абзацем
            #         pf.space_after = Pt(3)                             # Интервал после абзаца
            #         pf.line_spacing = 1.0                              # Одинарный межстрочный интервал
            #         pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            #         pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY      # Выровнять по ширине

            #         # Безопасный доступ к run
            #         if p.runs:
            #             run = p.runs[0]
            #         else:
            #             run = p.add_run()

            #         # Настройка шрифта
            #         run.font.name = 'Times New Roman'
            #         run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            #         run.font.size = Pt(11)
            elif block == "Блок5":
                
                # for para in paragraphs:
                #     #para = apply_typographic_fixes(para)
                #     # Обработка только если есть тире
                #     if "–" in para or "-" in para:
                #         dash = "–" if "–" in para else "-"
                #         parts = para.split(dash, 1)
                #         if len(parts) == 2:
                #             term, desc = parts
                #             p = formatted_doc.add_paragraph()
                #             font = p.add_run().font
                #             font.highlight_color = WD_COLOR_INDEX.GRAY_50
                #             #tabs = p.paragraph_format.tab_stops
                #             #tabs.add_tab_stop(Inches(1.0))
                #             # Настройка абзаца
                #             pf = p.paragraph_format
                #            # pf.first_line_indent = Inches(0.3)                 # Отступ первой строки
                #             pf.space_before = Pt(0)
                #             pf.space_after = Pt(3)
                #            # pf.left_indent = Inches(0.3)
                #             pf.line_spacing = 1.0
                #             pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                #             pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                #             # Часть до тире
                #             run1 = p.add_run(" –\t" + term.strip() )#+ " –\t" + desc.strip())
                #             run1.font.name = 'Times New Roman'
                #             run1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                #             run1.font.size = Pt(11)
                #             #run.font.bold = True

                #             # Часть после тире
                #             run2 = p.add_run(desc.strip())
                #             run2.font.name = 'Times New Roman'
                #             run2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                #             run2.font.size = Pt(11)
                #             #run2.font.bold = Falses
                #     else:
                #         # Без тире – обычный абзац
                #         p = formatted_doc.add_paragraph(para)
                #         pf = p.paragraph_format
                #         #pf.first_line_indent = Inches(0.3)
                #         pf.space_before = Pt(0)
                #         pf.space_after = Pt(3)
                #         pf.line_spacing = 1.0
                #         pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                #         pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                #         run = p.add_run()
                #         run = p.runs[0] if p.runs else p.add_run()
                #         run.text = para
                #         run.font.name = 'Times New Roman'
                #         run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                #         run.font.size = Pt(11)
                if paragraphs:
                    
                    first_para = formatted_doc.add_paragraph()
                    #first_para = clean_text(first_para)
                    pf = first_para.paragraph_format
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(3)
                    pf.line_spacing = 1.0
                    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                    run = first_para.add_run(paragraphs[0].strip())
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    run.font.size = Pt(11)
                    shade_paragraph(first_para)
                for para in paragraphs[1:]:
        # Проверка на наличие тире
                    para = clean_text(para)
                    if "–" in para or "-" in para:
                        dash = "–" if "–" in para else "-"
                        parts = para.split(dash, 1)

                        if len(parts) == 2:
                            term, desc = parts

                            # Создаем абзац
                            p = formatted_doc.add_paragraph()
                            pf = p.paragraph_format

                            # font = p.add_run().font
                            # font.highlight_color = WD_COLOR_INDEX.GRAY_50
                            
                            # Висячий отступ
                            pf.left_indent = Inches(0.3)
                            pf.first_line_indent = Inches(-0.3)
                            pf.tab_stops.add_tab_stop(Inches(0.3), WD_TAB_ALIGNMENT.LEFT)
                            # Прочее форматирование
                            pf.space_before = Pt(0)
                            pf.space_after = Pt(3)
                            pf.line_spacing = 1.0
                            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                            pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                            # Добавляем tab stop (необязательно, но полезно для ручного редактирования)
                            #pf.tab_stops.add_tab_stop(Inches(0.3), WD_TAB_ALIGNMENT.LEFT)

                            # Добавляем тире, табуляцию и текст
                            run = p.add_run("–\t" + desc.strip())
                            run.font.name = 'Times New Roman'
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                            run.font.size = Pt(11)
                            #run.font.highlight_color = WD_COLOR_INDEX.GRAY_50
                            shade_paragraph(p)
                    else:
                        # Абзац без тире — просто отформатированный
                        p = formatted_doc.add_paragraph()
                        pf = p.paragraph_format
                        
                        font = p.add_run().font
                        font.highlight_color = WD_COLOR_INDEX.GRAY_50
                        
                        # Такой же висячий отступ для выравнивания
                        pf.left_indent = Inches(0.3)
                        pf.first_line_indent = Inches(-0.3)
                        #pf.left_indent = Inches(0.3)
                        #pf.first_line_indent = Inches(-0.3)
                        pf.tab_stops.add_tab_stop(Inches(0.3), WD_TAB_ALIGNMENT.LEFT)
                        
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(3)
                        pf.line_spacing = 1.0
                        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                        #pf.tab_stops.add_tab_stop(Inches(0.3), WD_TAB_ALIGNMENT.LEFT)
                        run = p.add_run("\t" + para.strip())
                        run.font.name = 'Times New Roman'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run.font.size = Pt(11)
                        #run.font.highlight_color = WD_COLOR_INDEX.GRAY_50
                        shade_paragraph(p)
            # elif block == "Блок7":
            #     for para in paragraphs:
            #         formatted_doc.add_paragraph()
            #         if re.match(r"^\d+\.", para):
            #             p = formatted_doc.add_paragraph(para)
            #             p.text = fix_docx_numbering(p.text)
            #             apply_typographic_fixes(p.text)
                        
            #         else:
            #             p = formatted_doc.add_paragraph(para)
            #         set_format(p)
            elif block == "Блок7":
                #* * *
                text = "* * *"
                ps = formatted_doc.add_paragraph(text)
                psf = ps.paragraph_format
                psf.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                
                runf = ps.runs[0]
                runf.font.name = 'Times New Roman'
                runf._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                runf.font.size = Pt(12)
                
                for idx, para in enumerate(paragraphs):
                    #para = clean_text(para)
                    if idx == 0:
                        # Первый параграф (вступление) — без табуляции и висячего отступа
                        p = formatted_doc.add_paragraph()
                        pf = p.paragraph_format
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(3)
                        pf.line_spacing = 1.0
                        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                        run = p.add_run(para.strip())
                        run.font.name = 'Times New Roman'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run.font.size = Pt(12)
                        continue  # переходим к следующему абзацу

                    # Далее — обычная обработка
                    if re.match(r"^\d+\.", para.strip()):
                        p = formatted_doc.add_paragraph()
                        pf = p.paragraph_format
                        pf.left_indent = Inches(0.5)
                        pf.first_line_indent = Inches(-0.5)
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(3)
                        pf.line_spacing = 1.0
                        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                        
                        # tab_stops = p.paragraph_format.tab_stops
                        # tab_stops.add_tab_stop(Inches(1.0), alignment=WD_TAB_ALIGNMENT.LEFT)
                       
                        numbered_text = fix_docx_numbering(para.strip())
                        numbered_text = re.sub(r"^(\d+\.)\s*", r"\1\t", numbered_text)
                        tab_stops = p.paragraph_format.tab_stops
                        tab_stops.add_tab_stop(Inches(1.0), alignment=WD_TAB_ALIGNMENT.LEFT)
                       
                        #tab_stops = p.paragraph_format.tab_stops
                #       #tab_stops.add_tab_stop(Inches(5.0), alignment=WD_TAB_ALIGNMENT.LEFT)
                        
                        run = p.add_run( numbered_text)
                        run.font.name = 'Times New Roman'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run.font.size = Pt(12)
                        shade_paragraph(p)
                    else:
                        p = formatted_doc.add_paragraph()
                        pf = p.paragraph_format
                        pf.left_indent = Inches(0.5)
                        pf.first_line_indent = Inches(0.5)
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(3)
                        pf.line_spacing = 1.0
                        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                        run = p.add_run("\t" + para.strip())
                        run.font.name = 'Times New Roman'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run.font.size = Pt(12)
                        shade_paragraph(p)
            elif block == "Блок8":
                
                # bold_words = ["Председатель Правления", "Заместитель Председателя Правления", "Советник Председателя Правления", "Управляющий директор", "ПМ"]  # add as needed
                # for para in paragraphs:
                #     #p = formatted_doc.add_paragraph(para)
                #     #bold_keywords(para, bold_words) 

                #     parts = re.split(r"\t+|\s{2,}", para)
                #     if len(parts) >= 2:
                #         #add_signature_table(parts[0].strip(), parts[1].strip())
                #         print(parts[0].strip())
                #         print("parts")
                #         print(parts[1].strip())
                #         role = parts[0].strip()
                #         name = parts[1].strip()
                #         staff = [(role, name)]
                #         print(len(staff))
                #         for role, name in staff:
                #             p = formatted_doc.add_paragraph()
                #             #bold_keywords(p, bold_words)
                            
                #             # Add tab stop at 4 inches (adjust for your layout)
                #             tab_stops = p.paragraph_format.tab_stops
                #             tab_stops.add_tab_stop(Inches(5.0), alignment=WD_TAB_ALIGNMENT.LEFT)
                            
                #             # Add text with tab
                #             run = p.add_run(f"{role}\t{name}")
                            
                #             # Set font to Times New Roman
                #             run.font.name = "Times New Roman"
                #             run.font.size = Pt(12)

                #             if role in bold_words:
                #                 run.bold = True
                #             # Optional: Align left
                #             p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                #     else:
                #         p = formatted_doc.add_paragraph(para)
                #         bold_keywords(p, bold_words)
                #         set_format(p)
                #--------------------------------------
                # bold_words = [
                #         "Председатель Правления", 
                #         "Заместитель Председателя Правления", 
                #         "Советник Председателя Правления", 
                #         "Управляющий директор", 
                #         "ПМ"
                # ]

                # for para in paragraphs:
                #     #para = clean_text(para)
                #     parts = re.split(r"\t+|\s{2,}", para)

                #     if len(parts) >= 2:
                #         role = parts[0].strip()
                #         name = parts[1].strip()
                #         p = formatted_doc.add_paragraph()

                #         # Устанавливаем табуляцию
                #         tab_stops = p.paragraph_format.tab_stops
                #         tab_stops.add_tab_stop(Inches(4.5), alignment=WD_TAB_ALIGNMENT.LEFT)

                #         # Добавляем роль с выборочным выделением жирным
                #         i = 0
                #         while i < len(role):
                #             matched = False
                #             for bw in bold_words:
                #                 if role[i:].startswith(bw):
                #                     run = p.add_run(bw)
                #                     run.bold = True
                #                     run.font.name = "Times New Roman"
                #                     run.font.size = Pt(12)
                #                     i += len(bw)
                #                     matched = True
                #                     break
                #             if not matched:
                #                 run = p.add_run(role[i])
                #                 run.font.name = "Times New Roman"
                #                 run.font.size = Pt(12)
                #                 i += 1

                #         # Добавляем табуляцию и имя
                #         p.add_run("\t")
                #         run_name = p.add_run(name)
                #         run_name.font.name = "Times New Roman"
                #         run_name.font.size = Pt(12)
                #         run_name.font.bold = True
                #         p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                #     else:
                #         p = formatted_doc.add_paragraph(para)
                #         run_name = p.add_run()
                #         run_name.font.bold = True
                #         bold_keywords(p, bold_words)
                #         set_format(p)
            #------------------------------------------------
                bold_words = [
                    "Председатель Правления", 
                    "Заместитель Председателя Правления", 
                    "Советник Председателя Правления", 
                    "Управляющий директор", 
                    "ПМ"
                ]
                formatted_doc.add_paragraph()
                formatted_doc.add_paragraph()
                formatted_doc.add_paragraph()
                formatted_doc.add_paragraph()
                formatted_doc.add_paragraph()
                for para in paragraphs:
                    para = para.strip()
                    matched_role = None

                    # Check if paragraph starts with one of the roles
                    for bw in bold_words:
                        if para.startswith(bw):
                            matched_role = bw
                            break

                    if matched_role:
                        name = para[len(matched_role):].strip()  # extract the rest as name
                        p = formatted_doc.add_paragraph()

                        # Set tab stop
                        tab_stops = p.paragraph_format.tab_stops
                        tab_stops.add_tab_stop(Inches(4.5), alignment=WD_TAB_ALIGNMENT.LEFT)

                        # Add role (bold)
                        run_role = p.add_run(matched_role)
                        run_role.font.name = "Times New Roman"
                        run_role.font.size = Pt(12)
                        run_role.bold = True

                        # Add tab + name (bold)
                        p.add_run("\t")
                        run_name = p.add_run(name)
                        run_name.font.name = "Times New Roman"
                        run_name.font.size = Pt(12)
                        run_name.bold = True

                        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                    else:
                        # No role matched — just output as is
                        p = formatted_doc.add_paragraph(para)
                        run = p.runs[0] if p.runs else p.add_run()
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(12)
                        run.bold = True
                        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            elif block == "Блок9":
                formatted_doc.add_paragraph()
                for para in paragraphs: 
                    para = clean_text(para)
                    if para:
                        p = formatted_doc.add_paragraph()
                        run = p.add_run(para)
                        apply_format(p,10,False,WD_PARAGRAPH_ALIGNMENT.LEFT)
                formatted_doc.add_paragraph()
            elif block == "Блок10":
                p = formatted_doc.add_paragraph()
                run = p.add_run("Приложения")
                font = run.font
                font.name = 'Times New Roman'
                font.size = Pt(11)
                font.bold = True
                for para in paragraphs:
                    para = clean_text(para)
                    if para:
                        p = formatted_doc.add_paragraph(para, style='List Number') 
                    # set_format(p)
                        run = p.runs[0] if p.runs else p.add_run()
                        font = run.font
                        font.name = 'Times New Roman'
                        font.size = Pt(11)
                        font.bold = False
                        font.color.rgb = RGBColor(0, 0, 0)
                        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                        apply_typographic_fixes(p.text)

            else:
                for para in paragraphs:
                    para = clean_text(para)
                    if para:
                        fixed_text = apply_typographic_fixes(para) 
                        p = formatted_doc.add_paragraph(para)
                        set_format(p)
                        #apply_typographic_fixes(p.text)

        # === Save result ===
        #formatted_doc.save(output_path)
        #print(f"✅ Reformatted document saved to: {output_path}")
    if doc.paragraphs and doc.paragraphs[0].text.strip() == "БЮЛЛИТЕНЬ":
        
        
        
        print("here")
    insert_page_numbers_except_first(formatted_doc)
    insert_page_numbers_except_first(doc)
    buffer = io.BytesIO()
    formatted_doc.save(buffer)
    buffer.seek(0)
    st.download_button("Download Reformatted DOCX", data=buffer, file_name="Reformatted_Change.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
