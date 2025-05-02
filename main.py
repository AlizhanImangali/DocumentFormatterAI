import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT,WD_LINE_SPACING
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

    # === Split paragraphs by blocks ===
    blocks = {}
    current_block = None
    for para in doc.paragraphs:
        text = para.text.strip()
        block_header = re.match(r"^Блок(\d+)", text)
        if block_header:
            current_block = f"Блок{block_header.group(1)}"
            blocks[current_block] = []
        elif current_block:
            blocks[current_block].append(text)

    # === Process blocks with specific styles ===

    for block, paragraphs in blocks.items():
        formatted_doc.add_paragraph(block).runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red block header

        if block == "Блок1":
            text = "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА"
            p = formatted_doc.add_paragraph()
            run = p.add_run(text)
            set_character_spacing(run,60)
            apply_format(p,14,True,WD_PARAGRAPH_ALIGNMENT.CENTER)
            for para in paragraphs:
                p = formatted_doc.add_paragraph(para)
                set_format(p, 12, True, WD_PARAGRAPH_ALIGNMENT.CENTER)
                apply_typographic_fixes(p.text)
        elif block == "Блок2":
            text = "Условные (сокращенные) обозначения, использованные в пояснительной записке"
            p = formatted_doc.add_paragraph()
            run = p.add_run(text)
            apply_format(p,11,True,WD_PARAGRAPH_ALIGNMENT.LEFT)
            for para in paragraphs:
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
            table = formatted_doc.add_table(rows=1, cols=1)
            cell = table.cell(0, 0)
            set_cell_shading(cell, 'D9D9D9')

            for para in paragraphs:
                p = cell.add_paragraph(para)
                run = p.runs[0] if p.runs else p.add_run()
                font = run.font
                font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                font.size = Pt(11)
                font.bold = False
                font.color.rgb = RGBColor(0, 0, 0)

                pf = p.paragraph_format
                pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                pf.space_after = Pt(3)
                pf.space_before = Pt(0)
                pf.line_spacing = 1.0
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

                if para.strip() == "Основание выноса вопроса на рассмотрение Советом директоров":
                    font.bold = True

                #if para == "Основание выноса вопроса на рассмотрение Советом директоров":
                #  font.bold = True
        elif block == "Блок4":
            for para in paragraphs: 
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
        
            for para in paragraphs:
                #para = apply_typographic_fixes(para)
                # Обработка только если есть тире
                if "–" in para or "-" in para:
                    dash = "–" if "–" in para else "-"
                    parts = para.split(dash, 1)
                    if len(parts) == 2:
                        term, desc = parts
                        p = formatted_doc.add_paragraph()

                        #tabs = p.paragraph_format.tab_stops
                        #tabs.add_tab_stop(Inches(1.0))
                        # Настройка абзаца
                        pf = p.paragraph_format
                       # pf.first_line_indent = Inches(0.3)                 # Отступ первой строки
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(3)
                       # pf.left_indent = Inches(0.3)
                        pf.line_spacing = 1.0
                        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                        # Часть до тире
                        run1 = p.add_run(" –\t" + term.strip() )#+ " –\t" + desc.strip())
                        run1.font.name = 'Times New Roman'
                        run1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run1.font.size = Pt(11)
                        #run.font.bold = True

                        # Часть после тире
                        run2 = p.add_run(desc.strip())
                        run2.font.name = 'Times New Roman'
                        run2._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                        run2.font.size = Pt(11)
                        #run2.font.bold = Falses
                else:
                    # Без тире – обычный абзац
                    p = formatted_doc.add_paragraph(para)
                    pf = p.paragraph_format
                    #pf.first_line_indent = Inches(0.3)
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(3)
                    pf.line_spacing = 1.0
                    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    pf.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                    run = p.add_run()
                    run = p.runs[0] if p.runs else p.add_run()
                    run.text = para
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    run.font.size = Pt(11)


        elif block == "Блок6":
            for para in paragraphs: 
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
                pf.space_before = Pt(6)
                pf.line_spacing = 1.0
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

        elif block == "Блок7":
            for para in paragraphs:
                formatted_doc.add_paragraph()
                if re.match(r"^\d+\.", para):
                    p = formatted_doc.add_paragraph(para)
                    p.text = fix_docx_numbering(p.text)
                    apply_typographic_fixes(p.text)
                    
                else:
                    p = formatted_doc.add_paragraph(para)
                set_format(p)

        elif block == "Блок8":
            for para in paragraphs:
                parts = re.split(r"\t+|\s{2,}", para)
                if len(parts) >= 2:
                    add_signature_table(parts[0].strip(), parts[1].strip())
                else:
                    p = formatted_doc.add_paragraph(para)
                    set_format(p)

        #elif block == "Блок8":
        #    for para in paragraphs:
        #         p = formatted_doc.add_paragraph()
        #         apply_format(p,10,False,WD_PARAGRAPH_ALIGNMENT.LEFT)
        
        elif block == "Блок9":
            for para in paragraphs: 
                if para:
                    p = formatted_doc.add_paragraph()
                    run = p.add_run(para)
                    apply_format(p,10,False,WD_PARAGRAPH_ALIGNMENT.LEFT)

        elif block == "Блок10":
            p = formatted_doc.add_paragraph()
            run = p.add_run("Приложения")
            font = run.font
            font.name = 'Times New Roman'
            font.size = Pt(11)
            font.bold = True
            for para in paragraphs:
                if para:
                    p = formatted_doc.add_paragraph(para, style='List Number') 
                    set_format(p)
                    apply_typographic_fixes(p.text)

        else:
            for para in paragraphs:
                if para:
                    fixed_text = apply_typographic_fixes(para) 
                    p = formatted_doc.add_paragraph(para)
                    set_format(p)
                    #apply_typographic_fixes(p.text)

    # === Save result ===
    #formatted_doc.save(output_path)
    #print(f"✅ Reformatted document saved to: {output_path}")
    buffer = io.BytesIO()
    formatted_doc.save(buffer)
    buffer.seek(0)
    st.download_button("Download Reformatted DOCX", data=buffer, file_name="Reformatted_Change.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
