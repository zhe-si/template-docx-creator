from docx import Document
from docx.text import font
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run


def add_style(t_d: Document, f_s_n: str, f_d=Document()):
    _style = f_d.styles[f_s_n]
    t_d.styles.element.append(_style.element)


def copy_font(from_font: font, to_font: font):
    to_font.name = from_font.name
    to_font.bold = from_font.bold
    to_font.italic = from_font.italic
    to_font.underline = from_font.underline
    to_font.size = from_font.size
    to_font.color.rgb = from_font.color.rgb
    to_font.strike = from_font.strike
    to_font.outline = from_font.outline
    to_font.highlight_color = from_font.highlight_color
    to_font.shadow = from_font.shadow
    to_font.imprint = from_font.imprint
    to_font.double_strike = from_font.double_strike
    to_font.subscript = from_font.subscript
    to_font.superscript = from_font.superscript


def copy_run_style(from_run: Run, to_run: Run):
    """拷贝run的样式"""
    to_run.style = from_run.style
    to_run.bold = from_run.bold
    to_run.italic = from_run.italic
    to_run.underline = from_run.underline
    copy_font(from_run.font, to_run.font)


def copy_paragraph_format(to_paragraph: ParagraphFormat, from_paragraph: ParagraphFormat):
    """拷贝paragraphFormat的格式，主要影响段落对其方式"""
    to_paragraph.alignment = from_paragraph.alignment
    to_paragraph.first_line_indent = from_paragraph.first_line_indent
    to_paragraph.keep_together = from_paragraph.keep_together
    to_paragraph.keep_with_next = from_paragraph.keep_with_next
    to_paragraph.left_indent = from_paragraph.left_indent
    to_paragraph.line_spacing = from_paragraph.line_spacing
    to_paragraph.line_spacing_rule = from_paragraph.line_spacing_rule
    to_paragraph.page_break_before = from_paragraph.page_break_before
    to_paragraph.right_indent = from_paragraph.right_indent
    to_paragraph.space_after = from_paragraph.space_after
    to_paragraph.space_before = from_paragraph.space_before


def copy_paragraph_style(to_paragraph: Paragraph, from_paragraph: Paragraph, from_run: Run = None):
    """拷贝paragraph的样式"""
    if from_run is None:
        from_run = from_paragraph.runs[0]
    to_paragraph.style = from_paragraph.style
    to_paragraph.alignment = from_paragraph.alignment
    copy_paragraph_format(to_paragraph.paragraph_format, from_paragraph.paragraph_format)
    for run in to_paragraph.runs:
        copy_run_style(from_run, run)


def delete_paragraph(paragraph: Paragraph):
    """删除段落"""
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None
