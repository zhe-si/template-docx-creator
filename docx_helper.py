from docx import Document
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, CT_R, CT_P
from docx.oxml.ns import qn
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


def delete_paragraph(paragraph: [Paragraph, Run]):
    """删除段落"""
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def add_hyperlink(paragraph: Paragraph, text: str, url: str):
    """给段落末尾添加超链接

    :param paragraph: 追加的段落
    :param text: 文本
    :param url: 链接
    """
    r_id = paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)  # 关联超链接

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    run = paragraph.add_run(text)
    run.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    run.font.underline = True
    hyperlink.append(run.element)
    paragraph._element.append(hyperlink)
    return hyperlink


def set_hyperlink(run_index: int, paragraph: Paragraph, text: str, url: str):
    """设置 run 为超链接，抹除原内容，添加超链接样式调整，其他样式保留"""
    run = paragraph.runs[run_index]

    # 在 document 的 part （rels文档资源） 中创建一个新的超链接资源并获取其 id
    r_id = paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)  # 关联超链接

    # 创建一个超链接的ooxml对象并将申请得到的超链接资源 id 作为其 id 属性
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # 创建链接内的内容 run
    new_r = OxmlElement('w:r')
    new_run = Run(new_r, paragraph)

    # 设置链接内容的文本与样式
    new_run.text = text
    new_run.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    new_run.font.underline = True

    # 将链接内容添加到超链接中
    hyperlink.append(new_r)

    # 将超链接加到内容标签的run后，并删除内容标签run
    run.element.addnext(hyperlink)
    delete_paragraph(run)

    return hyperlink
