from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .schema import CaseData

FONT = 'Times New Roman'
FONT_SIZE = 12


def _set_run_font(run, bold: bool | None = None):
    run.font.name = FONT
    run.font.size = Pt(FONT_SIZE)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        rfonts.set(qn(f'w:{attr}'), FONT)
    if bold is not None:
        run.bold = bold


def _normalize_court_district(value: str) -> str:
    """Clean spacing while preserving the court's official wording verbatim."""
    return re.sub(r'\s+', ' ', (value or '').strip()).rstrip('.,; ')


def _normalize_order_subject(value: str) -> str:
    text = re.sub(r'\s+', ' ', (value or '').strip()).rstrip('.; ')
    if text.casefold().startswith('о '):
        text = text[2:].lstrip()
    if text.casefold().startswith('взыскание'):
        text = 'взыскании' + text[len('взыскание'):]
    elif text.casefold().startswith('задолженности'):
        remainder = text[len('задолженности'):]
        text = 'взыскании задолженности' + remainder
    elif text.casefold().startswith('задолженность'):
        remainder = text[len('задолженность'):]
        text = 'взыскании задолженности' + remainder
    return text


def _capitalize_sentence(text: str) -> str:
    text = (text or '').strip()
    if not text:
        return text
    return text[0].upper() + text[1:]


def _ensure_sentence_end(text: str) -> str:
    text = (text or '').strip()
    if text and text[-1] not in '.!?':
        text += '.'
    return text


def _document_copy_title(item: str) -> str:
    """Produce a readable attachment title that explicitly denotes a copy."""
    text = re.sub(r'\s+', ' ', (item or '').strip()).rstrip('.; ')
    if not text:
        return ''
    text = re.sub(r'\bгос\s*услуг\b', 'Госуслуг', text, flags=re.IGNORECASE)
    text = re.sub(r'\bс Госуслуг\b', 'из Госуслуг', text, flags=re.IGNORECASE)

    if re.match(r'^копия\b', text, re.IGNORECASE):
        return _ensure_sentence_end(_capitalize_sentence(text))
    if re.match(r'^копии\b', text, re.IGNORECASE):
        return _ensure_sentence_end(_capitalize_sentence(text))

    lower = text.casefold()
    replacements = [
        ('уведомление ', 'уведомления '),
        ('договор ', 'договора '),
        ('постановление ', 'постановления '),
        ('скриншот ', 'скриншота '),
        ('выписка ', 'выписки '),
        ('справка ', 'справки '),
        ('почтовое уведомление ', 'почтового уведомления '),
        ('документ ', 'документа '),
    ]
    for start, changed in replacements:
        if lower.startswith(start):
            text = changed + text[len(start):]
            break
    return _ensure_sentence_end('Копия ' + text)


def _disagreement_sentence(basis: str) -> str:
    text = re.sub(r'\s+', ' ', (basis or '').strip()).rstrip('.; ')
    if not text:
        return 'Основанием для несогласия с судебным приказом является моё несогласие с заявленным требованием.'
    if re.match(r'^не согласен\b', text, re.IGNORECASE):
        return _ensure_sentence_end('Основанием для несогласия с судебным приказом является то, что я ' + text)
    if re.match(r'^не согласна\b', text, re.IGNORECASE):
        return _ensure_sentence_end('Основанием для несогласия с судебным приказом является то, что я ' + text)
    if re.match(r'^(с наличием|с размером|с основанием|с обоснованностью)\b', text, re.IGNORECASE):
        return _ensure_sentence_end('Основанием для несогласия с судебным приказом является моё несогласие ' + text)
    if text.casefold().startswith('тем, что '):
        return _ensure_sentence_end('Основанием для несогласия с судебным приказом является ' + text)
    return _ensure_sentence_end('Основанием для несогласия с судебным приказом считаю следующее: ' + text)


def _set_paragraph(para, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent_cm=1.25,
                   left_indent_cm=0, right_indent_cm=0, line_spacing=1.5,
                   before_pt=0, after_pt=6, keep_together=False):
    pf = para.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_indent_cm) if first_indent_cm is not None else None
    pf.left_indent = Cm(left_indent_cm) if left_indent_cm else None
    pf.right_indent = Cm(right_indent_cm) if right_indent_cm else None
    pf.line_spacing = line_spacing
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    pf.keep_together = keep_together
    pf.widow_control = True


def _set_table_borders_none(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn('w:tblBorders'))
    if borders is None:
        borders = OxmlElement('w:tblBorders')
        tbl_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = 'w:' + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn('w:val'), 'nil')


def _set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tc_mar.find(qn(f'w:{m}'))
        if node is None:
            node = OxmlElement(f'w:{m}')
            tc_mar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')


def generate_order_cancellation(case: CaseData, output: str | Path) -> str:
    """Generate a formatted Russian court filing from validated CaseData."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(1.5)

    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(FONT_SIZE)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)
    normal_rpr = normal.element.get_or_add_rPr()
    normal_rfonts = normal_rpr.rFonts
    if normal_rfonts is None:
        normal_rfonts = OxmlElement('w:rFonts')
        normal_rpr.insert(0, normal_rfonts)
    for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        normal_rfonts.set(qn(f'w:{attr}'), FONT)

    def add_para(text='', *, bold=False, kind='body', keep=False):
        para = doc.add_paragraph()
        run = para.add_run(text)
        _set_run_font(run, bold=bold)
        if kind == 'header':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent_cm=0,
                           line_spacing=1.0, before_pt=0, after_pt=1, keep_together=True)
        elif kind == 'header_gap':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent_cm=0,
                           line_spacing=1.0, before_pt=0, after_pt=2)
        elif kind == 'title':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
                           line_spacing=1.0, before_pt=2, after_pt=3, keep_together=True)
        elif kind == 'request_heading':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
                           line_spacing=1.5, before_pt=3, after_pt=4, keep_together=True)
        elif kind == 'numbered':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent_cm=-0.65,
                           left_indent_cm=0.65, line_spacing=1.5, before_pt=0, after_pt=3)
        elif kind == 'attachment':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=-0.55,
                           left_indent_cm=0.55, line_spacing=1.5, before_pt=0, after_pt=1)
        elif kind == 'attachments_heading':
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
                           line_spacing=1.5, before_pt=6, after_pt=3, keep_together=True)
        else:
            _set_paragraph(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent_cm=1.25,
                           line_spacing=1.5, before_pt=0, after_pt=3, keep_together=keep)
        return para

    district = _normalize_court_district(case.court_district)
    subject = _normalize_order_subject(case.order_subject)
    court_line = f'СУДЕБНОГО УЧАСТКА № {case.court_number}'
    # Court/party header: right-aligned and compact, as requested.
    add_para('МИРОВОМУ СУДЬЕ', kind='header', bold=True)
    add_para(court_line, kind='header', bold=True)
    if district:
        add_para(district, kind='header')
    if case.court_address:
        add_para(case.court_address, kind='header')
    add_para('', kind='header_gap')

    add_para(f'Заявитель: {case.applicant_name}', kind='header', bold=True)
    for label, value in [
        ('Дата рождения', case.birth_date),
        ('Паспорт', case.passport),
        ('Выдан', case.passport_issued),
        ('ИНН', case.inn),
        ('СНИЛС', case.snils),
        ('Адрес', case.applicant_address),
        ('Тел.', case.phone),
        ('E-mail', case.email),
    ]:
        if (value or '').strip():
            add_para(f'{label}: {value.strip()}', kind='header')
    add_para('', kind='header_gap')
    add_para(f'Взыскатель: {case.creditor_name}', kind='header', bold=True)
    if case.creditor_inn:
        add_para(f'ИНН: {case.creditor_inn}', kind='header')
    if case.creditor_address:
        add_para(f'Адрес: {case.creditor_address}', kind='header')
    add_para('', kind='header_gap')
    if case.case_number:
        add_para(f'Дело № {case.case_number}', kind='header')
    add_para('', kind='header_gap')

    add_para('ЗАЯВЛЕНИЕ', kind='title', bold=True)
    add_para('ОБ ОТМЕНЕ СУДЕБНОГО ПРИКАЗА', kind='title', bold=True)
    add_para('')

    add_para(
        f'{case.order_date} в отношении меня мировым судьёй судебного участка № {case.court_number} '
        f'{district} по делу № {case.case_number} был вынесен судебный приказ о {subject}.'
    )
    if case.ip_date or case.ip_number:
        ip_parts = []
        if case.ip_date:
            ip_parts.append(case.ip_date)
        if case.ip_number:
            ip_parts.append(f'№ {case.ip_number}')
        add_para('На основании указанного судебного приказа в отношении меня ' + ' '.join(ip_parts) + ' возбуждено исполнительное производство.')

    if case.awareness_date:
        add_para(f'О наличии судебного приказа мне стало известно {case.awareness_date.strip()}.')
    if case.awareness_source:
        source = _ensure_sentence_end(_capitalize_sentence(case.awareness_source.strip()))
        add_para(f'Источник информации: {source}')

    if case.missed_reason:
        reason = _ensure_sentence_end(_capitalize_sentence(case.missed_reason.strip()))
        add_para('Обстоятельства пропуска срока: ' + reason)
    if case.supporting_documents:
        add_para('В подтверждение изложенных обстоятельств прилагаю копии документов, перечисленных в разделе «Приложения (копии)».')

    add_para('В соответствии со ст. 112 ГПК РФ лицам, пропустившим установленный федеральным законом процессуальный срок по причинам, признанным судом уважительными, пропущенный срок может быть восстановлен.')
    add_para('На основании ст. 129 ГПК РФ должник вправе обжаловать судебный приказ в течение 10 дней с момента его получения.')
    add_para(f'Прошу исчислять срок для подачи возражений относительно исполнения судебного приказа с {case.effective_start_date}.')
    add_para('Настоящим заявляю о своём несогласии с судебным приказом.')
    add_para(_disagreement_sentence(case.disagreement_basis))
    add_para('На основании изложенного и руководствуясь статьями 112, 128, 129 ГПК РФ,')
    add_para('ПРОШУ:', kind='request_heading', bold=True)
    add_para(f'1. Восстановить срок на обжалование судебного приказа мирового судьи судебного участка № {case.court_number} {district} по делу № {case.case_number}.', kind='numbered')
    add_para(f'2. Отменить судебный приказ от {case.order_date} по делу № {case.case_number} о {subject}, вынесенный мировым судьёй судебного участка № {case.court_number} {district}.', kind='numbered')

    corr = case.correspondence_address or case.applicant_address
    if case.fssp_office:
        fssp = case.fssp_office.strip().rstrip('.,; ')
        add_para(f'3. Копию определения об отмене судебного приказа направить в {fssp} и заявителю по адресу: {corr}.', kind='numbered')
    else:
        add_para(f'3. Копию определения об отмене судебного приказа направить в соответствующее подразделение ФССП России и заявителю по адресу: {corr}.', kind='numbered')

    add_para('Приложения (копии):', kind='attachments_heading', bold=True)
    clean_docs = []
    for item in case.supporting_documents:
        title = _document_copy_title(item)
        if title and title not in clean_docs:
            clean_docs.append(title)
    for number, item in enumerate(clean_docs, start=1):
        add_para(f'{number}. {item}', kind='attachment')

    add_para('', kind='header_gap')
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(5.0)
    table.columns[1].width = Cm(11.5)
    _set_table_borders_none(table)
    left, right = table.rows[0].cells
    left.width = Cm(5.0)
    right.width = Cm(11.5)
    for cell in (left, right):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_cell_margins(cell)
    date_text = case.composition_date.strip() if case.composition_date.strip() else datetime.now().strftime('%d.%m.%Y')
    p_date = left.paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_date.paragraph_format.space_after = Pt(0)
    p_date.paragraph_format.line_spacing = 1.0
    r = p_date.add_run(f'Дата: {date_text}')
    _set_run_font(r)
    p_sign = right.paragraphs[0]
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.space_after = Pt(0)
    p_sign.paragraph_format.line_spacing = 1.0
    r = p_sign.add_run(f'________________ / {case.applicant_name}')
    _set_run_font(r)

    Path(output).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output))
    return str(output)
