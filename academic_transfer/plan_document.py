# -*- coding: utf-8 -*-
"""Формирование индивидуального плана (.docx) без внешнего шаблона."""
from __future__ import annotations

import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

DEFAULT_PLAN_META = {
    'student_genitive': 'Егорова Никиты Сергеевича',
    'student_sign_fio': 'Егоров Н.С.',
    'direction': '09.03.02 «Информационные системы и технологии»',
    'group': 'ЦИС-36',
    'profile': 'Информационные системы и технологии',
    'deadline': 'до «1» февраля 2027 г.',
    'sign_year': '2026',
    'deputy_title': 'Заместитель директора Института цифровых систем',
    'deputy_fio': 'Бойков С.Ю.',
    'prorector_fio': 'В.А. Голкина',
}

def _run_font(run, size_pt=12, bold=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.bold = bold


def _para(doc, text, bold=False, align=None, size_pt=12):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    _run_font(r, size_pt=size_pt, bold=bold)
    return p


def _cell_write(cell, text, size_pt=11):
    cell.text = ''
    p = cell.paragraphs[0]
    r = p.add_run(str(text) if text is not None else '')
    _run_font(r, size_pt=size_pt, bold=False)


def _fmt_ze(h):
    if h is None or h == '':
        return '—'
    try:
        if float(h) == int(float(h)):
            return str(int(float(h)))
        return str(h)
    except (TypeError, ValueError):
        return str(h)


def _fmt_sem(v) -> str:
    if v is None or v == '':
        return '—'
    try:
        if isinstance(v, float) and not v == int(v):
            return str(v)
        if float(v) == int(float(v)):
            return str(int(float(v)))
        return str(v)
    except (TypeError, ValueError):
        s = str(v).strip()
        return s if s else '—'


def _fmt_grade_cell(grade: Any, normalized: Any = None) -> str:
    if normalized is not None:
        if normalized in (5, 4, 3, 2):
            return str(normalized)
        if normalized == 'зачет':
            return 'зачтено'
        if normalized == 'не зачет':
            return 'не зачтено'
    if grade is None or grade == '':
        return '—'
    s = str(grade).strip()
    sl = s.lower()
    if sl in ('не сдано', 'не зачтено', 'не зачет') or 'не зач' in sl:
        return 'не зачтено'
    if 'удовлетвор' in sl:
        return '3'
    if 'хорошо' in sl and 'удовлетвор' not in sl:
        return '4'
    if 'отличн' in sl:
        return '5'
    if 'зачт' in sl or sl == 'зачет':
        return 'зачтено'
    if s.isdigit():
        return s
    if len(s) > 24:
        return s[:21] + '…'
    return s


def _to_float(val: Any) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '.')
    if not s or s == '—':
        return 0.0
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    if not m:
        return 0.0
    try:
        return float(m.group(0))
    except ValueError:
        return 0.0


def _sum_key(rows: List[Dict[str, Any]], key: str) -> str:
    total = sum(_to_float(r.get(key)) for r in rows)
    return _fmt_ze(total)


def _split_deadline(deadline: str) -> Tuple[str, str, str]:
    """
    Преобразует строку 'до «1» февраля 2027 г.' -> ('1', 'февраля', '2027').
    """
    if not deadline:
        return '', '', ''
    dday, dmonth, dyear = '', '', ''

    q = re.search(r'«\s*([^»]+?)\s*»', deadline)
    if q:
        dday = q.group(1).strip()

    year_m = re.search(r'(\d{4})', deadline)
    if year_m:
        dyear = year_m.group(1)

    # Берем слово между закрывающей кавычкой и годом, если есть
    after_quote = deadline.split('»', 1)[1] if '»' in deadline else deadline
    parts = [p for p in re.split(r'\s+', after_quote.replace('г.', '').replace('г', '').strip()) if p]
    # Обычно: [<месяц>, <год>] или [до, <день>, <месяц>, <год>]
    for p in parts:
        if re.fullmatch(r'\d{4}', p):
            break
        if not p.isdigit():
            dmonth = p
            break
    return dday, dmonth, dyear


def _apply_ygstu_table_merges(table):
    n = len(table.rows)
    if n < 15:
        return
    try:
        table.cell(8, 4).merge(table.cell(9, 4))
        table.cell(8, 5).merge(table.cell(9, 5))
        table.cell(10, 4).merge(table.cell(11, 4))
        table.cell(10, 5).merge(table.cell(11, 5))
        table.cell(13, 0).merge(table.cell(13, 1))
    except (IndexError, ValueError):
        pass


def _set_table_columns_6(table):
    table.autofit = False
    for row in table.rows:
        row.cells[0].width = Cm(0.9)
        row.cells[1].width = Cm(7.2)
        row.cells[2].width = Cm(1.4)
        row.cells[3].width = Cm(1.4)
        row.cells[4].width = Cm(2.4)
        row.cells[5].width = Cm(2.7)


def _add_data_table_programmatic(doc, headers, body_rows, min_body_rows=14):
    n_body = max(len(body_rows), min_body_rows)
    table = doc.add_table(rows=1 + n_body, cols=6)
    table.style = 'Table Grid'

    hdr = table.rows[0].cells
    for j, h in enumerate(headers):
        _cell_write(hdr[j], h, size_pt=10)
        hdr[j].paragraphs[0].runs[0].bold = True

    for i, row_data in enumerate(body_rows):
        row = table.rows[i + 1].cells
        for j, val in enumerate(row_data):
            _cell_write(row[j], val, size_pt=11)

    for i in range(len(body_rows), n_body):
        row = table.rows[i + 1].cells
        for j in range(6):
            _cell_write(row[j], '', size_pt=11)

    _apply_ygstu_table_merges(table)
    _set_table_columns_6(table)
    doc.add_paragraph()


def _build_programmatic(final_results: Dict[str, Any], meta: Dict[str, Any]) -> BytesIO:
    m = {**DEFAULT_PLAN_META, **(meta or {})}
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    approve_lines = [
        'УТВЕРЖДАЮ',
        'Проректор по',
        'образовательной',
        'деятельности и',
        'воспитательной работе',
        f'___________    {m.get("prorector_fio", "В.А. Голкина")}',
        '«____» ____________ 20__ г.',
    ]
    for line in approve_lines:
        pr = doc.add_paragraph()
        pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rr = pr.add_run(line)
        _run_font(rr, size_pt=12)

    doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_title.add_run('Индивидуальный план\nликвидации разницы в учебных планах')
    _run_font(r1, size_pt=14, bold=True)

    doc.add_paragraph()

    line2 = (
        f'{m["student_genitive"]}, обучающегося по направлению подготовки '
        f'{m["direction"]} группы {m["group"]}'
    )
    _para(doc, line2, size_pt=12)
    _para(doc, f'Направленность (профиль): «{m["profile"]}»', size_pt=12)
    _para(doc, f'Срок ликвидации разницы в учебных планах: {m["deadline"]}', size_pt=12)

    doc.add_paragraph()
    _para(doc, 'Перечень перезачитываемых дисциплин или практик:', bold=True, size_pt=12)
    doc.add_paragraph()

    rec_rows = []
    n = 1
    headers = (
        '№\nп/п',
        'Наименование дисциплины, практики',
        'Всего зачётных единиц',
        'Семестр',
        'Форма промежуточной и итоговой аттестации, отметка',
        'Отметка работодателя (при необходимости)',
    )

    for item in final_results.get('recreditable', []):
        note = f'Зачтено по УП: {item.get("matched_to", "—")}'
        rec_rows.append((
            str(n),
            item.get('name', '—'),
            _fmt_ze(item.get('hours')),
            _fmt_sem(item.get('semester')),
            _fmt_grade_cell(item.get('grade'), item.get('normalized_grade')),
            note
        ))
        n += 1

    for item in final_results.get('reattestation', []):
        note = f'По УП: {item.get("matched_to", "—")} (переаттестация)'
        rec_rows.append((
            str(n),
            item.get('name', '—'),
            _fmt_ze(item.get('hours')),
            _fmt_sem(item.get('semester')),
            _fmt_grade_cell(item.get('grade'), item.get('normalized_grade')),
            note
        ))
        n += 1

    if not rec_rows:
        rec_rows.append(('1', '—', '—', '—', '—', 'Нет записей по данным сопоставления'))

    _add_data_table_programmatic(doc, headers, rec_rows, min_body_rows=14)

    _para(doc, 'Перечень дисциплин или практик, подлежащих изучению или прохождению', bold=True, size_pt=12)
    doc.add_paragraph()

    study_rows = []
    n = 1
    for item in final_results.get('need_study', []):
        study_rows.append((
            str(n),
            item.get('name', '—'),
            _fmt_ze(item.get('hours')),
            '—',
            '—',
            'К изучению / прохождению',
        ))
        n += 1

    if not study_rows:
        study_rows.append(('1', '—', '—', '—', '—', 'Нет записей по данным сопоставления'))

    _add_data_table_programmatic(doc, headers, study_rows, min_body_rows=14)

    doc.add_paragraph()
    _para(doc, 'СОГЛАСОВАНО:', bold=True, size_pt=12)
    doc.add_paragraph()

    y = m.get('sign_year', '2026')
    st = m.get('student_sign_fio', '_______________')
    dep_t = m.get('deputy_title', '')
    dep_f = m.get('deputy_fio', '_______________')

    sig_block = (
        f'Обучающийся\n'
        f'«____» ____________ {y} г.\t\t_________________\t\t{st}\n\n'
        f'{dep_t}\n'
        f'«____» ____________ {y} г.\t\t_________________\t\t{dep_f}'
    )
    _para(doc, sig_block, size_pt=12)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def build_individual_plan_docx(final_results: Dict[str, Any], meta: Optional[Dict[str, Any]] = None) -> BytesIO:
    return _build_programmatic(final_results, meta or {})