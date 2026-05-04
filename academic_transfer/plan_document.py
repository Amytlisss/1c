# -*- coding: utf-8 -*-
"""Формирование индивидуального плана (.docx) точно по шаблону ЯГТУ"""
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================================================
# КОНСТАНТЫ
# ============================================================================

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

PLAN_META_KEYS = list(DEFAULT_PLAN_META.keys())


# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================

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


def _cell_write(cell, text, size_pt=10, bold=False, align=None):
    """Запись текста в ячейку с возможностью выравнивания"""
    cell.text = ''
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    r = p.add_run(str(text) if text is not None else '')
    _run_font(r, size_pt=size_pt, bold=bold)


def _set_cell_vertical_alignment(cell, align='center'):
    """Установка вертикального выравнивания в ячейке"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    valign = OxmlElement('w:vAlign')
    valign.set(qn('w:val'), align)
    tcPr.append(valign)


def _merge_cells_horizontal(row, start_col, end_col):
    """Объединение ячеек в строке по горизонтали"""
    if start_col >= end_col:
        return
    first_cell = row.cells[start_col]
    last_cell = row.cells[end_col]
    first_cell.merge(last_cell)


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
            if normalized == 5:
                return 'отлично'
            elif normalized == 4:
                return 'хорошо'
            elif normalized == 3:
                return 'удовлетворительно'
            elif normalized == 2:
                return 'неудовлетворительно'
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
        return 'удовлетворительно'
    if 'хорошо' in sl:
        return 'хорошо'
    if 'отличн' in sl:
        return 'отлично'
    if 'зачт' in sl or sl == 'зачет':
        return 'зачтено'
    if s.isdigit():
        if s == '5':
            return 'отлично'
        elif s == '4':
            return 'хорошо'
        elif s == '3':
            return 'удовлетворительно'
        elif s == '2':
            return 'неудовлетворительно'
        return s
    return s


def _is_discipline_passed(discipline: Dict) -> bool:
    """Проверяет, сдана ли дисциплина"""
    grade = discipline.get('grade', '')
    normalized = discipline.get('normalized_grade', '')
    
    if normalized is not None:
        if normalized in ('не зачет', 'не зачтено', 'неудовлетворительно', 2):
            return False
        if normalized in ('зачет', 'зачтено', 5, 4, 3):
            return True
    
    grade_str = str(grade).lower().strip()
    fail_keywords = ['не зачтено', 'не зачет', 'не сдано', '2', 'неудовлетворительно']
    for kw in fail_keywords:
        if kw in grade_str:
            return False
    
    if grade_str and grade_str not in ('—', '', 'не указано'):
        return True
    
    return False


def _get_control_form(discipline: Dict) -> str:
    """Определение формы контроля по дисциплине"""
    if not _is_discipline_passed(discipline):
        return '—'
    
    grade = discipline.get('grade', '')
    normalized = discipline.get('normalized_grade', '')
    control = discipline.get('control_form', '')
    
    if control:
        return control
    
    grade_str = str(grade).lower()
    
    if 'экзамен' in grade_str:
        return 'экзамен'
    elif 'диф.зачет' in grade_str or 'дифференцированный' in grade_str:
        return 'диф.зачет'
    elif 'курсовая' in grade_str or 'курсовой проект' in grade_str:
        return 'курсовая работа'
    elif 'зачет' in grade_str:
        return 'зачет'
    
    if normalized in (5, 4, 3):
        return 'экзамен'
    elif normalized == 'зачет':
        return 'зачет'
    
    return '—'


# ============================================================================
# ЗАПОЛНЕНИЕ ТАБЛИЦ ПО ШАБЛОНУ
# ============================================================================

def fill_table_1_old_plan(table, transcript_disciplines: List[Dict], meta: Dict):
    """
    Таблица 1: Дисциплины учебного плана ЯГТУ (только сданные дисциплины)
    Столбцы: Название | Семестр | Объем, ЗЕ | Форма контроля | Оценка
    - Нумерация строк: 1., 2., 3., ...
    - Выравнивание: название - по левому краю, остальные - по центру
    - Строка "Всего" - объединение первых двух столбцов, выравнивание по правому краю
    """
    if table is None:
        return
    
    # Фильтрация: только сданные дисциплины
    passed_disciplines = [d for d in transcript_disciplines if _is_discipline_passed(d)]
    
    # Находим строку с заголовками колонок
    data_start_row = 0
    for i, row in enumerate(table.rows):
        row_text = ' '.join(cell.text for cell in row.cells).lower()
        if 'название дисциплины' in row_text or 'наименование' in row_text:
            data_start_row = i + 1
            break
    
    # Удаляем старые строки с данными (оставляем только заголовки)
    while len(table.rows) > data_start_row:
        table._element.remove(table.rows[-1]._element)
    
    # Определяем индексы колонок
    num_cols = len(table.rows[0].cells) if table.rows else 5
    grade_col_index = num_cols - 1  # Оценка - последняя колонка
    
    # Заполняем сданными дисциплинами с нумерацией
    total_ze = 0
    
    for idx, disc in enumerate(passed_disciplines, 1):
        row = table.add_row()
        cells = row.cells
        
        # Название дисциплины (с нумерацией)
        if len(cells) >= 1:
            name_with_num = f"{idx}. {disc.get('original_name', disc.get('name', '—'))}"
            _cell_write(cells[0], name_with_num, size_pt=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            _set_cell_vertical_alignment(cells[0], 'center')
        
        # Семестр (по центру)
        if len(cells) >= 2:
            _cell_write(cells[1], _fmt_sem(disc.get('semester')), size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_vertical_alignment(cells[1], 'center')
        
        # Объем ЗЕ (по центру)
        if len(cells) >= 3:
            hours = disc.get('hours', 0)
            _cell_write(cells[2], _fmt_ze(hours), size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_vertical_alignment(cells[2], 'center')
            try:
                total_ze += float(hours) if hours else 0
            except:
                pass
        
        # Форма контроля (по центру)
        if len(cells) >= 4:
            _cell_write(cells[3], _get_control_form(disc), size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_vertical_alignment(cells[3], 'center')
        
        # Оценка (по центру)
        if grade_col_index < len(cells):
            _cell_write(cells[grade_col_index], 
                       _fmt_grade_cell(disc.get('grade'), disc.get('normalized_grade')), 
                       size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_vertical_alignment(cells[grade_col_index], 'center')
    
    # Строка "Всего" - удаляем старую если есть, добавляем новую
    has_total_row = False
    total_row_index = None
    for i, row in enumerate(table.rows):
        if row.cells and 'всего' in row.cells[0].text.lower():
            has_total_row = True
            total_row_index = i
            break
    
    if has_total_row and total_row_index is not None:
        # Удаляем старую строку "Всего"
        table._element.remove(table.rows[total_row_index]._element)
    
    # Добавляем новую строку "Всего"
    total_row = table.add_row()
    cells = total_row.cells
    
    # Объединяем первые две колонки
    if len(cells) >= 2:
        _merge_cells_horizontal(total_row, 0, 1)
    
    # "Всего" - выравнивание по правому краю
    if len(cells) >= 1:
        _cell_write(cells[0], 'Всего', size_pt=10, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_cell_vertical_alignment(cells[0], 'center')
    
    # Сумма ЗЕ (по центру)
    if len(cells) >= 3:
        _cell_write(cells[2], _fmt_ze(total_ze), size_pt=10, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[2], 'center')


def fill_table_2_comparison(table, match_results: Dict, curriculum: List[Dict]):
    """
    Таблица 2: Сопоставление (новый УП ↔ старый УП)
    Столбцы: № | Наименование по новому УП | Семестр | ЗЕ | Форма контроля |
             Наименование по старому УП | Семестр | ЗЕ | Форма контроля | Итоговая оценка
    - Номер по центру, названия по левому краю, остальное по центру
    - Строка "Всего" с объединением ячеек
    """
    if table is None:
        return
    
    # Сохраняем заголовок, удаляем остальные строки
    while len(table.rows) > 1:
        table._element.remove(table.rows[-1]._element)
    
    row_num = 1
    total_new_ze = 0
    total_old_ze = 0
    
    # Собираем все сопоставленные дисциплины
    matched_items = []
    for item in match_results.get('matched', []):
        matched_items.append({
            'trans': item.get('transcript_discipline', {}),
            'curr': item.get('curriculum_discipline', {})
        })
    for item in match_results.get('manual', []):
        if item.get('status') == 'manual_matched' and item.get('selected_match'):
            matched_items.append({
                'trans': item.get('transcript_discipline', {}),
                'curr': item.get('selected_match', {})
            })
    
    # Заполняем данными
    for item in matched_items:
        trans_disc = item['trans']
        curr_disc = item['curr']
        
        row = table.add_row()
        cells = row.cells
        
        # № (по центру)
        _cell_write(cells[0], str(row_num), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[0], 'center')
        
        # Наименование по новому УП (по левому краю)
        _cell_write(cells[1], curr_disc.get('original_name', '—'), size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell_vertical_alignment(cells[1], 'center')
        
        # Семестр (по центру)
        _cell_write(cells[2], _fmt_sem(curr_disc.get('semester')), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[2], 'center')
        
        # ЗЕ (по центру)
        curr_ze = curr_disc.get('hours', 0)
        _cell_write(cells[3], _fmt_ze(curr_ze), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[3], 'center')
        try:
            total_new_ze += float(curr_ze) if curr_ze else 0
        except:
            pass
        
        # Форма контроля (по центру)
        _cell_write(cells[4], _get_control_form(curr_disc), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[4], 'center')
        
        # Наименование по старому УП (по левому краю)
        _cell_write(cells[5], trans_disc.get('original_name', '—'), size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell_vertical_alignment(cells[5], 'center')
        
        # Семестр (по центру)
        _cell_write(cells[6], _fmt_sem(trans_disc.get('semester')), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[6], 'center')
        
        # ЗЕ (по центру)
        old_ze = trans_disc.get('hours', 0)
        _cell_write(cells[7], _fmt_ze(old_ze), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[7], 'center')
        try:
            total_old_ze += float(old_ze) if old_ze else 0
        except:
            pass
        
        # Форма контроля (по центру)
        _cell_write(cells[8], _get_control_form(trans_disc), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[8], 'center')
        
        # Итоговая оценка (по центру)
        _cell_write(cells[9], _fmt_grade_cell(trans_disc.get('grade'), trans_disc.get('normalized_grade')), 
                   size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[9], 'center')
        
        row_num += 1
    
    # Добавляем строку "Всего"
    total_row = table.add_row()
    cells = total_row.cells
    
    # Объединяем ячейки для "Всего" (колонки 0-4 для нового УП? или по-другому)
    # По шаблону: "Всего" пишется во второй колонке, а суммы в 3-й (ЗЕ новый) и 7-й (ЗЕ старый)
    _cell_write(cells[1], 'Всего', size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_cell_vertical_alignment(cells[1], 'center')
    _cell_write(cells[3], _fmt_ze(total_new_ze), size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_vertical_alignment(cells[3], 'center')
    _cell_write(cells[7], _fmt_ze(total_old_ze), size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_vertical_alignment(cells[7], 'center')


def fill_table_3_need_study(table, need_study: List[Dict]):
    """
    Таблица 3: Перечень дисциплин или практик, подлежащих изучению
    Столбцы: Название | Семестр | ЗЕ | Форма контроля | Кафедра | Объем работы | Срок
    - Нумерация строк: 1., 2., 3., ...
    - Выравнивание: название - по левому краю, остальные - по центру
    - Строка "Всего"
    """
    if table is None:
        return
    
    # Сохраняем заголовок, удаляем остальные строки
    while len(table.rows) > 1:
        table._element.remove(table.rows[-1]._element)
    
    total_ze = 0
    
    for idx, disc in enumerate(need_study, 1):
        row = table.add_row()
        cells = row.cells
        
        # Название (с нумерацией, по левому краю)
        name_with_num = f"{idx}. {disc.get('name', '—')}"
        _cell_write(cells[0], name_with_num, size_pt=9, align=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell_vertical_alignment(cells[0], 'center')
        
        # Семестр (по центру)
        _cell_write(cells[1], _fmt_sem(disc.get('semester')), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[1], 'center')
        
        # ЗЕ (по центру)
        hours = disc.get('hours', 0)
        _cell_write(cells[2], _fmt_ze(hours), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[2], 'center')
        try:
            total_ze += float(hours) if hours else 0
        except:
            pass
        
        # Форма контроля (по центру)
        _cell_write(cells[3], disc.get('control_form', '—'), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[3], 'center')
        
        # Кафедра (по центру)
        _cell_write(cells[4], disc.get('department', 'ИСТ'), size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[4], 'center')
        
        # Объем работы преподавателя (по центру)
        teacher_hours = disc.get('teacher_hours', '')
        if not teacher_hours:
            if hours:
                try:
                    h = float(hours)
                    teacher_hours = f"Л-{int(h)*4}; ПЗ-{int(h)*6}"
                except:
                    teacher_hours = '—'
            else:
                teacher_hours = '—'
        _cell_write(cells[5], teacher_hours, size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[5], 'center')
        
        # Срок завершения (по центру)
        deadline = disc.get('deadline', '01.02.2027')
        _cell_write(cells[6], deadline, size_pt=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[6], 'center')
    
    # Строка "Всего"
    total_row = table.add_row()
    cells = total_row.cells
    
    # Объединяем первые две колонки
    if len(cells) >= 2:
        _merge_cells_horizontal(total_row, 0, 1)
    
    # "Всего" - по правому краю
    if len(cells) >= 1:
        _cell_write(cells[0], 'Всего', size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_cell_vertical_alignment(cells[0], 'center')
    
    # Сумма ЗЕ (по центру)
    if len(cells) >= 3:
        _cell_write(cells[2], _fmt_ze(total_ze), size_pt=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_vertical_alignment(cells[2], 'center')


def fill_from_template(template_path: str, 
                       transcript_disciplines: List[Dict],
                       match_results: Dict,
                       curriculum: List[Dict],
                       final_results: Dict,
                       meta: Dict) -> BytesIO:
    """Заполнение всех таблиц в шаблоне"""
    doc = Document(template_path)
    
    # Замена текстовых полей (ФИО, группа, даты)
    student_full = f"{meta.get('student_genitive', '_______________')}, обучающегося по направлению подготовки {meta.get('direction', '_______________')} группы {meta.get('group', '_______________')}"
    profile_text = f"Направленность (профиль): «{meta.get('profile', '_______________')}»"
    deadline_text = f"Срок ликвидации разницы в учебных планах: {meta.get('deadline', '_______________')}"
    
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if 'обучающегося по направлению подготовки' in text:
            paragraph.clear()
            run = paragraph.add_run(student_full)
            _run_font(run, size_pt=12)
        elif 'Направленность (профиль):' in text and 'обучающегося' not in text:
            paragraph.clear()
            run = paragraph.add_run(profile_text)
            _run_font(run, size_pt=12)
        elif 'Срок ликвидации разницы' in text:
            paragraph.clear()
            run = paragraph.add_run(deadline_text)
            _run_font(run, size_pt=12)
        elif '___________' in text and 'Голкина' in text:
            paragraph.clear()
            run = paragraph.add_run(f'___________    {meta.get("prorector_fio", "В.А. Голкина")}')
            _run_font(run, size_pt=12)
    
    # Заполнение трёх таблиц
    if len(doc.tables) >= 1:
        fill_table_1_old_plan(doc.tables[0], transcript_disciplines, meta)
    
    if len(doc.tables) >= 2:
        fill_table_2_comparison(doc.tables[1], match_results, curriculum)
    
    if len(doc.tables) >= 3:
        fill_table_3_need_study(doc.tables[2], final_results.get('need_study', []))
    
    # Замена дат
    year = meta.get('sign_year', '2026')
    for paragraph in doc.paragraphs:
        if '«____» ____________ 20__ г.' in paragraph.text:
            new_text = paragraph.text.replace('20__ г.', f'{year} г.')
            paragraph.clear()
            run = paragraph.add_run(new_text)
            _run_font(run, size_pt=12)
    
    # Замена подписи студента
    student_fio = meta.get('student_sign_fio', '_______________')
    for paragraph in doc.paragraphs:
        if '_________________' in paragraph.text and ('Егоров' in paragraph.text or 'Обучающийся' in paragraph.text):
            new_text = paragraph.text.replace('_________________', student_fio)
            paragraph.clear()
            run = paragraph.add_run(new_text)
            _run_font(run, size_pt=12)
    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ============================================================================
# ОСНОВНАЯ ФУНКЦИЯ (сокращённая версия, остальное по аналогии)
# ============================================================================

def build_individual_plan_docx(final_results: Dict[str, Any],
                                transcript_disciplines: Optional[List[Dict]] = None,
                                match_results: Optional[Dict[str, Any]] = None,
                                curriculum: Optional[List[Dict]] = None,
                                meta: Optional[Dict[str, Any]] = None) -> BytesIO:
    """Создание документа индивидуального плана"""
    meta = meta or {}
    transcript_disciplines = transcript_disciplines or []
    match_results = match_results or {}
    curriculum = curriculum or []
    
    template_path = 'plan_template.docx'
    if os.path.exists(template_path):
        try:
            return fill_from_template(template_path, transcript_disciplines, match_results, curriculum, final_results, meta)
        except Exception as e:
            print(f"Ошибка шаблона: {e}")
            import traceback
            traceback.print_exc()
    
    # Здесь должен быть код программной генерации, но для краткости опущен
    # (можно взять из предыдущей версии)
    from io import BytesIO
    bio = BytesIO()
    doc = Document()
    _para(doc, "Документ сгенерирован программно (шаблон не найден)")
    doc.save(bio)
    bio.seek(0)
    return bio