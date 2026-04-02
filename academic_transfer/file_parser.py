import pandas as pd
import re
from difflib import SequenceMatcher

def parse_curriculum(file_path):
    """
    Парсинг файла учебного плана ЯГТУ
    """
    try:
        # Определяем формат файла
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8')
        else:  # xlsx
            df = pd.read_excel(file_path, header=None)
        
        # Ищем строку с заголовками
        header_row = None
        for idx, row in df.iterrows():
            row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
            if 'название дисциплины' in row_str.lower() or 'шифр' in row_str.lower() and 'дисцип' in row_str.lower():
                header_row = idx
                break
        
        if header_row is None:
            # Если не нашли стандартную шапку, ищем "Название дисциплины" в колонке B
            for idx, row in df.iterrows():
                if pd.notna(row[1]) and 'название дисциплины' in str(row[1]).lower():
                    header_row = idx
                    break
        
        if header_row is not None:
            # Устанавливаем заголовки
            df.columns = df.iloc[header_row]
            df = df.drop(index=list(range(header_row + 1)))
        else:
            # Используем первую строку как заголовок
            df.columns = df.iloc[0]
            df = df[1:]
        
        # Нормализуем названия колонок
        df.columns = [str(col).lower().strip() if pd.notna(col) else f'col_{i}' for i, col in enumerate(df.columns)]
        
        # Ищем колонку с названиями дисциплин
        name_col = None
        for col in df.columns:
            if any(keyword in col.lower() for keyword in ['название дисциплины', 'дисциплина', 'name', 'discipline']):
                name_col = col
                break
        
        # Если не нашли, пробуем колонку B (индекс 1)
        if name_col is None and len(df.columns) > 1:
            name_col = df.columns[1]
        
        if name_col is None:
            raise ValueError("Не найдена колонка с названиями дисциплин")
        
        # Формируем список дисциплин
        disciplines = []
        seen_names = {}  # Для обработки дубликатов
        
        for idx, row in df.iterrows():
            discipline_name = row[name_col]
            
            # Пропускаем пустые и служебные строки
            if pd.isna(discipline_name):
                continue
            
            discipline_name = str(discipline_name).strip()
            
            # Пропускаем строки-заголовки
            if any(skip in discipline_name.lower() for skip in ['блок', 'обязательная часть', 'дисциплины по выбору', 'итого', 'всего']):
                continue
            
            # Пропускаем слишком короткие названия
            if len(discipline_name) < 3:
                continue
            
            # Убираем звездочки и другие маркеры
            clean_name = discipline_name.replace('*', '').strip()
            
            # Ищем колонку с часами/кредитами
            hours = None
            for col in df.columns:
                if any(keyword in col for keyword in ['часы', 'зачетные единицы', 'hours', 'credits', 'зе']):
                    if pd.notna(row[col]):
                        try:
                            hours = float(row[col])
                            break
                        except:
                            pass
            
            # Если есть дубликаты, добавляем семестр для уникальности
            base_name = clean_name
            if clean_name in seen_names:
                seen_names[clean_name] += 1
                clean_name = f"{clean_name} (семестр {seen_names[clean_name]})"
            else:
                seen_names[clean_name] = 1
            
            discipline = {
                'id': idx,
                'name': normalize_name(clean_name),
                'original_name': clean_name,
                'hours': hours,
                'semester': None  # Можно добавить парсинг семестра
            }
            disciplines.append(discipline)
        
        return disciplines
    
    except Exception as e:
        raise Exception(f"Ошибка при чтении файла учебного плана: {str(e)}")

def parse_transcript(file_path):
    """
    Парсинг файла ведомости с оценками ЯГТУ
    """
    try:
        # Определяем формат файла
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8')
        else:  # xlsx
            df = pd.read_excel(file_path)
        
        # Нормализуем названия колонок
        df.columns = [str(col).lower().strip() if pd.notna(col) else f'col_{i}' for i, col in enumerate(df.columns)]
        
        # Ищем нужные колонки
        name_col = None
        grade_col = None
        hours_col = None
        status_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in ['наименование предмета', 'предмет', 'дисциплина', 'name', 'discipline']):
                name_col = col
            if any(keyword in col_lower for keyword in ['оценка', 'grade', 'mark', 'балл']):
                grade_col = col
            if any(keyword in col_lower for keyword in ['часы', 'кредиты', 'hours', 'credits', 'зед']):
                hours_col = col
            if any(keyword in col_lower for keyword in ['вид контроля', 'status']):
                status_col = col
        
        if name_col is None:
            raise ValueError("Не найдена колонка с названиями дисциплин")
        
        if grade_col is None and status_col is None:
            raise ValueError("Не найдена колонка с оценками или видом контроля")
        
        # Формируем список дисциплин
        disciplines = []
        
        for idx, row in df.iterrows():
            # Пропускаем пустые строки
            if pd.isna(row[name_col]):
                continue
            
            discipline_name = str(row[name_col]).strip()
            
            # Пропускаем заголовки и пояснения
            if any(skip in discipline_name.lower() for skip in ['примечание', 'итого', 'всего', 'средний']):
                continue
            
            # Убираем звездочки
            clean_name = discipline_name.replace('*', '').strip()
            
            # Получаем оценку
            grade = None
            if grade_col and pd.notna(row[grade_col]):
                grade = str(row[grade_col]).strip()
            elif status_col and pd.notna(row[status_col]):
                # Если есть вид контроля, но нет оценки, значит не сдано
                status = str(row[status_col]).strip()
                if 'экзамен' in status.lower() or 'зачет' in status.lower():
                    grade = 'не сдано'
            
            # Нормализуем оценку
            normalized_grade = normalize_grade(grade) if grade else None
            
            # Пропускаем дисциплины без оценки (если они не сданы)
            if not normalized_grade and grade != 'не сдано':
                continue
            
            # Получаем часы/кредиты
            hours = None
            if hours_col and pd.notna(row[hours_col]):
                try:
                    hours = float(row[hours_col])
                except:
                    pass
            
            discipline = {
                'id': idx,
                'name': normalize_name(clean_name),
                'original_name': clean_name,
                'grade': grade,
                'normalized_grade': normalized_grade,
                'hours': hours,
                'semester': None  # Можно добавить парсинг семестра
            }
            disciplines.append(discipline)
        
        return disciplines
    
    except Exception as e:
        raise Exception(f"Ошибка при чтении файла ведомости: {str(e)}")

def normalize_name(name):
    """
    Приведение названия дисциплины к единому формату
    """
    if not name:
        return ""
    
    # Приводим к нижнему регистру
    name = name.lower().strip()
    
    # Удаляем лишние пробелы
    name = ' '.join(name.split())
    
    # Удаляем специальные символы (кроме пробелов)
    name = re.sub(r'[^\w\s\(\)]', '', name)
    
    # Удаляем указания на семестр в скобках для лучшего сопоставления
    name = re.sub(r'\(семестр \d+\)', '', name).strip()
    
    return name

def normalize_grade(grade):
    """
    Нормализация оценки для сравнения
    """
    if not grade:
        return None
    
    grade_lower = grade.lower().strip()
    
    # Числовые оценки
    if grade_lower in ['5', 'отлично']:
        return 5
    elif grade_lower in ['4', 'хорошо']:
        return 4
    elif grade_lower in ['3', 'удовлетворительно']:
        return 3
    elif grade_lower in ['2', 'неудовлетворительно']:
        return 2
    elif grade_lower in ['зачет', 'зачтено']:
        return 'зачет'
    elif grade_lower in ['не зачет', 'не зачтено']:
        return 'не зачет'
    
    return grade_lower

def find_best_match(transcript_name, curriculum_list, threshold=0.6):
    """
    Улучшенный поиск совпадений с учетом похожести названий
    """
    best_match = None
    best_ratio = 0
    
    for curr in curriculum_list:
        # Сравниваем нормализованные названия
        ratio = SequenceMatcher(None, transcript_name, curr['name']).ratio()
        
        # Бонус за точное совпадение
        if transcript_name == curr['name']:
            ratio = 1.0
        
        # Бонус за частичное совпадение ключевых слов
        transcript_words = set(transcript_name.split())
        curr_words = set(curr['name'].split())
        word_overlap = len(transcript_words & curr_words) / max(len(transcript_words), 1)
        ratio = max(ratio, word_overlap * 0.8)
        
        if ratio > best_ratio and ratio >= threshold:
            best_ratio = ratio
            best_match = curr
    
    return best_match, best_ratio