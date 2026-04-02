from difflib import SequenceMatcher
from file_parser import find_best_match

def auto_match(transcript_disciplines, curriculum_disciplines):
    """
    Улучшенное автоматическое сопоставление дисциплин
    """
    results = {
        'matched': [],      # Перезачитываются
        'manual': []        # Требуют ручной обработки
    }
    
    used_curriculum_ids = set()
    
    for trans_disc in transcript_disciplines:
        # Пропускаем несданные дисциплины
        if trans_disc.get('normalized_grade') in ['не сдано', 'не зачет', 2]:
            continue
        
        # Ищем лучшее совпадение
        best_match, similarity = find_best_match(trans_disc['name'], curriculum_disciplines)
        
        # Проверяем, не использована ли уже эта дисциплина
        if best_match and best_match['id'] not in used_curriculum_ids:
            results['matched'].append({
                'transcript_discipline': trans_disc,
                'curriculum_discipline': best_match,
                'similarity': similarity,
                'status': 'matched'
            })
            used_curriculum_ids.add(best_match['id'])
        else:
            # Ищем похожие для ручного выбора
            similar_matches = find_similar_matches(trans_disc, curriculum_disciplines, used_curriculum_ids)
            
            results['manual'].append({
                'transcript_discipline': trans_disc,
                'possible_matches': similar_matches,
                'selected_match': None,
                'status': 'manual'
            })
    
    return results

def find_similar_matches(transcript_disc, curriculum_disciplines, used_ids=None, threshold=0.3):
    """
    Поиск похожих дисциплин для предложения пользователю
    """
    if used_ids is None:
        used_ids = set()
    
    similar = []
    
    for curr_disc in curriculum_disciplines:
        if curr_disc['id'] in used_ids:
            continue
            
        similarity = SequenceMatcher(None, transcript_disc['name'], curr_disc['name']).ratio()
        
        if similarity >= threshold:
            similar.append({
                'discipline': curr_disc,
                'similarity': round(similarity, 2)
            })
    
    # Сортируем по убыванию схожести
    similar.sort(key=lambda x: x['similarity'], reverse=True)
    
    return similar[:10]  # Возвращаем топ-10 похожих

def apply_manual_match(match_results, transcript_id, curriculum_discipline_id, curriculum_disciplines):
    """
    Применение ручного сопоставления
    """
    for item in match_results.get('manual', []):
        if item['transcript_discipline']['id'] == transcript_id:
            # Находим выбранную дисциплину из учебного плана
            selected = None
            for curr in curriculum_disciplines:
                if curr['id'] == curriculum_discipline_id:
                    selected = curr
                    break
            
            if selected:
                item['selected_match'] = selected
                item['status'] = 'manual_matched'
                return True
            else:
                return False
    
    return False

def mark_as_study(transcript_id, match_results):
    """
    Пометка дисциплины как "требует изучения"
    """
    for item in match_results.get('manual', []):
        if item['transcript_discipline']['id'] == transcript_id:
            item['status'] = 'needs_study'
            item['selected_match'] = None
            return True
    
    return False

def get_final_results(match_results, curriculum_disciplines):
    """
    Формирование итогового списка по категориям
    """
    final_results = {
        'recreditable': [],      # Перезачитываются
        'reattestation': [],     # На переаттестацию
        'need_study': []         # Требуют изучения
    }
    
    # Добавляем автоматически сопоставленные
    for item in match_results.get('matched', []):
        final_results['recreditable'].append({
            'name': item['transcript_discipline']['original_name'],
            'grade': item['transcript_discipline'].get('grade', ''),
            'hours': item['transcript_discipline'].get('hours', ''),
            'matched_to': item['curriculum_discipline']['original_name']
        })
    
    # Обрабатываем ручные сопоставления
    for item in match_results.get('manual', []):
        if item.get('status') == 'manual_matched' and item.get('selected_match'):
            final_results['reattestation'].append({
                'name': item['transcript_discipline']['original_name'],
                'grade': item['transcript_discipline'].get('grade', ''),
                'hours': item['transcript_discipline'].get('hours', ''),
                'matched_to': item['selected_match']['original_name']
            })
        elif item.get('status') == 'needs_study':
            # Добавляем дисциплину из учебного плана, которую нужно изучить
            final_results['need_study'].append({
                'name': item['transcript_discipline']['original_name'],
                'hours': item['transcript_discipline'].get('hours', '')
            })
    
    # Добавляем дисциплины из учебного плана, которые не были сопоставлены
    matched_curriculum_ids = set()
    for item in match_results.get('matched', []):
        matched_curriculum_ids.add(item['curriculum_discipline']['id'])
    for item in match_results.get('manual', []):
        if item.get('selected_match'):
            matched_curriculum_ids.add(item['selected_match']['id'])
    
    for curr in curriculum_disciplines:
        if curr['id'] not in matched_curriculum_ids:
            # Проверяем, не добавлена ли уже эта дисциплина
            already_added = False
            for need in final_results['need_study']:
                if need['name'] == curr['original_name']:
                    already_added = True
                    break
            if not already_added:
                final_results['need_study'].append({
                    'name': curr['original_name'],
                    'hours': curr.get('hours', '')
                })
    
    return final_results

def get_matching_stats(match_results):
    """
    Получение статистики по сопоставлениям
    """
    stats = {
        'total_matched': len(match_results.get('matched', [])),
        'total_manual': len(match_results.get('manual', [])),
        'manual_matched': len([m for m in match_results.get('manual', []) if m.get('status') == 'manual_matched']),
        'needs_study': len([m for m in match_results.get('manual', []) if m.get('status') == 'needs_study']),
        'pending': len([m for m in match_results.get('manual', []) if m.get('status') == 'manual'])
    }
    
    return stats