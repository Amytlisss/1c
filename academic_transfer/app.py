from flask import Flask, render_template, request, redirect, url_for, session, jsonify, Response, send_file
import os
import shutil
import csv
from io import StringIO
from werkzeug.utils import secure_filename
from file_parser import parse_curriculum, parse_transcript
from matcher import auto_match, apply_manual_match, mark_as_study, get_final_results
from plan_document import build_individual_plan_docx, DEFAULT_PLAN_META, PLAN_META_KEYS

PLAN_META_KEYS = list(DEFAULT_PLAN_META.keys())

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'csv'}

# Создаем папку для загрузок, если её нет
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def clear_upload_folder():
    """Очистка папки с временными файлами"""
    folder = app.config['UPLOAD_FOLDER']
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # Обработка загрузки учебного плана
        if 'curriculum' in request.files:
            file = request.files['curriculum']
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                try:
                    curriculum = parse_curriculum(filepath)
                    
                    # Логируем результат парсинга для отладки
                    print(f"Учебный план: загружено {len(curriculum)} дисциплин")
                    if len(curriculum) > 0:
                        print(f"Пример: {curriculum[0]['original_name']}")
                    
                    session['curriculum'] = curriculum
                    session['curriculum_filename'] = filename
                    message = f"Учебный план загружен, количество дисциплин: {len(curriculum)}"
                    
                    # Удаляем файл после обработки
                    os.remove(filepath)
                    
                    # Если ведомость уже загружена, запускаем сопоставление
                    if 'transcript' in session:
                        match_results = auto_match(session['transcript'], curriculum)
                        session['match_results'] = match_results
                        message += f" Автоматическое сопоставление выполнено. Совпадений: {len(match_results['matched'])}, требуется ручной обработки: {len(match_results['manual'])}"
                    
                    return jsonify({'success': True, 'message': message, 'count': len(curriculum)})
                except Exception as e:
                    print(f"Ошибка при обработке учебного плана: {str(e)}")
                    return jsonify({'success': False, 'message': str(e)})
            else:
                return jsonify({'success': False, 'message': 'Неверный формат файла. Поддерживаются XLSX и CSV'})
        
        # Обработка загрузки ведомости
        elif 'transcript' in request.files:
            file = request.files['transcript']
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                try:
                    transcript = parse_transcript(filepath)
                    
                    # Логируем результат парсинга для отладки
                    print(f"Ведомость: загружено {len(transcript)} дисциплин")
                    if len(transcript) > 0:
                        print(f"Пример: {transcript[0]['original_name']} - Оценка: {transcript[0]['grade']}")
                    
                    session['transcript'] = transcript
                    session['transcript_filename'] = filename
                    message = f"Ведомость загружена, количество дисциплин: {len(transcript)}"
                    
                    # Запускаем автоматическое сопоставление, если учебный план уже загружен
                    if 'curriculum' in session:
                        match_results = auto_match(transcript, session['curriculum'])
                        session['match_results'] = match_results
                        message += f" Автоматическое сопоставление выполнено. Совпадений: {len(match_results['matched'])}, требуется ручной обработки: {len(match_results['manual'])}"
                        
                        # Логируем результаты сопоставления
                        print(f"Сопоставление: найдено {len(match_results['matched'])} совпадений, {len(match_results['manual'])} требует обработки")
                    
                    # Удаляем файл после обработки
                    os.remove(filepath)
                    
                    return jsonify({'success': True, 'message': message, 'count': len(transcript)})
                except Exception as e:
                    print(f"Ошибка при обработке ведомости: {str(e)}")
                    return jsonify({'success': False, 'message': str(e)})
            else:
                return jsonify({'success': False, 'message': 'Неверный формат файла. Поддерживаются XLSX и CSV'})
    
    # GET запрос - отображаем страницу загрузки
    curriculum_loaded = 'curriculum' in session
    transcript_loaded = 'transcript' in session
    
    return render_template('upload.html', 
                         curriculum_loaded=curriculum_loaded,
                         transcript_loaded=transcript_loaded,
                         curriculum_count=len(session.get('curriculum', [])),
                         transcript_count=len(session.get('transcript', [])))

@app.route('/manual_matching', methods=['GET', 'POST'])
def manual_matching():
    if 'match_results' not in session:
        flash('Сначала загрузите файлы', 'warning')
        return redirect(url_for('upload'))
    
    match_results = session.get('match_results', {})
    curriculum = session.get('curriculum', [])
    
    if request.method == 'POST':
        # Обработка ручного сопоставления
        data = request.json
        
        if data.get('action') == 'match':
            success = apply_manual_match(
                match_results, 
                data['transcript_id'], 
                data['curriculum_id'],
                curriculum
            )
            if success:
                session['match_results'] = match_results
                session.modified = True
                return jsonify({'success': True, 'message': 'Сопоставление сохранено'})
            return jsonify({'success': False, 'message': 'Ошибка при сохранении'})
            
        elif data.get('action') == 'study':
            success = mark_as_study(data['transcript_id'], match_results)
            if success:
                session['match_results'] = match_results
                session.modified = True
                return jsonify({'success': True, 'message': 'Дисциплина помечена как "требует изучения"'})
            return jsonify({'success': False, 'message': 'Ошибка при сохранении'})
            
        elif data.get('action') == 'skip':
            # Пропустить дисциплину (будет в итоговом списке как несопоставленная)
            for item in match_results.get('manual', []):
                if item['transcript_discipline']['id'] == data['transcript_id']:
                    item['status'] = 'skipped'
                    session['match_results'] = match_results
                    session.modified = True
                    return jsonify({'success': True, 'message': 'Дисциплина пропущена'})
            return jsonify({'success': False, 'message': 'Ошибка'})
    
    # Для GET запроса отображаем страницу
    manual_disciplines = [item for item in match_results.get('manual', []) 
                         if item['status'] == 'manual']
    
    # Подготавливаем список дисциплин для отображения с возможными совпадениями
    disciplines_for_display = []
    for item in manual_disciplines:
        display_item = {
            'id': item['transcript_discipline']['id'],
            'name': item['transcript_discipline']['original_name'],
            'grade': item['transcript_discipline']['grade'],
            'hours': item['transcript_discipline']['hours'],
            'possible_matches': []
        }
        
        # Добавляем возможные совпадения
        for match in item.get('possible_matches', [])[:10]:
            display_item['possible_matches'].append({
                'id': match['discipline']['id'],
                'name': match['discipline']['original_name'],
                'hours': match['discipline']['hours'],
                'similarity': match['similarity']
            })
        
        # Также добавляем все дисциплины из учебного плана для выбора
        display_item['all_curriculum'] = [
            {'id': curr['id'], 'name': curr['original_name'], 'hours': curr['hours']}
            for curr in curriculum
        ]
        
        disciplines_for_display.append(display_item)
    
    return render_template('manual_matching.html', 
                         disciplines=disciplines_for_display,
                         total_count=len(manual_disciplines))

@app.route('/results')
def results():
    if 'match_results' not in session:
        return redirect(url_for('upload'))
    
    match_results = session.get('match_results', {})
    curriculum = session.get('curriculum', [])
    transcript = session.get('transcript', [])
    
    final_results = get_final_results(match_results, curriculum)
    
    # Добавляем статистику
    stats = {
        'total_transcript': len(transcript),
        'matched': len(match_results.get('matched', [])),
        'manual_total': len(match_results.get('manual', [])),
        'manual_matched': len([m for m in match_results.get('manual', []) if m.get('status') == 'manual_matched']),
        'needs_study': len([m for m in match_results.get('manual', []) if m.get('status') == 'needs_study']),
        'recreditable': len(final_results.get('recreditable', [])),
        'reattestation': len(final_results.get('reattestation', [])),
        'need_study': len(final_results.get('need_study', []))
    }
    
    plan_meta = {**DEFAULT_PLAN_META, **session.get('plan_meta', {})}
    return render_template('results.html', results=final_results, stats=stats, plan_meta=plan_meta)

@app.route('/export_results')
def export_results():
    if 'match_results' not in session:
        return redirect(url_for('upload'))
    
    match_results = session.get('match_results', {})
    curriculum = session.get('curriculum', [])
    
    final_results = get_final_results(match_results, curriculum)
    
    # Создаем CSV файл с разделителем точка с запятой для Excel
    output = StringIO()
    writer = csv.writer(output, delimiter=';')
    
    # Заголовки
    writer.writerow(['Категория', 'Название дисциплины', 'Оценка', 'Часы/Кредиты', 'Сопоставлено с', 'Примечание'])
    
    # Перезачитываются
    for item in final_results['recreditable']:
        writer.writerow([
            'Перезачитываются', 
            item['name'], 
            item.get('grade', ''), 
            item.get('hours', ''),
            item.get('matched_to', ''),
            'Автоматическое сопоставление'
        ])
    
    # На переаттестацию
    for item in final_results['reattestation']:
        writer.writerow([
            'На переаттестацию', 
            item['name'], 
            item.get('grade', ''), 
            item.get('hours', ''),
            item.get('matched_to', ''),
            'Ручное сопоставление'
        ])
    
    # Требуют изучения
    for item in final_results['need_study']:
        writer.writerow([
            'Требуют изучения', 
            item['name'], 
            '', 
            item.get('hours', ''),
            '',
            'Нет в ведомости'
        ])
    
    # Несопоставленные дисциплины из ведомости (если есть)
    for item in match_results.get('manual', []):
        if item.get('status') == 'skipped' or (item.get('status') == 'manual' and not item.get('selected_match')):
            writer.writerow([
                'Требует решения', 
                item['transcript_discipline']['original_name'],
                item['transcript_discipline'].get('grade', ''),
                item['transcript_discipline'].get('hours', ''),
                '',
                'Не сопоставлено'
            ])
    
    output.seek(0)
    
    # Отправляем файл
    return Response(
        output,
        mimetype="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": "attachment;filename=results.csv",
            "Content-Type": "text/csv; charset=utf-8"
        }
    )

@app.route('/export_plan_docx', methods=['POST'])
def export_plan_docx():
    if 'match_results' not in session:
        return redirect(url_for('upload'))
    
    match_results = session.get('match_results', {})
    curriculum = session.get('curriculum', [])
    transcript = session.get('transcript', [])
    
    final_results = get_final_results(match_results, curriculum)
    
    meta = {**DEFAULT_PLAN_META}
    for k in PLAN_META_KEYS:
        v = request.form.get(k, '').strip()
        if v:
            meta[k] = v
    
    session['plan_meta'] = {k: meta.get(k, '') for k in PLAN_META_KEYS}
    session.modified = True
    
    buf = build_individual_plan_docx(
        final_results=final_results,
        transcript_disciplines=transcript,
        match_results=match_results,
        curriculum=curriculum,
        meta=meta
    )
    
    return send_file(
        buf,
        as_attachment=True,
        download_name='individual_plan.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )

@app.route('/reset')
def reset():
    # Очищаем сессию
    session.clear()
    # Очищаем папку с загрузками
    clear_upload_folder()
    return redirect(url_for('index'))

@app.route('/get_status')
def get_status():
    """Получить статус текущей сессии"""
    return jsonify({
        'curriculum_loaded': 'curriculum' in session,
        'transcript_loaded': 'transcript' in session,
        'matched_loaded': 'match_results' in session,
        'curriculum_count': len(session.get('curriculum', [])),
        'transcript_count': len(session.get('transcript', [])),
        'matched_count': len(session.get('match_results', {}).get('matched', [])),
        'manual_count': len(session.get('match_results', {}).get('manual', []))
    })

def flash(message, category='info'):
    """Простая реализация flash сообщений через сессию"""
    if '_flashes' not in session:
        session['_flashes'] = []
    session['_flashes'].append({'message': message, 'category': category})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)