import uuid
import tempfile
import pickle
import os
from typing import Dict, Any

# Временное хранилище для данных сессий
# Ключ: session_id, значение: словарь с данными
_session_data: Dict[str, Dict[str, Any]] = {}

def get_session_data(session_id: str) -> Dict[str, Any]:
    """Получить данные сессии по ID"""
    if session_id not in _session_data:
        _session_data[session_id] = {}
    return _session_data[session_id]

def set_session_data(session_id: str, key: str, value: Any):
    """Установить значение в данных сессии"""
    if session_id not in _session_data:
        _session_data[session_id] = {}
    _session_data[session_id][key] = value

def clear_session_data(session_id: str):
    """Очистить данные сессии"""
    if session_id in _session_data:
        del _session_data[session_id]

def save_to_temp_file(data: Any) -> str:
    """Сохранить данные во временный файл и вернуть путь"""
    file_id = str(uuid.uuid4())
    filepath = os.path.join(tempfile.gettempdir(), f"session_{file_id}.pkl")
    with open(filepath, 'wb') as f:
        pickle.dump(data, f)
    return filepath

def load_from_temp_file(filepath: str) -> Any:
    """Загрузить данные из временного файла"""
    with open(filepath, 'rb') as f:
        return pickle.load(f)

def delete_temp_file(filepath: str):
    """Удалить временный файл"""
    if os.path.exists(filepath):
        os.remove(filepath)