from file_parser import parse_curriculum, parse_transcript

print("="*60)
print("ТЕСТ ПАРСИНГА ФАЙЛОВ")
print("="*60)

print("\n1. Читаем учебный план...")
try:
    curriculum = parse_curriculum('учебный план.xlsx')
    print(f"\n✅ Учебный план прочитан успешно!")
    print(f"   Найдено дисциплин: {len(curriculum)}")
except Exception as e:
    print(f"\n❌ Ошибка при чтении учебного плана: {e}")

print("\n" + "="*60)
print("\n2. Читаем ведомость...")
try:
    transcript = parse_transcript('Ведомость1.xlsx')
    print(f"\n✅ Ведомость прочитана успешно!")
    print(f"   Найдено дисциплин: {len(transcript)}")
except Exception as e:
    print(f"\n❌ Ошибка при чтении ведомости: {e}")

print("\n" + "="*60)
print("\n3. Проверка сопоставления...")
if 'curriculum' in locals() and 'transcript' in locals():
    from matcher import auto_match
    results = auto_match(transcript, curriculum)
    print(f"\nРезультаты сопоставления:")
    print(f"  - Автоматически сопоставлено: {len(results['matched'])}")
    print(f"  - Требуют ручной обработки: {len(results['manual'])}")