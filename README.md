# Thesis-work-Word-Template
Репозиторий хранит word шаблон и прочие настройки выпускной квалификационной работы.
Разработана клубом PAIS Курганского государственного университета

----------
### Use builder

В корне репозитория лежит python файл: `pyDocxPropsFlash.py`.
Это терминальный скрипт, задачей которого является вшивание переменных из JSON файла в кастомные свойства Word документа (docx).
Соответственно, чтобы его использовать нужен .docx файл и properties.json файл.
Структура JSON файла должна быть следующей:
```json
{
    "universityName": "Название университета",
    "departmentName": "Название кафедры",
    "shortDhName": "Иванов И. И. Краткое ФИО заведующего кафедры",
    "fullDhName": "Иванов Иван Иванович. Полное ФИО заведующего кафедры",
    "shortScientificAdvisor": "Иванов И. И. Краткое ФИО науного руководителя",
    "fullScientificAdvisor": "Иванов Иван Иванович. Полное ФИО научного руководителя",
    "degreeScientificAdvisor": "должность и научное звание научного руководителя",
    "directionName": "Название направления с кодом направления",
    "specialization": "Название специализации с кодом специализации",
    "educationLevel": "уровень образования",
    "shortStName": "Краткое ФИО студента",
    "fullStName": "Полное ФИО студента",
    "stGroup": "Номер группы",
    "thesisTopic": "Тема / название ВКР",
    "thesisType": "Тип ВКР",
    "orgDev": "Уникальный номер организации (3375842)",
    "thesisCode": "Код ВКР (ДП24)",
    "passbookCode": "Номер зачетной книжки"
}
```

Теперь, чтобы использовать скрипт нужно выполнить следующий код
```shell
python pyDocxPropsFlash.py --file <path to document> --custom_props <path to properties> --new <path for new document>
```
По умолчанию, если в документе нет какого-то свойства из JSON файла, то оно не добавляется, а в терминале пишется предупреждение
```shell
key: stGroup not in properties in current document
```
Если мы знаем что делаем и хотим добавить все новые свойства в документ, то нужно добавить `--force` при исполнении команды
```shell
python pyDocxPropsFlash.py --file <path to document> --custom_props <path to properties> --new <path for new document> --force
```