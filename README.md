# Автообновление отчета из Bitrix в Google Sheets

Скрипт берет лиды или сделки из Bitrix24, считает количество записей по комбинации `utm_source + utm_medium + utm_campaign` и перестраивает тело отчета в Google Sheets без изменения оформления шапки и форматирования.

## Что умеет

- берет данные из `crm.lead.list` или `crm.deal.list`
- учитывает записи любых стадий
- считает только записи с заполненными UTM-полями
- умеет брать UTM из `UTM_*` и автоматически падать обратно на `UF_LEAD_FIRST_UTM_*`, если именно там лежат реальные данные
- определяет `Источник` по `utm_source`
- собирает дополнительные метрики по сделкам:
  - `Одобрена ипотека`
  - `Проведена встреча/показ`
  - `Зафиксирована бронь`
  - `Закрыто сделок`
- строит историю начиная с `2026-03-01`
- делает вложенную структуру: итог месяца -> итог дня -> детальные строки
- создает row groups в Google Sheets, чтобы месяцы и дни можно было сворачивать
- создает отдельный лист `Итог по источникам`
- дополнительно собирает статический дашборд в папку `dashboard`
- обновляет только значения в диапазоне `GOOGLE_ALLOWED_RANGE`
- запускается локально или по расписанию через GitHub Actions
- может публиковаться в GitHub Pages без прямого доступа браузера к Bitrix и Google API

## Локальная настройка

1. Установи Python 3.11+
2. Создай виртуальное окружение:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

3. Заполни `bitrix.env`

Минимально нужны:

```env
BITRIX_WEBHOOK_URL=https://your-company.bitrix24.ru/rest/1/your_webhook/
BITRIX_ENTITY_TYPE=deal
BITRIX_DATE_FIELD=DATE_CREATE
BITRIX_DEAL_CATEGORY_ID=1
BITRIX_UTM_SOURCE_FIELD=UTM_SOURCE
BITRIX_UTM_MEDIUM_FIELD=UTM_MEDIUM
BITRIX_UTM_CAMPAIGN_FIELD=UTM_CAMPAIGN
BITRIX_STAGE_FIELD=STAGE_ID
BITRIX_APPROVED_MORTGAGE_FIELD=UF_CRM_APPROVED_MORTGAGE
BITRIX_MEETING_SHOW_FIELD=UF_CRM_MEETING_SHOW
BITRIX_RESERVATION_FIELD=UF_CRM_RESERVATION
BITRIX_SUCCESS_STAGE_IDS=WON,CLOSED
BITRIX_REQUEST_TIMEOUT=120

GOOGLE_SHEET_ID=1HJkKDM0k7ZtytAURms7CYX5bjOyLgTsfQW0UU34vceg
GOOGLE_SHEET_NAME=Показатели
GOOGLE_ALLOWED_RANGE=A1:Z60601
GOOGLE_SOURCE_SUMMARY_SHEET_NAME=Итог по источникам
GOOGLE_MEETING_LOG_SHEET_ID=1CNT1xTe5uBHo4W4ZLUh3qZLmgWy7wxe7nSsCtDXwwIo
GOOGLE_MEETING_LOG_SHEET_NAME=Meetings
GOOGLE_SERVICE_ACCOUNT_FILE=Credentials/otchety-493307-a6f2a3a4c466.json

REPORT_TIMEZONE=Asia/Yekaterinburg
REPORT_PERIOD_MODE=from_start_date
REPORT_START_DATE=2026-03-01
REPORT_REQUIRE_ALL_UTM=true
REPORT_UNKNOWN_SOURCE=Не определено
```

## Как заполнить `GOOGLE_SERVICE_ACCOUNT_JSON`

Локально лучше не использовать `GOOGLE_SERVICE_ACCOUNT_JSON`. Вместо этого укажи:

```env
GOOGLE_SERVICE_ACCOUNT_FILE=Credentials/имя-файла.json
```

`GOOGLE_SERVICE_ACCOUNT_JSON` нужен в GitHub Secrets. Туда вставляется полное содержимое JSON-файла service account, целиком.

Если `GOOGLE_SERVICE_ACCOUNT_FILE` не указан, скрипт попробует сам найти один `.json` файл в папке `Credentials`.

## Запуск локально

Проверка без записи в таблицу:

```powershell
python scripts/sync_formula_report.py --env-file bitrix.env --dry-run
```

Боевой запуск:

```powershell
python scripts/sync_formula_report.py --env-file bitrix.env
```

После каждого запуска дополнительно обновляются файлы:

- `dashboard/data/report-data.json`
- `dashboard/data/report-data.js`

Именно они используются GitHub Pages-дэшбордом.

## Логика заполнения

- скрипт берет данные из Bitrix начиная с `REPORT_START_DATE`
- учитываются записи любых стадий, но только с заполненными `utm_source`, `utm_medium`, `utm_campaign`
- для всех сделочных метрик используются только сделки из воронки `Льготная ипотека` (`BITRIX_DEAL_CATEGORY_ID=1`)
- строки перестраиваются так:
  - строка `Итого за <месяц>`
  - внутри нее строки `Итого за <день>`
  - под каждым днем детальные строки по UTM-комбинациям
- если `UTM_SOURCE / UTM_MEDIUM / UTM_CAMPAIGN` пусты, скрипт для лидов использует `UF_LEAD_FIRST_UTM_SOURCE / MEDIUM / CAMPAIGN`
- `Источник` вычисляется так:
  - `leadit` -> `Лидген КЦ`
  - `selfwalk` -> `Самоход`
  - `avito` -> `Авито` для любых `utm_medium / utm_campaign`
  - `recommendation` -> `Рекомендация`
- `Проведена встреча/показ` считается не по полю сделки, а по отдельной таблице встреч:
  - берутся только строки со статусом `Прошла успешно`
  - из строки читается `ID сделки` или `Ссылка на сделку`
  - затем встреча относится в отчет по `DATE_CREATE` и UTM этой сделки
- остальные дополнительные метрики берутся по полям сделки и успешным стадиям закрытия
- если `utm_source` не попал в словарь, в `Источник` пишется само значение `utm_source`
- для месяцев и дней создаются группировки строк Google Sheets, как на скриншоте
- на отдельном листе `Итог по источникам` формируется summary без группировок

## GitHub Actions

Workflow лежит в `.github/workflows/update-report.yml`.

По умолчанию он запускается:

- вручную через `workflow_dispatch`
- при пуше в `main`, если изменились `scripts/`, `dashboard/`, `requirements.txt` или сам workflow
- автоматически каждый день в `01:00 UTC`, это `06:00` по `Asia/Yekaterinburg`

Если нужно другое время, измени `cron`.

В одном прогоне workflow:

- пересчитывает отчет и обновляет Google Sheets
- пересобирает `dashboard/data`
- публикует папку `dashboard` в GitHub Pages

## GitHub Pages

Дашборд лежит в папке `dashboard` и использует только уже подготовленные статические данные.

Для публикации нужно один раз включить:

- `Settings -> Pages -> Build and deployment -> Source = GitHub Actions`
- `Settings -> Actions -> General -> Allow all actions and reusable workflows`

После этого workflow сам будет деплоить страницу.

## Какие секреты создать в GitHub

В репозитории открой `Settings -> Secrets and variables -> Actions` и создай:

- `BITRIX_WEBHOOK_URL`
- `BITRIX_ENTITY_TYPE`
- `BITRIX_DATE_FIELD`
- `BITRIX_UTM_SOURCE_FIELD`
- `BITRIX_UTM_MEDIUM_FIELD`
- `BITRIX_UTM_CAMPAIGN_FIELD`
- `BITRIX_STAGE_FIELD`
- `BITRIX_APPROVED_MORTGAGE_FIELD`
- `BITRIX_MEETING_SHOW_FIELD`
- `BITRIX_RESERVATION_FIELD`
- `BITRIX_SUCCESS_STAGE_IDS`
- `BITRIX_REQUEST_TIMEOUT`
- `GOOGLE_SHEET_ID`
- `GOOGLE_SHEET_NAME`
- `GOOGLE_ALLOWED_RANGE`
- `GOOGLE_SOURCE_SUMMARY_SHEET_NAME`
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_MEETING_LOG_SHEET_ID`
- `GOOGLE_MEETING_LOG_SHEET_NAME`
- `REPORT_TIMEZONE`
- `REPORT_PERIOD_MODE`
- `REPORT_START_DATE`
- `REPORT_REQUIRE_ALL_UTM`
- `REPORT_UNKNOWN_SOURCE`

## Важно

- service account должен иметь доступ редактора к Google-таблице
- `GOOGLE_ALLOWED_RANGE` должен включать шапку и всю рабочую область отчета
- если таблица встреч переедет, измени `GOOGLE_MEETING_LOG_SHEET_ID` и `GOOGLE_MEETING_LOG_SHEET_NAME`
- если в `GOOGLE_SHEET_NAME` указано название книги, а не вкладки, скрипт все равно возьмет единственную вкладку
- скрипт не меняет шапку, ширины, цвета и форматирование, он обновляет значения и группировки строк
