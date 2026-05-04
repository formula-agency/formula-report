# Переменные для автоматического отчета

Секреты в чат не отправляй. Заполни локально файл `.env` на основе `.env.example`, а потом эти же значения перенеси в `GitHub Secrets`.

## Что заполнить в `.env`

### Bitrix

- `BITRIX_WEBHOOK_URL` - полный URL входящего вебхука Bitrix24
- `BITRIX_ENTITY_TYPE` - что считать: `lead` или `deal`
- `BITRIX_DATE_FIELD` - поле даты для отбора, обычно `DATE_CREATE`
- `BITRIX_UTM_SOURCE_FIELD` - код поля `utm_source`
- `BITRIX_UTM_MEDIUM_FIELD` - код поля `utm_medium`
- `BITRIX_UTM_CAMPAIGN_FIELD` - код поля `utm_campaign`
- `BITRIX_REQUEST_TIMEOUT` - таймаут запроса к Bitrix в секундах, например `120`

### Google Sheets

- `GOOGLE_SHEET_ID` - ID таблицы из URL
- `GOOGLE_SHEET_NAME` - имя листа
- `GOOGLE_ALLOWED_RANGE` - диапазон, который можно обновлять, например `A2:J200`
- `GOOGLE_SERVICE_ACCOUNT_FILE` - локальный путь до JSON-ключа, например `Credentials/otchety-493307-a6f2a3a4c466.json`
- `GOOGLE_SERVICE_ACCOUNT_JSON` - JSON-ключ service account для GitHub Secrets

### Логика отчета

- `REPORT_TIMEZONE` - часовой пояс отчета
- `REPORT_PERIOD_MODE` - режим периода, для этого отчета нужен `from_start_date`
- `REPORT_START_DATE` - стартовая дата загрузки, для этого отчета `2026-03-01`
- `REPORT_REQUIRE_ALL_UTM` - `true`, если учитывать только записи с заполненными всеми UTM
- `REPORT_UNKNOWN_SOURCE` - что писать, если источник не распознан

## Что потом перенести в GitHub Secrets

Создай в репозитории одинаковые по названию секреты:

- `BITRIX_WEBHOOK_URL`
- `BITRIX_ENTITY_TYPE`
- `BITRIX_DATE_FIELD`
- `BITRIX_UTM_SOURCE_FIELD`
- `BITRIX_UTM_MEDIUM_FIELD`
- `BITRIX_UTM_CAMPAIGN_FIELD`
- `BITRIX_REQUEST_TIMEOUT`
- `GOOGLE_SHEET_ID`
- `GOOGLE_SHEET_NAME`
- `GOOGLE_ALLOWED_RANGE`
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `REPORT_TIMEZONE`
- `REPORT_PERIOD_MODE`
- `REPORT_START_DATE`
- `REPORT_REQUIRE_ALL_UTM`
- `REPORT_UNKNOWN_SOURCE`

## Важно

- `.env` не коммить в git
- `bitrix.env` тоже не коммить в git
- вебхук и JSON-ключ не пересылай сообщением
- локально проще использовать `GOOGLE_SERVICE_ACCOUNT_FILE`, а в GitHub хранить `GOOGLE_SERVICE_ACCOUNT_JSON`
- service account нужно выдать доступ редактора к Google-таблице
- внешний вид таблицы не изменится, если скрипт будет обновлять только значения в разрешенном диапазоне
