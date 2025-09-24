## Google Sheets automation

The `apps_script/ContentPlanStats.gs` file contains a ready-to-use Google Apps Script that pulls post-performance metrics from the LiveDune API into the content-plan spreadsheet.

Key features:

- Automatically adds a **Статистика → Собрать показатели → <месяц>** menu.
- Supports Telegram, VKontakte and Odnoklassники posts (based on the platform name in the sheet).
- Looks up post links in the `Ссылка` column and writes metrics (views/reach, interactions, ER) back into the columns that already exist in the sheet.
- Caches API calls within a single run to avoid duplicate requests for the same link.

How to use:

1. Open the [content-plan spreadsheet](https://docs.google.com/spreadsheets/d/1qTU42nME1EdGHg90NV8VOd81kSz-0nFZtB1hJUVWC6s/edit?gid=0#gid=0) and go to **Extensions → Apps Script**.
2. Replace the default script with the contents of `apps_script/ContentPlanStats.gs`.
3. Make sure the header row contains the columns referenced in `SHEET_CONFIG` and that the metric columns match the names listed in `PLATFORM_CONFIG` (you can adjust the arrays if your headers differ).
4. Click **Run** once to authorise the script, then use the **Статистика** menu to collect monthly metrics.
# ---
Скрипт для гугл-таблиц, который позволяет автоматически показатели постов в соцсетях по ссылкам. Запускается по кнопке.
