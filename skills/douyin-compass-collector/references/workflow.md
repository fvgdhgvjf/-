# Douyin Compass Automation Workflow

## Files

The online skill bundles a runnable automation package at:

`assets/automation-package`

Files in that package:

- `运行抖音罗盘采集.cmd`: double-click entry; it changes to its own directory and launches PowerShell.
- `douyin_compass_collect.ps1`: main automation script.
- `自动化流程配置.xlsx`: template config workbook beside the script.

When installing for an end user, copy `assets/automation-package` to a normal working folder, then ask the user to fill `自动化流程配置.xlsx`.

## Config Contract

Only read the `店铺` sheet from `自动化流程配置.xlsx`.

Do not require a `配置` sheet. Login is manual, so no account or password is needed.

## Runtime Flow

1. Start Chrome with remote debugging and a local profile under `.chrome-compass-profile`.
2. Navigate to `https://compass.jinritemai.com/login`.
3. Wait 10 seconds, then check `location.href`.
4. If URL does not contain `login`, treat the user as already logged in.
5. If URL contains `login`, wait for the user to complete login manually.
6. Loop over shops from the `店铺` sheet.
7. Switch shop by hovering XPath `//div[contains(@class, "userName")]`, clicking `切换数据视角`, selecting the target shop, waiting 5-10 seconds, and confirming current shop context.
8. Navigate to `https://compass.jinritemai.com/shop/live-overview`.
9. Query yesterday live rooms via API, filtered to self-operated accounts.
10. For each live room, query products with `explain_cnt > 0`.
11. For each product, query detail APIs for explain durations.
12. Write result workbook in the same folder.
13. Close the Chrome instance started by the script.

## API Endpoints

Live room list:

`/compass_api/shop/live/live_overview/live_room_detail_v2`

Product list:

`/compass_api/content_live/shop/live_screen/product_list_after_live`

Use:

- `index_selected=pay_amt,avg_max_pay_amt_min,pay_combo_cnt,product_click_ucnt,product_click_pay_ucnt_ratio,explain_cnt`
- `data_range=0`
- `sort_field=explain_cnt`
- `is_asc=false`

Product explain segments:

`/compass_api/shop/live/live_screen/product_explain_analysis`

Use this to sum all `data_result[].explain_duration.value` entries for the `讲解时长` column.

Platform total explain duration:

`/compass_api/shop/live/live_screen/product_explain_detail`

Use:

- `room_id`
- `product_id`
- `show_feature=true`
- `without_explain_duration=false`

Read `data.explain_duration.value` for the `平台合计讲解时长` column. Treat this value as seconds and format as Chinese duration text.

## Excel Output

Headers must be exactly:

`店铺、直播间id、直播标题、直播间url、商品id、商品名称、讲解次数、讲解时长、平台合计讲解时长`

Write `直播间id` and `商品id` as text.

The workbook writer must loop over the headers dynamically. Do not hardcode only eight data columns.

## Validation Checklist

After edits:

1. Run PowerShell parser validation on the edited script.
2. Search for unwanted dependencies: `配置`, `Account`, `Password`, `Get-LoginFillExpression`, `compass_debug_`.
3. Confirm result headers include `平台合计讲解时长`.
4. Confirm `Write-ResultsWorkbook` writes values by iterating over `$Headers`.
