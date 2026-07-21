---
name: douyin-compass-collector
description: Maintain, install, package, run, or troubleshoot the Douyin E-commerce Compass live product explain collection workflow. Use when the user asks about 抖音电商罗盘, 电商罗盘, 直播间商品讲解次数, 讲解时长, 平台合计讲解时长, shop switching, or installing/downloading this automation from an online skill.
---

# Douyin Compass Collector

## Quick Start

This skill includes an automation package at:

`assets/automation-package`

To install the runnable workflow for a user, copy that folder to any convenient location, for example:

`C:\Users\<user>\Documents\抖音罗盘采集自动化流程`

Then run:

`运行抖音罗盘采集.cmd`

The package contains:

- `运行抖音罗盘采集.cmd`
- `douyin_compass_collect.ps1`
- `自动化流程配置.xlsx`

The script reads `自动化流程配置.xlsx` from the same folder as `douyin_compass_collect.ps1`. Do not reintroduce a hardcoded `C:\配置\自动化流程配置.xlsx` dependency unless explicitly requested.

## Operating Rules

- Treat login as manual. Do not read account/password from Excel and do not automate captcha bypass.
- Only require the `店铺` sheet in `自动化流程配置.xlsx`.
- Keep final user-facing output to one result `.xlsx` file in the workflow folder.
- Do not generate normal-run debug JSON snapshots.
- Keep `直播间id` and `商品id` as text in Excel to avoid scientific notation.
- Add 2-4 seconds of delay after each Compass API request.
- Close the Chrome instance started by the script when the run finishes.

## Maintenance

When modifying or diagnosing the workflow, read `references/workflow.md` for the implementation contract, API endpoints, Excel columns, and validation checklist.

After any script edit, validate PowerShell syntax:

```powershell
$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile('<path-to>\douyin_compass_collect.ps1',[ref]$tokens,[ref]$errors) | Out-Null; if ($errors.Count) { $errors | ForEach-Object { $_.Message }; exit 1 } else { 'parse ok' }
```
