param(
  [string]$WorkbookPath = "",
  [string]$LoginUrl = "https://compass.jinritemai.com/login",
  [string]$LiveOverviewUrl = "https://compass.jinritemai.com/shop/live-overview",
  [int]$Port = 9222,
  [int]$DelayMs = 1200,
  [int]$RoomApiTimeoutSeconds = 180,
  [switch]$CaptureApi,
  [switch]$AutoSwitchShop,
  [switch]$AutoFilterLiveList,
  [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProfileDir = Join-Path $ScriptDir ".chrome-compass-profile"
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
  $WorkbookPath = Join-Path $ScriptDir "自动化流程配置.xlsx"
}

function Resolve-ChromePath {
  $candidates = @(
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  foreach ($path in $candidates) {
    if (Test-Path -LiteralPath $path) { return $path }
  }
  throw "Chrome was not found. Please install Google Chrome first."
}

function Normalize-Text {
  param([object]$Value)
  if ($null -eq $Value) { return "" }
  return ([string]$Value).Trim()
}

function Release-ComObject {
  param([object]$Value)
  if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
  }
}

function Convert-ColumnLettersToNumber {
  param([string]$Letters)
  $value = 0
  foreach ($ch in $Letters.ToUpper().ToCharArray()) {
    if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
    $value = ($value * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $value
}

function Get-OpenXmlWorksheetGrid {
  param(
    [string]$Path,
    [string]$SheetName
  )

  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

  $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
  try {
    $workbookEntry = $zip.GetEntry("xl/workbook.xml")
    $relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
    if (-not $workbookEntry -or -not $relsEntry) {
      throw "Workbook structure is not recognized."
    }

    $ns = New-Object System.Xml.XmlNamespaceManager((New-Object System.Xml.NameTable))
    $ns.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $relsNs = New-Object System.Xml.XmlNamespaceManager((New-Object System.Xml.NameTable))
    $relsNs.AddNamespace("d", "http://schemas.openxmlformats.org/package/2006/relationships")

    $workbookDoc = New-Object System.Xml.XmlDocument
    $workbookDoc.Load($workbookEntry.Open())
    $sheetNode = $null
    foreach ($candidate in $workbookDoc.SelectNodes("//d:sheets/d:sheet", $ns)) {
      if (($candidate.GetAttribute("name")) -eq $SheetName) {
        $sheetNode = $candidate
        break
      }
    }
    if (-not $sheetNode) {
      throw "Sheet not found: $SheetName"
    }
    $relId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relsDoc = New-Object System.Xml.XmlDocument
    $relsDoc.Load($relsEntry.Open())
    $relNode = $relsDoc.SelectSingleNode("//d:Relationship[@Id='$relId']", $relsNs)
    if (-not $relNode) {
      throw "Relationship for sheet not found: $SheetName"
    }
    $sheetPath = "xl/" + $relNode.GetAttribute("Target")
    if ($sheetPath.StartsWith("xl/../")) {
      $sheetPath = $sheetPath.Replace("xl/../", "xl/")
    }

    $sharedStrings = @()
    $sharedEntry = $zip.GetEntry("xl/sharedStrings.xml")
    if ($sharedEntry) {
      $sharedDoc = New-Object System.Xml.XmlDocument
      $sharedDoc.Load($sharedEntry.Open())
      $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedDoc.NameTable)
      $sharedNs.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      foreach ($si in $sharedDoc.SelectNodes("//d:sst/d:si", $sharedNs)) {
        $parts = @()
        foreach ($t in $si.SelectNodes(".//d:t", $sharedNs)) {
          $parts += $t.InnerText
        }
        $sharedStrings += ($parts -join "")
      }
    }

    $sheetEntry = $zip.GetEntry($sheetPath)
    if (-not $sheetEntry) {
      throw "Worksheet file not found: $sheetPath"
    }
    $sheetDoc = New-Object System.Xml.XmlDocument
    $sheetDoc.Load($sheetEntry.Open())
    $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetDoc.NameTable)
    $sheetNs.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $cells = @{}
    $maxRow = 0
    $maxCol = 0
    foreach ($row in $sheetDoc.SelectNodes("//d:sheetData/d:row", $sheetNs)) {
      foreach ($cell in $row.SelectNodes("d:c", $sheetNs)) {
        $ref = $cell.GetAttribute("r")
        if ($ref -notmatch '^([A-Z]+)(\d+)$') { continue }
        $colLetters = $Matches[1]
        $rowNumber = [int]$Matches[2]
        $colNumber = Convert-ColumnLettersToNumber $colLetters
        $type = $cell.GetAttribute("t")
        $value = ""
        if ($type -eq "s") {
          $indexNode = $cell.SelectSingleNode("d:v", $sheetNs)
          if ($indexNode) {
            $index = [int]$indexNode.InnerText
            if ($index -ge 0 -and $index -lt $sharedStrings.Count) {
              $value = $sharedStrings[$index]
            }
          }
        } elseif ($type -eq "inlineStr") {
          $inlineNode = $cell.SelectSingleNode("d:is", $sheetNs)
          if ($inlineNode) {
            $parts = @()
            foreach ($t in $inlineNode.SelectNodes(".//d:t", $sheetNs)) {
              $parts += $t.InnerText
            }
            $value = $parts -join ""
          }
        } else {
          $valueNode = $cell.SelectSingleNode("d:v", $sheetNs)
          if ($valueNode) { $value = $valueNode.InnerText }
        }
        $cells["$rowNumber,$colNumber"] = Normalize-Text $value
        if ($rowNumber -gt $maxRow) { $maxRow = $rowNumber }
        if ($colNumber -gt $maxCol) { $maxCol = $colNumber }
      }
    }

    $grid = New-Object 'object[,]' $maxRow, $maxCol
    for ($r = 1; $r -le $maxRow; $r++) {
      for ($c = 1; $c -le $maxCol; $c++) {
        $key = "$r,$c"
        $value = ""
        if ($cells.ContainsKey($key)) { $value = $cells[$key] }
        $grid[($r - 1), ($c - 1)] = $value
      }
    }

    return [pscustomobject]@{
      Rows = $maxRow
      Cols = $maxCol
      Grid = $grid
    }
  } finally {
    if ($zip) { $zip.Dispose() }
  }
}

function Read-WorksheetGrid {
  param(
    [string]$Path,
    [string]$SheetName
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Workbook not found: $Path"
  }

  $excel = $null
  $workbook = $null
  $sheet = $null
  $range = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($Path, $null, $true)
    $sheet = $workbook.Worksheets.Item($SheetName)
    $range = $sheet.UsedRange
    $rows = [int]$range.Rows.Count
    $cols = [int]$range.Columns.Count
    $grid = New-Object 'object[,]' $rows, $cols
    for ($r = 1; $r -le $rows; $r++) {
      for ($c = 1; $c -le $cols; $c++) {
        $grid[($r - 1), ($c - 1)] = Normalize-Text $range.Cells.Item($r, $c).Text
      }
    }
    return [pscustomobject]@{
      Rows = $rows
      Cols = $cols
      Grid = $grid
    }
  } catch {
    return Get-OpenXmlWorksheetGrid -Path $Path -SheetName $SheetName
  } finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $range
    Release-ComObject $sheet
    Release-ComObject $workbook
    Release-ComObject $excel
  }
}

function Get-ShopNames {
  param([object[,]]$Grid)

  $rows = $Grid.GetLength(0)
  $cols = $Grid.GetLength(1)
  $candidateColumns = New-Object System.Collections.Generic.List[int]

  for ($c = 0; $c -lt $cols; $c++) {
    $header = Normalize-Text $Grid[0, $c]
    if ($header -match '店铺|店名|名称|shop') {
      $candidateColumns.Add($c)
    }
  }

  $shops = New-Object System.Collections.Generic.List[string]
  if ($candidateColumns.Count -gt 0) {
    foreach ($col in $candidateColumns) {
      for ($r = 1; $r -lt $rows; $r++) {
        $value = Normalize-Text $Grid[$r, $col]
        if ($value -and $value -notmatch '店铺|店名|名称|shop') {
          $shops.Add($value)
        }
      }
    }
  } else {
    for ($r = 0; $r -lt $rows; $r++) {
      for ($c = 0; $c -lt $cols; $c++) {
        $value = Normalize-Text $Grid[$r, $c]
        if ($value -and $value -notmatch '店铺|店名|名称|shop') {
          $shops.Add($value)
          break
        }
      }
    }
  }

  $unique = @{}
  foreach ($shop in $shops) {
    if (-not $unique.ContainsKey($shop)) { $unique[$shop] = $true }
  }
  return @($unique.Keys)
}

function Read-CompassWorkbook {
  param([string]$Path)

  $shopSheet = Read-WorksheetGrid -Path $Path -SheetName "店铺"
  $shops = Get-ShopNames -Grid $shopSheet.Grid

  if (-not $shops -or $shops.Count -eq 0) { throw "Cannot find any shop name in sheet '店铺'." }

  return [pscustomobject]@{
    Shops = $shops
  }
}

function Wait-HttpJson {
  param([string]$Url, [int]$TimeoutSeconds = 30)
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  do {
    try {
      return Invoke-RestMethod -Uri $Url -UseBasicParsing
    } catch {
      Start-Sleep -Milliseconds 500
    }
  } while ((Get-Date) -lt $deadline)
  throw "Cannot connect to Chrome debugging endpoint: $Url"
}

function New-CdpConnection {
  param([string]$WebSocketUrl)
  $ws = [System.Net.WebSockets.ClientWebSocket]::new()
  $ct = [Threading.CancellationToken]::None
  $null = $ws.ConnectAsync([Uri]$WebSocketUrl, $ct).GetAwaiter().GetResult()
  return [pscustomobject]@{ Socket = $ws; NextId = 1 }
}

function Receive-CdpMessage {
  param($Conn)
  $ct = [Threading.CancellationToken]::None
  $buffer = New-Object byte[] 1048576
  $ms = [System.IO.MemoryStream]::new()
  do {
    $segment = [ArraySegment[byte]]::new($buffer)
    $result = $Conn.Socket.ReceiveAsync($segment, $ct).GetAwaiter().GetResult()
    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
      throw "Chrome debugging connection was closed."
    }
    $ms.Write($buffer, 0, $result.Count)
  } while (-not $result.EndOfMessage)
  $text = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
  $ms.Dispose()
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

function Invoke-Cdp {
  param(
    $Conn,
    [string]$Method,
    [hashtable]$Params = @{}
  )
  $id = $Conn.NextId
  $Conn.NextId += 1
  $payload = @{ id = $id; method = $Method; params = $Params } | ConvertTo-Json -Depth 20 -Compress
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
  $segment = [ArraySegment[byte]]::new($bytes)
  $Conn.Socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
  while ($true) {
    $message = Receive-CdpMessage $Conn
    if ($message -and $message.id -eq $id) {
      if ($message.error) { throw ("CDP " + $Method + " failed: " + ($message.error | ConvertTo-Json -Compress)) }
      return $message.result
    }
  }
}

function Invoke-PageJson {
  param(
    $Conn,
    [string]$Expression,
    [int]$TimeoutSeconds = 120
  )
  $result = Invoke-Cdp $Conn "Runtime.evaluate" @{
    expression = $Expression
    awaitPromise = $true
    returnByValue = $true
    timeout = $TimeoutSeconds * 1000
  }
  if ($result.exceptionDetails) {
    throw ("Page script failed: " + ($result.exceptionDetails | ConvertTo-Json -Depth 8 -Compress))
  }
  $value = $result.result.value
  if ([string]::IsNullOrWhiteSpace($value)) { return $null }
  return $value | ConvertFrom-Json
}

function Start-CompassChrome {
  param(
    [string]$ChromePath,
    [string]$StartUrl,
    [string]$ProfileDir,
    [int]$DebugPort
  )
  Write-Host ("Starting Chrome: {0}" -f $StartUrl)
  New-Item -ItemType Directory -Force -Path $ProfileDir | Out-Null
  return Start-Process -FilePath $ChromePath -ArgumentList @(
    "--remote-debugging-port=$DebugPort",
    "--user-data-dir=$ProfileDir",
    "--new-window",
    $StartUrl
  ) -PassThru
}

function Stop-CompassChrome {
  param(
    $Conn,
    $ChromeProcess
  )
  try {
    if ($Conn) {
      Invoke-Cdp $Conn "Browser.close" | Out-Null
      Start-Sleep -Seconds 2
    }
  } catch {
    Write-Host ("Browser.close failed: {0}" -f $_.Exception.Message)
  }
  try {
    if ($ChromeProcess -and -not $ChromeProcess.HasExited) {
      Stop-Process -Id $ChromeProcess.Id -Force -ErrorAction SilentlyContinue
    }
  } catch {
    Write-Host ("Chrome process cleanup failed: {0}" -f $_.Exception.Message)
  }
}

function Get-YesterdayTokens {
  $y = (Get-Date).AddDays(-1)
  return @(
    $y.ToString("yyyy-MM-dd"),
    $y.ToString("yyyy/MM/dd"),
    $y.ToString("MM-dd"),
    $y.ToString("M月d日"),
    "昨天",
    "昨日"
  )
}

function Get-ShopSelectExpression {
  param([string]$ShopName)
  $cfgJson = @{ shopName = $ShopName } | ConvertTo-Json -Depth 5 -Compress
  return @"
(async () => {
  const cfg = $cfgJson;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (s) => (s || '').replace(/\s+/g, '').toLowerCase();
  const visible = (el) => !!(el && el.getClientRects && el.getClientRects().length);
  const rawText = (el) => ((el && (el.innerText || el.textContent)) || '').replace(/\s+/g, ' ').trim();
  const textOf = (el) => norm(rawText(el));
  const byXPath = (xpath) => {
    const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    return result.singleNodeValue;
  };
  const hover = async (el) => {
    if (!el) return false;
    el.scrollIntoView({ block: 'center', inline: 'center' });
    await sleep(150);
    el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, cancelable: true, view: window }));
    await sleep(900);
    return true;
  };
  const fireClick = async (el, waitMs = 900) => {
    if (!el) return false;
    el.scrollIntoView({ block: 'center', inline: 'center' });
    await sleep(120);
    el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
    el.click();
    if (waitMs > 0) await sleep(waitMs);
    return true;
  };
  const allClickables = () => [...document.querySelectorAll('button,a,[role="button"],[role="menuitem"],[role="option"],li,div,span')]
    .filter(visible)
    .filter((el) => rawText(el));
  const clickShopIfVisible = async () => {
    const targetName = norm(cfg.shopName);
    const nodes = allClickables()
      .filter((el) => textOf(el).includes(targetName))
      .sort((a, b) => rawText(a).length - rawText(b).length);
    for (const node of nodes) {
      const clickedText = rawText(node);
      await fireClick(node, 0);
      return { clicked: true, text: clickedText };
    }
    return { clicked: false, text: '' };
  };
  const openByUserName = async () => {
    const userName = byXPath('//div[contains(@class, "userName")]');
    if (!userName || !visible(userName)) return { opened: false, reason: 'userName not found' };
    await hover(userName);
    const switchNodes = allClickables()
      .filter((el) => /切换数据视角/.test(rawText(el)))
      .sort((a, b) => rawText(a).length - rawText(b).length);
    if (!switchNodes.length) return { opened: false, reason: 'switch data view not found', userNameText: rawText(userName) };
    await fireClick(switchNodes[0]);
    await sleep(1200);
    return { opened: true, openerText: rawText(switchNodes[0]), userNameText: rawText(userName) };
  };
  const clickOpeners = async () => {
    const openerPatterns = [
      /切换店铺|店铺切换|更换店铺|选择店铺|当前店铺/,
      /店铺.*切换|店铺.*选择/,
      /切换|更换/
    ];
    const nodes = allClickables()
      .filter((el) => {
        const text = rawText(el);
        if (/直播|商品|数据|订单|首页|概览|客服|消息|帮助/.test(text)) return false;
        return openerPatterns.some((p) => p.test(text));
      })
      .sort((a, b) => rawText(a).length - rawText(b).length);
    for (const node of nodes.slice(0, 8)) {
      await fireClick(node);
      await sleep(1000);
      if ([...document.querySelectorAll('input')].filter(visible).some((inp) => /搜索|店铺|shop/i.test([inp.placeholder, inp.name, inp.id, inp.type].join(' ')))) {
        return { opened: true, openerText: rawText(node) };
      }
      const visibleShop = await clickShopIfVisible();
      if (visibleShop.clicked) return { opened: true, openerText: rawText(node), clickedShopText: visibleShop.text };
    }
    return { opened: false, openerText: '' };
  };

  let direct = await clickShopIfVisible();
  if (direct.clicked) {
    return JSON.stringify({ ok: true, mode: 'direct-visible', matchedText: direct.text, clicked: true, url: location.href, title: document.title });
  }

  let opener = await openByUserName();
  if (!opener.opened) {
    opener = await clickOpeners();
  }
  await sleep(600);

  const setInputValue = (input, value) => {
    if (!input) return false;
    input.focus();
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', bubbles: true }));
    return true;
  };
  const inputs = [...document.querySelectorAll('input')].filter(visible);
  const search = inputs.find((inp) => {
    const meta = [inp.placeholder, inp.name, inp.id, inp.type].join(' ');
    return /搜索|店铺|shop/i.test(meta);
  }) || inputs[0];
  let searched = false;
  if (search) {
    searched = setInputValue(search, cfg.shopName);
    await sleep(1500);
  }

  const selected = await clickShopIfVisible();
  if (selected.clicked) {
    return JSON.stringify({ ok: true, mode: 'search-click', opened: opener.opened, openerText: opener.openerText, searched, matchedText: selected.text, clicked: true, url: location.href, title: document.title });
  }
  return JSON.stringify({ ok: false, opened: opener.opened, openerText: opener.openerText, searched, matchedText: '', url: location.href, title: document.title });
})()
"@
}

function Get-ShopReadyExpression {
  param([string]$ShopName)
  $cfgJson = @{ shopName = $ShopName } | ConvertTo-Json -Depth 5 -Compress
  return @"
JSON.stringify((() => {
  const cfg = $cfgJson;
  const norm = (s) => (s || '').replace(/\s+/g, '').toLowerCase();
  const rawText = (el) => ((el && (el.innerText || el.textContent)) || '').replace(/\s+/g, ' ').trim();
  const byXPath = (xpath) => {
    const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    return result.singleNodeValue;
  };
  const currentShop = rawText(byXPath('//div[contains(@class, "userName")]'));
  const body = document.body ? (document.body.innerText || document.body.textContent || '') : '';
  const target = norm(cfg.shopName);
  return {
    ok: norm(currentShop).includes(target) || norm(body).includes(target),
    currentShop,
    url: location.href,
    title: document.title,
    bodySample: body.replace(/\s+/g, ' ').trim().slice(0, 300)
  };
})())
"@
}

function Get-LiveRoomCollectionExpression {
  param(
    [string]$BeginDate,
    [string]$EndDate,
    [int]$PageSize = 100
  )
  $cfgJson = @{
    beginDate = $BeginDate
    endDate = $EndDate
    pageSize = $PageSize
  } | ConvertTo-Json -Depth 5 -Compress
  return @"
(async () => {
  const cfg = $cfgJson;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const requestDelayMs = () => 2000 + Math.floor(Math.random() * 2001);
  const baseQuery = {
    page_size: String(cfg.pageSize),
    page_no: '1',
    index_selected: 'new_pay_amt,pay_amt,watch_cnt,pay_cnt,net_pay_cnt,ad_costed_amt,stat_cost,use_coupon_pay_amt,coupon_pay_amt_ratio',
    date_type: '2',
    begin_date: cfg.beginDate + ' 00:00:00',
    end_date: cfg.endDate + ' 00:00:00',
    a_type: '0',
    activity_id: ''
  };
  const buildUrl = (aType) => {
    const q = new URLSearchParams(baseQuery);
    q.set('a_type', String(aType));
    return '/compass_api/shop/live/live_overview/live_room_detail_v2?' + q.toString();
  };
  async function fetchList(aType) {
    const res = await fetch(buildUrl(aType), { credentials: 'include' });
    const json = await res.json();
    await sleep(requestDelayMs());
    if (json.st !== 0) throw new Error('live_room_detail_v2 failed: ' + JSON.stringify(json).slice(0, 2000));
    const detail = json.data && json.data.module_data && json.data.module_data.shop_live_list_room_detail && json.data.module_data.shop_live_list_room_detail.compass_general_table_value;
    const rows = ((detail && detail.data) || []).map((item) => {
      const ci = item.cell_info || {};
      const room = (ci.room && ci.room.room) || {};
      const author = room.author || {};
      const authorType = ci.author_type && ci.author_type.value ? ci.author_type.value.value : null;
      return {
        liveRoomId: String(room.live_room_id || ''),
        title: room.live_room_title || '',
        authorNickName: author.nick_name || '',
        authorShortId: author.short_id || '',
        authorId: author.author_id || '',
        authorType: authorType,
        showBigScreen: ci.show_big_screen && ci.show_big_screen.value ? ci.show_big_screen.value.value : null,
        showDetail: ci.show_detail && ci.show_detail.value ? ci.show_detail.value.value : null,
        liveStartTs: room.live_start_ts || (ci.live_start_ts && ci.live_start_ts.value ? ci.live_start_ts.value.value_str : ''),
        liveEndTs: room.live_end_ts || '',
        liveStatus: room.live_status || null,
        watchCnt: ci.watch_cnt && ci.watch_cnt.index_values && ci.watch_cnt.index_values.value ? ci.watch_cnt.index_values.value.value : null,
        payCnt: ci.pay_cnt && ci.pay_cnt.index_values && ci.pay_cnt.index_values.value ? ci.pay_cnt.index_values.value.value : null,
        netPayCnt: ci.net_pay_cnt && ci.net_pay_cnt.index_values && ci.net_pay_cnt.index_values.value ? ci.net_pay_cnt.index_values.value.value : null,
        newPayAmt: ci.new_pay_amt && ci.new_pay_amt.index_values && ci.new_pay_amt.index_values.value ? ci.new_pay_amt.index_values.value.value : null,
        payAmt: ci.pay_amt && ci.pay_amt.index_values && ci.pay_amt.index_values.value ? ci.pay_amt.index_values.value.value : null
      };
    });
    return {
      ok: true,
      aType: aType,
      total: detail && detail.page_result ? detail.page_result.total : rows.length,
      rows: rows
    };
  }
  let result = await fetchList(0);
  if (!result.rows.length) {
    result = await fetchList(-1);
  }
  result.rows = result.rows.filter((row) => row.authorType === 0 || row.authorType === '0');
  result.rows = result.rows.filter((row) => row.showDetail === 1 || row.showDetail === '1' || row.showBigScreen === 1 || row.showBigScreen === '1');
  result.rows = result.rows.map((row) => Object.assign({}, row, {
    href: row.liveRoomId ? ('https://compass.jinritemai.com/shop/live-detail?live_room_id=' + row.liveRoomId + '&prepages%5B0%5D=%2Fshop%2Flive-list') : ''
  }));
  result.count = result.rows.length;
  result.items = result.rows;
  result.url = location.href;
  result.title = document.title;
  return JSON.stringify(result);
})()
"@
}

function Get-LiveRoomProductsExpression {
  param(
    [string]$LiveRoomId,
    [string]$BeginDate,
    [string]$EndDate,
    [int]$PageSize = 100
  )
  $cfgJson = @{
    liveRoomId = $LiveRoomId
    beginDate = $BeginDate
    endDate = $EndDate
    pageSize = $PageSize
  } | ConvertTo-Json -Depth 5 -Compress
  return @"
(async () => {
  const cfg = $cfgJson;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const requestDelayMs = () => 2000 + Math.floor(Math.random() * 2001);
  const attempts = [];
  const liveScreenUrl = 'https://compass.jinritemai.com/screen/live/shop?live_room_id=' + encodeURIComponent(cfg.liveRoomId) + '&live_app_id=1128&source=compass-live-detail&from_page=%2Fshop%2Flive-detail';
  const listEndpoint = '/compass_api/content_live/shop/live_screen/product_list_after_live';
  const detailEndpoint = '/compass_api/shop/live/live_screen/product_explain_analysis';
  const platformDetailEndpoint = '/compass_api/shop/live/live_screen/product_explain_detail';
  const getCookie = (name) => {
    const parts = String(document.cookie || '').split(';').map((x) => x.trim());
    const hit = parts.find((x) => x.startsWith(name + '='));
    return hit ? decodeURIComponent(hit.slice(name.length + 1)) : '';
  };
  const fp = getCookie('s_v_web_id');

  const scalar = (value) => {
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
    if (typeof value !== 'object') return '';
    for (const key of ['value', 'value_str', 'valueString', 'text', 'name', 'title', 'display_value', 'display']) {
      if (value[key] != null && typeof value[key] !== 'object') return value[key];
    }
    return '';
  };

  const toNumber = (value) => {
    const raw = String(scalar(value) || '').replace(/,/g, '');
    const match = raw.match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) : 0;
  };

  const flatten = (node, path = '', out = [], depth = 0) => {
    if (depth > 12) return out;
    if (node == null) return out;
    if (typeof node !== 'object') {
      out.push({ path, value: node });
      return out;
    }
    if (Array.isArray(node)) {
      node.forEach((item, index) => flatten(item, path + '[' + index + ']', out, depth + 1));
      return out;
    }
    for (const [key, value] of Object.entries(node)) {
      const next = path ? path + '.' + key : key;
      if (value == null || typeof value !== 'object') {
        out.push({ path: next, value });
      } else {
        const direct = scalar(value);
        if (direct !== '') out.push({ path: next, value: direct });
        flatten(value, next, out, depth + 1);
      }
    }
    return out;
  };

  const pickText = (flat, patterns, excludes = []) => {
    const row = flat.find((x) => patterns.some((p) => p.test(x.path)) && !excludes.some((p) => p.test(x.path)) && String(x.value || '').trim());
    return row ? String(row.value).trim() : '';
  };

  const pickNumber = (flat, patterns) => {
    const row = flat.find((x) => patterns.some((p) => p.test(x.path)) && toNumber(x.value) > 0);
    return row ? toNumber(row.value) : 0;
  };

  const findArrays = (node, path = '', out = []) => {
    if (node == null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      if (node.some((item) => item && typeof item === 'object' && !Array.isArray(item))) {
        const text = JSON.stringify(node.slice(0, 5)).slice(0, 6000);
        const score =
          (/商品|product|goods|item|sku/i.test(text) ? 2 : 0) +
          (/讲解|explain|explanation|explain_cnt/i.test(text) ? 4 : 0) +
          (/次数|count|cnt|num|times/i.test(text) ? 1 : 0) +
          (/product_id|商品ID|商品id/i.test(text) ? 2 : 0);
        out.push({ path, score, rows: node });
      }
      node.forEach((item, index) => findArrays(item, path + '[' + index + ']', out));
      return out;
    }
    for (const [key, value] of Object.entries(node)) findArrays(value, path ? path + '.' + key : key, out);
    return out;
  };

  const extractPageRows = (json) => {
    const arrays = findArrays(json) || [];
    const chosen = arrays
      .filter((x) => x.score > 0)
      .sort((a, b) => b.score - a.score || b.rows.length - a.rows.length)
      .slice(0, 6);
    const products = [];
    for (const group of chosen) {
      for (const row of group.rows) {
        if (!row || typeof row !== 'object' || Array.isArray(row)) continue;
        const flat = flatten(row);
        const explainCount = pickNumber(flat, [
          /(^|\.|_)explain_cnt$/i,
          /(^|\.|_)explain.*cnt/i,
          /讲解.*次数/i,
          /explain.*(count|cnt|num|times)/i,
          /(count|cnt|num|times).*explain/i,
          /explanation.*(count|cnt|num|times)/i
        ]);
        if (!(explainCount > 0)) continue;
        const productTitle = pickText(flat, [
          /product.*(title|name)/i,
          /product_info.*title/i,
          /product_info.*name/i,
          /goods.*(title|name)/i,
          /item.*(title|name)/i,
          /sku.*(title|name)/i,
          /商品.*(标题|名称|名)/i,
          /commodity.*(title|name)/i,
          /title$/i,
          /name$/i
        ], [/author/i, /room/i, /shop/i, /达人/i, /直播间/i]);
        const productId = pickText(flat, [
          /(^|\.|_)product_id$/i,
          /product.*id/i,
          /goods.*id/i,
          /item.*id/i,
          /commodity.*id/i,
          /sku.*id/i,
          /商品.*id/i
        ], [/room/i, /author/i, /shop/i]);
        const explainTime = pickText(flat, [
          /讲解.*(时长|时间|市场)/i,
          /explain.*(duration|time)/i,
          /explanation.*(duration|time)/i
        ]);
        products.push({
          productTitle,
          productId,
          explainCount,
          explainTime,
          platformTotalExplainTime: '',
          explainDetail: '',
          sourceEndpoint: listEndpoint,
          sourcePath: group.path,
          rawSummary: JSON.stringify(row).slice(0, 1200)
        });
      }
    }
    const seen = new Set();
    return products
      .sort((a, b) => b.explainCount - a.explainCount)
      .filter((item) => {
        const key = [item.productId, item.productTitle, item.explainCount, item.sourceEndpoint].join('|');
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  };

  const extractExplainTime = (json, fallback) => {
    const normalizeTime = (text) => {
      const value = String(text || '').trim();
      if (!value) return '';
      if (/^time$/i.test(value)) return '';
      if (!/[\d:秒分时ms]/i.test(value)) return '';
      return value;
    };
    const formatDuration = (totalSeconds) => {
      const seconds = Math.max(0, Math.floor(totalSeconds));
      const hours = Math.floor(seconds / 3600);
      const minutes = Math.floor((seconds % 3600) / 60);
      const secs = seconds % 60;
      return [
        hours > 0 ? hours + '小时' : '',
        (hours > 0 || minutes > 0) ? minutes + '分钟' : '',
        secs + '秒'
      ].join('');
    };
    const toNumber = (text) => {
      const value = String(text || '').replace(/,/g, '').trim();
      if (!value) return null;
      const match = value.match(/-?\d+(?:\.\d+)?/);
      return match ? Number(match[0]) : null;
    };
    const flat = flatten(json);
    const durationValues = flat
      .filter((x) => /(^|[._])explain_duration\.value$/i.test(String(x.path || '')))
      .map((x) => toNumber(x.value))
      .filter((x) => x != null && x > 0);
    if (durationValues.length > 0) {
      const total = durationValues.reduce((sum, value) => sum + value, 0);
      return formatDuration(total);
    }
    const startNode = flat.find((x) => /(^|[._])explain_start_ts$/i.test(String(x.path || '')));
    const endNode = flat.find((x) => /(^|[._])explain_end_ts$/i.test(String(x.path || '')));
    if (startNode && endNode) {
      const start = toNumber(startNode.value);
      const end = toNumber(endNode.value);
      if (start != null && end != null && end > start) {
        const useMs = Math.max(start, end) > 1e12;
        const diffSeconds = useMs ? (end - start) / 1000 : (end - start);
        if (diffSeconds > 0) return formatDuration(diffSeconds);
      }
    }
    const candidates = flat
      .filter((x) => {
        const path = String(x.path || '');
        const value = x.value;
        const text = String(scalar(value) || value || '').trim();
        if (!text) return false;
        const lower = path.toLowerCase();
        if (/(create|update|start|end|finish|publish|page|request|event|insert|modify|last|collect)_time/.test(lower)) return false;
        if (!/(explain|讲解|product_explain|duration|时长)/i.test(lower)) return false;
        if (!/[\d:秒分时ms]/i.test(text)) return false;
        return (
          /explain|讲解|product_explain|duration|时长/.test(lower) ||
          (/(^|[._])(time|length)$/.test(lower) || /_time$/.test(lower))
        );
      })
      .map((x) => {
        const path = String(x.path || '');
        const text = String(scalar(x.value) || x.value || '').trim();
        const score =
          (/explain/i.test(path) ? 3 : 0) +
          (/product/i.test(path) ? 2 : 0) +
          (/duration|时长/i.test(path) ? 3 : 0) +
          (/\btime\b/i.test(path) ? 1 : 0) +
          (/ms|millisecond/i.test(path) ? 1 : 0);
        return { path, text, score };
      })
      .sort((a, b) => b.score - a.score || b.text.length - a.text.length);
    if (candidates.length > 0) return normalizeTime(candidates[0].text);
    return normalizeTime(fallback);
  };

  const extractPlatformTotalExplainTimeInfo = (json) => {
    const formatDuration = (totalSeconds) => {
      const seconds = Math.max(0, Math.floor(Number(totalSeconds) || 0));
      const hours = Math.floor(seconds / 3600);
      const minutes = Math.floor((seconds % 3600) / 60);
      const secs = seconds % 60;
      return [
        hours > 0 ? hours + '小时' : '',
        (hours > 0 || minutes > 0) ? minutes + '分钟' : '',
        secs + '秒'
      ].join('');
    };
    const readDurationNumber = (node) => {
      if (node == null) return null;
      if (typeof node === 'number') return node > 0 ? node : null;
      if (typeof node === 'string') {
        const value = toNumber(node);
        return value > 0 ? value : null;
      }
      if (typeof node !== 'object') return null;
      if (node.value != null) {
        const value = readDurationNumber(node.value);
        if (value != null) return value;
      }
      if (node.value_str != null) {
        const value = readDurationNumber(node.value_str);
        if (value != null) return value;
      }
      if (node.display_value != null) {
        const value = readDurationNumber(node.display_value);
        if (value != null) return value;
      }
      return null;
    };
    const directNode = json && json.data ? json.data.explain_duration : null;
    const directNumber = readDurationNumber(directNode);
    if (directNumber != null) {
      return {
        value: formatDuration(directNumber),
        seconds: directNumber,
        path: 'data.explain_duration.value',
        raw: JSON.stringify(directNode).slice(0, 500)
      };
    }

    const flat = flatten(json);
    const durationCandidates = flat
      .filter((x) => /(^|[._])data\.explain_duration(\.|$)/i.test(String(x.path || '')))
      .map((x) => ({
        path: String(x.path || ''),
        raw: x.value,
        seconds: toNumber(x.value),
        score:
          (/\.value(\.|$)/i.test(String(x.path || '')) ? 4 : 0) +
          (/duration/i.test(String(x.path || '')) ? 2 : 0) -
          (/unit$/i.test(String(x.path || '')) ? 5 : 0)
      }))
      .filter((x) => x.seconds > 0)
      .sort((a, b) => b.score - a.score || b.seconds - a.seconds);
    if (durationCandidates.length > 0) {
      const best = durationCandidates[0];
      return {
        value: formatDuration(best.seconds),
        seconds: best.seconds,
        path: best.path,
        raw: String(best.raw).slice(0, 500)
      };
    }
    return { value: '', seconds: 0, path: '', raw: '' };
  };

  async function callApi(label, endpoint, params) {
    const q = new URLSearchParams(params);
    const url = endpoint + '?' + q.toString();
    const startedAt = new Date().toISOString();
    try {
      const res = await fetch(url, {
        credentials: 'include',
        referrer: liveScreenUrl,
        headers: { accept: 'application/json, text/plain, */*' }
      });
      const text = await res.text();
      let json = null;
      try { json = JSON.parse(text); } catch (_) {}
      attempts.push({
        label,
        endpoint,
        url,
        startedAt,
        status: res.status,
        ok: res.ok,
        st: json && json.st,
        msg: json && json.msg,
        keys: json && typeof json === 'object' ? Object.keys(json).slice(0, 20) : [],
        textSample: text.slice(0, 800)
      });
      return { ok: res.ok, status: res.status, json, text };
    } catch (err) {
      attempts.push({
        label,
        endpoint,
        url,
        startedAt,
        ok: false,
        error: err && err.message ? err.message : String(err)
      });
      return { ok: false, status: 0, json: null, text: '' };
    } finally {
      await sleep(requestDelayMs());
    }
  }

  let products = [];
  const pageSize = Math.min(Number(cfg.pageSize) || 20, 20);
  for (let pageNo = 1; pageNo <= 20; pageNo++) {
    const listResult = await callApi('product_list_after_live', listEndpoint, {
      index_selected: 'pay_amt,avg_max_pay_amt_min,pay_combo_cnt,product_click_ucnt,product_click_pay_ucnt_ratio,explain_cnt',
      data_range: '0',
      room_id: cfg.liveRoomId,
      page_no: String(pageNo),
      page_size: String(pageSize),
      sort_field: 'explain_cnt',
      is_asc: 'false',
      verifyFp: fp,
      fp: fp
    });
    const rows = listResult.json ? extractPageRows(listResult.json) : [];
    attempts[attempts.length - 1].productCount = rows.length;
    products = products.concat(rows);
    if (!rows.length || rows.length < pageSize) break;
  }

  const seen = new Set();
  products = products
    .sort((a, b) => b.explainCount - a.explainCount)
    .filter((item) => {
      const key = [item.productId, item.productTitle, item.explainCount].join('|');
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

  for (const product of products) {
    if (!product.productId) continue;
    const detailResult = await callApi('product_explain_analysis', detailEndpoint, {
      room_id: cfg.liveRoomId,
      product_id: product.productId,
      role: 'shop',
      page_no: '1',
      page_size: '20',
      verifyFp: fp,
      fp: fp
    });
    product.explainTime = detailResult.json ? extractExplainTime(detailResult.json, product.explainTime) : product.explainTime;
    const platformResult = await callApi('product_explain_detail', platformDetailEndpoint, {
      room_id: cfg.liveRoomId,
      product_id: product.productId,
      show_feature: 'true',
      without_explain_duration: 'false',
      verifyFp: fp,
      fp: fp
    });
    const platformInfo = platformResult.json ? extractPlatformTotalExplainTimeInfo(platformResult.json) : { value: '', seconds: 0, path: '', raw: '' };
    product.platformTotalExplainTime = platformInfo.value || '';
    product.platformTotalExplainSeconds = platformInfo.seconds || 0;
    product.platformTotalExplainPath = platformInfo.path || '';
    product.platformTotalExplainRaw = platformInfo.raw || '';
    product.platformApiOk = platformResult.ok;
    product.platformApiStatus = platformResult.status || 0;
    product.platformApiSt = platformResult.json && platformResult.json.st != null ? platformResult.json.st : '';
    product.platformApiMsg = platformResult.json && platformResult.json.msg != null ? platformResult.json.msg : '';
    product.platformApiSample = platformResult.text ? String(platformResult.text).slice(0, 300) : '';
    if (attempts.length > 0) {
      const lastAttempt = attempts[attempts.length - 1];
      if (lastAttempt && lastAttempt.label === 'product_explain_detail') {
        lastAttempt.platformDuration = product.platformTotalExplainTime;
        lastAttempt.platformDurationSeconds = product.platformTotalExplainSeconds;
        lastAttempt.platformDurationPath = product.platformTotalExplainPath;
        lastAttempt.platformDurationRaw = product.platformTotalExplainRaw;
      }
    }
    product.sourceEndpoint = listEndpoint + ' + ' + detailEndpoint + ' + ' + platformDetailEndpoint;
  }

  return JSON.stringify({
    ok: true,
    liveRoomId: cfg.liveRoomId,
    count: products.length,
    items: products,
    attempts,
    url: location.href,
    title: document.title
  });
})()
"@
}

function Write-JsonFile {
  param(
    [string]$Path,
    [object]$Value
  )
  $jsonText = $Value | ConvertTo-Json -Depth 80
  [System.IO.File]::WriteAllText($Path, $jsonText, [System.Text.UTF8Encoding]::new($false))
}

function Write-ResultsWorkbook {
  param(
    [string]$Path,
    [string[]]$Headers,
    [object[]]$Rows
  )

  $excel = $null
  $workbook = $null
  $sheet = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
      $workbook.Worksheets.Item(2).Delete()
    }
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "结果"

    for ($c = 1; $c -le $Headers.Count; $c++) {
      $cell = $sheet.Cells.Item(1, $c)
      $cell.Value2 = $Headers[$c - 1]
      $cell.Font.Bold = $true
    }

    $textColumns = @(2, 4, 5)
    foreach ($col in $textColumns) {
      $sheet.Columns.Item($col).NumberFormat = "@"
    }

    $rowIndex = 2
    foreach ($row in $Rows) {
      for ($c = 1; $c -le $Headers.Count; $c++) {
        $header = $Headers[$c - 1]
        $prop = $row.PSObject.Properties[$header]
        $value = if ($prop) { [string]$prop.Value } else { "" }
        $cell = $sheet.Cells.Item($rowIndex, $c)
        if ($textColumns -contains $c) {
          $cell.NumberFormat = "@"
        }
        $cell.Value2 = $value
      }
      $rowIndex++
    }

    $sheet.Rows.Item(1).AutoFilter() | Out-Null
    $sheet.Columns.AutoFit() | Out-Null
    $workbook.SaveAs($Path, 51)
  } finally {
    if ($workbook) { $workbook.Close($true) }
    if ($excel) { $excel.Quit() }
    Release-ComObject $sheet
    Release-ComObject $workbook
    Release-ComObject $excel
  }
}

function Wait-ManualStep {
  param([string]$Message)
  Write-Host ""
  Write-Host $Message
  if (-not $NoPause) {
    Read-Host "完成后按 Enter 继续"
  }
}

function Get-InstallApiRecorderExpression {
  return @"
JSON.stringify((() => {
  if (window.__compassApiRecorderInstalled) {
    return { ok: true, alreadyInstalled: true, count: (window.__compassApiLog || []).length };
  }

  const storageKey = '__compassApiLog';
  const loadStoredLogs = () => {
    try {
      const raw = localStorage.getItem(storageKey);
      return raw ? JSON.parse(raw) : [];
    } catch (_) {
      return [];
    }
  };
  const saveStoredLogs = () => {
    try {
      localStorage.setItem(storageKey, JSON.stringify(window.__compassApiLog || []));
    } catch (_) {}
  };

  window.__compassApiLog = window.__compassApiLog || loadStoredLogs();
  window.__compassApiRecorderInstalled = true;
  const now = () => new Date().toISOString();
  const trimText = (value, max) => {
    const text = value == null ? '' : String(value);
    return text.length > max ? text.slice(0, max) + '...<trimmed>' : text;
  };
  const shouldKeep = (url) => {
    const text = String(url || '');
    return /compass|jinritemai|ecom|live|room|product|goods|author|shop|screen|explain|analysis|list|detail|api/i.test(text);
  };
  const parseMaybeJson = (text) => {
    try { return JSON.parse(text); } catch (_) { return null; }
  };
  const summarizeJson = (value) => {
    if (!value || typeof value !== 'object') return {};
    const keys = Array.isArray(value) ? ['arrayLength:' + value.length] : Object.keys(value).slice(0, 40);
    const text = JSON.stringify(value);
    return {
      keys,
      hasListLikeData: /list|rows|items|data|records|room|product|goods|explain|times|count|duration|讲解|商品|直播/i.test(text),
      textSample: trimText(text, 3000)
    };
  };
  const addLog = (entry) => {
    try {
      if (!shouldKeep(entry.url)) return;
      entry.index = window.__compassApiLog.length + 1;
      window.__compassApiLog.push(entry);
      if (window.__compassApiLog.length > 500) window.__compassApiLog.shift();
      saveStoredLogs();
    } catch (_) {}
  };

  const originalFetch = window.fetch;
  window.fetch = async function(input, init) {
    const startedAt = now();
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    const method = ((init && init.method) || (input && input.method) || 'GET').toUpperCase();
    const requestBody = init && init.body ? trimText(init.body, 8000) : '';
    try {
      const response = await originalFetch.apply(this, arguments);
      const clone = response.clone();
      clone.text().then((text) => {
        const json = parseMaybeJson(text);
        addLog({
          transport: 'fetch',
          startedAt,
          endedAt: now(),
          pageUrl: location.href,
          method,
          url: new URL(url, location.href).href,
          status: response.status,
          requestBody,
          responseSummary: json ? summarizeJson(json) : { textSample: trimText(text, 3000) }
        });
      }).catch(() => {});
      return response;
    } catch (error) {
      addLog({
        transport: 'fetch',
        startedAt,
        endedAt: now(),
        pageUrl: location.href,
        method,
        url: String(url),
        status: 0,
        requestBody,
        error: String(error && error.message || error)
      });
      throw error;
    }
  };

  const OriginalXHR = window.XMLHttpRequest;
  window.XMLHttpRequest = function() {
    const xhr = new OriginalXHR();
    const meta = { method: 'GET', url: '', requestBody: '', startedAt: '' };
    const originalOpen = xhr.open;
    const originalSend = xhr.send;
    xhr.open = function(method, url) {
      meta.method = String(method || 'GET').toUpperCase();
      meta.url = new URL(url, location.href).href;
      return originalOpen.apply(xhr, arguments);
    };
    xhr.send = function(body) {
      meta.startedAt = now();
      meta.requestBody = body ? trimText(body, 8000) : '';
      xhr.addEventListener('loadend', () => {
        const text = xhr.responseType && xhr.responseType !== 'text' && xhr.responseType !== '' ? '' : (xhr.responseText || '');
        const json = parseMaybeJson(text);
        addLog({
          transport: 'xhr',
          startedAt: meta.startedAt,
          endedAt: now(),
          pageUrl: location.href,
          method: meta.method,
          url: meta.url,
          status: xhr.status,
          requestBody: meta.requestBody,
          responseSummary: json ? summarizeJson(json) : { textSample: trimText(text, 3000) }
        });
      });
      return originalSend.apply(xhr, arguments);
    };
    return xhr;
  };
  window.XMLHttpRequest.prototype = OriginalXHR.prototype;

  return { ok: true, alreadyInstalled: false, count: window.__compassApiLog.length };
})())
"@
}

function Get-ApiCaptureExpression {
  return @"
JSON.stringify((() => {
  const logs = window.__compassApiLog || [];
  const classify = (entry) => {
    const blob = [
      entry.url || '',
      entry.requestBody || '',
      entry.responseSummary && entry.responseSummary.textSample || '',
      (entry.responseSummary && entry.responseSummary.keys || []).join(',')
    ].join(' ');
    const labels = [];
    if (/直播间|live|room|场次|开播|下播/i.test(blob)) labels.push('live_room_candidate');
    if (/自营|account|author|达人/i.test(blob)) labels.push('self_account_filter_candidate');
    if (/商品|product|goods|item/i.test(blob)) labels.push('product_candidate');
    if (/讲解|explain|duration|times|count|次数|时长/i.test(blob)) labels.push('explain_candidate');
    if (/大屏|screen|dashboard/i.test(blob)) labels.push('live_screen_candidate');
    if (/detail|详情/i.test(blob)) labels.push('detail_candidate');
    return labels;
  };
  const grouped = {};
  for (const entry of logs) {
    const url = new URL(entry.url, location.href);
    const key = entry.method + ' ' + url.origin + url.pathname;
    grouped[key] = grouped[key] || { count: 0, labels: {}, samples: [] };
    grouped[key].count += 1;
    for (const label of classify(entry)) grouped[key].labels[label] = (grouped[key].labels[label] || 0) + 1;
    if (grouped[key].samples.length < 3) grouped[key].samples.push(entry);
  }
  const candidates = Object.entries(grouped)
    .map(([endpoint, data]) => ({ endpoint, count: data.count, labels: data.labels, samples: data.samples }))
    .sort((a, b) => {
      const score = (x) => Object.values(x.labels).reduce((sum, n) => sum + n, 0) + x.count;
      return score(b) - score(a);
    });
  return {
    url: location.href,
    title: document.title,
    capturedAt: new Date().toISOString(),
    count: logs.length,
    candidates,
    logs
  };
})())
"@
}

function Get-ClearApiCaptureExpression {
  return @"
JSON.stringify((() => {
  try { localStorage.removeItem('__compassApiLog'); } catch (_) {}
  window.__compassApiLog = [];
  return { ok: true };
})())
"@
}

function Save-ApiCapture {
  param(
    $Conn,
    [string]$Name
  )
  $capture = Invoke-PageJson $Conn (Get-ApiCaptureExpression) 30
  $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $safe = ($Name -replace '[\\/:*?"<>|]', '_')
  $path = Join-Path $ScriptDir ("compass_api_capture_{0}_{1}.json" -f $safe, $stamp)
  Write-JsonFile -Path $path -Value $capture
  Write-Host ("API capture: {0}" -f $path)
  Write-Host ("Captured API calls: {0}" -f $capture.count)
  return $capture
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "Cannot find workbook: $WorkbookPath. Please put 自动化流程配置.xlsx in the same folder as this script or pass -WorkbookPath."
}
Write-Host ("Using workbook: {0}" -f $WorkbookPath)
$config = Read-CompassWorkbook -Path $WorkbookPath
$chromePath = Resolve-ChromePath
$chromeProcess = Start-CompassChrome -ChromePath $chromePath -StartUrl $LoginUrl -ProfileDir $ProfileDir -DebugPort $Port

Wait-HttpJson "http://127.0.0.1:$Port/json/version" 45 | Out-Null
Write-Host "Chrome debugging port is ready."
Start-Sleep -Seconds 2

$targets = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json/list" -UseBasicParsing
$page = $targets | Where-Object { $_.type -eq "page" -and $_.url -like "https://compass.jinritemai.com/*" } | Select-Object -First 1
if (-not $page) {
  throw "Compass page was not found. Please confirm Chrome opened the login page."
}

$conn = New-CdpConnection $page.webSocketDebuggerUrl
Invoke-Cdp $conn "Runtime.enable" | Out-Null
Invoke-Cdp $conn "Page.enable" | Out-Null
Invoke-Cdp $conn "Page.navigate" @{ url = $LoginUrl } | Out-Null
Write-Host "Waiting 10 seconds for existing login session redirect..."
Start-Sleep -Seconds 10

$entryStatus = Invoke-PageJson $conn @"
JSON.stringify((() => {
  const url = location.href;
  return {
    loginInUrl: /login/i.test(url),
    url,
    title: document.title
  };
})())
"@ 10

$loginOk = $false
if (-not $entryStatus.loginInUrl) {
  $loginOk = $true
  Write-Host ("Already entered Compass backend, skip login. Current URL: {0}" -f $entryStatus.url)
} else {
  Write-Host ("Login page detected. Current URL: {0}" -f $entryStatus.url)
  Write-Host "Please sign in manually in Chrome. The script will continue after Compass backend is detected."

  for ($i = 0; $i -lt 240; $i++) {
    $status = Invoke-PageJson $conn @"
JSON.stringify((() => {
  const url = location.href;
  const text = document.body ? (document.body.innerText || '') : '';
  return {
    ok: url.includes('/app/#/') || url.includes('/shop/') || /直播|店铺|罗盘/.test(text),
    url,
    title: document.title,
    text: text.slice(0, 300)
  };
})())
"@ 10
    if ($status.ok) { $loginOk = $true; break }
    if (($i % 10) -eq 0) { Write-Host "Waiting for login..." }
    Start-Sleep -Seconds 2
  }
}

if (-not $loginOk) {
  throw "Login wait timed out. Please complete the captcha and sign in, then rerun the script."
}

if ($CaptureApi) {
  Write-Host ""
  Write-Host "API capture mode is enabled."
  Invoke-PageJson $conn (Get-ClearApiCaptureExpression) 10 | Out-Null
  Invoke-Cdp $conn "Page.addScriptToEvaluateOnNewDocument" @{ source = (Get-InstallApiRecorderExpression) } | Out-Null
  Invoke-PageJson $conn (Get-InstallApiRecorderExpression) 10 | Out-Null

  Invoke-Cdp $conn "Page.navigate" @{ url = $LiveOverviewUrl } | Out-Null
  Start-Sleep -Seconds 5
  Invoke-PageJson $conn (Get-InstallApiRecorderExpression) 10 | Out-Null

  Wait-ManualStep "请按真实流程操作一遍：切到一个目标店铺 -> 进入直播概览 -> 筛选昨天 + 自营账号。不要进入直播间详情或直播大屏。"
  $capture = Save-ApiCapture -Conn $conn -Name "manual_flow"

  Write-Host ""
  Write-Host "Top API candidates:"
  $capture.candidates | Select-Object -First 12 | ForEach-Object {
    $labels = ($_.labels.PSObject.Properties | ForEach-Object { $_.Name }) -join ","
    Write-Host ("- {0} | count={1} | labels={2}" -f $_.endpoint, $_.count, $labels)
  }

  Write-Host ""
  Write-Host "Capture complete. Send me the compass_api_capture_*.json file if you want me to turn these into pure API calls."
  Stop-CompassChrome -Conn $conn -ChromeProcess $chromeProcess
  if (-not $NoPause) {
    Read-Host "Press Enter to close"
  }
  exit 0
}

Write-Host "Login confirmed. Starting shop switch loop..."

$allResults = New-Object System.Collections.Generic.List[object]
$apiDebugs = New-Object System.Collections.Generic.List[object]

foreach ($shop in $config.Shops) {
  Write-Host ""
  Write-Host ("Switching shop: {0}" -f $shop)
  $selectResult = $null
  $selectInterrupted = $false
  try {
    $selectResult = Invoke-PageJson $conn (Get-ShopSelectExpression -ShopName $shop) 45
  } catch {
    $selectInterrupted = $true
    Write-Host ("Auto shop switch click may have triggered page refresh: {0}" -f $_.Exception.Message)
  }

  if ($selectResult -and $selectResult.ok) {
    Write-Host ("Auto shop switch clicked: {0}. Waiting for shop context..." -f $selectResult.matchedText)
  } elseif ($selectInterrupted) {
    Write-Host "Waiting for shop context after refresh..."
  } else {
    Write-Host ("Auto shop switch did not click target shop for: {0}" -f $shop)
  }

  Start-Sleep -Seconds 5
  $switchConfirmed = $false
  $readyResult = $null
  for ($waitIndex = 1; $waitIndex -le 10; $waitIndex++) {
    try {
      $readyResult = Invoke-PageJson $conn (Get-ShopReadyExpression -ShopName $shop) 15
      if ($readyResult.ok) {
        $switchConfirmed = $true
        Write-Host ("Shop switch confirmed: {0}; waited {1}s" -f $readyResult.currentShop, (5 + $waitIndex))
        break
      }
    } catch {
      Write-Host ("Shop context check retry {0}: {1}" -f $waitIndex, $_.Exception.Message)
    }
    Start-Sleep -Seconds 1
  }

  if (-not $switchConfirmed) {
    Write-Host ("Auto shop switch did not confirm selection for: {0}; fallback to manual." -f $shop)
    Wait-ManualStep ("请在浏览器里手动切换到店铺【{0}】。" -f $shop)
  }

  Invoke-Cdp $conn "Page.navigate" @{ url = $LiveOverviewUrl } | Out-Null
  Start-Sleep -Seconds 5
  Write-Host "Fetching yesterday live rooms via API..."
  $yesterdayDate = (Get-Date).AddDays(-1).ToString("yyyy/MM/dd")
  $roomPayload = Invoke-PageJson $conn (Get-LiveRoomCollectionExpression -BeginDate $yesterdayDate -EndDate $yesterdayDate -PageSize 100) 60
  if (-not $roomPayload.ok) {
    Write-Host "Live overview scan failed, skipping this shop."
    continue
  }
  Write-Host ("Candidate live rooms: {0}" -f $roomPayload.count)
  if ($roomPayload.count -eq 0) {
    Write-Host "No live rooms came back from the API."
  }

  $rooms = @($roomPayload.items)
  $roomIndex = 0
  foreach ($room in $rooms) {
    $roomIndex += 1
    Write-Host ("[{0}/{1}] Room: {2}" -f $roomIndex, $rooms.Count, $room.title)

    if (-not $room.liveRoomId) {
      Write-Host "Room has no live_room_id, skipping."
      continue
    }

    Start-Sleep -Milliseconds $DelayMs
    $productPayload = Invoke-PageJson $conn (Get-LiveRoomProductsExpression -LiveRoomId $room.liveRoomId -BeginDate $yesterdayDate -EndDate $yesterdayDate -PageSize 100) $RoomApiTimeoutSeconds
    if (-not $productPayload.ok) {
      Write-Host "Product API failed for this room."
      $apiDebugs.Add([pscustomobject]@{
        ShopName = $shop
        RoomId = $room.liveRoomId
        RoomTitle = $room.title
        Error = "Invoke failed"
        Attempts = ""
      })
      continue
    }
    Write-Host ("Products with explain count > 0 via API: {0}" -f $productPayload.count)
    $apiDebugs.Add([pscustomobject]@{
      ShopName = $shop
      RoomId = $room.liveRoomId
      RoomTitle = $room.title
      ProductCount = $productPayload.count
      Attempts = ($productPayload.attempts | ConvertTo-Json -Depth 20 -Compress)
    })
    if ($productPayload.count -eq 0) {
      continue
    }

    $products = @($productPayload.items)
    $productIndex = 0
    foreach ($product in $products) {
      $productIndex += 1
      Write-Host ("  [{0}/{1}] Product: {2} | Count: {3} | Platform duration: {4}" -f $productIndex, $products.Count, $product.productTitle, $product.explainCount, $(if ($product.platformTotalExplainTime) { $product.platformTotalExplainTime } else { "EMPTY" }))
      if (-not $product.platformTotalExplainTime) {
        Write-Host ("    Platform detail empty. status={0}, st={1}, msg={2}, sample={3}" -f $product.platformApiStatus, $product.platformApiSt, $product.platformApiMsg, $product.platformApiSample)
      }

      $allResults.Add([pscustomobject]@{
        ShopName = [string]$shop
        RoomId = [string]$room.liveRoomId
        RoomTitle = [string]$room.title
        RoomUrl = [string]$room.href
        ProductId = [string]$product.productId
        ProductTitle = [string]$product.productTitle
        ExplainCount = [string]$product.explainCount
        ExplainTime = [string]$product.explainTime
        PlatformTotalExplainTime = [string]$product.platformTotalExplainTime
        PlatformTotalExplainSeconds = [string]$product.platformTotalExplainSeconds
        PlatformTotalExplainPath = [string]$product.platformTotalExplainPath
        PlatformTotalExplainRaw = [string]$product.platformTotalExplainRaw
        PlatformApiStatus = [string]$product.platformApiStatus
        PlatformApiSt = [string]$product.platformApiSt
        PlatformApiMsg = [string]$product.platformApiMsg
        SourceEndpoint = [string]$product.sourceEndpoint
        SourcePath = [string]$product.sourcePath
        RawSummary = [string]$product.rawSummary
        SourceUrl = [string]$productPayload.url
      })
    }
  }
}

$stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$safeName = ("douyin_compass_live_explain_{0}" -f $stamp) -replace '[\\/:*?"<>|]', '_'
$xlsxPath = Join-Path $ScriptDir ($safeName + ".xlsx")

$exportRows = foreach ($row in $allResults) {
  [pscustomobject]@{
    店铺 = $row.ShopName
    直播间id = $row.RoomId
    直播标题 = $row.RoomTitle
    直播间url = $row.RoomUrl
    商品id = $row.ProductId
    商品名称 = $row.ProductTitle
    讲解次数 = $row.ExplainCount
    讲解时长 = $row.ExplainTime
    平台合计讲解时长 = $row.PlatformTotalExplainTime
  }
}

$headers = @("店铺","直播间id","直播标题","直播间url","商品id","商品名称","讲解次数","讲解时长","平台合计讲解时长")

Write-ResultsWorkbook -Path $xlsxPath -Headers $headers -Rows $exportRows

Write-Host ""
Write-Host "Collection complete."
Write-Host ("Rows: {0}" -f $allResults.Count)
Write-Host ("XLSX: {0}" -f $xlsxPath)

Stop-CompassChrome -Conn $conn -ChromeProcess $chromeProcess

if (-not $NoPause) {
  Read-Host "Press Enter to close"
}
