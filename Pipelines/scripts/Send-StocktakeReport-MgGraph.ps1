param (
    [Parameter(Mandatory = $true)]
    [string] $TenantID,

    [Parameter(Mandatory = $true)]
    [string] $ClientID,

    [Parameter(Mandatory = $true)]
    [string] $CertificateFilePath,

    [string] $CertificatePassword = "",

    [Parameter(Mandatory = $true)]
    [string] $Sender,

    [Parameter(Mandatory = $true)]
    [string] $Recipient,

    [string] $DatabasePath = "stocktake-data.json",
    [string] $LogoPath = "Spark_Logo_RGB.png",
    [string] $ReportFolder = "reports",
    [switch] $BuildOnly
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = (Resolve-Path (Join-Path $scriptRoot "..\..")).ProviderPath
if (![System.IO.Path]::IsPathRooted($DatabasePath)) {
    $DatabasePath = Join-Path $repoRoot $DatabasePath
}
$DatabasePath = [System.IO.Path]::GetFullPath($DatabasePath)

if (![System.IO.Path]::IsPathRooted($LogoPath)) {
    $LogoPath = Join-Path $repoRoot $LogoPath
}
$LogoPath = [System.IO.Path]::GetFullPath($LogoPath)

if (![System.IO.Path]::IsPathRooted($ReportFolder)) {
    $ReportFolder = Join-Path $repoRoot $ReportFolder
}
$ReportFolder = [System.IO.Path]::GetFullPath($ReportFolder)

function Get-StockStatus {
    param ($Item)

    if ($Item.retired) { return "Retired" }
    if ([int]$Item.quantity -le 0) { return "Out of stock" }
    if ([int]$Item.quantity -le [int]$Item.minimum) { return "Low stock" }
    return "In stock"
}

function New-EmailRecipient {
    param ([string] $Address)

    @{
        EmailAddress = @{
            Address = $Address
        }
    }
}

function Get-ReportSummary {
    param ($Items)

    $activeItems = @($Items | Where-Object { -not $_.retired })
    $lowStockCount = @($activeItems | Where-Object { [int]$_.quantity -gt 0 -and [int]$_.quantity -le [int]$_.minimum }).Count
    $outOfStockCount = @($activeItems | Where-Object { [int]$_.quantity -le 0 }).Count
    $retiredCount = @($Items | Where-Object { $_.retired }).Count

    [PSCustomObject]@{
        TotalItems = $activeItems.Count
        LowStock = $lowStockCount
        OutOfStock = $outOfStockCount
        Retired = $retiredCount
        Categories = @($activeItems | Select-Object -ExpandProperty category -Unique | Where-Object { $_ }).Count
    }
}

function Get-DisplayDate {
    param ([string] $Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "Not updated"
    }

    try {
        return [DateTime]::Parse($Value).ToString("dd MMM yyyy, HH:mm")
    }
    catch {
        return $Value
    }
}

function ConvertTo-XmlText {
    param ($Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function New-ExcelTextCell {
    param (
        [string] $Reference,
        $Value,
        [int] $Style = 0
    )

    $safeValue = ConvertTo-XmlText $Value
    return "<c r=`"$Reference`" t=`"inlineStr`" s=`"$Style`"><is><t>$safeValue</t></is></c>"
}

function New-ExcelNumberCell {
    param (
        [string] $Reference,
        [int] $Value,
        [int] $Style = 0
    )

    return "<c r=`"$Reference`" s=`"$Style`"><v>$Value</v></c>"
}

function Get-ExcelColumnName {
    param ([int] $ColumnNumber)

        $name = ""
        while ($ColumnNumber -gt 0) {
            $mod = ($ColumnNumber - 1) % 26
            $name = [string][char](65 + $mod) + $name
            $ColumnNumber = [math]::Floor(($ColumnNumber - $mod) / 26)
        }

    return $name
}

function New-ExcelRow {
    param (
        [int] $RowNumber,
        [object[]] $Cells,
        [double] $Height = 0
    )

    $cellXml = $Cells -join ""
    $heightXml = if ($Height -gt 0) { " ht=`"$Height`" customHeight=`"1`"" } else { "" }
    return "<row r=`"$RowNumber`"$heightXml>$cellXml</row>"
}

function New-EmailBodyHtml {
    param (
        [DateTime] $ReportDate,
        [object] $Summary
    )

    $generatedDate = ConvertTo-XmlText ($ReportDate.ToString("dddd, dd MMMM yyyy"))
    $generatedTime = ConvertTo-XmlText ($ReportDate.ToString("HH:mm"))
    $needsAttention = [int]$Summary.LowStock + [int]$Summary.OutOfStock
    $healthText = if ($needsAttention -eq 0) { "Inventory is looking healthy" } else { "$needsAttention item group(s) need attention" }
    $healthColor = if ($needsAttention -eq 0) { "#137a3d" } else { "#b45309" }
    $healthBg = if ($needsAttention -eq 0) { "#e8f7ee" } else { "#fff4de" }
    $healthBorder = if ($needsAttention -eq 0) { "#bfe8ce" } else { "#f5d69a" }
    $healthText = ConvertTo-XmlText $healthText

    return @"
<!doctype html>
<html>
  <body style="margin:0; padding:0; background:#eef2f7; font-family:Segoe UI, Arial, sans-serif; color:#14233d;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f7; margin:0; padding:28px 0;">
      <tr>
        <td align="center" style="padding:0 14px;">
          <table role="presentation" width="680" cellpadding="0" cellspacing="0" style="width:100%; max-width:680px; background:#ffffff; border:1px solid #dbe3ee; border-radius:8px; overflow:hidden;">
            <tr>
              <td style="background:#103b77; padding:28px 32px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                  <tr>
                    <td style="vertical-align:top;">
                      <div style="font-size:12px; line-height:16px; color:#c9dcf7; font-weight:700; letter-spacing:0.08em; text-transform:uppercase;">Spark IT Monthly Report</div>
                      <div style="font-size:30px; line-height:36px; color:#ffffff; font-weight:700; margin-top:6px;">Stocktake Summary</div>
                      <div style="font-size:14px; line-height:20px; color:#dbeafe; margin-top:8px;">Generated $generatedDate at $generatedTime</div>
                    </td>
                    <td align="right" style="vertical-align:top; padding-left:16px;">
                      <img src="cid:spark-logo" width="150" alt="Spark IT" style="display:block; width:150px; max-width:150px; height:auto; margin:0 0 12px auto; border:0;">
                      <span style="display:inline-block; background:#ffffff; color:#103b77; border-radius:4px; padding:8px 10px; font-size:12px; line-height:14px; font-weight:700;">Excel attached</span>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding:24px 32px 8px 32px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:$healthBg; border:1px solid $healthBorder; border-radius:8px;">
                  <tr>
                    <td style="padding:16px 18px;">
                      <div style="font-size:12px; line-height:16px; color:#5f6f85; font-weight:700; text-transform:uppercase;">Inventory status</div>
                      <div style="font-size:20px; line-height:26px; color:$healthColor; font-weight:700; margin-top:3px;">$healthText</div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding:16px 32px 4px 32px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="20%" style="padding:6px;">
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #dbe3ee; border-radius:8px; background:#f8fafc;">
                        <tr><td style="padding:14px 12px;"><div style="font-size:12px; color:#5f6f85; font-weight:700;">Active items</div><div style="font-size:26px; line-height:32px; color:#14233d; font-weight:800;">$($Summary.TotalItems)</div></td></tr>
                      </table>
                    </td>
                    <td width="20%" style="padding:6px;">
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #f5d69a; border-radius:8px; background:#fff8ea;">
                        <tr><td style="padding:14px 12px;"><div style="font-size:12px; color:#8a5a00; font-weight:700;">Low stock</div><div style="font-size:26px; line-height:32px; color:#b45309; font-weight:800;">$($Summary.LowStock)</div></td></tr>
                      </table>
                    </td>
                    <td width="20%" style="padding:6px;">
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #f2b8b5; border-radius:8px; background:#fff1f0;">
                        <tr><td style="padding:14px 12px;"><div style="font-size:12px; color:#a3342b; font-weight:700;">Out of stock</div><div style="font-size:26px; line-height:32px; color:#c7362e; font-weight:800;">$($Summary.OutOfStock)</div></td></tr>
                      </table>
                    </td>
                    <td width="20%" style="padding:6px;">
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #dbe3ee; border-radius:8px; background:#f8fafc;">
                        <tr><td style="padding:14px 12px;"><div style="font-size:12px; color:#5f6f85; font-weight:700;">Retired</div><div style="font-size:26px; line-height:32px; color:#475569; font-weight:800;">$($Summary.Retired)</div></td></tr>
                      </table>
                    </td>
                    <td width="20%" style="padding:6px;">
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #c7d8f4; border-radius:8px; background:#eff6ff;">
                        <tr><td style="padding:14px 12px;"><div style="font-size:12px; color:#24539a; font-weight:700;">Categories</div><div style="font-size:26px; line-height:32px; color:#1d4ed8; font-weight:800;">$($Summary.Categories)</div></td></tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding:20px 32px 26px 32px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #e5ebf3;">
                  <tr>
                    <td style="padding-top:18px;">
                      <div style="font-size:16px; line-height:22px; font-weight:700; color:#14233d;">Attached workbook includes</div>
                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;">
                        <tr>
                          <td style="padding:7px 0; font-size:14px; line-height:20px; color:#334155;">Full inventory list with current quantities, minimums, locations, notes, and latest update time.</td>
                        </tr>
                        <tr>
                          <td style="padding:7px 0; font-size:14px; line-height:20px; color:#334155;">Filters, frozen headings, summary totals, and colour-coded stock status for quick review.</td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="background:#f8fafc; border-top:1px solid #e5ebf3; padding:16px 32px;">
                <div style="font-size:12px; line-height:18px; color:#64748b;">This report was generated automatically by Spark IT Stocktake Report Mailer.</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
"@
}

function New-ExcelReport {
    param (
        [string] $OutputPath,
        [DateTime] $ReportDate,
        [object[]] $Rows,
        [object] $Summary,
        [string] $LogoPath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("stocktake-xlsx-{0}" -f [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tempRoot | Out-Null

    try {
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "_rels") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "docProps") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\_rels") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\drawings") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\drawings\_rels") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\media") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\worksheets") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl\worksheets\_rels") | Out-Null

        $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>
"@

        $rootRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

        $workbookRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

        $sheetRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"@

        $drawingRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/Spark_Logo_RGB.png"/>
</Relationships>
"@

        $drawing = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>6</xdr:col>
      <xdr:colOff>120000</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>100000</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="1050000" cy="795000"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Spark IT Logo" descr="Spark IT Logo"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
"@

        $workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Stocktake Report" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

        $createdAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Spark IT Monthly Stocktake Report</dc:title>
  <dc:creator>Spark IT Stocktake Report Mailer</dc:creator>
  <cp:lastModifiedBy>Spark IT Stocktake Report Mailer</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$createdAt</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$createdAt</dcterms:modified>
</cp:coreProperties>
"@

        $app = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Spark IT Stocktake Report Mailer</Application>
</Properties>
"@

        $styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="7">
    <font><sz val="11"/><color rgb="FF14233D"/><name val="Segoe UI"/></font>
    <font><b/><sz val="24"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><sz val="11"/><color rgb="FFEAF2FF"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF24539A"/><name val="Segoe UI"/></font>
    <font><b/><sz val="20"/><color rgb="FF14233D"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF14233D"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="10">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF103B77"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF103B77"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7F7ED"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF3DE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE9E9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFECEFF5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFE7ECF3"/></left><right style="thin"><color rgb="FFE7ECF3"/></right><top style="thin"><color rgb="FFE7ECF3"/></top><bottom style="thin"><color rgb="FFE7ECF3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
    <xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"@

        $summaryLabels = @("Total active items", "Low stock", "Out of stock", "Retired items", "Categories")
        $summaryValues = @($Summary.TotalItems, $Summary.LowStock, $Summary.OutOfStock, $Summary.Retired, $Summary.Categories)
        $headers = @("Item", "Category", "Location", "Qty", "Min", "Status", "Notes", "Updated")

        $sheetRows = New-Object System.Collections.Generic.List[string]
        $sheetRows.Add((New-ExcelRow -RowNumber 1 -Height 34 -Cells @((New-ExcelTextCell -Reference "A1" -Value "Spark IT Stocktake Summary" -Style 1))))
        $sheetRows.Add((New-ExcelRow -RowNumber 2 -Height 24 -Cells @((New-ExcelTextCell -Reference "A2" -Value ("Monthly report generated on {0}" -f $ReportDate.ToString("dd MMM yyyy")) -Style 2))))

        $labelCells = for ($i = 0; $i -lt $summaryLabels.Count; $i++) {
            New-ExcelTextCell -Reference ("{0}4" -f (Get-ExcelColumnName ($i + 1))) -Value $summaryLabels[$i] -Style 3
        }
        $valueCells = for ($i = 0; $i -lt $summaryValues.Count; $i++) {
            New-ExcelNumberCell -Reference ("{0}5" -f (Get-ExcelColumnName ($i + 1))) -Value $summaryValues[$i] -Style 4
        }
        $sheetRows.Add((New-ExcelRow -RowNumber 4 -Height 22 -Cells $labelCells))
        $sheetRows.Add((New-ExcelRow -RowNumber 5 -Height 30 -Cells $valueCells))

        $headerCells = for ($i = 0; $i -lt $headers.Count; $i++) {
            New-ExcelTextCell -Reference ("{0}7" -f (Get-ExcelColumnName ($i + 1))) -Value $headers[$i] -Style 5
        }
        $sheetRows.Add((New-ExcelRow -RowNumber 7 -Height 24 -Cells $headerCells))

        $rowNumber = 8
        foreach ($row in $Rows) {
            $statusStyle = switch ($row.Status) {
                "Out of stock" { 10 }
                "Low stock" { 9 }
                "Retired" { 11 }
                default { 8 }
            }

            $cells = @(
                (New-ExcelTextCell -Reference "A$rowNumber" -Value $row.Item -Style 6),
                (New-ExcelTextCell -Reference "B$rowNumber" -Value $row.Category -Style 6),
                (New-ExcelTextCell -Reference "C$rowNumber" -Value $row.Location -Style 6),
                (New-ExcelNumberCell -Reference "D$rowNumber" -Value ([int]$row.Quantity) -Style 7),
                (New-ExcelNumberCell -Reference "E$rowNumber" -Value ([int]$row.Minimum) -Style 7),
                (New-ExcelTextCell -Reference "F$rowNumber" -Value $row.Status -Style $statusStyle),
                (New-ExcelTextCell -Reference "G$rowNumber" -Value $row.Notes -Style 13),
                (New-ExcelTextCell -Reference "H$rowNumber" -Value (Get-DisplayDate -Value $row.UpdatedAt) -Style 12)
            )
            $sheetRows.Add((New-ExcelRow -RowNumber $rowNumber -Height 21 -Cells $cells))
            $rowNumber++
        }

        $lastRow = [math]::Max(7, $rowNumber - 1)
        $sheetData = $sheetRows -join [Environment]::NewLine

        $sheet = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="7" topLeftCell="A8" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A8" sqref="A8"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="31" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="3" width="20" customWidth="1"/>
    <col min="4" max="5" width="10" customWidth="1"/>
    <col min="6" max="6" width="16" customWidth="1"/>
    <col min="7" max="7" width="24" customWidth="1"/>
    <col min="8" max="8" width="22" customWidth="1"/>
  </cols>
  <sheetData>
$sheetData
  </sheetData>
  <autoFilter ref="A7:H$lastRow"/>
  <mergeCells count="2">
    <mergeCell ref="A1:F1"/>
    <mergeCell ref="A2:F2"/>
  </mergeCells>
  <pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
  <drawing r:id="rId1"/>
</worksheet>
"@

        Set-Content -LiteralPath (Join-Path $tempRoot "[Content_Types].xml") -Value $contentTypes -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "_rels\.rels") -Value $rootRelationships -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "docProps\core.xml") -Value $core -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "docProps\app.xml") -Value $app -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\workbook.xml") -Value $workbook -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\_rels\workbook.xml.rels") -Value $workbookRelationships -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\drawings\drawing1.xml") -Value $drawing -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\drawings\_rels\drawing1.xml.rels") -Value $drawingRelationships -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\styles.xml") -Value $styles -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\worksheets\sheet1.xml") -Value $sheet -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $tempRoot "xl\worksheets\_rels\sheet1.xml.rels") -Value $sheetRelationships -Encoding UTF8
        Copy-Item -LiteralPath $LogoPath -Destination (Join-Path $tempRoot "xl\media\Spark_Logo_RGB.png") -Force

        if (Test-Path $OutputPath) {
            Remove-Item -Path $OutputPath -Force
        }

        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $OutputPath)
    }
    finally {
        if (Test-Path $tempRoot) {
            Remove-Item -Path $tempRoot -Recurse -Force
        }
    }
}

if ((Test-Path $DatabasePath) -eq $False) {
    throw "Database file not found: $DatabasePath"
}

if ((Test-Path $LogoPath) -eq $False) {
    throw "Logo file not found: $LogoPath"
}

if ((Test-Path $ReportFolder) -eq $False) {
    New-Item -ItemType Directory -Path $ReportFolder | Out-Null
}

$Database = Get-Content $DatabasePath -Raw | ConvertFrom-Json
$ReportDate = Get-Date
$ReportFileDate = $ReportDate.ToString("yyyy-MM-dd")
$WorkbookPath = Join-Path $ReportFolder ("StocktakeReport-{0}.xlsx" -f $ReportFileDate)

$ReportRows = @($Database.items | ForEach-Object {
    [PSCustomObject]@{
        Item = $_.name
        Category = $_.category
        Location = $_.location
        Quantity = [int]$_.quantity
        Minimum = [int]$_.minimum
        Status = Get-StockStatus $_
        Notes = $_.notes
        UpdatedAt = $_.updatedAt
    }
})

$Summary = Get-ReportSummary -Items $Database.items
$EmailBody = New-EmailBodyHtml -ReportDate $ReportDate -Summary $Summary
New-ExcelReport -OutputPath $WorkbookPath -ReportDate $ReportDate -Rows $ReportRows -Summary $Summary -LogoPath $LogoPath

Write-Host "Styled Excel report created: $WorkbookPath"

if ($BuildOnly) {
    return
}

if ((Test-Path $CertificateFilePath) -eq $False) {
    throw "Certificate file not found: $CertificateFilePath"
}

if (!(Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions)) {
    Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Users.Actions

$Certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    $CertificateFilePath,
    $CertificatePassword,
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
)

Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -Certificate $Certificate -NoWelcome

$WorkbookAttachmentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($WorkbookPath))
$LogoAttachmentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($LogoPath))

$Message = @{
    Subject = "Spark IT Monthly Stocktake Report - $($ReportDate.ToString('dd MMM yyyy'))"
    Body = @{
        ContentType = "HTML"
        Content = $EmailBody
    }
    ToRecipients = @(
        $Recipient -split "," | ForEach-Object {
            New-EmailRecipient $_.Trim()
        }
    )
    Attachments = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = [IO.Path]::GetFileName($WorkbookPath)
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ContentBytes = $WorkbookAttachmentBytes
        },
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = [IO.Path]::GetFileName($LogoPath)
            ContentType = "image/png"
            ContentId = "spark-logo"
            IsInline = $true
            ContentBytes = $LogoAttachmentBytes
        }
    )
}

Send-MgUserMail -UserId $Sender -Message $Message -SaveToSentItems -ErrorAction Stop
Write-Host "Stocktake report sent to $Recipient"
