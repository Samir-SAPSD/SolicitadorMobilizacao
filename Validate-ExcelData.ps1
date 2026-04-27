param(
    [Parameter(Mandatory=$false)]
    [string]$ExcelPath = ".\DadosParaImportar.xlsx",
    [Parameter(Mandatory=$false)]
    [string]$SheetName = "PESSOAS"
)

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

function Convert-ExcelColumnToIndex {
    param([string]$ColumnLetters)
    $sum = 0
    foreach ($ch in $ColumnLetters.ToUpper().ToCharArray()) {
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Get-OpenXmlCellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [array]$SharedStrings,
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $cellType = $Cell.GetAttribute("t")
    $valueNode = $Cell.SelectSingleNode("x:v", $Ns)
    $inlineNode = $Cell.SelectSingleNode("x:is/x:t", $Ns)

    if ($cellType -eq "inlineStr" -and $inlineNode) {
        return $inlineNode.InnerText
    }

    if (-not $valueNode) {
        return $null
    }

    $raw = $valueNode.InnerText
    if ($cellType -eq "s") {
        $idx = 0
        if ([int]::TryParse($raw, [ref]$idx) -and $idx -ge 0 -and $idx -lt $SharedStrings.Count) {
            return $SharedStrings[$idx]
        }
        return $raw
    }

    return $raw
}

function Read-ExcelOpenXml {
    param(
        [string]$Path,
        [string]$WorksheetName,
        [bool]$IncludeEmptyColumns
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbookEntry = $zip.GetEntry("xl/workbook.xml")
        if (-not $workbookEntry) { throw "Arquivo workbook.xml não encontrado no xlsx." }

        $relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
        if (-not $relsEntry) { throw "Arquivo workbook.xml.rels não encontrado no xlsx." }

        [xml]$workbookXml = New-Object System.Xml.XmlDocument
        $wbStream = $workbookEntry.Open()
        try { $workbookXml.Load($wbStream) } finally { $wbStream.Dispose() }

        [xml]$relsXml = New-Object System.Xml.XmlDocument
        $relsStream = $relsEntry.Open()
        try { $relsXml.Load($relsStream) } finally { $relsStream.Dispose() }

        $wbNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
        $wbNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $relNs.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships")

        $sheetNode = $workbookXml.SelectSingleNode("//x:sheets/x:sheet[@name='$WorksheetName']", $wbNs)
        if (-not $sheetNode) {
            $sheetNames = @()
            $workbookXml.SelectNodes("//x:sheets/x:sheet", $wbNs) | ForEach-Object { $sheetNames += $_.GetAttribute("name") }
            throw "Aba '$WorksheetName' não encontrada. Abas disponíveis: $($sheetNames -join ', ')"
        }

        $relId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ([string]::IsNullOrWhiteSpace($relId)) { throw "Relacionamento da aba '$WorksheetName' não encontrado." }

        $targetNode = $relsXml.SelectSingleNode("//pr:Relationship[@Id='$relId']", $relNs)
        if (-not $targetNode) { throw "Target da aba '$WorksheetName' não encontrado em workbook.xml.rels." }

        $target = $targetNode.GetAttribute("Target")
        if ([string]::IsNullOrWhiteSpace($target)) { throw "Target da aba '$WorksheetName' vazio." }

        if ($target.StartsWith("/")) {
            $sheetPath = $target.TrimStart('/')
        } elseif ($target.StartsWith("xl/")) {
            $sheetPath = $target
        } else {
            $sheetPath = "xl/$target"
        }

        $sheetEntry = $zip.GetEntry($sheetPath)
        if (-not $sheetEntry) { throw "Worksheet XML não encontrado: $sheetPath" }

        $sharedStrings = @()
        $ssEntry = $zip.GetEntry("xl/sharedStrings.xml")
        if ($ssEntry) {
            [xml]$ssXml = New-Object System.Xml.XmlDocument
            $ssStream = $ssEntry.Open()
            try { $ssXml.Load($ssStream) } finally { $ssStream.Dispose() }

            $ssNs = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ssNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

            $ssXml.SelectNodes("//x:si", $ssNs) | ForEach-Object {
                $parts = @()
                $_.SelectNodes(".//x:t", $ssNs) | ForEach-Object { $parts += $_.InnerText }
                $sharedStrings += ($parts -join "")
            }
        }

        [xml]$sheetXml = New-Object System.Xml.XmlDocument
        $sheetStream = $sheetEntry.Open()
        try { $sheetXml.Load($sheetStream) } finally { $sheetStream.Dispose() }

        $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $sheetNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $rows = $sheetXml.SelectNodes("//x:sheetData/x:row", $sheetNs)
        if (-not $rows -or $rows.Count -lt 2) {
            throw "Planilha vazia ou apenas cabeçalho."
        }

        $headerMap = @{}
        $headerRow = $rows[0]
        foreach ($cell in $headerRow.SelectNodes("x:c", $sheetNs)) {
            $ref = $cell.GetAttribute("r")
            if ($ref -match '^([A-Za-z]+)') {
                $colIdx = Convert-ExcelColumnToIndex -ColumnLetters $matches[1]
                $headerText = Get-OpenXmlCellText -Cell $cell -SharedStrings $sharedStrings -Ns $sheetNs
                if (-not [string]::IsNullOrWhiteSpace("$headerText")) {
                    $headerMap[$colIdx] = "$headerText"
                }
            }
        }

        $items = @()
        for ($r = 1; $r -lt $rows.Count; $r++) {
            $rowNode = $rows[$r]
            $obj = New-Object PSCustomObject
            $hasData = $false

            $valueByCol = @{}
            foreach ($cell in $rowNode.SelectNodes("x:c", $sheetNs)) {
                $ref = $cell.GetAttribute("r")
                if ($ref -match '^([A-Za-z]+)') {
                    $colIdx = Convert-ExcelColumnToIndex -ColumnLetters $matches[1]
                    $valueByCol[$colIdx] = Get-OpenXmlCellText -Cell $cell -SharedStrings $sharedStrings -Ns $sheetNs
                }
            }

            foreach ($colIdx in ($headerMap.Keys | Sort-Object)) {
                $header = $headerMap[$colIdx]
                $cellValue = $null
                if ($valueByCol.ContainsKey($colIdx)) { $cellValue = $valueByCol[$colIdx] }

                if ($IncludeEmptyColumns) {
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value $cellValue -Force
                } elseif (-not [string]::IsNullOrWhiteSpace("$cellValue")) {
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value "$cellValue" -Force
                }

                if (-not [string]::IsNullOrWhiteSpace("$cellValue")) {
                    $hasData = $true
                }
            }

            if ($hasData) {
                $items += $obj
            }
        }

        return $items
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
}

# ===================================================================
# SCRIPT DE VALIDAÇÃO - Totalmente separado do upload
# Conecta ao SharePoint, obtém metadados das colunas, lê o Excel
# e valida TODAS as linhas. Retorna JSON no final com o resultado.
# ===================================================================

$TestMode = $false
$SiteUrl = "https://vestas.sharepoint.com/sites/CC-ControleService-BR"

if ($TestMode) {
    $ListId = "ea1e6a2e-8df6-4171-825e-1b7ecfbea7a0"
} else {
    $ListId = "2d72b0f5-d3a3-4add-a8b0-3f94de786223"
}

if (-not $ListId) {
    $result = @{ status = "error"; errors = @("ID da lista não configurado.") }
    Write-Output "---VALIDATION_JSON_START---"
    Write-Output ($result | ConvertTo-Json -Compress)
    Write-Output "---VALIDATION_JSON_END---"
    exit 1
}

# Configura TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12


# --- Instalar/Importar PnP.PowerShell apenas se necessário ---
$TargetPnPVersion = $null
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $TargetPnPVersion = "1.12.0"
}

# Só instala provider se não existir e só se for instalar módulo
$needInstallPnP = $false
if ($TargetPnPVersion) {
    $needInstallPnP = -not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version -eq $TargetPnPVersion })
} else {
    $needInstallPnP = -not (Get-Module -ListAvailable -Name PnP.PowerShell)
}

if ($needInstallPnP) {
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
    }
    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        $InstallArgs = @{ Name = "PnP.PowerShell"; Scope = "CurrentUser"; Force = $true; AllowClobber = $true; ErrorAction = "Stop" }
        if ($TargetPnPVersion) { $InstallArgs["RequiredVersion"] = $TargetPnPVersion }
        Install-Module @InstallArgs
    } catch {
        $result = @{ status = "error"; errors = @("Não foi possível instalar PnP.PowerShell: $_") }
        Write-Output "---VALIDATION_JSON_START---"
        Write-Output ($result | ConvertTo-Json -Compress)
        Write-Output "---VALIDATION_JSON_END---"
        exit 1
    }
}

try {
    if ($TargetPnPVersion) {
        Import-Module PnP.PowerShell -RequiredVersion $TargetPnPVersion -ErrorAction Stop
    } else {
        Import-Module PnP.PowerShell -ErrorAction Stop
    }
} catch {
    $result = @{ status = "error"; errors = @("Falha ao importar PnP.PowerShell: $_") }
    Write-Output "---VALIDATION_JSON_START---"
    Write-Output ($result | ConvertTo-Json -Compress)
    Write-Output "---VALIDATION_JSON_END---"
    exit 1
}

# --- Conectar ao SharePoint ---
Write-Host "[Validação] Conectando ao SharePoint..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "[Validação] Conectado com sucesso!" -ForegroundColor Green
} catch {
    $result = @{ status = "error"; errors = @("Falha ao conectar ao SharePoint: $_") }
    Write-Output "---VALIDATION_JSON_START---"
    Write-Output ($result | ConvertTo-Json -Compress)
    Write-Output "---VALIDATION_JSON_END---"
    exit 1
}

# --- Obter metadados das colunas ---
Write-Host "[Validação] Obtendo metadados das colunas do SharePoint..." -ForegroundColor Cyan
try {
    $AllFields = Get-PnPField -List $ListId | Select-Object InternalName, Title, Required, DefaultValue, TypeAsString
    Write-Host "[Validação] Campos carregados: $($AllFields.Count)" -ForegroundColor Gray
} catch {
    $result = @{ status = "error"; errors = @("Não foi possível obter campos da lista SharePoint: $_") }
    Write-Output "---VALIDATION_JSON_START---"
    Write-Output ($result | ConvertTo-Json -Compress)
    Write-Output "---VALIDATION_JSON_END---"
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}

# --- Ler Excel ---
Write-Host "[Validação] Lendo arquivo Excel: $ExcelPath (Aba: $SheetName)" -ForegroundColor Cyan
$ExcelFilePath = $ExcelPath
if ([string]::IsNullOrWhiteSpace($SheetName)) { $SheetName = "PESSOAS" }

$ItensParaValidar = @()

if (-not (Test-Path $ExcelFilePath)) {
    $result = @{ status = "error"; errors = @("Arquivo Excel não encontrado: $ExcelFilePath") }
    Write-Output "---VALIDATION_JSON_START---"
    Write-Output ($result | ConvertTo-Json -Compress)
    Write-Output "---VALIDATION_JSON_END---"
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}

$readSuccess = $false

# 1. Leitura via COM
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $workbook = $excel.Workbooks.Open((Resolve-Path $ExcelFilePath).Path)
    
    try {
        $worksheet = $workbook.Worksheets.Item($SheetName)
    } catch {
        $allSheets = foreach($s in $workbook.Worksheets) { $s.Name }
        throw "Aba '$SheetName' não encontrada. Abas disponíveis: $($allSheets -join ', ')"
    }

    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    if ($rowCount -lt 2) { throw "Planilha vazia ou apenas cabeçalho." }

    $valueArray = $usedRange.Value2
    $headers = @()
    for ($c = 1; $c -le $colCount; $c++) {
        $headers += $valueArray[1, $c]
    }

    for ($r = 2; $r -le $rowCount; $r++) {
        $obj = New-Object PSCustomObject
        $hasData = $false
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $valueArray[$r, $c]
            $header = $headers[$c-1]
            if (-not [string]::IsNullOrWhiteSpace($header)) {
                # Adiciona TODAS as colunas (inclusive vazias) para validar campos obrigatórios
                $obj | Add-Member -MemberType NoteProperty -Name $header -Value $val -Force
                if (-not [string]::IsNullOrWhiteSpace($val)) { $hasData = $true }
            }
        }
        if ($hasData) { $ItensParaValidar += $obj }
    }
    
    Write-Host "[Validação] Leitura via COM: $($ItensParaValidar.Count) linhas com dados" -ForegroundColor Green
    $readSuccess = $true
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    Write-Warning "[Validação] Falha COM: $($_.Exception.Message)"
    if ($excel) { 
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}

# 2. Fallback: Import-Excel
if (-not $readSuccess) {
    Write-Warning "[Validação] Tentando fallback OpenXML (sem Excel/Import-Excel)..."
    try {
        $ItensParaValidar = Read-ExcelOpenXml -Path (Resolve-Path $ExcelFilePath).Path -WorksheetName $SheetName -IncludeEmptyColumns $true
        if ($ItensParaValidar -and $ItensParaValidar.Count -gt 0) {
            Write-Host "[Validação] Leitura via OpenXML: $($ItensParaValidar.Count) linhas" -ForegroundColor Green
            $readSuccess = $true
        }
    } catch {
        Write-Warning "[Validação] Falha OpenXML: $($_.Exception.Message)"
    }
}

if (-not $readSuccess) {
    if (-not (Get-Module -ListAvailable -Name Import-Excel)) {
        try {
            Install-Module -Name Import-Excel -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
            Import-Module Import-Excel -ErrorAction Stop
        } catch {
            $result = @{ status = "error"; errors = @("Não foi possível instalar Import-Excel e leitura via COM falhou.") }
            Write-Output "---VALIDATION_JSON_START---"
            Write-Output ($result | ConvertTo-Json -Compress)
            Write-Output "---VALIDATION_JSON_END---"
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            exit 1
        }
    }
    if (-not (Get-Module -Name Import-Excel)) { Import-Module Import-Excel -ErrorAction SilentlyContinue }

    try {
        $ItensParaValidar = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName -ErrorAction Stop
        if (-not $ItensParaValidar) {
            $result = @{ status = "error"; errors = @("Nenhum dado encontrado na aba '$SheetName'.") }
            Write-Output "---VALIDATION_JSON_START---"
            Write-Output ($result | ConvertTo-Json -Compress)
            Write-Output "---VALIDATION_JSON_END---"
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            exit 1
        }
        Write-Host "[Validação] Leitura via Import-Excel: $($ItensParaValidar.Count) linhas" -ForegroundColor Green
    } catch {
        $result = @{ status = "error"; errors = @("Erro ao ler Excel: $_") }
        Write-Output "---VALIDATION_JSON_START---"
        Write-Output ($result | ConvertTo-Json -Compress)
        Write-Output "---VALIDATION_JSON_END---"
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        exit 1
    }
}

# ===================================================================
# VALIDAÇÕES
# ===================================================================
Write-Host ""
Write-Host "[Validação] Iniciando validações de $($ItensParaValidar.Count) linhas..." -ForegroundColor Cyan
$ValidationErrors = @()

# --- Funções auxiliares ---
function Get-RowFieldValue {
    param([PSCustomObject]$Row, [string]$FieldTitle, [string]$FieldInternalName)
    $prop = $Row.PSObject.Properties | Where-Object { 
        $_.Name.Trim() -eq $FieldTitle -or $_.Name.Trim() -eq $FieldInternalName 
    } | Select-Object -First 1
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-DateFromExcel {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) { return $null }
    $ptBR = [System.Globalization.CultureInfo]::new("pt-BR")
    try {
        # Número = OADate do Excel (independente de locale)
        if ("$Value" -match '^\d+(\.\d+)?$') {
            return [DateTime]::FromOADate([double]$Value)
        } else {
            # Forçar parse no formato PT-BR (dd/MM/yyyy), ignorando o locale da máquina
            $dateFormats = @(
                "dd/MM/yyyy",
                "d/M/yyyy",
                "dd/MM/yyyy HH:mm:ss",
                "dd/MM/yyyy H:mm:ss",
                "d/M/yyyy HH:mm:ss"
            )
            foreach ($fmt in $dateFormats) {
                $parsed = [DateTime]::MinValue
                if ([DateTime]::TryParseExact("$Value", $fmt, $ptBR, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                    return $parsed
                }
            }
            # Fallback: Parse explicitamente com cultura pt-BR
            return [DateTime]::Parse("$Value", $ptBR)
        }
    } catch {
        return $null
    }
}

function Get-ColumnMappingKey {
    param([string]$ColumnName)

    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        return [PSCustomObject]@{ DisplayKey = ""; InternalKey = "" }
    }

    $col = "$ColumnName".Trim()
    if ($col -match '_x005F_') {
        $col = $col -replace '_x005F_', '_'
    }

    # Permite cabeçalho no formato: "Display Name [InternalName]"
    if ($col -match '^(?<disp>.+?)\s*\[(?<internal>[^\[\]]+)\]\s*$') {
        $disp = "$($matches['disp'])".Trim().ToLowerInvariant()
        $internal = "$($matches['internal'])".Trim().ToLowerInvariant()
        return [PSCustomObject]@{ DisplayKey = $disp; InternalKey = $internal }
    }

    $key = $col.ToLowerInvariant()
    return [PSCustomObject]@{ DisplayKey = $key; InternalKey = $key }
}

function Get-MappedItemValues {
    param(
        [PSCustomObject]$Row,
        [array]$AllFields,
        [hashtable]$FieldMap,
        [array]$RequiredFields
    )

    $itemValues = @{}

    $Row.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | ForEach-Object {
        $val = $_.Value
        $colName = $_.Name.Trim()

        $colMap = Get-ColumnMappingKey -ColumnName $colName
        $colNameNormalized = if ($colMap.DisplayKey) { $colMap.DisplayKey } else { "$colName".ToLowerInvariant() }
        $internalHint = $colMap.InternalKey

        if ($null -eq $val -or [string]::IsNullOrWhiteSpace("$val")) {
            return
        }

        $fieldInfo = $null
        if (-not [string]::IsNullOrWhiteSpace($internalHint) -and $FieldMap.ContainsKey($internalHint)) {
            $fieldInfo = $FieldMap[$internalHint]
        }
        elseif ($FieldMap.ContainsKey($colNameNormalized)) {
            $fieldInfo = $FieldMap[$colNameNormalized]
        }
        if ($fieldInfo) {
            $realColName = $fieldInfo.InternalName

            if ($fieldInfo.TypeAsString -match "DateTime") {
                $parsedDate = ConvertTo-DateFromExcel -Value $val
                if ($parsedDate) {
                    $itemValues[$realColName] = $parsedDate
                } else {
                    $itemValues[$realColName] = $val
                }
            } else {
                $itemValues[$realColName] = $val
            }
        } elseif ($colName -ieq "Title") {
            $itemValues["Title"] = $val
        }
    }

    if (-not $itemValues.ContainsKey("Title")) {
        $possibleTitleCols = @("Nome", "Equipamento", "Modelo", "Descricao", "Tag", "Serial")
        foreach ($col in $possibleTitleCols) {
            $foundCol = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $col } | Select-Object -ExpandProperty Name -First 1
            if ($foundCol -and -not [string]::IsNullOrWhiteSpace("$($Row.$foundCol)")) {
                $itemValues["Title"] = $Row.$foundCol
                break
            }
        }
    }

    foreach ($field in $RequiredFields) {
        $internalName = $field.InternalName
        $defaultValue = $field.DefaultValue
        $isEmpty = (-not $itemValues.ContainsKey($internalName)) -or ([string]::IsNullOrWhiteSpace("" + $itemValues[$internalName]))
        if ($isEmpty -and -not [string]::IsNullOrWhiteSpace("$defaultValue")) {
            $itemValues[$internalName] = $defaultValue
        }
    }

    return $itemValues
}

$FieldMap = @{}
foreach ($f in $AllFields) {
    if (-not [string]::IsNullOrWhiteSpace("$($f.InternalName)")) {
        $FieldMap[$f.InternalName.ToLowerInvariant()] = $f
    }
    if (-not [string]::IsNullOrWhiteSpace("$($f.Title)")) {
        $FieldMap[$f.Title.ToLowerInvariant()] = $f
    }
}

# --- Identificar campos obrigatórios ---
$RequiredFields = $AllFields | Where-Object { $_.Required -eq $true }
Write-Host "[Validação] Campos obrigatórios do SharePoint: $($RequiredFields.Count)" -ForegroundColor Gray
foreach ($rf in $RequiredFields) {
    Write-Host "  - $($rf.Title) ($($rf.InternalName)) | Default: '$($rf.DefaultValue)'" -ForegroundColor Gray
}

# --- Identificar campos de data ---
$DateFields = $AllFields | Where-Object { $_.TypeAsString -eq "DateTime" }
$SolicitacaoField = $DateFields | Where-Object { $_.Title -match "Solicita" } | Select-Object -First 1
$AcessoField = $DateFields | Where-Object { $_.Title -match "Acesso" } | Select-Object -First 1
$DesmobField = $DateFields | Where-Object { $_.Title -match "Desmob" } | Select-Object -First 1

if ($SolicitacaoField) { Write-Host "[Validação] Campo Data Solicitação: $($SolicitacaoField.Title)" -ForegroundColor Gray }
if ($AcessoField) { Write-Host "[Validação] Campo Data Acesso: $($AcessoField.Title)" -ForegroundColor Gray }
if ($DesmobField) { Write-Host "[Validação] Campo Data Desmobilização: $($DesmobField.Title)" -ForegroundColor Gray }

# --- Obter nomes das colunas do Excel para debug ---
if ($ItensParaValidar.Count -gt 0) {
    $excelColumns = $ItensParaValidar[0].PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | ForEach-Object { $_.Name }
    Write-Host "[Validação] Colunas do Excel: $($excelColumns -join ', ')" -ForegroundColor Gray
}

# --- Loop de validação por linha ---
for ($i = 0; $i -lt $ItensParaValidar.Count; $i++) {
    $row = $ItensParaValidar[$i]
    $lineNum = $i + 2  # Linha 1 = cabeçalho no Excel

    # Pular linhas completamente vazias
    $hasAnyData = $false
    foreach ($prop in $row.PSObject.Properties) {
        if ($prop.MemberType -eq 'NoteProperty' -and -not [string]::IsNullOrWhiteSpace("$($prop.Value)")) {
            $hasAnyData = $true
            break
        }
    }
    if (-not $hasAnyData) { continue }

    Write-Host "[Validação] Validando linha $lineNum..." -ForegroundColor Gray

    $mappedValues = Get-MappedItemValues -Row $row -AllFields $AllFields -FieldMap $FieldMap -RequiredFields $RequiredFields

    # === VALIDAÇÃO 1: Campos obrigatórios ===
    foreach ($field in $RequiredFields) {
        $value = $null
        if ($mappedValues.ContainsKey($field.InternalName)) {
            $value = $mappedValues[$field.InternalName]
        }

        $isEmpty = ($null -eq $value) -or ([string]::IsNullOrWhiteSpace("$value"))
        $fieldDisplayName = $field.Title
        if ($isEmpty) {
            $hasDefault = -not [string]::IsNullOrWhiteSpace("$($field.DefaultValue)")
            if (-not $hasDefault) {
                $errMsg = "Linha ${lineNum}: Campo '$fieldDisplayName' é obrigatório e não possui valor padrão."
                $ValidationErrors += $errMsg
                Write-Host "  [FALHA] $errMsg" -ForegroundColor Red
            } else {
                Write-Host "  [OK] Campo '$fieldDisplayName' vazio mas possui default: '$($field.DefaultValue)'" -ForegroundColor DarkYellow
            }
        }
    }

    # === VALIDAÇÃO 2: Datas ===
    $dataSolicitacao = $null
    $dataAcesso = $null
    $dataDesmob = $null

    if ($SolicitacaoField) {
        $val = $null
        if ($mappedValues.ContainsKey($SolicitacaoField.InternalName)) {
            $val = $mappedValues[$SolicitacaoField.InternalName]
        }
        $dataSolicitacao = ConvertTo-DateFromExcel -Value $val
    }
    if ($AcessoField) {
        $val = $null
        if ($mappedValues.ContainsKey($AcessoField.InternalName)) {
            $val = $mappedValues[$AcessoField.InternalName]
        }
        $dataAcesso = ConvertTo-DateFromExcel -Value $val
    }
    if ($DesmobField) {
        $val = $null
        if ($mappedValues.ContainsKey($DesmobField.InternalName)) {
            $val = $mappedValues[$DesmobField.InternalName]
        }
        $dataDesmob = ConvertTo-DateFromExcel -Value $val
    }

    # Data de Acesso >= Data de Solicitação
    if ($dataAcesso -and $dataSolicitacao) {
        if ($dataAcesso -lt $dataSolicitacao) {
            $errMsg = "Linha ${lineNum}: Data de Acesso ($($dataAcesso.ToString('dd/MM/yyyy'))) é anterior à Data de Solicitação ($($dataSolicitacao.ToString('dd/MM/yyyy')))."
            $ValidationErrors += $errMsg
            Write-Host "  [FALHA] $errMsg" -ForegroundColor Red
        }
    }

    # Data de Desmobilização >= Data de Acesso
    if ($dataDesmob -and $dataAcesso) {
        if ($dataDesmob -lt $dataAcesso) {
            $errMsg = "Linha ${lineNum}: Data de Desmobilização ($($dataDesmob.ToString('dd/MM/yyyy'))) é anterior à Data de Acesso ($($dataAcesso.ToString('dd/MM/yyyy')))."
            $ValidationErrors += $errMsg
            Write-Host "  [FALHA] $errMsg" -ForegroundColor Red
        }
    }
}

# ===================================================================
# RESULTADO DA VALIDAÇÃO (JSON estruturado)
# ===================================================================
Write-Host ""
if ($ValidationErrors.Count -gt 0) {
    Write-Host "[Validação] FALHOU: $($ValidationErrors.Count) erro(s) encontrado(s)." -ForegroundColor Red
    $result = @{
        status = "failed"
        total_lines = $ItensParaValidar.Count
        error_count = $ValidationErrors.Count
        errors = $ValidationErrors
    }
} else {
    Write-Host "[Validação] SUCESSO: Todas as $($ItensParaValidar.Count) linhas validadas." -ForegroundColor Green
    $result = @{
        status = "success"
        total_lines = $ItensParaValidar.Count
        error_count = 0
        errors = @()
    }
}

Write-Output "---VALIDATION_JSON_START---"
Write-Output ($result | ConvertTo-Json -Compress)
Write-Output "---VALIDATION_JSON_END---"

Disconnect-PnPOnline -ErrorAction SilentlyContinue

if ($ValidationErrors.Count -gt 0) {
    exit 1
} else {
    exit 0
}
