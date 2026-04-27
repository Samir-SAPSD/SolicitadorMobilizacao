param(
    [Parameter(Mandatory=$false)]
    [string]$ExcelPath = ".\DadosParaImportar.xlsx"
)

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

# Importar módulo de regras de Analista Responsável
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$analistRulesPath = Join-Path $scriptDir "Apply-AnalistRules.ps1"
if (Test-Path $analistRulesPath) {
    . $analistRulesPath
}

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

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
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
        if ($fileStream) { $fileStream.Dispose() }
    }
}

function Resolve-SharePointDefaultValue {
    param(
        $Field,
        $DefaultValue
    )

    if ([string]::IsNullOrWhiteSpace("$DefaultValue")) {
        return $null
    }

    if ($Field.TypeAsString -notmatch "DateTime") {
        return $DefaultValue
    }

    $dv = ("$DefaultValue").Trim()
    if ($dv -match '^\[today\]$' -or $dv -match '^today\(\)$' -or $dv -match '^=today\(\)$') {
        return (Get-Date).Date
    }

    try {
        return [DateTime]::Parse($dv)
    }
    catch {
        # Default de data nao parseavel localmente: deixa o campo ausente para o SharePoint aplicar o default nativo.
        return $null
    }
}

function Get-DetailedErrorMessage {
    param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    $parts = @()

    if ($null -eq $ErrorRecord) {
        return "Erro desconhecido (ErrorRecord nulo)."
    }

    if ($ErrorRecord.Exception -and -not [string]::IsNullOrWhiteSpace("$($ErrorRecord.Exception.Message)")) {
        $parts += "Mensagem: $($ErrorRecord.Exception.Message)"
    }

    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.InnerException -and -not [string]::IsNullOrWhiteSpace("$($ErrorRecord.Exception.InnerException.Message)")) {
        $parts += "InnerException: $($ErrorRecord.Exception.InnerException.Message)"
    }

    if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace("$($ErrorRecord.ErrorDetails.Message)")) {
        $parts += "ErrorDetails: $($ErrorRecord.ErrorDetails.Message)"
    }

    if ($ErrorRecord.ScriptStackTrace -and -not [string]::IsNullOrWhiteSpace("$($ErrorRecord.ScriptStackTrace)")) {
        $parts += "ScriptStackTrace: $($ErrorRecord.ScriptStackTrace)"
    }

    if ($ErrorRecord.Exception) {
        foreach ($propName in @("ServerErrorTypeName", "ServerErrorCode", "ServerErrorValue")) {
            $prop = $ErrorRecord.Exception.PSObject.Properties[$propName]
            if ($prop -and -not [string]::IsNullOrWhiteSpace("$($prop.Value)")) {
                $parts += "${propName}: $($prop.Value)"
            }
        }
    }

    if ($parts.Count -eq 0) {
        return "Erro sem detalhes adicionais no ErrorRecord."
    }

    return ($parts -join " | ")
}

function Test-SharePointFieldValueCompatibility {
    param(
        [Parameter(Mandatory = $true)]
        $Field,
        [Parameter(Mandatory = $true)]
        $Value
    )

    $errors = @()
    $fieldType = "$($Field.TypeAsString)"

    if ($Field.ReadOnlyField -eq $true) {
        $errors += "campo somente leitura"
        return $errors
    }

    if ($Field.Hidden -eq $true) {
        $errors += "campo oculto"
        return $errors
    }

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) {
        return $errors
    }

    if ($fieldType -eq "DateTime") {
        if ($Value -isnot [DateTime]) {
            $parsedDate = [DateTime]::MinValue
            if (-not [DateTime]::TryParse("$Value", [ref]$parsedDate)) {
                $errors += "valor '$Value' inválido para DateTime"
            }
        }
    }
    elseif ($fieldType -eq "Number" -or $fieldType -eq "Currency") {
        $parsedNumber = 0.0
        if (-not [double]::TryParse("$Value", [ref]$parsedNumber)) {
            $errors += "valor '$Value' inválido para $fieldType"
        }
    }
    elseif ($fieldType -eq "Boolean") {
        $allowedBoolean = @("true", "false", "1", "0", "sim", "nao", "não")
        if ($Value -isnot [bool]) {
            $valueNorm = "$Value".Trim().ToLowerInvariant()
            if (-not ($allowedBoolean -contains $valueNorm)) {
                $errors += "valor '$Value' inválido para Boolean"
            }
        }
    }
    elseif ($fieldType -eq "Choice") {
        if ($Field.Choices -and $Field.Choices.Count -gt 0) {
            $rawValue = "$Value"
            if (-not ($Field.Choices -contains $rawValue)) {
                $choicesPreview = ($Field.Choices | Select-Object -First 8) -join ", "
                $errors += "valor '$rawValue' fora das opções válidas (exemplos: $choicesPreview)"
            }
        }
    }
    elseif ($fieldType -eq "Lookup") {
        if (-not ("$Value" -match '^\d+$')) {
            $errors += "valor '$Value' inválido para Lookup (esperado ID numérico)"
        }
    }
    elseif ($fieldType -eq "LookupMulti") {
        if ($Value -is [string]) {
            $parts = ("$Value" -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            if ($parts.Count -eq 0) {
                $errors += "valor '$Value' inválido para LookupMulti"
            }
            foreach ($p in $parts) {
                if ($p -notmatch '^\d+$') {
                    $errors += "valor '$Value' inválido para LookupMulti (IDs devem ser numéricos)"
                    break
                }
            }
        }
    }

    return $errors
}

# ──────────────────────────────────────────────────────────────────────────────
# Funções auxiliares: ID de Mobilização e leitura segura de abas Excel
# ──────────────────────────────────────────────────────────────────────────────

function Get-MaxMobilizacaoId {
    # Retorna o valor numérico (long) do maior ID_Mobilizacao existente na lista.
    # Busca apenas o campo ID_Mobilizacao e calcula o máximo em PowerShell.
    param([string]$ListId)
    try {
        $items = Get-PnPListItem -List $ListId -Fields "ID_Mobilizacao" -PageSize 5000 -ErrorAction Stop
        $maxVal = [long]0
        foreach ($item in $items) {
            $idStr = $item.FieldValues["ID_Mobilizacao"]
            if (-not [string]::IsNullOrWhiteSpace($idStr)) {
                try {
                    $val = [Convert]::ToInt64($idStr.Trim(), 16)
                    if ($val -gt $maxVal) { $maxVal = $val }
                } catch { }
            }
        }
        return $maxVal
    }
    catch {
        Write-Warning "Não foi possível ler ID_Mobilizacao máximo: $_"
        return [long]0
    }
}

function Format-MobilizacaoId {
    param([long]$Value)
    $hex = $Value.ToString("X")
    if ($hex.Length -lt 4) { $hex = $hex.PadLeft(4, '0') }
    return $hex
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

function Read-ExcelSheetSafe {
    param(
        [string]$ExcelFilePath,
        [string]$SheetName
    )

    # 1. TENTATIVA PRIORITÁRIA: OPENXML
    try {
        $items = @(Read-ExcelOpenXml -Path (Resolve-Path $ExcelFilePath).Path -WorksheetName $SheetName -IncludeEmptyColumns $false)
        Write-Host "  Aba '$SheetName': OpenXML OK — $($items.Count) item(ns)." -ForegroundColor Green
        return $items
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match "não encontrada|Planilha vazia") {
            Write-Host "  Aba '$SheetName': não encontrada ou vazia (OpenXML)." -ForegroundColor Yellow
            return @()
        }
        Write-Warning "  Aba '$SheetName': falha OpenXML ($msg)."
    }

    # 2. TENTATIVA SECUNDÁRIA: VIA COM (EXCEL INSTALADO)
    $comExcel = $null; $comWorkbook = $null; $comWorksheet = $null; $comUsedRange = $null
    try {
        Write-Host "  Aba '$SheetName': tentando via COM..." -ForegroundColor Gray
        $comExcel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $comExcel.Visible = $false
        $comExcel.DisplayAlerts = $false
        $comWorkbook = $comExcel.Workbooks.Open((Resolve-Path $ExcelFilePath).Path)
        try { $comWorksheet = $comWorkbook.Worksheets.Item($SheetName) } catch { }
        if (-not $comWorksheet) {
            Write-Host "  Aba '$SheetName': não encontrada (COM)." -ForegroundColor Yellow
            return @()
        }
        $comUsedRange = $comWorksheet.UsedRange
        $rowCount = $comUsedRange.Rows.Count
        $colCount = $comUsedRange.Columns.Count
        if ($rowCount -lt 2) {
            Write-Host "  Aba '$SheetName': sem dados (COM)." -ForegroundColor Yellow
            return @()
        }
        $valueArray = $comUsedRange.Value2
        $headers = @(); for ($c = 1; $c -le $colCount; $c++) { $headers += $valueArray[1, $c] }
        $comItems = @()
        for ($r = 2; $r -le $rowCount; $r++) {
            $obj = New-Object PSCustomObject
            $hasData = $false
            for ($c = 1; $c -le $colCount; $c++) {
                $val = $valueArray[$r, $c]
                if (-not [string]::IsNullOrWhiteSpace("$val")) {
                    $header = $headers[$c - 1]
                    if (-not [string]::IsNullOrWhiteSpace("$header")) {
                        $obj | Add-Member -MemberType NoteProperty -Name $header -Value "$val" -Force
                        $hasData = $true
                    }
                }
            }
            if ($hasData) { $comItems += $obj }
        }
        Write-Host "  Aba '$SheetName': COM OK — $($comItems.Count) item(ns)." -ForegroundColor Green
        return $comItems
    }
    catch {
        Write-Warning "  Aba '$SheetName': falha COM ($($_.Exception.Message))."
    }
    finally {
        try { if ($comUsedRange)  { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comUsedRange) } }  catch {}
        try { if ($comWorksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comWorksheet) } } catch {}
        try { if ($comWorkbook)  { $comWorkbook.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comWorkbook) } } catch {}
        try { if ($comExcel)     { $comExcel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comExcel) } }     catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    # 3. TENTATIVA TERCIÁRIA: IMPORT-EXCEL
    try {
        Write-Warning "  Aba '$SheetName': tentando via Import-Excel..."
        if (-not (Get-Module -ListAvailable -Name Import-Excel)) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            try {
                Install-Module -Name Import-Excel -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
            }
            catch {
                if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction SilentlyContinue
                }
                Install-Package -Name Import-Excel -Source "https://www.powershellgallery.com/api/v2" -Scope CurrentUser -Force -ErrorAction Stop
            }
        }
        if (-not (Get-Module -Name Import-Excel)) { Import-Module Import-Excel -ErrorAction SilentlyContinue }
        $ieItems = @(Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName -ErrorAction Stop)
        Write-Host "  Aba '$SheetName': Import-Excel OK — $($ieItems.Count) item(ns)." -ForegroundColor Green
        return $ieItems
    }
    catch {
        Write-Warning "  Aba '$SheetName': falha Import-Excel ($($_.Exception.Message))."
    }

    return @()
}

# Configurações
$TestMode = $false # Altere para $false para usar a lista de produção
$SiteUrl = "https://vestas.sharepoint.com/sites/CC-ControleService-BR"

# Centralização da lista alvo (conforme solicitado, a lista é a mesma para Pessoas e Equipamentos)
if ($TestMode) {
    Write-Host "--- MODO DE TESTE ATIVADO ---" -ForegroundColor Yellow
    $ListId = "ea1e6a2e-8df6-4171-825e-1b7ecfbea7a0"
} else {
    $ListId = "2d72b0f5-d3a3-4add-a8b0-3f94de786223"
}

if (-not $ListId) {
    Write-Error "ID da lista não configurado."
    exit
}

# Configura TLS 1.2 (necessário para PSGallery)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Verifica/Instala o provedor NuGet antes de tentar instalar módulos
if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Warning "Instalando provedor NuGet..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
}

# Lógica de Versão PnP: Windows PowerShell 5.1 detectado.
$TargetPnPVersion = $null
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "Ambiente: Windows PowerShell 5.1 detectado."
    Write-Warning "Forçando uso da versão legacy 1.12.0 do PnP.PowerShell (versões 2.0+ requerem PowerShell 7)."
    $TargetPnPVersion = "1.12.0"
}

# Verifica se a versão correta está instalada
$IsInstalled = if ($TargetPnPVersion) {
    Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version -eq $TargetPnPVersion }
} else {
    Get-Module -ListAvailable -Name PnP.PowerShell
}

if (-not $IsInstalled) {
    $vMsg = if ($TargetPnPVersion) { " v$TargetPnPVersion" } else { "" }
    Write-Warning "Instalando PnP.PowerShell$vMsg..."
    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        
        $InstallArgs = @{
            Name = "PnP.PowerShell"
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            ErrorAction = "Stop"
        }
        if ($TargetPnPVersion) { $InstallArgs["RequiredVersion"] = $TargetPnPVersion }

        Install-Module @InstallArgs
    }
    catch {
        Write-Error "ERRO CRÍTICO: Não foi possível instalar o módulo. Detalhes: $_"
        Write-Host "Execute manualmente:" -ForegroundColor Yellow
        $cmd = "Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
        if ($TargetPnPVersion) { $cmd += " -RequiredVersion $TargetPnPVersion" }
        Write-Host $cmd -ForegroundColor Yellow
        exit 1
    }
}

# Importação Explícita da Versão Correta
try {
    if ($TargetPnPVersion) {
        # Tenta carregar a versão específica
        Import-Module PnP.PowerShell -RequiredVersion $TargetPnPVersion -ErrorAction Stop
    } else {
        Import-Module PnP.PowerShell -ErrorAction Stop
    }
}
catch {
    Write-Error "ERRO CRÍTICO: Falha ao importar PnP.PowerShell. Detalhes: $_"
    Write-Host "Se você está no PowerShell 5.1 e a atualização falhou, tente instalar o PowerShell 7 manualmente: https://aka.ms/PS7" -ForegroundColor Yellow
    exit 1
}

# Conectar ao SharePoint
# O parâmetro -Interactive abrirá uma janela para login via navegador (MFA suportado)
# Se preferir usar ClientId/ClientSecret ou credenciais, altere este comando.
try {
    Write-Host "Conectando ao site: $SiteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Conectado com sucesso!" -ForegroundColor Green
    
    if ($TestMode) {
        # try {
        #     # Tenta obter usuário, se falhar usa um padrão
        #     try { $currentUserEmail = (Get-PnPCurrentUser).Email } catch { $currentUserEmail = "sapsd@vestas.com" }
            
        #     Write-Host "Modo de Teste: Regenerando dados (Smart Generator)..." -ForegroundColor Yellow
        #     .\Smart-Generate-TestData.ps1 -UserEmail $currentUserEmail
        # }
        # catch {
        #     Write-Warning "Não foi possível gerar dados de teste: $_"
        # }
        Write-Host "Modo de Teste: Usando arquivo Excel existente." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Falha ao conectar ao SharePoint: $_"
    exit
}

# Mapeamento de Campos
Write-Host "Mapeando campos da lista..." -ForegroundColor Cyan

# LISTA MESTRE PARA LOOKUPS DE "PARQUE" - Definida pelo usuário
$ParqueLookupListId = "678f10f9-8d46-404b-a451-70dfe938a1ee"

try {
    # Obter TODOS os campos para mapear Excel Title -> SharePoint InternalName
    $AllFields = Get-PnPField -List $ListId | Select-Object InternalName, Title, TypeAsString, LookupList, LookupField, Required, DefaultValue, ReadOnlyField, Hidden, Choices
    $LookupFields = $AllFields | Where-Object { $_.TypeAsString -eq "Lookup" }
    $LookupCache = @{} # Cache: "FieldName:Value" -> ID
    $LookupDatasetCache = @{} # Cache: "ListId|Field" -> Itens da lista de lookup
    $LookupExactIndexCache = @{} # Cache: "ListId|Field" -> Hashtable(normalizedValue -> Id)
    $FieldMap = @{} # Cache: lower(title|internalname) -> FieldInfo
    $FieldByInternalName = @{} # Cache: internalname -> FieldInfo

    foreach ($f in $AllFields) {
        if (-not [string]::IsNullOrWhiteSpace("$($f.InternalName)")) {
            $FieldMap[$f.InternalName.ToLowerInvariant()] = $f
            $FieldByInternalName[$f.InternalName] = $f
        }
        if (-not [string]::IsNullOrWhiteSpace("$($f.Title)")) {
            $FieldMap[$f.Title.ToLowerInvariant()] = $f
        }
    }
    
    if ($AllFields) {
        Write-Host "Campos da lista carregados: $($AllFields.Count)" -ForegroundColor Gray
    }
}
catch {
    Write-Warning "Não foi possível obter campos da lista. A importação pode falhar se os nomes das colunas não forem exatos."
}

# Configurações do Arquivo Excel
$ExcelFilePath = $ExcelPath

# Ler dados do Excel — abas PESSOAS e EQUIPAMENTOS (submissão unificada)
if (-not (Test-Path $ExcelFilePath)) {
    Write-Error "Arquivo Excel não encontrado: $ExcelFilePath"
    exit 1
}

Write-Host "Lendo arquivo Excel: $ExcelFilePath" -ForegroundColor Cyan
$ItensPessoas      = @(Read-ExcelSheetSafe -ExcelFilePath $ExcelFilePath -SheetName "PESSOAS")
$ItensEquipamentos = @(Read-ExcelSheetSafe -ExcelFilePath $ExcelFilePath -SheetName "EQUIPAMENTOS")

foreach ($it in $ItensPessoas) {
    try { $it | Add-Member -MemberType NoteProperty -Name "OrigemAba" -Value "PESSOAS" -Force } catch {}
}
foreach ($it in $ItensEquipamentos) {
    try { $it | Add-Member -MemberType NoteProperty -Name "OrigemAba" -Value "EQUIPAMENTOS" -Force } catch {}
}

$ItensParaAdicionar = @()
if ($ItensPessoas.Count -gt 0) {
    Write-Host "  PESSOAS: $($ItensPessoas.Count) item(ns) carregado(s)." -ForegroundColor Cyan
    $ItensParaAdicionar += $ItensPessoas
}
if ($ItensEquipamentos.Count -gt 0) {
    Write-Host "  EQUIPAMENTOS: $($ItensEquipamentos.Count) item(ns) carregado(s)." -ForegroundColor Cyan
    $ItensParaAdicionar += $ItensEquipamentos
}

if ($ItensParaAdicionar.Count -eq 0) {
    Write-Error "Nenhum dado encontrado nas abas PESSOAS ou EQUIPAMENTOS."
    exit 1
}
Write-Host "Total de itens a processar: $($ItensParaAdicionar.Count)" -ForegroundColor Cyan

# === APLICAR REGRAS DE PREENCHIMENTO AUTOMÁTICO DO ANALISTA RESPONSÁVEL ===
Write-Host ""
Write-Host "Aplicando Regras de Preenchimento - Analista Responsável" -ForegroundColor Cyan

# Validar se a função Apply-AnalistRules foi carregada
if (Get-Command -Name Apply-AnalistRules -ErrorAction SilentlyContinue) {
    try {
        $ItensParaAdicionar = @(Apply-AnalistRules -Items $ItensParaAdicionar -ParqueLookupListId $ParqueLookupListId -SiteUrl $SiteUrl -DetailedLog)
        Write-Host "Regras aplicadas com sucesso!" -ForegroundColor Green
    } catch {
        Write-Warning "Erro ao aplicar regras de Analista: $_"
    }
} else {
    Write-Warning "Módulo Apply-AnalistRules não disponível. Pulando preenchimento automático do Analista."
}

Write-Host ""

# Loop para preparar e validar os itens antes do envio (evita envio parcial)
$ExecutionReport = @()
$PreparedItems = @()
$BlockingErrors = @()
$RequiredFields = $AllFields | Where-Object { $_.Required -eq $true }

for ($rowIndex = 0; $rowIndex -lt $ItensParaAdicionar.Count; $rowIndex++) {
    $Row = $ItensParaAdicionar[$rowIndex]
    $lineNum = $rowIndex + 2 # Linha 1 = cabeçalho

    try {
        # Converte a linha do Excel (PSCustomObject) para Hashtable
        $ItemValues = @{}
        $DisplayValues = @{} # Valores legíveis para exibição no relatório (lookup: nome em vez de ID)
        $Row.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | ForEach-Object {
            $val = $_.Value
            $colName = $_.Name.Trim() # Remove espaços extras do nome da coluna

            $colMap = Get-ColumnMappingKey -ColumnName $colName
            $colNameNormalized = if ($colMap.DisplayKey) { $colMap.DisplayKey } else { "$colName".ToLowerInvariant() }
            $internalHint = $colMap.InternalKey

            # Verifica se não é nulo e se, convertido para string, não é vazio ou apenas espaços
            if ($null -ne $val -and "$val".Trim().Length -gt 0) {
                
                # 1. Identificar o campo SharePoint correto (por Title ou InternalName)
                $fieldInfo = $null
                if (-not [string]::IsNullOrWhiteSpace($internalHint) -and $FieldMap.ContainsKey($internalHint)) {
                    $fieldInfo = $FieldMap[$internalHint]
                }
                elseif ($FieldMap.ContainsKey($colNameNormalized)) {
                    $fieldInfo = $FieldMap[$colNameNormalized]
                }
                if ($fieldInfo) {
                    $realColName = $fieldInfo.InternalName

                    if ($fieldInfo.ReadOnlyField -eq $true -or $fieldInfo.Hidden -eq $true) {
                        continue
                    }

                    # 2. Se for Lookup, resolve o ID
                    if ($fieldInfo.TypeAsString -match "Lookup") {
                        if ($val -match '^\d+$') {
                            $ItemValues[$realColName] = $val
                            # Tenta resolver nome a partir do ID para exibição no relatório
                            $targetListId2 = $fieldInfo.LookupList
                            $targetField2  = if ($fieldInfo.LookupField) { $fieldInfo.LookupField } else { "Title" }
                            if ($realColName -match "Parque" -or $colName -match "Parque") {
                                $targetListId2 = $ParqueLookupListId
                                $targetField2  = "Title"
                            }
                            $datasetKey2 = "$targetListId2|$targetField2"
                            if ($LookupDatasetCache.ContainsKey($datasetKey2)) {
                                $matchItem = $LookupDatasetCache[$datasetKey2] | Where-Object { $_.Id -eq [int]$val } | Select-Object -First 1
                                if ($matchItem) { $DisplayValues[$realColName] = "$($matchItem.FieldValues[$targetField2])" }
                            }
                            if (-not $DisplayValues.ContainsKey($realColName)) { $DisplayValues[$realColName] = "$val" }
                        }
                        else {
                            $cacheKey = "${realColName}:${val}"
                            if ($LookupCache.ContainsKey($cacheKey)) {
                                $ItemValues[$realColName] = $LookupCache[$cacheKey]
                                $DisplayValues[$realColName] = "$val"
                            }
                            else {
                                # Write-Host "Resolvendo Lookup '$colName' ($realColName) para valor '$val'..." -NoNewline -ForegroundColor Gray
                                try {
                                    $targetListId = $fieldInfo.LookupList
                                    $targetInternalField = $fieldInfo.LookupField
                                    
                                    # Fallback "Parque"
                                    if ($realColName -match "Parque" -or $colName -match "Parque") {
                                        $targetListId = $ParqueLookupListId
                                        $targetInternalField = "Title"
                                    }

                                    $searchVal = "$val".Trim()
                                    $datasetKey = "$targetListId|$targetInternalField"

                                    if (-not $LookupDatasetCache.ContainsKey($datasetKey)) {
                                        $allItems = Get-PnPListItem -List $targetListId -PageSize 2000 -ErrorAction SilentlyContinue
                                        if (-not $allItems) { $allItems = @() }
                                        $LookupDatasetCache[$datasetKey] = $allItems

                                        $exactIndex = @{}
                                        foreach ($li in $allItems) {
                                            $fv = $li.FieldValues[$targetInternalField]
                                            if ($null -ne $fv) {
                                                $normalized = ("$fv").Trim().ToLowerInvariant()
                                                if (-not [string]::IsNullOrWhiteSpace($normalized) -and -not $exactIndex.ContainsKey($normalized)) {
                                                    $exactIndex[$normalized] = $li.Id
                                                }
                                            }
                                        }
                                        $LookupExactIndexCache[$datasetKey] = $exactIndex
                                    }

                                    $foundId = $null
                                    $searchNorm = $searchVal.ToLowerInvariant()

                                    if ($LookupExactIndexCache.ContainsKey($datasetKey) -and $LookupExactIndexCache[$datasetKey].ContainsKey($searchNorm)) {
                                        $foundId = $LookupExactIndexCache[$datasetKey][$searchNorm]
                                    }

                                    if (-not $foundId) {
                                        $allItems = $LookupDatasetCache[$datasetKey]
                                        $foundItem = $allItems | Where-Object { $_.FieldValues[$targetInternalField] -ilike "*$searchVal*" } | Select-Object -First 1
                                        if ($foundItem) { $foundId = $foundItem.Id }
                                    }

                                    if ($foundId) {
                                        $LookupCache[$cacheKey] = $foundId
                                        $ItemValues[$realColName] = $foundId
                                        $DisplayValues[$realColName] = $searchVal
                                        # Write-Host " [OK ID: $foundId]" -ForegroundColor Green
                                    } else {
                                        Write-Host " [Parque não encontrado]" -ForegroundColor Red
                                    }
                                } catch {
                                    Write-Host " [Erro]" -ForegroundColor Red
                                }
                            }
                        }
                    } 
                    elseif ($fieldInfo.TypeAsString -match "DateTime") {
                        # 3. Tratamento especial para Datas (Excel via COM retorna números OLE Automation)
                        try {
                            if ($val -match '^\d+(\.\d+)?$') {
                                # É um número (ex: 45302.5), converte de OADate
                                $ItemValues[$realColName] = [DateTime]::FromOADate([double]$val)
                            } else {
                                # Tenta converter string de data
                                $ItemValues[$realColName] = [DateTime]::Parse("$val")
                            }
                        }
                        catch {
                            $BlockingErrors += "Linha ${lineNum}: valor de data invalido para o campo '$realColName': '$val'."
                        }
                    }
                    else {
                        # 4. Campo comum: Mapeia para o InternalName correto
                        $ItemValues[$realColName] = $val
                    }
                } else {
                    $normalizedColName = Normalize-TextForCompare -Text $colName
                    if ($normalizedColName.Contains("ANALISTA") -and $normalizedColName.Contains("RESP")) {
                        $ItemValues["AnalistaRespons_x00e1_vel"] = $val
                    }
                    if ($colName -ieq "Title") { $ItemValues["Title"] = $val }
                }
            }
        }

        # GARANTIA DE 'TITLE': Se a coluna Title estiver vazia mas houver colunas comuns de Equipamento, mapeia para Title
        if (-not $ItemValues.ContainsKey("Title")) {
            $possibleTitleCols = @("Nome", "Equipamento", "Modelo", "Descricao", "Tag", "Serial")
            foreach ($col in $possibleTitleCols) {
                # Procura nas chaves originais da linha do Excel
                $foundCol = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $col } | Select-Object -ExpandProperty Name -First 1
                if ($foundCol -and $Row.$foundCol) {
                    $ItemValues["Title"] = $Row.$foundCol
                    # Write-Host " [Auto-mapeamento $foundCol -> Title]" -ForegroundColor Gray
                    break
                }
            }
        }

        # Preencher campos obrigatórios vazios com DefaultValue do SharePoint
        foreach ($field in $RequiredFields) {
            $internalName = $field.InternalName
            $defaultValue = $field.DefaultValue
            $isEmpty = (-not $ItemValues.ContainsKey($internalName)) -or ([string]::IsNullOrWhiteSpace("" + $ItemValues[$internalName]))
            if ($isEmpty -and -not [string]::IsNullOrWhiteSpace($defaultValue)) {
                $resolvedDefault = Resolve-SharePointDefaultValue -Field $field -DefaultValue $defaultValue
                if ($null -ne $resolvedDefault) {
                    $ItemValues[$internalName] = $resolvedDefault
                }
            }
        }

        # Bloqueia se existir campo obrigatório sem valor e sem default
        $missingRequired = @()
        foreach ($field in $RequiredFields) {
            $internalName = $field.InternalName
            $fieldTitle = $field.Title
            $defaultValue = $field.DefaultValue
            $isEmpty = (-not $ItemValues.ContainsKey($internalName)) -or ([string]::IsNullOrWhiteSpace("" + $ItemValues[$internalName]))
            $hasDefault = -not [string]::IsNullOrWhiteSpace("$defaultValue")

            if ($isEmpty -and -not $hasDefault) {
                $missingRequired += "$fieldTitle ($internalName)"
            }
        }

        if ($missingRequired.Count -gt 0) {
            $BlockingErrors += "Linha ${lineNum}: campos obrigatórios sem valor e sem default no SharePoint: $($missingRequired -join ', ')"
            continue
        }

        # Se não houver nenhuma coluna com dados, pula a linha
        if ($ItemValues.Count -eq 0) {
            Write-Warning "Linha sem dados encontrada no Excel. Ignorando..."
            continue
        }

        # Validação de compatibilidade de tipo antes do Add-PnPListItem
        $rowTypeErrors = @()
        foreach ($key in $ItemValues.Keys) {
            if ($FieldByInternalName.ContainsKey($key)) {
                $fieldDef = $FieldByInternalName[$key]
                $compatErrors = @(Test-SharePointFieldValueCompatibility -Field $fieldDef -Value $ItemValues[$key])
                foreach ($ce in $compatErrors) {
                    $rowTypeErrors += "Campo '$key': $ce"
                }
            }
        }

        if ($rowTypeErrors.Count -gt 0) {
            $BlockingErrors += "Linha ${lineNum}: incompatibilidade de tipo detectada: $($rowTypeErrors -join ' | ')"
            continue
        }

        # Remover ID_Mobilizacao eventualmente presente no Excel (será gerado pelo programa)
        $ItemValues.Remove("ID_Mobilizacao") | Out-Null

        $PreparedItems += [PSCustomObject]@{
            LinhaExcel = $lineNum
            Values = $ItemValues
            DisplayValues = $DisplayValues
            OrigemAba = if ($Row.PSObject.Properties['OrigemAba']) { "$($Row.OrigemAba)" } else { "DESCONHECIDA" }
        }
    }
    catch {
        $errorDetail = $_.Exception.Message
        $BlockingErrors += "Linha ${lineNum}: erro ao preparar dados para envio: $errorDetail"
    }
}

if ($BlockingErrors.Count -gt 0) {
    Write-Host "" 
    Write-Host "UPLOAD CANCELADO: Foram encontrados campos obrigatórios sem valor e sem default." -ForegroundColor Red
    foreach ($be in $BlockingErrors) {
        Write-Host " - $be" -ForegroundColor Red
    }
    exit 1
}

if ($PreparedItems.Count -eq 0) {
    Write-Error "Nenhum item válido para envio foi preparado."
    exit 1
}

# === GERAÇÃO DO ID DE MOBILIZAÇÃO E SUBMISSÃO UNIFICADA ===
Write-Host ""
Write-Host "Gerando ID de Mobilização único para esta submissão..." -ForegroundColor Cyan

$ID_Mobilizacao   = $null
$MaxMobRetries    = 5
$MobRetryCount    = 0
$MobSuccess       = $false
$CreatedItemSPIds = @()

while (-not $MobSuccess -and $MobRetryCount -lt $MaxMobRetries) {

    if ($MobRetryCount -gt 0) {
        $delay = 500 + (Get-Random -Minimum 100 -Maximum 500)
        Write-Warning "  Tentativa $($MobRetryCount + 1)/$MaxMobRetries — aguardando ${delay}ms..."
        Start-Sleep -Milliseconds $delay
    }

    # Leitura única: busca apenas o ID máximo atual via CAML (1 item ordenado DESC)
    $maxVal      = Get-MaxMobilizacaoId -ListId $ListId
    $candidateId = Format-MobilizacaoId -Value ($maxVal + 1)

    Write-Host "  ID_Mobilizacao gerado: $candidateId" -ForegroundColor Cyan

    # Submissão de todos os itens (PESSOAS + EQUIPAMENTOS) com o candidateId
    $submissionOk    = $true
    $CreatedItemSPIds = @()

    foreach ($prepared in $PreparedItems) {
        try {
            $ItemValues = $prepared.Values
            $ItemValues["ID_Mobilizacao"] = $candidateId
            $origemAba = if ([string]::IsNullOrWhiteSpace("$($prepared.OrigemAba)")) { "DESCONHECIDA" } else { "$($prepared.OrigemAba)" }

            $reportTitle = if ($ItemValues.Title) { $ItemValues.Title } else { "Item" }
            Write-Host "Adicionando item [$origemAba]: $reportTitle..." -NoNewline

            $novoItem = Add-PnPListItem -List $ListId -Values $ItemValues -ErrorAction Stop
            $CreatedItemSPIds += $novoItem.Id
            Write-Host " [OK] (ID SP: $($novoItem.Id))" -ForegroundColor Green
            Write-Host "--- RESULT: SUCCESS:$origemAba ---" -ForegroundColor Gray

            $ExecutionReport += [PSCustomObject]@{
                "Linha"          = $prepared.LinhaExcel
                "Item"           = $reportTitle
                "ID_SP"          = $novoItem.Id
                "ID_Mobilizacao" = $candidateId
                "Status"         = "Sucesso"
            }
        }
        catch {
            Write-Host ' [ERRO]' -ForegroundColor Red
            $errorDetail    = Get-DetailedErrorMessage -ErrorRecord $_
            $payloadPreview = ($ItemValues.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key)='$($_.Value)'" }) -join "; "
            Write-Error "Erro ao adicionar item (linha $($prepared.LinhaExcel)): $errorDetail"
            Write-Host "Campos enviados: $($ItemValues.Keys -join ', ')" -ForegroundColor DarkYellow
            Write-Host "Payload: $payloadPreview" -ForegroundColor DarkYellow
            Write-Host "--- RESULT: ERROR:$origemAba ---" -ForegroundColor Gray

            $ExecutionReport += [PSCustomObject]@{
                "Linha"          = $prepared.LinhaExcel
                "Item"           = "Erro na linha"
                "ID_SP"          = "N/A"
                "ID_Mobilizacao" = $candidateId
                "Status"         = "Erro: $errorDetail"
            }
            $submissionOk = $false
        }
    }

    # Se houve erros de envio (não relacionados ao ID), encerra o loop sem retry de ID
    if (-not $submissionOk) {
        $MobSuccess = $true
        break
    }

    # ── Validação pós-criação: confirmar unicidade do ID_Mobilizacao ──────────
    Write-Host ""
    Write-Host "  Validando unicidade de '$candidateId'..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 800  # Aguarda propagação no SharePoint

    try {
        $camlQuery   = "<View><Query><Where><Eq><FieldRef Name='ID_Mobilizacao'/><Value Type='Text'>$candidateId</Value></Eq></Where></Query><RowLimit>500</RowLimit></View>"
        $verifyItems = @(Get-PnPListItem -List $ListId -Query $camlQuery -ErrorAction Stop)

        if ($verifyItems.Count -gt $CreatedItemSPIds.Count) {
            # Colisão confirmada: outros itens gravaram o mesmo ID concorrentemente
            Write-Warning "  Colisão confirmada: $($verifyItems.Count) itens com '$candidateId', esperado $($CreatedItemSPIds.Count)."
            Write-Host "  Corrigindo IDs dos itens criados nesta submissão..." -ForegroundColor Yellow

            $maxValAfter = Get-MaxMobilizacaoId -ListId $ListId
            $correctedId = Format-MobilizacaoId -Value ($maxValAfter + 1)

            $correctionOk = $true
            foreach ($spId in $CreatedItemSPIds) {
                try {
                    Set-PnPListItem -List $ListId -Identity $spId -Values @{ "ID_Mobilizacao" = $correctedId } -ErrorAction Stop
                }
                catch {
                    Write-Warning "  Falha ao corrigir item SP $spId : $_"
                    $correctionOk = $false
                }
            }

            if ($correctionOk) {
                Write-Host "  Correção concluída. Novo ID_Mobilizacao: $correctedId" -ForegroundColor Green
                # Atualiza o relatório com o ID corrigido
                foreach ($rep in $ExecutionReport) {
                    if ($rep.ID_Mobilizacao -eq $candidateId) {
                        $rep | Add-Member -MemberType NoteProperty -Name "ID_Mobilizacao" -Value $correctedId -Force
                    }
                }
                $candidateId = $correctedId
            }
            else {
                Write-Warning "  Correção incompleta. Verifique manualmente os itens SP: $($CreatedItemSPIds -join ', ')"
            }
        }
        else {
            Write-Host "  Unicidade confirmada: $($verifyItems.Count) item(ns) com ID '$candidateId'." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "  Validação pós-criação falhou: $_ — Verifique manualmente."
    }

    $ID_Mobilizacao = $candidateId
    $MobSuccess     = $true
}

if (-not $MobSuccess -and $MobRetryCount -ge $MaxMobRetries) {
    Write-Error "Não foi possível gerar um ID_Mobilizacao único após $MaxMobRetries tentativas."
    exit 1
}

Write-Host ""
Write-Host '=== RELATORIO DE IMPORTACAO ===' -ForegroundColor Cyan
if ($ExecutionReport) {
    $ExecutionReport | Format-Table -Property Linha, Item, ID_SP, ID_Mobilizacao, Status -AutoSize | Out-String | Write-Host
} else {
    Write-Host 'Nenhum item foi processado.' -ForegroundColor Yellow
}

Write-Host '=== FINAL SUMMARY ===' -ForegroundColor Cyan
if ($ID_Mobilizacao) {
    Write-Host "ID de Mobilizacao desta submissao: $ID_Mobilizacao" -ForegroundColor Green
}
Write-Host 'Processo finalizado.' -ForegroundColor Cyan

# === GERAÇÃO DE JSON PARA RELATÓRIO EXCEL ===
try {
    $submissionDate = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    # Obter nome e e-mail do solicitante (usuário conectado ao SharePoint)
    $requesterName  = ""
    $requesterEmail = ""
    try {
        $currentUser    = Get-PnPCurrentUser -ErrorAction SilentlyContinue
        if ($currentUser) {
            $requesterName  = "$($currentUser.Title)"
            $requesterEmail = "$($currentUser.Email)"
        }
    } catch { }

    # Mapear InternalName -> Display Name para todos os campos enviados
    $fieldDisplayMap = @{}
    foreach ($prepared in $PreparedItems) {
        foreach ($key in @($prepared.Values.Keys)) {
            if ($key -eq "ID_Mobilizacao") { continue }
            if (-not $fieldDisplayMap.ContainsKey($key)) {
                $displayName = if ($FieldByInternalName.ContainsKey($key)) { $FieldByInternalName[$key].Title } else { $key }
                $fieldDisplayMap[$key] = $displayName
            }
        }
    }

    # Construir lista de itens do relatório
    $reportItems = [System.Collections.Generic.List[object]]::new()
    foreach ($rep in $ExecutionReport) {
        $matching = $PreparedItems | Where-Object { $_.LinhaExcel -eq $rep.Linha } | Select-Object -First 1
        $fieldValues = [ordered]@{}
        if ($matching) {
            foreach ($key in @($matching.Values.Keys)) {
                if ($key -eq "ID_Mobilizacao") { continue }
                $dispName = if ($fieldDisplayMap.ContainsKey($key)) { $fieldDisplayMap[$key] } else { $key }
                # Preferir valor legível (DisplayValues) para campos Lookup
                if ($matching.DisplayValues -and $matching.DisplayValues.ContainsKey($key)) {
                    $rawVal = $matching.DisplayValues[$key]
                } else {
                    $rawVal = $matching.Values[$key]
                }
                # Datas: formatar para PT-BR
                if ($rawVal -is [DateTime]) {
                    $rawVal = $rawVal.ToString("dd/MM/yyyy")
                }
                $fieldValues[$dispName] = "$rawVal"
            }
        }
        $reportItems.Add([PSCustomObject]@{
            id_sp   = "$($rep.ID_SP)"
            status  = "$($rep.Status)"
            linha   = $rep.Linha
            fields  = $fieldValues
        })
    }

    $reportPayload = [PSCustomObject]@{
        submission_datetime = $submissionDate
        id_mobilizacao      = "$ID_Mobilizacao"
        requester_name      = $requesterName
        requester_email     = $requesterEmail
        items               = @($reportItems)
    }

    $jsonStr = $reportPayload | ConvertTo-Json -Depth 5 -Compress
    Write-Output "---REPORT_JSON_START---"
    Write-Output $jsonStr
    Write-Output "---REPORT_JSON_END---"
}
catch {
    Write-Warning "Não foi possível gerar dados do relatório: $_"
}

$hasUploadErrors = $ExecutionReport | Where-Object { "$($_.Status)" -like "Erro:*" } | Select-Object -First 1
if ($hasUploadErrors) {
    exit 1
}

exit 0
