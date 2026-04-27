<#
.SYNOPSIS
    Aplica regras de atribuição automática de Analista Responsável aos itens de mobilização.

.DESCRIPTION
    Implementa 4 regras de prioridade para preenchimento da coluna "Analista Responsável":
    
    Regra 1: FORNECEDOR em lista específica -> "MAFDO"
    Regra 2: TIPO DE MOBILIZAÇÃO = "Máquinas e Equipamentos" -> "SAIOI / LPHDS"
    Regra 3: TIPO DE ATIVIDADE inicia com "ST - " -> "SAIOI / LPHDS"
    Regra 4: Fallback - compara Parque com mapeamento do Excel -> Analista

.PARAMETER Items
    Array de PSCustomObjects com os dados do Excel (linhas a processar)

.PARAMETER ParqueLookupListId
    ID da lista de Parques no SharePoint (para Regra 4)

.PARAMETER SiteUrl
    URL do site SharePoint (para Regra 4)

.OUTPUTS
    Array de PSCustomObjects com a coluna "Analista Responsável" preenchida
#>

function Normalize-TextForCompare {
    <#
    .SYNOPSIS
        Normaliza texto removendo acentos para comparação case-insensitive
    #>
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    $normalized = $Text.Trim().ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder

    foreach ($ch in $normalized.ToCharArray()) {
        $unicodeCategory = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($unicodeCategory -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }

    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Get-FieldValue {
    <#
    .SYNOPSIS
        Extrai o valor de um campo de um objeto, tratando nulos e espaços
        Procura por nomes exatos e variações (case-insensitive, com/sem acentos)
    #>
    param(
        [PSCustomObject]$Object,
        [string]$FieldName,
        [string[]]$AlternateNames = @()
    )

    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($FieldName)) {
        return ""
    }

    $namesToCheck = @($FieldName) + $AlternateNames
    
    # Primeiro tenta correspondência exata (case-insensitive)
    foreach ($name in $namesToCheck) {
        $prop = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($prop) {
            $val = $prop.Value
            if ($null -ne $val) { return "$val".Trim() }
        }
    }

    # Se não encontrar exato, procura por substring (case-insensitive)
    foreach ($name in $namesToCheck) {
        $prop = $Object.PSObject.Properties | Where-Object {
            $_.Name -ilike "*$name*"
        } | Select-Object -First 1
        if ($prop) {
            $val = $prop.Value
            if ($null -ne $val) { return "$val".Trim() }
        }
    }

    return ""
}

function Test-AnalistaEmpty {
    <#
    .SYNOPSIS
        Verifica se a coluna Analista está vazia/nula
    #>
    param(
        [PSCustomObject]$Object
    )

    $val = Get-FieldValue -Object $Object -FieldName "Analista" -AlternateNames @("AnalistaResponsavel", "AnalistaRespons_x00e1_vel", "Analyst", "Responsible")
    return [string]::IsNullOrWhiteSpace($val)
}

function Set-AnalistaValue {
    <#
    .SYNOPSIS
        Define o valor da coluna Analista no objeto
    #>
    param(
        [PSCustomObject]$Object,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }

    $internalAnalistaName = "AnalistaRespons_x00e1_vel"

    # Procura a propriedade "Analista" com variações
    $propriedadeBuscada = @("Analista Responsável", "AnalistaResponsavel", "AnalistaRespons_x00e1_vel", "Analista", "Analyst", "Responsible Analyst")
    
    foreach ($propName in $propriedadeBuscada) {
        $prop = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $propName } | Select-Object -First 1
        if ($prop) {
            $prop.Value = $Value
            $internalProp = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $internalAnalistaName } | Select-Object -First 1
            if ($internalProp) {
                $internalProp.Value = $Value
            } else {
                $Object | Add-Member -MemberType NoteProperty -Name $internalAnalistaName -Value $Value -Force
            }
            return
        }
    }

    # Se não encontrar, tenta com "like" (substring)
    $prop = $Object.PSObject.Properties | Where-Object {
        $_.Name -ilike "*Analista*" -and $_.Name -ilike "*Responsavel*"
    } | Select-Object -First 1

    if ($prop) {
        $prop.Value = $Value
        $internalProp = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $internalAnalistaName } | Select-Object -First 1
        if ($internalProp) {
            $internalProp.Value = $Value
        } else {
            $Object | Add-Member -MemberType NoteProperty -Name $internalAnalistaName -Value $Value -Force
        }
        return
    }

    # Se ainda não encontrar, cria uma nova propriedade com o nome mais comum
    $Object | Add-Member -MemberType NoteProperty -Name "Analista Responsável" -Value $Value -Force
    $Object | Add-Member -MemberType NoteProperty -Name $internalAnalistaName -Value $Value -Force
}

function Get-FirstNonEmpty {
    <#
    .SYNOPSIS
        Retorna o primeiro valor não vazio de uma lista de candidatos
    #>
    param([object[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if ($null -eq $candidate) { continue }
        $asString = "$candidate"
        if (-not [string]::IsNullOrWhiteSpace($asString)) {
            return $asString
        }
    }

    return $null
}

function Get-ItemLineLabel {
    <#
    .SYNOPSIS
        Resolve um identificador amigável da linha/item para logs
    #>
    param([PSCustomObject]$Item)

    $idRaw = Get-FieldValue -Object $Item -FieldName "ID" -AlternateNames @("Id", "LinhaExcel", "Linha", "Row", "RowNumber")
    if ([string]::IsNullOrWhiteSpace($idRaw)) {
        return "N/A"
    }

    return $idRaw
}

function Get-FieldText {
    <#
    .SYNOPSIS
        Extrai texto de campos heterogêneos (string, lookup, objetos)
    #>
    param($Value)

    if ($null -eq $Value) { return "" }

    if ($Value -is [string]) {
        if ($Value.Contains(";#")) {
            $parts = $Value -split ";#"
            return $parts[$parts.Length - 1].Trim()
        }
        return $Value.Trim()
    }

    if ($Value.PSObject -and $Value.PSObject.Properties["LookupValue"]) {
        return [string]$Value.LookupValue
    }

    if ($Value.PSObject -and $Value.PSObject.Properties["Label"]) {
        return [string]$Value.Label
    }

    if ($Value.PSObject -and $Value.PSObject.Properties["Value"]) {
        return [string]$Value.Value
    }

    return "$Value".Trim()
}

function Load-CleverParqueAnalistaMap {
    <#
    .SYNOPSIS
        Carrega o mapeamento Parque -> Analista a partir do Excel online (Clever)
    #>
    param(
        [string]$WorkbookUniqueId,
        [string]$WorksheetOrTableName,
        [string]$ColParquePreferida,
        [string]$ColAnalistaPreferida,
        [switch]$DetailedLog
    )

    $mapParqueAnalista = @{}

    try {
        $fileInfo = Invoke-PnPSPRestMethod `
            -Method GET `
            -Url "/_api/web/GetFileById('$WorkbookUniqueId')?`$select=ServerRelativeUrl" `
            -ErrorAction Stop
        $excelServerRelativeUrl = $fileInfo.ServerRelativeUrl
        if ($DetailedLog) { Write-Host "  Arquivo Clever encontrado: $excelServerRelativeUrl" -ForegroundColor Gray }
    } catch {
        if ($DetailedLog) { Write-Host "  Aviso: não foi possível localizar o arquivo Clever por UniqueId." -ForegroundColor Yellow }
        return $mapParqueAnalista
    }

    $tmpFolder = Join-Path $env:TEMP "SolicitadorMobilizacao"
    if (-not (Test-Path $tmpFolder)) {
        New-Item -Path $tmpFolder -ItemType Directory | Out-Null
    }

    $localExcelPath = Join-Path $tmpFolder "db_sites_clever.xlsx"
    try {
        Get-PnPFile -Url $excelServerRelativeUrl -Path $tmpFolder -FileName "db_sites_clever.xlsx" -AsFile -Force -ErrorAction Stop
    } catch {
        if ($DetailedLog) { Write-Host "  Aviso: não foi possível baixar o arquivo Clever." -ForegroundColor Yellow }
        return $mapParqueAnalista
    }

    $excelRows = @()
    $loadedByImportExcel = $false

    if (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue) {
        try {
            $excelRows = @(Import-Excel -Path $localExcelPath -WorksheetName $WorksheetOrTableName -ErrorAction Stop)
            $loadedByImportExcel = $true
        } catch {
            try {
                $excelRows = @(Import-Excel -Path $localExcelPath -ErrorAction Stop)
                $loadedByImportExcel = $true
            } catch {
                $loadedByImportExcel = $false
            }
        }
    }

    if (-not $loadedByImportExcel) {
        # Fallback COM (Excel instalado), para não depender do módulo ImportExcel.
        $excel = $null
        $workbook = $null
        $worksheet = $null
        $usedRange = $null
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($localExcelPath)
            try {
                $worksheet = $workbook.Worksheets.Item($WorksheetOrTableName)
            } catch {
                $worksheet = $workbook.Worksheets.Item(1)
            }

            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count
            if ($rowCount -lt 2) {
                return $mapParqueAnalista
            }

            $valueArray = $usedRange.Value2
            $headers = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $headers += "$($valueArray[1, $c])"
            }

            for ($r = 2; $r -le $rowCount; $r++) {
                $obj = [PSCustomObject]@{}
                for ($c = 1; $c -le $colCount; $c++) {
                    $header = $headers[$c - 1]
                    if ([string]::IsNullOrWhiteSpace($header)) { continue }
                    $cell = $valueArray[$r, $c]
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value $cell -Force
                }
                $excelRows += $obj
            }
        } catch {
            if ($DetailedLog) { Write-Host "  Aviso: não foi possível ler Excel Clever (ImportExcel/COM)." -ForegroundColor Yellow }
            return $mapParqueAnalista
        } finally {
            try { if ($usedRange) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) } } catch {}
            try { if ($worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } } catch {}
            try { if ($workbook) { $workbook.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } } catch {}
            try { if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }

    if (-not $excelRows -or $excelRows.Count -eq 0) {
        return $mapParqueAnalista
    }

    $headers = $excelRows[0].PSObject.Properties.Name
    $headerParque = $null
    $headerAnalista = $null

    foreach ($h in $headers) {
        if ($h -ieq $ColParquePreferida) { $headerParque = $h; break }
    }
    if (-not $headerParque) {
        foreach ($h in $headers) {
            $hn = Normalize-TextForCompare -Text $h
            if ($hn.Contains("CLEVER") -and ($hn.Contains("PAQUE") -or $hn.Contains("PARQUE") -or $hn.Contains("PAQ"))) {
                $headerParque = $h
                break
            }
        }
    }

    foreach ($h in $headers) {
        if ($h -ieq $ColAnalistaPreferida) { $headerAnalista = $h; break }
    }
    if (-not $headerAnalista) {
        foreach ($h in $headers) {
            $hn = Normalize-TextForCompare -Text $h
            if ($hn.Contains("ANALISTA")) {
                $headerAnalista = $h
                break
            }
        }
    }

    if (-not $headerParque -or -not $headerAnalista) {
        if ($DetailedLog) { Write-Host "  Aviso: colunas Parque/Analista não encontradas no Excel Clever." -ForegroundColor Yellow }
        return $mapParqueAnalista
    }

    foreach ($row in $excelRows) {
        $parqueRaw = Get-FieldText -Value $row.$headerParque
        $analistaRaw = Get-FieldText -Value $row.$headerAnalista

        $parqueKey = Normalize-TextForCompare -Text $parqueRaw
        $analistaVal = "$analistaRaw".Trim()

        if ([string]::IsNullOrWhiteSpace($parqueKey)) { continue }
        if ([string]::IsNullOrWhiteSpace($analistaVal)) { continue }

        if (-not $mapParqueAnalista.ContainsKey($parqueKey)) {
            $mapParqueAnalista[$parqueKey] = $analistaVal
        }
    }

    return $mapParqueAnalista
}

function Apply-AnalistRules {
    <#
    .SYNOPSIS
        Aplica todas as regras de atribuição de Analista
    #>
    param(
        [Parameter(Mandatory=$true)]
        [Array]$Items,

        [Parameter(Mandatory=$false)]
        [string]$ParqueLookupListId = "678f10f9-8d46-404b-a451-70dfe938a1ee",

        [Parameter(Mandatory=$false)]
        [string]$SiteUrl = "https://vestas.sharepoint.com/sites/CC-ControleService-BR",

        [Parameter(Mandatory=$false)]
        [switch]$DetailedLog
    )

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "Nenhum item para processar." -ForegroundColor Gray
        return $Items
    }

    # Configuração das regras
    $FornecedoresRegra1 = @("FLEXWIND", "ARTHWIND", "TETRACE", "REVTECH", "TECH SERVICES", "AERIS", "PRIME WIND", "BELA VISTA", "DRONE BASE")
    $ValorRegra1 = "MAFDO"
    $TipoMobilizacaoRegra2 = "Maquinas e Equipamentos"
    $ValorRegra2 = "SAIOI / LPHDS"
    $PrefixoRegra3 = "ST - "
    $ValorRegra3 = "SAIOI / LPHDS"

    # Configuração da Regra 4 (Excel online Clever)
    $CleverWorkbookUniqueId = "CA5D2574-187D-444A-9C15-D5E966223565"
    $CleverWorksheetOrTableName = "tbSitesClever"
    $ColCleverParquePreferida = "MRP Parque"
    $ColCleverAnalistaPreferida = "Analista"

    # Estatísticas
    $stats = @{
        Regra1 = 0
        Regra2 = 0
        Regra3 = 0
        Regra4 = 0
        Vazio  = 0
    }

    # === REGRA 1: FORNECEDOR ===
    if ($DetailedLog) { Write-Host "Aplicando Regra 1: FORNECEDOR..." -ForegroundColor Gray }
    foreach ($item in $Items) {
        if (-not (Test-AnalistaEmpty -Object $item)) { continue }

        # Procura por "Fornecedor", "FORNECEDOR", ou nomes similares
        $fornecedor = Get-FieldValue -Object $item -FieldName "Fornecedor" -AlternateNames @("FORNECEDOR", "Supplier")
        if ([string]::IsNullOrWhiteSpace($fornecedor)) { continue }

        $fornStr = $fornecedor.Trim().ToUpper()
        if ($FornecedoresRegra1 -contains $fornStr) {
            Set-AnalistaValue -Object $item -Value $ValorRegra1
            $stats.Regra1++
            if ($DetailedLog) {
                $lineLabel = Get-ItemLineLabel -Item $item
                Write-Host "  [R1] ID/Linha: $lineLabel -> Fornecedor: $fornecedor" -ForegroundColor Green
            }
        }
    }

    # === REGRA 2: TIPO DE MOBILIZAÇÃO ===
    if ($DetailedLog) { Write-Host "Aplicando Regra 2: TIPO DE MOBILIZAÇÃO..." -ForegroundColor Gray }
    $tipoMobNorm = Normalize-TextForCompare -Text $TipoMobilizacaoRegra2
    foreach ($item in $Items) {
        if (-not (Test-AnalistaEmpty -Object $item)) { continue }

        # Procura por variações de "Tipo de Mobilização", "TipoMobilizacao", etc
        $tipo = Get-FieldValue -Object $item -FieldName "Tipo" -AlternateNames @("TipoMobilizacao", "Mobilizacao", "TipoMobiliza", "Mobilization Type")
        if ([string]::IsNullOrWhiteSpace($tipo)) { continue }

        $tipoNorm = Normalize-TextForCompare -Text $tipo
        if ($tipoNorm -eq $tipoMobNorm) {
            Set-AnalistaValue -Object $item -Value $ValorRegra2
            $stats.Regra2++
            if ($DetailedLog) {
                $lineLabel = Get-ItemLineLabel -Item $item
                Write-Host "  [R2] ID/Linha: $lineLabel -> Tipo: $tipo" -ForegroundColor Green
            }
        }
    }

    # === REGRA 3: TIPO DE ATIVIDADE (inicia com "ST - ") ===
    if ($DetailedLog) { Write-Host "Aplicando Regra 3: TIPO DE ATIVIDADE..." -ForegroundColor Gray }
    foreach ($item in $Items) {
        if (-not (Test-AnalistaEmpty -Object $item)) { continue }

        # Procura por "Atividade", "TIPODEATIVIDADE", "Activity", etc
        $atividade = Get-FieldValue -Object $item -FieldName "Atividade" -AlternateNames @("TIPODEATIVIDADE", "TipoAtividade", "Activity", "Type")
        if ([string]::IsNullOrWhiteSpace($atividade)) { continue }

        $atividadeStr = $atividade.Trim()
        if ($atividadeStr.StartsWith($PrefixoRegra3, [System.StringComparison]::OrdinalIgnoreCase)) {
            Set-AnalistaValue -Object $item -Value $ValorRegra3
            $stats.Regra3++
            if ($DetailedLog) {
                $lineLabel = Get-ItemLineLabel -Item $item
                Write-Host "  [R3] ID/Linha: $lineLabel -> Atividade: $atividade" -ForegroundColor Green
            }
        }
    }

    # === REGRA 4: PARQUE -> Excel Mapping (fallback) ===
    if ($DetailedLog) { Write-Host "Aplicando Regra 4: PARQUE (fallback)..." -ForegroundColor Gray }
    
    # Carregar mapeamento de Parque -> Analista do Excel online (Clever)
    $parqueAnalistaMap = @{}
    try {
        if ($DetailedLog) { Write-Host "  Carregando mapeamento de Parques..." -ForegroundColor Gray }

        $parqueAnalistaMap = Load-CleverParqueAnalistaMap `
            -WorkbookUniqueId $CleverWorkbookUniqueId `
            -WorksheetOrTableName $CleverWorksheetOrTableName `
            -ColParquePreferida $ColCleverParquePreferida `
            -ColAnalistaPreferida $ColCleverAnalistaPreferida `
            -DetailedLog:$DetailedLog

        if ($DetailedLog) { Write-Host "  Mapeamento carregado: $($parqueAnalistaMap.Count) parques" -ForegroundColor Gray }
    } catch {
        if ($DetailedLog) { Write-Host "  Aviso: Não foi possível carregar mapeamento do Excel Clever ($($_.Exception.Message))" -ForegroundColor Yellow }
    }

    # Aplicar Regra 4 para itens com Parque
    foreach ($item in $Items) {
        if (-not (Test-AnalistaEmpty -Object $item)) { continue }

        # Procura por "Parque" com variações
        $parque = Get-FieldValue -Object $item -FieldName "Parque" -AlternateNames @("Park", "WindPark", "LocationPark")
        if ([string]::IsNullOrWhiteSpace($parque)) { continue }

        $parqueKey = Normalize-TextForCompare -Text $parque
        if ($parqueAnalistaMap.ContainsKey($parqueKey)) {
            $analistaValor = $parqueAnalistaMap[$parqueKey]
            Set-AnalistaValue -Object $item -Value $analistaValor
            $stats.Regra4++
            if ($DetailedLog) {
                $lineLabel = Get-ItemLineLabel -Item $item
                Write-Host "  [R4] ID/Linha: $lineLabel -> Parque: $parque = $analistaValor" -ForegroundColor Green
            }
        }
    }

    # Contar quantos ainda estão vazios
    foreach ($item in $Items) {
        if (Test-AnalistaEmpty -Object $item) {
            $stats.Vazio++
        }
    }

    # Resumo
    if ($DetailedLog -or $true) {
        Write-Host ""
        Write-Host "Resumo de Preenchimento do Analista Responsável:" -ForegroundColor Cyan
        Write-Host "  Regra 1 (Fornecedor)          : $($stats.Regra1)"
        Write-Host "  Regra 2 (Tipo de Mobilização): $($stats.Regra2)"
        Write-Host "  Regra 3 (Tipo de Atividade)  : $($stats.Regra3)"
        Write-Host "  Regra 4 (Parque)             : $($stats.Regra4)"
        Write-Host "  Ainda Vazios                 : $($stats.Vazio)"
        Write-Host "  Total Processado             : $($Items.Count)"
        Write-Host ""
    }

    return $Items
}

if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    Export-ModuleMember -Function Apply-AnalistRules, Test-AnalistaEmpty, Set-AnalistaValue, Get-FieldValue, Normalize-TextForCompare
}
