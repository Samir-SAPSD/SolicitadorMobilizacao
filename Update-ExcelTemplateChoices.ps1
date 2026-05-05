param(
    [Parameter(Mandatory=$true)]
    [string]$TemplatePath
)

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

$TestMode = $false
$SiteUrl = "https://vestas.sharepoint.com/sites/CC-ControleService-BR"

if ($TestMode) {
    $ListId = "ea1e6a2e-8df6-4171-825e-1b7ecfbea7a0"
} else {
    $ListId = "2d72b0f5-d3a3-4add-a8b0-3f94de786223"
}

if (-not (Test-Path $TemplatePath)) {
    Write-Error "Template não encontrado: $TemplatePath"
    exit 1
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$TargetPnPVersion = $null
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $TargetPnPVersion = "1.12.0"
}

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

    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    $InstallArgs = @{
        Name = "PnP.PowerShell"
        Scope = "CurrentUser"
        Force = $true
        AllowClobber = $true
        ErrorAction = "Stop"
    }
    if ($TargetPnPVersion) {
        $InstallArgs["RequiredVersion"] = $TargetPnPVersion
    }

    Install-Module @InstallArgs
}

if ($TargetPnPVersion) {
    Import-Module PnP.PowerShell -RequiredVersion $TargetPnPVersion -ErrorAction Stop
} else {
    Import-Module PnP.PowerShell -ErrorAction Stop
}

Write-Host "Conectando ao SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue

Write-Host "Consultando campos Choice da lista..." -ForegroundColor Cyan
$allFields = Get-PnPField -List $ListId | Select-Object Title, InternalName, TypeAsString, Choices, Hidden, ReadOnlyField, Sealed

$choiceFields = $allFields | Where-Object {
    ($_.TypeAsString -eq "Choice" -or $_.TypeAsString -eq "MultiChoice") -and
    -not $_.Hidden -and
    -not $_.ReadOnlyField -and
    -not $_.Sealed
} | Select-Object Title, InternalName, TypeAsString, Choices

if (-not $choiceFields -or $choiceFields.Count -eq 0) {
    Write-Error "Nenhuma coluna Choice/MultiChoice encontrada na lista."
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}

$choiceByTitle = @{}
$choiceByInternalName = @{}
$fieldByTitle = @{}
$fieldByInternalName = @{}

foreach ($field in $allFields) {
    $titleKey = "$($field.Title)".Trim().ToLowerInvariant()
    $internalKey = "$($field.InternalName)".Trim().ToLowerInvariant()

    if (-not [string]::IsNullOrWhiteSpace($titleKey)) {
        $fieldByTitle[$titleKey] = $field
    }
    if (-not [string]::IsNullOrWhiteSpace($internalKey)) {
        $fieldByInternalName[$internalKey] = $field
    }
}

foreach ($field in $choiceFields) {
    $titleKey = "$($field.Title)".Trim().ToLowerInvariant()
    $internalKey = "$($field.InternalName)".Trim().ToLowerInvariant()

    if (-not [string]::IsNullOrWhiteSpace($titleKey)) {
        $choiceByTitle[$titleKey] = $field
    }
    if (-not [string]::IsNullOrWhiteSpace($internalKey)) {
        $choiceByInternalName[$internalKey] = $field
    }
}

$excel = $null
$workbook = $null
$worksheet = $null

try {
    Write-Host "Abrindo template Excel..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $resolvedTemplate = (Resolve-Path $TemplatePath).Path
    $workbook = $excel.Workbooks.Open($resolvedTemplate)

    try {
        $worksheet = $workbook.Worksheets.Item("Lista Suspensa")
    } catch {
        throw "A aba 'Lista Suspensa' não foi encontrada no template."
    }

    # Atualiza cabeçalhos das abas de entrada para exibir DisplayName sem perder referência ao InternalName
    $headerSheets = @("PESSOAS", "EQUIPAMENTOS")
    foreach ($sheetName in $headerSheets) {
        try {
            $wsHeader = $workbook.Worksheets.Item($sheetName)
        } catch {
            Write-Warning "A aba '$sheetName' não foi encontrada no template."
            continue
        }

        $firstHeaderRaw = $wsHeader.Cells.Item(1, 1).Value2
        $firstHeader = "$firstHeaderRaw".Trim()
        if ($firstHeader -ne "GRUPO") {
            $wsHeader.Columns.Item(1).Insert() | Out-Null
            $wsHeader.Cells.Item(1, 1).Value2 = "GRUPO"
            Write-Host "Aba '$sheetName': coluna GRUPO inserida na posição A." -ForegroundColor Green
        }

        $usedRangeHeader = $wsHeader.UsedRange
        $lastHeaderColumnSheet = [Math]::Max($usedRangeHeader.Columns.Count, 1)

        for ($col = 1; $col -le $lastHeaderColumnSheet; $col++) {
            $headerRaw = $wsHeader.Cells.Item(1, $col).Value2
            $header = "$headerRaw".Trim()
            if ([string]::IsNullOrWhiteSpace($header)) {
                continue
            }

            $match = $null
            $headerKey = $header.ToLowerInvariant()
            $internalHint = $null

            # Suporta cabeçalho já no formato "Display [Internal]"
            if ($header -match '^(?<disp>.+?)\s*\[(?<internal>[^\[\]]+)\]\s*$') {
                $internalKey = "$($matches['internal'])".Trim().ToLowerInvariant()
                $internalHint = $internalKey
                if ($fieldByInternalName.ContainsKey($internalKey)) {
                    $match = $fieldByInternalName[$internalKey]
                }
            }

            if (-not $match) {
                if ($fieldByInternalName.ContainsKey($headerKey)) {
                    $match = $fieldByInternalName[$headerKey]
                } elseif ($fieldByTitle.ContainsKey($headerKey)) {
                    $match = $fieldByTitle[$headerKey]
                }
            }

            if (-not $match -and -not [string]::IsNullOrWhiteSpace($internalHint) -and $fieldByInternalName.ContainsKey($internalHint)) {
                $match = $fieldByInternalName[$internalHint]
            }

            if ($match) {
                $newHeader = "$($match.Title) [$($match.InternalName)]"
                $wsHeader.Cells.Item(1, $col).Value2 = $newHeader
            }
        }
    }

    $usedRange = $worksheet.UsedRange
    $lastHeaderColumn = [Math]::Max($usedRange.Columns.Count, 1)

    $updatedColumns = 0
    $notMatchedColumns = @()

    for ($col = 1; $col -le $lastHeaderColumn; $col++) {
        $headerRaw = $worksheet.Cells.Item(1, $col).Value2
        $header = "$headerRaw".Trim()
        if ([string]::IsNullOrWhiteSpace($header)) {
            continue
        }

        $headerKey = $header.ToLowerInvariant()
        $field = $null
        $internalHint = $null

        if ($header -match '^(?<disp>.+?)\s*\[(?<internal>[^\[\]]+)\]\s*$') {
            $internalHint = "$($matches['internal'])".Trim().ToLowerInvariant()
        }

        if (-not [string]::IsNullOrWhiteSpace($internalHint) -and $choiceByInternalName.ContainsKey($internalHint)) {
            $field = $choiceByInternalName[$internalHint]
        } elseif ($choiceByTitle.ContainsKey($headerKey)) {
            $field = $choiceByTitle[$headerKey]
        } elseif ($choiceByInternalName.ContainsKey($headerKey)) {
            $field = $choiceByInternalName[$headerKey]
        }

        if (-not $field) {
            $notMatchedColumns += $header
            continue
        }

        $choices = @($field.Choices | Sort-Object)

        # Limpa as linhas antigas da coluna para evitar lixo de opções anteriores.
        $worksheet.Range($worksheet.Cells.Item(2, $col), $worksheet.Cells.Item(5000, $col)).ClearContents() | Out-Null

        if ($choices.Count -gt 0) {
            $row = 2
            foreach ($choice in $choices) {
                $worksheet.Cells.Item($row, $col).Value2 = "$choice"
                $row++
            }
            $updatedColumns++
            Write-Host "Atualizado: $header ($($choices.Count) opções)" -ForegroundColor Green
        } else {
            Write-Host "Sem opções para: $header" -ForegroundColor Yellow
        }
    }

    $workbook.Save()

    Write-Host "Resumo: $updatedColumns coluna(s) atualizada(s)." -ForegroundColor Cyan
    if ($notMatchedColumns.Count -gt 0) {
        Write-Host "Colunas sem correspondência em Choice: $($notMatchedColumns -join ', ')" -ForegroundColor Yellow
    }
}
finally {
    if ($workbook) {
        try { $workbook.Close($true) } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Write-Host "Template atualizado com sucesso." -ForegroundColor Green
exit 0
