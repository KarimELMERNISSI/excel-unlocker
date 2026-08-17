# ==============================================================================
# Excel Protection Remover - PowerShell Natif (100% Sans fichier .exe)
# Fonctionne directement sur Windows 10/11 sans aucune installation
# ==============================================================================

[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Files,

    [Alias("b")]
    [switch]$Backup
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Inspect-ExcelProtections {
    param([string]$FullPath)

    $items = @()
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FullPath)
        
        # 1. Protection classeur
        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/workbook.xml" }
        $wbContent = ""
        if ($wbEntry) {
            $stream = $wbEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $wbContent = $reader.ReadToEnd()
            $reader.Dispose()
            $stream.Dispose()

            $hasWbProt = ($wbContent -match "<workbookProtection[^>]*(\/>|>.*?<\/workbookProtection>)") -or
                         ($wbContent -match '\slockStructure=["'']?(?:true|1)["'']?') -or
                         ($wbContent -match '\slockWindows=["'']?(?:true|1)["'']?')

            $hasFileSharing = ($wbContent -match "<fileSharing[^>]*(\/>|>.*?<\/fileSharing>)")

            if ($hasWbProt) {
                $items += [PSCustomObject]@{
                    Id = "wb_structure"
                    Title = "Structure du Classeur (Feuilles & Fenêtres)"
                    TargetPath = "xl/workbook.xml"
                    IsProtected = $true
                }
            }
            if ($hasFileSharing) {
                $items += [PSCustomObject]@{
                    Id = "wb_sharing"
                    Title = "Partage / Mot de passe de modification"
                    TargetPath = "xl/workbook.xml"
                    IsProtected = $true
                }
            }
        }

        # 2. Relations des feuilles
        $relMap = @{}
        $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/_rels/workbook.xml.rels" }
        if ($relsEntry) {
            $rStream = $relsEntry.Open()
            [xml]$relsXml = (New-Object System.IO.StreamReader($rStream, [System.Text.Encoding]::UTF8)).ReadToEnd()
            $rStream.Dispose()
            foreach ($rel in $relsXml.Relationships.Relationship) {
                $target = $rel.Target.TrimStart('/')
                if (-not $target.StartsWith('xl/')) { $target = "xl/$target" }
                $relMap[$rel.Id] = $target
            }
        }

        # 3. Feuilles
        if ($wbContent) {
            [xml]$wbXml = $wbContent
            $sheetIdx = 1
            foreach ($sheet in $wbXml.workbook.sheets.sheet) {
                $sName = $sheet.name
                $rId = $sheet.id
                $targetPath = $relMap[$rId]
                if (-not $targetPath) { $targetPath = "xl/worksheets/sheet$sheetIdx.xml" }
                $sheetIdx++

                $sEntry = $zip.Entries | Where-Object { $_.FullName -eq $targetPath }
                $isProt = $false
                if ($sEntry) {
                    $sStream = $sEntry.Open()
                    $sText = (New-Object System.IO.StreamReader($sStream, [System.Text.Encoding]::UTF8)).ReadToEnd()
                    $sStream.Dispose()
                    $isProt = ($sText -match "<sheetProtection[^>]*(\/>|>.*?<\/sheetProtection>)") -or
                              ($sText -match "<protectedRanges[^>]*(\/>|>.*?<\/protectedRanges>)")
                }

                $items += [PSCustomObject]@{
                    Id = "sheet_$sheetIdx"
                    Title = "Feuille « $sName »"
                    TargetPath = $targetPath
                    IsProtected = $isProt
                }
            }
        }

        $zip.Dispose()
    }
    catch {
        Write-Warning "Erreur lors de l'inspection : $_"
    }

    return $items
}

function Unlock-ExcelFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string[]]$SelectedTargets = $null,
        [bool]$MakeBackup = $false
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item -or -not $item.Exists) {
        Write-Host "[ERREUR] Fichier introuvable : $Path" -ForegroundColor Red
        return
    }

    $fullPath = $item.FullName

    if ($MakeBackup) {
        $backupPath = "$fullPath.bak"
        Copy-Item -LiteralPath $fullPath -Destination $backupPath -Force
        Write-Host "[INFO] Sauvegarde créée : $(Split-Path $backupPath -Leaf)" -ForegroundColor Cyan
    }

    $tempFile = [System.IO.Path]::Combine($item.DirectoryName, ".~$($item.BaseName)_tmp$($item.Extension)")

    try {
        $zipIn = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
        $zipOut = [System.IO.Compression.ZipFile]::Open($tempFile, [System.IO.Compression.ZipArchiveMode]::Create)

        $modifiedEntries = 0

        foreach ($entry in $zipIn.Entries) {
            $newEntry = $zipOut.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $entry.Open()
            $newStream = $newEntry.Open()

            $shouldModify = $false
            if ($null -eq $SelectedTargets) {
                $shouldModify = $true
            } else {
                $shouldModify = ($SelectedTargets -contains $entry.FullName) -or 
                                ($entry.FullName -eq "xl/workbook.xml" -and ($SelectedTargets -contains "wb_structure" -or $SelectedTargets -contains "wb_sharing"))
            }

            if ($shouldModify -and (($entry.FullName -like "*.xml") -or ($entry.FullName -like "*.rels"))) {
                $reader = New-Object System.IO.StreamReader($entryStream, [System.Text.Encoding]::UTF8)
                $content = $reader.ReadToEnd()
                $reader.Dispose()
                $originalContent = $content

                # Classeur
                if ($entry.FullName -eq "xl/workbook.xml") {
                    if ($null -eq $SelectedTargets -or ($SelectedTargets -contains "wb_structure")) {
                        $content = [regex]::Replace($content, "<workbookProtection[^>]*(\/>|>.*?<\/workbookProtection>)", "", "IgnoreCase,Singleline")
                        $content = [regex]::Replace($content, '\s+lockStructure=["'']?(?:true|1)["'']?', "", "IgnoreCase")
                        $content = [regex]::Replace($content, '\s+lockWindows=["'']?(?:true|1)["'']?', "", "IgnoreCase")
                        $content = [regex]::Replace($content, '\s+lockRevision=["'']?(?:true|1)["'']?', "", "IgnoreCase")
                    }
                    if ($null -eq $SelectedTargets -or ($SelectedTargets -contains "wb_sharing")) {
                        $content = [regex]::Replace($content, "<fileSharing[^>]*(\/>|>.*?<\/fileSharing>)", "", "IgnoreCase,Singleline")
                    }
                }
                # Feuilles
                else {
                    $content = [regex]::Replace($content, "<sheetProtection[^>]*(\/>|>.*?<\/sheetProtection>)", "", "IgnoreCase,Singleline")
                    $content = [regex]::Replace($content, "<protectedRanges[^>]*(\/>|>.*?<\/protectedRanges>)", "", "IgnoreCase,Singleline")
                    $content = [regex]::Replace($content, "<dialogsheetProtection[^>]*(\/>|>.*?<\/dialogsheetProtection>)", "", "IgnoreCase,Singleline")
                }

                if ($content -ne $originalContent) { $modifiedEntries++ }

                $writer = New-Object System.IO.StreamWriter($newStream, (New-Object System.Text.UTF8Encoding($false)))
                $writer.Write($content)
                $writer.Flush()
                $writer.Dispose()
            }
            else {
                $entryStream.CopyTo($newStream)
            }

            $newStream.Dispose()
            $entryStream.Dispose()
        }

        $zipIn.Dispose()
        $zipOut.Dispose()

        Move-Item -LiteralPath $tempFile -Destination $fullPath -Force

        if ($modifiedEntries -gt 0) {
            Write-Host "[SUCCÈS] Déverrouillé avec succès : $($item.Name) ($modifiedEntries élément(s) modifié(s))" -ForegroundColor Green
        } else {
            Write-Host "[INFO] $($item.Name) traité (aucune modification requise)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[ERREUR] Sur $($item.Name) : $_" -ForegroundColor Red
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- Entrée principale ---
$targetFile = $null
if ($Files -and $Files.Count -gt 0) {
    $targetFile = $Files[0]
}
elseif (Test-Path "mon_fichier.xlsx") {
    $targetFile = "mon_fichier.xlsx"
}
else {
    $found = @(Get-ChildItem -Path . -Filter "*.xls*" | Where-Object { $_.Name -notlike ".~*" -and $_.Name -notlike "~$*" })
    if (Test-Path "data") {
        $found += @(Get-ChildItem -Path "data" -Filter "*.xls*" | Where-Object { $_.Name -notlike ".~*" -and $_.Name -notlike "~$*" })
    }
    if ($found.Count -gt 0) {
        Write-Host "Fichiers Excel détectés :" -ForegroundColor Cyan
        for ($i = 0; $i -lt $found.Count; $i++) {
            $folder = if ($found[$i].Directory.Name -ne "xlsx_unlocker") { " ($($found[$i].Directory.Name))" } else { "" }
            Write-Host "  [$($i+1)] $($found[$i].Name)$folder"
        }
        $choice = Read-Host "Entrez le numéro du fichier (ou 'q' pour quitter)"
        if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $found.Count) {
            $targetFile = $found[[int]$choice - 1].FullName
        }
    }
}

if ($targetFile -and (Test-Path $targetFile)) {
    Write-Host "`n--- Analyse des protections de $(Split-Path $targetFile -Leaf) ---" -ForegroundColor Cyan
    $elements = Inspect-ExcelProtections -FullPath (Resolve-Path $targetFile).Path

    if ($elements.Count -gt 0) {
        for ($i = 0; $i -lt $elements.Count; $i++) {
            $e = $elements[$i]
            $status = if ($e.IsProtected) { "[PROTÉGÉ]" } else { "[Libre]" }
            $color = if ($e.IsProtected) { "Red" } else { "Gray" }
            Write-Host "  [$($i+1)] $status $($e.Title)" -ForegroundColor $color
        }

        Write-Host "`nOptions :" -ForegroundColor Yellow
        Write-Host "  [Entrée]  : Déverrouiller TOUT par défaut"
        Write-Host "  [1,2,...] : Déverrouiller certains numéros ciblés (ex: 1,3)"
        
        $userChoice = Read-Host "Votre choix"
        if ($userChoice -eq "" -or $userChoice -eq "all") {
            Unlock-ExcelFile -Path $targetFile -SelectedTargets $null -MakeBackup $Backup
        }
        else {
            $selected = @()
            $nums = $userChoice -split ',' | ForEach-Object { $_.Trim() }
            foreach ($n in $nums) {
                if ($n -match '^\d+$' -and [int]$n -ge 1 -and [int]$n -le $elements.Count) {
                    $item = $elements[[int]$n - 1]
                    if ($item.Id -like "wb_*") { $selected += $item.Id }
                    else { $selected += $item.TargetPath }
                }
            }
            Unlock-ExcelFile -Path $targetFile -SelectedTargets $selected -MakeBackup $Backup
        }
    }
    else {
        Unlock-ExcelFile -Path $targetFile -SelectedTargets $null -MakeBackup $Backup
    }
}
