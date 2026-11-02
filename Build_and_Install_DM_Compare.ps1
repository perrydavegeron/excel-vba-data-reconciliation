# DM Compare Tool v3.1 - Excel VBA Add-in builder / installer (Windows PowerShell 5.1)
# Requirements: Desktop Microsoft Excel; Trust access to the VBA project object model.
# This script creates the add-in with YOUR local Excel installation; it does not download code.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Add-FormControl {
    param($Designer, [string]$ProgId, [string]$Name,
          [double]$Left, [double]$Top, [double]$Width, [double]$Height,
          [string]$Caption = '')
    $control = $Designer.Controls.Add($ProgId, $Name, $true)
    $control.Left = $Left
    $control.Top = $Top
    $control.Width = $Width
    $control.Height = $Height
    if ($ProgId -eq 'Forms.Label.1' -or $ProgId -eq 'Forms.CommandButton.1') {
        $control.Caption = $Caption
    }
    try { $control.Font.Name = 'Aptos'; $control.Font.Size = 10 } catch { }
    return $control
}

function Read-ZipText {
    param($zip, [string]$name)
    $part = $zip.GetEntry($name)
    if ($null -eq $part) { throw "OpenXML component missing: $name" }
    $reader = [System.IO.StreamReader]::new($part.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-ZipText {
    param($zip, [string]$name, [string]$text)
    $existing = $zip.GetEntry($name)
    if ($null -ne $existing) { $existing.Delete() }
    $part = $zip.CreateEntry($name)
    $stream = $part.Open()
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($text) } finally { $writer.Dispose() }
}

function Inject-Ribbon {
    param([string]$AddInFile)
    $ribbon = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="dmCompareTab" label="DM Compare">
        <group id="dmCompareGroup" label="File comparison">
          <button id="dmCompareButton" label="Compare Files" size="large"
            imageMso="FileOpen" onAction="DM_RibbonCompare"
            screentip="Compare Excel files"
            supertip="Compare File A and File B using a common Unique ID." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@
    $fileStream = [System.IO.File]::Open($AddInFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    $zip = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
    try {
        Write-ZipText $zip 'customUI/customUI14.xml' $ribbon

        [xml]$rels = Read-ZipText $zip '_rels/.rels'
        $relNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
        $rel = $rels.CreateElement('Relationship', $relNS)
        $rel.SetAttribute('Id', 'rIdDMCompare')
        $rel.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $rel.SetAttribute('Target', 'customUI/customUI14.xml')
        [void]$rels.DocumentElement.AppendChild($rel)
        Write-ZipText $zip '_rels/.rels' $rels.OuterXml

        [xml]$types = Read-ZipText $zip '[Content_Types].xml'
        $typeNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
        $override = $types.CreateElement('Override', $typeNS)
        $override.SetAttribute('PartName', '/customUI/customUI14.xml')
        $override.SetAttribute('ContentType', 'application/vnd.ms-office.customUI+xml')
        [void]$types.DocumentElement.AppendChild($override)
        Write-ZipText $zip '[Content_Types].xml' $types.OuterXml
    }
    finally { $zip.Dispose() }
}

$excel = $null
$book = $null
try {
    $source = Join-Path $PSScriptRoot 'modDMCompare.bas'
    $formSource = Join-Path $PSScriptRoot 'frmFileCompare_code.txt'
    if (!(Test-Path -LiteralPath $source) -or !(Test-Path -LiteralPath $formSource)) {
        throw 'Required .bas or UserForm source file is missing. Extract the complete ZIP first.'
    }
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        throw 'Close all Excel windows (save your work) before building/installing this add-in.'
    }
    $targetDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
    [void](New-Item -ItemType Directory -Path $targetDir -Force)
    $target = Join-Path $targetDir 'DM Compare Tool.xlam'
    Write-Host "Creating Excel add-in at: $target" -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    try {
        $project = $book.VBProject
        $null = $project.VBComponents.Count
    }
    catch {
        throw "Excel is blocking VBA project access. Enable 'Trust access to the VBA project object model' in File > Options > Trust Center > Trust Center Settings > Macro Settings, then retry. Details: $_"
    }
    $mod = $project.VBComponents.Import($source)
    $mod.Name = 'modDMCompare'
    $form = $project.VBComponents.Add(3) # vbext_ct_MSForm
    $form.Name = 'frmFileCompare'
    # IMPORTANT: design-time UserForm dimensions belong to the VBComponent
    # Properties collection, not to VBComponent.Designer. Designer.Width
    # produces a COM property-not-found error on desktop Excel.
    Write-Host 'Configuring UserForm layout...' -ForegroundColor Cyan
    $form.Properties.Item('Caption').Value = 'DM Compare Files v3.1'
    $form.Properties.Item('Width').Value = 520
    $form.Properties.Item('Height').Value = 350
    $d = $form.Designer

    # Two simple side-by-side source configuration panels.
    Add-FormControl $d 'Forms.Label.1' 'lblTitleA' 14 12 220 18 'FILE A (Original)' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblTitleB' 266 12 220 18 'FILE B (Updated)' | Out-Null
    Add-FormControl $d 'Forms.CommandButton.1' 'btnBrowseA' 14 34 100 27 'Browse A...' | Out-Null
    Add-FormControl $d 'Forms.CommandButton.1' 'btnBrowseB' 266 34 100 27 'Browse B...' | Out-Null
    $fileA = Add-FormControl $d 'Forms.TextBox.1' 'txtFileA' 14 67 233 24
    $fileB = Add-FormControl $d 'Forms.TextBox.1' 'txtFileB' 266 67 233 24
    $fileA.Locked = $true; $fileB.Locked = $true
    Add-FormControl $d 'Forms.Label.1' 'lblSheetA' 14 97 210 17 'Worksheet A' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblSheetB' 266 97 210 17 'Worksheet B' | Out-Null
    $comboA = Add-FormControl $d 'Forms.ComboBox.1' 'cboSheetA' 14 114 233 23
    $comboB = Add-FormControl $d 'Forms.ComboBox.1' 'cboSheetB' 266 114 233 23
    $comboA.Style = 2; $comboB.Style = 2
    Add-FormControl $d 'Forms.Label.1' 'lblHeaderA' 14 145 104 17 'Header row A' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblDataA' 139 145 108 17 'Data start row A' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblHeaderB' 266 145 104 17 'Header row B' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblDataB' 391 145 108 17 'Data start row B' | Out-Null
    Add-FormControl $d 'Forms.TextBox.1' 'txtHeaderA' 14 162 100 23 | Out-Null
    Add-FormControl $d 'Forms.TextBox.1' 'txtStartA' 139 162 108 23 | Out-Null
    Add-FormControl $d 'Forms.TextBox.1' 'txtHeaderB' 266 162 100 23 | Out-Null
    Add-FormControl $d 'Forms.TextBox.1' 'txtStartB' 391 162 108 23 | Out-Null

    Add-FormControl $d 'Forms.Label.1' 'lblID' 14 204 270 20 'Common Unique ID (choose one)' | Out-Null
    $comboID = Add-FormControl $d 'Forms.ComboBox.1' 'cboUniqueID' 14 226 337 24
    $comboID.Style = 2
    Add-FormControl $d 'Forms.CommandButton.1' 'btnRefreshIDs' 361 226 138 24 'Refresh headers' | Out-Null
    Add-FormControl $d 'Forms.Label.1' 'lblNote' 14 260 485 27 '' | Out-Null
    Add-FormControl $d 'Forms.CommandButton.1' 'btnCompare' 281 290 119 30 'Compare Files' | Out-Null
    Add-FormControl $d 'Forms.CommandButton.1' 'btnCancel' 409 290 90 30 'Cancel' | Out-Null
    Write-Host 'Adding UserForm event code...' -ForegroundColor Cyan
    $form.CodeModule.AddFromString([System.IO.File]::ReadAllText($formSource, [Text.Encoding]::UTF8))

    # xlOpenXMLAddIn = 55. This compiles real VBA through desktop Excel.
    Write-Host 'Saving the Excel add-in...' -ForegroundColor Cyan
    $book.SaveAs($target, 55)
    $book.Close($false)
    $book = $null
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null

    Write-Host 'Installing the Ribbon button...' -ForegroundColor Cyan
    Inject-Ribbon $target

    # Register the newly created add-in with Excel.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host 'Registering add-in with Excel...' -ForegroundColor Cyan
    # Excel's AddIns.Add can raise error 1004 when there are NO visible workbooks.
    # Create a normal workbook before invoking the Excel Add-In Manager.
    # Microsoft KB 280290 documents this exact behavior.
    $registrationBook = $excel.Workbooks.Add()
    try {
        # On an upgrade, re-use the existing entry rather than adding a duplicate.
        $item = $null
        for ($j = 1; $j -le $excel.AddIns.Count; $j++) {
            $candidate = $excel.AddIns.Item($j)
            if ([string]::Equals([string]$candidate.FullName, [string]$target,
                    [System.StringComparison]::OrdinalIgnoreCase)) {
                $item = $candidate
                break
            }
        }
        if ($null -eq $item) {
            $item = $excel.AddIns.Add($target, $false)
        }
        $item.Installed = $true
        Write-Host 'Add-in registered and enabled.' -ForegroundColor Green
    }
    finally {
        $registrationBook.Close($false)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    Write-Host ''
    Write-Host 'SUCCESS: DM Compare Tool installed.' -ForegroundColor Green
    Write-Host "Add-in: $target"
    Write-Host 'Open Excel and use the DM Compare tab > Compare Files.'
    Write-Host 'You can switch OFF VBA project object model access again after setup.'
}
catch {
    Write-Host ''
    Write-Host "SETUP FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'If the error occurred during registration, the .xlam may still have been created.' -ForegroundColor Yellow
    Write-Host 'Try Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse and select:' -ForegroundColor Yellow
    if ($target) { Write-Host "  $target" -ForegroundColor Yellow }
    Write-Host 'See README.txt for additional troubleshooting instructions.'
    exit 1
}
finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
