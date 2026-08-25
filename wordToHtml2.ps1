<#
.SYNOPSIS
    Batch-converts Microsoft Word (.docx) files into clean, structural HTML.
.DESCRIPTION
    Uses native Windows 11 PowerShell and Word COM automation with zero external dependencies.
#>

param (
    [Parameter(Mandatory=$true, HelpMessage="Path to a single .docx file or a folder containing .docx files")]
    [string]$Path,

    [Parameter(Mandatory=$false, HelpMessage="Optional custom output directory")]
    [string]$OutputPath
)

# Enums & constants for Word COM automation
$wdFormatFilteredHTML = 10
$wdDoNotSaveChanges   = 0
$missing              = [System.Reflection.Missing]::Value

# --- 1. Input Path Resolution (LiteralPath for wildcard safety) ---
$targetItem = Get-Item -LiteralPath $Path -ErrorAction Stop
$isFolder   = $targetItem -is [System.IO.DirectoryInfo]

$filesToProcess = @()
if ($isFolder) {
    # Get .docx files, excluding temporary lock files (~$*)
    $filesToProcess = Get-ChildItem -LiteralPath $targetItem.FullName -Filter "*.docx" -File | 
                      Where-Object { $_.Name -notlike "~$*" }
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $targetItem.FullName -ChildPath "Clean_HTML"
    }
} else {
    if ($targetItem.Extension -notmatch "^\.docx$") {
        Write-Host "Error: The provided file is not a .docx file." -ForegroundColor Red
        exit
    }
    
    if ($targetItem.Name -like "~$*") {
        Write-Host "Error: Cannot process Word temporary lock files." -ForegroundColor Red
        exit
    }
    
    $filesToProcess = @($targetItem)
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = $targetItem.DirectoryName
    }
}

if ($filesToProcess.Count -eq 0) {
    Write-Host "No valid .docx files found to process." -ForegroundColor Yellow
    exit
}

# Ensure Output Directory exists
if (-not (Test-Path -LiteralPath $OutputPath)) {
    Write-Host "Creating output directory: $OutputPath" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}
$AbsoluteOutPath = (Get-Item -LiteralPath $OutputPath).FullName

Write-Host "Starting conversion for $($filesToProcess.Count) file(s)..." -ForegroundColor Cyan
Write-Host "Output location: $AbsoluteOutPath" -ForegroundColor Cyan
Write-Host "---------------------------------------------------" -ForegroundColor DarkGray

# Encoding setup: UTF-8 without BOM using native .NET
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

# Initialize COM Object Variables
$wordApp = $null
$processedCount = 0

try {
    # --- 2. Initialize Word COM Application ---
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0 # wdAlertsNone

    # --- 3. Processing Loop ---
    foreach ($file in $filesToProcess) {
        $processedCount++
        $prefix = "[$processedCount/$($filesToProcess.Count)]"
        Write-Host "$prefix Processing: $($file.Name)... " -NoNewline -ForegroundColor White

        $htmlFileName = $file.BaseName + ".html"
        $htmlFilePath = Join-Path -Path $AbsoluteOutPath -ChildPath $htmlFileName
        
        $document = $null

        try {
            # Open document in Read-Only mode
            $document = $wordApp.Documents.Open([ref]$file.FullName, [ref]$missing, [ref]$true)
            
            # Save as Filtered HTML
            $document.SaveAs([ref]$htmlFilePath, [ref]$wdFormatFilteredHTML)
            
            # Close document immediately
            $document.Close([ref]$wdDoNotSaveChanges)
            $document = $null

            # --- 4. Regex Cleaning Pipeline ---
            $htmlContent = [System.IO.File]::ReadAllText($htmlFilePath, [System.Text.Encoding]::UTF8)

            # 1. Remove XML namespaces and XML declarations
            $htmlContent = $htmlContent -replace '<\?xml.*?\?>', ''
            
            # 2. Remove all <style> blocks
            $htmlContent = $htmlContent -replace '(?si)<style[^>]*>.*?</style>', ''
            
            # 3. Remove MSO conditional comments
            $htmlContent = $htmlContent -replace '(?si)<!--\[if.*?<!\[endif\]-->', ''
            
            # 4. Remove proprietary <o:p> wrappers
            $htmlContent = $htmlContent -replace '(?i)</?o:p[^>]*>', ''
            
            # 5. Remove inline attributes: style, class, lang
            $htmlContent = $htmlContent -replace '(?i)\s+(class|style|lang)="[^"]*"', ''
            $htmlContent = $htmlContent -replace "(?i)\s+(class|style|lang)='[^']*'", ''
            
            # 6. Clean up trailing spaces in tags (e.g. <p > to <p>)
            $htmlContent = $htmlContent -replace '\s+>', '>'
            
            # 7. Strip non-semantic wrappers (span, font)
            $htmlContent = $htmlContent -replace '(?i)</?(span|font)[^>]*>', ''
            
            # 8. Strip meta and link tags
            $htmlContent = $htmlContent -replace '(?si)<meta[^>]*>', ''
            $htmlContent = $htmlContent -replace '(?si)<link[^>]*>', ''

            # 9. Dynamic multi-pass empty-tag cleanup
            $emptyTagPattern = '(?i)<(p|b|i|strong|em)>(\s|&nbsp;)*</\1>'
            do {
                $prevLength = $htmlContent.Length
                $htmlContent = $htmlContent -replace $emptyTagPattern, ''
            } while ($htmlContent.Length -ne $prevLength)

            # 10. Normalize whitespace and blank lines
            $htmlContent = $htmlContent -replace '(?m)^\s+$', ''
            $htmlContent = $htmlContent -replace '[\r\n]{2,}', "`r`n"

            # Write clean HTML back to disk (UTF-8 without BOM)
            [System.IO.File]::WriteAllText($htmlFilePath, $htmlContent, $utf8NoBom)

            Write-Host "Success!" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed! ($($_.Exception.Message))" -ForegroundColor Red
            
            if ($null -ne $document) {
                $document.Close([ref]$wdDoNotSaveChanges)
                $document = $null
            }
        }
    }
}
finally {
    # --- 5. COM Cleanup ---
    Write-Host "---------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Releasing Word COM objects from system memory..." -ForegroundColor Cyan

    if ($null -ne $wordApp) {
        $wordApp.Quit([ref]$wdDoNotSaveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        $wordApp = $null
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Done." -ForegroundColor Green
}
