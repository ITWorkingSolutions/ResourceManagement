param(
    [string]$ProjectDir
)

# Normalize path (VS sometimes adds trailing slash)
$ProjectDir = $ProjectDir.Trim('"').TrimEnd('\')

# Paths relative to project root
$jsonPath   = Join-Path $ProjectDir "SQLite\Schema\Json\ExcelRuleViewMap.json"
$outputPath = Join-Path $ProjectDir "Generated\ViewMetadata.vb"

Write-Host "JSON path:   $jsonPath"
Write-Host "Output path: $outputPath"

# Ensure the output folder exists
$outputDir = Split-Path $outputPath
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}


# Load JSON
$json = Get-Content $jsonPath -Raw | ConvertFrom-Json

# Extract DisplayNames and remove spaces
$names = $json.Views |
    Where-Object { -not $_.IsConceptual } |
    ForEach-Object { $_.DisplayName.Replace(" ", "") }

# Join into a single string
$joined = ($names -join ", ")

# Generate VB code
$vb = @"
' Auto-generated. Do not edit.
Namespace GeneratedViewMetadata
    Public Module Views
        Public Const ValidViewNames As String = "$joined"
    End Module
End Namespace
"@

# Ensure output folder exists
$outputDir = Split-Path $outputPath
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Write file
Set-Content -Path $outputPath -Value $vb -Encoding UTF8