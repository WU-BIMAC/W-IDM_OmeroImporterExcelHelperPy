$script = Join-Path $PSScriptRoot "\fetch_images.py"
$excelFilePath = $args[0]
$isDebug = $args[1]
$result = $false

$activateScript = Join-Path $PSScriptRoot "\.venv\Scripts\Activate.ps1"
if (-not (Test-Path $activateScript)) {
   $activateScript = Join-Path $PSScriptRoot "\venv3\Scripts\Activate.ps1"
}
. $activateScript

if($excelFilePath -and $excelFilePath.Trim().Length -gt 0 -and (Test-Path $excelFilePath)) {
    python $script $excelFilePath
}
 
deactivate

if ($isDebug -and [bool]::TryParse($isDebug, [ref]$result) -and $result -eq $true) {
    Write-Host -NoNewLine 'Press any key to continue...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}