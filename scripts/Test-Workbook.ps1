param(
    [string]$Workbook,
    [string]$DemoOutput
)
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path $PSScriptRoot -Parent
if (-not $Workbook) {
    $taskCandidates = @(Get-ChildItem -LiteralPath $taskRoot -Filter '*.xlsm')
    if ($taskCandidates.Count -ne 1) { throw 'Specify exactly one workbook with -Workbook.' }
    $Workbook = $taskCandidates[0].FullName
}
$taskSource = (Resolve-Path -LiteralPath $Workbook).Path
if ($DemoOutput) {
    $DemoOutput = [IO.Path]::GetFullPath($DemoOutput)
    if ([IO.Path]::GetExtension($DemoOutput) -ne '.xlsm') { throw 'DemoOutput must end in .xlsm.' }
    if (Test-Path -LiteralPath $DemoOutput) { throw 'DemoOutput already exists; choose a new path.' }
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('vba-smoke-' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $taskTemp)
$taskCopy = Join-Path $taskTemp ([IO.Path]::GetFileName($taskSource))
Copy-Item -LiteralPath $taskSource -Destination $taskCopy
$taskExcel = $null
$taskBook = $null
try {
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    # This setting belongs only to the new disposable Excel instance.
    $taskExcel.AutomationSecurity = 3
    $taskBook = $taskExcel.Workbooks.Open($taskCopy,0,$false)
    $taskTests = $taskBook.VBProject.VBComponents.Import((Join-Path $taskRoot 'tests/vba/SmokeTests.bas'))
    $taskBook.Save()
    $taskBook.Close($false)
    $taskBook = $null
    $taskExcel.AutomationSecurity = 1
    $taskBook = $taskExcel.Workbooks.Open($taskCopy,0,$false)
    $taskResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunSmokeTests")
    if (-not $taskResult.StartsWith('PASS:')) { throw $taskResult }
    Write-Output $taskResult
    if ($DemoOutput) {
        $null = $taskExcel.Run("'$($taskBook.Name)'!BuildDemo")
        $taskBook.VBProject.VBComponents.Remove($taskBook.VBProject.VBComponents.Item('SmokeTests'))
        $taskExcel.CalculateFull()
        $taskBook.SaveAs($DemoOutput,52)
        Write-Output ('Synthetic demo workbook: ' + $DemoOutput)
    }
} finally {
    if ($null -ne $taskBook) { $taskBook.Close($false) }
    if ($null -ne $taskExcel) {
        $taskExcel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel)
    }
    if (Test-Path -LiteralPath $taskCopy) { Remove-Item -LiteralPath $taskCopy -Force }
    if (@(Get-ChildItem -LiteralPath $taskTemp -Force).Count -eq 0) { Remove-Item -LiteralPath $taskTemp }
}
