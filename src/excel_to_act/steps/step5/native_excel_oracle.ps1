param(
    [Parameter(Mandatory = $true)][string]$RequestPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$request = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json
$excel = $null
$book = $null

function Get-CellRecord($cell) {
    $text = [string]$cell.Text
    $formula = [string]$cell.Formula
    $value = $cell.Value2
    $errorCode = $null
    if ($text -match '^#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!)$') {
        $errorCode = $text
        $value = $null
    }
    return @{
        value = $value
        display = $text
        formula = $formula
        excel_error = $errorCode
    }
}

function Get-RangeRecord($workbook, $item) {
    $worksheet = $workbook.Worksheets.Item([string]$item.sheet)
    $range = $worksheet.Range([string]$item.address)
    $rows = [System.Collections.Generic.List[object]]::new()
    $formulas = [System.Collections.Generic.List[object]]::new()
    $errors = [System.Collections.Generic.List[object]]::new()
    for ($row = 1; $row -le $range.Rows.Count; $row++) {
        $valueRow = [System.Collections.Generic.List[object]]::new()
        $formulaRow = [System.Collections.Generic.List[object]]::new()
        $errorRow = [System.Collections.Generic.List[object]]::new()
        for ($column = 1; $column -le $range.Columns.Count; $column++) {
            $record = Get-CellRecord $range.Cells.Item($row, $column)
            $valueRow.Add($record.value)
            $formulaRow.Add($record.formula)
            $errorRow.Add($record.excel_error)
        }
        $rows.Add(@($valueRow.ToArray()))
        $formulas.Add(@($formulaRow.ToArray()))
        $errors.Add(@($errorRow.ToArray()))
    }
    return @{
        address = [string]$item.qualified_address
        values = @($rows.ToArray())
        formulas = @($formulas.ToArray())
        excel_errors = @($errors.ToArray())
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open([string]$request.copy_path, 0, $true)

    $primary = @{}
    $main = $book.Worksheets.Item('Main')
    foreach ($address in $request.primary_input_addresses) {
        $primary[[string]$address] = Get-CellRecord $main.Range([string]$address)
    }

    $beforeTargets = @{}
    foreach ($name in $request.targets) {
        $namedRange = $book.Names.Item([string]$name).RefersToRange
        $beforeTargets[[string]$name] = Get-CellRecord $namedRange.Cells.Item(1, 1)
    }
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.CalculateFullRebuild()
    $timer.Stop()

    $afterTargets = @{}
    foreach ($name in $request.targets) {
        $namedRange = $book.Names.Item([string]$name).RefersToRange
        $afterTargets[[string]$name] = Get-CellRecord $namedRange.Cells.Item(1, 1)
    }
    $afterRanges = @{}
    foreach ($item in $request.ranges) {
        $afterRanges[[string]$item.qualified_address] = Get-RangeRecord $book $item
    }

    $settings = @{
        iteration = [bool]$excel.Iteration
        max_iterations = [int]$excel.MaxIterations
        max_change = [double]$excel.MaxChange
        calculation = [int]$excel.Calculation
        version = [string]$excel.Version
    }
    $afterCells = @{}
    $standardNames = @{
        AnnuityDue = 'H1'
        PVLoading = 'H2'
        PVFB = 'H3'
        GP = 'J1'
    }
    foreach ($name in $request.targets) {
        if ($standardNames.ContainsKey([string]$name)) {
            $afterCells[$standardNames[[string]$name]] = $afterTargets[[string]$name]
        }
    }
    if ($afterRanges.ContainsKey('Premium!B9:CD115')) {
        $premium = $afterRanges['Premium!B9:CD115']
        $premiumValues = @()
        foreach ($row in $premium.values) {
            $premiumValues += ,@($row)
        }
        $premiumFormulas = @()
        foreach ($row in $premium.formulas) {
            $premiumFormulas += ,@($row)
        }
    } else {
        $premiumValues = @()
        $premiumFormulas = @()
    }
    $result = @{
        source_path = [string]$request.source_path
        source_sha256 = [string]$request.source_sha256
        source_copy_sha256 = [string]$request.source_sha256
        engine = 'Microsoft Excel'
        engine_settings = $settings
        full_rebuild_seconds = $timer.Elapsed.TotalSeconds
        calculation_state = [int]$excel.CalculationState
        macros_executed = $false
        input_overrides = @()
        primary_inputs = $primary
        before_explicit_rebuild = $beforeTargets
        after_full_rebuild = $afterCells
        named_targets = $afterTargets
        ranges = $afterRanges
        premium_values_B9_CD115 = $premiumValues
        premium_formulas_B9_CD115 = $premiumFormulas
    }
    $json = $result | ConvertTo-Json -Depth 100
    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($OutputPath, $json, $encoding)
} finally {
    if ($book -ne $null) {
        $book.Close($false)
    }
    if ($excel -ne $null) {
        $excel.Quit()
    }
    if ($book -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($excel -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
