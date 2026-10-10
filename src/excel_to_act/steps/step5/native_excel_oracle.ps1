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

function Get-TargetRecord($namedRange, $shape) {
    $targetShape = $shape
    if ($null -eq $targetShape) {
        $targetShape = @()
    } elseif ($targetShape -isnot [array]) {
        $targetShape = @($targetShape)
    }
    if ($targetShape.Count -eq 0) {
        return Get-CellRecord $namedRange.Cells.Item(1, 1)
    }
    if ($targetShape.Count -gt 2) {
        throw 'Named result targets must be scalar, vector, or table values.'
    }
    foreach ($size in $targetShape) {
        if ([int]$size -lt 0) {
            throw 'Named result target dimensions cannot be negative.'
        }
    }
    if ($namedRange.Areas.Count -ne 1) {
        throw 'A non-scalar named result target must be one contiguous range.'
    }
    if (($targetShape | Where-Object { [int]$_ -eq 0 }).Count -gt 0) {
        $emptyValue = if ($targetShape.Count -eq 1) { @() } else {
            $emptyRows = [System.Collections.Generic.List[object]]::new()
            for ($row = 0; $row -lt [int]$targetShape[0]; $row++) {
                $emptyRows.Add(@())
            }
            @($emptyRows.ToArray())
        }
        return @{ value = $emptyValue; shape = $targetShape; formulas = @(); excel_errors = @() }
    }
    $rowsCount = [int]$namedRange.Rows.Count
    $columnsCount = [int]$namedRange.Columns.Count
    if ($targetShape.Count -eq 1) {
        $expectedCount = [int]$targetShape[0]
        if (($rowsCount -ne 1 -and $columnsCount -ne 1) -or ($rowsCount * $columnsCount -ne $expectedCount)) {
            throw "Named vector range $($namedRange.Address) has $rowsCount rows and $columnsCount columns; declared shape is [$expectedCount]."
        }
    } elseif ($rowsCount -ne [int]$targetShape[0] -or $columnsCount -ne [int]$targetShape[1]) {
        throw "Named table range $($namedRange.Address) has shape [$rowsCount,$columnsCount]; declared shape is [$($targetShape -join ',')]."
    }

    $values = [System.Collections.Generic.List[object]]::new()
    $formulas = [System.Collections.Generic.List[object]]::new()
    $errors = [System.Collections.Generic.List[object]]::new()
    if ($targetShape.Count -eq 1) {
        for ($index = 1; $index -le $rowsCount * $columnsCount; $index++) {
            if ($rowsCount -eq 1) {
                $cell = $namedRange.Cells.Item(1, $index)
            } else {
                $cell = $namedRange.Cells.Item($index, 1)
            }
            $record = Get-CellRecord $cell
            $values.Add($record.value)
            $formulas.Add($record.formula)
            $errors.Add($record.excel_error)
        }
        return @{ value = @($values.ToArray()); shape = $targetShape;
            formulas = @($formulas.ToArray()); excel_errors = @($errors.ToArray()) }
    }

    for ($row = 1; $row -le $rowsCount; $row++) {
        $valueRow = [System.Collections.Generic.List[object]]::new()
        $formulaRow = [System.Collections.Generic.List[object]]::new()
        $errorRow = [System.Collections.Generic.List[object]]::new()
        for ($column = 1; $column -le $columnsCount; $column++) {
            $record = Get-CellRecord $namedRange.Cells.Item($row, $column)
            $valueRow.Add($record.value)
            $formulaRow.Add($record.formula)
            $errorRow.Add($record.excel_error)
        }
        $values.Add(@($valueRow.ToArray()))
        $formulas.Add(@($formulaRow.ToArray()))
        $errors.Add(@($errorRow.ToArray()))
    }
    return @{ value = @($values.ToArray()); shape = $targetShape;
        formulas = @($formulas.ToArray()); excel_errors = @($errors.ToArray()) }
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
        $shapeProperty = $request.target_shapes.PSObject.Properties[[string]$name]
        $shape = if ($null -ne $shapeProperty) { $shapeProperty.Value } else { @() }
        $beforeTargets[[string]$name] = Get-TargetRecord $namedRange $shape
    }
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.CalculateFullRebuild()
    $timer.Stop()

    $afterTargets = @{}
    foreach ($name in $request.targets) {
        $namedRange = $book.Names.Item([string]$name).RefersToRange
        $shapeProperty = $request.target_shapes.PSObject.Properties[[string]$name]
        $shape = if ($null -ne $shapeProperty) { $shapeProperty.Value } else { @() }
        $afterTargets[[string]$name] = Get-TargetRecord $namedRange $shape
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
