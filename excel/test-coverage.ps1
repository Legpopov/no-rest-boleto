param([Parameter(Mandatory=$true)][string]$WorkbookPath,[string]$PreviewPath='')
$ErrorActionPreference='Stop'
$app=New-Object -ComObject Excel.Application
$app.Visible=$false
$app.DisplayAlerts=$false
$book=$null
$passed=0
function Check($condition,[string]$message) {
    if (-not $condition) { throw "FAIL: $message" }
    $script:passed++
}
try {
    $book=$app.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path,0,$true)
    # Scenario inputs are written in batches; each assertion below follows an
    # explicit CalculateFull, avoiding incidental recalculation after every cell.
    $app.Calculation=-4135
    foreach ($name in 'modPlanningRegression','modCoverageRegression') {
        $book.VBProject.VBComponents.Import((Join-Path $PSScriptRoot "tests/$name.bas")) | Out-Null
    }
    $n=$app.Run("'$($book.Name)'!RunPlanningRegression",(Join-Path $PSScriptRoot '../tests/fixtures/planning-contract.json'),(Join-Path $PSScriptRoot '../tests/fixtures/planning-catalog-boundary.json'))
    if ($n -is [string]) { throw $n }; $passed += [int]$n
    $n=$app.Run("'$($book.Name)'!RunCoverageRegression")
    if ($n -is [string]) { throw $n }; $passed += [int]$n
    $trips=$book.Worksheets.Item('Командировки и выезды')
    $output=$book.Worksheets.Item('Прогноз эк.эфф-ти')
    $params=$book.Worksheets.Item('Параметры цен')
    $book.Worksheets.Item('Паспорт').Range('G23').Value2=2027
    $app.CalculateFull()
    Check ($trips.Range('J4').Value2 -eq '') 'empty transport remains empty'
    Check ($trips.Range('J5').Formula -match 'PlanStayForecast') 'indexed hotel formula installed'
    Check ($book.Names.Item('PlanningCities').RefersToRange.Rows.Count -gt 248) 'destinations extend beyond initial catalog'
    # Native sheet calculates a two-leg non-Moscow journey.
    $trips.Range('C4').Value2=2
    $trips.Range('L4').Value2='Воронеж'
    $trips.Range('F4').Value2='Москва'
    $trips.Range('Q4').Value2='поезд, авиа, автобус'
    $trips.Range('M4').Value2='test-journey'
    $trips.Range('N4').Value2=1
    $trips.Range('A9').Value2=$trips.Range('A4').Value2
    $trips.Range('B9').Value2=$trips.Range('B4').Value2
    $trips.Range('C9').Value2=2
    $trips.Range('L9').Value2='Москва'
    $trips.Range('F9').Value2='Казань'
    $trips.Range('Q9').Value2='поезд, авиа, автобус'
    $trips.Range('M9').Value2='test-journey'
    $trips.Range('N9').Value2=2
    $app.CalculateFull()
    Check ($trips.Range('J4').Value2 -is [double] -and $trips.Range('J4').Value2 -gt 0) 'new origin calculates'
    Check ($trips.Range('J9').Value2 -is [double] -and $trips.Range('J9').Value2 -gt 0) 'second leg calculates'
    Check ($trips.Range('W4').Text -match 'индекс пассажирского транспорта') 'sector proxy is labelled'
    $sum=[double]$trips.Range('J4').Value2+[double]$trips.Range('J9').Value2
    Check ([Math]::Abs([double]$output.Range('D27').Value2-$sum) -lt .01) 'final forecast sums legs once'
    $trips.Range('L9').Value2='Санкт-Петербург'
    $app.CalculateFull()
    Check ($trips.Range('J4').Text -match '#' -and $trips.Range('W4').Text -match 'Разрыв') 'broken chain rejected'
    $trips.Range('M4:N4').ClearContents()
    $trips.Range('M9:N9').ClearContents()
    $trips.Range('L4').Value2='Несуществующий тестовый город'
    $app.CalculateFull()
    Check ($trips.Range('J4').Text -match '#' -and $trips.Range('W4').Text -match 'вне покрытия') 'unknown origin is not replaced by Moscow'
    Check ($output.Range('D27').Text -match '#') 'missing fare propagates into forecast'
    $hotel=$book.Names.Item('PlanningHotels').RefersToRange
    $base=$null
    foreach ($r in $hotel.Rows) { if ($r.Cells.Item(1,1).Value2 -eq 'Москва') { $base=$r; break } }
    Check ($null -ne $base -and $base.Cells.Item(1,10).Value2 -eq 'ok') 'fresh filtered Moscow hotel base'
    $expected=0
    # Spring semester for launch year 2027 is February--June 2028.
    foreach ($m in 2..6) {
        $elapsed=(2028-[int]$base.Cells.Item(1,3).Value2)*12+$m-[int]$base.Cells.Item(1,4).Value2
        $expected += [double]$base.Cells.Item(1,2).Value2*[Math]::Pow(1+[double]$base.Cells.Item(1,5).Value2,$elapsed/12)
    }
    $actual=$app.Run("'$($book.Name)'!PlanStayForecast",2,'Москва','пребывание',3,2027,'1','2',$params.Range('O3:O6'),$hotel,'price')
    Check ([Math]::Abs([double]$actual-[Math]::Round(6*$expected/5,2)) -lt .01) 'hotel index matches independent calendar calculation'
    $trips.Range('L4').Value2='Воронеж'
    $trips.Range('C5').Value2=2
    $trips.Range('F5').Value2='Москва'
    $trips.Range('O5').Value2='пребывание'
    $trips.Range('P5').Value2=3
    $app.CalculateFull()
    Check ($trips.Range('J5').Value2 -is [double] -and $trips.Range('J5').Value2 -gt 0) 'hotel expense formula recalculates'
    $trips.PageSetup.PrintArea='$L$2:$W$5'
    $trips.PageSetup.Orientation=2
    $trips.PageSetup.Zoom=$false
    $trips.PageSetup.FitToPagesWide=1
    $trips.PageSetup.FitToPagesTall=1
    $preview=if ($PreviewPath) { [System.IO.Path]::GetFullPath($PreviewPath) } else { [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../_coverage-work/coverage-preview.pdf')) }
    $trips.ExportAsFixedFormat(0,$preview)
    "PASS: $passed native Excel checks; workbook and test inputs were not saved"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
