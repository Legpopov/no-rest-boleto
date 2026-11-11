param([Parameter(Mandatory=$true)][string]$WorkbookPath)
# Read-only, unsaved native Excel checks; synthetic rows stay only in memory.
$ErrorActionPreference = 'Stop'
$excelPlan = New-Object -ComObject Excel.Application
$excelPlan.Visible = $false
$excelPlan.DisplayAlerts = $false
$bookPlan = $null
$passedPlan = 0
function Assert-Plan($Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:passedPlan++
}
try {
    $bookPlan = $excelPlan.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path,0,$true)
    # Test-only module is imported into the unsaved session, never the delivered book.
    $bookPlan.VBProject.VBComponents.Import((Join-Path $PSScriptRoot 'tests/modPlanningRegression.bas')) | Out-Null
    $regressionPlan = $excelPlan.Run("'$($bookPlan.Name)'!RunPlanningRegression", (Join-Path $PSScriptRoot '../tests/fixtures/planning-contract.json'), (Join-Path $PSScriptRoot '../data/planning.json'))
    if ($regressionPlan -is [string]) { throw $regressionPlan }
    $passedPlan += [int]$regressionPlan
    $tripsPlan = $bookPlan.Worksheets.Item('Командировки и выезды')
    $paramsPlan = $bookPlan.Worksheets.Item('Параметры цен')
    $dataPlan = $bookPlan.Worksheets.Item('Данные плана')
    $macroPlan = "'$($bookPlan.Name)'!PlanFare"
    $catalogPlan = $bookPlan.Names.Item('PlanningData').RefersToRange
    $coveragePlan = $bookPlan.Names.Item('PlanningCoverage').RefersToRange
    $calendarPlan = $paramsPlan.Range('O3:O6')
    $bookPlan.Worksheets.Item('Паспорт').Range('G23').Value2 = 2027
    $excelPlan.CalculateFull()
    Assert-Plan ($bookPlan.Worksheets.Item('Параметры').Range('C2').Value2 -eq 2027) 'launch-year input reaches planning formulas'
    Assert-Plan ($tripsPlan.Range('J4').Value2 -eq '') 'empty transport input is empty'
    $paramsPlan.Range('O5').Value2 = 5
    $paramsPlan.Range('O6').Value2 = 2
    $actualPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','авиа, поезд, автобус','ДА',$calendarPlan,$catalogPlan,'price')
    $recordsPlan = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../data/planning.json') | ConvertFrom-Json
    $matchedPlan = @($recordsPlan.records | Where-Object { $_.origin -eq 'Москва' -and $_.destination -eq 'Санкт-Петербург' -and $_.baggage_included -and $_.month -in 5,6 })
    $expectedPlan = [Math]::Round((($matchedPlan | ForEach-Object { $_.price * [Math]::Pow(1+$_.annual_rate,2) } | Measure-Object -Average).Average),2)
    Assert-Plan ([Math]::Abs($actualPlan-$expectedPlan) -lt 0.01) '2028 compound forecast matches independent source calculation'
    $otherOriginPlan = $excelPlan.Run($macroPlan,'Екатеринбург','Казань',2027,'1 год набора','2 семестр','авиа','ДА',$calendarPlan,$catalogPlan,'price')
    Assert-Plan ($otherOriginPlan -is [double] -and $otherOriginPlan -gt 0) 'non-Moscow origin'
    $duplicatePlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','авиа, авиа','ДА',$calendarPlan,$catalogPlan,'status')
    Assert-Plan ($duplicatePlan -match 'повтор') 'duplicate preference rejected'
    $paramsPlan.Range('O5').Value2 = 2
    $paramsPlan.Range('O6').Value2 = 5
    $missingPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','авиа','ДА',$calendarPlan,$catalogPlan,'status')
    Assert-Plan ($missingPlan -match '2028-02' -and $missingPlan -match '2028-04') 'missing semester months identified'
    $tripsPlan.Range('C4').Value2 = 1
    $tripsPlan.Range('F4').Value2 = 'Санкт-Петербург'
    $tripsPlan.Range('L4').Value2 = 'Москва'
    $tripsPlan.Range('Q4').Value2 = 'авиа'
    $excelPlan.CalculateFull()
    Assert-Plan ($tripsPlan.Range('J4').Text -match '#') 'incomplete populated row returns Excel error'
    $outputPlan = $bookPlan.Worksheets.Item('Прогноз эк.эфф-ти')
    $errorsPlan = 0
    foreach ($cellPlan in $outputPlan.Range('D26:O30')) {
        if ($cellPlan.Text -match '#') { $errorsPlan++ }
    }
    Assert-Plan ($errorsPlan -gt 0) 'incomplete expense propagates into forecast output'
    # A second transport leg within the same cohort/semester.
    $tripsPlan.Range('A9').Value2 = $tripsPlan.Range('A4').Value2
    $tripsPlan.Range('B9').Value2 = $tripsPlan.Range('B4').Value2
    $tripsPlan.Range('C9').Value2 = 1
    $tripsPlan.Range('M4').Value2 = 'test-route'
    $tripsPlan.Range('M9').Value2 = 'test-route'
    $tripsPlan.Range('N4').Value2 = 1
    $tripsPlan.Range('N9').Value2 = 2
    $tripsPlan.Range('L9').Value2 = 'Санкт-Петербург'
    $tripsPlan.Range('F9').Value2 = 'Казань'
    $guardPlan = "'$($bookPlan.Name)'!PlanRouteIssue"
    $issuePlan = $excelPlan.Run($guardPlan,4,$tripsPlan.Range('L3:N152'),$tripsPlan.Range('F3:F152'),$tripsPlan.Range('A3:C152'),$tripsPlan.Range('H3:H152'))
    Assert-Plan ($issuePlan -eq '') 'connected route accepted'
    $tripsPlan.Range('L9').Value2 = 'Екатеринбург'
    $issuePlan = $excelPlan.Run($guardPlan,4,$tripsPlan.Range('L3:N152'),$tripsPlan.Range('F3:F152'),$tripsPlan.Range('A3:C152'),$tripsPlan.Range('H3:H152'))
    Assert-Plan ($issuePlan -match 'Разрыв') 'disconnected route rejected'
    $tripsPlan.Range('N9').Value2 = 1
    $issuePlan = $excelPlan.Run($guardPlan,4,$tripsPlan.Range('L3:N152'),$tripsPlan.Range('F3:F152'),$tripsPlan.Range('A3:C152'),$tripsPlan.Range('H3:H152'))
    Assert-Plan ($issuePlan -match 'Повтор') 'duplicate leg rejected'
    # Two synthetic rail months, deliberately preferred over air. Never saved.
    $endPlan = $dataPlan.Cells($dataPlan.Rows.Count,1).End(-4162).Row
    foreach ($monthPlan in 5,6) {
        $endPlan++
        $dataPlan.Cells($endPlan,1).Value2 = 'Москва'
        $dataPlan.Cells($endPlan,2).Value2 = 'Санкт-Петербург'
        $dataPlan.Cells($endPlan,3).Value2 = 'rail'
        $dataPlan.Cells($endPlan,4).Value2 = [double]$monthPlan
        $dataPlan.Cells($endPlan,5).Value2 = [double]2026
        $dataPlan.Cells($endPlan,6).Value2 = [double]100
        $dataPlan.Cells($endPlan,7).Value2 = -0.1
        $dataPlan.Cells($endPlan,8).Value2 = $false
        $dataPlan.Cells($endPlan,9).Value2 = 'synthetic test only'
        $dataPlan.Cells($endPlan,15).Value2 = 'test_class'
        $dataPlan.Cells($endPlan,16).Value2 = $true
        $dataPlan.Cells($endPlan,17).Value2 = [double]1
    }
    $catalogPlan = $dataPlan.Range("A2:S$endPlan")
    $paramsPlan.Range('O5').Value2 = 5
    $paramsPlan.Range('O6').Value2 = 2
    $modePlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','поезд, авиа','ДА',$calendarPlan,$catalogPlan,'mode','test_class')
    Assert-Plan ($modePlan -eq 'Поезд') 'user preference chooses rail'
    $railPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','поезд, авиа','ДА',$calendarPlan,$catalogPlan,'price','test_class')
    Assert-Plan ([Math]::Abs($railPlan-81) -lt 0.01) 'negative annual rates compound correctly'
    $stayPlan = "'$($bookPlan.Name)'!PlanStay"
    $hotelSourcePlan = $bookPlan.Worksheets.Item('Справочник').Range('BK4:BR248')
    $transitPlan = $excelPlan.Run($stayPlan,2,'Санкт-Петербург','пересадка','',$hotelSourcePlan,'price')
    Assert-Plan ($transitPlan -eq 0) 'transit does not invent hotel nights'
    $missingNightsPlan = $excelPlan.Run($stayPlan,2,'Санкт-Петербург','пребывание','',$hotelSourcePlan,'status')
    Assert-Plan ($missingNightsPlan -match 'укажите число ночей') 'stay requires explicit nights'
    $hotelPricePlan = $excelPlan.Run($stayPlan,2,'Санкт-Петербург','пересадка',3,$hotelSourcePlan,'price')
    Assert-Plan ($hotelPricePlan -gt 0) 'explicit hotel nights can be charged during transit'
    $hotelStatusPlan = $excelPlan.Run($stayPlan,2,'Санкт-Петербург','пребывание',3,$hotelSourcePlan,'status')
    Assert-Plan ($hotelStatusPlan -match 'без индексации') 'hotel forecasting limitation is visible'
    # Real rail/bus references, independently recomputed from source snapshot.
    $catalogPlan = $bookPlan.Names.Item('PlanningData').RefersToRange
    $surfacePlan = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../data/surface/latest.json') | ConvertFrom-Json
    $paramsPlan.Range('O5').Value2 = 2
    $paramsPlan.Range('O6').Value2 = 5
    $rawRailPlan = @($surfacePlan.routes | Where-Object { $_.mode -eq 'rail' -and $_.origin -eq 'Москва' -and $_.destination -eq 'Санкт-Петербург' })[0]
    $coupePlan = @($rawRailPlan.items | Where-Object { $_.firm -and $_.classes.coupe } | ForEach-Object { $_.classes.coupe })
    $expectedRailPlan = [Math]::Round((($coupePlan | Measure-Object -Average).Average)*[Math]::Pow(1.04,2),2)
    $actualRailPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','поезд','ДА',$calendarPlan,$catalogPlan,'price','coupe:firm',$coveragePlan)
    Assert-Plan ([Math]::Abs($actualRailPlan-$expectedRailPlan) -lt .02) 'firm-coupe filter and real rail forecast'
    $rawBusPlan = @($surfacePlan.routes | Where-Object { $_.mode -eq 'bus' -and $_.origin -eq 'Москва' -and $_.destination -eq 'Санкт-Петербург' })[0]
    $expectedBusPlan = [Math]::Round((($rawBusPlan.items.price | Measure-Object -Average).Average)*[Math]::Pow(1.04,2),2)
    $actualBusPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','автобус','ДА',$calendarPlan,$catalogPlan,'price','coupe',$coveragePlan)
    Assert-Plan ([Math]::Abs($actualBusPlan-$expectedBusPlan) -lt .02) 'real bus forecast'
    $surfaceStatusPlan = $excelPlan.Run($macroPlan,'Москва','Санкт-Петербург',2027,'1 год набора','2 семестр','автобус','ДА',$calendarPlan,$catalogPlan,'status','coupe',$coveragePlan)
    Assert-Plan ($surfaceStatusPlan -match 'без сезонности' -and $surfaceStatusPlan -match 'не индекс транспорта') 'reserve limitations are visible'
    $noDirectPlan = $excelPlan.Run($macroPlan,'Калининград','Москва',2027,'1 год набора','2 семестр','автобус','ДА',$calendarPlan,$catalogPlan,'status','coupe',$coveragePlan)
    Assert-Plan ($noDirectPlan -match 'прямого рейса нет') 'no-direct status differs from missing historical months'
    $outsidePlan = $excelPlan.Run($macroPlan,'Воронеж','Москва',2027,'1 год набора','2 семестр','авиа, поезд, автобус','ДА',$calendarPlan,$catalogPlan,'status','coupe',$coveragePlan)
    Assert-Plan ($outsidePlan -match 'Город отправления вне покрытия') 'outside origin is not substituted with Moscow'
    $outsideDestPlan = $excelPlan.Run($macroPlan,'Москва','Воронеж',2027,'1 год набора','2 семестр','авиа, поезд, автобус','ДА',$calendarPlan,$catalogPlan,'status','coupe',$coveragePlan)
    Assert-Plan ($outsideDestPlan -match 'Город прибытия вне покрытия') 'outside destination is explicit'
    $tripsPlan.Range('M4:N4').ClearContents()
    $tripsPlan.Range('M9:N9').ClearContents()
    $tripsPlan.Range('L4').Value2 = 'Воронеж'
    $tripsPlan.Range('F4').Value2 = 'Москва'
    $excelPlan.CalculateFull()
    Assert-Plan ($tripsPlan.Range('J4').Text -match '#' -and $tripsPlan.Range('W4').Text -match 'Город отправления вне покрытия') 'actual worksheet blocks unsupported origin and explains it'
    $tripsPlan.Range('L4').Value2 = 'Москва'
    $tripsPlan.Range('F4').Value2 = 'Санкт-Петербург'
    $tripsPlan.Range('Q4').Value2 = 'поезд, автобус, авиа'
    $excelPlan.CalculateFull()
    Assert-Plan ($tripsPlan.Range('J4').Value2 -is [double] -and $tripsPlan.Range('J4').Value2 -gt 0 -and $tripsPlan.Range('V4').Value2 -eq 'Поезд') 'actual populated worksheet calculates rail budget'
    $expectedTotalPlan = [double]$tripsPlan.Range('J4').Value2 + [double]$tripsPlan.Range('J9').Value2
    Assert-Plan ([Math]::Abs([double]$outputPlan.Range('D27').Value2-$expectedTotalPlan) -lt .01) 'native forecast SUMIFS includes both transport legs'
    "PASS: $passedPlan native Excel checks; workbook was not saved"
} finally {
    if ($null -ne $bookPlan) { $bookPlan.Close($false) }
    $excelPlan.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelPlan) | Out-Null
}
