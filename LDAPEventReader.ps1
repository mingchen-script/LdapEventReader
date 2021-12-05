# Read me
# This script will convert LDAP events 1644 into Excel pivot tables for workload analysis by:
#    1. Scan all evtx files in script directory for event 1644, and export to CSV.
#    2. Calls into Excel to import resulting CSV, create pivot tables for common ldap search analysis scenarios. 
# Script requires Excel 2013 installed. 64bits Excel will allow generation of larger worksheet.
#
# To use the script:
#  1. Convert pre-2008 evt to evtx using later OS. (Please note, pre-2008 does not contain all 16 data fields. So some pivot tables might not display correctly.)

# LdapEventReader.ps1 v2.14 12/4/2021(timerange + event numbers in [More Info])
	#		Steps: 
	#   	1. Copy Directory Service EVTX from target DC(s) to same directory as this script.
	#     		Tip: When copying Directory Service EVTX, filter on event 1644 to reduce EVTX size for quicker transfer. 
	#					Note: Script will process all *.EVTX in script directory when run.
	#   	2. Run script

# Script info:    https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance
#   Latest:       https://github.com/mingchen-script/LdapEventReader
# AD Schema:      https://docs.microsoft.com/en-us/windows/win32/adschema/active-directory-schema
# AD Attributes:  https://docs.microsoft.com/en-us/windows/win32/adschema/attributes

#------Script variables block, modify to fit your needs ---------------------------------------------------------------------
$g_StartTime = '2010/01/12  09:53'    # Earliest 1644 event to export, in the form of M/d/yyyy H:m:s tt' example: '3/19/2010 1:11:49 AM'. Use this to filter events after changes.
$g_LookBackDays = 0 #2080             # 0 means script start list events after $g_StartTime, when set to an interger, script will list events in last $g_LookBackDays days. For examle: 1 will list events occurs in last 24 hours. Use this to filter events after changes.
$g_MaxExports = 5000                  # Max number of 1644 events to export per each EVTX. Use this for quicker spot checks.
$g_MaxThreads = 4                     # Max concurrent Evtx to CSV export threads (jobs), hight number might hit File/IO bottleneck since all files are in one directory.
$g_ColorBar   = $True                 # Can set to $false to speed up excel import & reduce memory requirement. 
$g_ColorScale = $True                 # Can set to $false to speed up excel import & reduce memory requirement. Color Scale requires '$g_ColorBar = $True' for color index. 
$ErrorActionPreference = "SilentlyContinue"
function Set-PivotField { param ( $PivotField = $null, $Orientation = $null, $NumberFormat = $null, $Function = $null, $Calculation = $null, $Name = $null, $Group = $null )
    if ($null -ne $Orientation) {$PivotField.Orientation = $Orientation}
    if ($null -ne $NumberFormat) {$PivotField.NumberFormat = $NumberFormat}
    if ($null -ne $Function) {$PivotField.Function = $Function}
    if ($null -ne $Calculation) {$PivotField.Calculation = $Calculation}
    if ($null -ne $Name) {$PivotField.Name = $Name}
    if ($null -ne $Group) {($PivotField.DataRange.Item($group)).group($true,$true,1,($false, $true, $true, $true, $false, $false, $false)) | Out-Null}
}
function Set-PivotPageRows { param ( $Sheet = $null, $PivotTable = $null, $Page = $null, $Rows = $null  )
    $xlRowField   = 1 #XlPivotFieldOrientation 
    $xlPageField  = 3 #XlPivotFieldOrientation 
    Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$Page") -Orientation $xlPageField
    $i=0
    ($Rows).foreach({
      $i++
      If ($i -lt ($Rows).count) {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField}
      else {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField -Group $i}
    })
}
function Set-TableFormats { param ( $Sheet = $null, $Table = $null, $ColumnWidth = $null, $label = $null, $Name = $null, $ColorScale = $null, $ColorBar = $null, $SortColumn = $null, $Hide = $null, $ColumnHiLite = $null, $NoteColumn = $null, $Note = $null )
  $Sheet.PivotTables("$Table").HasAutoFormat = $False
    $Column = 1
    $ColumnWidth.foreach({ $Sheet.columns.item($Column).columnwidth = $_
      $Column++
    })
    $Sheet.Application.ActiveWindow.SplitRow = 3
    $Sheet.Application.ActiveWindow.SplitColumn = 2
    $Sheet.Application.ActiveWindow.FreezePanes = $true
    $Sheet.Cells.Item(3,1) = $label
    $Sheet.Name = $Name
    if ($null -ne $SortColumn) {$null = $Sheet.Cells.Item($SortColumn,4).Sort($Sheet.Cells.Item($SortColumn,4),2)}
    if ($null -ne $Hide) {$Hide.foreach({($Sheet.PivotTables("$Table").PivotFields($_)).ShowDetail = $false})}
    if ($null -ne $ColumnHiLite) {
      $Sheet.Range("A4:"+[char]($sheet.UsedRange.Cells.Columns.count+64)+[string](($Sheet.UsedRange.Cells).Rows.count-1)).interior.Color = 16056319
      $ColumnHiLite.ForEach({$sheet.Range(($_+"3")).interior.ColorIndex = 37})
    }
    if (($null -ne $ColorBar) -and ($g_ColorBar -eq $true)) {
      $ColorRange='$'+$ColorBar+'$4:$'+$ColorBar+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddDatabar()
    }
    if (($null -ne $ColorScale) -and ($g_ColorScale -eq $true)) {
      $ColorRange='$'+$ColorScale+'$4:$'+$ColorScale+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddColorScale(3)
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).type = 1
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).FormatColor.Color = 8109667
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(2).FormatColor.Color = 8711167
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).type = 2 
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).FormatColor.Color = 7039480
    }
    $Sheet.Cells.Item(1,$NoteColumn)= "[More Info]" #--Add log info
      $null = $Sheet.Cells.Item(1,$NoteColumn).addcomment()
      $null = $Sheet.Cells.Item(1,$NoteColumn).comment.text($Note)
      $Sheet.Cells.Item(1,$NoteColumn).comment.shape.textframe.Autosize = $true
  }
function Export-1644CSV { param ( $InFile = $null, $OutFile = $null, $StartTime = $null, $MaxExports = $null )
  $1644s = Get-WinEvent -Path $InFile -FilterXPath "Event[ System[ EventID = 1644 and Channel = 'Directory Service' and TimeCreated[@SystemTime>='$StartTime'] ] ] " -MaxEvents $MaxExports -ErrorAction SilentlyContinue
  If ($null -ne $1644s) {
  	$Header=$null
    $1644 = New-Object System.Object
  	($1644s).foreach({
      $1644 | Add-Member -MemberType NoteProperty -Name LDAPServer -force -Value $_.MachineName 
      $1644 | Add-Member -MemberType NoteProperty -Name TimeGenerated -force -Value $_.TimeCreated	
      $1644 | Add-Member -MemberType NoteProperty -Name ClientIP -force -Value $_.Properties[4].Value.Substring(0,$_.Properties[4].Value.LastIndexOf(":"))
      $1644 | Add-Member -MemberType NoteProperty -Name ClientPort -force -Value $_.Properties[4].Value.Substring($_.Properties[4].Value.LastIndexOf(":")+1)
      $1644 | Add-Member -MemberType NoteProperty -Name StartingNode -force -Value $_.Properties[0].Value
      $1644 | Add-Member -MemberType NoteProperty -Name Filter -force -Value $_.Properties[1].Value
      $1644 | Add-Member -MemberType NoteProperty -Name SearchScope -force -Value $_.Properties[5].Value
      $1644 | Add-Member -MemberType NoteProperty -Name AttributeSelection -force -Value $_.Properties[6].Value
      $1644 | Add-Member -MemberType NoteProperty -Name ServerControls -force -Value $_.Properties[7].Value
      $1644 | Add-Member -MemberType NoteProperty -Name VisitedEntries -force -Value $_.Properties[2].Value
      $1644 | Add-Member -MemberType NoteProperty -Name ReturnedEntries -force -Value $_.Properties[3].Value
      $1644 | Add-Member -MemberType NoteProperty -Name UsedIndexes -force -Value $_.Properties[8].Value 
      $1644 | Add-Member -MemberType NoteProperty -Name PagesReferenced -force -Value $_.Properties[9].Value
      $1644 | Add-Member -MemberType NoteProperty -Name PagesReadFromDisk -force -Value $_.Properties[10].Value
      $1644 | Add-Member -MemberType NoteProperty -Name PagesPreReadFromDisk -force -Value $_.Properties[11].Value
      $1644 | Add-Member -MemberType NoteProperty -Name CleanPagesModified -force -Value $_.Properties[12].Value
      $1644 | Add-Member -MemberType NoteProperty -Name DirtyPagesModified -force -Value $_.Properties[13].Value
      $1644 | Add-Member -MemberType NoteProperty -Name SearchTimeMS -force -Value $_.Properties[14].Value
      $1644 | Add-Member -MemberType NoteProperty -Name AttributesPreventingOptimization -force -Value $_.Properties[15].Value	
      $1644 | Add-Member -MemberType NoteProperty -Name User -force -Value $_.Properties[16].Value	
      If ($null -eq $Header) {
        $1644 | Select-Object -Property LDAPServer,TimeGenerated,StartingNode,Filter,VisitedEntries,ReturnedEntries,ClientIP,ClientPort,SearchScope,AttributeSelection,ServerControls,UsedIndexes,PagesReferenced,PagesReadFromDisk,PagesPreReadFromDisk,CleanPagesModified,DirtyPagesModified,SearchTimeMS,AttributesPreventingOptimization,User | Export-Csv $OutFile -NoTypeInformation
        $Header = $True
      } else { $1644 | Select-Object -Property LDAPServer,TimeGenerated,StartingNode,Filter,VisitedEntries,ReturnedEntries,ClientIP,ClientPort,SearchScope,AttributeSelection,ServerControls,UsedIndexes,PagesReferenced,PagesReadFromDisk,PagesPreReadFromDisk,CleanPagesModified,DirtyPagesModified,SearchTimeMS,AttributesPreventingOptimization,User | Export-Csv $OutFile -NoTypeInformation -Append }
    })
  } else {  
    # Write-Host '    No event 1644 found in' $InFile  -ForegroundColor Red 
  }
}

$Export1644CSV = {
  function Export-1644CSV { param ( $InFile = $null, $OutFile = $null, $StartTime = $null, $MaxExports = $null )
    $1644s = Get-WinEvent -Path $InFile -FilterXPath "Event[ System[ EventID = 1644 and Channel = 'Directory Service' and TimeCreated[@SystemTime>='$StartTime'] ] ] " -MaxEvents $MaxExports -ErrorAction SilentlyContinue
    If ($null -ne $1644s) {
      $Header=$null
      $1644 = New-Object System.Object
      ($1644s).foreach({
        $1644 | Add-Member -MemberType NoteProperty -Name LDAPServer -force -Value $_.MachineName 
        $1644 | Add-Member -MemberType NoteProperty -Name TimeGenerated -force -Value $_.TimeCreated	
        $1644 | Add-Member -MemberType NoteProperty -Name ClientIP -force -Value $_.Properties[4].Value.Substring(0,$_.Properties[4].Value.LastIndexOf(":"))
        $1644 | Add-Member -MemberType NoteProperty -Name ClientPort -force -Value $_.Properties[4].Value.Substring($_.Properties[4].Value.LastIndexOf(":")+1)
        $1644 | Add-Member -MemberType NoteProperty -Name StartingNode -force -Value $_.Properties[0].Value
        $1644 | Add-Member -MemberType NoteProperty -Name Filter -force -Value $_.Properties[1].Value
        $1644 | Add-Member -MemberType NoteProperty -Name SearchScope -force -Value $_.Properties[5].Value
        $1644 | Add-Member -MemberType NoteProperty -Name AttributeSelection -force -Value $_.Properties[6].Value
        $1644 | Add-Member -MemberType NoteProperty -Name ServerControls -force -Value $_.Properties[7].Value
        $1644 | Add-Member -MemberType NoteProperty -Name VisitedEntries -force -Value $_.Properties[2].Value
        $1644 | Add-Member -MemberType NoteProperty -Name ReturnedEntries -force -Value $_.Properties[3].Value
        $1644 | Add-Member -MemberType NoteProperty -Name UsedIndexes -force -Value $_.Properties[8].Value 
        $1644 | Add-Member -MemberType NoteProperty -Name PagesReferenced -force -Value $_.Properties[9].Value
        $1644 | Add-Member -MemberType NoteProperty -Name PagesReadFromDisk -force -Value $_.Properties[10].Value
        $1644 | Add-Member -MemberType NoteProperty -Name PagesPreReadFromDisk -force -Value $_.Properties[11].Value
        $1644 | Add-Member -MemberType NoteProperty -Name CleanPagesModified -force -Value $_.Properties[12].Value
        $1644 | Add-Member -MemberType NoteProperty -Name DirtyPagesModified -force -Value $_.Properties[13].Value
        $1644 | Add-Member -MemberType NoteProperty -Name SearchTimeMS -force -Value $_.Properties[14].Value
        $1644 | Add-Member -MemberType NoteProperty -Name AttributesPreventingOptimization -force -Value $_.Properties[15].Value	
        $1644 | Add-Member -MemberType NoteProperty -Name User -force -Value $_.Properties[16].Value	
        If ($null -eq $Header) {
          $1644 | Select-Object -Property LDAPServer,TimeGenerated,StartingNode,Filter,VisitedEntries,ReturnedEntries,ClientIP,ClientPort,SearchScope,AttributeSelection,ServerControls,UsedIndexes,PagesReferenced,PagesReadFromDisk,PagesPreReadFromDisk,CleanPagesModified,DirtyPagesModified,SearchTimeMS,AttributesPreventingOptimization,User | Export-Csv $OutFile -NoTypeInformation
          $Header = $True
        } else { $1644 | Select-Object -Property LDAPServer,TimeGenerated,StartingNode,Filter,VisitedEntries,ReturnedEntries,ClientIP,ClientPort,SearchScope,AttributeSelection,ServerControls,UsedIndexes,PagesReferenced,PagesReadFromDisk,PagesPreReadFromDisk,CleanPagesModified,DirtyPagesModified,SearchTimeMS,AttributesPreventingOptimization,User | Export-Csv $OutFile -NoTypeInformation -Append }
      })
    } else {  
      # Write-Host '    No event 1644 found in' $InFile  -ForegroundColor Red 
    }
  }
}

#------Main---------------------------------
$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
  $TotalSteps = ((Get-ChildItem -Path $ScriptPath -Filter '*.evtx').count)+9
  $Step=1
$TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
$StartTime = ([datetime]$g_StartTime).ToUniversalTime().ToString("s")
  if ($g_LookBackDays -ne 0) { $StartTime = ((Get-Date).AddDays(0-$g_LookBackDays)).ToUniversalTime().ToString("s") }
(Get-ChildItem -Path $ScriptPath -Filter '*.evtx').foreach({
  Write-Progress -Activity "Generating $_ to CSV" -PercentComplete (($Step++/$TotalSteps)*100)
  If ($g_MaxThreads -le 1) {
    Export-1644CSV -InFile "$ScriptPath\$_" -OutFile ("$ScriptPath\$TimeStamp-Temp1644-"+$_.BaseName+".csv") -StartTime $StartTime -MaxExports $g_MaxExports
  } else { #--- Start $g_MaxThreads jobs-
    Start-Job -name $_.BaseName -InitializationScript $Export1644CSV -ArgumentList @("$ScriptPath\$_", ("$ScriptPath\$TimeStamp-Temp1644-"+$_.BaseName+".csv"), $StartTime, $g_MaxExports) -ScriptBlock{
      Export-1644CSV -InFile $Args[0] -OutFile $Args[1] -StartTime $Args[2] -MaxExports $Args[3]
    } | Out-Null
    While((Get-Job -State 'Running').Count -ge $g_MaxThreads) { Start-Sleep -Milliseconds 10 }
  }
})
  While((Get-Job -State 'Running').Count -gt 0) { Start-Sleep -Milliseconds 10  }
    Get-Job -State Completed | Remove-Job 
#---------Find logs's time range Info----------
  $OldestTimeStamp = $NewestTimeStamp = $LogsInfo = $null
  (Get-ChildItem -Path $ScriptPath\* -include ('*.csv') ).foreach({
    $FirstTimeStamp = [DateTime]((Get-Content $_ -Tail 1) -split ',' | Select-Object -skip 1 -first 1 | ForEach-Object { $_ -replace '"',$null})
    $LastTimeStamp = [DateTime]((Get-Content $_ -Head 2) -split ',' | Select-Object -skip 21 -first 1 | ForEach-Object { $_ -replace '"',$null})
      if ($OldestTimeStamp -eq $null) { $OldestTimeStamp = $NewestTimeStamp = $FirstTimeStamp }
      If ($OldestTimeStamp -gt $FirstTimeStamp) {$OldestTimeStamp = $FirstTimeStamp }
      If ($NewestTimeStamp -lt $LastTimeStamp) {$NewestTimeStamp = $LastTimeStamp }
      $LogsInfo = $LogsInfo + ($_.name+"`n   "+$FirstTimeStamp+' ~ '+$LastTimeStamp+"`t   Log range = "+($LastTimeStamp-$FirstTimeStamp).Days+" Days "+($LastTimeStamp-$FirstTimeStamp).Hours+" Hours "+($LastTimeStamp-$FirstTimeStamp).Minutes+" min "+($LastTimeStamp-$FirstTimeStamp).Seconds+" sec. ("+((Get-Content $_ | Measure-Object -line).lines-1)+" Events.)`n`n")
  })
    $LogTimeRange = ($NewestTimeStamp-$OldestTimeStamp)
    $LogRangeText += ("Script info:`n   https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance`n") 
    $LogRangeText += ("Github latest download:`n   https://github.com/mingchen-script/LdapEventReader`n`n") 
    $LogRangeText += ("AD Schema:`n   https://docs.microsoft.com/en-us/windows/win32/adschema/active-directory-schema`n") 
    $LogRangeText += ("AD Attributes:`n   https://docs.microsoft.com/en-us/windows/win32/adschema/attributes`n`n") 
    $LogRangeText += ("#-------------------------------`n  [Overall EventRange]: "+$OldestTimeStamp+' ~ '+$NewestTimeStamp+"`n  [Overall TimeRange]: "+$LogTimeRange.Days+' Days '+$LogTimeRange.Hours+' Hours '+$LogTimeRange.Minutes+' Minutes '+$LogTimeRange.Seconds+" Seconds `n`n") + ($LogsInfo -replace "$TimeStamp-Temp1644-")
#-----Combine CSV(s) into one for faster Excel import
  $OutTitle1 = 'LDAP searches'
  $OutFile1 = "$ScriptPath\$TimeStamp-$OutTitle1.csv"
  Write-Progress -Activity "Generating $OutTitle1" -PercentComplete (($Step++/$TotalSteps)*100)
    Get-ChildItem -Path $ScriptPath -Filter "$TimeStamp-Temp1644-*.csv" | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv  $OutFile1 -NoTypeInformation -Append 
    $null = Get-ChildItem -Path $ScriptPath -Filter "$TimeStamp-Temp1644-*.csv" | Remove-Item
#----Excel COM variables-------------------------------------------------------------------
  $fmtNumber  = "###,###,###,###,###"
  $fmtPercent = "#0.00%"
  $xlDataField  = 4 #XlPivotFieldOrientation 
  $xlAverage    = -4106 #XlConsolidationFunction
  $xlSum        = -4157 #XlConsolidationFunction 
  $xlPercentOfTotal = 8 #XlPivotFieldCalculation 
#-------Import to Excel
If (Test-Path $OutFile1) { 
  $Excel = New-Object -ComObject excel.application
  Write-Progress -Activity "Import to Excel $OutTitle1" -PercentComplete (($Step++/$TotalSteps)*100)
    # $Excel.visible = $true
    $Excel.Workbooks.OpenText("$OutFile1")
    $Sheet0 = $Excel.Workbooks[1].Worksheets[1]
      $Sheet0.Application.ActiveWindow.SplitRow=1  
      $Sheet0.Application.ActiveWindow.FreezePanes = $true
      $null = $Sheet0.Columns.AutoFit() = $Sheet0.Range("A1").AutoFilter()
        ("C","D","J","K","L").ForEach({$Sheet0.Columns.Item($_).columnwidth = 70})
        ("E","F","H","M","N","O","P","Q","R").ForEach({$Sheet0.Columns.Item($_).numberformat = $fmtNumber})
        $Sheet0.Columns.Item("B").numberformat = "m/d/yyyy h:mm:s AM/PM"
      $Sheet0.Name = $OutTitle1
      $null = $Sheet0.ListObjects.Add(1, $Sheet0.Application.ActiveCell.CurrentRegion, $null ,0)
    #----Pivot Table 1-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount StartingNode Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet1 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet1!R1C1")
      Set-PivotPageRows -Sheet $sheet1 -PivotTable "PivotTable1" -Page "LDAPServer" -Rows ("StartingNode","Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet1 -Table "PivotTable1" -ColumnWidth (60,12,14,12,14) -label 'StartingNode grouping' -Name '1.TopCount StartingNode' -SortColumn 4 -Hide ('ClientIP','Filter','StartingNode') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
      #----Pivot Table 2-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet2 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet2!R1C1")
      Set-PivotPageRows -Sheet $sheet2 -PivotTable "PivotTable2" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet2 -Table "PivotTable2" -ColumnWidth (60,12,19,12) -label 'IP grouping' -Name '2.TopCount IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 3-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount Filters Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet3 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet3!R1C1")
      Set-PivotPageRows -Sheet $sheet3 -PivotTable "PivotTable3" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet3 -Table "PivotTable3" -ColumnWidth (70,12,19,12) -label 'Filter grouping' -Name '3.TopCount Filters' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 4-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet4 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet4!R1C1")
      Set-PivotPageRows -Sheet $sheet4 -PivotTable "PivotTable4" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet4 -Table "PivotTable4" -ColumnWidth (50,21,12,19) -label 'IP grouping' -Name '4.TopTime IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 5-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime Filter Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet5 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet5!R1C1")
      Set-PivotPageRows -Sheet $sheet5 -PivotTable "PivotTable5" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet5 -Table "PivotTable5" -ColumnWidth (70,21,12,19) -label 'IP grouping' -Name '5.TopTime Filter' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 6-------------------------------------------------------------------
    Write-Progress -Activity "Creating Top Users Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet6 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet6!R1C1")
      Set-PivotPageRows -Sheet $Sheet6 -PivotTable "PivotTable6" -Page "LDAPServer" -Rows ("User","ClientIP","Filter")
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet6 -Table "PivotTable6" -ColumnWidth (70,21,12,19) -label 'User grouping' -Name '6.Top User IP Filter' -SortColumn 4 -Hide ('Filter','ClientIP','User') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 7-------------------------------------------------------------------
    Write-Progress -Activity "Creating Top Filter Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet7 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet7!R1C1")
      Set-PivotPageRows -Sheet $Sheet7 -PivotTable "PivotTable7" -Page "LDAPServer" -Rows ("AttributesPreventingOptimization","Filter","ClientIP")
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet7 -Table "PivotTable7" -ColumnWidth (70,21,12,19) -label 'Attributes Preventing Optimization' -Name '7.Attributes Need Optimization' -SortColumn 4 -Hide ('ClientIP','Filter','AttributesPreventingOptimization') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #---General Tab Operations-------------------------------------------------------------------
    ($Sheet1,$Sheet2,$Sheet3).ForEach{$_.Tab.ColorIndex = 35}
    ($Sheet4,$Sheet5).ForEach{$_.Tab.ColorIndex = 36}
    ($Sheet6,$Sheet7).ForEach{$_.Tab.ColorIndex = 37}
      $WorkSheetNames = New-Object System.Collections.ArrayList  #---Sort by sheetName-
      foreach($WorkSheet in $Excel.Workbooks[1].Worksheets) { $null = $WorkSheetNames.add($WorkSheet.Name) }
        $null = $WorkSheetNames.Sort()
        For ($i=0; $i -lt $WorkSheetNames.Count-1; $i++){ ($Excel.Workbooks[1].Worksheets.Item($WorkSheetNames[$i])).Move($Excel.Workbooks[1].Worksheets.Item($i+1)) }
    $Sheet1.Activate()
    $Excel.Workbooks[1].SaveAs($ScriptPath+'\'+$TimeStamp+'-'+$OutTitle1,51)
    Remove-Item "$ScriptPath\$TimeStamp-$OutTitle1.csv"
    $Excel.visible = $true
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
      # Stop-process -Name Excel 
} else {
	Write-Host 'No LogParser CSV found. Please confirm evtx contain event 1644.' -ForegroundColor Red
}
