net use M: \\100.100.100.126\CARS /user:genesiis\contractor G5ntech
$path = "M:\accCARS.mdb"
$date_dateNum=Get-Date -format %d
$date_dateNum=$date_dateNum-1
$date=Get-Date -format MM/$date_dateNum/yyyy" "hh:mm:ss" "tt
$connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
$connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= $path;User Id=admin;Persist Security Info=False;Jet OLEDB:Database Password=320033"
$command = $connection.CreateCommand()

$repDate=Get-Date -format dd-MM-yyyy
$extensions=445,490,461,492,479,498,405
$Names='Silmiya','Tina','Sajani','Sithara','Srimani','Thilini','Charani'
$NameIterate=0
$NumOG=0
foreach ($extn in $Extensions) {
    
	$Query = "SELECT FORMAT(fldStart,'dd-MM-YYYY hh:mm:ss') AS DATE_TIME, 	
	IIF(fldCalledNo='0779612819' OR fldCalledNo='0716577686' OR fldCalledNo='0774562693' OR fldCalledNo='0773676797' OR fldCalledNo='0711214178' ,'DIVERT',fldCalledNo+'--OG') AS Called_NO,
			fldDuration AS CALL_DURATION FROM tblCDR where fldStart>=#$date#  and fldExtNo='$extn' ORDER BY fldStart"
	
	$CountOG = "SELECT COUNT(*)  FROM tblCDR WHERE fldStart>=#$date# and NOT(fldCalledNo='0779612819' OR fldCalledNo='0716577686' OR fldCalledNo='0774562693' OR fldCalledNo='0773676797' OR fldCalledNo='0711214178') and fldExtNo='$extn' "
	
	$QueryInc = "SELECT FORMAT(fldStart,'dd-MM-YYYY hh:mm:ss') AS DATE_TIME, fldCallerNo+'--INC' AS Caller_NO, fldDuration AS CALL_DURATION FROM tblCDRI 
	where fldStart>=#$date#  and fldExtNo='$extn' ORDER BY fldStart"        
	
    $CountInc = "SELECT COUNT(*)  FROM tblCDRI WHERE fldStart>=#$date# and fldExtNo='$extn'"	
	
    $CountDivert = "SELECT COUNT(*)  FROM tblCDR WHERE fldStart>=#$date# and (fldCalledNo='0779612819' OR fldCalledNo='0716577686' OR fldCalledNo='0774562693' OR fldCalledNo='0773676797' OR fldCalledNo='0711214178') and fldExtNo='$extn' "
		
$csv = "C:\Program Files\CallReportGenerate\table$extn.csv"
$csvInc = "C:\Program Files\CallReportGenerate\tableInc$extn.csv"
$csvOgNum= "C:\Program Files\CallReportGenerate\OG$extn.csv"
$csvIncNum= "C:\Program Files\CallReportGenerate\INC$extn.csv"
$csvDivertNum= "C:\Program Files\CallReportGenerate\DRT$extn.csv"

$command.CommandText = $Query
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$dataset.Tables[0] | export-csv  $csv -NoTypeInformation

$command.CommandText = $QueryInc
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$dataset.Tables[0] | export-csv  $csvInc  -NoTypeInformation

$command.CommandText = $CountOG
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$dataset.Tables[0] | export-csv  $csvOgNum  -NoTypeInformation 

$command.CommandText = $CountInc
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$dataset.Tables[0] | export-csv  $csvIncNum  -NoTypeInformation 
$IncNum =(Get-Content $csvIncNum)[1] 

$command.CommandText = $CountDivert
$adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
$dataset = New-Object -TypeName System.Data.DataSet
$adapter.Fill($dataset)
$dataset.Tables[0] | export-csv  $csvDivertNum  -NoTypeInformation 
$DivertNum =(Get-Content $csvDivertNum)[1] 
$IncNumTot = [int]$IncNum  + [int]$DivertNum

$report="C:\Program Files\CallReportGenerate\Report_$repDate.csv"
$Names[$NameIterate]+" ExtNo - "+ $extn| Out-File $report -Append -Encoding ASCII
" OG - "+ (Get-Content $csvOgNum)[1] | Out-File  $report -Append -Encoding ASCII
" INC - "+ $IncNumTot | Out-File  $report -Append -Encoding ASCII
" " | Out-File  $report -Append -Encoding ASCII

[System.IO.File]::ReadAllText($csv) | Out-File  $report -Append -Encoding ASCII 
(Get-Content $csvInc)[1..$IncNum]  | Out-File  $report -Append -Encoding ASCII
"- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - " | Out-File  $report -Append -Encoding ASCII
" " | Out-File  $report -Append -Encoding ASCII
$connection.Close()
$NameIterate=$NameIterate+1
$IncNumTot = $IncNum = $DivertNum = 0
}
net use M: /delete
Rename-Item Report_$repDate.csv Report_$repDate.txt

cmd.exe /c 'C:\Program Files\CallReportGenerate\mail_send.bat'



