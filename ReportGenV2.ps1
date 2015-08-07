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
foreach ($extn in $Extensions) {
    
	$Query = "SELECT FORMAT(fldStart,'dd-MM-YYYY hh:mm:ss') AS DATE_TIME, 	
	IIF(fldCalledNo='0779612819' OR fldCalledNo='0716577686' OR fldCalledNo='0774562693' OR fldCalledNo='0773676797' OR fldCalledNo='0711214178' ,'DIVERT',fldCalledNo+'--OG') AS Called_NO,
			fldDuration AS CALL_DURATION FROM tblCDR where fldStart>=#$date#  and fldExtNo='$extn' ORDER BY fldStart"
			
	$QueryInc = "SELECT FORMAT(fldStart,'dd-MM-YYYY hh:mm:ss') AS DATE_TIME, fldCallerNo+'--INC' AS Caller_NO, fldDuration AS CALL_DURATION FROM tblCDRI 
	where fldStart>=#$date#  and fldExtNo='$extn' ORDER BY fldStart"        
						
			
$csv = "C:\Program Files\CallReportGenerate\table$extn.csv"
$csvInc = "C:\Program Files\CallReportGenerate\tableInc$extn.csv"

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


$report="C:\Program Files\CallReportGenerate\Report_$repDate.csv"
$Names[$NameIterate]+" - "+ $extn| Out-File $report -Append -Encoding Unicode
#$extn | Out-File $report -Append -Encoding Unicode
[System.IO.File]::ReadAllText($csv) | Out-File  $report -Append -Encoding Unicode 
[System.IO.File]::ReadAllText($csvInc) | Out-File  $report -Append -Encoding Unicode
$connection.Close()
$NameIterate=$NameIterate+1
}
net use M: /delete
cmd.exe /c 'C:\Program Files\CallReportGenerate\mail_send.bat'


