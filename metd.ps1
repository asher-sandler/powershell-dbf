#-----------------------------------------------------
function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
[System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
#-----------------------------------------------------

function upd_kva($isodig,$nrush,$period1,$isolat,$period2,$vcurse1,$vcurse2){


$sqlConnection = New-Object System.Data.SqlClient.SqlConnection

$sqlConnection.ConnectionString = "Server=server;Database=db;Integrated Security=True"

$sqlConnection.Open()




$query = "SELECT TOP (1) ISO_LAT3,ISO_DIG FROM KVAL WHERE ISO_DIG = '$isodig'"


$SQLCommand = New-Object System.Data.SqlClient.SqlCommand($query, $sqlConnection)
$SQLReader = $SQLCommand.ExecuteScalar()


## проверяем есть значение в базе или нет
## $SQLReader
if ($SQLReader){
	$sqlstring = "update KVAL set NAME_RUSH='$nrush', ISO_LAT3 = '$isolat', PERIOD2 = CONVERT(date,'$period2',104)  where ISO_DIG = '$isodig'"
	}
else{
	
	$sqlstring = "Insert Into KVAL (ISO_LAT3,NAME_RUSH,PERIOD1,ISO_DIG,PERIOD2) values('$isolat','$nrush',convert(date,'$period1',104),'$isodig' ,convert(date,'$period2',104))"

	}

$sqlstring



$cmd = $sqlConnection.CreateCommand()

$Cmd.CommandText=$sqlstring


$cmd.ExecuteNonQuery() > arj.out


$query = "SELECT TOP 1 DAT,ISO_DIG  FROM [VALUTA].[dbo].[VALCOURSE] WHERE ISO_DIG = '$isodig' AND DAT = convert(date,'$period1',104)"

$SQLCommand = New-Object System.Data.SqlClient.SqlCommand($query, $sqlConnection)
$SQLReader = $SQLCommand.ExecuteScalar()
if ($SQLReader){
	$sqlstring = "update [VALUTA].[dbo].[VALCOURSE] set SCALE = 1, COURSE = $vcurse1, COURSE_BUY = $vcurse2  WHERE ISO_DIG = '$isodig' AND DAT = convert(date,'$period1',104)"
	}
else{
	
	$sqlstring = "Insert Into [VALUTA].[dbo].[VALCOURSE] (DAT,ISO_DIG,COURSE,COURSE_BUY,SCALE) values(convert(date,'$period1',104),'$isodig',$vcurse1,$vcurse2,1)"

	}

$sqlstring

$cmd = $sqlConnection.CreateCommand()

$Cmd.CommandText=$sqlstring


$cmd.ExecuteNonQuery() > arj.out

$sqlConnection.Close()


## Release-Ref($SQLCommand)
## Release-Ref($sqlConnection)



}


Function ParceDBF($DataPath, $TableName){


$DataPath = $DataPath.Replace(":\",":\\")
# $TableName
# read-host

$ConnString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=$DataPath;"

$dbfc = New-Object -comobject ADODB.Connection
$dbfc.Open($ConnString)
$dbfsql =  "Select KOD,DAT,QUOTE_BUY,QUOTE_SELL,Name_Rus from $TableName"


$Record = New-Object -comobject ADODB.RecordSet
$Record.Open($dbfsql,$dbfc)

$Record.MoveFirst()
# $out = $False

do {

        ## $current = "" | Select @{n='ISO_DAT';e={$Record.Fields.Item("Iso_Dig").Value}},
        ## @{n='Iso_LAT3';e={$Record.Fields.Item("Iso_LAT3").Value}},
		## @{n='DAT';e={$Record.Fields.Item("DAT").Value}},
 		## @{n='Scale';e={$Record.Fields.Item("Scale").Value}},
		## @{n='Curse';e={$Record.Fields.Item("Curse").Value}},
		## @{n='Name_Rush';e={$Record.Fields.Item("Name_Rush").Value}},
		## @{n='PR_Ecu';e={$Record.Fields.Item("PR_Ecu").Value}}
        ## $current | select DAT, Name_Rush,scale, Curse 
        # if ($out -eq $false){
        #     write-host $Record.Fields.Item("DAT").Value
        #     $out = $true
        # }    
    
        #$arr += $current
    $isodig = $Record.Fields.Item("KOD").Value;
    $nrush  = $Record.Fields.Item("Name_Rus").Value;
    $p1     = $($($Record.Fields.Item("DAT").Value).ToString()).substring(0,10); 
    $isolat = $Record.Fields.Item("KOD").Value;
    $p2     = $($($Record.Fields.Item("DAT").Value).ToString()).substring(0,10);
    $vcurse2 = $Record.Fields.Item("QUOTE_SELL").Value;
    $vcurse1 = $Record.Fields.Item("QUOTE_BUY").Value;
    # заполняем ,базу
    upd_kva $isodig $nrush $p1 $isolat $p2 $vcurse1 $vcurse2 $p1 $p2
	
        
	
        #$Record.Fields.Item("Iso_Dig").Value;
	#$Record.Fields.Item("Iso_LAT3").Value;
	#$Record.Fields.Item("DAT").Value;
	#$Record.Fields.Item("Scale").Value;
	#$Record.Fields.Item("Curse").Value;
	#$Record.Fields.Item("Name_Rush").Value;
	#$Record.Fields.Item("PR_Ecu").Value;







 	$Record.MoveNext()

} until ($Record.EOF)

$Record.Close()
$dbfc.Close()

Release-Ref($Record)
Release-Ref($dbfc)


}


$tmpdir = "F:\AdminDir\2Val";
foreach ($file in $(Get-childItem 'F:\curd\' -include MET*.ARJ -Recurse ))
{
$hostout = $file.FullName + "..."
write-host $hostout
$arjtmp=$tmpdir+"\"

get-childitem $tmpdir -include *.dbf -recurse| remove-item 
F:\AdminDir\2Val\arj.exe x -y $file.FullName $arjtmp > arj.out
foreach ($dbf in $(Get-childItem $tmpdir -include *.dbf -Recurse )){
		
		$hostout = $dbf.Fullname + ":  За дату"; 		write-host $hostout
		
		ParceDBF $dbf.DirectoryName $dbf.Name
		
		}

}

#$arr | Sort "Name_Rush" 

#$arr | ft

###   $dbfc.execute($dbfsql)