$FILEPATH = "C:\Users\diegoas.SARAIVA\Desktop\emailZ.csv"
$Stuff = Import-Csv $FILEPATH

$Result = ForEach ($Name in $Stuff){
$T= $Name.Email
Get-ADUser -Identity $T -Properties LastLogonDate,Enabled,Department,Description,EmployeeID,SamAccountName
}
$accountenable = $Result | Where-Object {$_.Enabled -like "true"}|Where-Object{$_.EmployeeID -notlike "1*"} |Where-Object{$_.EmployeeID -notlike "T*"} |Where-Object{$_.EmployeeID -notlike "4*"}|Where-Object{$_.EmployeeID -notlike "2*"} |Where-Object{$_.EmployeeID -notlike "7*"} | Sort Department | FT Name,Enabled,Department -AutoSize
$accountenable >> C:\accountenabledservicess.txt
<#
$accountenable >> C:\AccountEnableADepartamento.txt
$accountdisabled.count

$Result | Where-Object {$_.LastLogonDate -lt $date_with_offset}
#>