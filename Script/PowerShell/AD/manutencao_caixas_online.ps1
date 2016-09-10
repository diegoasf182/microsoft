import-module activedirectory

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

# Lista usuarios que foram modificadas ha mais de um mes 

$today=get-date
$datacorte=$today.AddMonths(-1)
$hoje=get-date -Format 'yyyy.MM.dd-hh.mm'

start-transcript "\\srvhsmbox01\pst\logs\limpeza-mailbox-usuarios-desativados-busca-$hoje.log"
$lista_usuarios = get-aduser -filter 'Enabled -eq $false -and (Employeeid -ne "Admin" -or employeeid -ne "Geral")' -Properties LastLogonDate,EmailAddress,employeeid,displayname

$contacts=get-mailcontact

foreach ($contact in $contacts) { 
	if ($contact.Hiddenfromaddresslistsenabled -eq $false) { 
		$alias=$contact.alias;  
		$user=get-aduser -filter 'samaccountname -eq $alias' ; 
		if ($user.enabled -eq $false) {
			write-host $user.SamAccountName $user.enabled $contact.name $contact.alias
			set-mailcontact $alias -HiddenFromAddressListsEnabled $true
		}
	}
}
stop-transcript

start-transcript "\\srvhsmbox01\pst\logs\limpeza-mailbox-usuarios-mailbox-desativadas-$hoje.log"
$lista_usuarios = get-aduser -filter 'Enabled -eq $false -and (Employeeid -ne "Admin" -or employeeid -ne "Geral") -and Lastlogondate -lt $datacorte' -Properties LastLogonDate,EmailAddress,employeeid,displayname

write-host -Separator ';' "SamAccountName" "Mailbox" "Server" "Database" "Displayname" "Deleted Item Size" "Total Item Size"
foreach ($user in $lista_usuarios) {
    #
    # Ajusta Displayname para evitar que tenha espacos no inicio ou no final
    #
    $displayname=$user.displayname
    $displaynamecorrigido=$displayname.trim()
    if ($displayname -ne $displaynamecorrigido) {
        set-aduser $user.samaccountname -displayname $displaynamecorrigido
    }
    if ($user.emailaddress -ne $null) { 
		$samaccountname=$user.samaccountname
        $mailbox=Get-Mailbox -filter 'alias -eq "$samaccountname"' # -identity $user.emailaddress
        if ($mailbox -ne $null) { 
#	    Set-Mailbox -Identity $user.emailaddress -HiddenFromAddressListsEnabled $true
            $mailboxstats=Get-MailboxStatistics $mailbox
            write-host -Separator ';' $user.samaccountname $user.emailaddress $mailbox.servername $mailboxstats.Displayname $mailboxstats.database $mailboxstats.totaldeleteditemsize $mailboxstats.totalitemsize
            
	    disable-mailbox -identity $user.emailaddress -confirm:$false
        }
    } 
}
stop-transcript
Remove-PSSession $Session