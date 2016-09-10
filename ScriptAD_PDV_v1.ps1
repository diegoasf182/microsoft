Set-ExecutionPolicy Unrestricted
Import-Module ActiveDirectory

################################################################################

Write-Host "Variaveis de Entrada"
[int]$NumberLOJA = Read-Host "Digite o Numero da Loja com 3 digitos Ex(050)"
$NameLoja = Read-Host "Digite o Nome da Loja Ex:(Anália Franco)"

########################## Concatena Numero da Loja + Nome ######################

if ($NumberLOJA -lt 100){
                            [string]$NumberLOJA="0"+$NumberLOJA
                        }
$FullnameLoja = "$NumberLOJA - $NameLoja"


##### Caminho OUs ###############################################################

$PathOUBASELOJAS = "OU=SARAIVA,OU=LOJAS,OU=Varejo,DC=SARAIVA,DC=CORP"
$PathOU = "OU=$FullnameLoja,OU=SARAIVA,OU=LOJAS,OU=Varejo,DC=SARAIVA,DC=CORP"

#### User ######################################################################
$userpdv = "l"+$NumberLOJA+"pdv"
$userfunction = "l"+$NumberLOJA+"pdv"
$GetCNUserFunction = Get-DistinguishedName $userfunction
#### Funcoes ####################################################################


Function CheckOU {
     param ($ou)
     
     $script:OUpathpdv = $ou
     $retorno = $false
     $GetOUpdv = Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$ou'"    
     if ($GetOUpdv -eq $null) {
                 $retorno = $false
                 Write-Host -ForegroundColor Red "$ou OU não Existe." 
}      else {
                 $retorno = $true
                 Write-Host -ForegroundColor Green "$ou OU já Existe!"
      }
      return $retorno
}
Function CheckUser {
     param ($user)
     
     $script:Userpdv = $user
     $retorno = $false
     $GetUserpdv = Get-ADUser -Filter "name -eq '$user'"    
     if ($GetUserpdv -eq $null) {
                 $retorno = $false
                 Write-Host -ForegroundColor Red "$user OU não Existe." 
}      else {
                 $retorno = $true
                 Write-Host -ForegroundColor Green "$user OU já Existe!"
      }
      return $retorno
}
Function CheckUserOU {
     param ($userOU)
     
     $script:Userpdv = $userOU
     $retorno = $false
     $GetUserpdv = "$GetCNUserFunction -eq '$userOU'"    
     if ($GetUserpdv -eq $null) {
                 $retorno = $false
                 Write-Host -ForegroundColor Red "$userOU O usuario não está na OU correta." 
}      else {
                 $retorno = $true
                 Write-Host -ForegroundColor Green "$userOU O usuario esta na OU correta!"
      }
      return $retorno
}
Function Get-DistinguishedName ($strUserName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$strUserName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 

#### Verifica existencia da OU Principal LOJA ###################################

$OUStatusraiz = &CheckOU($PathOU)

if ($OUStatusraiz -eq $true) {
    Write-Host " A OU Default da loja $FullnameLoja já Existe"

#### Verfica a existencia da OU Usuarios ########################################

    $OUpathusers = "OU=Usuarios,$PathOU"
    $OUStatususers = &CheckOU($OUpathusers)
    if ($OUStatususers -eq $true) {
        Write-Host "A OU Usuarios da loja $FullnameLoja já existe"
    } else {
    Write-Host "$OUStatususers OU Não Existe"
    $CreateOUUsers = Read-Host "Deseja Criar a OU Usuarios para a Loja $FullnameLoja 1-Sim ou 2-Nao?"
    if($CreateOUUsers -eq 1) {
    New-ADOrganizationalUnit  -Name Usuarios -Path "$PathOU"
    Write-Host "OU Usuarios Criada com Sucesso!"
       }
      }          
#### Verifique a existencia da OU PDV dentro de Usuarios ########################

    $OUpathpdv = "OU=PDV,OU=Usuarios,$PathOU"
    $OUpathpdv = &CheckOU($OUpathpdv)                                                                                 
    if ($OUpathpdv -eq $true) {
        Write-Host "Já existe OU PDV"
    } else {
    Write-Host "$OUpathpdv OU Não Existe"
    $CreateOUUserspdv = Read-Host "Deseja Criar a OU PDV para a Loja $FullnameLoja 1-Sim ou 2-Nao?"
    if($CreateOUUserspdv -eq 1) {
    New-ADOrganizationalUnit  -Name PDV -Path "OU=Usuarios,$PathOU"
    Write-Host "OU PDV Criada com Sucesso!"
       }
      }
     

#### Verifica a existencia da OU Desktops ########################

    $OUpathdesks = "OU=Desktops,$PathOU"
    $OUpathdesks = &CheckOU($OUpathdesks)
    if ($OUpathdesks -eq $true) {
        Write-Host "A OU Desktops da loja $FullnameLoja já existe"
    } else {
    Write-Host "$OUpathdesks OU Não Existe"
    $CreateOUDesks = Read-Host "Deseja Criar a OU Desktops para a Loja $FullnameLoja 1-Sim ou 2-Nao?"
    if($CreateOUDesks -eq 1) {
    New-ADOrganizationalUnit  -Name Desktops -Path "$PathOU"
    Write-Host "OU Desktops Criada com Sucesso!"
       }
      }          
#### Verifique a existencia da OU PDV dentro de Desktops ########################

    $OUpathpdvdesk = "OU=PDV,OU=Desktops,$PathOU"
    $OUpathpdvdesk = &CheckOU($OUpathpdvdesk)                                                                                 
    if ($OUpathpdvdesk -eq $true) {
        Write-Host "Já existe OU PDV"
    } else {
    Write-Host "$OUpathpdvdesk OU Não Existe"
    $CreateOUUserspdvdesk = Read-Host "Deseja Criar a OU PDV para a Loja $FullnameLoja 1-Sim ou 2-Nao?"
    if($CreateOUUserspdvdesk -eq 1) {
    New-ADOrganizationalUnit  -Name PDV -Path "OU=Desktops,$PathOU"
    Write-Host "OU PDV Criada com Sucesso!"
       }
      }
#### Cria OU Default Loja ###############################################

} else{
           Write-Host "$OUpathpdv OU Não Existe"
    $CreateOULOJA = Read-Host "Deseja Criar a OU Default para a Loja $FullnameLoja 1-Sim ou 2-Nao?"
    
    if($CreateOULOJA -eq 1) {
    New-ADOrganizationalUnit  -Name $FullnameLoja -Path "$PathOUBASELOJAS"
    New-ADOrganizationalUnit  -Name Usuarios      -Path "OU=$FullnameLoja,$PathOUBASELOJAS"
    New-ADOrganizationalUnit  -Name Desktops      -Path "OU=$FullnameLoja,$PathOUBASELOJAS"
    New-ADOrganizationalUnit  -Name PDV           -Path "OU=Usuarios,OU=$FullnameLoja,$PathOUBASELOJAS"
    New-ADOrganizationalUnit  -Name PDV           -Path "OU=Desktops,OU=$FullnameLoja,$PathOUBASELOJAS"
    Write-Host "OU PDV Criada com Sucesso" 
 }
}
#### Verifica se Existe o Usuario ###############################################

$CheckUserpdv = $userpdv
$CheckUserpdv = &CheckUser($CheckUserpdv)

if ($CheckUserpdv -eq $true) {
    Write-Host " O Usuario $userpdv da loja $FullnameLoja já Existe"

#### Ajusta O usuario PDV para Não Expirar a Senha ##############################
Set-aduser $userpdv -PasswordNeverExpires $true
Write-Host "A senha do $userpdv não Expira"
#### Permite a Troca de senha para sem senha #####################################
Set-aduser $userpdv -PasswordNotRequired $true
Write-Host "O usuario aceita a senha em branco"
#### Troca a senha do usuario para senha em Branco ###############################
$emptypwd = Read-Host -AsSecureString "Insira a Senha em Branco"
Set-ADAccountPassword -reset $userpdv -NewPassword $emptypwd

#### Verifica se o usuario esta na OU correta PDV ################################
$CheckUserOU = "CN=$userpdv,OU=PDV,OU=Usuarios,OU=$PathOU"
$CheckUserOU = &CheckUserOU($CheckUserOU)

if ($CheckUserOU -eq $true) {
    Write-Host "O Usuario $userpdv da Loja $FullnameLoja está na OU Correta"

#### Move O user PDV para OU Correta #############################################

}else{
    $userpdv2 = "l"+$NumberLOJA+"pdv"
    $MoveUserspdv = Read-Host " Deseja Mover o $userpdv2 para OU correta 1-SIM 2-NAO"
    if($MoveUserspdv -eq 1) { 
     $GetCNUser = Get-DistinguishedName $userpdv2 
     Move-ADObject -Identity $GetCNUser -TargetPath "OU=PDV,OU=Usuarios,$PathOU"
    Write-Host "Usuario movido com Sucesso!"
       }
}
### Cria o user PDV ##############################################################
}else{
        Write-Host " O Usuario $userpdv da loja $FullnameLoja não Existe"

}
