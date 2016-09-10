Import-Module ActiveDirectory

$odpnet="C:\oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"
If (Test-Path $odpnet){
  # // File exists
  Add-Type -Path $odpnet  
} else {
  $odpnet="E:\APPSAR\Oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"
  If (Test-Path $odpnet){
        # // File exists
        Add-Type -Path $odpnet
  }
}
$connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(Host=172.27.0.168)(Port=1521)))(CONNECT_DATA=(SID=FPW)));User ID=interportal_fpw;Password=interportal_fpw"
$connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)

function Capitalize {
    param($nome)

    $a=$null
    $x=$null
    $b=$null
    $temp=$nome.trim();
    $a=$nome.split(" ");

    foreach ($x in $a) {
        switch ($x) {
            "DA" { $b=$b+" da";break}
            "DE" { $b=$b+" de"; break}
            "DO" { $b=$b+" do"; break}
            "DI" { $b=$b+" di"; break}
            "DOS" { $b=$b+" dos"; break}
            "DAS" { $b=$b+" das"; break}
            "VP" { if ($b -eq $null) { $b="VP"} else {$b=$b+" VP"}; break}
            "TI" { if ($b -eq $null) { $b="TI"} else {$b=$b+" TI"}; break}
            "COM" { if ($b -eq $null) { $b="com"} else {$b=$b+" com"}; break}
            "II" { if ($b -eq $null) { $b="II"} else {$b=$b+" II"}; break}
            "III" { if ($b -eq $null) { $b="III"} else {$b=$b+" III"}; break}
            "PMO" { if ($b -eq $null) { $b="PMO"} else {$b=$b+" PMO"}; break}
            "AMS" { if ($b -eq $null) { $b="AMS"} else {$b=$b+" AMS"}; break}
            "SP" { if ($b -eq $null) { $b="SP"} else {$b=$b+" SP"}; break}
            "RJ" { if ($b -eq $null) { $b="RJ"} else {$b=$b+" RJ"}; break}
            "MG" { if ($b -eq $null) { $b="MG"} else {$b=$b+" Mg"}; break}
            "BP" { if ($b -eq $null) { $b="BP"} else {$b=$b+" BP"}; break}
            "UX" { if ($b -eq $null) { $b="UX"} else {$b=$b+" UX"}; break}
            "D&D" { if ($b -eq $null) { $b="D&D"} else {$b=$b+" D&D"}; break}
            "S.A." { if ($b -eq $null) { $b="S.A."} else {$b=$b+" S.A."}; break}
            "S.A" { if ($b -eq $null) { $b="S.A."} else {$b=$b+" S.A."}; break}
            "S/A" { if ($b -eq $null) { $b="S/A"} else {$b=$b+" S/A"}; break}

            default { 
                if ($x.length -ge 2) {
                    if ($b -eq $null) {
                        $b = $x.substring(0,1).toupper()+$x.substring(1).tolower()
                    } else {
                        $b = $b+" "+$x.substring(0,1).toupper()+$x.substring(1).tolower()
                    }
                }
            }
        }
    }
    if ($b -ne $null) {
        $b=$b.trim()
    }
    return $b
}

function FormataCC {
    param($cc)

# 30012201239998 - Length 14 
# 3.001.2.201.239.998
#
# 3.643.2.200.205.998
# 3.703.2.200.205.998
# 1.001.2.200.222
# 1.001.2.200

    $x=$cc.trim();
    if ($x.length -eq 14) {
        $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)+'.'+$x.substring(4,1)+'.'+$x.substring(5,3)+'.'+$x.substring(8,3)+'.'+$x.substring(11,3)
    } else {
            if ($x.length -eq 11) {
                $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)+'.'+$x.substring(4,1)+'.'+$x.substring(5,3)+'.'+$x.substring(8,3)+'.994'
            } else {
                if ($x.length -eq 8) {
                    $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)+'.'+$x.substring(4,1)+'.'+$x.substring(5,3)+'.998.994'
                } else {
                    $b = $x
                }
            }
    }
    if ($x.length -eq 15) {
        $b = $x.substring(0,2)+'.'+$x.substring(2,3)+'.'+$x.substring(5,1)+'.'+$x.substring(6,3)+'.'+$x.substring(9,3)+'.'+$x.substring(12,3)
    } else {
            if ($x.length -eq 12) {
                $b = $x.substring(0,2)+'.'+$x.substring(2,3)+'.'+$x.substring(5,1)+'.'+$x.substring(6,3)+'.'+$x.substring(9,3)+'.994'
            } else {
                if ($x.length -eq 9) {
                    $b = $x.substring(0,2)+'.'+$x.substring(2,3)+'.'+$x.substring(5,1)+'.'+$x.substring(6,3)+'.998.994'
                } else {
                    $b = $x
                }
            }
    }

    return $b
}

function FormataCPF {
    param($cpf)

# 222.597.318.05
# 31750156890
#  3798432503

    $x=$cpf.trim();
    while ($x.length -lt 11) {
        $x='0'+$x
    }
    if ($x.length -eq 11) {
        $b = $x.substring(0,3)+'.'+$x.substring(3,3)+'.'+$x.substring(6,3)+'-'+$x.substring(9,2)
        $b = $x.substring(0,3)+'.***.'+$x.substring(6,3)+'-'+$x.substring(9,2)
    } else {
        $b = $x
    }
    return $b
}

function Hash {
    param($cc)

# 30012201239998 - Length 14 
# 3.001.2.201.239.998
#
# 3.643.2.200.205.998

    $x=$cc.trim();
    if ($x.length -eq 14) {
        $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)
    }
    if ($x.length -eq 11) {
        $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)
    }
    if ($x.length -eq 8) {
        $b = '0'+$x.substring(0,1)+'.'+$x.substring(1,3)
    }
    if ($x.length -eq 15) {
        $b = $x.substring(0,2)+'.'+$x.substring(2,3)
    }
    if ($x.length -eq 12) {
        $b = $x.substring(0,2)+'.'+$x.substring(2,3)
    }
    if ($x.length -eq 9) {
        $b = $x.substring(0,2)+'.'+$x.substring(2,3)
    }

    return $b
}

function FormataData {
    param($dt)

    if ($dt.length -eq 8) {
        $b=$dt.substring(6,2)+'/'+$dt.substring(4,2)+'/'+$dt.substring(0,4)
    } else {
        $b = $dt
    }
    return $b
}

function ConsultaFPWMatricula {
    param($matricula)

 #   if ($matricula -gt 0 -and $matricula -lt 1000000) {
        $queryString = "SELECT a.id, a.cod_empresa, a.desc_empresa, a.cpf, a.nr_matricula,
            a.nom_funcionario, a.centrocusto, a.desc_ccusto,
            a.grp_hierarquico, a.desc_ghierarquico, a.mat_sup_direto,
            a.sup_direto, a.cod_situacao, a.des_situacao, a.cod_cargo,
            a.des_cargo, a.dt_admissao, a.dt_nascimento, a.qt_horasmensal,
            a.email1, a.email2, a.cod_lotacao, a.desc_lotacao, a.end_lotacao,
            a.bairro_lotacao, a.cep_lotacao, a.ddd_lotacao, a.tel_lotacao,
            a.dt_deslig, a.dt_modif, a.user_modif, a.st_modif
            FROM inter_ad a where a.nr_matricula = '$matricula' " # and a.dt_deslig is null
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText=$querystring
        $result = $command.ExecuteReader()
        $table = new-object “System.Data.DataTable”
        $table.Load($result)
        $connection.Close()
        $retorno=@{ } 
        $contador=0
        foreach ($row in $table.rows) { 
            $matricula =  $row."NR_MATRICULA".ToString()
            $nome = Capitalize $row."NOM_FUNCIONARIO".ToString()
            $empresa = Capitalize $row."DESC_EMPRESA".ToString()
            $cargo = Capitalize $row."DES_CARGO".ToString()
            $matriculamanager = $row."MAT_SUP_DIRETO".ToString()
            $manager = Capitalize $row."SUP_DIRETO".ToString()
            $situacao = $row."DES_SITUACAO".ToString()
            $departamento = Capitalize $row."DESC_GHIERARQUICO".ToString()
            $cpf = FormataCPF $row."CPF".ToString()
            $centrocusto = FormataCC $row."CENTROCUSTO".ToString()
            $hash = Hash $row."CENTROCUSTO".ToString()
            $admissao = FormataData $row."DT_ADMISSAO".ToString()
            $datanascimento = FormataData $row."DT_NASCIMENTO".ToString()
            $recisao = $row."DT_DESLIG".ToString()

            $retorno[$contador]=@{"matricula"=$matricula;
                     "nome"=$nome;
                     "empresa"=$empresa;
                     "cargo"=$cargo;
                     "matriculamanager"=$matriculamanager;
                     "manager"=$manager;
                     "situacao"=$situacao;
                     "departamento"=$departamento;
                     "cpf"=$cpf;
                     "centrocusto"=$centrocusto;
                     "hash"=$hash;
                     "admissao"=$admissao;
                     "nascimento"=$datanascimento
                     "recisao"=$recisao }
            $contador=$contador+1
        }
        $table=$null
        return $retorno
}

function RemoveAcentos {
    param($string)

    $string=$string -replace "[á,ã,à,Á,À,Ã,â,Â]", "a"
    $string=$string -replace "[é,É,ê,Ê]", "e"
    $string=$string -replace "[í,Í]", "i"
    $string=$string -replace "[ó,Ó,õ,Õ,ô,Ô]", "o"
    $string=$string -replace "[ú,Ú]", "u"
    $string=$string -replace "[ç,Ç]", "c"
    
    return $string
}

$enderecos = @{ }

$codenderecos='\\srvhsmbox01\pst\csv\cod_x_enderecos.csv'

If (Test-Path $codenderecos){
  # // File exists
    write-host "Utilizando informacoes de $codenderecos"
} else {
    $codenderecos='C:\Users\joao.galdino\Documents\powershell\csv\cod_x_enderecos.csv'
    If (Test-Path $codenderecos){
        # // File exists
        write-host "Utilizando informacoes de $codenderecos"
    }
}

If (Test-Path $codenderecos){
Import-Csv $codenderecos -Delimiter:';' -Encoding Default | ForEach-Object {
    $empresa = $_."Empresa"
    $local = $_."Local"
    $descricao = $_."Descrição"
    $end = $_."Endereço"
    $cid = $_."Cidade"
    $est = $_."UF"
    $cepcode = $_."CEP"

    if ($empresa.Length -lt 2) {
        $empresa='0'+$empresa
    }
    while ($local.length -lt 3) {
        $local='0'+$local
    }
    $hash=$empresa+'.'+$local
    $enderecos.Add("$hash",$_)
}
}
write-host -Separator ';' "Matricula" "Nome" "Cargo" "Departamento" "Centro de Custo" "MatriculaSUP" "Manager" "Situacao" "NomeAD" "CargoAD" "DepartamentoAD" "ManagerAD" "ManagerDN" "Avaliacao"

$logon=Read-Host "Informe o logon do usuario a ser ajustado"
$users = get-aduser -Properties memberof,EmailAddress,Displayname,Employeeid,Manager,Title,Department,info,Company,PostalCode,Office,StreetAddress,State,City,OfficePhone,Country $logon

foreach ($user in $users) {
    $usuario=$null
    $hash=$null
    $matricula =  $user.employeeid
    $dadosfpw=$null
    $nome=RemoveAcentos($user.Name)
    $nomepattern=$nome.Replace('.','*')
    $nomepattern=$nomepattern.Replace(' ','*')
    $nomepattern=$nomepattern+'*'
    if ($matricula -ne "Terceiro" -and $matricula -ne "Geral" -and $matricula -ne "Admin" -and $matricula -notmatch "Erica*" ) {
        $dadosfpw = ConsultaFPWMatricula($matricula)
        if ($dadosfpw.count -gt 1) {
           for ($i=0;$i -lt $dadosfpw.count;$i++) {
               $nomefpw=RemoveAcentos($dadosfpw[$i].nome)
               $nomefpwpattern=$nomefpw.Replace('.','*')
               $nomefpwpattern=$nomefpwpattern.Replace(' ','*')
               $nomefpwpattern=$nomefpwpattern+'*'
               if ($dadosfpw[$i].situacao -ne "TRANSFERENCIA S/ONUS P/CEDENTE" -and $dadosfpw[$i].situacao -cnotmatch "RESC" -and ($nomefpw -like $nomepattern -or $nome -like $nomefpwpattern)) {
                  $usuario=$dadosfpw[$i]                  
               }
           }
        } else {
            $usuario=$dadosfpw[0]
        }         
        if ($usuario -ne $null) { 
#           Encontrado apeas um usuario, começando a verificar se preciso atualizar algo no AD
            $nome=$usuario.nome
            $empresa=$usuario.empresa
            $cargo=$usuario.cargo
            $matriculamanager=$usuario.matriculamanager
            $manager = $usuario.manager
            $situacao = $usuario.situacao
            $departamento = $usuario.departamento
            $cpf = $usuario.cpf
            $centrocusto = $usuario.centrocusto
            $hash = $usuario.hash
            $admissao = $usuario.admissao
            $datanascimento = $usuario.nascimento
            write-host "Nome: $nome Hash: $hash Endereco: " $enderecos[$hash].Endereço "CC $centrocusto CPF $cpf"
            $info = "Matricula:$matricula`r`nCentro de Custo: $centrocusto`r`nCPF: $cpf`r`nData Nascimento:$datanascimento`r`nData Admissão: $admissao"
            $nomead=$user.name
            $information=$user.info
            $department=$user.department
            $title=$user.title  
            $displayname=$user.Displayname
            if ($displayname.Contains('-')) {
               #Ajusta DisplayName
               $temp=$displayname.split('-');
               $temp.trim() 
               set-aduser $user.SamAccountName -Displayname $temp[0]
            }
            if ($user.City -eq $null-and $enderecos[$hash].Cidade -ne $null) {
               set-aduser $user.SamAccountName -City $enderecos[$hash].Cidade
            }
            if ($user.Country -eq $null -or $user.Country -ne 'BR') {
               set-aduser  $user.SamAccountName -Country 'BR'
            }
            if ($user.StreetAddress -eq $null-and $enderecos[$hash].Endereço -ne $null) {
               set-aduser  $user.SamAccountName -StreetAddress $enderecos[$hash].Endereço
            }
            if ($user.OfficePhone -eq $null -and $enderecos[$hash].Telefone -ne $null) {
               set-aduser $user.SamAccountName -OfficePhone $enderecos[$hash].Telefone
            }
            if ($user.Office -ne $enderecos[$hash].Descrição) {
               set-aduser  $user.SamAccountName -Office $enderecos[$hash].Descrição
            }
            if ($user.Office -eq 'HS' -and $departamento.StartsWith("Equipe Loja")) {
               $temp=$departamento.Substring(7)
               set-aduser  $user.SamAccountName -Office $temp
            }
            if ($user.Office -eq 'HS' -and $departamento.StartsWith("Equipe itown")) {
               $temp=$departamento.Substring(7)
               set-aduser  $user.SamAccountName -Office $temp
            }
            if ($user.PostalCode -eq $null-and $enderecos[$hash].CEP -ne $null) {
               set-aduser  $user.SamAccountName -PostalCode $enderecos[$hash].CEP
            }
            if ($user.State -ne $enderecos[$hash].UF) {
               set-aduser  $user.SamAccountName -State $enderecos[$hash].UF
            }
            if ($user.title -ne $cargo) {
               #Ajusta Cargo
               set-aduser $user.SamAccountName -Title $cargo
            }
            if ($user.Department -ne $departamento) {
               #Ajusta Departamento
               set-aduser  $user.SamAccountName -Department $departamento
            }
            if ($user.Company -ne $empresa) {
               #Ajusta Empresa
               set-aduser  $user.SamAccountName -Company $empresa
            }
            if ($user.info -ne $info) {
               #Ajusta Informacoes
               set-aduser  $user.SamAccountName -Replace @{Info=$info}
            }
            if ($enderecos[$hash].GrupoLocal -ne "" -and $enderecos[$hash].GrupoLocal -ne $null) {
                if (($user.EmailAddress -ne $null) -and ($user.Enabled)) {
#                   write-host "Ajuste Listas;Empresa $empresa" $user.samaccountname $user.Enabled $user.EmailAddress $enderecos[$hash].GrupoLocal
                    Add-ADGroupMember -Identity $enderecos[$hash].GrupoLocal -Members $user.SamAccountName 
                }
            }
            if ($enderecos[$hash].GrupoGlobal -ne "" -and $enderecos[$hash].GrupoGlobal -ne $null) {
                if (($user.EmailAddress -ne $null) -and ($user.Enabled)) {
#                   write-host "Ajuste Listas;Empresa $empresa" $user.samaccountname $user.Enabled $user.EmailAddress $enderecos[$hash].GrupoGlobal
                    Add-ADGroupMember -Identity $enderecos[$hash].GrupoGlobal -Members $user.SamAccountName 
                }
            }
            if ($enderecos[$hash].GrupoJira -ne "" -and $enderecos[$hash].GrupoJira -ne $null) {
                if ($user.Enabled) {
					foreach ($grupo in $user.memberof) {
						if ($grupo -like "*Grupo TI - Clientes Jira*" -and $grupo -ne $enderecos[$hash].GrupoJira) {
							remove-adgroupmember -Identity $grupo -Members $user.SamAccountName -confirm:$false
						}
					}
#                   write-host "Ajuste Listas;Empresa $empresa" $user.samaccountname $user.Enabled $user.EmailAddress $enderecos[$hash].GrupoGlobal
                    Add-ADGroupMember -Identity $enderecos[$hash].GrupoJira -Members $user.SamAccountName 
                }
            }

            if ($enderecos[$hash].OU -ne "" -and $enderecos[$hash].OU -ne $null) {
                if (($user.distinguishedname -notcontains $enderecos[$hash].OU) -and ($user.enabled)) {
                    # Usuario em OU incorreta, movendo
                    get-aduser $user | Move-ADObject -TargetPath $enderecos[$hash].OU
                }
            }


            if ($user.manager -eq $null) {
                # Atualizar manager - vazio
                $managerdn=$null
                $mgrs = Get-AdUser -filter ('(Employeeid -eq $matriculamanager -or Employeenumber -eq $matriculamanager) -and Enabled -eq $true') -Properties Employeeid,Manager
                foreach ($mgr in $mgrs) {
                    $managerdn = $mgr.DistinguishedName
                    $managerad=$mgr.name
                }
                if ($managerdn -ne $null) {
                    $nomead=$user.name
                    $title=$user.title            
                    write-host -Separator ';' "$matricula" "$nome" "$cargo" "$depatamento" "$centrocusto" "$matriculamanager" "$manager" "$situacao" "$nomead" "$title" "$department" "$managerad" "$managerdn" "ajustar cargo e cadastrando gerente"
                    set-aduser  $user.SamAccountName -Manager $managerdn
                }
            } else {
                $teste = $user.manager
                if ( $teste.StartsWith('CN') ) {
                    # Manager esta configurado corretamente
                    $mgrs = Get-AdUser -filter ('(Employeeid -eq $matriculamanager -or Employeenumber -eq $matriculamanager) -and Enabled -eq $true') -Properties Employeeid,Manager
                    foreach ($mgr in $mgrs) {
                        $managerdn = $mgr.DistinguishedName
                        $managerad=$mgr.name
                    }
                    $nomead=$user.name
                    $information=$user.info
                    $department=$user.department
                    $title=$user.title               
                    if ($teste -ne $managerdn) {
                       write-host -Separator ';' "$matricula" "$nome" "$cargo" "$depatamento" "$centrocusto" "$matriculamanager" "$manager" "$situacao" "$nomead" "$title" "$department" "$managerad" "$managerdn" "gerente diferente corrigindo"
                       set-aduser -Identity $user.SamAccountName -Manager $managerdn
                    } else {
                       write-host -Separator ';' "$matricula" "$nome" "$cargo" "$depatamento" "$centrocusto" "$matriculamanager" "$manager" "$situacao" "$nomead" "$title" "$department" "$managerad" "$managerdn" "configurado ok"
                    }
                } else {
                    # Atualizar manager - vazio
                    $managerdn=$null
                    $mgrs = Get-AdUser -filter ('(Employeeid -eq $matriculamanager -or Employeenumber -eq $matriculamanager) -and Enabled -eq $true') -Properties Employeeid,Manager
                    foreach ($mgr in $mgrs) {
                        $managerdn = $mgr.DistinguishedName
                    }
                    if ($managerdn -ne $null) {
                       write-host -Separator ';' "$matricula" "$nome" "$cargo" "$depatamento" "$centrocusto" "$matriculamanager" "$manager" "$situacao" "$nomead" "$title" "$department" "$managerad" "$managerdn" "ajustar gerente mal preenchido e cargo"
                       set-aduser $user.SamAccountName -Manager $managerdn
                    }
                }
            }

         } else {
            write-host "TROUBLE - Usuario no AD nao encontrado na folha, configra se a matricula esta correta" $user.name " Matricula: " $user.employeeid
        }
    }
}