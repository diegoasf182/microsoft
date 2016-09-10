
#Add-Type -Path "C:\oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"
Add-Type -Path "E:\APPSAR\Oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"

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

function ConsultaFPWNome {
    param($nome)

    $nome=RemoveAcentos($nome)
    $nome=$nome.ToUpper()
    $nome=$nome.Replace(' ','%')
    $nome=$nome+'%'
    write-host "Pesquisa pelo padrao $nome"

 #   if ($matricula -gt 0 -and $matricula -lt 1000000) {
        $queryString = "SELECT a.id, a.cod_empresa, a.desc_empresa, a.cpf, a.nr_matricula,
            a.nom_funcionario, a.centrocusto, a.desc_ccusto,
            a.grp_hierarquico, a.desc_ghierarquico, a.mat_sup_direto,
            a.sup_direto, a.cod_situacao, a.des_situacao, a.cod_cargo,
            a.des_cargo, a.dt_admissao, a.dt_nascimento, a.qt_horasmensal,
            a.email1, a.email2, a.cod_lotacao, a.desc_lotacao, a.end_lotacao,
            a.bairro_lotacao, a.cep_lotacao, a.ddd_lotacao, a.tel_lotacao,
            a.dt_deslig, a.dt_modif, a.user_modif, a.st_modif
            FROM inter_ad a where a.nom_funcionario like '$nome' " # and a.dt_deslig is null
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
            $situacao = Capitalize $row."DES_SITUACAO".ToString()
            $departamento = Capitalize $row."DESC_GHIERARQUICO".ToString()
            $cpf = FormataCPF $row."CPF".ToString()
            $centrocusto = FormataCC $row."CENTROCUSTO".ToString()
            $hash = Hash $row."CENTROCUSTO".ToString()
            $admissao = FormataData $row."DT_ADMISSAO".ToString()
            $datanascimento = FormataData $row."DT_NASCIMENTO".ToString()
            $recisao = $row."DT_DESLIG".ToString()
            $modificado = $row."DT_MODIF".ToString()
			$gh = $row."grp_hierarquico".ToString()
			$ghdesc = $row."desc_ghierarquico".toString()
			$lot = $row."cod_lotacao".ToString()
			$desclot = $row."desc_lotacao".ToString()
			
            write-host " "
            write-host "Matricula         : $matricula"
            write-host "Nome              : $nome"
            write-host "Empresa           : $empresa"
            write-host "Cargo             : $cargo"
            write-host "Matricula Manager : $matriculamanager"
            write-host "Manager           : $manager"
            write-host "Situacao          : $situacao"
            write-host "Departamento      : $departamento"
            write-host "CPF               : $cpf"
            write-host "Centro de Custo   : $centrocusto"
            write-host "Data Admissao     : $admissao"
            write-host "Data Nascimento   : $datanascimento"
            write-host "Data Recisao      : $recisao"
            write-host "Data Mudanca      : $modificado"
			write-host "GH                : $gh"
			write-host "GH Descricao      : $ghdesc"
			write-host "Lotação           : $lot"
			write-host "Desc. Lotação     : $desclot"
            
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

$nome = read-host "Informe o nome a ser pesquisado"
write-host "Padrao informado: $nome"

$usuario = ConsultaFPWNome($nome)
