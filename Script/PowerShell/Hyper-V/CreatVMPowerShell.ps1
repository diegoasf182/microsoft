Set-ExecutionPolicy Unrestricted
Import-Module Hyper-V
Write-Host "Variaveis de Entrada"
$VMNAME = Read-Host "Digite o Nome da VM"
$VMLocation = "C:\Users\Public\Documents\Hyper-V\Virtual hard disks\"
$VMVLAN = Read-Host "Digite o Numero da VLAN"
$VMVHD = "$VMLocation\$VMNAME\${VMNAME}_OS_C_Drive.VHDX"
$VHTEMPLATE = "C:\Users\Public\Documents\Hyper-V\Virtual hard disks\TEMPLATES\WINDOWS SERVER 2012 R2 STANDARD - USAR SYSPREP.vhdx"

[int]$VMProcessor = Read-Host "Digite a quantidade Processador Max 4"


    if ([int]$VMProcessor -gt "4") 

        {
        
            Do
             {
                Write-Warning "Não é possível criar essa quantidade"
                [int]$VMProcessor = Read-Host "Digite a quantidade Processador Max 4"

              } While([int]$VMProcessor -gt "4")
            
        }

[int64]$VMMEMORY=$null
#[string]
$VMMEMORY = Read-Host "Digite a quantidade de Memória em GB"

 

if ($VMMEMORY -gt "16")

   {
        
        Do
          {
              Write-Warning "Não é possível criar essa quantidade"
              $VMMEMORY = Read-Host "Digite a quantidade de Memória em GB Max 16GB"
    
          } While($VMMEMORY -gt "16")

    }

    $VMEMORYbytes=$VMMEMORY*1024*1024*1024

echo "Create VM Folder"
MD $VMLocation -ErrorAction SilentlyContinue
 
echo "Criando VM"
New-VM -Name "$VMNAME" -Path $VMLocation -MemoryStartupBytes $VMEMORYbytes -SwitchName "Vswitch01" -Generation 2

echo "Adicionando VLAN"
Set-VMNetworkAdapterVlan -VMName $VMNAME -Access -VlanId $VMVLAN

echo "Adicionando Processador"
Set-VMProcessor $VMNAME -Count $VMProcessor

echo "Copiando VHD"
Convert-VHD -path $VHTEMPLATE -DestinationPath $VMVHD
Add-VMHardDiskDrive -VMName $VMNAME -ControllerType SCSI -ControllerNumber 0 -ControllerLocation 0 -Path $VMVHD

Echo "Setando VHD Boot"
$vhd = Get-VMHardDiskDrive -VMName $VMNAME
Set-VMFirmware -VMName $VMNAME -FirstBootDevice $vhd

Echo "Iniciando VM"
Start-VM $VMNAME


