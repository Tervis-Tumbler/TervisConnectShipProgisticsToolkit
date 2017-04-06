$ModulePath = (Get-Module -ListAvailable TervisConnectShipProgisticsToolkit).ModuleBase
$TervisConnectShipDataPathLocal = "C:\ProgramData\TervisConnectShip"

$EnvironmentState = [PSCustomObject][Ordered]@{
    EnvironmentName = "Production"
    ProgisticsUserPasswordEntryID = 4108
},
[PSCustomObject][Ordered]@{
    EnvironmentName = "Epsilon"
    ProgisticsUserPasswordEntryID = 4107
},
[PSCustomObject][Ordered]@{
    EnvironmentName = "Delta"
    ProgisticsUserPasswordEntryID = 3659
}

function Get-EnvironmentState {
    param (
        [Parameter(Mandatory)]$EnvironmentName
    )
    $EnvironmentState | where EnvironmentName -eq $EnvironmentName
}

function Invoke-ProgisticsProvision {
    param (
        $EnvironmentName
    )
    Invoke-ClusterApplicationProvision -ClusterApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisClusterApplicationNode -ClusterApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes | Add-WCSODBCDSN -ODBCDSNTemplateName Tervis
    $Nodes | Set-TervisConnectShipToolkitResponseFile
    Foreach ($Node in $Nodes) {

        $TervisConnectShipDataPathOnNode = $TervisConnectShipDataPathLocal | 
        ConvertTo-RemotePath -ComputerName $Node.ComputerName

        Copy-Item -Path "\\tervis.prv\applications\Chocolatey\progistics.6.5.nupkg" -Destination $TervisConnectShipDataPathOnNode
        Install-TervisChocolateyPackage -ComputerName $Node.ComputerName -PackageName Progistics -Version 6.5 -PackageParameters "$TervisConnectShipDataPathLocal\INST.ini"
    }
    $Nodes | Set-TervisConnectShipProgisticsLicense
}

function Set-TervisConnectShipToolkitResponseFile {
    param (
        [Parameter(ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(ValueFromPipelineByPropertyName)]$EnvironmentName
    )
    begin {
        $ConnectShipWebSite = Get-PasswordstateCredential -PasswordID 2602

        $Global:MembersAreaUser = $ConnectShipWebSite.UserName
        $Global:MembersAreaPassword = $ConnectShipWebSite.GetNetworkCredential().password
    }
    process {
        $EnvironmentState = Get-EnvironmentState -EnvironmentName $EnvironmentName
        $ProgisticsCredential = Get-PasswordstateCredential -PasswordID $EnvironmentState.ProgisticsUserPasswordEntryID
        $Global:ProgisticsPassword = $ProgisticsCredential.GetNetworkCredential().password

        $TervisConnectShipDataPathOnNode = $TervisConnectShipDataPathLocal | 
        ConvertTo-RemotePath -ComputerName $ComputerName
        
        New-Item -ItemType Directory -Path $TervisConnectShipDataPathOnNode -Force | Out-Null

        "$ModulePath\INST.ini.pstemplate" | 
        Invoke-ProcessTemplateFile |
        Out-File -Encoding utf8 -NoNewline "$TervisConnectShipDataPathOnNode\INST.ini"
    }
}

function Set-TervisConnectShipProgisticsLicense {
    param (
        [parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $LicenseCred = Get-PasswordstateCredential -PasswordID 3923
    }
    process {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Import-Module "C:\Program Files (x86)\ConnectShip\Progistics\bin\Progistics.Management.dll"
            Set-License -Credentials $Using:LicenseCred
        }
    }    
}