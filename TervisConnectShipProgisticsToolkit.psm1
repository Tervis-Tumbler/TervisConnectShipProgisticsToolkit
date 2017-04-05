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
    $Nodes | Set-TervisConnectShipToolkitResponseFile -ComputerName
    Foreach ($Node in $Nodes) {
        Install-TervisChocolateyPackage -ComputerName $Node.ComputerName -PackageName Progistics -Version 6.5 -PackageParameters "/RESPONSE=$TervisConnectShipDataPathLocal\INST.ini /ACCEPTEULA=YES"
    }
}

function Set-TervisConnectShipToolkitResponseFile {
    param (
        [Parameter(ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(ValueFromPipelineByPropertyName)]$EnvironmentName
    )
    begin {
        $ConnectShipWebSite = Get-PasswordstateCredential -PasswordID 2602

        $MembersAreaUser = $ConnectShipWebSite.UserName
        $MembersAreaPassword = $ConnectShipWebSite.GetNetworkCredential().password
    }
    process {
        $EnvironmentState = Get-EnvironmentState -EnvironmentName $EnvironmentName
        $ProgisticsCredential = Get-PasswordstateCredential -PasswordID $EnvironmentState.ProgisticsUserPasswordEntryID
        $ProgisticsPassword = $ProgisticsCredential.GetNetworkCredential().password

        $TervisConnectShipDataPathOnNode = $TervisConnectShipDataPathLocal | 
        ConvertTo-RemotePath -ComputerName $ComputerName

        "$ModulePath\INST.ini.pstemplate" | 
        Invoke-ProcessTemplateFile |
        Out-File -Encoding utf8 -NoNewline "$TervisConnectShipDataPathOnNode\INST.ini"
    }
}