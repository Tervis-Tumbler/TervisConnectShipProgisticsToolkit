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
    $Nodes | Install-TervisConnectShipProgistics
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

function Copy-TervisConnectShipConfigurationFiles {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $Domain = Get-ADDomain
        $SourceRootPath = "\\$($Domain.DNSRoot)\Applications\GitRepository\ConnectShip\Progistics"
        $ProgisticsPath = "C:\Program Files (x86)\ConnectShip\Progistics"
        $XMLFileProviderPath = "$ProgisticsPath\AdditionalComponents\DocumentProviders\ConnectShip\XMLFileProvider"
        $ServerPath = "$ProgisticsPath\XML_Processor\Server"
        $AMPServicePath = "$ProgisticsPath\AMP\AMPService"
        $CustomScriptingPath = "$ProgisticsPath\AMP\CustomScripting"
    }
    process {
        $XMLFileProviderRemotePath = $XMLFileProviderPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $ServerRemotePath = $ServerPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $AMPServiceRemotePath = $AMPServicePath | ConvertTo-RemotePath -ComputerName $ComputerName
        $CustomScriptingRemotePath = $CustomScriptingPath | ConvertTo-RemotePath -ComputerName $ComputerName
        Copy-Item -Path $SourceRootPath\XMLFileProvider -Destination $XMLFileProviderRemotePath -Recurse -Force 
        Copy-Item -Path $SourceRootPath\Server -Destination $ServerRemotePath -Recurse -Force
        Copy-Item -Path $SourceRootPath\AMPService -Destination $AMPServiceRemotePath -Recurse -Force        
        # New-Item -ItemType Directory -Path $CustomScriptingRemotePath -Force | Out-Null
        Copy-Item -Path $SourceRootPath\CustomScripting -Destination $CustomScriptingRemotePath -Recurse -Force
    }
}
function Install-TervisConnectShipProgistics {
    param (
        [parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $ADDomain = Get-ADDomain
        $ProgisticsPackageFilePath = "\\$($ADDomain.DNSRoot)\applications\Chocolatey\progistics.6.5.nupkg"
    }
    process {
        $TervisConnectShipDataPathOnNode = $TervisConnectShipDataPathLocal | 
        ConvertTo-RemotePath -ComputerName $ComputerName
        if (-not (Test-Path -Path $TervisConnectShipDataPathOnNode\progistics.6.5.nupkg)) {            
            Copy-Item -Path $ProgisticsPackageFilePath -Destination $TervisConnectShipDataPathOnNode
        }
        Install-TervisChocolateyPackage -ComputerName $ComputerName -PackageName Progistics -Version 6.5 -PackageParameters "$TervisConnectShipDataPathLocal\INST.ini" -Source $TervisConnectShipDataPathLocal
    }    
}