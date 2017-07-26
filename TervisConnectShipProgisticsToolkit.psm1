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
    $Script:EnvironmentState | where EnvironmentName -eq $EnvironmentName
}

function Invoke-ProgisticsProvision {
    param (
        $EnvironmentName
    )
    Invoke-ApplicationProvision -ApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisApplicationNode -ApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes | Add-WCSODBCDSN -ODBCDSNTemplateName Tervis
    $Nodes | Set-TervisConnectShipToolkitResponseFile
    $Nodes | Install-TervisConnectShipProgistics
    $Nodes | Set-TervisConnectShipProgisticsLicense
    $Nodes | Copy-TervisConnectShipConfigurationFiles
    $Nodes | Set-ProgisticsSmartPostSettings
    $Nodes | Install-TervisConnectShipProgisticsScheduledTasks
    $Nodes | Set-SQLTCPEnabled -InstanceName CSI_Data -Architecture x86
    $Nodes | Set-SQLTCPIPAllTcpPort -InstanceName CSI_Data -Architecture x86
    $Nodes | New-SQLNetFirewallRule
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
        $ConnectShipComponentsPath = "$ProgisticsPath\AdditionalComponents\DocumentProviders\ConnectShip"
        $XMLProcessorPath = "$ProgisticsPath\XML_Processor"
        $AMPPath = "$ProgisticsPath\AMP"
    }
    process {
        $ConnectShipComponentsRemotePath = $ConnectShipComponentsPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $XMLProcessorRemotePath = $XMLProcessorPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $AMPRemotePath = $AMPPath | ConvertTo-RemotePath -ComputerName $ComputerName        
        Copy-Item -Path $SourceRootPath\XMLFileProvider -Destination $ConnectShipComponentsRemotePath -Recurse -Force 
        Copy-Item -Path $SourceRootPath\Server -Destination $XMLProcessorRemotePath -Recurse -Force
        Copy-Item -Path $SourceRootPath\AMPService -Destination $AMPRemotePath -Recurse -Force                
        Copy-Item -Path $SourceRootPath\CustomScripting -Destination $AMPRemotePath -Recurse -Force
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

function Install-TervisConnectShipProgisticsScheduledTasks {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$EnvironmentName
    )
    begin {
        $SystemCredential = New-Object System.Management.Automation.PSCredential ('System',(New-Object System.Security.SecureString))
    }
    process {
        $DomainName = Get-DomainName -ComputerName $ComputerName
        $Execute = "\\WCSJavaApplication.$EnvironmentName.$DomainName\QcSoftware\Bin\transapi_cleanup.vbs"
        Install-TervisScheduledTask -Credential $SystemCredential -TaskName TransAPI_Cleanup -Execute $Execute -RepetitionIntervalName EveryDayAt2am -ComputerName $ComputerName
    }
}

function Get-TervisConnectShipProgisticsControllerConfigurationData {
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        $Database
    )
    process {
        $QueryResponse = Invoke-SQL -dataSource $ComputerName -database $Database -sqlCommand @"
SELECT MAX(MSN)
    FROM packageList

SELECT MAX(bundleId)
    FROM packageList

SELECT MAX(pkgListId)
    FROM packageList

SELECT MAX(groupId)
    FROM groupList

SELECT MAX(shipperId)
    FROM packageList
"@
        [PSCustomObject][Ordered]@{
            MSN = $QueryResponse[0].Column1
            bundleId = $QueryResponse[1].Column1
            pkgListId = $QueryResponse[2].Column1
            groupId = $QueryResponse[3].Column1
            shipperId = $QueryResponse[4].Column1
        }
    }
}

function Set-TervisConnectShipProgisticsControllerConfigurationDataMSN {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ComputerName,
        [Parameter(Mandatory)]$MSN
    )
    process {
        Invoke-SQL -dataSource $ComputerName -database CSI_CONTROLLER_CONFIG -sqlCommand @"
UPDATE [CSI_CONTROLLER_CONFIG].[dbo].[controller_sequences]
SET sequenceValue = $MSN
WHERE sequenceId = 1
"@
    }
}

function Set-ProgisticsMSNToHigherThanWCSSybaseConnectShipMSNPreviouslyUsed {
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$EnvironmentName
    )
    process {
        $MSNMax = Get-WCSSQLConnectShipShipmentMSNMax -EnvironmentName $EnvironmentName
        Set-TervisConnectShipProgisticsControllerConfigurationDataMSN -ComputerName $ComputerName -MSN $MSNMax
    }
}

function Import-ProgisticsManagementModule {
    param (
        $ComputerName,
        $Prefix
    )
    $Session = New-PSSession -ComputerName $ComputerName
    Invoke-Command -Session $Session -ScriptBlock { ipmo -force "C:\Program Files (x86)\ConnectShip\Progistics\bin\Progistics.Management.dll" }

    if ($Prefix) {
        Import-module (Import-PSSession -Session $Session -Module Progistics.Management -DisableNameChecking -AllowClobber) -Global -Prefix $Prefix
    } else {
        Import-module (Import-PSSession -Session $Session -Module Progistics.Management -DisableNameChecking -AllowClobber) -Global
    }
}

function Compare-ProgisticsConfiguration {
    param (
        [Parameter(Mandatory)]$ReferenceComputerName,
        [Parameter(Mandatory)]$DifferenceComputerName
    )
    Import-ProgisticsManagementModule -ComputerName $ReferenceComputerName -Prefix Reference
    Import-ProgisticsManagementModule -ComputerName $DifferenceComputerName -Prefix Difference
    
    $ProgisticsObjectNounsForSimpleComparison = @"
Carrier
Country
EventHandler
License
ServerComponent
Service
Shipper
TransAPIField
Update
Version
"@ -split "`r`n"
    $ProgisticsObjectNounsForSimpleComparison | Compare-ProgisticsObject

$ProgisticsObjectNounsForShipperComparison = @"
ShipperConfiguration
ShipperControl
"@ -split "`r`n"
 
    $ReferenceShipper = Get-ReferenceShipper
    foreach ($Shipper in $ReferenceShipper) {
        foreach ($CarrierSymbol in $Shipper.RegisteredCarriers) {
            $ProgisticsObjectNounsForShipperComparison  | 
            Compare-ProgisticsObjectRequiringShipperSymbolAndCarrierSymbol -ShipperSymbol $Shipper.ShipperSymbol -CarrierSymbol $CarrierSymbol
        }
    }

#NameAddressCandidate
}

function Compare-ProgisticsObjectRequiringShipperSymbolAndCarrierSymbol {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Noun,
        [Parameter(Mandatory)]$ShipperSymbol,
        [Parameter(Mandatory)]$CarrierSymbol,
        [Switch]$IncludeEqual
    )
    process {
        $Reference = Invoke-Expression "Get-Reference$Noun -ShipperSymbol $ShipperSymbol -CarrierSymbol $CarrierSymbol"
        $Difference = Invoke-Expression "Get-Difference$Noun -ShipperSymbol $ShipperSymbol -CarrierSymbol $CarrierSymbol"
        Compare-Object -ReferenceObject $Reference -DifferenceObject $Difference -IncludeEqual:$IncludeEqual -Property Symbol, Data
    }

}

function Compare-ProgisticsObject {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Noun,
        [Switch]$IncludeEqual
    )
    process {
        $Reference = Invoke-Expression "Get-Reference$Noun"
        $Difference = Invoke-Expression "Get-Difference$Noun"
        Compare-Object -ReferenceObject $Reference -DifferenceObject $Difference -IncludeEqual:$IncludeEqual
    }
}

function Set-ProgisticsSmartPostSettings {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        Import-ProgisticsManagementModule -ComputerName $ComputerName
        Set-ShipperConfiguration -ShipperSymbol tervis -CarrierSymbol TANDATA_FEDEXFSMS.fedex -Field SP_ENABLED -Value $true
        $FedexSmartPostAccountNumber = (Get-PasswordstateCredential -PasswordID 4347).GetNetworkCredential().Password
        Set-ShipperConfiguration -ShipperSymbol tervis -CarrierSymbol TANDATA_FEDEXFSMS.fedex -Field SP_ACCOUNT -Value $FedexSmartPostAccountNumber
        Set-ShipperConfiguration -ShipperSymbol tervis -CarrierSymbol TANDATA_FEDEXFSMS.fedex -Field SP_CARRIER -Value 1
    }
}
