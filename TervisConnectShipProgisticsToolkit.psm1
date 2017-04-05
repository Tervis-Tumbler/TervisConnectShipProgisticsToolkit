function Invoke-ProgisticsProvision {
    param (
        $EnvironmentName
    )
    Invoke-ClusterApplicationProvision -ClusterApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisClusterApplicationNode -ClusterApplicationName Progistics -EnvironmentName $EnvironmentName
    $Nodes | Install-WCSPrintersForBartenderCommander
    $Nodes | Add-WCSODBCDSN -ODBCDSNTemplateName Tervis
    $Nodes | Add-WCSODBCDSN -ODBCDSNTemplateName tervisBartender
    $Nodes | Set-TervisBTLMIniFile
    $Nodes | Set-TervisCommanderTaskList
    $Nodes | New-BartenderCommanderFirewallRules
}
