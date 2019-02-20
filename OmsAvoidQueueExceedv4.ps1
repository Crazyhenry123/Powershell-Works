param 
(
    [string]$WSID,
    [string]$FCPath
 )

Set-Location $FCPath


function TestFireWall()
{
.\FirewallChecker.exe /workspaceid:$WSID
}

function Test-EventLog {
    Param(
        [Parameter(Mandatory=$true)]
        [string] $SourceName
    )

    [System.Diagnostics.EventLog]::SourceExists($SourceName)
}

$EventSource = Test-EventLog "OmsTestConnectionScripts"
$TestResult= TestFireWall
$Matches = Select-String -InputObject $TestResult -Pattern "SUCCESS" -AllMatches
$HSS = get-service healthservice |ft status |out-string
$HSS = ($HSS -split '\n')[3]
# echo $HSS

If ($HSS -like "Run*")
{
$HSS = 1
}
else
{
$HSS = 0
}

if ( $EventSource -eq 0 )
{
New-EventLog -LogName Application -Source "OmsTestConnectionScripts"
}

# echo $Matches.Matches.Count
# echo $HSS

Switch ($HSS)
{
1
{
If ($Matches.Matches.Count -ne 3)
{
Stop-Service HealthService
Write-EventLog -LogName Application -Source "OmsTestConnectionScripts" -EntryType Error -EventId 3 -Message "OMS Connection Stopped. Health Service has been stopped."
}
}
0
{
If ($Matches.Matches.Count -eq 3)
{
Start-Service HealthService
Write-EventLog -LogName Application -Source "OmsTestConnectionScripts" -EntryType Information -EventId 1 -Message "OMS Connection Resumed. Health Service has been started."
}
else 
{
Write-EventLog -LogName Application -Source "OmsTestConnectionScripts" -EntryType Warning -EventId 2 -Message "OMS Connection is still unavailable. Health Service was stopped."
}
}
}





