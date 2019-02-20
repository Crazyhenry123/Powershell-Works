#if (condition) {
#    
#}
#param 
#(   
#[string]$alias
#)


$con = new-object "System.data.sqlclient.SQLconnection"
#Set Connection String
$con.ConnectionString =(“Data Source=shmsdsql.database.windows.net;Initial Catalog=CaseFinder;User ID=msdadmin;Password=PasSw0rd01;Connect Timeout=30;Encrypt=True;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False”)




$con.open()


$adapter = New-Object System.Data.sqlclient.sqlDataAdapter $sqlcmd
$dataset = New-Object System.Data.DataSet

$sqlcmd = new-object "System.data.sqlclient.sqlcommand"
$sqlcmd.connection = $con
$sqlcmd.CommandTimeout = 600000
$sqlcmd.CommandText = “SELECT Alias FROM dbo.[Engineer] WHERE Manager = 'chuchiu'”


$DataSet = New-Object System.Data.DataSet

$adapter.SelectCommand = $sqlcmd

$adapter.Fill($DataSet)

$con.Close()

Write-Host "Engineers list to be scanned: "
Write-Output $dataset.Tables.Alias

$queryagentid = "$alias@microsoft.com"
$assembly = [Reflection.Assembly]::LoadFile("C:\Users\denl\Desktop\IRMonitor\FederatedAuthentication.dll")
$date = Get-Date -Format d
$Encoding = New-Object System.Text.UTF8Encoding
$token = "xxxx"
$AuthResult = [StefanG.MSTools.FedAuth]::Authenticate_for_ServiceDesk("https://servicedesk.microsoft.com", [ref]$token, $null)


Write-Output "Token string is $token"

foreach ($alias in $DataSet.Tables.alias)

{

Write-Host "Scanning cases for owner $alias"

 Function CreateEventSource ()
 {
 
 $EventSourceExist = [System.Diagnostics.EventLog]::SourceExists("IrMonitorEvents")

 if ($EventSourceExist -eq 0 )

 {
 
 New-EventLog -LogName Application -Source "IrMonitorEvents"
 
 }

 }

 $queryid = "CASE_QUERY"
 
 Add-Type -AssemblyName System.Net.Http
 $handler = New-Object  System.Net.Http.HttpClientHandler

 $handler.CookieContainer = New-Object System.Net.CookieContainer
 $handler.AllowAutoRedirect = $true;

 $HTTPClient= New-Object System.Net.Http.HttpClient($handler);

 $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
 $HttpClient.DefaultRequestHeaders.Accept.Add($Header)

 $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("text/plain")
 $HttpClient.DefaultRequestHeaders.Accept.Add($Header)

 $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("*/*")
 $HttpClient.DefaultRequestHeaders.Accept.Add($Header)

 $header = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer",$token)

 $HTTPClient.DefaultRequestHeaders.Authorization = $Header
 
 $testcontent = @"
 {"queryid":"CASE_QUERY","query_parameters":[{"name":"caseType","values":["Case"]},{"name":"state","values":["Open"]},{"name":"selectedAgentId","values":["$alias@microsoft.com"]},{"name":"agentId","values":["$alias@microsoft.com"]}]}
"@

 $RequestContent = New-Object System.Net.Http.StringContent($testcontent,$Encoding,"application/json")

 # $HTTPResponses = Invoke-WebRequest -Authentication Bearer -ContentType "application/json;charset=UTF-8"  -uri "https://api.support.microsoft.com/v0/queryidresult" -Method Post -Body $body -Token $token

 $HTTPResponses = $HTTPClient.PostAsync("https://api.support.microsoft.com/v0/queryidresult",$RequestContent).Result


#}

$ResponsefromSD = $HTTPResponses.Content.ReadAsStringAsync().Result 

$ResponsefromSD = ConvertFrom-Json $ResponsefromSD


$caseproperties = $ResponsefromSD.table_parameters.header_names
$caseinfoall = $ResponsefromSD.table_parameters[0].table_parameter_result

# }


[hashtable]$CaseTableAll = @{}


for ( $k=0; $k -lt $caseinfoall.count; $k= $k+1)
{

$caseinfo = $caseinfoall[$k]
[hashtable]$CaseTable = @{}

for ( $i=0; $i -lt 44; $i=$i+1)

{

$CaseTable.add($caseproperties[$i],$caseinfo[$i])

}

Write-Host "Scanning case: "
Write-Output  $CaseTable.CaseNumber
 
if (($casetable.SlaState -eq "Pending") -and ($casetable.SlaCompletedOn -eq ""))
{

[String]$CaseNumberString = $CaseTable.CaseNumber
[string]$caseTitle = $CaseTable.Title
[string]$SLAEndTime = $CaseTable.SlaExpiresOn

$ErrorMessage = @”
$CaseNumberString ($alias) <br><br> https://servicedesk.microsoft.com/#/customer/case/$CaseNumberString <br><br>
“@

#$ErrorMessage 

Write-EventLog -LogName Application -Source "IrMonitorEvents" -EntryType Error -EventId 555 -Message $ErrorMessage 

$LAbody = @"
{"CaseId":"$CaseNumberString","OwnerAlias":"$alias","Owneraddress":"$alias@microsoft.com","SLAEnd":"$SLAEndTime"}
"@

$URL = 'https://prod-09.eastasia.logic.azure.com:443/workflows/e5ad4cd3a88f4db580f6946a7d16bedf/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2vpEv5ND1m50lVrJCVku3rmF090rWLP5x_8FC69xHIY'

Invoke-WebRequest $URL -Body $LAbody -Method 'POST' -ContentType "application/json"

[hashtable]$CaseTable = @{}
}

}

}


