$con = new-object "System.data.sqlclient.SQLconnection"
#Set Connection String. The string is commetted for security concern. 
$SQLServer = "tcp:shmsdsql.database.windows.net,1433"
$uid = "######"
$pwd = "######"
$con.ConnectionString =("Server=$SQLServer;Initial Catalog=CaseFinder;Persist Security Info=False;User ID=$uid;Password=$pwd;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;
")

$con2 = new-object "System.data.sqlclient.SQLconnection"
$con2.ConnectionString =("Server=$SQLServer;Initial Catalog=SDCases;Persist Security Info=False;User ID=$uid;Password=$pwd;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;
")

$con.open()

$adapter = New-Object System.Data.Sqlclient.SqlDataAdapter 
$dataset = New-Object System.Data.DataSet

$sqlcmd = new-object "System.data.sqlclient.sqlcommand"
$sqlcmd.connection = $con
$sqlcmd.CommandTimeout = 600000
$sqlcmd.CommandText = "SELECT Alias FROM dbo.[Engineer] WHERE Manager = 'chuchiu'"

$DataSet = New-Object System.Data.DataSet

$adapter.SelectCommand = $sqlcmd
$adapter.Fill($DataSet)

$con.Close()

Write-Host "Engineers list to be scanned: "
Write-Output $dataset.Tables.Alias

$queryagentid = "$alias@microsoft.com"
$assembly = [Reflection.Assembly]::LoadFile("d:\Document\Projects\IRMonitor2\IRMonitor2\FederatedAuthentication.dll")
$date = Get-Date -Format d
$token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyIsImtpZCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyJ9.eyJhdWQiOiI1NmJjMDQ0Yy03NjQ0LTRiZTYtYmM1Mi0zN2QxNzhiNWE2NDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNTUyMzgwMTE1LCJuYmYiOjE1NTIzODAxMTUsImV4cCI6MTU1MjM4NDAxNSwiYWlvIjoiQVNRQTIvOEtBQUFBYnJxb0lHMWFZU3NjUWJYY2lGWCtyTmVtd1lJT2VRSHdSWmpWeVZyQ3Awcz0iLCJhbXIiOlsid2lhIl0sImZhbWlseV9uYW1lIjoiWmhhbmciLCJnaXZlbl9uYW1lIjoiUnVpd2VuIiwiaW5fY29ycCI6InRydWUiLCJpcGFkZHIiOiIxNjcuMjIwLjI1NS42NSIsIm5hbWUiOiJSdWl3ZW4gWmhhbmciLCJub25jZSI6ImQzYzMwOGE3LWE0YWEtNGU2OC1iYjIxLTBkYmEzZjJhMmM5YSIsIm9pZCI6ImY1Y2FlNTc2LTg5MzQtNGRmNS05MzIwLTQwNjAwZGRkYTQ2OCIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMTQ2NzczMDg1LTkwMzM2MzI4NS03MTkzNDQ3MDctMjQ1ODY3MCIsInN1YiI6Ijc2Um5SMk5tcHhScjdsOGlEY0c5a3hna05zOHBrV2VOUlhweU5iazhSaXciLCJ0aWQiOiI3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJ1bmlxdWVfbmFtZSI6InQtcnVpemhAbWljcm9zb2Z0LmNvbSIsInVwbiI6InQtcnVpemhAbWljcm9zb2Z0LmNvbSIsInV0aSI6ImxQQkVCQUFFSGttMWU5VnE5LTRNQUEiLCJ2ZXIiOiIxLjAifQ.sYBOkt173-IJzx6P6hFJiFFmYFPeJcgWvPi0eqXqMrbkqEiw4Or0gJirvKWk_KiZkZy0TG8P1DKITo1qIK3rs8bQWOaaQI4f6bYhVYpXnS4ZjSyyYaMiRyvdP8MtGLDY7FsFmhjv0GwIIKrZqQak-gDNOpU5iuVR4h9UnuL9oP9j55HJIc5aTuguiCwIBf_YzRmNeDwdN3My1kbaE0xt7QHdz8PTWwfKuWraMfYnH_7l8Yu5Do1N7LM9uvjK6SyfZLdN32-oDyLnNueYJeGBvZRP8Q52_pHDbFlLophyg_pjnuBHzJi5ovhvCP-RvUQEFMUO9asV6k49nH3MHfOWVA"
$AuthResult = [StefanG.MSTools.FedAuth]::Authenticate_for_ServiceDesk("https://servicedesk.microsoft.com", [ref]$token, $null)


Write-Output "Token string is $token"

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

$adapter2 = New-Object System.Data.Sqlclient.SqlDataAdapter 
$sqlcmd2 = New-Object System.Data.Sqlclient.Sqlcommand
$sqlcmd2.connection = $con2
$sqlcmd2.CommandTimeout = 600000
$sqlcmd3 = New-Object System.Data.Sqlclient.Sqlcommand
$sqlcmd3.connection = $con2
$sqlcmd3.CommandTimeout = 600000

function GetResponsefromSD ($alias, $HTTPClient){
    $Encoding = New-Object System.Text.UTF8Encoding
    $testcontent = @"
    {"queryid":"CASE_QUERY","query_parameters":[{"name":"caseType","values":["Case"]},
    {"name":"state","values":["Open"]},{"name":"selectedAgentId","values":["$alias@microsoft.com"]},
    {"name":"agentId","values":["$alias@microsoft.com"]}]}
"@
     $RequestContent = New-Object System.Net.Http.StringContent($testcontent,$Encoding,"application/json")
     
     #$HTTPResponses = Invoke-WebRequest -Authentication Bearer -ContentType "application/json;charset=UTF-8"  -uri "https://api.support.microsoft.com/v0/queryidresult" -Method Post -Body $body -Token $token
     $HTTPResponses = $HTTPClient.PostAsync("https://api.support.microsoft.com/v0/queryidresult",$RequestContent).Result
     
     $ResponsefromSD = $HTTPResponses.Content.ReadAsStringAsync().Result 
     $ResponsefromSD = ConvertFrom-Json $ResponsefromSD

     return $ResponsefromSD
}

function GetCustomerAttribute ([String]$customerString) {
    $jSriptSer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $customerDictList = $jSriptSer.DeserializeObject($customerString)
    $customerDict = $customerDictList[0]
    if ($customerDictList.Count -eq 2) {
        $contactDict = $customerDictList[1]
    }
    else {
        $contactDict = $customerDict["Contacts"][0]
    }
    $customerDict["Contacts"] = $contactDict["id"]

    return $customerDict, $contactDict 
}

function GetTimeString($datetime){
    $datetime = [String]$datetime
    if ($datetime -eq "") {
        $datetime = "null"
    }
    else {
        $datetime = $datetime.Split(".")
        $datetime = "'" + $datetime[0]+ "'"
    }
    return $datetime 
}

function GetCaseInsertString ($caseInfo, $alias){
    $commandText = "INSERT INTO `"Case`" VALUES(" 
    for($i = 0; $i -lt $caseInfo.Count; $i += 1){
        if($i -eq 34){
            $customerDict, $contactDict = GetCustomerAttribute $caseInfo[$i]
            $customerDict.Add("CaseId", $caseinfo[0])
            $commandText += "'$($customerDict["id"])'"
        }
        elseif($caseInfo[$i] -eq ""){
            $commandText += "null"
        }
        elseif ($i -eq 0 -or $i -eq 9 -or $i -eq 14 -or $i -eq 15 -or $i -eq 25) {
            #int
            $commandText += "$($caseinfo[$i])"
        }
        elseif ($i -eq 5 -or $i -eq 6 -or $i -eq 8 -or ($i -ge 10 -and $i -le 13) -or $i -eq 16 -or $i -eq 17) {
            #datetime
            $dataTime = GetTimeString $caseinfo[$i]
            $commandText += "$dataTime"
        }
        elseif ($i -eq 34) {
            $text = [string]$caseinfo[$i].ToUpper()
            $commandText += "'$text'"
        }
        else {
            #nvarchar
            $text = [string]$caseinfo[$i]
            $text = $text.Replace("'", "''")
            $commandText += "'$text'"
        }

        if ($i -lt $caseInfo.Count - 1) {
            $commandText += ", "
        }
        else {
            $commandText += ", '$alias');"
        }
    }

    return $commandText, $customerDict, $contactDict 
}
function GetCustomerInsertString ($customerDict){
    $updatedOn = GetTimeString $customerDict["UpdatedOn"]
    $createdOn = GetTimeString $customerDict["CreatedOn"]
    $customerDict["id"] = $customerDict["id"].ToUpper()
    $customerDict["Contacts"] = $customerDict["Contacts"].ToUpper()
    $commandText = @"
    INSERT INTO Customer VALUES('$($customerDict["id"])', '$($customerDict["UpdatedBy"])',
    $updatedOn, '$($customerDict["Contacts"])', '$($customerDict["CustomerId"])',
    '$($customerDict["CustomerIdSource"])', '$($customerDict["CustomerType"])', '$($customerDict["CustomerName"])',
    '$($customerDict["CreatedBy"])', $createdOn, '$($customerDict["CaseId"])');
"@
     return $commandText
}

function GetContactInsertString ($contactDict){
    $updatedOn = GetTimeString  $contactDict["UpdatedOn"]
    $createdOn = GetTimeString  $contactDict["CreatedOn"]
    $contactDict["id"] = $contactDict["id"].ToUpper()
    $commandText = @"
    INSERT INTO Contact VALUES('$($contactDict["Email"])', '$($contactDict["FirstName"])',
    '$($contactDict["IncludeInCommunication"])', '$($contactDict["IsPrimaryContact"])', '$($contactDict["LastName"])',
    '$($contactDict["Phone"])', '$($contactDict["PreferredContactChannel"])', '$($contactDict["ContactId"])',
    '$($contactDict["ContactIdSource"])', '$($contactDict["CreatedBy"])', $createdOn, '$($contactDict["UpdatedBy"])',
    $updatedOn, '$($contactDict["id"])');
"@
     return $commandText
}

$con2.Open()
#delete all data in Case, Customer, Contact
$sqlcmd2.CommandText = "DeleteAllData"
$sqlcmd2.CommandType = [System.Data.CommandType]::StoredProcedure
$sqlcmd2.ExecuteNonQuery() | Tee-Object -Variable output | Out-Null
if ($output -eq -1) {
    $sqlcmd2.ExecuteNonQuery() | Out-Null
}
$sqlcmd2.Dispose()

#insert new cases, customers, contacts by engineer
foreach ($alias in $DataSet.Tables.alias){
    Write-Host "Scanning cases for owner $alias"
    $responseFromSD = GetResponsefromSD $alias $HTTPClient
    $caseInfoAll = $ResponsefromSD.table_parameters[0].table_parameter_result
    foreach ($caseInfo in $caseInfoAll){
        Write-Host $caseInfo[0]
        $commandTextCase, $customerDict, $contactDict = GetCaseInsertString $caseInfo $alias
        $commandTextCustomer = GetCustomerInsertString $customerDict
        $commandTextContact = GetContactInsertString $contactDict
        $sqlcmd3.CommandText = $commandTextContact
        $adapter2.InsertCommand = $sqlcmd3
        $adapter2.InsertCommand.ExecuteNonQuery() | Tee-Object -Variable output | Out-Null
        if ($output -eq -1) {
            $adapter2.InsertCommand.ExecuteNonQuery() | Out-Null
        }
        $sqlcmd3.CommandText = $commandTextCustomer
        $adapter2.InsertCommand = $sqlcmd3
        $adapter2.InsertCommand.ExecuteNonQuery() | Tee-Object -Variable output | Out-Null
        if ($output -eq -1) {
            $adapter2.InsertCommand.ExecuteNonQuery() | Out-Null
        }
        $sqlcmd3.CommandText = $commandTextCase
        $adapter2.InsertCommand = $sqlcmd3
        $adapter2.InsertCommand.ExecuteNonQuery() | Tee-Object -Variable output | Out-Null
        if ($output -eq -1) {
            $adapter2.InsertCommand.ExecuteNonQuery() | Out-Null
        }
    }
}

$sqlcmd3.Dispose()
$con2.Close()

