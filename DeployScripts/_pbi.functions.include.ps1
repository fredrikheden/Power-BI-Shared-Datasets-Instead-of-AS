<#
    Created by Fredrik Hedenström 2019-08-09
#>

$baseUri = "https://api.powerbi.com/v1.0/myorg"

function GetNewPowerBIToken
{
    $LoginErrorOccurred = $False
    try 
    {
        $var = Get-PowerBIAccessToken
        Set-Variable -Name PBIAuthHeader -Scope Global -Value $var
    } 
    catch
    {
        if ( $_.Exception.Message -notlike "*Login-PowerBIServiceAccount*" )
        {
             throw  $_.Exception
        }
        $LoginErrorOccurred = $True
    }
    if ( $LoginErrorOccurred ) 
    {
        $loginInfo = Login-PowerBIServiceAccount
        $var2 = Get-PowerBIAccessToken
        Set-Variable -Name PBIAuthHeader -Scope Global -Value $var2 -Force
    }
}

function EnsureAuthHeader 
{
    if ( -not (Test-Path variable:global:PBIAuthHeader) ) 
    {
        GetNewPowerBIToken
    }
    $var = Get-Variable -Name PBIAuthHeader -Scope Global
    if ( $v.Value -eq $null ) 
    {
        GetNewPowerBIToken
    }
}

function GetPBIAuthHeader
{
    EnsureAuthHeader
    $v = Get-Variable -Name PBIAuthHeader -Scope Global
    [System.Collections.Hashtable]$vRet = $v.Value
    return $vRet
}

function Invoke-RestMethod-PBI
{
    param(
        [Parameter(Mandatory=$False)]$uriAction,
        [Parameter(Mandatory=$False)]$method="GET",
        [Parameter(Mandatory=$False)]$contentType,
        [Parameter(Mandatory=$False)]$body
    )
    $headers = GetPBIAuthHeader
    $uri = "$baseUri/$uriAction"
    $tokenErrorOccurred = $False
    try
    {
        $r = Invoke-RestMethod -Headers $headers -Uri $uri -Method $method -ContentType $contentType -Body $body
    } 
    catch
    {
        $resp = $_.Exception.Response
        if ( $resp -eq $Null )
        {
            throw $_.Exception
        }
        $reqstream = $resp.GetResponseStream()
        $reqstream.Position = 0
        $sr = New-Object System.IO.StreamReader $reqstream
        $errResp = $sr.ReadToEnd()  | ConvertFrom-Json
        if ( $errResp.error.code -ne "TokenExpired" )
        {
            throw "Error occurred in REST call: $($errResp.error)"
        } 
        $tokenErrorOccurred = $True
    }
    if ( $tokenErrorOccurred ) 
    {
        Write-Verbose "PBI token expired, getting a new one and doing the REST call again."
        GetNewPowerBIToken
        $headers = GetPBIAuthHeader
        $r = Invoke-RestMethod -Headers $headers -Uri $uri -Method $method -ContentType $contentType -Body $body
    }
    return $r
}

function GetReportsInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId
    )
    $r = Invoke-RestMethod-PBI -uriAction "groups/$groupId/reports"
    return $r.value
}

function GetReportInGroup 
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$False)]$reportName,
        [Parameter(Mandatory=$False)]$reportId
    )
    if ( $reportId -ne $null )
    {
        $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/reports/$reportId"
    } 
    else 
    {
        $r1 = GetReportsInGroup -groupId $groupId | where { $_.name -eq $reportName }
        $r = $r1[0]
    }
    return $r
}

function ReportExistsInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$reportName
    )
    $r = GetReportsInGroup -groupId $groupId | Where { $_.name -eq $reportName }
    return $r -ne $Null
}

function GetDatasetsInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId
    )
    $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets"
    return $r.value
}

function GetDatasetInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$False)]$datasetId,
        [Parameter(Mandatory=$False)]$datasetName
    )
    if ( $datasetId -ne $null )
    {
        $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets/$datasetId"
    }
    else
    {
        $r1 = GetDatasetsInGroup -groupId $groupId | where { $_.name -eq $datasetName }
        $r = $r1[0]
    }
    return $r
}

function GetDataSourcesInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$datasetId
    )
    $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets/$datasetId/datasources"
    return $r.value
}

function WaitForRefreshToFinnish
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$datasetId
    )
    for( $i=0; $i -lt 10 ) {
        write-verbose "Getting refresh status..."
        $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets/$datasetId/refreshes?$top=1"
        $status = $r.value[0].status
        if ( $status -ne "Unknown" ) 
        {
            break
        }
        write-verbose "Refresh status: $status"
        Sleep -Seconds 5
    }
    if ( $status -ne "Completed" ) {
        Write-Warning "Refresh did not complete successfully."
    }
}

function RefreshDatasetInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$datasetId,
        [Parameter(Mandatory=$True)]$waitForRefreshToFinnish
    )

    try 
    {
        $r2 = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets/$datasetId/refreshes" `
            -method Post -contentType "application/json"
    }
    catch
    {
        if ( $_.Exception.Message -like "*24 hours exceeded*" ) 
        {
            $lnk = "https://app.powerbi.com/groups/$groupId/settings/datasets/$datasetId"
            write-warning "Dataset refresh rate exceeded. Please refresh the dataset manually and press enter when completed:`r`n$lnk" 
            Pause
        }
        else 
        {
            throw $_Exception
        }
    }
    if ( $waitForRefreshToFinnish ) 
    {
        WaitForRefreshToFinnish -groupId $groupId -datasetId $datasetId
    }
}

function UpdateDatasourcesInGroup_SqlServer
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$datasetId,
        [Parameter(Mandatory=$True)]$targetServer,
        [Parameter(Mandatory=$True)]$targetDatabase
    )
    
    $currentDataSource = (GetDataSourcesInGroup  -groupId $groupId -datasetId $datasetId)[0]
    $body = @"
{
  "updateDetails": [
    {
      "datasourceSelector": {
        "datasourceType": "$($currentDataSource.dataSourceType)",
        "connectionDetails": {
          "server": "$($currentDataSource.connectionDetails.server)",
          "database": "$($currentDataSource.connectionDetails.database)"
        }
      },
      "connectionDetails": {
        "server": "$targetServer",
        "database": "$targetDatabase"
      }
    }
  ]
}
"@
    Write-Verbose "Updating datasource..."
    $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/datasets/$datasetId/Default.UpdateDatasources" `
            -method Post -body $body -contentType "application/json"

    return $r.value
} 

function WaitForImportToFinnish()
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$True)]$importId
    )
    Write-Verbose $importId
    $notFoundIds = 0
    for($i=0; $i -lt 10; $i++) 
    {
        $r = Invoke-RestMethod-PBI -UriAction "groups/$groupId/imports/$importId"
        if ( $r -eq $null ) 
        {
            # No import with that id, return.
            Write-Verbose "No import id found."
            $notFoundIds ++
            if ($notFoundIds > 5 )
            {
                return
            }

        }
        if ( $r.importState -eq "Succeeded" ) {
            # The import is succeeded, return
            Write-Verbose "Import id found, succeeded."
            return $r
        } 
        else 
        {
            Write-Verbose "Import id found with status: $($r.importState) "
        }
        sleep -Seconds 1
    }
    write-host "Error! The import did not succeed." -ForegroundColor Red
}

function PostImportInGroup
{
    param(
        [Parameter(Mandatory=$True)]$pbixFile,
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$False)]$datasetDisplayName,
        [Parameter(Mandatory=$False)]$waitToFinnish=$True,
        [Parameter(Mandatory=$False)]$nameConflict="Ignore", <# Ignore, Overwrite, CreateOrOverwrite #>
        [Parameter(Mandatory=$False)]$skipReport=$False
    )
    $fileName = [uri]::EscapeDataString([IO.Path]::GetFileName($pbixFile))

    if ( $datasetDisplayName -eq $null ) 
    {
        # Default is to use the report file name.
        $datasetDisplayName = $fileName
    }

    $boundary = [System.Guid]::NewGuid().ToString("N")   
    $fileBin = [IO.File]::ReadAllBytes($pbixFile)	      
    $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
    $fileEnc = $enc.GetString($fileBin)	
    $LF = [System.Environment]::NewLine
    $bodyLines = (
			"--$boundary",
			"Content-Disposition: form-data; name=`"file0`"; filename=`"$fileName`"; filename*=UTF-8''$fileName",
			"Content-Type: application/x-zip-compressed$LF",
			$fileEnc,
			"--$boundary--$LF"
		) -join $LF	

    $import = Invoke-RestMethod-PBI -UriAction "groups/$groupId/imports?datasetDisplayName=$datasetDisplayName&nameConflict=$nameConflict&skipReport=$skipReport" `
                -ContentType "multipart/form-data; boundary=--$boundary" -Body $bodyLines -Method Post
                
    if ( $waitToFinnish ) 
    {
        $r2 = WaitForImportToFinnish -groupId $groupId -importId $import.id
        return $r2.reports[0].id
    } 
}

function RebindReportInGroup
{
    param(
        [Parameter(Mandatory=$True)]$groupId,
        [Parameter(Mandatory=$False)]$reportId,
        [Parameter(Mandatory=$False)]$targetDatasetId
    )
    
    $body = @{datasetId=$targetDatasetId}

    $r = Invoke-RestMethod-PBI -ContentType "application/json" -Method Post -Body ($body | ConvertTo-Json) -UriAction "groups/$groupId/reports/$reportId/Rebind"

    return $r
    
}

