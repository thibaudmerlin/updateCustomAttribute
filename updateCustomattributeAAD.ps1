<#
    .SYNOPSIS
    updateCustomAttribute.ps1
	
    .DESCRIPTION
    Update Custom Device Attribute in Entra ID for all devices sync from AD matching deviceId

    .LINK  
    www.github.com/thibaudmerlin

    .NOTES
    Written by: Thibaud MERLIN
    Website:    www.kyos.ch

    .LICENSE
    The MIT License (MIT)
    Copyright (c) 2024 MERLIN THIBAUD
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    .CHANGELOG
    V1.00, 20/03/2024 - Initial version

    .REQUIRED
    The script requires AD powershell module
#>

#Region Variables
$client = "Company"
$logPath = "$ENV:ProgramData\$client\Logs"
$logDate = Get-Date -Format "dd-MM-yyyy"
$logFile = "$logPath\updateCustomAttribute-$logDate.log"

$appId = "" # Application ID
$tenantId = "" # Tenant ID
$secret = "" # Secret

$organizationalUnit = "OU=Devices,DC=company,DC=local"
$extensionAttribute1value = "Custom Value"
#EndRegion

#region logging
if (!(Test-Path -Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}
Start-Transcript -Path $logFile -Force
#endregion

#region Functions
function Get-AuthHeader {
    param (
        [Parameter(mandatory = $true)]
        [string]$tenant_id,
        [Parameter(mandatory = $true)]
        [string]$client_id,
        [Parameter(mandatory = $true)]
        [string]$client_secret,
        [Parameter(mandatory = $true)]
        [string]$resource_url
    )
    $body = @{
        resource      = $resource_url
        client_id     = $client_id
        client_secret = $client_secret
        grant_type    = "client_credentials"
        scope         = "openid"
    }
    try {
        $response = Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$tenant_id/oauth2/token" -Body $body -ErrorAction Stop
        $headers = @{ }
        $headers.Add("Authorization", "Bearer " + $response.access_token)
        return $headers
    }
    catch {
        Write-Error $_.Exception
    }
}
Function Get-JsonFromGraph {
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]    
        $token,
        [Parameter(Mandatory = $true)]
        $strQuery,
        [parameter(mandatory = $true)] [ValidateSet('v1.0', 'beta')]
        $ver

    )
    #proxy pass-thru
    $webClient = new-object System.Net.WebClient
    $webClient.Headers.Add("user-agent", "PowerShell Script")
    $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    try { 
        $header = $token
        if ($header) {
            #create the URL
            $url = "https://graph.microsoft.com/$ver/$strQuery"
        
            #Invoke the Restful call and display content.
            Write-Verbose $url
            $query = Invoke-RestMethod -Method Get -Headers $header -Uri $url -ErrorAction STOP
            if ($query) {
                if ($query.value) {
                    #multiple results returned. handle it
                    $query = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/$ver/$strQuery" -Headers $header
                    $result = @()
                    while ($query.'@odata.nextLink') {
                        Write-Verbose "$($query.value.Count) objects returned from Graph"
                        $result += $query.value
                        Write-Verbose "$($result.count) objects in result array"
                        $query = Invoke-RestMethod -Method Get -Uri $query.'@odata.nextLink' -Headers $header
                    }
                    $result += $query.value
                    Write-Verbose "$($query.value.Count) objects returned from Graph"
                    Write-Verbose "$($result.count) objects in result array"
                    return $result
                }
                else {
                    #single result returned. handle it.
                    $query = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/$ver/$strQuery" -Headers $header
                    return $query
                }
            }
            else {
                $errorMsg = @{
                    errNumber = 404
                    errMsg    = "No results found. Either there literally is nothing there or your query was malformed."
                }
            }
            throw;
        }
        else {
            $errorMsg = @{
                errNumber = 401
                errMsg    = "Authentication Failed during attempt to create Auth header."
            }
            throw;
        }
    }
    catch {
        return $errorMsg
    }
}

Function Update-DeviceInGraph {
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]    
        $token,
        [Parameter(Mandatory = $true)]
        $deviceId,
        [Parameter(Mandatory = $true)]
        $json,
        [parameter(mandatory = $true)] [ValidateSet('v1.0', 'beta')]
        $ver
    )

    # Create the URL
    $url = "https://graph.microsoft.com/$ver/devices/$deviceId"

    try { 
        # Invoke the Restful call to update the device
        $response = Invoke-WebRequest -Method Patch -Headers $token -Uri $url -Body $json -ContentType "application/json" -ErrorAction Stop
        return $response
    }
    catch {
        Write-Error $_.Exception
    }
}

function Get-ComputersInOU {
    param(
        [Parameter(Mandatory=$true)]
        [string] $OU
    )

    # Import the Active Directory module if not already loaded
    if (!(Get-Module -Name ActiveDirectory)) {
        Import-Module ActiveDirectory
    }

    # Get all computers in the specified OU
    $computers = Get-ADComputer -Filter * -SearchBase $OU

    # Return the array of computers
    return $computers
}

function Compare-Devices {
    param(
        [Parameter(Mandatory=$true)]
        [array] $adComputers,

        [Parameter(Mandatory=$true)]
        [array] $entraDevices,

        [Parameter(Mandatory=$true)]
        [string] $extensionAttribute1value
    )

    # Initialize an array to hold the matching devices
    $matchingDevices = @()

    # Loop through each on-premises computer
    foreach ($device in $entraDevices) {
        if ($device.extensionAttributes.extensionAttribute1 -ne $extensionAttribute1value) {
            # Loop through each Azure AD device
            foreach ($computer in $adComputers) {
                # Convert the ObjectGUID to a format that matches the deviceId in Azure AD
                $guid = $computer.ObjectGUID.Guid
                # If the deviceId matches the ObjectGUID, add the device to the matching devices array
                if ($device.deviceId -eq $guid) {
                    $matchingDevices += $device
                }
            }
        }
    }

    # Return the array of matching devices
    return $matchingDevices
}
#endregion

#region process
$params = @{
    tenant_id = $tenantId
    client_id = $appId
    client_secret = $secret
    resource_url = "https://graph.microsoft.com"
}
$token = Get-AuthHeader @params

$adComputers = Get-ComputersInOU -OU $organizationalUnit
$entraDevices = Get-JsonFromGraph -token $token -strQuery "devices?`$filter=onPremisesSyncEnabled eq true" -ver "v1.0"
try {
    $matchingDevices = Compare-Devices -adComputers $adComputers -entraDevices $entraDevices -extensionAttribute1value $extensionAttribute1value
    foreach ($device in $matchingDevices) {
        $json = @{
            extensionAttributes = @{
                extensionAttribute1 = $extensionAttribute1value
            }
        } | ConvertTo-Json
        $response = Update-DeviceInGraph -token $token -deviceId $device.id -json $json -ver "v1.0"
        if ($response.StatusCode -eq 204) {
            Write-Output "Success, the device $($device.displayName) has been updated with extensionAttribute1 value: $extensionAttribute1value"
        } else {
            Write-Output "Failed to update device $($device.displayName) with status code: $($response.StatusCode)"
        }
    }
    Stop-Transcript
}
catch {
    Write-Error $_.Exception
    Stop-Transcript
}
#endregion
