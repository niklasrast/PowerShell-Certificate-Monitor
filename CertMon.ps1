<#
    .SYNOPSIS 
    PowerShell Certificate Monitoring

    .DESCRIPTION
    RUN:   C:\Windows\SysNative\WindowsPowershell\v1.0\PowerShell.exe -ExecutionPolicy Bypass -Command .\CertMon.ps1
    
    .ENVIRONMENT
    PowerShell 5.0
    
    .AUTHOR
    Niklas Rast
#>
# Logging
Start-Transcript -Path "$PSScriptRoot\ScriptLog.log" -Append -Force

# Settings
$ErrorActionPreference = "SilentlyContinue"

# Functions
function SendWarningMailGraph {
    param (
        [string]$title,
        [string]$customer,
        [string]$expirationdate,
        [string]$servicedeskaddress
    )

    $MailTenantID = ''
    $MailClientID = ''
    $MailClientsecret = ''
    $MailSender = ""

    #Connect to GRAPH API
    $MailtokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $MailClientID
        Client_Secret = $MailClientsecret
    }
    $MailtokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$MailTenantID/oauth2/v2.0/token" -Method POST -Body $MailtokenBody
    $Mailheaders = @{
        "Authorization" = "Bearer $($MailtokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    #Send Mail    
    $URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
$BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "WPS Certificate Monitoring",
                          "body": {
                            "contentType": "HTML",
                            "content": "Hallo,
                             <br><br>
                             the following certificate is about to expire.
                             <br><br>
                             <b>Customer</b>           = $customer
                             <br>
                             <b>Certificate name</b> = $title
                             <br>
                             <b>Expiration date</b>     = $expirationdate
                             <br><br>
                             Best regards
                             <br>
                             Modern Workplace Services ($ENV:COMPUTERNAME)
                            "
                          },
                          "toRecipients": [
                            {
                              "emailAddress": {
                                "address": "$servicedeskaddress"
                              }
                            }
                          ]
                        },
                        "saveToSentItems": "true"
                      }
"@

    Invoke-RestMethod -Method POST -Uri $URLsend -Headers $Mailheaders -Body $BodyJsonsend
    Write-Host "Function SendReportMailGraph finished." -ForegroundColor Green
}

# Expiration dates
$30DaysAway = (Get-Date).AddDays(30)
$Expiration30Days = (Get-Date -Date $30DaysAway -Format "dd.MM.yyyy")

$20DaysAway = (Get-Date).AddDays(20)
$Expiration20Days = (Get-Date -Date $20DaysAway -Format "dd.MM.yyyy")

$10DaysAway = (Get-Date).AddDays(10)
$Expiration10Days = (Get-Date -Date $10DaysAway -Format "dd.MM.yyyy")

$5DaysAway = (Get-Date).AddDays(5)
$Expiration5Days = (Get-Date -Date $5DaysAway -Format "dd.MM.yyyy")

# Read cert file
[xml]$Certificates = Get-Content "$PSScriptRoot\certificates.xml"

foreach ($item in $Certificates.Certificates.Cert) {
    if ($Expiration30Days -eq $item.ExpirationDate) {
        SendWarningMailGraph -title $item.Title -customer $item.Customer -expirationdate $item.ExpirationDate -servicedeskaddress $item.ServiceDeskAddress
    }

    if ($Expiration20Days -eq $item.ExpirationDate) {
        SendWarningMailGraph -title $item.Title -customer $item.Customer -expirationdate $item.ExpirationDate -servicedeskaddress $item.ServiceDeskAddress
    }

    if ($Expiration10Days -eq $item.ExpirationDate) {
        SendWarningMailGraph -title $item.Title -customer $item.Customer -expirationdate $item.ExpirationDate -servicedeskaddress $item.ServiceDeskAddress
    }

    if ($Expiration5Days -eq $item.ExpirationDate) {
        SendWarningMailGraph -title $item.Title -customer $item.Customer -expirationdate $item.ExpirationDate -servicedeskaddress $item.ServiceDeskAddress
    }
}

Stop-Transcript