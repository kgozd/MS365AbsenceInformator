# Install and import the AzureAD module
# Install-Module -Name AzureAD
# Install-Module -Name ExchangeOnlineManagement
using namespace System.Collections.Generic

# Import-Module AzureAD
# Import-Module ExchangeOnlineManagement

# Passwords and Logins are in this json for development,( please use more secure method of handling credentials in production env)
$ConfigData = Get-Content -Path ".\CredsConfig.json" | ConvertFrom-Json

function connect_ms_apps {
    param ($ConfigData)

    $Username = $ConfigData.MS365Username
    $Password = ConvertTo-SecureString $ConfigData.MS365Password -AsPlainText -Force
    $UserCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)

    Connect-AzureAD -Credential $UserCredential
    Connect-ExchangeOnline -Credential $UserCredential
}

function close_ms_sessions {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-AzureAD
}

#This function retrieves data(list of users with fixed mail forwarding) from MS Azure AD and MS Exchange Online
function get_user_data {
    # Retrieve users with assigned licenses (active accounts)
    $Users = Get-AzureADUser -All $true | Where-Object { $_.AssignedLicenses.Count -gt 0 -and $_.AccountEnabled -eq $true }

    # Create a table to store the user information
    $UserInformation = @()

    # Retrieve information about user managers and forwarding addresses using ExO
    foreach ($User in $Users) {
        $Manager = Get-AzureADUserManager -ObjectId $User.ObjectId
        
        $ForwardingAddress = Get-Mailbox -Identity $User.UserPrincipalName | 
        Select-Object -ExpandProperty ForwardingSmtpAddress
    
        if ($Manager -and $ForwardingAddress) {
            $ManagerEmail = Get-Mailbox -Identity $Manager.UserPrincipalName |
            Select-Object -ExpandProperty PrimarySmtpAddress
            $ForwardingAddress = $ForwardingAddress.Replace("smtp:", "").trim()
            try { 
                $ForwardingUserName = Get-Mailbox -Identity  $ForwardingAddress -ErrorAction Stop
            }
            catch {
                $ForwardingUserName = "No UserName"
            }
        
            $UserInfo = @{
                UserDisplayName    = $User.DisplayName
                UserMail           = $User.UserPrincipalName
                ManagerDisplayName = $Manager.DisplayName
                ManagerEmail       = $ManagerEmail
                ForwardingUserName = $ForwardingUserName
                ForwardingAddress  = $ForwardingAddress
            }

            $UserInformation += $UserInfo
        }
       
    }
    return $UserInformation
}

# Funaction creates email structure for every manager
function send_emails {
    param ($ConfigData, $UserInformation, $ListOfManagers)
    # Creating email with data from get_user_data and get_sorted_managers 
    foreach ( $Manager in $ListOfManagers) {
        $EmailHTML = "<html><body><h2 style='font-size: 17px;'>Below is a list of employees from your department who are currently absent (have email forwarding):</h2>
            <table style='border-collapse: collapse; border: 2px solid black; margin: 10px; font-size: 14px; width: 80%;' cellpadding='10'><tr>
            <th style='border: 1px solid black; font-weight: bold;'>UserDisplayName</th><th style='border: 1px solid black; font-weight: bold;'>UserMail</th>
            <th style='border: 1px solid black; font-weight: bold;'>Forwarded To</th>
            <th style='border: 1px solid black; font-weight: bold;'>ForwardingUserName</th><th style='border: 1px solid black; font-weight: bold;'>ForwardingAddress</th></tr>"

        # Adding specific data to certain manager
        $UserInformation | ForEach-Object {
            $UserInfo = $_
            if ($Manager -eq $($UserInfo.ManagerEmail)) {
                $EmailHTML += "<tr><td style='border: 1px solid black;'>$($UserInfo.UserDisplayName)</td>
                <td style='border: 1px solid black;'>$($UserInfo.UserMail)</td>
                <td style='border: 1px solid black;'><b>---------------></b></td>
                <td style='border: 1px solid black;'>$($UserInfo.ForwardingUserName)</td>
                <td style='border: 1px solid black;'>$($UserInfo.ForwardingAddress)</td></tr>"
            }
        }
        # End of creating emailL
        $EmailHTML += "</table></body></html>"
        $EmailHTML += "<p style='color: red; font-weight: bold;'>This message has been generated automatically; please do not reply to this email!</p>"



        # Smtp server config; this below is for gmail 
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = 587

        # credentials for email(remember for $Password variable to generate an "application secret" on gmail otherwise it wouldn't work)
        $Username = $ConfigData.DispatchMail
        $Password = ConvertTo-SecureString $ConfigData.DispatchPassword -AsPlainText -Force
        $SMTPCredential = New-Object System.Management.Automation.PSCredential ($Username, $Password)

        
        # DispatchMail is an email which will be using for automatic mail sending to every manager on list
        $EmailTo = $Manager
        $EmailFrom = $ConfigData.DispatchMail
        $EmailSubject = "List of absent employees"

        Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -BodyAsHtml $EmailHTML -SmtpServer $SMTPServer `
            -Port $SMTPPort -UseSsl -Credential $SMTPCredential -Encoding UTF8

    }
    
}

# this function creates a list with managers with absent employees
function get_sorted_managers {
    param ($UserInformation)
    
    $ManagersEmailsList = New-Object List[string]
    $UserInformation | ForEach-Object {
        $UserInfo = $_
        $ManagerEmail = $UserInfo.ManagerEmail
    
        # Sprawdzamy, czy ManagerEmail nie istnieje już na liście
        if (-not $ManagersEmailsList.Contains($ManagerEmail)) {
            $ManagersEmailsList.Add($ManagerEmail)
        }
    }
    
    return $ManagersEmailsList
}


function Main {
    param ($ConfigData)
    &connect_ms_apps  -ConfigData $ConfigData
    $UserInformation = &get_user_data
    
    &close_ms_sessions

    $ManagersEmailsList = get_sorted_managers -UserInformation $UserInformation
    
    &send_emails  -ConfigData $ConfigData -UserInformation $UserInformation -ListOfManagers $ManagersEmailsList



}

&Main -ConfigData $ConfigData
