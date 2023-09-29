# MS365AbsenceInformator
This script is a great tool for every SMB which uses MS365 suite. This Powershell script connect to MS365 apps and send to managers information which users from their department has email forwarding.




## Requirements

This script for proper working requires several things:

- PowerShell in version 5.1 (Not Core Edition!!!)
- Active MS365 account 
- An email adress from which messeges would be send (for example gmail)
- Installed 2 powershell Modules (AzureAD, ExchangeOnlineManagement)
- Enabled script execution in PowerShell

## Installation
Set execution policy in powershell to allow cript execution

Use these commands to install required modules

```PowerShell
  Install-Module -Name AzureAD
  Install-Module -Name ExchangeOnlineManagement
```
    
Next go to Google account > security> 2FA > Application Password and generate one

In script folder create JSON file like this:

```JSON
{
    "MS365Username": "username@mycompany.com",
    "MS365Password": "password_to_ms365",

    "DispatchMail": "example@gmail.com",
    "DispatchPassword": "generated_application_password"
}
```
Save it as: CredsConfig.json
