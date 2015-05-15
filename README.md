This is a script module for Powershell 3.0 and up that adds some new commands for managing Office 365 tenants. At present it includes the following:

Functions
---------
####Connect-O365

Connect-O365 connects to Office 365 service. You can choose to connect to Exchange, Lync, or Sharepoint, Azure AD or any combination of the four.

####Disconnect-O365

Disconnects sessions from any of the above mentioned services.

####Get-O365UserLicense

Outputs O365 user license information in an easily consumable object format.

####Set-O365UserLicense

License management via Powershell is a real pain, but this function can make it quite a bit easier. Now when you load the O365Admin module for the first time, you'll be prompted to log in with Office 365 admin credentials. This is to initialize your environment and populate your available licenses. Set-O365UserLicense provides the -AccountSkuId and -ServicePlans parameters and tab-completion/Intellisense for licenses available in your tenant.

####Get-O365PrincipalGroupMembership

The ActiveDirectory module includes cmdlets for getting members of a group as well as getting the group membership of a security principal. The MSOL and Exchange Online cmdlets are more one-sided - you can only look at the members of a group. Get-O365PrincipalGroupMembership lets you see group membership from the security principal's perspective similar to how Get-ADPrincipalGroupMembership works.

####Set-O365PrincipalGroupMembership

See above. Adds or removes group membership for an Exchange Online user.

####Start-O365DirSync

A function that will initiate an Azure Active Directory Sync on the local or a remote computer.

Helper Functions
----------------
####Test-O365ExchangeSessionState
Checks for an active implicit remoting session to outlook.com

####Reconnect-O365Exchange
Works with Test-O365ExchangeSessionState to reconnect the remoting session if something went wrong.

 
Examples
--------
Connect to Exchange and Lync:

```PowerShell


$Credential = Get-Credential 
 Connect-O365 -Services Exchange,Lync -Credential $Credential
```

 Connect to Sharepoint:

```PowerShell


Connect-O365 -Services Sharepoint -SharepointUrl https://contoso-admin.sharepoint.com -Credential $Credential
```

 Disconnect from all services:

```PowerShell


Disconnect-O365 -Services Exchange,Lync,Sharepoint
```

Show available Office 365 licenses:

```PowerShell


$O365AccountSkus
```
 
Get assigned licenses for a user:

```PowerShell


Get-O365UserLicense -UserPrincipalName Abraham.Lincoln@whitehouse.gov
```
 
Assign user licenses and enabled service plans:

```PowerShell


Set-O365UserLicense -UserPrincipalName Abraham.Lincoln@whitehouse.gov -AccountSkuId ENTERPRISEPACK -ServicePlans EXCHANGE_S_STANDARD,SHAREPOINTENTERPRISE
```

###Installation Instructions

Right-click the downloaded .zip file and select Properties and then Unblock. Then use your favorite archive tool (7zip, etc.) to unpack it into your module directory. In a Powershell console or the ISE, run:

```PowerShell

Import-Module O365Admin
```

Connecting to Office 365 requires a credential in UPN (email address) format. To save a credential to disk for later use,
use Export-CliXml:

```PowerShell
Get-Credential | Export-CliXml -Path c:\creds\credential.xml
```

You can import this credential into your session later using Import-CliXml:

```PowerShell
Import-CliXml -Path c:\creds\credential.ps1
```

###Requirements:

Microsoft Online Services Sign-In Assistant

Azure Active Directory Module

Lync Online Module

Sharepoint Online Management Shell