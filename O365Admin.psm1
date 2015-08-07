#region Initialize
. "$PSScriptRoot\Init.ps1"
$CSharp = Get-Content -Path $CSharpPath -Raw
try {[O365Admin.AccountSkuId]}
catch [System.Management.Automation.RuntimeException]
{
    Add-Type -TypeDefinition $CSharp
    $Global:Error.RemoveAt(0)
}
#endregion

#region Public Functions
function Connect-O365
{
    <#
            .SYNOPSIS
            Connects to the Office 365 environment
            .DESCRIPTION
            Connects to Office 365 with options for Exchange, Skype and Sharepoint. You can also select
            AzureActiveDirectory only.
            .PARAMETER Services
            The Office 365 services you wish to connect to. Valid values are Exchange, Skype,
            and Sharepoint. To specify multiple values use a comma-separated list.
            .PARAMETER Credential
            The username or PSCredential to use to connect to Office 365 services.
            .PARAMETER SharepointUrl
            If Sharepoint is specified as an argument to -Services, you can use SharepointUrl to
            specify the URL to connect to.
            .EXAMPLE
            $Credential = Get-Credential
            Connect-O365 -Services Exchange,Skype -Credential $Credential
            .EXAMPLE
            Connect-O365 -Services Sharepoint -SharepointUrl https://contoso-admin.sharepoint.com -Credential $Credential
    #>
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory = $true)]
        [O365Services]$Services,

        [parameter(Mandatory = $true)]
        [System.Management.Automation.Credential()]
        $Credential
    )

    dynamicparam {
        if ($PSBoundParameters.Services -contains 'Sharepoint')
        {
            $ParamAttr = New-Object -TypeName System.Management.Automation.ParameterAttribute
            $ParamOptions = New-Object -TypeName System.Management.Automation.ValidatePatternAttribute `
                                       ('^https://[a-zA-Z0-9\-]+\.sharepoint\.com')
            $AttributeCollection = New-Object -TypeName 'Collections.ObjectModel.Collection[System.Attribute]'
            $AttributeCollection.Add($ParamAttr)
            $AttributeCollection.Add($ParamOptions)
            $Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter `
                                    -ArgumentList @('SharepointUrl', [string], $AttributeCollection)
            $Dictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
            $Dictionary.Add('SharepointUrl', $Parameter)
            $Dictionary
        }
    }

    begin
    {
        function Connect-O365Skype
        {
            param($Credential)
            Import-Module -Name LyncOnlineConnector -DisableNameChecking -Force
            $Option = New-PSSessionOption -IdleTimeout -1
            $SkypeSession = New-CsOnlineSession -Credential $Credential -SessionOption $Option
            $ModuleName = 'SkypeForBusiness'
            $ModulePath = "$PSScriptRoot\Bin\$ModuleName"
            $null = Export-PSSession -Session $SkypeSession -OutputModule $ModulePath -AllowClobber -Force
            Import-Module $ModulePath -Global -DisableNameChecking
        }

        function Connect-O365Sharepoint
        {
            param($Credential, $Url)
            $Params = @{
                Url = $Url
                Credential = $Credential
                WarningAction = 'SilentlyContinue'
            }
            
            Import-Module -Name Microsoft.Online.Sharepoint.Powershell -DisableNameChecking -Force
            Connect-SPOService @Params
        }

        function Connect-O365Exchange
        {
            param($Credential)
            
            $ExchParams = @{
                ConfigurationName = 'microsoft.exchange'
                ConnectionUri     = 'https://ps.outlook.com/powershell'
                Credential        = $Credential
                Authentication    = 'Basic'
                AllowRedirection  = $true
            }
            $ExchSession = New-PSSession @ExchParams
            $ModuleName = 'ExchangeOnline'
            $ModulePath = "$PSScriptRoot\Bin\$ModuleName"
            $null = Export-PSSession -Session $ExchSession -OutputModule $ModulePath -AllowClobber -Force
            Import-Module $ModulePath -Global -DisableNameChecking
        }
    }

    process
    {
        switch ($Services)
        {
            { $_.HasFlag([O365Services]::AzureActiveDirectory) -or $_.HasFlag([O365Services]::Exchange) }
            {
                Import-Module -Name MSOnline -DisableNameChecking -Force
                Connect-MsolService -Credential $Credential
            }

            { $_.HasFlag([O365Services]::Exchange) }
            { Connect-O365Exchange -Credential $Credential }

            { $_.HasFlag([O365Services]::Skype) }     
            { Connect-O365Skype -Credential $Credential }

            { $_.HasFlag([O365Services]::Sharepoint) }
            { Connect-O365Sharepoint -Credential $Credential -Url $PSBoundParameters.SharepointUrl }
        }
    }
}

function Disconnect-O365
{
    <#
            .SYNOPSIS
            Disconnects from Office 365 services and removes proxy commands and sessions
    #>

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true)]
        [O365Services]
        $Services
    )

    switch ($Services)
    {
        { $_.HasFlag([O365Services]::Exchange) }
        {
            Get-PSSession | Where-Object -Property ComputerName -Like -Value '*outlook.com' | Remove-PSSession
            Remove-Module -Name ExchangeOnline -ErrorAction SilentlyContinue
        }

        { $_.HasFlag([O365Services]::Skype) }
        {
            Get-PSSession | Where-Object -Property ComputerName -Like -Value '*online.lync.com' | Remove-PSSession
            Remove-Module -Name SkypeForBusiness -ErrorAction SilentlyContinue
        }

        { $_.HasFlag([O365Services]::Sharepoint) }
        { try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch [System.InvalidOperationException]{} }
    }
}

. "$PSScriptRoot\Licensing.ps1"

function Get-O365PrincipalGroupMembership
{
    <#
            .DESCRIPTION
            Lists distribution group membership for an Exchange Online recipient

            .PARAMETER Identity
            Specifies an Exchange Online recipient. Valid attributes are:

            Displayname
            Example: 'Matthew McNabb'

            UserPrincipalName
            Example: matt@domain.com

            DistinguishedName
            Example: 'CN=Matthew McNabb,OU=Domain.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=NAMPR07B091,DC=prod,DC=outlook,DC=com'

            Alias
            Example: matt

            .EXAMPLE
            Get-O365PrincipalGroupMembership -Identity matt

            This command retrieves the group membership for the recipient with alias 'matt'
    #>
    
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('Identity')]
        [string]
        $UserPrincipalName
    )
    
    Reconnect-O365Exchange

    $Recipient = Get-Recipient -Identity $Identity

    $Groups = Get-Group -ResultSize Unlimited -RecipientTypeDetails 'MailUniversalDistributionGroup','MailUniversalSecurityGroup'

    foreach ($Group in $Groups)
    {
        if ($Group.Members -contains $Recipient.DisplayName) { $Group.Identity }
    }
}

function Set-O365PrincipalGroupMembership
{
    <#
            .DESCRIPTION
            Configures group membership for an Exchange Online recipient.

            .PARAMETER Identity
            Specifies an Exchange Online recipient whose group membership you will modify.

            .PARAMETER MemberOf
            Specifies the group(s) that the recipient will be a member of.

            .PARAMETER Replace
            If -Replace is included with -MemberOf then any previous group membership will be removed.

            .PARAMETER Clear
            The -Clear switch parameter will remove all group membership for the recipient.

            .EXAMPLE
            Set-O365PrincipalGroupMembership -Identity matt -Memberof 'sales','IT'

            Adds the sales and IT groups to Matt's current group membership
            .EXAMPLE
            Set-O365PrincipalGroupMembership -Identity matt -Memberof 'sales','IT' -Replace

            Replaces and current group membership for Matt with the Sales and IT groups

            .EXAMPLE
            Set-O365PrincipalGroupMembership -Identity matt -Clear

            Removes all group membership for Matt
    #>
    [CmdletBinding(DefaultParameterSetName = 'MemberOf')]
    Param
    (
        # Office 365 user to modify
        [parameter(ParameterSetName='MemberOf')]
        [parameter(ParameterSetName='Clear')]
        [parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('Identity')]
        $UserPrincipalName,

        # An array of groups to add the user to. Separate group names with a comma.
        [parameter(ParameterSetName='MemberOf')]
        [ValidateNotNullorEmpty()]
        [string[]]$MemberOf,

        # If set then the user will be removed from any distribution groups not specified.
        [parameter(ParameterSetName='MemberOf')]
        [switch]$Replace,

        [parameter(ParameterSetName='Clear')]
        [switch]$Clear
    )

    Reconnect-O365Exchange

    If ($Replace -or $Clear)
    {
        Get-O365PrincipalGroupMembership  -Identity $Identity |
        ForEach-Object  -Process {
            $Params = @{
                Identity = $_
                Member = $Identity
                Confirm = $false
                BypassSecurityGroupManagerCheck = $true
            }
            Remove-DistributionGroupMember @params
        }
    }

    If ($MemberOf -eq $null) {return}

    $MemberOf |
    ForEach-Object  -Process {
        $Params = @{
                Identity = $_
                Member = $Identity
                Confirm = $false
                BypassSecurityGroupManagerCheck = $true
            }
        Add-DistributionGroupMember @Params
    }
}

function Start-O365DirSync
{
    <#
            .SYNOPSIS
            Initiates a directory import using Office 365 Azure Active Directory Sync.
            .DESCRIPTION
            Initiates a directory import using Office 365 Azure Active Directory Sync on the local or remote computer.
            Can run an incremental sync or a full import sync. If run against a remote computer, requires that PSRemoting
            is enabled.
            .PARAMETER ComputerName
            The computer on which Azure Active Directory Sync is running.
            .PARAMETER Path
            The path to the directory sync console file. You can ignore this parameter if you installed Azure Active 
            Directory Sync to the default installation folder.
            .PARAMETER Credential
            You can provide an alternate credential for the remote computer.
            .PARAMETER FullImport
            Lets Azure Active Directory Sync know to run a full import sync instead of an incremental sync.
            .EXAMPLE
            Start-O365DirSync -ComputerName DirSyncServer -Credential 'whitehouse\alincoln'
            .NOTES

            .LINK

    #>
    [CmdletBinding()]
    param
    (
        [string]
        $ComputerName,
        
        $Path = "$env:ProgramFiles\Microsoft Online Directory Sync\DirSyncConfigShell.psc1",

        [switch]
        $FullImport,

        [System.Management.Automation.CredentialAttribute()]
        $Credential
    )

    $SB = {
        if ($Using:FullImport)
        {
            $RegSplat = @{
                Path   = 'HKLM:\Software\Microsoft\MSOLCoExistence'
                Name = 'FullSyncNeeded'
                Value  = 1
            }
            Set-ItemProperty @RegSplat
        }
        
        & Powershell.exe -PsConsoleFile $Using:Path -Command 'Start-OnlineCoexistenceSync'
    }

    $CmdSplat = @{
        ComputerName = $ComputerName
        ScriptBlock = $SB
    }

    If ($Credential) {$CmdSplat.Credential = $Credential}

    Invoke-Command @CmdSplat
}

function Update-O365AccountSku
{
    param
    (
        $Parameter1
    )
    
    try
    {
        $o365AccountSkus = Get-O365AccountSku
    }
    catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]
    {
        Write-Warning "Credentials are needed to connect to Azure Active Directory."
        Write-Warning "Please enter Office 365 administrative credentials."
        $Credential = Get-Credential
        Connect-MsolService -Credential $Credential
        $O365AccountSkus = Get-O365AccountSku
    }
    
    Set-O365LicenseCodeGen -O365AccountSkus $O365AccountSkus
    Import-Module $ExecutionContext.SessionState.Module.Path
}

#endregion

#region Helpers, variables, and aliases

function Reconnect-O365Exchange
{
    <#
            .DESCRIPTION
            Checks for available connections to Exchange Online and 
            reconects if session is in a disconnected or broken state.    
    #>

    try {Get-PSSession | Test-O365ExchSessionState}
    catch
    {
        Write-Warning -Message 'The connection to Exchange Online has timed out. Reconnecting...'
        Disconnect-O365 -Services Exchange
        Connect-O365 -Services Exchange
    }
}

function Test-O365ExchSessionState
{
    <#
            .DESCRIPTION
            Checks for a current Exchange Online session. If session does
            not exist or is broken, an error is returned.
    #>
    param
    (
        [parameter(ValueFromPipeline=$true)]
        [System.Management.Automation.Runspaces.PSSession]
        $Session
    )

    begin {
        $ExchSessions = @()
    }

    process {
        if ($Session.computername -like '*.outlook.com')
        {
            $ExchSessions += $Session
            if (($Session.state -ne 'Opened') -or ($Session.availability -ne 'Available'))
            {
                throw
            }
        }
    }

    end {
        if (!$ExchSessions)
        {
            throw
        }
    }
}

Set-Alias -Name O365 -Value Connect-O365
Set-Alias -Name gogm -Value Get-O365PrincipalGroupMembership
Set-Alias -Name sogm -Value Set-O365PrincipalGroupMembership
#requires -version 3.0
#endregion

Export-ModuleMember -Function `
    'Connect-O365',
    'Disconnect-O365',
    'Get-O365UserLicense',
    'Get-O365PrincipalGroupMembership',
    'Set-O365UserLicense',
    'Set-O365PrincipalGroupMembership',
    'Start-O365Dirsync',
    'Update-O365AccountSku' -Alias * -Variable O365AccountSkus