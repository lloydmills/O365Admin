function Get-O365UserLicense
{
    <#
        .SYNOPSIS
        Returns Office 365 licenses and enabled service plans for a
        given user
        .PARAMETER UserPrincipalName
        The full UPN for the user to report on
        .EXAMPLE
        Get-O365License -UserPrincipalName user@domain.com
        .NOTES
        Author: Matt McNabb
    #>
    [CmdletBinding()]
    param
    (
        [parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true,Mandatory=$true)]
        [string]
        $UserPrincipalName
    )

    begin
    {
        #Reconnect-O365Exchange
    }

    process
    {
        $MSOLUser = Get-MsolUser -UserPrincipalName $UserPrincipalName
        foreach ($License in $MSOLUser.Licenses)
        {
            $EnabledPlans = @()
            foreach ($ServicePlan in $License.ServiceStatus)
            {
                if ($ServicePlan.ProvisioningStatus -ne 'Disabled')
                {
                    $EnabledPlans += $ServicePlan.ServicePlan.ServiceName
                }
            }

            [PSCustomObject]@{AccountSkuId = $License.AccountSkuId; EnabledPlans = $EnabledPlans}
        }
    }
}

function Set-O365UserLicense
{
    <#
        .SYNOPSIS
        Sets licenses for Office 365 users.
        .PARAMETER UserPrincipalName
        The UPN value of the Office 365 user
        .PARAMETER AccountSkuId
        The Account SKU ID of the Office 365 plan.
        .PARAMETER ServicePlans
        The service plans to enable (if available).
        .EXAMPLE
        user@domain.com | Set-O365License -Sku Faculty -ServicePlans Exchange,Sharepoint
        .INPUTS
        [string]
        [Microsoft.Online.Administration.User]
        [Microsoft.ActiveDirectory.Management.ADUser]
        .NOTES
        Author: Matt McNabb
        Date: 6/12/2014

        To-do:
        - create -Remove switch parameter to identify licenses to remove
        - method for handling disabled plans is removing plans from the $O365AccountSkus variable - gotta fix this
    #>

    [cmdletbinding(SupportsShouldProcess = $true)]
    param
    (
        [parameter(Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [string]
        $UserPrincipalName,
        
        [Parameter(Mandatory = $true)]
        [O365Library.AccountSkuId]
        $AccountSkuId
    )

    dynamicparam
    {
        $Dictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary

        [System.Collections.ArrayList]$AvailablePlans = $O365AccountSkus |
            Where-Object AccountSkuID -eq $AccountSkuId |
            Select-Object -ExpandProperty AvailablePlans |
            Where-Object {$_ -notlike '*yammer*'}
        # This line removed the service plan from $O365AccountSkus inexplicably; weird!
        #[System.Collections.ArrayList]$Options = ($O365AccountSkus | Where-Object AccountSkuID -eq $AccountSkuId).AvailablePlans
        $ParamAttr = New-Object -TypeName System.Management.Automation.ParameterAttribute
        $ParamOptions = New-Object -TypeName System.Management.Automation.ValidateSetAttribute -ArgumentList $AvailablePlans
        $AttributeCollection = New-Object -TypeName 'Collections.ObjectModel.Collection[System.Attribute]'
        $AttributeCollection.Add($ParamAttr)
        $AttributeCollection.Add($ParamOptions)
        $Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @('ServicePlans', [string[]], $AttributeCollection)
        $Dictionary.Add('ServicePlans', $Parameter)
        $Dictionary
    }

    begin
    {
        [System.Collections.ArrayList]$DisabledPlans = $AvailablePlans
        [string]$AccountSkuId = "$AccountSkuId"
        $DomainName = $O365AccountSkus |
            Where-Object {$_.AccountSkuId -eq $AccountSkuId} |
            Select-Object -ExpandProperty DomainName
        $FullSku = "$DomainName`:$AccountSkuId"
        $Params = @{AddLicenses = $FullSku}

        if ($PSBoundParameters.ServicePlans)
        {
            foreach ($ServicePlan in $PSBoundParameters.ServicePlans)
            {
                $DisabledPlans.Remove($ServicePlan)
            }

            $LicenseOptions = @{AccountSkuId  = $FullSku; DisabledPlans = $DisabledPlans}
            $Params.Add('LicenseOptions', (New-MsolLicenseOptions @LicenseOptions))
        }

        Write-Verbose -Message "License SKU: $AccountSkuId"
        Write-Verbose -Message "Service plans: $($ServicePlans -join ',')"
    }

    process
    {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName))
        {
            Write-Verbose -Message "Processing: $UserPrincipalName"

            $Params.UserPrincipalName = $UserPrincipalName

            try
            {
                Write-Verbose -Message "Attempting to add license $AccountSkuId..."
                $null = Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation 'US'
                Set-MsolUserLicense @Params -ErrorAction Stop
            }
            catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]
            {
                Write-Verbose -Message "User could not be licensed for $AccountSkuId. Attempting to set service plans..."
                $Params.Remove('AddLicenses')
                Set-MsolUserLicense @Params
            }
        }
    }
}