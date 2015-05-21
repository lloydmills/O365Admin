function Get-O365AccountSku
{
    $Skus = Get-MsolAccountSku -ErrorAction stop
    $CustomSkus = @()
    foreach ($Sku in $Skus)
    {
        $Elements = $Sku.AccountSkuId -split ':'
        $DomainName = $Elements[0]
        $AccountSkuId = $Elements[1]
        $CustomSku = [PsCustomObject]@{
            AccountSkuId = $AccountSkuId
            AvailablePlans = $Sku.Servicestatus.ServicePlan.ServiceName
            DomainName = $DomainName
        }
        $CustomSkus += $CustomSku
    }
    $CustomSkus
}

$TypeTemplate = @"
    namespace O365Admin
    {
	    public enum AccountSkuId
	    {
		    #Placeholder
	    }
    }
"@
$CSharpPath = "$PSScriptRoot\bin\O365AccountSkuIds.cs"
$XmlPath = "$PSScriptRoot\bin\O365AccountSkuIds.xml"

# Create AccountSkuId enum type definition file
if (!(Test-Path $CSharpPath))
{
    Write-Warning -Message "Setting up module for first use."
    Write-Warning -Message "Please enter credentials for an Office 365 administrative account."
    Connect-MsolService -Credential (Get-Credential -Message 'Enter O365 Administrative Credentials.') -ErrorAction Stop
    Get-O365AccountSku | Export-Clixml -Path $XmlPath
    $O365AccountSkus = Import-Clixml $XmlPath
    $TypeDefinition = $TypeTemplate -replace '#Placeholder', "$($O365AccountSkus.AccountSkuId -join ', ')"
    $TypeDefinition | Out-File $CSharpPath
}
else
{
    $O365AccountSkus = Import-Clixml $XmlPath
}

Add-Type -TypeDefinition @'
    using System;

    [Flags]
    public enum O365Services
    {
        AzureActiveDirectory,
        Exchange,
        Skype,
        Sharepoint,
        All = AzureActiveDirectory | Exchange | Sharepoint | Skype
    }
'@