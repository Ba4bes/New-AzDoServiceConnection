Function New-AzDoServiceConnection {
    <#
    .SYNOPSIS
    This function creates an Azure DevOps service connection for AzureRM.
    .DESCRIPTION
    A service principal with set permissions is created in Azure.
    This principal is used to create an AzureRM service connection in Azure DevOps
    .PARAMETER AzServicePrincipalName
    The name the Service Principal in Azure. Has to be unique
    .PARAMETER AzSubscriptionName
    The subscription that the service connection will connect to.
    If no resourcegroupscope is added, permissions will be set to this subscription
    .PARAMETER AzResourceGroupScope
    A resourcegroup that the Connection needs permissions to.
    If left empty, permissions will be set to the subscription.
    .PARAMETER AzRole
    The AzRoleDefinition that the Service principal needs
    .PARAMETER AzDoOrganizationName
    The organization name in Azure DevOps
    .PARAMETER AzDoProjectName
    The project name in Azure DevOps
    .PARAMETER AzDoConnectionName
    A name for the Azure DevOps Connection.
    If left empty, defaults to the name of the subscription without spaces
    .PARAMETER AzDoUserName
    The username to use to connect to Azure DevOps
    .PARAMETER AzDoToken
    The PAT token to use to connect to Azure DevOps
    .EXAMPLE
    $Parameters = @{
    AzServicePrincipalName = example
    AzSubscriptionName = "subscription01"
    AzResourceGroupScope = "RG01"
    AzRole = "owner"
    AzDoOrganizationName = AzDoCompany
    AzDoProjectName = AzureDeployment
    AzDoUserName = user@domain.com
    AzDoToken = "afweafawe3228faefa0w32f0A"
    }

    New-AzDoServiceConnection @Parameters
    ===

    Will create a serviceprincipal called example with owner permissions to the resourcegroup RG01.
    Will create a connection in Azure DevOps organization AzDoCompany for project AzureDeployment.
    }
    .NOTES
    PAT token needs permissions for Service Connections: Read, query, & manage
    Minimum permissions for Azure account:
    - Azure Application administrator
    - Owner on the resourcegroup or subscription that is scoped.

    Created by Barbara Forbes
    https://4bes.nl
    
    #>
    #Requires -Module Az.Resources, Az.Accounts
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$AzServicePrincipalName,

        [parameter(Mandatory = $true)]
        [ValidateScript( { Get-AzSubscription -SubscriptionName $_ })]
        [string]$AzSubscriptionName,

        [parameter(Mandatory = $false)]
        [string]$AzResourceGroupScope,

        [parameter(Mandatory = $false)]
        [ValidateScript( { Get-AzRoleAssignment -RoleDefinitionName $_ })]
        [string]$AzRole = "Contributor",

        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$AzDoOrganizationName,

        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$AzDoProjectName,

        [parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$AzDoConnectionName,

        [parameter(Mandatory = $false)]
        [ValidateNotNullorEmpty()]
        [string]$AzDoUserName,

        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$AzDoToken
    )

    Write-Verbose "Starting Function New-AzDoServiceConnection"

    # Create the header to authenticate to Azure DevOps
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $AzDoUserName, $AzDoToken)))
    $Header = @{
        Authorization = ("Basic {0}" -f $base64AuthInfo)
    }
    Remove-Variable AzDoToken
    try {
        $AzSubscription = Get-AzSubscription -SubscriptionName $AzSubscriptionName -ErrorAction Stop
    }
    Catch {
        Throw "Could not find subscription $AzSubscriptionName. Please verify it exists"
    }
    $AzSubscriptionID = $AzSubscription.Id
    $TenantId = $AzSubscription.TenantId
    if ($AzResourceGroupScope) {
        Write-Verbose "Changing Context to $AzSubscriptionName"
        $Null = Set-AzContext $AzSubscriptionID
        # Check if resourcegroup exists before setting it as the scope
        Try {
            $null = Get-AzResourceGroup -Name $AzResourceGroupScope -ErrorAction Stop
            Write-Verbose "Resourcegroup exists"
        }
        Catch {
            Throw "Resourcegroup $AzResourceGroupScope was not found"
        }
        $Scope = "/subscriptions/$AzSubscriptionID/resourceGroups/$AzResourceGroupScope"
    }
    else {
        $Scope = "/subscriptions/$AzSubscriptionID"
    }

    Write-Verbose "Scope set: $Scope"
    # Create the Service Principal
    Try {
        $Parameters = @{
            DisplayName = $AzServicePrincipalName
            Role        = $AzRole
            Scope       = $Scope
            ErrorAction = "Stop"
        }
        $ServicePrincipal = New-AzADServicePrincipal @Parameters
        Write-Verbose "Created ServicePrincipal $AzServicePrincipalName"
    }
    Catch {
        Throw "Could not create the ServicePrincipal: $_"
    }

    ## Get ProjectId
    $URL = "https://dev.azure.com/$AzDoOrganizationName/_apis/projects?api-version=6.0"
    Try {
        $AzDoProjectNameproperties = (Invoke-RestMethod $URL -Headers $Header -ErrorAction Stop).Value
        Write-Verbose "Collected Azure DevOps Projects"
    }
    Catch {
        if ($_ | Select-String -Pattern "Access Denied: The Personal Access Token used has expired.") {
            Throw "Access Denied: The Azure DevOps Personal Access Token used has expired."
        }
        else {
            $ErrorMessage = $_ | ConvertFrom-Json
            Throw "Could not collect project: $($ErrorMessage.message)"
        }
    }
    $AzDoProjectID = ($AzDoProjectNameproperties | Where-Object { $_.Name -eq $AzDoProjectName }).id
    Write-Verbose "Collected ID: $AzDoProjectID"

    if (-not $AzDoConnectionName) {
        $AzDoConnectionName = $AzSubscriptionName -replace " "
    }

    if ($PSVersionTable.PSVersion.Major -gt 7) {
        $PlainTextSecret = $ServicePrincipal.PasswordCredentials.SecretText
    }
    else {
        $PlainTextSecret = [System.Net.NetworkCredential]::new("", $ServicePrincipal.Secret).Password
    }

    # Create body for the API call
    $Body = @{
        data                             = @{
            subscriptionId   = $AzSubscriptionID
            subscriptionName = $AzSubscriptionName
            environment      = "AzureCloud"
            scopeLevel       = "Subscription"
            creationMode     = "Manual"
        }
        name                             = ($AzSubscriptionName -replace " ")
        type                             = "AzureRM"
        url                              = "https://management.azure.com/"
        authorization                    = @{
            parameters = @{
                tenantid            = $TenantId
                serviceprincipalid  = $ServicePrincipal.AppId
                authenticationType  = "spnKey"
                serviceprincipalkey = $PlainTextSecret
            }
            scheme     = "ServicePrincipal"
        }
        isShared                         = $false
        isReady                          = $true
        serviceEndpointProjectReferences = @(
            @{
                projectReference = @{
                    id   = $AzDoProjectID
                    name = $AzDoProjectName
                }
                name             = $AzDoConnectionName
            }
        )
    }
    Remove-Variable PlainTextSecret
    $URL = "https://dev.azure.com/$AzDoOrganizationName/$AzDoProjectName/_apis/serviceendpoint/endpoints?api-version=6.0-preview.4"
    $Parameters = @{
        Uri         = $URL
        Method      = "POST"
        Body        = ($Body | ConvertTo-Json -Depth 3)
        Headers     = $Header
        ContentType = "application/json"
        Erroraction = "Stop"
    }
    try {
        Write-Verbose "Creating Connection"
        $Result = Invoke-RestMethod @Parameters
    }
    Catch {
        $ErrorMessage = $_ | ConvertFrom-Json
        Throw "Could not create Connection: $($ErrorMessage.message)"
    }
    Write-Verbose "Connection Created"
    $Result
}