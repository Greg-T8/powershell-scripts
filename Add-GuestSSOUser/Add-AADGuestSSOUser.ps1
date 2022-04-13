#requires -modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement

[CmdletBinding()]
param(
)

$main = {    
    # Load configuration
    LoadConfigFile

    # Verify connectivity to Azure AD
    $sessionInfo = GetSessionInfo -TenantId $config.tenantId

    # Check permissions
    CheckRequiredRoles -RequiredPermissions $config.requiredPermissions -SessionInfo $sessionInfo

} 

function CheckRequiredRoles {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$RequiredPermissions,

        [Parameter(Mandatory=$true)]
        [Microsoft.Open.Azure.AD.CommonLibrary.PSAzureContext]$SessionInfo
    )

    # Get current user
    try {
        $currentUser = Get-AzureADUser -Filter "userPrincipalName eq '$($SessionInfo.Account.Id)'"
    } catch {
        Write-Warning "Issue getting AzureAD user. Check connectivity to AzureAD. Exiting script."
        exit
    }
    # Get all Azure AD roles
    $roleDefinitions = Get-AzureADMSRoleDefinition

    # Get required roles
    $requiredRoles = $roleDefinitions | 
        ForEach-Object {$_} -PipelineVariable roleDefinition | 
        Select-Object -ExpandProperty rolepermissions | 
        Select-Object -ExpandProperty allowedresourceactions |
        Where-Object {$_ -in $RequiredPermissions} |
        ForEach-Object {Write-Output $roleDefinition} 

    # Match assigned/current roles with required roles
    $assignedRoles = $requiredRoles | 
        ForEach-Object {
            $filter = "roleDefinitionId eq '$($_.Id)' and PrincipalId eq '$($currentUser.ObjectId)'"
            Get-AzureADMSRoleAssignment -Filter $filter
        }

    # Get friendly role names
    $assignedRoleNames = $roleDefinitions | 
        Where-Object {$_.id -in $assignedRoles.roleDefinitionId} |
        ForEach-Object {Write-Output $_.DisplayName}
    
    if (-not $assignedRoleNames) {
        $warningMessage
        Write-Warning "You do not have the required permissions to run this script. Activate a valid role."
        Write-Warning "Here is a list of valid roles: $($requiredRoles -join ', ')"
        Write-Warning "Exiting script."
        exit
    } else {
       Write-Host "Verified roles." 
    }
}

function LoadConfigFile {
    try {
        $configFileName = $PSCmdlet.MyInvocation.MyCommand.Name -replace ('.ps1', '.config.ps1')
        $configFile = Get-ChildItem "$PSScriptRoot\$configFileName" -ErrorAction stop
    } catch {
        Write-Warning "Unable to find config file. Exiting script."
        exit
    }
    try {
        # Creates $config variable
        Invoke-Expression -Command $configFile.FullName -ErrorAction stop
    } catch {
        Write-Warning "Unable to load config file. Exiting script."
        exit
    }
}

function GetSessionInfo {
    param(
        [string]$TenantId
    )

    [Microsoft.Graph.PowerShell.Authentication.AuthContext]$sessionInfo = Get-MGContext
    if ($sessionInfo) {
        Write-Host "Connected to Microsoft Graph using $($sessionInfo.Account)"
    } else {
        Write-Warning "Not connected to Microsoft Graph. Running Connect-MgGraph."
        $sessionInfo = Connect-MgGraph -TenantId $TenantId -ErrorAction stop
        Write-Host "Connected to Microsoft Graph using $($sessionInfo.Account)"
    }
    return $sessionInfo
}

& $main