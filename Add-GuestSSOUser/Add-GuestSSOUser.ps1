<#
.SYNOPSIS
Invites a guest user to Azure AD and gives access to an Azure AD SSO-based application.

.DESCRIPTION
This script is used to give guest users access to an Azure AD SSO-based application. Normally, the guest invitation
process involves sending an invitation message to the guest user. In many cases the user does not expect a welcome
message. This script suprresses the welcome message so that the user consents only when accessing the application
for the first time.

.PARAMETER TBD

.EXAMPLE
TBD

.INPUTS
TBD

.OUTPUTS
TBD

.LINK
TBD

.NOTES
This script uses the Microsoft Graph module to perform a number of checks to ensure script execution. These checks
include verifying connectivity to Microsoft Graph and verifying whether the user has the appropriate role activated
in order to execute the script.

#>

#requires -modules @{ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="1.9.5"}
#requires -modules @{ModuleName="Microsoft.Graph.Identity.DirectoryManagement"; ModuleVersion="1.9.5"}

[CmdletBinding(SupportsShouldProcess)]
param(
    [OutputType()]

    [Parameter(ParameterSetName='Bulk')]
    [ValidateScript({Test-Path -Path $_})]
    [string]$CsvPath,

    [Parameter(ParameterSetName='Adhoc')]
    [string]$Name,

    [Parameter(ParameterSetName='Adhoc')]
    [string]$EmailAddress,

    [Parameter(ParameterSetName='Adhoc')]
    [string]$FirstName,

    [Parameter(ParameterSetName='Adhoc')]
    [string]$LastName
    
)


$main = {    
    # Run initial checks before proceeding with main program. Exit script if any of these checks fail.

    # Load configuration
    LoadConfigFile

    # Verify connectivity to Microsoft Graph
    $sessionInfo = GetSessionInfo -TenantId $config.tenantId

    # Check permissions
    CheckRequiredRoles -RequiredPermissions $config.requiredPermissions -SessionInfo $sessionInfo

}

function LoadConfigFile {
    try {
        $configFileName = $PSCmdlet.MyInvocation.MyCommand.Name -replace ('.ps1', '.config.ps1')
        $configFile = Get-ChildItem "$PSScriptRoot\$configFileName" -ErrorAction stop
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Unable to find config file. Here's the error message: `n`t $errorMessage"
        Write-Warning "Exiting script."
        exit
    }
    try {
        # Loads the $config variable from the configuration file
        & $configFile.FullName -ErrorAction stop
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Unable to load config file. Here's the error mesage: `n`t $errorMessage"
        Write-Warning "Exiting script."
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
        try {
            Connect-MgGraph -TenantId $TenantId -ErrorAction stop | Out-Null
            $sessionInfo = Get-MGContext
        } catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "There was an issue connecting to Microsoft Graph. Here's the error message `n`t $errorMessage"
            Write-Warning "Exiting script."
            exit
        }
        Write-Host "Connected to Microsoft Graph using $($sessionInfo.Account)"
    }
    return $sessionInfo
}


function CheckRequiredRoles {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$RequiredPermissions,

        [Parameter(Mandatory=$true)]
        [Microsoft.Graph.PowerShell.Authentication.AuthContext]$SessionInfo
    )

    $currentUser = $SessionInfo.Account

    # Get all Azure AD roles
    $roleDefinitions = Get-MgDirectoryRoleTemplate

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

& $main