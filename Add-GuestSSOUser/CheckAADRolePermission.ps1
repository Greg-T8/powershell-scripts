#requires -modules AzureADPreview

# Check for connection to AzureAD
try {
    $sessionInfo = Get-AzureADCurrentSessionInfo -ErrorAction stop
} catch {
    Write-Warning "Not connected to AzureAD. Run Connect-AzureAD."
    exit
}

# Establish required permissions
$requiredPermissions = @(
    'microsoft.directory/users/basic/update'
    'microsoft.directory/users/allProperties/allTasks'
)

# Get current user
$currentUser = Get-AzureADUser -Filter "userPrincipalName eq '$($sessionInfo.account.id)'"

# Get all Azure AD roles
$roleDefinitions = Get-AzureADMSRoleDefinition

# Get required roles
$requiredRoles = $roleDefinitions | 
    ForEach-Object {$_} -PipelineVariable roleDefinition | 
    Select-Object -ExpandProperty rolepermissions | 
    Select-Object -ExpandProperty allowedresourceactions |
    Where-Object {$_ -in $requiredPermissions} |
    ForEach-Object {Write-Output $roleDefinition} 

# Determine role assignments
$assignedRoles = $requiredRoles | 
    ForEach-Object {
        $filter = "roleDefinitionId eq '$($_.Id)' and PrincipalId eq '$($currentUser.ObjectId)'"
        Get-AzureADMSRoleAssignment -Filter $filter
    }
$assignedRoleNames = $roleDefinitions | 
    Where-Object {$_.id -in $assignedRoles.roleDefinitionId} |
    ForEach-Object {Write-Output $_.DisplayName}

# Output results
if ($assignedRoleNames) {
    Write-Host "$($currentUser.DisplayName)`: Active roles -> $($assignedRoleNames -join ', ')"
} else {
    Write-Host "$($currentUser.DisplayName)`: No permissible roles assigned"
}
