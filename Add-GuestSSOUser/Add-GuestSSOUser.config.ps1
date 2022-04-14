$tenants = @{
    Sandbox = 'c20737c4-bc9c-49aa-a451-90862566d79c'
}

ActiveTenant = $tenants.Sandbox

$global:config = @{
    
    tenantID = $ActiveTenant

    requiredPermissions = @(
        'microsoft.directory/users/basic/update'
        'microsoft.directory/users/allProperties/allTasks'
    )

}
