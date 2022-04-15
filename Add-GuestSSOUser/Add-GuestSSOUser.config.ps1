$tenants = @{
    Sandbox = 'c20737c4-bc9c-49aa-a451-90862566d79c'
}
$activeTenant = $tenants.Sandbox

$global:config = @{
    
    tenantID = $activeTenant

    requiredPermissions = @(
        'microsoft.directory/users/basic/update'
        'microsoft.directory/users/allProperties/allTasks'
    )

}