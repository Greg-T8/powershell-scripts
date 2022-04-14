$Verbose = $true
$testCase = 'adhoc'

switch ($testCase) {
    'adhoc' {
        $params = @{
            Name = 'John Smith'
            EmailAddress = 'john@abc.com'
            FirstName = 'John'
            LastName = 'Smith'
        }
        .\Add-GuestSSOUser.ps1 @params
    }
}