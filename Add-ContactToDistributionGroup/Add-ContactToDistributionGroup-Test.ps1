$testCase = 'bulk'
$Verbose = $true
$WhatIf = $false
$DistributionList = 'testdl@tate365sandbox.dev'

switch ($testCase) {
    'bulk' {
        $params = @{
            # CSV file needs the following columns: Name, Email
            CsvPath = "$($PWD.Path)\Test.csv"
            DistributionListAddress = $DistributionList
            Verbose = $Verbose
            WhatIf = $WhatIf
        }
        .\Add-ContactToDistributionGroup.ps1 @params
        break
    }
    'adhoc' {
        $params = @{
            Name = 'Greg Tate'
            EmailAddress = 'greg@abc.com'
            DistributionListAddress = $DistributionList
            Verbose = $Verbose
            WhatIf = $WhatIf
        }
        .\Add-ContactToDistributionGroup.ps1 @params
        break
    }
}
