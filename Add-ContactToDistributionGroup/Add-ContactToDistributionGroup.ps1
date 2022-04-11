<#
.SYNOPSIS
Add users to a distribution list.

.DESCRIPTION
This script creates mail contacts for users and adds the mail contacts to a distribution list. 

Existing mail contacts that match the user's email address are added to the distribution list.

Existing mail contacts that match the users name but have a different email address are not added to the
distribution list.

Add users in bulk using the -CsvPath parameter. Add users individually using the -Name and -Email parameters.

Run the command first with the -WhatIf parameter to understand what changes will be made. For example, specify
the -WhatIf option to understand which mail contacts will be created, which mail contacts exist in the system, 
and which mail contacts that have conflicting records, i.e. same name but different email address.

.PARAMETER DistributionListAddress
The email address of the distribution list for which contacts will be added. 

.PARAMETER CsvPath
When adding users in bulk, specifies the path to a CSV file. The CSV file must have the following column 
headers: Name, Email.

The name column represents the display name of the contact, which is typically the first and last name together.
Additional columns representing first name and last name fields are not processed.

.PARAMETER Name
When adding users individually, specifies the display name of the user to be added to the distribution list. The 
display name typically includes the first and last name of the user.

.PARAMETER Email
When adding users individually, specifies the email address of the user to be added to the distribution list.

.EXAMPLE
Add a single users to a distribution list.

Add-ContactToDistributionGroup -DistributionListAddress 'dl@domain.com' -Name 'Tom Brady' -EmailAddress 'thomas.brady@domain.com'

.EXAMPLE
Add bulk users to a distribution list. Uses the $PWD variable to reference the CSV file in the same folder as the 
script.

.\Add-ContactToDistributionGroup.ps1 -DistributionListAddress 'dl@domain.com' -CsvPath "$($PWD.Path\users.csv"

.EXAMPLE
Run a check to determine which contacts will be created.

.\Add-ContactToDistributionGroup.ps1 -DistributionListAddress 'dl@domain.com' -CsvPath "$($PWD.Path)\users.csv" -WhatIf

.INPUTS
This command does not take any input from the pipeline.

.OUTPUTS
None. This command does not return any objects.

.NOTES
Version 1: 3/21/22 by Greg Tate

#>

#requires -modules ExchangeOnlineManagement

[CmdletBinding(SupportsShouldProcess)]
param(
    [OutputType()]

    [Parameter(Mandatory, ParameterSetName='Bulk')]
    [ValidateScript({Test-Path -Path $_})]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName='Adhoc')]
    [string]$Name,

    [Parameter(Mandatory, ParameterSetName='Adhoc')]
    [string]$EmailAddress,

    [Parameter(Mandatory)]
    [string]$DistributionListAddress 
)

$main = {

    # Run initial checks before proceeding w/ main program. Exit script if any of these checks fail.
    CheckExchangeOnlineConnectivity
    $distributionList = CheckExistenceOfDistributionList

    # Identify users to add 
    $users = @()
    switch ($PSCmdlet.ParameterSetName) {
        'Bulk' {
            CheckCSVFile
            $users = Import-CSV -Path $CsvPath
            break
        }
        'Adhoc' {
            $userProperties = @{
                Name = $Name
                Email = $EmailAddress
            }
            $user = New-Object -TypeName PSCustomObject -Property $userProperties
            $users += $user
            break
        }
    }

    $contactsToAdd = @()
    $contactsToOmit = @()

    Write-Verbose "Searching for existing contacts..."
    foreach ($u in $users) {
        $existingContactByEmail = Get-MailContact -Anr $u.Email -ErrorAction stop
        $existingContactByName = Get-MailContact -Anr $u.Name -ErrorAction stop
        if ($existingContactByEmail) {
            Write-Verbose "$($u.Name)`: Existing contact found with email address $($u.Email). Will add to distribution list."
            if ($PSCmdlet.ShouldProcess($u.Name, "Stage existing mail contact")) {
                $contactsToAdd += $existingContactByEmail
            } else {
                # WhatIf action: create a fake contact to be used for later WhatIf processing
                $fakeContactProperties = @{
                    DisplayName = $u.Name
                    PrimarySmtpAddress = $u.Email
                }
                $contactsToAdd += New-Object -TypeName PSCustomObject -Property $fakeContactProperties
            }
            continue
        }
        if ($existingContactByName) {
            Write-Warning "$($u.Name)`: Matching contact found with the same name but different email address `
            ($($existingContactByName.PrimarySmtpAddress)). Contact will not be added to distribution list."
            $contactsToOmit += $existingContactByName
            continue
        }
    }
    if ($contactsToAdd.count -eq 0) {
        Write-Verbose "No existing mail contacts found."
    }
    Write-Verbose "Completed contact search. `n"

    Write-Verbose "Creating mail contacts..."
    $contactsCreated = @()
    :parent foreach ($u in $users) {
        foreach ($c in $contactsToOmit) {
            if ($u.Name -eq $c.Name) {
                # Skip contact creation, as matching contact found w/ same name but different email address
                continue parent
            }
        }
        foreach ($c in $contactsToAdd) {
            if($u.Email -eq $c.PrimarySmtpAddress) {
                # Skip contact creation, as contact already exists in system
                continue parent
            }
        }
        # Mail contact doesn't exist; create new mail contact for the user
        if ($PSCmdlet.ShouldProcess($u.Name, "Create mail contact")) {
            Write-Verbose "$($u.Name)`: Creating mail contact w/ email address $($u.Email)."
            $contactProperties = @{
                Name = $u.Name
                ExternalEmailAddress = $u.Email
            }
            try {
                $createdContact = New-MailContact @contactProperties -ErrorAction stop
            } catch {
                $errorMessage = $_.Exception.Message
                $responseMessage = "Unable to create contact. Here's the error: `n`t $errorMessage"
                Write-Warning "$($u.name)`: $responseMessage"
                continue
            }
        } else {
            # WhatIf action: Create a fake contact to be used for later WhatIf processing
            $fakeContactProperties = @{
                DisplayName = $u.Name
                PrimarySmtpAddress = $u.Email
            }
            $createdContact = New-Object -TypeName PSCustomObject -Property $fakeContactProperties
        }
        $contactsCreated += $createdContact
        $contactsToAdd += $createdContact
    }
    if ($contactsCreated.count -eq 0) {
        Write-Verbose "No contacts created. All contacts exist in the system."
    }
    Write-Verbose "Completed addition of mail contacts. `n"

    Write-Verbose "Adding contacts to the $($distributionList.DisplayName) distribution list..."
    foreach ($c in $contactsToAdd) {
        $params = @{
            Identity = $distributionListAddress
            Member = $c.PrimarySmtpAddress
        }
        if ($PSCmdlet.ShouldProcess($c.PrimarySmtpAddress, "Add contact to distribution list")) {
            try {
                Add-DistributionGroupMember @params -ErrorAction stop
                Write-Host "$($c.DisplayName)`: Added to distribution list."
            } catch {
                $errorMessage = $_.Exception.Message
                switch -regex ($errorMessage) {
                    'The recipient .* is already a member' {
                        $responseMessage = "The recipient is already a member of the distribution list."
                        break
                    }
                    default {
                        $responseMessage = "Unable to add to distribution list. Here's the error: `n`t $errorMessage"
                    }
                }
                Write-Warning "$($c.DisplayName)`: $responseMessage"
                continue
            }
        }
    }
} # end of script execution

function CheckCSVFile {
    # Import CSV
    Write-Verbose "Checking content of CSV file..."
    try {
        $users = Import-Csv -Path $Script:CsvPath -ErrorAction stop
        if ($users.count -eq 0) {
            throw "There isn't any user date in the CSV file."
        }
        foreach ($user in $users) {
            $errorMessage = "Improperly-formatted CSV file. Make the following columns exist: Name, Email"
            if ($null -eq $user.Name -or $null -eq $user.Email) { throw $errorMessage}
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Issue importing CSV file. Here's the error message: `n`t $errorMessage"
        Write-Host "Exiting script."
        exit
    }
    Write-Verbose "Confirmed contents of CSV file.`n"
}

function CheckExchangeOnlineConnectivity {
    Write-Verbose "Checking connection to Exchange Online..."
    $getSessions = Get-PSSession | Select-Object -Property State, Name
    $isConnected = ( @($getSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*' ).count -gt 0
    if ($isConnected) {
        Write-Verbose "Confirmed connection to Exchange Online.`n"
    } else {
        Write-Warning "Not connected to Exchange Online. Run Connect-ExchangeOnline. Remember to activate role!"
        Write-Host "Exiting script."
        exit
    }
}

function CheckExistenceOfDistributionList {
    Write-Verbose "Checking for existence of distribution list $DistributionListAddress..."
    try {
        $distributionList = Get-DistributionGroup -Identity $DistributionListAddress -ErrorAction stop
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Unable to get distribution list. Here's the error message: `r`t $errorMessage"
        Write-Warning "Exiting script."
        exit
    }
    Write-Output $distributionList
    Write-Verbose "Confirmed existence of distribution list.`n"
}

& $main
