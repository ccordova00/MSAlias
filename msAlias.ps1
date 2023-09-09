<#
    Author: ccordova
    Date modified: 2023-09-08

    #ToDo: Fix delete menu item logic
#>
# Prompt user for username
$username = Read-Host "Please enter your username (e.g., user@example.com)"

# Connect to Exchange Online without showing progress
Connect-ExchangeOnline -UserPrincipalName $username -ShowProgress $false

# Get the user's email aliases (proxy addresses)
$user = Get-Mailbox -Identity $username
$aliases = ($user.EmailAddresses | Select-String -Pattern "smtp")
If($aliases -eq "")
{
 $aliases="-"
}
Write-Host "Current Aliases:"
$aliases

# Loop until a valid choice is made
while ($true) {
    # Ask the user what they want to do
    $action = Read-Host "Would you like to (1) create an alias, (2) delete an alias, or (3) exit? Enter 1, 2, or 3"

    # Create an alias
    if ($action -eq "1") {
        $newAlias = Read-Host "Enter the new alias you would like to create (e.g., alias@example.com)"
        Set-Mailbox $username -EmailAddresses @{Add=$newAlias}
        Write-Host "Alias created."
        #break
    }

    # Delete an alias
    elseif ($action -eq "2") {
        $deleteAlias = Read-Host "Enter the alias you would like to delete"
        
        # Extract actual email addresses from MatchInfo objects
        $aliasStrings = $aliases | ForEach-Object { $_.ToString() }

        if ($aliasStrings -contains $deleteAlias) {
            # Remove the alias
            $newAliases = $aliasStrings | Where-Object { $_ -ne $deleteAlias }
            
            # Update the mailbox
            Set-Mailbox -Identity $username -EmailAddresses $newAliases
            Write-Host "Alias deleted."
            #break
        }
        else {
            Write-Host "Alias not found."
        }
    }

    # Exit
    elseif ($action -eq "3") {
        Write-Host "Exiting."
        break
    }

    # Invalid choice
    else {
        Write-Host "Invalid choice. Please try again."
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
