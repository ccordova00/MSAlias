<#
    Author: ccordova
    Date modified: 2023-09-08

    #ToDo: 
#>

Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    [string]$Password
)
Function Connect_Exo
{
 #Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host `nConnecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 else
 {
  Connect-ExchangeOnline
 }
}


if($UserName -eq "")
{
# Prompt user for username
    $UserName = Read-Host "Please enter the username of the account you would like to change (e.g., user@example.com)"
}

# Call connect
Connect_Exo

# Get the user's email aliases (proxy addresses)
$user = Get-Mailbox -Identity $UserName
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
        Set-Mailbox $UserName -EmailAddresses @{Add=$newAlias}
        Write-Host "Alias created."
        #break
    }

    # Delete an alias
    elseif ($action -eq "2") {
        $deleteAlias = Read-Host "Enter the alias you would like to delete"
        
        # Extract actual email addresses from MatchInfo objects
        $aliasStrings = $aliases | ForEach-Object { $_.ToString().Substring(5) }

        if ($aliasStrings -contains $deleteAlias) {
            # Remove the alias
            $newAliases = $aliasStrings | Where-Object { $_ -ne $deleteAlias }
            
            # Update the mailbox
            Set-Mailbox -Identity $UserName -EmailAddresses $newAliases
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

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
