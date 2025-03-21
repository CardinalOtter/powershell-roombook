# --- Script Start ---

# --- Global Variables ---

# Controls the verbosity of the script.  "SilentlyContinue" (default) suppresses verbose output.
# "Continue" enables verbose output, providing detailed information for debugging.
$DebugPreference = "SilentlyContinue"

# --- Function Definitions ---

# Displays the main menu and gets the user's choice.
function Show-Menu {
  param() # No parameters needed for the menu display

  Clear-Host # Clear the console screen for a cleaner interface.

  # Display the menu title and options.
  Write-Host "Exchange Online Room Mailbox Configuration Menu" -ForegroundColor Green
  Write-Host "-----------------------------------------------" -ForegroundColor Green
  Write-Host "1. Set Calendar Processing Settings"
  Write-Host "2. Grant Editor Access to Coordinator Group"
  Write-Host "3. Set Default User Access Rights to Reviewer"
  Write-Host "4. Grant Reviewer Access to Digital Signage Group"
  Write-Host "5. Add Resource Delegates"
  Write-Host "6. Remove Existing Permissions for a User"
  Write-Host "7. Show Current Permissions"
  Write-Host "8. Show Calendar Settings"
  Write-Host "9. Enable/Disable Debug Mode (Current: $($DebugPreference))" -ForegroundColor Yellow
  Write-Host "10. Exit"
  Write-Host ""

  # Prompt the user for their choice and return the input.
  $choice = Read-Host "Enter your choice (1-10)"
  return $choice
}

# Sets the calendar processing settings for the specified room mailbox.
function Set-CalendarProcessingSettings {
  param(
    [string]$roomMailbox # The identity (email address) of the room mailbox.
  )

  Write-Verbose "Entering Set-CalendarProcessingSettings function" # Debug information

  # Inform the user about the action being performed.
  Write-Host "Setting Calendar Processing Settings for $($roomMailbox)..." -ForegroundColor Cyan

  # Use a try-catch block to handle potential errors during the operation.
  try {
    # Execute the Set-CalendarProcessing cmdlet to modify the settings.
    # -AddOrganizerToSubject $false:  Do not add the organizer's name to the subject.
    # -DeleteSubject $false: Do not delete the subject.
    # -ErrorAction Stop: Treat errors as terminating errors, causing the script to stop.
    Set-CalendarProcessing -Identity $roomMailbox -AddOrganizerToSubject $false -DeleteSubject $false -ErrorAction Stop
    Write-Host "Calendar processing settings updated successfully." -ForegroundColor Green
  }
  catch {
    # If an error occurs, display a user-friendly error message in red.
    Write-Host "Error updating calendar processing settings: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString() # Detailed error information in debug mode
  }

  Write-Verbose "Exiting Set-CalendarProcessingSettings function" # Debug information
  Start-Sleep -Seconds 2 # Pause for 2 seconds to allow the user to read the output.
}

# Grants "Editor" access rights to the specified coordinator group on the room mailbox's calendar.
function Grant-EditorAccess {
  param(
    [string]$roomMailbox # The identity (email address) of the room mailbox.
  )

  Write-Verbose "Entering Grant-EditorAccess function"

  # Prompt the user for the email address of the coordinator group.
  $coordinatorGroup = Read-Host "Enter the email address of the coordinator group"

  Write-Host "Granting Editor access to $($coordinatorGroup) on $($roomMailbox)'s calendar..." -ForegroundColor Cyan

  try {
    # Add-MailboxFolderPermission: Grants permissions to a folder within a mailbox.
    # -Identity "$roomMailbox:\Calendar": Specifies the calendar folder of the room mailbox.
    # -User $coordinatorGroup: The user or group to grant permissions to.
    # -AccessRights Editor: Grants the "Editor" access rights, allowing modification of calendar items.
    Add-MailboxFolderPermission -Identity "${roomMailbox}:\Calendar" -User $coordinatorGroup -AccessRights Editor -ErrorAction Stop
    Write-Host "Editor access granted to $($coordinatorGroup) successfully." -ForegroundColor Green
  }
  catch {
    Write-Host "Error granting Editor access: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Grant-EditorAccess function"
  Start-Sleep -Seconds 2
}

# Sets the default user access rights on the room mailbox's calendar to "Reviewer".
function Set-DefaultUserAccess {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Set-DefaultUserAccess function"

  Write-Host "Setting default user access rights to Reviewer on $($roomMailbox)'s calendar..." -ForegroundColor Cyan

  try {
    # Set-MailboxFolderPermission: Modifies permissions on a mailbox folder.
    # -User Default:  Refers to the default permissions for all users.
    # -AccessRights Reviewer: Grants the "Reviewer" access rights, allowing viewing of calendar items.
    Set-MailboxFolderPermission -Identity "${roomMailbox}:\Calendar" -User Default -AccessRights Reviewer -ErrorAction Stop
    Write-Host "Default user access rights set to Reviewer successfully." -ForegroundColor Green
  }
  catch {
    Write-Host "Error setting default user access rights: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Set-DefaultUserAccess function"
  Start-Sleep -Seconds 2
}

# Grants "Reviewer" access rights to the specified digital signage group on the room mailbox's calendar.
function Grant-SignageAccess {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Grant-SignageAccess function"

  $signageGroup = Read-Host "Enter the email address of the digital signage access group"
  
  if (-not ($signageGroup -match '^[^@]+@[^@]+\.[^@]+$') -or -not (Get-Recipient $signageGroup -ErrorAction SilentlyContinue)) {
      Write-Host "Invalid email format or non-existent recipient: $signageGroup" -ForegroundColor Red
      return
  }

  Write-Host "Granting Reviewer access to $($signageGroup) on $($roomMailbox)'s calendar..." -ForegroundColor Cyan

  try {
    Add-MailboxFolderPermission -Identity "${roomMailbox}:\Calendar" -User $signageGroup -AccessRights Reviewer -ErrorAction Stop
    Write-Host "Reviewer access granted to $($signageGroup) successfully." -ForegroundColor Green
  }
  catch {
    Write-Host "Error granting Reviewer access: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Grant-SignageAccess function"
  Start-Sleep -Seconds 2
}

# Adds the specified group as resource delegates for the room mailbox.
function Add-ResourceDelegates {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Add-ResourceDelegates function"

  $delegateGroup = Read-Host "Enter the email address of the delegate group"

  Write-Host "Adding $($delegateGroup) as resource delegates for $($roomMailbox)..." -ForegroundColor Cyan

  try {
    # -ResourceDelegates: Specifies the users or groups to be added as resource delegates.
    #  Resource delegates can manage meeting requests on behalf of the room mailbox.
    Set-CalendarProcessing -Identity $roomMailbox -ResourceDelegates $delegateGroup -ErrorAction Stop
    Write-Host "$($delegateGroup) added as resource delegates successfully." -ForegroundColor Green
  }
  catch {
    Write-Host "Error adding resource delegates: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Add-ResourceDelegates function"
  Start-Sleep -Seconds 2
}

# Removes existing permissions for a specified user from the room mailbox's calendar.
function Remove-UserPermissions {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Remove-UserPermissions function"

  $userToRemove = Read-Host "Enter the email address of the user to remove permissions for"
  
  if (-not ($userToRemove -match '^[^@]+@[^@]+\.[^@]+$') -or -not (Get-Recipient $userToRemove -ErrorAction SilentlyContinue)) {
      Write-Host "Invalid email format or non-existent recipient: $userToRemove" -ForegroundColor Red
      return
  }

  Write-Host "Removing permissions for $($userToRemove) from $($roomMailbox)'s calendar..." -ForegroundColor Cyan

  try {
    # Remove-MailboxFolderPermission: Removes permissions from a mailbox folder.
    # -Confirm:$false: Suppresses the confirmation prompt.
    Remove-MailboxFolderPermission -Identity "${roomMailbox}:\Calendar" -User $userToRemove -Confirm:$false -ErrorAction Stop
    Write-Host "Permissions removed for $($userToRemove) successfully." -ForegroundColor Green
  }
  catch {
    Write-Host "Error removing permissions: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Remove-UserPermissions function"
  Start-Sleep -Seconds 2
}

# Displays the current permissions on the room mailbox's calendar.
function Show-CurrentPermissions {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Show-CurrentPermissions function"

  Write-Host "Showing current permissions for $($roomMailbox)'s calendar..." -ForegroundColor Cyan

  try {
    # Get-MailboxFolderPermission: Retrieves permissions on a mailbox folder.
    $permissions = Get-MailboxFolderPermission -Identity "${roomMailbox}:\Calendar" -ErrorAction Stop
    if ($permissions) {
      $permissions | Format-Table # Display the permissions in a formatted table.
    }
    else {
      Write-Host "No permissions found for $($roomMailbox)'s calendar."
    }
  }
  catch {
    Write-Host "Error retrieving permissions: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Show-CurrentPermissions function"
  Start-Sleep -Seconds 2 # Keep output on screen
}

# Displays the current calendar processing settings for the room mailbox.
function Show-CalendarSettings {
  param(
    [string]$roomMailbox
  )

  Write-Verbose "Entering Show-CalendarSettings function"

  Write-Host "Showing calendar settings for $($roomMailbox)..." -ForegroundColor Cyan

  try {
    # Get-CalendarProcessing: Retrieves calendar processing settings for a mailbox.
    # Format-List: Displays the output as a detailed list.
    Get-CalendarProcessing -Identity $roomMailbox -ErrorAction Stop | Format-List
  }
  catch {
    Write-Host "Error retrieving calendar settings: $($_.Exception.Message)" -ForegroundColor Red
    Write-Verbose $_.Exception.ToString()
  }

  Write-Verbose "Exiting Show-CalendarSettings function"
  Start-Sleep -Seconds 2 # Keep output on screen
}

# Toggles the debug mode on or off.
function Toggle-DebugMode {
  param() # No parameters needed

  Write-Verbose "Entering Toggle-DebugMode function"

  # Check the current state of $DebugPreference and toggle it.
  if ($DebugPreference -eq "SilentlyContinue") {
    $DebugPreference = "Continue" # Enable debug mode
    Write-Host "Debug mode enabled." -ForegroundColor Yellow
  }
  else {
    $DebugPreference = "SilentlyContinue" # Disable debug mode
    Write-Host "Debug mode disabled." -ForegroundColor Yellow
  }

  Write-Verbose "Exiting Toggle-DebugMode function"
  Start-Sleep -Seconds 2
}

# --- Main Script Logic ---

# --- Module Installation and Execution Policy ---

# Check if the Exchange Online Management module is installed.
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
  # If not installed, prompt the user for confirmation to install it.
  $installModule = Read-Host "The ExchangeOnlineManagement PowerShell module is not installed. Do you want to install it now? (y/n)"
  if ($installModule -eq "y") {
      # Check and set Execution Policy
      if ((Get-ExecutionPolicy) -ne 'RemoteSigned') {
        $setExecutionPolicy = Read-Host "The execution policy needs to be set to RemoteSigned to install and use the ExchangeOnlineManagement module. Do you want to set it now? (y/n)"
          if ($setExecutionPolicy -eq "y") {
              try{
                Set-ExecutionPolicy RemoteSigned -Force -Confirm:$false -ErrorAction Stop # Set execution policy, suppressing prompts.
                Write-Host "Execution Policy set to RemoteSigned" -ForegroundColor Green
              } catch {
                Write-Host "Failed to set Execution Policy: $($_.Exception.Message)" -ForegroundColor Red
                exit
              }
          }
          else{
            Write-Host "Cannot proceed without setting the execution policy. Exiting." -ForegroundColor Red
            exit # Exit the script if the user declines.
          }
      }
    # Install the Exchange Online Management module, suppressing the confirmation prompt.
    try{
        Install-Module -Name ExchangeOnlineManagement -Confirm:$false -ErrorAction Stop
        Write-Host "ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
  }
  else {
    Write-Host "Cannot proceed without the ExchangeOnlineManagement module. Exiting." -ForegroundColor Red
    exit # Exit the script if the user declines.
  }
} else {
    Write-Verbose "ExchangeOnlineManagement module is already installed."
}

# --- Connect to Exchange Online ---

# Prompt the user for their Exchange Online administrator UPN.
# Secure credential prompt with MFA

# Validate Exchange Online module version
$minVersion = [version]"3.0"
$module = Get-Module ExchangeOnlineManagement -ListAvailable | Select-Object -First 1
if (-not $module -or $module.Version -lt $minVersion) {
    Write-Host "Requires ExchangeOnlineManagement module v3.0+ - Found: $($module.Version)" -ForegroundColor Red
    exit
}

# Use a try-catch block to handle potential connection errors.
try {
  Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
  Connect-ExchangeOnline -ErrorAction Stop
  Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
}
catch {
  # Display a user-friendly error message if the connection fails.
  Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
  Write-Verbose $_.Exception.ToString()
  exit # Exit the script since we cannot proceed without a connection.
}

# --- Get Room Mailbox Identity and Validate ---

# Prompt the user for the email address of the room mailbox.
$roomMailbox = Read-Host "Enter the email address of the room mailbox to configure"

# Use a try-catch block to handle potential errors when retrieving the mailbox.
try {
  # Attempt to retrieve the mailbox using Get-Mailbox.
  # -ErrorAction SilentlyContinue:  Suppresses error messages if the mailbox is not found.
  $mailbox = Get-Mailbox -Identity $roomMailbox -ErrorAction SilentlyContinue
  if (-not $mailbox) {
    # If the mailbox is not found, display an error message and exit.
    Write-Host "Room mailbox '$($roomMailbox)' not found. Exiting." -ForegroundColor Red
    exit
  }
  Write-Host "Room mailbox '$($roomMailbox)' found." -ForegroundColor Green
}
catch {
  # Display an error message if an exception occurs during mailbox retrieval.
  Write-Host "Error retrieving mailbox: $($_.Exception.Message)" -ForegroundColor Red
  Write-Verbose $_.Exception.ToString()
  exit # Exit since we cannot proceed without a valid mailbox.
}

# --- Main Menu Loop ---

# Start a loop that continues until the user chooses to exit (option 10).
do {
  $choice = Show-Menu # Display the menu and get the user's choice.

  # Use a switch statement to execute the selected menu option.
  switch ($choice) {
    1 { Set-CalendarProcessingSettings -roomMailbox $roomMailbox } # Call the function for option 1.
    2 { Grant-EditorAccess -roomMailbox $roomMailbox } # Call the function for option 2.
    3 { Set-DefaultUserAccess -roomMailbox $roomMailbox } # Call the function for option 3.
    4 { Grant-SignageAccess -roomMailbox $roomMailbox } # Call the function for option 4.
    5 { Add-ResourceDelegates -roomMailbox $roomMailbox } # Call the function for option 5.
    6 { Remove-UserPermissions -roomMailbox $roomMailbox } # Call the function for option 6.
    7 { Show-CurrentPermissions -roomMailbox $roomMailbox } # Call the function for option 7.
    8 { Show-CalendarSettings -roomMailbox $roomMailbox } # Call the function for option 8.
    9 { Toggle-DebugMode } # Call the function for option 9.
    10 { Write-Host "Exiting script." } # Inform the user that the script is exiting.
    default { Write-Host "Invalid choice. Please enter a number between 1 and 10." -ForegroundColor Red } # Handle invalid input.
  }
} until ($choice -eq 10) # Continue the loop until the user chooses option 10 (Exit).

# --- Script End ---