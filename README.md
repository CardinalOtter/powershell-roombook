## 🗄️ Exchange Online Room Mailbox Configuration Script

This script automates the configuration of Exchange Online room mailboxes, providing a menu-driven interface for common tasks.

**🚧 Development Status: This project is currently in its development stage.**

### Features

*   Interactive Menu: 💻 Guides the user through configuration steps.
*   Exchange Online Integration: 🔗 Connects to Exchange Online using administrator credentials (with MFA support).
*   Room Mailbox Configuration:
    *   Set `AddOrganizerToSubject` and `DeleteSubject` to `$false`.
    *   Grant "Editor" access to a coordinator group.
    *   Set default user access rights to "Reviewer".
    *   Grant "Reviewer" access to a digital signage group.
    *   Add resource delegates.
    *   Remove existing permissions.
    *   Show current permissions.
    *   Show current calendar settings.
*   Input Validation: ✅ Verifies the existence of the room mailbox.
*   Error Handling: ⚠️ Uses `try-catch` blocks and checks for non-terminating errors.
*   Debug Mode: 🐛 Provides verbose output for troubleshooting (🔍).
*   Module Installation: Checks for and installs the `ExchangeOnlineManagement` module if needed.
*   Execution Policy Handling: Checks and sets the execution policy to `RemoteSigned` if necessary.

### Prerequisites

*   `ExchangeOnlineManagement` PowerShell module. The script will prompt for installation if it's not found.
*   Exchange Online administrator credentials (with MFA support).
*   PowerShell execution policy set to `RemoteSigned`. The script will prompt to set this if necessary.

### Usage

1.  Run the script: `.\Configure-RoomMailbox.ps1`
2.  Follow the prompts to enter your Exchange Online admin UPN and the room mailbox email address.
3.  Use the interactive menu to configure the room mailbox:
    1.  Set Calendar Processing Settings
    2.  Grant Editor Access to Coordinator Group
    3.  Set Default User Access Rights to Reviewer
    4.  Grant Reviewer Access to Digital Signage Group
    5.  Add Resource Delegates
    6.  Remove Existing Permissions for a User
    7.  Show Current Permissions
    8.  Show Calendar Settings
    9.  Enable/Disable Debug Mode
    10. Exit

### Error Handling

The script uses `try-catch` blocks to handle potential errors. User-friendly error messages are displayed in red if an error occurs. Verbose output is available in debug mode for more detailed error information.