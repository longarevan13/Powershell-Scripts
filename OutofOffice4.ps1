function ConnectToExchangeOnline {
    param (
        [switch]$IsReconnect = $false
    )

    if (-not $IsReconnect) {
        $UserPrincipalName = "administrator@henges.com"
        $Password = ConvertTo-SecureString "mwu4jMUj!" -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential($UserPrincipalName, $Password)

        # Install and import the Exchange Online PowerShell module in the user scope
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        Import-Module ExchangeOnlineManagement -Force

        # Connect to Exchange Online using basic authentication
        Connect-ExchangeOnline -Credential $Credential
    }
    else {
        Write-Host "Already connected to Exchange Online."
    }
}

function ConfigureOOF {
    param(
        [string]$Identity,
        [string]$AutoReplyState,
        [string]$InternalMessage,
        [string]$ExternalMessage
    )

    Set-MailboxAutoReplyConfiguration -Identity $Identity -AutoReplyState $AutoReplyState -InternalMessage $InternalMessage -ExternalMessage $ExternalMessage -ExternalAudience All
}

function ConfigureEmailForwarding {
    param (
        [string]$Identity,
        [string]$ForwardingAddress,
        [switch]$Disable = $false
    )

    if ($Disable) {
        Set-Mailbox -Identity $Identity -ForwardingAddress $null -DeliverToMailboxAndForward $false
    }
    else {
        Set-Mailbox -Identity $Identity -ForwardingAddress $ForwardingAddress -DeliverToMailboxAndForward $true
    }
}

# Initialize the $isExchangeConnected variable
$isExchangeConnected = $false

# Connect to Exchange Online
ConnectToExchangeOnline -IsReconnect:$isExchangeConnected
$isExchangeConnected = $true

while ($true) {
    # Display menu
    Write-Host "1. Configure Out of Office"
    Write-Host "2. Configure Email Forwarding"
    Write-Host "3. Exit"

    # Get user choice
    $choice = Read-Host "Enter your choice (1, 2, or 3)"

    switch ($choice) {
        1 {
            # Configure Out of Office
            $Identity = Read-Host "Enter the mailbox identity (e.g., admin@henges.com)"
            $choice2 = Read-Host "Enter Your Choice (1.Enable, 2.Disable)"

            switch ($choice2) {
                1 {
                    $InternalMessage = Read-Host "Enter internal auto-reply message"
                    $ExternalMessage = Read-Host "Enter external auto-reply message"
                    $AutoReplyState = "Enable"
                }
                2 {
                    $InternalMessage = "NA"
                    $ExternalMessage = "NA"
                    $AutoReplyState = "Disable"
                }
            }

            # Call the function to configure OOF with provided parameters
            ConfigureOOF -Identity $Identity -AutoReplyState $AutoReplyState -InternalMessage $InternalMessage -ExternalMessage $ExternalMessage
            Write-Host "Out of Office configuration complete."
        }
        2 {
            # Configure Email Forwarding
            $Identity = Read-Host "Enter the mailbox identity (e.g., user@henges.com)"
            $DisableForwarding = $false

            $choice2 = Read-Host "Enter Your Choice (1.Enable, 2.Disable)"

            if ($choice2 -eq 2) {
                $DisableForwarding = $true
            }

            if (-not $DisableForwarding) {
                $ForwardingAddress = Read-Host "Enter the email address to forward messages to"
            }

            # Call the function to configure email forwarding
            ConfigureEmailForwarding -Identity $Identity -ForwardingAddress $ForwardingAddress -Disable $DisableForwarding
            Write-Host "Email forwarding configuration complete."
        }
        3 {
            # Exit the script
            Write-Host "Exiting the script."
            return
        }
        default {
            Write-Host "Invalid choice. Please try again."
        }
    }
Write-Host "Press enter to exit ..."
$null = Read-Host
}
