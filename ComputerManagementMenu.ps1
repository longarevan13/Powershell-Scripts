function ConfigureOOF {
    param(
        [string]$Identity,
        [string]$AutoReplyState,
        [string]$InternalMessage,
        [string]$ExternalMessage
    )

    Set-MailboxAutoReplyConfiguration -Identity $Identity -AutoReplyState $AutoReplyState -InternalMessage $InternalMessage -ExternalMessage $ExternalMessage -ExternalAudience All
}

function ConnectToExchangeOnline {
    param (
        [switch]$IsReconnect = $false
    )

    if (-not $IsReconnect) {
        $UserPrincipalName = Read-Host "Enter your Exchange Online admin account (e.x., admin@henges.com)"
        
        # Install and import the Exchange Online PowerShell module
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
        Import-Module ExchangeOnlineManagement -Force

        # Connect to Exchange Online
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
    }
    else {
        Write-Host "Already connected to Exchange Online."
    }
}

function CheckSystemUptime {
    $pcname = Read-Host "Enter PC Name (E.g., jhe-ad-pk)"

    # Specify the full path to psexec.exe
    $psexecPath = "C:\PStools\psexec.exe"

    # Check if psexec.exe exists
    if (-not (Test-Path $psexecPath -PathType Leaf)) {
        Write-Host "Error: psexec.exe not found at the specified path."
        return
    }

    # Generate a unique file name for the script
    $scriptFile = [System.IO.Path]::Combine($env:TEMP, "Script-{0}.ps1" -f (Get-Random))

    # Create the PowerShell script file
    $scriptContent = @"
    `$bootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    `$uptime = (Get-Date) - `$bootTime
    Write-Output ('System Uptime on {0}: {1} days, {2} hours, {3} minutes' -f `$env:COMPUTERNAME, `$uptime.Days, `$uptime.Hours, `$uptime.Minutes)
"@

# Run psexec.exe with the correct syntax to execute the PowerShell script content
$result = & $psexecPath -s \\$pcname powershell -ExecutionPolicy Bypass -Command "& { $scriptContent }" 2>&1

# Display the result without adding a newline
Write-Host -NoNewline ($result -join "`n")
}

# Main script
Write-Host "1. Computer Management (WIP)"
Write-Host "2. Exchange"
$choice3 = Read-Host "Enter Your Choice (1, 2)"
Switch ($choice3) {
    1 {
        CheckSystemUptime -ComputerName ($pcname)
    }
    2 {
        $isExchangeConnected = $false

        # Loop for Exchange options
        while ($true) {
            # Display menu
            Write-Host "Out of Office Configuration Script"
            Write-Host "1. Connect to ExchangeOnline"
            Write-Host "2. Configure Out of Office"
            Write-Host "3. Exit"

            # Get user choice
            $choice = Read-Host "Enter your choice (1, 2, or 3)"
            $choice2 = Read-Host "Enter Your Choice (1.Enable, 2.Disable)"

            switch ($choice) {
                1 {
                    ConnectToExchangeOnline -IsReconnect:$isExchangeConnected
                    $isExchangeConnected = $true
                }
                2 {
                    # Get input from the user
                    $Identity = Read-Host "Enter the mailbox identity (e.g., admin@henges.com)"
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
                    # Call the function to configure OOF
                    ConfigureOOF -Identity $Identity -AutoReplyState $AutoReplyState -InternalMessage $InternalMessage -ExternalMessage $ExternalMessage

                    Write-Host "Out of Office configuration complete."
                }
                3 {
                    Write-Host "Exiting the script."
                    break
                }
                default {
                    Write-Host "Invalid choice. Please try again."
                }
            }
        }
    }
}
