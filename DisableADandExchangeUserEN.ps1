# Attention!!! This script works correctly only as Administrator from PowerShell ISE, running on a Windows domain controller.
# The DisableADandExchangeUser script takes the user's login and does the following:
# 1. Connects to the Exchange server integrated into Windows Active Directory and exports a copy of the mailbox at the specified path to a pst file
# 2. Disables the specified user's mailbox
# 3. Disables a user in a Windows AD domain

# Declare a variable to concatenate multiple requests to export an Exchange mailbox.
$BatchName = 'MassRequest'

# Address of the Microsoft Exchange server where MyExchange.ms.com is the name of the server on the local network.
# If you need to enter the address of the Exchange server in the console, you can uncomment the "$exServer = Read-Host ..." line.
# And comment out the line "$exServer = 'MyExchange.ms.com'"
# $exServer = Read-Host -Prompt 'Specify the name of the Exchange server: '
$exServer = 'MyExchange.ms.com'


# Clear the screen
Clear-Host

# ---- Step one:
# connect to the Microsoft Exchange server,
# here you will need to enter your domain administrator username and password
#
Write-Host (Get-Date) '| Connecting to the Exchange Server'
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

# ---- Step two:
# get the name of the user to disable
Write-Host (Get-Date) '| Get the user's AD name'
$adUser = Read-Host -Prompt 'Enter the name of the user to logout (ctrl+c to exit)'

# ---- Step three:
# Clear the history of requests to export the history of the Exchange server
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false
Get-MailboxExportRequest -Status Failed | Remove-MailboxExportRequest -Confirm:$false
# Start creating a backup copy of the user's mailbox
Write-Host (Get-Date) '| Create a backup copy of the user's mail'
# Create a request to create a backup with the address to save the archive
New-MailboxExportRequest -Mailbox $adUser -FilePath D:\Archive_Mailbox\$adUser.pst
# We are waiting for the end of the backup
Write-Host (Get-Date) '| Waiting for the end of the backup'
# Display copy progress as a percentage
while ((Get-MailboxExportRequest -Identity $adUser\MailboxExport | Where {($_.Status -eq “Queued”) -or ($_.Status -eq “InProgress”)})) {
    sleep 10
    Write-Host "Done: "(Get-MailboxExportRequest -Mailbox $adUser | Get-MailboxExportRequestStatistics).PercentComplete"%"
}
Get-MailboxExportRequest

# Check for box export error
if ((Get-MailboxExportRequest -Mailbox $adUser | Get-MailboxExportRequestStatistics).Status -eq "Failed") {
    Write-Host "Archiving error"
}


$ConfirmKey = Read-Host -Prompt 'Make sure the copy was created successfully (Y to disable mail / N to skip mail disable)'
if ($ConfirmKey -eq "Y"){

Write-Host "I turn off the mail"

# ---- Step Four: Deactivate Your Account
# To disable a user's archive mailboxes, uncomment the following line, "Disable-Mailbox -Identity $adUser -Archive..."
# and comment out the line "Disable-Mailbox -Identity $adUser -Confirm:$false"
# Disable-Mailbox -Identity $adUser -Archive -Confirm:$false 
Disable-Mailbox -Identity $adUser -Confirm:$false
}

# ---- Step Five: Update the Global Address List so that clients see the changes in the address book
Write-Host "Initiating a mailbook update"
Get-GlobalAddressList | Update-GlobalAddressList
Get-OfflineAddressBook | Update-OfflineAddressBook
Get-AddressList | Update-AddressList

# ---- Step six: Install the module and import it into the PS session. Disabling a user account
Write-Host "Deactivating account"
Import-Module ActiveDirectory
Disable-ADAccount -Identity $adUser
