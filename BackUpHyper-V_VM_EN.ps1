##########PowerShell script for backup Microsoft Hyper-V (local)##########

#Declare variables

# vmNum - number of virtual machine in the array.
# TimeStampDIR - a variable containing the start date of the script.
# LastBackup - variable containing the date of the last backup.
# vmNameArray - array of names of virtual machines defined for backup
# BackupCount - size of the array of virtual machines.
# BackupDir - local directory where backups are located
# BackupPath - path to save the current backup 
# ExportLogName - name of the current backup file
# vmCopylog - log file of the copy process
# smtpServer - variable containing the address of the mail sending server
# encoding - text encoding of the mail (for more information https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7. 5)
# smtpPort - port of sending server
# from - variable containing sender's email 
# to - variable containing recipient's email
# attachment - variable containing local address of attachment
# smtpUsername - user name, to access the mail sending server
# smtpPassword - password (in open form) to access the mail sending server
# smtpSecurePassword - password protection conversion command
# smtpCredential - command to initialize connection to the mail sending server
# SMBpasswd - converted password of the user with access to the network share
# SMBcreds - command to initialize connection to the network share
# SMBdriveLetter - letter, of the mounted network share
# SMBdrivePath - path to the mounted disk
# SMBcopyMarker - marker file created during the copying process
# SMBshareDir - network share for placing the current backup

# Block of variables for backup work:
$vmNum = 0
$TimeStampDIR = $(get-date -f dd-MM-yyyy)
$LastBackup = Get-Date $(get-date).Adddays(-7) -Format "dd-MM-yyyy"
$vmNameArray = @("VM1", "VM2")
$BackupCount = ($vmNameArray.count -1)
$BackupDir = "D:\BackUp"
$BackupPath = "$BackupDir\$TimeStampDIR\"
$ExportLogName = "vmExport_log_$TimeStampDIR.log"
$vmCopylog = "$BackupDir\Logs\vmCopylog_$TimeStampDIR.log"

# Variable block for sending mail:
$smtpServer = "smtp.mail.com"
$encoding = "oem"
$smtpPort = 587
$from = "Sender@mail.com"
$to = "Recipient@mail.com"
$attachment = "$BackupPath\$ExportLogName"
$smtpUsername = "Username"
$smtpPassword = "Password"
$smtpSecurePassword = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
$smtpCredential = New-Object System.Management.Automation.PSCredential($smtpUsername, $smtpSecurePassword)

# Variable block for connecting the network share:
$SMBshareDir = "{\ServerSecureStorage/Folder"
$SMBpasswd = ConvertTo-SecureString "User Password" -AsPlainText -Force
$SMBcreds = New-Object System. Management.Automation.PSCredential ("Username", $SMBpasswd)
$SMBdriveLetter = "Drive Letter"
$SMBdrivePath = $SMBdriveLetter+":\"
$SMBcopyMarker = $SMBdrivePath + "start_copy.marker.txt"


#Start executing procedures

#Check if there is a backup catalog. If there is no catalog, report an error and terminate the script.
if (Test-Path -Path "$BackupDir"){
    #Create a new directory using the current script start date as the name.
    if (!(Test-Path -Path "$BackupDir\$TimeStampDIR")){
        New-Item -Path "$BackupDir\" -Name "$TimeStampDIR" -ItemType "Directory"
    }

    #Create new logging files.
    if (!(Test-Path -Path $BackupPath\$ExportLogName)){
        New-Item -Path $BackupPath\$ExportLogName -ItemType "File" -Force
        New-Item -Path $vmCopylog -ItemType "File" -Force
    }

    #Start the backup process.
    $(get-date -f "dd-MM-yyyy HH:mm")+" Starting a backup" >> $BackupPath\$ExportLogName
    
    while($vmNum -le $BackupCount){
        $vmName = $vmNameArray[$vmNum]
            $ExportJob = Export-VM -Name $vmName -Path $BackupPath -Asjob
            $ExportJob | Wait-Job
        Write-Output ($(get-date -f "dd-MM-yyyy HH:mm") + ";" + $vmName + ";" + $($ExportJob.Progress.PercentComplete) + "%;" + $($ExportJob.State)) >> $BackupPath\$ExportLogName

        #If the backup fails, move the last successful backup to the D:\LastGoodCopyVM directory.
        if ($ExportJob.State -ne "Completed") {
            New-Item -Path "D:\LastGoodCopyVM\" -Name "$LastBackup" -ItemType "Directory"
            Move-Item -Path D:\BackUp\$LastBackup\$vmName -Destination D:\LastGoodCopyVM\$LastBackup\$vmName
        }
    $vmNum++;
    }
    $(get-date -f "dd-MM-yyyy HH:mm")+" End of backup" >> $BackupPath\$ExportLogName

    #Check if there is access to a network share to host backups and mount the network share on the system
    New-PSDrive -Name $SMBdriveLetter -Root $SMBshareDir -Persist -PSProvider "FileSystem" -Credential $SMBcreds

    #If a network share is available:
	# 1. Delete copies on the NAS older than two weeks (-14)
	# 2. Create a marker file before starting to copy
	# 3. Copy backups
	# 4. Delete the marker file after copying
	# 5. Create/overwrite a marker file indicating the existence of a new backup in the network share

    If(Get-PSDrive | Where-Object DisplayRoot -EQ $SMBshareDir){
        Get-ChildItem -Path "$SMBdrivePath" -Directory -recurse| where {$_.LastWriteTime -le $(get-date).Adddays(-14)} | Remove-Item -recurse -force
        New-Item -Path $SMBcopyMarker -ItemType "File" -Force
        robocopy $BackupPath "$SMBdrivePath$TimeStampDIR" /E /J /B /R:3 /W:20 /NP /LOG:$vmCopylog
        Remove-Item -Path $SMBcopyMarker -Force
        "There is a new backup from $TimeStampDIR" > $SMBdrivePath\new_copy.marker.txt
        Remove-PSDrive $SMBdriveLetter
    } else { 
        "ERROR - network resource is unavailable" >> $vmCopylog
    }

    #Check for backup errors.
    $ErrorBackUpChecking = Select-String -Path $BackupPath\$ExportLogName -Pattern "Failed"
    $ErrorCopyChecking = Select-String -Path $vmCopylog -Pattern "ERROR"

       if ($ErrorBackUpChecking -ne $null) {
         Send-MailMessage -From $from -To $to -Subject "Error in the HV01 backup process." `
         -Body "Error during HV01 backup process. The last successful copy was moved to the D:\LastGoodCopyVM directory." `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Attachments $attachment -Encoding $encoding
        }

       if ($ErrorCopyChecking -ne $null) {
          Send-MailMessage -From $from -To $to -Subject "Error in the HV01 backup copy process." `
         -Body "Error while copying HV01 backups to network share $SMBshareDir" `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Attachments $attachment -Encoding $encoding
       }

    #Delete local copies older than one day (-1).
    Get-ChildItem -Path "D:\BackUp\" -Directory -recurse| where {$_.LastWriteTime -le $(get-date).Adddays(-1)} | Remove-Item -recurse -force
}
else
{
    #If the backup directory is unavailable, send an error message.
    Send-MailMessage -From $from -To $to -Subject "Backup error HV01." `
         -Body "Backup error HV01. The D:\BackUp\ directory is not available." `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Encoding $encoding
}