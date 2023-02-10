# The script allows you to get information about the actions in the user's mailbox and save the report to a CSV file
# Created based on the script provided on the page https://learn.microsoft.com/en-US/microsoft-365/troubleshoot/audit-logs/mailbox-audit-logs
# Tested on Exchange 2010
# ATTENTION!!! Make sure box auditing is enabled before use
# 1. Check if mailbox auditing is enabled: Get-Mailbox <useridentity> | ft AuditEnabled
# 2. Enable mailbox auditing: Set-Mailbox <useridentity> -AuditEnabled $true
# 3. Specify what to log: Set-Mailbox <useridentity> -AuditOwner "Create,HardDelete,Move,MoveToDeletedItems,SoftDelete,Update"
# The script can be used
# To disable logging: Set-Mailbox <useridentity> -AuditOwner $none

<#
Script parameters that can be passed by specifying
 Where:
        -Mailbox                Mailbox name
        -StartDate              Time from which to start searching for objects
        -EndDate                Time until which to search. It is recommended to specify + 1 day to include today's date
        -Subject                The subject of the email, enclosed in quotes. For example .\AuditExchangeRUS.ps1 -Subject "<Good News>"
        -IncludeFolderBind      Include in the report when the non-owner accessed the box.
        -ReturnObject           Prints information to the console without creating a report file
#>

# Optional script parameters
param ([PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[string]$Mailbox,
[PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[string]$StartDate,
[PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[string]$EndDate,
[PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[string]$Subject,
[PARAMETER(Mandatory=$False,ValueFromPipeline=$FALSE)]
[switch]$IncludeFolderBind,
[PARAMETER(Mandatory=$False,ValueFromPipeline=$FALSE)]
[switch]$ReturnObject)

# Body of the program

# In BEGIN, we set the fields and get data about the user and search dates.
BEGIN {
  [string[]]$LogParameters = @('Operation', 'LogonUserDisplayName', 'LastAccessed', 'DestFolderPathName', 'FolderPathName', 'ClientInfoString', 'ClientIPAddress', 'ClientMachineName', 'ClientProcessName', 'ClientVersion', 'LogonType', 'MailboxResolvedOwnerName', 'OperationResult')
  if ($Mailbox -eq "") {
      $Mailbox = Read-Host -Prompt 'Введите имя ящика (ctrl+c чтобы выйти)'
  }
  if ($StartDate -eq "") {
      $StartDate = Read-Host -Prompt 'Введите дату начала поиска. Образец 01-01-1999 (ctrl+c чтобы выйти)'
  }
  if ($EndDate -eq "") {
      $EndDate = Read-Host -Prompt 'Введите дату окончания поиска. Образец 01-01-1999 (ctrl+c чтобы выйти)'
  }
}
# In PROCESS we process the main function of the program
PROCESS {
  write-host -fore green 'Searching Mailbox Audit Logs...'
  $SearchResults = @(search-mailboxAuditLog $Mailbox -StartDate $StartDate -EndDate $EndDate -LogonTypes Owner, Admin, Delegate -ShowDetails -resultsize 50000)
  write-host -fore green 'Total entries Found: '$($SearchREsults.Count)
  if (-not $IncludeFolderBind)
  {
  write-host -fore green 'Removing FolderBind operations.'
  $SearchResults = @($SearchResults | ? {$_.Operation -notlike 'FolderBind'})
  write-host -fore green 'Filtered to Entries: '$($SearchREsults.Count)
  }
  $SearchResults = @($SearchResults | select ($LogParameters + @{Name='Subject';e={if (($_.SourceItems.Count -eq 0) -or ($_.SourceItems.Count -eq $null)){$_.ItemSubject} else {($_.SourceItems[0].SourceItemSubject).TrimStart(' ')}}},
  @{Name='CrossMailboxOp';e={if (@('SendAs','Create','Update') -contains $_.Operation) {'N/A'} else {$_.CrossMailboxOperation}}}))
  $LogParameters = @('Subject') + $LogParameters + @('CrossMailboxOp')
  If ($Subject -ne '' -and $Subject -ne $null)
  {
  write-host -fore green 'Searching for Subject: '$Subject
  $SearchResults = @($SearchResults | ? {$_.Subject -match $Subject -or $_.Subject -eq $Subject})
  write-host -fore green 'Filtered to Entries: '$SearchREsults.Count
  }
  $SearchResults = @($SearchResults | select $LogParameters)
}

# In END we set the final processing of the report and uploading to a csv file
END {
  if ($ReturnObject)
  {return $SearchResults}
  elseif ($SearchResults.count -gt 0)
  {
  $Date = get-date -Format "dd-MM-yyyy_HH-mm"
  $OutFileName = $Mailbox+"_AuditLog_$Date.csv"
  write-host
  write-host -fore green "Posting results to file: $OutfileName"
  $SearchResults | export-csv $OutFileName -notypeinformation -encoding UTF8
  }
}
