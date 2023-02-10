# Скрипт позволяет получить информацию о действиях в почтовом ящике пользователя и сохранить отчет в CSV файл
# Создан на основе скрипта предоставленного на странице https://learn.microsoft.com/ru-RU/microsoft-365/troubleshoot/audit-logs/mailbox-audit-logs
# Протестировано на Exchange 2010
# ВНИМАНИЕ!!! Перед использованием убедитесь, что аудит ящика включен
# 1. Проверить включен ли аудит ящика:      Get-Mailbox <useridentity> | ft AuditEnabled
# 2. Включить аудит ящика:                  Set-Mailbox <useridentity> -AuditEnabled $true
# 3. Указать что логируем:                  Set-Mailbox <useridentity> -AuditOwner "Create,HardDelete,Move,MoveToDeletedItems,SoftDelete,Update"
# Скрипт можно использовать
# Для отключения логирования:               Set-Mailbox <useridentity> -AuditOwner $none

<#
 Блок параметров скрипта, которые можно передавать указывая 
 где:
        -Mailbox                        Имя почтового ящика
        -StartDate                      Время с которого начинать поиск обьектов
        -EndDate                        Время до которого проводить поиск. Рекомендуется указывать + 1 день, чтобы вошло сегодняшнее число
        -Subject                        Тема письма, указывается в ковычках. Например .\AuditExchangeRUS.ps1 -Subject "<Good News>"
        -IncludeFolderBind              Включает в отчет, время доступа к ящику не владелцем.
        -ReturnObject                   Выводит информацию в консоль минуя создание файла отчета
#>
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


# Тело программы 

# В BEGIN задаем поля и получаем данные о пользователе и датах поиска.
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

# В PROCESS обрабатываем основную функцию программы
PROCESS {
  write-host -fore green 'Ищем логи Mailbox ...'
  $SearchResults = @(search-mailboxAuditLog $Mailbox -StartDate $StartDate -EndDate $EndDate -LogonTypes Owner, Admin, Delegate -ShowDetails -resultsize 50000)
  write-host -fore green 'Записей найдено: '$SearchREsults.Count
  if (-not $IncludeFolderBind)
  {
  write-host -fore green 'Исключаем FolderBind подробные данные доступов.'
  $SearchResults = @($SearchResults | ? {$_.Operation -notlike 'FolderBind'})
  write-host -fore green 'Исключено FolderBind записей: '$SearchREsults.Count
  }
  $SearchResults = @($SearchResults | select ($LogParameters + @{Name='Subject';e={if (($_.SourceItems.Count -eq 0) -or ($_.SourceItems.Count -eq $null)){$_.ItemSubject} else {($_.SourceItems[0].SourceItemSubject).TrimStart(' ')}}},
  @{Name='CrossMailboxOp';e={if (@('SendAs','Create','Update') -contains $_.Operation) {'N/A'} else {$_.CrossMailboxOperation}}}))
  $LogParameters = @('Subject') + $LogParameters + @('CrossMailboxOp')
  If ($Subject -ne '' -and $Subject -ne $null)
  {
  write-host -fore green 'Ищем тему письма: '$Subject
  $SearchResults = @($SearchResults | ? {$_.Subject -match $Subject -or $_.Subject -eq $Subject})
  write-host -fore green 'Найдено тем: '$($SearchREsults.Count)
  }
  $SearchResults = @($SearchResults | select $LogParameters)
}

# В END задаем конечную обработку отчета и выгрузку в csv файл
END {
if ($ReturnObject)
  {return $SearchResults}
elseif ($SearchResults.count -gt 0)
  {
  $Date = get-date -Format "dd-MM-yyyy_HH-mm"
  $OutFileName = $Mailbox+"_AuditLog_$Date.csv"
  write-host
  write-host -fore green "Сохраняем результаты в файл: $OutfileName"
  $SearchResults | export-csv $OutFileName -notypeinformation -encoding UTF8
  }
}
