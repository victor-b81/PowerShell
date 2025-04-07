##########Powershell скрипт резервного копирования Microsoft Hyper-V (локального)##########

#Объявляем переменные

# vmNum - номер виртуальной машины в в массиве.
# TimeStampDIR - переменная, содержащая дату начала работы скрипта.
# LastBackup - переменная, содержащая дату последнего резервного копирования.
# vmNameArray - массив названий виртуальных машин, определенных к резервному копированию
# BackupCount - размер массива виртуальных машин.
# BackupDir - локальный каталог, размещения резервных копий
# BackupPath - путь сохранения текущей резервной копии 
# ExportLogName - имя файла легирования текущей резервной копии
# vmCopylog - файл лога, процесса копирования
# smtpServer - переменная содержащая адрес сервера отправки почты
# encoding - текстовая кодировка письма (подробнее https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.5)
# smtpPort - порт сервера отправки
# from - переменная содержащая email отправителя 
# to - переменная содержащая email получателя
# attachment - переменная содержащая локальный адрес вложения
# smtpUsername - имя пользователя, для доступа к серверу отправки писем
# smtpPassword - пароль (в открытой форме) для доступа к серверу отправки писем
# smtpSecurePassword - команда преобразования защиты пароля
# smtpCredential - команда инициализации подключения к серверу отправки писем
# SMBpasswd - преобразованный пароль пользователя с доступом к сетевому ресурсу
# SMBcreds - команда инициализации подключения к сетевому ресурсу
# SMBdriveLetter - буква, монтируемого сетевого ресурса
# SMBdrivePath - путь к монтируемому диску
# SMBcopyMarker - файл-маркер создаваемый на время процесса копирования
# SMBshareDir - сетевой ресурс для размещения текущей резервной копии

# Блок переменных для работы с резервным копированием:
$vmNum = 0
$TimeStampDIR = $(get-date -f dd-MM-yyyy)
$LastBackup = Get-Date $(get-date).Adddays(-7) -Format "dd-MM-yyyy"
$vmNameArray = @("VM1", "VM2")
$BackupCount = ($vmNameArray.count -1)
$BackupDir = "D:\BackUp"
$BackupPath = "$BackupDir\$TimeStampDIR\"
$ExportLogName = "vmExport_log_$TimeStampDIR.log"
$vmCopylog = "$BackupDir\Logs\vmCopylog_$TimeStampDIR.log"

# Блок переменных для отправки почты:
$smtpServer = "smtp.mail.ru"
$encoding = "oem"
$smtpPort = 587
$from = "Отправитель@mail.ru"
$to = "Получатель@mail.ru"
$attachment = "$BackupPath\$ExportLogName"
$smtpUsername = "Имя пользователя"
$smtpPassword = "Пароль"
$smtpSecurePassword = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
$smtpCredential = New-Object System.Management.Automation.PSCredential($smtpUsername, $smtpSecurePassword)

# Блок переменных для подключения сетевой шары:
$SMBshareDir = "\\СерверРезервногоКопирования\Папка"
$SMBpasswd = ConvertTo-SecureString "Пароль Пользователя" -AsPlainText -Force
$SMBcreds = New-Object System.Management.Automation.PSCredential ("Имя пользователя", $SMBpasswd)
$SMBdriveLetter = "Буква диска"
$SMBdrivePath = $SMBdriveLetter+":\"
$SMBcopyMarker = $SMBdrivePath + "start_copy.marker.txt"


#Начало выполнения процедур

#Проверяем наличие каталога резервного копирования. Если каталога нет, сообщаем о ошибке и завершаем скрипт.
if (Test-Path -Path "$BackupDir"){
    #Создаем новый каталог, используя в качестве имени текущую дату запуска скрипта.
    if (!(Test-Path -Path "$BackupDir\$TimeStampDIR")){
        New-Item -Path "$BackupDir\" -Name "$TimeStampDIR" -ItemType "Directory"
    }

    #Создаем новые файлы логирования.
    if (!(Test-Path -Path $BackupPath\$ExportLogName)){
        New-Item -Path $BackupPath\$ExportLogName -ItemType "File" -Force
        New-Item -Path $vmCopylog -ItemType "File" -Force
    }

    #Запускаем процесс создания резервных копий.
    $(get-date -f "dd-MM-yyyy HH:mm")+" Начало резервного копирования" >> $BackupPath\$ExportLogName
    
    while($vmNum -le $BackupCount){
        $vmName = $vmNameArray[$vmNum]
            $ExportJob = Export-VM -Name $vmName -Path $BackupPath -Asjob
            $ExportJob | Wait-Job
        Write-Output ($(get-date -f "dd-MM-yyyy HH:mm") + ";" + $vmName + ";" + $($ExportJob.Progress.PercentComplete) + "%;" + $($ExportJob.State)) >> $BackupPath\$ExportLogName

        #Если резервное копирование завершилось неудачей, переносим последнюю удачную копию в каталог D:\LastGoodCopyVM.
        if ($ExportJob.State -ne "Completed") {
            New-Item -Path "D:\LastGoodCopyVM\" -Name "$LastBackup" -ItemType "Directory"
            Move-Item -Path D:\BackUp\$LastBackup\$vmName -Destination D:\LastGoodCopyVM\$LastBackup\$vmName
        }
    $vmNum++;
    }
    $(get-date -f "dd-MM-yyyy HH:mm")+" Конец резервного копирования" >> $BackupPath\$ExportLogName

    #Проверяем наличие доступа к сетевому ресурсу для размещения резервных копий и монтируем сетевой ресурс в системе
    New-PSDrive -Name $SMBdriveLetter -Root $SMBshareDir -Persist -PSProvider "FileSystem" -Credential $SMBcreds

    #Если сетевой ресурс доступен:
    # 1. Удаляем копии в сетевом хранилище, старше двух недель (-14)
    # 2. Создаем файл-маркер перед началом копирования
    # 3. Копируем резервные копии
    # 4. Удаляем файл-маркер после копирования
    # 5. Создаем/перезаписываем файл-маркер указывающий на существование новой резервной копии в сетевом ресурсе

    If(Get-PSDrive | Where-Object DisplayRoot -EQ $SMBshareDir){
        Get-ChildItem -Path "$SMBdrivePath" -Directory -recurse| where {$_.LastWriteTime -le $(get-date).Adddays(-14)} | Remove-Item -recurse -force
        New-Item -Path $SMBcopyMarker -ItemType "File" -Force
        robocopy $BackupPath "$SMBdrivePath$TimeStampDIR" /E /J /B /R:3 /W:20 /NP /LOG:$vmCopylog
        Remove-Item -Path $SMBcopyMarker -Force
        "Есть новая резервная копия от $TimeStampDIR" > $SMBdrivePath\new_copy.marker.txt
        Remove-PSDrive $SMBdriveLetter
    } else { 
        "ОШИБКА - сетевой ресурс недоступен" >> $vmCopylog
    }

    #Проверяем наличие ошибок при резервном копировании.
    $ErrorBackUpChecking = Select-String -Path $BackupPath\$ExportLogName -Pattern "Failed"
    $ErrorCopyChecking = Select-String -Path $vmCopylog -Pattern "ОШИБКА"

       if ($ErrorBackUpChecking -ne $null) {
         Send-MailMessage -From $from -To $to -Subject "Ошибка в процессе создания резервных копий HV01." `
         -Body "Ошибка в процессе резервного копирования HV01. Последняя удачная копию перемещена каталог D:\LastGoodCopyVM" `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Attachments $attachment -Encoding $encoding
        }

       if ($ErrorCopyChecking -ne $null) {
          Send-MailMessage -From $from -To $to -Subject "Ошибка в процессе копирования резервных копий HV01." `
         -Body "Ошибка в процессе копирования резервных копий HV01 в сетевой ресурс $SMBshareDir" `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Attachments $attachment -Encoding $encoding
       }

    #Удаляем локальные копии старше одного дня (-1).
    Get-ChildItem -Path "D:\BackUp\" -Directory -recurse| where {$_.LastWriteTime -le $(get-date).Adddays(-1)} | Remove-Item -recurse -force
}
else
{
    #Если каталог резервного копирования недоступен отправляем сообщение о ошибке.
    Send-MailMessage -From $from -To $to -Subject "Ошибка резервного копирования HV01." `
         -Body "Ошибка резервного копирования HV01. Недоступен каталог D:\BackUp\" `
         -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredential `
         -UseSsl -Encoding $encoding
}