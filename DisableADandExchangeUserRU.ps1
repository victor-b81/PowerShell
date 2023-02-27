# Внимание!!! Данный скрипт, корректно работает только от имени Администратора из PowerShell ISE, запущеный на контролере домена Windows.
# Скрипт DisableADandExchangeUser принимает логин пользователя и выполняет следующие действия:
#   1. Подключается к серверу Exchange интегрированному в Active Directory Windows и экспортирует копию почтового ящика по заданному пути в pst файл
#   2. Отключает почтовый ящик заданного пользователя
#   3. Отключает пользователя в домене AD Windows

# Объявляем переменную для объединения нескольких запросов на экспорт почтового ящика Exchange.
$BatchName = 'MassRequest'

# Адрес Microsoft Exchange сервера где MyExchange.ms.com имя сервера в локальной сети.
# Если будет нужно вводить адрес сервера Exchange в консоли, вы можете разкомментировать строку "$exServer = Read-Host ...".
# И закомментировать строку "$exServer = 'MyExchange.ms.com'"
# $exServer = Read-Host -Prompt 'Укажите имя Exchange сервера: '
$exServer = 'MyExchange.ms.com'

# Очищаем экран
Clear-Host

# ---- Шаг первый:
# подключаемся к серверу Microsoft Exchange,
# тут понадобиться ввести свое имя пользователя и пароль администратора домена
# 
Write-Host (Get-Date) '| Подключаемся к серверу Exchange'
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

# ---- Шаг второй:
# получаем имя отключаемого пользователя
Write-Host (Get-Date) '| Получаем имя AD пользователя'
$adUser = Read-Host -Prompt 'Введите имя отключаемого пользователя (ctrl+c чтобы выйти)'

# ---- Шаг третий:
# Очищаем историю запросов на экспорт истории Exchange сервера
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false
Get-MailboxExportRequest -Status Failed | Remove-MailboxExportRequest -Confirm:$false
# Запускаем создание резервной копии почтового ящика пользователя
Write-Host (Get-Date) '| Создаем резервную копию почты пользователя'
# Создаем запрос на создание резервной копии с указанием адреса сохранения архива
New-MailboxExportRequest -Mailbox $adUser -FilePath D:\Archive_Mailbox\$adUser.pst
# Ждем, окончания резервного копирования
Write-Host (Get-Date) '| Ждем, окончания резервного копирования'
# Отображаем прогресс копирования в процентах
while ((Get-MailboxExportRequest -Identity $adUser\MailboxExport | Where {($_.Status -eq “Queued”) -or ($_.Status -eq “InProgress”)})) {
    sleep 10
    Write-Host "Выполнено: "(Get-MailboxExportRequest -Mailbox $adUser | Get-MailboxExportRequestStatistics).PercentComplete"%"
}
Get-MailboxExportRequest

# Проверка на ошибку экспортирования ящика
if ((Get-MailboxExportRequest -Mailbox $adUser | Get-MailboxExportRequestStatistics).Status -eq "Failed") {
    Write-Host "Ошибка архивирования"
}


$ConfirmKey = Read-Host -Prompt 'Убедитесь, что копия создана успешно (Y чтобы отключить почту / N чтобы пропустить отключение почты)'
if ($ConfirmKey -eq "Y"){

Write-Host "Отключаю почту"

# ---- Шаг четвертый: Отключаем учетную запись
# Для отключения архивных почтовых ящиков пользователя раскомментируйте следующую строчку, "Disable-Mailbox -Identity $adUser -Archive..."
# а закомментируйте строку "Disable-Mailbox -Identity $adUser -Confirm:$false"
# Disable-Mailbox -Identity $adUser -Archive -Confirm:$false 
Disable-Mailbox -Identity $adUser -Confirm:$false
}

# ---- Шаг пятый: Обновляем Global Adress List, чтобы клиенты увидели изменения в адресной книге
Write-Host "Запускаю обдновление почтовой книги"
Get-GlobalAddressList | Update-GlobalAddressList
Get-OfflineAddressBook | Update-OfflineAddressBook
Get-AddressList | Update-AddressList

# ---- Шаг шестой: Установливаем модуль и импортируем его в сеанс PS. Отключаем учетную запись пользователя
Write-Host "Отключаю учетную запись"
Import-Module ActiveDirectory
Disable-ADAccount -Identity $adUser
