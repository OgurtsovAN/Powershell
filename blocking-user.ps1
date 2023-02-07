################################################################
#Имя:      .blocking(domen1)new.ps1 ver 05.10.2022
#Язык:     PowerShell 
#Описание: скрипт вызывается из Lotus Notes при выполнении заявок на блокирование  в базе "Доступ к АБС".
#          скрипт запускается только для домена domen1
################################################################
#domen1\s-Notes



param (
    
    [string]$documnumber    = $Null, # номер служебной записки no
    [string]$blockingreason = $Null, # причина блокирования
    [string]$dismissdate    = $Null, # дата увольнения date
    [string]$surname    = $Null, # Фамилия;
    [string]$name       = $Null, # имя пользователя
    [string]$patronymic = $Null, # Отчество;
    [string]$TabNumber =  $Null,  # Табельный номер;
    [string]$NewTabNum =  $Null  # Новый табельный номер;     
     
    )


Import-Module ActiveDirectory

# --- переменные ---
 
# список обязательных входных параметров
$ParamList = "documnumber", "blockingreason", "dismissdate", "surname", "name", "patronymic"

# родительский каталог
$ParentFolder = $MyInvocation.MyCommand.Definition | split-path -parent
 
# путь к лог-файлу
$LogFilePath = $ParentFolder +  "\Logs\blocking(domen1)new.txt"

# путь к каталогу хранения ошибок
$errorFilesPath = $ParentFolder +  "\Logs\Errors"

# путь к логу резолюции для лотуса
$LotusPath = $ParentFolder +  "\Logs\stdout.txt"

# FQDN имя сервера
$ComputerName = "$env:computername.$env:userdnsdomain"

# список шар с перемещаемыми профилями пользователей
$UsersRoamingProfiles = "\\SRVFILES-CRT\UsersData2","\\SRVFILES-KRML\UsersData2","\\srvfiles2\usersdata2","\\SRVFILES2\XA-UserProfiles","\\SRVFILES2\usersdata1","\\SRVFILESGRK\UsersData","\\SRVFILESESC\UsersData2"

# домен
$Domain = "domen1.akbars.ru"

# Ключ для паролей
$key_path = "F:\abbwin\cyph_keys\aes.key"

$SMTPMessage = New-Object System.Net.Mail.MailMessage

#---------------------------------------
# функции

### записываем лог в файл
function WriteToLogFile ([string]$strMsg)
{
try { Add-Content -Path $LogFilePath -Value $strMsg -Encoding UTF8 }
catch{}
}

### отправка smtp сообщения 
function SendEmail ([string]$strMessage,$recipient)
{   
$sender = 'noreply@akbars.ru'
$recipient += 'aydar@akbars.ru','Tisovski@akbars.ru','OgurtsovAN@akbars.ru'
$Subject = "Блокирование УЗ пользователей по заявке № $documnumber"
$SMTPServer = '192.168.8.243'
$Body = "Данное уведомление отправлено вам автоматически от 'Робот ОЭРМиСА'.`n`n"
$Body += $strMessage

     try { Send-MailMessage -To $recipient -From $sender -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Encoding utf8}  
     catch 
     {
      # текст ошибки
      $ErrMsg = $error[0]|format-list -force | Out-String
              
      # очищаем предыдущие ошибки
      $error.clear()
      # пишем ошибку в файл
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";ошибка отправки smtp сообщения $ErrMsg")
     }
}

### отправка smtp сообщения о удалении профиля сотрудника
function SendEmailProfile ([string]$recipient)
{ 
$sender = 'noreply@akbars.ru'  
$Subject = "Содержимое хранилища пользователя "+$user.Name+" будет доступно в течение 14 дн"
$SMTPServer = '192.168.8.243'
$Body = "Учетная запись пользователя "+$user.Name+" отключена из службы каталогов Active Directory по причине увольнения данного сотрудника <br />"
$Body += "Хранилище Перемещаемого профиля этого пользователя будет доступно в течение 14 дн. <br />
Если вы хотите сохранить содержимое до истечения 14-дневного срока хранения, необходимо написать заявку на перенос данных данного сотрудника в Jira по ссылке и скопировать данное сообщение в описание проблемы:
https://team.akbars.ru/servicedesk/customer/portal/4/create/87 <br /> "
$Body += "Через 14 дн. хранилище Перемещаемого профиля пользователя "+$user.name+" будет удалено навсегда, без возможности восстановления. <br />"
$Body += ((($UsersRP).where{$_.name -like $user.sAMAccountName}).fullname)
     try { Send-MailMessage -To $recipient -From $sender -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Encoding utf8 -BodyAsHtml}
     catch {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";ошибка отправки smtp сообщения")}   
}


############################################
#набор функций для генерации сложного пароля

# функция возвращает специальные символы
function Get-NonAlphaChars {
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true,Position=0)][int]$length,[Parameter(Mandatory=$true,Position=0)][int]$ratio)
    if ($ratio -lt 1){$ratio = 1}
    $no = 1..$ratio | Get-Random
    $special =  @"
~!@#$%^&*()_-='"|\{}[]:<>.,/
"@
    ($special.ToCharArray() | Get-Random -Count $no) -join ""
}

# функция возвращает числа
function Get-NumericChars {
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true,Position=0)][int]$length,[Parameter(Mandatory=$true,Position=0)][int]$ratio)
    if ($ratio -lt 1){$ratio = 1}
    $no = 1..$ratio | Get-Random
    $numericChars = (0..9) -join ""
    ($numericChars.ToCharArray() | Get-Random -Count $no) -join ""
}

# функция возвращает символы алфавита в верхнем регистре
function Get-UpperChars {
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true,Position=0)][int]$length,[Parameter(Mandatory=$true,Position=0)][int]$ratio)
    if ($ratio -lt 1){$ratio = 1}
    $no = 1..$ratio | Get-Random
    $upperChars = ((65..90) | ForEach-Object {[char]$_}) -join ""
    ($upperChars.ToCharArray() | Get-Random -Count $no) -join ""
}

# функция возвращает символы алфавита в нижнем регистре
function Get-LowerChars {
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true,Position=0)][int]$no)
    $lowerChars = ((97..122) | ForEach-Object {[char]$_}) -join ""
    ($lowerChars.ToCharArray() | Get-Random -Count $no) -join ""
}


# Функция генерации пароля, соответствующего требованиям (20 симв.,включать буквы (большие и маленькие), цифры и спец символы.)
Function GenerateComplexPassword ([Parameter(Mandatory=$true)][int]$passLen)
{
    $newPassword = @()

    $lengthRemaining = $passLen
    $ratio = $passLen / 4

    # генеирируем специальные символы
    $nonAlpha = Get-NonAlphaChars -length $passLen -ratio $ratio
    $newPassword += $nonAlpha
    $lengthRemaining = $lengthRemaining - $nonAlpha.Length

    # генеирируем числа
    $numericChars = Get-NumericChars -length $passLen -ratio $ratio
    $newPassword += $numericChars
    $lengthRemaining = $lengthRemaining - $numericChars.Length

    # генеирируем символы алфавита в верхнем регистре
    $upperChars = Get-UpperChars -length $passLen -ratio $ratio
    $newPassword += $upperChars
    $lengthRemaining = $lengthRemaining - $upperChars.Length

    # оставшуюся длину заполняем символами нижнего регистра
    if($lengthRemaining -lt 1){$lengthRemaining = 1}
    $newPassword += Get-LowerChars -no $lengthRemaining

    # перемешиваем строку
    $result = (($newPassword.ToCharArray() | Get-Random -Count $passLen) -join "").Replace(" ","")

    return $result
}

############################################
#конец набора функций для генерации сложного пароля

### установка случайного пароля для учетной записи пользователя
Function SetUserComplexPassword ([Parameter(Mandatory=$true)][string]$userName ,[string]$Domain, $Credential) {

    # все ошибки прерывающие
    $ErrorActionPreference = "stop"

    # подключаем скрипт для генерации сложного пароля. Назначаем пароль пользователю
    try {       
        Set-ADAccountPassword $userName -Server $Domain -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $(GenerateComplexPassword(20)) -Force -Verbose) -Credential $cred -PassThru
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Устанавлен сложный пароль для " +$username) }        
     catch {
        Write-Host "ERROR"
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Установить сложный пароль для $username не удалось")
         SendEmail -strMessage "Установить сложный пароль для $username не удалось" -recipient $ADDrecipient
          }
}
 


### функция отключения п\я, постановка на удержание и перемещению mailbox в другую БД
function SetMailboxDisabled([string]$uvolen)
{
try
{
          
          $exch_pwd_path = "F:\abbwin\cyph_keys\pwd_exchdis.txt"
          $exch_password = Get-Content $exch_pwd_path | ConvertTo-SecureString -Key (Get-Content $key_path)
          $Cred = New-Object System.Management.Automation.PsCredential("domen1\s-exch",$exch_password) 


$Session_mlx01 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://p-exch.domen1.akbars.ru/PowerShell/ -Authentication Kerberos -Credential $Cred
Import-PSSession $Session_mlx01 -DisableNameChecking -CommandName set-mailbox,Set-MailboxAutoReplyConfiguration,Set-CASMailbox,Get-Mailbox,New-MoveRequest,Get-MobileDevice,Get-MailboxDatabase,Remove-MobileDevice -AllowClobber

#Выбор домен контроллера
$DCExch = (Get-ADDomainController -Discover -DomainName "domen1.akbars.ru" -service "PrimaryDC").Name
try
{
#ставим на судебное удержание, прописываем подсказку, скрываем из адресной книги
Set-Mailbox $uvolen -LitigationHoldEnabled $true -RetentionComment 'сотрудник уволен' -MailTip 'Адресат не найден в справочнике Банка. Возможно он уже не работает в данной организации' -HiddenFromAddressListsEnabled $true -DomainController $DCExch

#прописываем автоответы для писем
Set-MailboxAutoReplyConfiguration $uvolen -InternalMessage 'Адресат не найден в справочнике Банка. Возможно он уже не работает в данной организации' -ExternalMessage 'Адресат не найден в справочнике Банка. Возможно он уже не работает в данной организации' -AutoReplyState Enabled -DomainController $DCExch

#Удаляем подключенные устройства
Get-MobileDevice -Mailbox $uvolen -DomainController $DCExch | Remove-MobileDevice -Confirm:$false -DomainController $DCExch

#отключаем все протоколы клиентского доступа
Set-CASMailbox $uvolen -ActiveSyncEnabled $false -EWSEnabled $False -IMAPEnabled $False -POPEnabled $False -MAPIEnabled $False -OWAEnabled $False -OWAForDevicesEnabled $False -EWSAllowMacOutlook $False -EWSAllowOutlook $False -EWSAllowEntourage $False -MAPIBlockOutlookRpcHttp $True -ECPEnabled $False -DomainController $DCExch

WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Блокируем почтовый ящик пользователя" +$uvolen)
}
catch
{WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Почтовый ящик пользователя" +$uvolen+"НЕ заблокирован или заблокирован с ошибками !!!")
 SendEmail -strMessage "Почтовый ящик пользователя $uvolen НЕ заблокирован или заблокирован с ошибками !!!" }

#смотрим базы с начальным названием DEL, размером меньше 3TB, сортируем по размеру и выбираем меньшую. мигрируем ящик в эту базу
$TargetDB = Get-MailboxDatabase -Status | Where-Object {$_.Name -like "DEL*" -and [math]::Round($_.DatabaseSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1TB,2) -lt '3' }| Sort-Object DatabaseSize | select name -first 1
try
{Get-Mailbox $uvolen -DomainController $DCExch | New-MoveRequest -TargetDatabase $TargetDB.Name -ArchiveTargetDatabase $TargetDB.Name -DomainController $DCExch
WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Почтовый ящик пользователя " +$uvolen+" перемещен в базу "+ $TargetDB.Name)}
catch { WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Почтовый ящик пользователя" +$uvolen+"НЕ был перемещен в архивную БД !!!") 
  SendEmail -strMessage "Почтовый ящик пользователя $uvolen НЕ был перемещен в архивную БД !!!"  }

Remove-PSSession $Session_mlx01 
}
catch 
    {
    $ErrMsg = $error[0]|format-list -force | Out-String
       $ErrorDescription += $ErrMsg+ " Ошибка блокировки почты в домене $Domain."
   WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";$ErrorDescription")
   WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Для сорудника: " +$uvolen)  
    }
}


### функция блокирования УЗ в домене
function BlockingUsers ($userslist,$cred,$Domain,$tpath)
{
     $samname = $userslist.sAMAccountName
         
     # очищаем предыдущие ошибки
     $error.clear()

     try 
     {     
     # устанавливаем сложный пароль
     SetUserComplexPassword -userName $userslist.SamAccountName -Domain $Domain -Credential $cred
       
     #блокируем пользователя
     $userDescription = $description + " / " + $userslist.Description  
     $pathOU = $userslist.distinguishedName.split(",",2)[1]  

     if ($Domain -ne "domen3.akbars.ru") 
     {Set-ADUser $userslist.SamAccountName -ChangePasswordAtLogon $true -Enabled $False -Description $userDescription -Replace @{extensionAttribute15=$pathOU} -Clear "pager" -Server $Domain  -Credential $cred}
     else {Set-ADUser $userslist.SamAccountName -ChangePasswordAtLogon $true -Enabled $False -Description $userDescription -Clear "pager" -Server $Domain  -Credential $cred}   
            
     Set-ADAccountExpiration -Identity $userslist -Server $Domain -DateTime (Get-Date) -Credential $cred     
   
     # пишем в лог
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname заблокирована в домене $Domain") 
     }
     catch 
     {
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname не заблокирована в домене $Domain") 
     SendEmail -strMessage "УЗ $samname не заблокирована в домене $Domain; $error" -recipient $ADDrecipient
     }

     try 
     {    
     # перемещаем УЗ в OU для отключенных
     Move-ADObject $userslist -Server $Domain -TargetPath $tpath -Credential $cred
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname перемещена в OU $tpath") 
     }
     catch 
     {
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname НЕ перемещена в OU $tpath") 
     SendEmail -strMessage "УЗ $samname НЕ перемещена в $tpath; $error" -recipient $ADDrecipient
     }
     try
     {
     # переименовываем УЗ, добавляем в конце _
     $newname = $userslist.name + "_"                                                      
     Rename-ADObject -Identity $userslist.ObjectGUID -NewName $newname -Server $Domain -Credential $cred 
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname переименована _") 
     }
     catch
     {
      # переименовываем УЗ, добавляем в конце _
     $newname = $userslist.name + "__"  
     try { Rename-ADObject -Identity $userslist.ObjectGUID -NewName $newname -Server $Domain -Credential $cred }
     catch {
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ $samname НЕ переименована _") 
     SendEmail -strMessage "УЗ $samname НЕ переименована _ ; $error" -recipient $ADDrecipient
     }
     }     

}


# функция отправки уведомления руководителю если у сотрудника есть перемещаемый профиль
Function FindRoamingUserProfile ($user)
{
# собираем массив профилей пользователей
foreach ($share in $UsersRoamingProfiles) {$UsersRP += Get-ChildItem -Path $share -Directory}

     if (($UsersRP).where{$_.name -like $user.sAMAccountName})
     {    
          # ищем руководителя сотрудника
          if ($user.sAMAccountName -like "z-*")
          {         
          $managerABB = Get-ADUser  (Get-ADUser $user.sAMAccountName.split("-")[1] -Server $domain -Properties manager).Manager -Server "domen1.akbars.ru" -Properties mail
          $managerFIL = Get-ADUser  (Get-ADUser $user.sAMAccountName.split("-")[1] -Server $domain -Properties manager).Manager -Server "fil.domen1.akbars.ru" -Properties mail        
          }
          else
          {         
          $managerABB = Get-ADUser  $user.Manager -Server "domen1.akbars.ru" -Properties mail
          $managerFIL = Get-ADUser  $user.manager -Server "fil.domen1.akbars.ru" -Properties mail        
          }
      
          # проверяем что УЗ руководителя включена и имеется почта, отправляем уведомление на почту о удалении профиля пользователя
          if ($managerABB.Enabled -eq $true -and $managerABB.mail -ne $null) 
          { 
          SendEmailProfile -recipient  $managerABB.mail 
          WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найден руководитель в АББ для УЗ "+$user.userPrincipalName)
          }
          elseif ($managerFIL.enabled -eq $true -and $managerABB.mail -ne $null) 
          { 
          SendEmailProfile -recipient  $managerFIL.mail 
          WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найден руководитель в FIL для УЗ "+$user.userPrincipalName)
          } 
          else {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Перемещаемый профиль есть, но актуальный руководитель не найден для УЗ "+$user.userPrincipalName)}    
     }
     else {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Перемещаемый профиль или папки не найдены для УЗ "+$user.userPrincipalName)}
}



####### главная функция  ##########
function MainFunc
{ 
 # очищаем предыдущие ошибки
 $error.clear()
 
  # креды на выполнение оперции для домена 
  $ab_pwd_path = "F:\abbwin\cyph_keys\pwd_abw.txt"
  $ab_password = Get-Content $ab_pwd_path | ConvertTo-SecureString -Key (Get-Content $key_path)
  $ABBcred = New-Object System.Management.Automation.PsCredential("domen1\ad-expert",$ab_password)

  # креды на выполнение оперции для домена FIL
  $f_pwd_path = "F:\abbwin\cyph_keys\pwd_fil.txt"
  $f_password = Get-Content $f_pwd_path | ConvertTo-SecureString -Key (Get-Content $key_path)
  $FILcred = New-Object System.Management.Automation.PsCredential("fil\access",$f_password)

  # креды на выполнение оперции для домена domen3
  $mofs_pwd_path = "F:\abbwin\cyph_keys\pwd_mof.txt"
  $mofs_password = Get-Content $mofs_pwd_path | ConvertTo-SecureString -Key (Get-Content $key_path)
  $MOcred = New-Object System.Management.Automation.PsCredential("domen3\abs",$mofs_password)

  # OU с уволенными для каждого домена         
  $ABBpath = "OU=Увольнение,OU=Отключенные пользователи,OU=Пользователи и группы,OU=Головной офис,DC=domen1,DC=akbars,DC=ru"
  $FILpath = "OU=Увольнение,OU=Отключенные пользователи,DC=fil,DC=domen1,DC=akbars,DC=ru"
  $MOpath = "OU=Увольнение,OU=Отключенные пользователи,DC=domen3,DC=akbars,DC=ru"

 # определяем тип заявки
  if ($blockingreason -eq "2") {$blockingreason = "Перевод из другого подразделения"}
  elseif ($blockingreason -eq "21") {$blockingreason = "Увольнение"}
  elseif ($blockingreason -eq "23") {$blockingreason = "Отпуск по уходу за ребенком"}
  elseif ($blockingreason -eq "24") {$blockingreason = "Очередной отпуск"}
  else {SendEmail -strMessage "Тип заявки $documnumber не определен" }

  # получаем ФИО УЗ   
  $AdUserName = $surname+" "+$name+" "+$patronymic
  
  # пишем в лог
  WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";поиск в домене $Domain")
  WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";поиск по ФИО $AdUserName")

  # комментарий в AD
  $description ="$blockingreason с $dismissdate (№ $documnumber)"
   
   # ищем пользователей в домене
   $Userslist = Get-ADUser -Filter "((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*'))"  -Server $Domain -Properties displayName,Description,homeMDB,wWWHomePage,pager,extensionAttribute15 -Credential $ABBcred | where {$_.distinguishedName -notlike $ABBpath }
   $UsersListCount=@($Userslist).Count
   $TabUserslist = Get-ADUser -Filter "(((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*')) -and (SamAccountName -notlike 'z-*'))"  -Server $Domain -Properties displayName,Description,homeMDB,wWWHomePage,pager,manager,extensionAttribute15 -Credential $ABBcred | where {$_.distinguishedName -notlike $ABBpath -and $_.wWWHomePage -like $TabNumber}
   $TabUsersListCount=@($TabUserslist).Count
   
   ### выполняем тип заявки "Увольнение"
   if ($UsersListCount -ne 0 -and $blockingreason -eq "Увольнение")
   {
    ### обьявляем переменные
    
     # список доменов для поиска Z УЗ
     $ABBSrv = "domen1.akbars.ru"
     $FILSrv = "fil.domen1.akbars.ru"
     $MOSrv = "domen3.domen1.akbars.ru"     
     
     # поиск УЗ без ТН
    $NoTabUserslist = Get-ADUser -Filter "(((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*')) -and (SamAccountName -notlike 'z-*'))"  -Server $Domain -Properties displayName,Description,homeMDB,wWWHomePage,pager,extensionAttribute15,manager -Credential $ABBcred | where {$_.distinguishedName -notlike $ABBpath -and $_.wWWHomePage.count -like 0}
    $NoTabUsersListCount=@($NoTabUserslist).Count
    # поиск Z УЗ для каждого домена
    $ZuserslistABB = Get-ADUser -Filter "(((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*')) -and (SamAccountName -like 'z-*'))"  -Server $ABBSrv -Credential $ABBcred -Properties Description,pager,extensionAttribute15 | where {$_.distinguishedName -notlike $ABBpath}
    $ZuserslistCountABB = @($ZuserslistABB).Count
    $ZuserslistFIL = Get-ADUser -Filter "(((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*')) -and (SamAccountName -like 'z-*'))"   -Server $FILSrv -Credential $FILcred -Properties Description,pager,extensionAttribute15 | where {$_.distinguishedName -notlike $FILpath}
    $ZuserslistCountFIL = @($ZuserslistFIL).Count
    try {
    $ZuserslistMO = Get-ADUser -Filter "(((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*')) -and (SamAccountName -like 'z-*'))"   -Server $MOSrv -Credential $MOcred -Properties Description,pager | where {$_.distinguishedName -notlike $MOpath}
    $ZuserslistCountMO = @($ZuserslistMO).Count}
    catch {}
      
      ### если найден пользователь с табельным номером
      if ($TabUsersListCount -eq 1)
      {
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная УЗ $TabUserslist c табельным номером $TabNumber в домене $Domain") 
        # выполняем блокировку учетной записи с ТН
        BlockingUsers -userslist $TabUserslist -cred $ABBcred -Domain $Domain -tpath $ABBpath
        # проверяем есть ли перемещаемый профиль у сотрудника
        FindRoamingUserProfile -user $TabUserslist
        # записываем результат
        $documdescription = " $Domain\$($TabUsersList.SamAccountName)"
        # отправляем результат в LotusNotes
        Write-Output "Исполнено" 
        Write-Output $documdescription
       "1" | Out-File  -FilePath $LotusPath -Encoding unicode -Append
       "$documdescription заблокирована" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      elseif ( $TabUsersListCount -gt 1)
      {
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! УЗ $TabUserslist c табельным номером $TabNumber в домене $Domain") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! УЗ $TabUserslist c табельным номером $TabNumber в домене $Domain"
         # отправляем результат в LotusNotes
        Write-Output "Исполнено"
        Write-Output "Найдено несколько УЗ в домене $Domain"
        "1" | Out-File   -FilePath $LotusPath -Encoding unicode -Append
        "Найдено несколько УЗ $TabUserslist в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      else
      {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ c табельным номером $TabNumber для домена $Domain не найдена") 
       # отправляем результат в LotusNotes
         Write-Output "Исполнено"
         Write-Output "Учетная запись $AdUserName с табельным номером отсутствует в домене $Domain"
        "1" | Out-File   -FilePath $LotusPath -Encoding unicode -Append
        "Учетная запись $AdUserName с табельным номером отсутствует в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      

      ### если найден пользователь без табельного номера
      if ($NoTabUsersListCount -eq 1)
      {
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная УЗ $NoTabUsersList без ТН в домене $Domain") 
        # выполняем блокировку учетной записи без ТН
        BlockingUsers -userslist $NoTabUsersList -cred $ABBcred -Domain $Domain -tpath $ABBpath
        # проверяем есть ли перемещаемый профиль у сотрудника
        FindRoamingUserProfile -user $NoTabUserslist
        # отправляем результат в LotusNotes
        $documdescription = " $Domain\$($NoTabUserslist.SamAccountName)"
        Write-Output "Исполнено" 
        Write-Output $documdescription
       "1" | Out-File  -FilePath $LotusPath -Encoding unicode -Append
       "$documdescription заблокирована" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      elseif ($NoTabUsersListCount -gt 1)
      {
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! УЗ $NoTabUsersList без ТН в домене $Domain") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! УЗ $NoTabUserslist без ТН в домене $Domain"
         # отправляем результат в LotusNotes
        Write-Output "Исполнено"
        Write-Output "Найдено несколько УЗ в домене $Domain"
        "1" | Out-File   -FilePath $LotusPath -Encoding unicode -Append
        "Найдено несколько УЗ $NoTabUserslist в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      else
      {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";УЗ без ТН для домена $Domain не найдена")
        # отправляем результат в LotusNotes
         Write-Output "Исполнено"
         Write-Output "Учетная запись без ТН отсутствует в домене $Domain"
        "1" | Out-File   -FilePath $LotusPath -Encoding unicode -Append
        "Учетная запись без ТН отсутствует в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
       }

      
      ### если найдена Z УЗ в домене 
      if ($ZuserslistCountABB -eq 1)
      {
       $ADDrecipient = 'aba@akbars.ru'   
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная Z УЗ $ZuserslistABB  в домене $ABBSrv") 
        # выполняем блокировку Z учетной записи в домене ABB
        BlockingUsers -userslist $ZuserslistABB -cred $ABBcred -Domain $ABBSrv -tpath $ABBpath
        # проверяем есть ли перемещаемый профиль у сотрудника
        FindRoamingUserProfile -user $ZuserslistABB
      }
      elseif ($ZuserslistCountABB -gt 1)
      {
       $ADDrecipient = 'aba@akbars.ru'   
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! Z УЗ $ZuserslistABB в домене $ABBSrv") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! Z УЗ $ZuserslistABB  в домене $ABBSrv" -recipient $ADDrecipient
      }
      else
      {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + "; Z УЗ для домена $ABBSrv не найдена") }


       ### если найдена Z УЗ в домене FIL
      if ($ZuserslistCountFIL -eq 1)
      {
        $ADDrecipient = 'aba@akbars.ru'      
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная Z УЗ $ZuserslistFIL  в домене $FILSrv") 
        # выполняем блокировку Z учетной записи в домене FIL
        BlockingUsers -userslist $ZuserslistFIL -cred $FILcred -Domain $FILSrv -tpath $FILpath
        # проверяем есть ли перемещаемый профиль у сотрудника
        FindRoamingUserProfile -user $ZuserslistFIL
      }
      elseif ($ZuserslistCountFIL -gt 1)
      {
        $ADDrecipient = 'aba@akbars.ru'   
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! Z УЗ $ZuserslistFIL в домене $FILSrv") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! Z УЗ $ZuserslistFIL  в домене $FILSrv" -recipient $ADDrecipient
      }
      else
      {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + "; Z УЗ для домена $FILSrv не найдена") }


       ### если найдена Z УЗ в домене domen3
      if ($ZuserslistCountMO -eq 1)
      {
        $ADDrecipient = 'aba@akbars.ru'   
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная Z УЗ $ZuserslistMO  в домене $MOSrv") 
        # выполняем блокировку Z учетной записи в домене domen3
        BlockingUsers -userslist $ZuserslistMO -cred $MOcred -Domain $MOSrv -tpath $MOpath
        # проверяем есть ли перемещаемый профиль у сотрудника
        FindRoamingUserProfile -user $ZuserslistMO
      }
      elseif ($ZuserslistCountMO -gt 1)
      {
        $ADDrecipient = 'aba@akbars.ru'   
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! Z УЗ $ZuserslistMO в домене $MOSrv") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! Z УЗ $ZuserslistMO  в домене $MOSrv" -recipient $ADDrecipient
      }
      else
      {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + "; Z УЗ для домена $MOSrv не найдена") }



   }

   ### выполняем тип заявки "Отпуск по уходу за ребенком"
   elseif ($UsersListCount -ne 0 -and $blockingreason -eq "Отпуск по уходу за ребенком")
   { 
     # OU для декретных
     $ABBpathD = "OU=Декретный отпуск,OU=Отключенные пользователи,OU=Пользователи и группы,OU=Головной офис,DC=domen1,DC=akbars,DC=ru"
     $FILpatD = "OU=Декретный отпуск,OU=Отключенные пользователи,DC=fil,DC=domen1,DC=akbars,DC=ru"
     $MOpathD = "OU=Декретный отпуск,OU=Отключенные пользователи,DC=domen3,DC=akbars,DC=ru"

     if ($TabUsersListCount -eq 1)
     {
      # пишем в лог
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдена единственная УЗ $TabUsersList  в домене $Domain") 
      # выполняем блокировку учетной записи
      BlockingUsers -userslist $TabUsersList -cred $ABBcred -Domain $Domain -tpath $ABBpathD
      
      # отправляем результат в LotusNotes
      $documdescription = " $Domain\$($TabUsersList.SamAccountName)"
       Write-Output "Исполнено" 
       Write-Output $documdescription
      "1" | Out-File  -FilePath $LotusPath -Encoding unicode
      "$documdescription заблокирована" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
     }
     elseif ( $TabUsersListCount -gt 1)
      {
        # пишем в лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Найдено несколько!!! УЗ $TabUserslist c табельным номером $TabNumber в домене $Domain") 
        # отправляем уведомление в почту
        SendEmail -strMessage "Найдено несколько!!! УЗ $TabUserslist c табельным номером $TabNumber в домене $Domain"
        # отправляем результат в LotusNotes
        Write-Output "Исполнено"
        Write-Output "Найдено несколько УЗ в домене $Domain"
        "1" | Out-File   -FilePath $LotusPath -Encoding unicode
        "Найдено несколько УЗ $TabUsersList в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
      }
      
   }

   ### выполняем тип заявки "Перевод из другого подразделения"
   elseif ($UsersListCount -ne 0 -and $blockingreason -eq "Перевод из другого подразделения")
   {     
     # выполняем дополнительный поиск по новому табельному номеру
     $NewTabUserslist = Get-ADUser -Filter "((name -like '*$AdUserName*') -or (displayName -like '*$AdUserName*'))"  -Server $Domain -Properties displayName,Description,wWWHomePage -Credential $ABBcred| where {$_.distinguishedName -notlike $ABBpath -and $_.wWWHomePage -like $NewTabNum}
     $NewTabUsersListCount=@($NewTabUserslist).Count
     
     if ($NewTabUsersListCount -eq 1)
     {
       WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";найдена единственная УЗ $AdUserName с новым ТН $NewTabNum")
    
       # отправляем результат в LotusNotes
        $documdescription = " $Domain\$($NewTabUserslist.SamAccountName)"
        Write-Output "Исполнено"
        Write-Output "$documdescription табельный номер уже новый"
       "1" | Out-File  -FilePath $LotusPath -Encoding unicode
       "$documdescription табельный номер уже новый" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
     }

     # если новый табельный номер имеется
     elseif ($TabUsersListCount -eq 1 -and $NewTabNum -like "*0*")
     {
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";найдена единственная УЗ $AdUserName с  ТН $TabNumber")

      # меняем ТН у пользователя
      try
      {
      $userDescription = $description + " / " + $userslist.Description  
      Set-ADUser $TabUserslist.SamAccountName -Replace @{wWWHomePage = $NewTabNum;employeeID=$NewTabNum } -Description $userDescription  -Server $Domain -Credential $ABBcred
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Меняем табельный номер пользователя $TabUserslist")  
      }
      catch {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Табельный номер пользователя Не изменен! $TabUserslist") 
      SendEmail -strMessage "Табельный номер пользователя Не изменен! $TabUserslist" }
       
       # отправляем результат в LotusNotes       
      $documdescription = " $Domain\$($TabUserslist.SamAccountName)"               
      Write-Output "Исполнено"
      Write-Output "$documdescription табельный номер изменен"
      "1" | Out-File  -FilePath $LotusPath -Encoding unicode
      "$documdescription табельный номер изменен" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
     }

     # если НЕТ нового табельного номер
     elseif ($TabUsersListCount -eq 1 -and $NewTabNum -notlike "*0*" )
     {
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";найдена единственная УЗ $AdUserName с  ТН $TabNumber")
      $userDescription = $description + " / " + $userslist.Description  
      try {
      Set-ADUser $TabUserslist.SamAccountName  -Description $userDescription -Server $Domain   -Credential $ABBcred
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Прописываем Description у пользователя $TabUserslist")  
      }
      catch {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";Description у пользователя Не изменен! $TabUserslist") 
      SendEmail -strMessage "Description у пользователя Не изменен! $TabUserslist" }

       # отправляем результат в LotusNotes       
      $documdescription = " $Domain\$($TabUserslist.SamAccountName)"               
      Write-Output "Исполнено"
      Write-Output "$documdescription новый табельный номер отсутствует"
      "1" | Out-File  -FilePath $LotusPath -Encoding unicode
      "$documdescription новый табельный номер отсутствует" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
     }
     elseif ($NewTabUsersListCount -gt 1 -or $TabUsersListCount -gt 1 )
     {
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";Найдено несколько УЗ $NewTabUsersList или $TabUsersList") 
      SendEmail -strMessage "Найдено несколько УЗ $NewTabUsersList или $TabUsersList" 

       # отправляем результат в LotusNotes
      Write-Output "Исполнено"
      Write-Output "Найдено несколько УЗ в домене $Domain"
      "1" | Out-File   -FilePath $LotusPath -Encoding unicode
      "Найдено несколько УЗ $TabUsersList $NewTabUsersList в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
       }
     
   }
   
   ### выполняем тип заявки "Очередной отпуск"
   elseif ($UsersListCount -ne 0 -and $blockingreason -eq "Очередной отпуск")
   {
    if ($TabUsersListCount -eq 1)
    {
    # пишем в лог
     WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";найдена единственная УЗ $AdUserName с  ТН $TabNumber")     

      # отправляем результат в LotusNotes
      $documdescription = " $Domain\$($TabUsersList.SamAccountName)"
      Write-Output "Исполнено"
      Write-Output "$documdescription очередной отпуск"
      "1" | Out-File  -FilePath $LotusPath -Encoding unicode
      "$documdescription очередной отпуск" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
    }
    elseif ($TabUsersListCount -gt 1)
    {     
      WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";Найдено несколько УЗ $TabUsersList") 
      SendEmail -strMessage "Найдено несколько УЗ $TabUsersList"  
      # отправляем результат в LotusNotes
      Write-Output "Исполнено"
      Write-Output "Найдено несколько УЗ в домене $Domain"
      "1" | Out-File   -FilePath $LotusPath -Encoding unicode
      "Найдено несколько УЗ $TabUsersList в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
    }
    
  }

   else {WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() +";УЗ с ФИО $AdUserName в домене $Domain не найдены") 
   # отправляем результат в LotusNotes
    Write-Output "Исполнено"
    Write-Output "Учетная запись $AdUserName отсутствует в домене $Domain"
    "1" | Out-File   -FilePath $LotusPath -Encoding unicode -Append
    "Учетная запись $AdUserName отсутствует в домене $Domain" | Out-File  -Append -FilePath $LotusPath -Encoding unicode
   }

}


# тело скрипта ---
try
{  
    # создаем каталог Logs если его нет
    if(!(Test-Path -Path ($ParentFolder +  "\Logs") )){New-Item -ItemType directory -Path ($ParentFolder +  "\Logs")  | Out-Null}
    # создаем каталог Errors если его нет
    if(!(Test-Path -Path ($errorFilesPath) )){New-Item -ItemType directory -Path ($errorFilesPath)  | Out-Null}

    # создаем лог файл если отсутствует
    if(!(Test-Path -Path $LogFilePath))
      {
        Add-Content -Value "Date;Message" -Path $LogFilePath -Encoding UTF8 
      }
    else # удаляем записи из лога старше 90 дней
     {
       $LogLines = Import-Csv -Path $LogFilePath -Header Date, Message -Delimiter ";"  | where {$_.Date -ne "Date"} |
       Where-Object {[datetime]::ParseExact($_.Date,'dd.MM.yyyy HH:mm:ss',$null) -gt (get-date).addDays(-90)} 
       Set-Content -Value "Date;Message" -Path $LogFilePath -Encoding UTF8 
       foreach ($Lines in $LogLines) 
       {Add-Content -Value  ($Lines.Date + ";" + $Lines.Message) -Path $LogFilePath -Encoding UTF8 }
     }

    Add-Content -Value ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";--- скрипт запущен под пользователем " + [Environment]::UserName + " ---") -Path $LogFilePath -Encoding UTF8

    
}
catch
{
 $ErrMsg = $error[0]|format-list -force | Out-String
 $tmperrMsg = "Ошибка при создании, редактировании логфайлов скриптом $ScriptPath"
 $tmperrMsg += "`nКомпьютер " + "$env:computername.$env:userdnsdomain" + "`nКаталог " + $LogFilePath
 $tmperrMsg += "`nCкрипт запущен под пользователем " + [Environment]::UserName
 $tmperrMsg += "`nТекст ошибки:`n" + $ErrMsg 
 # очищаем предыдущие ошибки
 $error.clear()
 # отправляем smtp сообщение
 SendEmail $tmperrMsg
}

# проверяем значения обязательных параметров
foreach ($ParamName in $ParamList) 
{ 
    # получаем значение параметра
    $ParamValue = Get-Variable -Name $ParamName -ValueOnly 

    # если один из обязательных параметров пуст:
    if ($ParamValue -eq $null -or $ParamValue -eq "") 
    {
        $strMsg = "Ошибка: один из следующих входных параметров пуст: "
        $ParamList | foreach { $strMsg += [string]::Format("'{0}'='{1}' ",$_,(Get-Variable -Name $_ -ValueOnly)) }
        # пишем в основной лог
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";$strMsg")       
        $strMsg += "`nСкрипт $ScriptPath"
        $strMsg += "`nКомпьютер " + "$env:computername.$env:userdnsdomain" + "`nКаталог " + $LogFilePath 
        # отправляем smtp сообщение
        SendEmail $strMsg
        WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";--- один из входных параметров пуст, скрипт остановлен ---")
        exit
    }
}


WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";запускаем скрипт для заявки $documnumber")

# запускаем основной процесс
MainFunc

WriteToLogFile ((Get-Date -Format 'dd.MM.yyyy HH:mm:ss').ToString() + ";--- скрипт остановлен ---")
# конец скрипта ----------
