cls

$dateLOG = Get-Date -Format "dd_MM_yyyy"
# путь к логам действий
$path = "\\Ogurtsov\Users_Changes_$dateLOG.csv"
# путь к логам ошибок
$LogFilePath  = "\\domen1\Ogurtsov\Users_Changes_Errors_$dateLOG.txt"

# домены
$domainABB = "domen1.akbars.ru"
$domainFIL = "domen2.akbars.ru"

# SQL сервер
$server = "sql123.domen1.akbars.ru\db4,1454"
$database = "DWDATA_123"

# креды для подключения к SQL
$CredSQL  = New-Object -typename System.Management.Automation.PSCredential -argumentlist $usernameSQL, $passwordSQL

# креды для внесения изменения ABB
$ABBcred     = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ABBusername, $ABBpassword

# креды на выполнение оперции для домена FIL
$FILcred     = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $FILusername, $FILpassword

# загружаем массив пользователей для проверки
$Users = Invoke-Sqlcmd  -ServerInstance $server -Database $database -Credential $CredSQL -Query "select email, city_phone, phone, fio, boss,build, department, dep, dep1, podr, tab_numb, title, phcode,[group], dokko from LOTUS.PHONE_BOOK" 
#$Users = $Users | select -First 100

$Global:Tables = @()

### записываем лог ошибок в файл
function WriteToLogFile ([string]$strMsg)
{
try { Add-Content -Path $LogFilePath -Value $strMsg -Encoding UTF8 }
catch{}
}


# функция проверки атрибутов пользователя
function CheckUserChanges ($UserChanges,$Cred,$domain)
{$report = @()
      # сверяем имя пользователя
      if ($user.fio.Split(" ")[1] -ne $UserChanges.GivenName -and $user.fio) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="GivenName";"Before"=$UserChanges.GivenName;"After"=$user.fio.Split(" ")[1]}          
           Set-ADUser $UserChanges.ObjectGUID -GivenName $user.fio.Split(" ")[1] -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить Имя для "+$UserChanges.UserPrincipalName+" на "+$user.fio.Split(" ")[1]+" не удалось")}
      }

      # сверяем фамилию пользователя
      if ($user.fio.Split(" ")[0] -ne $UserChanges.Surname -and $user.fio) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Surname";"Before"=$UserChanges.Surname;"After"=$user.fio.Split(" ")[0]}           
           Set-ADUser $UserChanges.ObjectGUID -Surname $user.fio.Split(" ")[0] -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить Фамилию для "+$UserChanges.UserPrincipalName+" на "+$user.fio.Split(" ")[0]+" не удалось")}      
      }

      # сверяем отображаемое ФИО пользователя
      if ($user.fio -ne $UserChanges.DisplayName -and $user.fio) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="DisplayName";"Before"=$UserChanges.DisplayName;"After"=$user.fio}           
           Set-ADUser $UserChanges.ObjectGUID -DisplayName $user.fio -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить отображаемое ФИО для "+$UserChanges.UserPrincipalName+" на "+$user.fio+" не удалось")}        
      }

      # сверяем ФИО пользователя
      if ($user.fio -ne $UserChanges.Name -and $user.fio) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Name";"Before"=$UserChanges.Name;"After"=$user.fio}         
           Rename-ADObject $UserChanges.ObjectGUID -NewName $user.fio -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить ФИО для "+$UserChanges.UserPrincipalName+" на "+$user.fio+" не удалось")}      
      }

      # сверяем табельный номер пользователя
      if (($user.tab_numb -ne $UserChanges.wWWHomePage -or $user.tab_numb -ne $UserChanges.employeeID) -and $user.tab_numb)  
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="EmployeeID";"Before"=$UserChanges.EmployeeID;"After"=$user.tab_numb}
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="wWWHomePage";"Before"=$UserChanges.wWWHomePage;"After"=$user.tab_numb}
           Set-ADUser $UserChanges.ObjectGUID -EmployeeID $user.tab_num  -HomePage $user.tab_num -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить табельный для "+$UserChanges.UserPrincipalName+" на "+$user.tab_numb+" не удалось")}        
      }
      elseif (!$user.tab_numb) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="EmployeeID";"Before"=$UserChanges.EmployeeID;"After"="нет табельного"}
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="wWWHomePage";"Before"=$UserChanges.wWWHomePage;"After"="нет табельного"}         
           Set-ADUser $UserChanges.ObjectGUID -Clear EmployeeID,wWWHomePage -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Удалить табельный для "+$UserChanges.UserPrincipalName+" не удалось")}        
      }

      # сверяем должность сотрудника
      if ($user.title -ne $UserChanges.title -and $user.title) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Title";"Before"=$UserChanges.Title;"After"=$user.title}           
           Set-ADUser $UserChanges -Title $user.title -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить должность для "+$UserChanges.UserPrincipalName+" на "+$user.title+" не удалось")}        
      }

      # сверяем группу сотрудника
      if ($user.group -ne $UserChanges.extensionAttribute5 -and $user.group) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Otdel";"Before"=$UserChanges.extensionAttribute5;"After"=$user.group}          
           Set-ADUser $UserChanges.ObjectGUID -Replace @{extensionAttribute5 = $user.group} -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить отдел для "+$UserChanges.UserPrincipalName+" на "+$user.group+" не удалось")}       
      }
      elseif (!$user.group -and $UserChanges.extensionAttribute5)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Otdel";"Before"=$UserChanges.extensionAttribute5;"After"="нет группы"}         
           Set-ADUser $UserChanges.ObjectGUID -Clear extensionAttribute5 -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Удалить отдел для "+$UserChanges.UserPrincipalName+" не удалось")}  
      }

      # сверяем отдел сотрудника
      if ($user.dep1 -ne $UserChanges.extensionAttribute4 -and $user.dep1) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Otdel";"Before"=$UserChanges.extensionAttribute4;"After"=$user.dep1}          
           Set-ADUser $UserChanges.ObjectGUID -Replace @{extensionAttribute4 = $user.dep1} -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить отдел для "+$UserChanges.UserPrincipalName+" на "+$user.dep1+" не удалось")}       
      }
      elseif (!$user.dep1 -and $UserChanges.extensionAttribute4)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Otdel";"Before"=$UserChanges.extensionAttribute4;"After"="нет отдела"}         
           Set-ADUser $UserChanges.ObjectGUID -Clear extensionAttribute4 -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Удалить отдел для "+$UserChanges.UserPrincipalName+" не удалось")}  
      }

      # сверяем управление сотрудника
      if ($user.dep -ne $UserChanges.extensionAttribute3 -and $User.dep) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Upravlenie";"Before"=$UserChanges.extensionAttribute3;"After"=$user.dep}           
           Set-ADUser $UserChanges.ObjectGUID -Replace @{extensionAttribute3 = $user.dep} -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить управление для "+$UserChanges.UserPrincipalName+" на "+$user.dep+" не удалось")}       
      }
      elseif (!$user.dep -and $UserChanges.extensionAttribute3)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Upravlenie";"Before"=$UserChanges.extensionAttribute3;"After"="нет управления"}           
           Set-ADUser $UserChanges.ObjectGUID -Clear extensionAttribute3 -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Удалить управление для "+$UserChanges.UserPrincipalName+" не удалось")}  
      }


      # сверяем департАмент сотрудника
      if ($user.department -ne $UserChanges.extensionAttribute2 -and $user.department) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="DepartAment";"Before"=$UserChanges.extensionAttribute2;"After"=$user.department}             
           Set-ADUser $UserChanges.ObjectGUID -Replace @{extensionAttribute2 = $user.department} -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить департАмент для "+$UserChanges.UserPrincipalName+" на "+$user.department+" не удалось")}      
      }
      elseif (!$user.department -and $UserChanges.extensionAttribute2)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="DepartAment";"Before"=$UserChanges.extensionAttribute2;"After"="нет департамента"}            
           Set-ADUser $UserChanges.ObjectGUID -Clear extensionAttribute2 -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Удалить департАмент для "+$UserChanges.UserPrincipalName+" не удалось")}  
      }


      # сверяем департмент сотрудника
      if ($user.department -notmatch $UserChanges.Department -and $user.department) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Department";"Before"=$UserChanges.Department;"After"=$user.department}           
           Set-ADUser $UserChanges.ObjectGUID -Department $user.department.subString(0,[System.Math]::Min(64,$user.department.Length)) -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить Department для "+$UserChanges.UserPrincipalName+" на "+$user.department+" не удалось")}      
      }
      elseif (!$user.department -and $user.dep -notmatch $UserChanges.Department -and $user.dep)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Department";"Before"=$UserChanges.Department;"After"=$user.department}   
           Set-ADUser $UserChanges.ObjectGUID -Department $user.dep.subString(0,[System.Math]::Min(64,$user.dep.Length)) -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить Department для "+$UserChanges.UserPrincipalName+" на "+$user.dep+" не удалось")}  
      }
      elseif (!$user.department -and $user.dep1 -notmatch $UserChanges.Department -and !$user.dep -and $user.dep1)
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Department";"Before"=$UserChanges.Department;"After"=$user.department}   
           Set-ADUser $UserChanges.ObjectGUID -Department $user.dep1.subString(0,[System.Math]::Min(64,$user.dep1.Length)) -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить Department для "+$UserChanges.UserPrincipalName+" на "+$user.dep1+" не удалось")}  
      }


      # сверяем менеджера сотрудника
      try {
      $ABBboss = (Get-ADUser $UserChanges.Manager -Server "domen1.akbars.ru" -Properties mail).mail
      $FILboss = (Get-ADUser $UserChanges.Manager -Server "domen2.akbars.ru" -Properties mail).mail}
      catch{}
      if (($user.boss -ne $ABBboss -or $user.boss -ne $FILboss) -and $user.boss -and $user.boss -notlike "*CN=*") 
      {
           try 
           {
                if (@($ABBboss).Count -eq 1)
                {[string]$BossEmail = $user.boss
                $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Manager";"Before"=$ABBboss;"After"=$user.boss}           
                Set-ADUser $UserChanges.ObjectGUID -Manager (Get-ADUser -Filter {mail -eq $BossEmail} -Properties mail).DistinguishedName -Server $domain -Credential $cred
                }
                elseif (@($FILboss).Count -eq 1)
                {[string]$BossEmail = $user.boss
                $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="Manager";"Before"=$FILboss;"After"=$user.boss}           
                Set-ADUser $UserChanges.ObjectGUID -Manager (Get-ADUser -Filter {mail -eq $BossEmail} -Properties mail).DistinguishedName -Server $domain -Credential $cred
                }
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить менеджера для "+$UserChanges.UserPrincipalName+" на "+$user.boss+" не удалось")}      
      }    
      

      # сверяем адрес места работы пользователя
      if ($user.build -ne $UserChanges.StreetAddress) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="StreetAddress";"Before"=$UserChanges.StreetAddress;"After"=$user.build}             
           Set-ADUser $UserChanges.ObjectGUID -StreetAddress $user.build -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить стрит адрес для "+$UserChanges.UserPrincipalName+" на "+$user.build+" не удалось")}       
      }


      # сверяем внутренний номер телефона
      if ($user.phone -ne $UserChanges.ipPhone -and $user.phone) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="ipPhone";"Before"=$UserChanges.ipPhone;"After"=$user.phone}          
           Set-ADUser $UserChanges.ObjectGUID -Replace @{ipPhone = $user.phone} -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить внутренний номер телефона для "+$UserChanges.UserPrincipalName+" на "+$user.phone+" не удалось")}                
      }


      # сверяем городской номер телефона
      if ((-join ($user.phcode+$user.city_phone).Split("-")) -notmatch $UserChanges.telephoneNumber -and (-join ($user.phcode+$user.city_phone).Split("-")) ) 
      {
           try 
           {
           $report += [PSCustomObject] @{"Date"=$date;"Name"=$UserChanges.Name;"userPrincipalName"=$UserChanges.UserPrincipalName;"Parametr"="telephoneNumber";"Before"=$UserChanges.OfficePhone;"After"=(-join ($user.phcode+$user.city_phone).Split("-"))}            
           Set-ADUser $UserChanges.ObjectGUID -OfficePhone (-join ($user.phcode+$user.city_phone).Split("-")) -Server $domain -Credential $cred
           }
           catch {WriteToLogFile (($date).ToString() + ";Изменить внутренний номер телефона для "+$UserChanges.UserPrincipalName+" на "+(-join ($user.phcode+$user.city_phone).Split("-"))+" не удалось")}      
      }

$Global:Tables += $report
}




# начинаем выполнять действия

foreach ($user in $users)
{
$double = @()
$UserEmail = $null
$userABB =$null
$userFIL =$null
# текущая дата и время
$date = Get-Date -Format "dd.MM.yyyy HH.mm"
# ищем пользователей по Email
[string]$UserEmail = $user.email
try {
$userABB = Get-ADUser -Filter {(Enabled -eq $true) -and (mail -eq $UserEmail) -and (SamAccountName -notlike "*-*") -and (description -notlike "*1/40*")} -Properties mail,description,DisplayName,title,department,streetAddress,ipPhone,OfficePhone,wWWHomePage,HomePage,employeeID,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,manager  -Server $domainABB
$userFIL = Get-ADUser -Filter {(Enabled -eq $true) -and (mail -eq $UserEmail) -and (SamAccountName -notlike "*-*") -and (description -notlike "*1/40*")} -Properties mail,description,DisplayName,title,department,streetAddress,ipPhone,OfficePhone,wWWHomePage,HomePage,employeeID,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,manager  -Server $domainFIL
}
catch {}
     # если пользователь из ABB
     if ($userABB -and @($userABB).Count -eq 1) { CheckUserChanges -UserChanges $userABB -domain $domainABB -Cred $ABBcred}
     # если пользователь из FIL
     elseif ($userFIL -and @($userABB).Count -eq 1) { CheckUserChanges -UserChanges $userFIL -domain $domainFIL -Cred $FILcred}
     elseif (@($userABB).Count -gt 1) {$double +=$userABB}
     elseif (@($userFIL).Count -gt 1) {$double +=$userFIL}
    
}



$Global:Tables | Export-Csv $path -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Append

