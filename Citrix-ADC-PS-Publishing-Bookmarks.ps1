$Credential = Get-Credential -Message "Enter your username and password to log in to Citrix Netscaler!"

Import-Module ".\PS-NITRO-master\NitroConfigurationFunctions\NITROConfigurationFunctions.psm1" -Force

#Масси DNS имен Netscaler для публикации
[String[]]$NSNameArr = "ns-01.test.com", "ns-02.test.com"

# Имя сервера на NS для проверки доступности по RDP
$NSServerRDPcheck = "Nitro_server"

# Имя севиса на NS для проверки доступности по RDP
$NSServiceRDPcheck = "LB_SVC_rdp_nitro"

#Файл со списком серверов для создания групп и публикации на rdplocal
[String[]]$computersArr = Get-Content -Path ".\servers.txt"

#OU в которой будет создана группа AD
$OUPath = "OU=TestGroups,DC=Test,DC=com"

#Шаблон для имени группы в AD, подставляется перед именем сервера
[String]$groupNameAhead = "Test-RDP-"

#Шаблон для описания группы в AD
$DescriptionAhead = "rdp bookmark | "

$authzPolicyNameAhead = "Authz_pol_"

$DateStr = (Get-Date).ToString("dd.MM.yyyy")

#Путь для файла логов
$logFilePath = ".\NSNitroLogs\" + $DateStr + "_logNSNitro.log"

#Получаем имя DC с ролью PDCEmulator для многодоменной инфраструктуры, запрос на создание группы AD будем отправлять на него,
#если введен домен, то запрос на получение PDC из этого домена

if ($Credential.UserName -match "(\w+)\\") {
    $DC = Get-ADDomain $Matches[1] -Credential $Credential | Select-Object -ExpandProperty PDCEmulator
}
else {
    $DC = Get-ADDomain -Credential $Credential | Select-Object -ExpandProperty PDCEmulator
}

#Функция записи лога в CSV файл
function WriteLogs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $logString,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $logFilePath
    )   
    $logString | Out-File $logFilePath -Encoding utf8 -Append
}

#Функция создания сессии на NS
function New-NSSession {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $NSName,
        [Parameter(Mandatory = $true)]
        [pscredential]
        $Credential
    )
    try {
        Set-NSMgmtProtocol -Protocol https
        $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
        $Session = Connect-NSAppliance -NSName $NSName -NSUserName $Credential.UserName -NSPassword $password
        return $Session
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Invalid username or password") {
            Write-Error -Message "Invalid username or password when logging in to $NSName" 
            return $null
        }
        if ($Error.ErrorDetails.Message -match "The remote name could not be resolved") {
            Write-Error -Message "Invalid Citrix Netscaler name or no network access $NSName" 
            return $null
        }
        else {
            Write-Error -Message "Error connecting to Citrix Netscaler $NSName"
            return $null
        }
    }
}

#Функция проверки является ли текущий NS Primary в HA
function Get-NSHAPrimary {
    param (
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )
    
    try {
        $NSName = $NSSession.Endpoint
        $NodeName = $NSSession.Endpoint
        if ($NodeName -match '(?<Name>.+)\.(?<domain>\w+\..+)') {
            $NodeName = $Matches.Name
        }
        $Responce = Invoke-NSNitroRestApi -OperationMethod GET -ResourceType hanode -NSSession $NSSession
        foreach ($Node in $Responce.hanode) {
            if (($Node.name -eq $NodeName) -and ($Node.state -eq "Primary")) {
                Write-Verbose "$NSName This server status is Primary"
                return $true
            }
        }
        return $false 
        
    }
    catch {
        Write-Error -Message "Request error HA Primary status Citrix Netscaler $NSName"
        return $null
    }
}

#Функция создания группы в AD, в параметрах принимает имя группы AD, имя сервера, OU группы
function New-GroupAD {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupNameAD,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $computerName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $destinationionPath,
        [Parameter(Mandatory = $true)]
        [pscredential]
        $Credential,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $DCName
    )

    try {
        $DescriptionGroup = $DescriptionAhead + $computerName
        New-ADGroup -Server $DCName $groupNameAD -path $destinationionPath -GroupScope Global -Description $DescriptionGroup -Credential $Credential -PassThru | Out-Null
        Add-ADGroupMember -Server $DCName -Identity ALL-GSG-CTX-RDPLOCAL-Access -Members $groupNameAD -Credential $Credential | Out-Null
        Write-Verbose "$groupNameAD The group was created in AD"
        
    }
    catch {
        $log = $groupNameAD + ";Error creating a group in AD"
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }
}


#Функция добаления группы на netscaler, в параметрах принимает имя группы и сессию c NS
function Add-NSgroup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )
    try {
        $NSName = $NSSession.Endpoint
        Invoke-NSNitroRestApi -ResourceType aaagroup -NSSession $NSSession -OperationMethod POST -Payload @{
            groupname = $groupName
        }
        Write-Verbose "$groupName The AD group was created on citrix netscaler $NSName"
    }
    catch {
        $log = $groupName + ";Error creating a group on Citrix Netscaler;" + $NSName
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }
}


#Функция добалвения Bookmark-а на netscaler, в параметрах принимает имя сервера и сессию с NS
function Add-NSBookmark {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $bookmarkName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    $url = "rdp://" + ([System.Net.Dns]::GetHostEntry($bookmarkName)).HostName
    try {
        $NSName = $NSSession.Endpoint
        Invoke-NSNitroRestApi -ResourceType vpnurl -NSSession $NSSession -OperationMethod POST -Payload @{
            urlname          = $bookmarkName
            linkname         = $bookmarkName
            actualurl        = $url
            clientlessaccess = "ON"
        }
        Write-Verbose "$bookmarkName The bookmark was created on citrix netscaler $NSName"
    }
    catch {
        $log = $bookmarkName + ";Error creating a bookmark on Citrix Netscaler;" + $NSName
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }
}


#Функция добавление Authorization Policy на netscaler, в параметрах принимает имя политики и сессию с NS
function Add-NSAuthzPolicy {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $authzPolicyName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $computerName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $computerRule = '"' + $computerName + '"'
        Invoke-NSNitroRestApi -ResourceType authorizationpolicy -NSSession $NSSession -OperationMethod POST -Payload @{
            name   = $authzPolicyName
            rule   = "HTTP.REQ.URL.CONTAINS($computerRule)"
            action = "ALLOW"
        } 
        Write-Verbose "$AuthzPolName The Authorization policy was created on citrix netscaler $NSName"

    }
    catch {
        $log = $authzPolName + ";Error creating authorization policy in Citrix Netscaler;" + $NSName
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }
}


#Функция биндинга Bookmark-а к группе, в параметрах принимает имя сервера, имя группы и сессию с NS
function Add-NSBindingBookmark {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $bookmarkName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        Invoke-NSNitroRestApi -ResourceType aaagroup_vpnurl_binding -NSSession $NSSession -OperationMethod PUT -Payload @{
            groupname = $groupName
            urlname   = $bookmarkName
        }  
        Write-Verbose "$bookmarkName The bookmark has been binding to a group $groupName Citrix Netscaler $NSName"
    }
    catch {
        $log = $bookmarkName + ";Error binding the Bookmark to the group;" + $groupName + ";Citrix Netscaler;" + $NSName
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }

}


#Функция биндинга Authorization Policy к группе, в параметрах принимает имя политики, имя группы и сессию с NS
function Add-NSBindingAuthzPolicy {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $authzPolicyName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        Invoke-NSNitroRestApi -ResourceType aaagroup_authorizationpolicy_binding -NSSession $NSSession -OperationMethod PUT -Payload @{
            groupname = $groupName
            policy    = $authzPolicyName
            priority  = "100"
        } 
        Write-Verbose "$authzPolicyName The authorization policy has been successfully binding to the group $groupName $NSName"
    }
    catch {
        $log = $authzPolicyName + ";Error binding the Authorization Policy to the group;" + $groupName + ";Citrix Netscaler;" + $NSName
        Write-Error -Message $log.Replace(";", " ")
        WriteLogs -logString $log -logFilePath $logFilePath
    }

}


#Функция проверки наличия группы в AD, в параметрах принимает имя группы AD 
function Get-GroupADExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupNameAD,
        [Parameter(Mandatory = $true)]
        [pscredential]
        $Credential,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $DCName
    )

    try {
        $ADGroup = Get-ADGroup -Server $DCName -Identity $groupNameAD -Credential $Credential
        Write-Verbose "$ADGroup.Name This group was successfully found in AD"
        return $true
    }
    catch {
        if ($Error.exception.Message -match "Cannot find an object with identity") {
            Write-Verbose -Message "$groupNameAD This group was not found in AD"
            return $false
        }
        else {
            Write-Error -Message "$groupNameAD group request error from AD"
            return $null
        }
    }

}


#Функция проверки на наличие группы на netscaler, в параметрах принимает имя группы AD и сессию с NS
function Get-NSGroupExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $NSGroup = Invoke-NSNitroRestApi -ResourceType aaagroup -NSSession $NSSession -OperationMethod GET -ResourceName $groupName
        Write-Verbose "$groupName This group already exists on Citrix Netscaler $NSName"
        return $true
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Group does not exist") {
            Write-Verbose -Message "$groupName This group was not found on Citrix Netscaler $NSName"
            return $false
        }
        else {
            Write-Error -Message "$groupName error when requesting a group NS $NSName"
            return $null
        }
    }

}

#Функция проверки на наличие Bookmark на netscaler, в параметрах принимает имя сервера и сессию с NS
function Get-NSBookmarkExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $bookmarkName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $NSGroup = Invoke-NSNitroRestApi -ResourceType vpnurl -NSSession $NSSession -OperationMethod GET -ResourceName $bookmarkName
        Write-Verbose "$bookmarkName This bookmark is already in Citrix Netscaler $NSName"
        return $true
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Action does not exist") {
            Write-Verbose "$bookmarkName This bookmark was not found on Citrix Netscaler $NSName"
            return $false
        }
        else {
            Write-Error -Message "$bookmarkName bookmark request error with NS $NSName"
            return $null
        }
    }

}


#Функция проверки на биндинг Bookmark к группе на netscaler, в параметрах принимает имя сервера, имя группы и сессию с NS
function Get-NSBookmarkBinding {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $bookmarkName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $NSGroupBinding = Invoke-NSNitroRestApi -ResourceType aaagroup_vpnurl_binding -NSSession $NSSession -OperationMethod GET -ResourceName $groupName 
        $ReturnBookmarkBind = $NSGroupBinding.aaagroup_vpnurl_binding.urlname -notcontains $bookmarkName
        if ($true -eq $ReturnBookmarkBind) {
            Write-Verbose "$bookmarkName This bookmark is not linked to this group $groupName $NSName"
        }
        else {
            Write-Verbose "$bookmarkName This bookmark is already binding to this group $groupName $NSName"
        }
        return $ReturnBookmarkBind
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Group does not exist") {
            Write-Error -Message "$groupName error when requesting a group to check the bookmark binding $bookmarkName $NSName"
            return $null
        }
        else {
            Write-Error -Message "Error function Get-NSBookmarkBinding group $groupName Bookmark $bookmarkName $NSName"
            return $null
        }
    }
}


#Функция проверка на наличие Authorization Policy на NS, в параметрах принимает имя политики и сессию с NS
function Get-NSAuthzPolicyExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $authzPolicyName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $AuthzPolicy = $null
        $AuthzPolicy = Invoke-NSNitroRestApi -ResourceType authorizationpolicy -NSSession $NSSession -OperationMethod GET -ResourceName $authzPolicyName
        if ($null -ne $AuthzPolicy) {
            Write-Verbose "$authzPolicyName This authorization policy already exists in Citrix netscaler $NSName"
            return $true
        }
    }
    catch {
        if ($Error.ErrorDetails.Message -match "No such policy exists") {
            Write-Verbose -Message "$authzPolicyName This authorization policy was not found on Citrix Netscaler $NSName"
            return $false
        }
        else {
            Write-Error -Message "$authzPolicyName Error function Get-NSAuthzPolicyExists authorization policy request $NSName"
            return $null
        }
    }

}

#Функция проверка на биндинг Authorization Policy к группе на NS, в параметрах принимает имя политики, имя группы и сессию с NS
function Get-NSAuthzPolicyBinding {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $authzPolicyName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $groupName,
        [Parameter(Mandatory = $true)]
        [psobject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $GetAuthzPol = Invoke-NSNitroRestApi -ResourceType aaagroup_authorizationpolicy_binding -NSSession $NSSession -OperationMethod GET -ResourceName $groupName
        $ReturnAuthzPolicy = $GetAuthzPol.aaagroup_authorizationpolicy_binding.policy -notcontains $authzPolicyName
        if ($true -eq $ReturnAuthzPolicy) {
            Write-Verbose -Message "$authzPolicyName This authorization policy is not linked to this group $groupName $NSName"
        }
        else {
            Write-Verbose "$authzPolicyName This authorization policy is already linked to this group $groupName $NSName"
        }
        return $ReturnAuthzPolicy
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Group does not exist") {
            Write-Error -Message "$groupName error when requesting a group to check the Authorization Policy binding $authzPolicyName $NSName"
            return $null
        }
        else {
            Write-Error -Message "Error function Get-NSAuthzPolicyBinding $groupName Authorization Policy $authzPolicyName $NSName"
            return $null
        }
    }

}

#Функция проверки есть ли сервер на NS
function Get-NSServerExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $NSServerName,
        [Parameter(Mandatory = $true)]
        [PSObject]
        $NSSession
    )
    
    try {
        $responseServer = Get-NSServer -ServerName $NSServerName -NSSession $NSSession
        if ($responseServer.Name -eq $NSServerName) {
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        if ($Error.ErrorDetails.Message -match "No such resource") {
            return $false
        }
        else {
            return $null
        }

    }
}

#Функция проверки есть ли сервис на NS
function Get-NSServiceExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $NSServiceName,
        [Parameter(Mandatory = $true)]
        [PSObject]
        $NSSession
    )
    try {
        $responseService = Get-NSService -Name $NSServiceName -NSSession $NSSession -Verbose:$false
        if ($responseService.Name -eq $NSServiceName) {
            return $true
        }
    }
    catch {
        if ($Error.ErrorDetails.Message -match "No such resource") {
            return $false
        }
        else {
            return $null
        }

    }
}

#Функция проверки доступа с NS до конечного компьютера по порту 3389
function Get-NSRDPAccessCheck {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $ipAddress,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $computerName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $NSServer,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $NSService,
        [Parameter(Mandatory = $true)]
        [PSObject]
        $NSSession
    )     

    $NSName = $NSSession.Endpoint
    if (!(Get-NSServerExists -NSServerName $NSServer -NSSession $NSSession)) {
        Write-Warning -Message "$NSServer This server for the test was not found on NS $NSName"
        try {

            Add-NSServer -Name $NSServer -IPAddress $ipAddress -NSSession $NSSession -Verbose:$false   
        }
        catch {
            if ($Error.ErrorDetails.Message -match "Server already exists") {
                Write-Error -Message "$NSServer Error when creating test RDP server, $ipAddress this IP address is used by another server $NSName"
            }
            else {
                Write-Error -Message "$NSServer Error when creating an access testing server on port 3389 $NSName"
            }
        }
        if (Get-NSServerExists -NSServerName $NSServer -NSSession $NSSession) {
            Write-Verbose "$NSServer The test server was created on NS $NSName"
        }
    }
    if (!(Get-NSServiceExists -NSServiceName $NSService -NSSession $NSSession)) {
        Write-Warning -Message "$NSService This service for the test was not found on NS $NSName"
        try {

            Add-NSService -Name $NSService -ServerName $NSServer -Protocol RDP -Port 3389 -NSSession $NSSession -Verbose:$false
        }
        catch {
            if ($Error.ErrorDetails.Message -match "No such server") {
                Write-Error -Message "$NSService Error when creating test RDP service, server $NSServer was not found for this service $NSName"
            }
            else {
                Write-Error -Message "$NSService Error when creating an access testing service on port 3389! $NSName"
            }
        }
        if (Get-NSServiceExists -NSServiceName $NSService -NSSession $NSSession) {
            Write-Verbose "$NSService The test service was created on NS $NSName"
        }
    }

    try {
        Update-NSServer -Name $NSServer -IPAddress $ipAddress -NSSession $NSSession -Verbose:$false
        Start-Sleep -Seconds 1
        $ResponceNSServer = Get-NSServer -ServerName $NSServer -NSSession $NSSession
    }
    catch {
        if ($Error.ErrorDetails.Message -match "Resource already exists") {
            $log = $computerName + ";" + $ipAddress + ";This server already exists on NS checking port 3389 failed!" + $NSName
            Write-Warning -Message $log.Replace(";", " ")
            WriteLogs -logString $log -logFilePath $logFilePath
        }
        else {
            $log = $computerName + ";" + $ipAddress + ";The availability check failed with an error!;" + $NSName
            Write-Error -Message $log.Replace(";", " ") 
            WriteLogs -logString $log -logFilePath $logFilePath
        }
    }
        
    if ($ResponceNSServer.ipaddress -ne $ipAddress) {
        Write-Error -Message "Failed to change the IP Address for the verification server $NSServer"
        return
    }
    $ResponceNSService = Get-NSService -Name $NSService -NSSession $NSSession -Verbose:$false
    if ($ResponceNSService.svrstate -eq "UP") {
        Write-Verbose "$computerName Available on port 3389 with NS $NSName"
    }
    else {
        Write-Verbose -Message "$computerName No access on port 3389 with NS $NSName"
    }
}

function Get-NSServer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $ServerName,
        [Parameter(Mandatory = $true)]
        [PSObject]
        $NSSession
    )

    try {
        $NSName = $NSSession.Endpoint
        $ResponceNSServer = Invoke-NSNitroRestApi -ResourceType server -NSSession $NSSession -OperationMethod GET -ResourceName $ServerName
        If ($ResponceNSServer.PSObject.Properties['server']) {
            return $ResponceNSServer.server
        }
        else {
            return $null
        }
    }
    catch {
        return $null
    }
}

#Функция проверки на наличие записи DNS для сервера, в параметрах принимает имя сервера
function Get-DNSRecordExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $computerName
    )

    try {
        $computerFQDN = ([System.Net.Dns]::GetHostEntry($computerName))
        [String]$computerFQDNString = $ComputerFQDN.HostName
        Write-Verbose "$computerFQDNString DNS record found"
        return $true
    }
    catch {
        if ($Error.exception.Message -match "No such host is known") {
            Write-Warning -Message "$computerName DNS record not found"
            return $false
        }
        else {
            Write-Error -Message "$computerName DNS record request error"
            return $null
        }
    }
}

#Функция получения IP адреса, в параметрах принимает DNS имя компьютера
function Get-ComputerIPv4Addres {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()][String]
        $computerName
    )
    try {
        (([System.Net.Dns]::GetHostAddresses($computerName)).IPAddressToString) -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' | Out-Null
        [String]$IPv4Address = $Matches.Values
        if (!([string]::IsNullOrEmpty($IPv4Address))) {
            Write-Verbose "$computerName IPv4 address $IPv4Address" 
            return $IPv4Address
        }
        else {
            return $null
        }
    }
    catch {
        Write-Error -Message "$computerName The IPv4 address of the computer could not be found" 
        return $null
    }
}

foreach ($NSName in $NSNameArr) {
    #Создаем сессию с Citrix Netscaler
    $Session = New-NSSession -NSName $NsName -Credential $Credential
    $NSNameEndpoint = $Session.Endpoint
    if (($null -ne $Session) -and (Get-NSHAPrimary -NSSession $Session -Verbose)) {

        #Циклом получаем имя компьютера из массива имен
        foreach ($computerName in $computersArr) {
            $computerName = ($computerName.ToLower()).Trim()
            #В переменную $groupNameFinal помещаем имя группы AD, соеденив шаблон имени и имя компьютера
            $groupNameFinal = $groupNameAhead + $computerName
            #В переменную $authzPolicyName помещаем имя политики авторизации, соеденив шаблон имени и имя компьютера
            $authzPolicyName = $authzPolicyNameAhead + $computerName
    
            #Проверяем есть ли DNS запись для компьютера и группа в AD, если нет группы, то создаем
            if ((Get-DNSRecordExists -computerName $computerName -Verbose) -and (!(Get-GroupADExists -groupNameAD $groupNameFinal -Credential $Credential -DCName $DC -Verbose))) {        
                New-GroupAD -groupNameAD $groupNameFinal -computerName $computerName -destinationionPath $OUPath -Credential $Credential -DCName $DC -Verbose
            }

            #Проверяем есть ли DNS запись для компьютера и есть ли в AD группа для добалвения на NS
            if ((Get-DNSRecordExists -computerName $computerName) -and (Get-GroupADExists -groupNameAD $groupNameFinal -Credential $Credential -DCName $DC)) {
                #Проверяем добавлена ли группа на NS, если нет, то добавляем
                if (!(Get-NSGroupExists -groupName $groupNameFinal -NSSession $Session -Verbose)) {
                    Add-NSgroup -groupName $groupNameFinal -NSSession $Session -Verbose
                }

                #Проверяем добавлен ли Bookmark на NS, если нет, то добавляем
                if (!(Get-NSBookmarkExists -bookmarkName $computerName -NSSession $Session -Verbose)) {
                    Add-NSBookmark -bookmarkName $computerName -NSSession $Session -Verbose
                }

                #Проверка на биндинг Bookmark-а к группе на NS, если нет, то выполняем биндинг
                if (Get-NSBookmarkBinding -bookmarkName $computerName -groupName $groupNameFinal -NSSession $Session -Verbose) {
                    Add-NSBindingBookmark -bookmarkName $computerName -groupName $groupNameFinal -NSSession $Session -Verbose
                }

                #Проверяем есть ли на NS Authorization Policy, если нет, то создаем
                if (!(Get-NSAuthzPolicyExists -authzPolicyName $authzPolicyName -NSSession $Session -Verbose)) {
                    Add-NSAuthzPolicy -authzPolicyName $authzPolicyName -computerName $computerName -NSSession $Session -Verbose
                }

                #Проверка на биндинг Authorization Policy к группе на NS, если нет, то выполняем биндинг
                if (Get-NSAuthzPolicyBinding -authzPolicyName $authzPolicyName -groupName $groupNameFinal -NSSession $Session -Verbose) {
                    Add-NSBindingAuthzPolicy -authzPolicyName $authzPolicyName -groupName $groupNameFinal -NSSession $Session -Verbose
                }

                $ipAddress = Get-ComputerIPv4Addres -computerName $computerName
                if ($null -ne $ipAddress) {
                    Get-NSRDPAccessCheck -ipAddress $ipAddress -computerName $computerName -NSServer $NSServerRDPcheck -NSService $NSServiceRDPcheck -NSSession $Session -Verbose
                }
        
            }  
            Write-Verbose " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" -Verbose

        }

        #Сохраняем config
        try {
            Save-NSConfig -NSSession $Session
            Write-Verbose "$NSNameEndpoint NS config saved successfully" -Verbose
        }
        catch {
            Write-Error -Message "Error $NSNameEndpoint Failed to save config"
        }
    }
}


Read-Host "To exit, press 'Enter'"
