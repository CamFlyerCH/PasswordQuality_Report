# 04.12.2023 - Prepare_PasswordQuality_Report.ps1 by https://github.com/CamFlyerCH

# Set path to sorted password hash file from haveibeenpwnd
$SortedHashFile = "D:\pwnedpasswords_ntlm.txt"

# Init
Import-Module DSInternals
Import-Module ActiveDirectory
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path -Parent $ScriptPath
$AccountList = @()
$ADDomain = Get-ADDomain

# Run PassowrdQualityCheck
$PWQualityData = Get-ADReplAccount -All -Server $ADDomain.PDCEmulator -NamingContext $ADDomain.DistinguishedName | Test-PasswordQuality -WeakPasswordHashesFile $SortedHashFile
$PWQualityData > ($ScriptDir + "\PasswordQuality_" + $ADDomain.DNSRoot + "_PWQ-Data.txt")
$PWQualityData | Export-Clixml -Path ($ScriptDir + "\PasswordQuality_" + $ADDomain.DNSRoot + "_PWQ-Data.xml")

# Get accounts
$ADUsers = Get-ADUser -Filter {Enabled -eq $True} -SearchScope Subtree -Properties CanonicalName,Created,Modified,Manager,Description,DisplayName,LastLogonDate,PasswordLastSet,PasswordNeverExpires,Enabled,Mail,userPrincipalName | Sort-Object CanonicalName
ForEach ($ADUser in $ADUsers){
    $Manager = $NULL
    If ($ADUser.Manager){
        Try{
            $Manager = Get-ADObject -Identity $ADUser.Manager -Properties sAMAccountName | Select-Object -ExpandProperty sAMAccountName
        } Catch {Continue}
    }
    $AccountList += $ADUser | Select-Object sAMAccountName,Created,Modified,@{n="LastLogon";e={$_.LastLogonDate}},@{n="PwLastSet";e={$_.PasswordLastSet}},@{n="PwNeverExpires";e={$_.PasswordNeverExpires}},@{n="ManagedBy";e={$Manager}},Description,DisplayName,Mail,@{n="UPN";e={$_.userPrincipalName}},CanonicalName,@{n="ObjectType";e={"User"}}
}

$ADComputers = Get-ADComputer -Filter {Enabled -eq $True} -SearchScope Subtree -Properties CanonicalName,Created,Modified,Manager,Description,DisplayName,LastLogonDate,PasswordLastSet,Enabled,OperatingSystem,dNSHostName | Sort-Object CanonicalName
ForEach ($ADComputer in $ADComputers){
    $Manager = $NULL
    If ($ADComputer.Manager){
        Try{
            $Manager = Get-ADObject -Identity $ADComputer.Manager -Properties sAMAccountName | Select-Object -ExpandProperty sAMAccountName
        } Catch {Continue}
    }
    $AccountList += $ADComputer | Select-Object sAMAccountName,Created,Modified,@{n="LastLogon";e={$_.LastLogonDate}},@{n="PwLastSet";e={$_.PasswordLastSet}},@{n="PwNeverExpires";e={$Null}},@{n="ManagedBy";e={$Manager}},Description,@{n="DisplayName";e={$_.DisplayName}},@{n="Mail";e={""}},@{n="UPN";e={$_.dNSHostName}},CanonicalName,@{n="ObjectType";e={"Computer"}}
}

# Export accounts
$AccountList | Export-Clixml -Path ($ScriptDir + "\PasswordQuality_" + $ADDomain.DNSRoot + "_Accounts.xml")
