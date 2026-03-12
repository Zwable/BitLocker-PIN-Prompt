#Global variables
[string]$UserName = [Environment]::UserName
[string]$Global:LogLocation = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
[string]$RegistryKeyPath = "HKLM:\Software\EndpointAdmin\BitlockerPin"
[string]$Global:LogName = "Invoke-TestBitlockerPIN"
[int]$Global:ExitCode = 0 #Global for usage inside functions
$ErrorActionPreference = 'SilentlyContinue'

function Write-Log {
    param 
    (
        [Parameter(Mandatory=$true, HelpMessage="Provide a message")][string]$LogOutput,
        [Parameter(Mandatory=$true, HelpMessage="Provide the function name")][string]$ComponentName,
        [Parameter(Mandatory=$false, HelpMessage="Provide the scriptlinenumber")][string]$ScriptLine,
        [Parameter(Mandatory=$false, HelpMessage="Provide path, default is .\Logs")][string]$Path = "$PSScriptRoot\Logs",
        [Parameter(Mandatory=$false, HelpMessage="Provide name for the log")][string]$Name,
        [Parameter(Mandatory=$false, HelpMessage="Provide level, 1 = default, 2 = warning 3 = error")][ValidateSet(1, 2, 3)][int]$LogLevel = 1,
        [Parameter(Mandatory=$false, HelpMessage="Provide level, 1 = default, 2 = warning 3 = error")][bool]$WriteToEventViewer = $false
    )

    #If the scriptline is not defined then use from the invocation
    If(!($ScriptLine)){
        $ScriptLine = $($MyInvocation.ScriptLineNumber)
    }

    if($LogOutput){

        #Date for the lognaming
        $FullLogName = ($Path + "\" + $Name + ".log")
        $FullSecodaryLogName = ($FullLogName).Replace(".log",".lo_")

        #If the log has reached over xx mb then rename it
        if(Test-Path $FullLogName){
            if((Get-Item $FullLogName).Length -gt 1024kb){
                if(Test-Path $FullSecodaryLogName){
                    Remove-Item -Path $FullSecodaryLogName -force
                }
                Rename-Item -Path $FullLogName -NewName $FullSecodaryLogName
            }
        }

        #First check if folder/logfile exists, if not then create it
        if(!(Test-Path $Path)){
            New-Item -ItemType Directory -Force -Path $Path -ErrorAction SilentlyContinue | Out-Null
        }

        #Get current date and time to write to log
        $TimeGenerated =  (Get-Date -Format "HH':'mm':'ss.fffffff")

        #Construct the logline format
        $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">'

        #Define line
        $LineFormat = $LogOutput, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($ComponentName):$($Scriptline)",$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),$LogLevel, ([Threading.Thread]::CurrentThread.ManagedThreadId)

        #Create line
        $Line = $Line -f $LineFormat

        #Write log
        try {

            #Write file (using Set-Content or Out-File would result in UFT8WithBom even with the encoding parameter defined, resulting in wrong encoding)
            Write-Host ("[$($ComponentName):$($Scriptline)]" + $logOutput)
            $Encoding = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::AppendAllLines($FullLogName, [string[]]$Line,$Encoding)
        }
        catch {
            # Write-Host "$_"
        }

        #write windows event
        if(($LogLevel -ne 1) -and ($WriteToEventViewer -ne $false)){
            Write-EventLog -LogName "Application" -Source $ComponentName -EventID 3005 -EntryType Information -Message $logOutput -Category $LogLevel# -RawData 10,20
        }
    }
}
function Convert-RegistryKey {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Provide a Key location to convert")][string]$Key
    )

    #Convert the registry key hive to the full path, only match if at the beginning of the line
    If ($Key -match '^HKLM') {
        $Key = $Key -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\' -replace '^HKLM:', 'HKEY_LOCAL_MACHINE\' -replace '^HKLM\\', 'HKEY_LOCAL_MACHINE\'
    }elseif ($Key -match '^HKCR') {
        $Key = $Key -replace '^HKCR:\\', 'HKEY_CLASSES_ROOT\' -replace '^HKCR:', 'HKEY_CLASSES_ROOT\' -replace '^HKCR\\', 'HKEY_CLASSES_ROOT\'
    }elseif ($Key -match '^HKCU') {
        $Key = $Key -replace '^HKCU:\\', 'HKEY_CURRENT_USER\' -replace '^HKCU:', 'HKEY_CURRENT_USER\' -replace '^HKCU\\', 'HKEY_CURRENT_USER\'
    }elseif ($Key -match '^HKU') {
        $Key = $Key -replace '^HKU:\\', 'HKEY_USERS\' -replace '^HKU:', 'HKEY_USERS\' -replace '^HKU\\', 'HKEY_USERS\'
    }elseif ($Key -match '^HKCC') {
        $Key = $Key -replace '^HKCC:\\', 'HKEY_CURRENT_CONFIG\' -replace '^HKCC:', 'HKEY_CURRENT_CONFIG\' -replace '^HKCC\\', 'HKEY_CURRENT_CONFIG\'
    }elseif ($Key -match '^HKPD') {
        $Key = $Key -replace '^HKPD:\\', 'HKEY_PERFORMANCE_DATA\' -replace '^HKPD:', 'HKEY_PERFORMANCE_DATA\' -replace '^HKPD\\', 'HKEY_PERFORMANCE_DATA\'
    }
    If ($Key -notmatch '^Registry::') {[string]$Key = "Registry::$Key" }
    
    #Return the key
    return $key
}
function Get-RegistryHiveDisplayName {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Provide a Key location to get displayname")][string]$Key
    )

    #Convert the registry key hive to the full path, only match if at the beginning of the line
    If (($Key -match '^HKLM') -or ($Key -match '^HKEY_LOCAL_MACHINE')-or ($Key -match '^Registry::HKEY_LOCAL_MACHINE')) {
        $DisplayName = "LocalMachine"
    }elseif (($Key -match '^HKCR') -or ($Key -match '^HKEY_CLASSES_ROOT') -or ($Key -match '^Registry::HKEY_CLASSES_ROOT')) {
        $DisplayName = "ClassesRoot"
    }elseif (($Key -match '^HKCU') -or ($Key -match '^HKEY_CURRENT_USER') -or ($Key -match '^Registry::HKEY_CURRENT_USER')) {
        $DisplayName = "CurrentUser"
    }elseif (($Key -match '^HKU') -or ($Key -match '^HKEY_USERS') -or ($Key -match '^Registry::HKEY_USERS')) {
        $DisplayName = "Users"
    }elseif (($Key -match '^HKCC') -or ($Key -match '^HKEY_CURRENT_CONFIG') -or ($Key -match '^Registry::HKEY_CURRENT_CONFIG')) {
        $DisplayName = "CurrentConfig"
    }

    #Return
    Return $DisplayName
}
function Test-RegistryKey {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Provide a Key location")][string]$Key,
        [Parameter(Mandatory=$false, HelpMessage="Provide the Name of the registry")][string]$Name,
        [Parameter(Mandatory=$false, HelpMessage="Provide the Data of the registry")]$Value,
        [Parameter(Mandatory=$false, HelpMessage="Provide the SID of the user")][string]$SID
    )

    #Replace sid if present
    $Key = (Convert-RegistryKey -Key $Key) -f $SID

    #Get the path properties
    if (Test-Path -Path $Key) {
        if ($Name) {
            if ($Name -like '(Default)') {
                [string]$KeyValue = $(Get-Item -LiteralPath $Key).GetValue($null,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            }else{
                [string]$KeyValue = $(Get-Item -LiteralPath $Key).GetValue($Name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            }
            if ($KeyValue -eq $Value) {
                Write-Log -LogOutput ("Registry key '{0}' with the name '{1}' with the data '{2}' exist" -f $Key,$Name,$Value) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
                Return $true
            }else {
                Write-Log -LogOutput ("Registry key '{0}' with the name '{1}' with the data '{2}' does not exist" -f $Key,$Name,$Value) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName -LogLevel 2
                $Global:ExitCode = 1
                Return $false
            }
        }else {
            Write-Log -LogOutput ("Registry key '{0}' exist" -f $Key) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
            Return $true
        }
    }else {
        Write-Log -LogOutput ("Registry key '{0}' with the name '{1}' with the data '{2}' does not exist" -f $Key,$Name,$Value) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName -LogLevel 2
        $Global:ExitCode = 1
        Return $false
    }
}
function Get-RegistryKey {
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Provide a Key location")][string]$Key,
        [Parameter(Mandatory = $true, HelpMessage = "Provide the Name of the registry")][string]$Name
    )

    # Get the path properties
    $regKey = $(Get-Item -LiteralPath $Key -ErrorAction SilentlyContinue)

    if($null -ne $regKey) {
        [string]$keyValue = $regKey.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    }

    # Return the key
    return $keyValue
}
function Test-IfAnyLoggedOnUsers {

    #Using Query session (property names returned by quser is OS language specific)
    $LoggedonUsers = ((quser) -replace '\s{2,}', ',') -replace '>','' | ConvertFrom-Csv
    if ($LoggedonUsers) {
        foreach ($LoggedOnUser in $LoggedonUsers) {
            Write-Log -LogOutput ("Found logged on user") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        }
    }else{
        Write-Log -LogOutput ("No logged on users found..") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
    return $true
}
function Test-RunningOOBE {
    # Check if OOBE / ESP is running [credit Michael Niehaus]
    $TypeDef = @"

using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Api
{
 public class Kernel32
 {
   [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   public static extern int OOBEComplete(ref int bIsOOBEComplete);
 }
}
"@

        Add-Type -TypeDefinition $TypeDef -Language CSharp
        $IsOOBEComplete = $false
        $hr = [Api.Kernel32]::OOBEComplete([ref] $IsOOBEComplete)
        If (!($IsOOBEComplete)) {
            Write-Log -LogOutput ("OOBE is running..") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
            return $true
        }else {
            Write-Log -LogOutput ("OOBE is not running..") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
            return $false
        }
}
function Test-BitLockerShouldBeResumed {

    # Determine if BitLocker is not currently awaiting a reboot and protection status is off, this indicates that BitLocker has not been correctly resumed
    [int]$SuspendCount = (Get-CimInstance -Namespace "ROOT/CIMV2/Security/MicrosoftVolumeEncryption" -Class Win32_EncryptableVolume -Filter "DriveLetter='$env:SystemDrive'" | Invoke-CimMethod -MethodName "GetSuspendCount").SuspendCount
    [string]$ProtectionStatus = (Get-BitLockerVolume -MountPoint $env:SystemDrive).ProtectionStatus
    if (!($SuspendCount -gt 0) -and $ProtectionStatus -eq "Off") {
        Write-Log -LogOutput ("Current SuspendCount is '{0}' and the ProtectionStatus is '{1}', BitLocker should be resumed" -f $SuspendCount,$ProtectionStatus) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $true
    }else {
        Write-Log -LogOutput ("Current SuspendCount is '{0}' and the ProtectionStatus is '{1}', BitLocker should not be resumed" -f $SuspendCount,$ProtectionStatus) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
}
function Test-BitLockerPINProtectorSet {

    # Determine if PIN is configured
    $PINSet = $(Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector | Where-Object { $_.KeyProtectorType -eq 'TpmPin' }
    if ($PINSet) {
        Write-Log -LogOutput ("BitLocker PIN is defined") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $true
    }else{
        Write-Log -LogOutput ("BitLocker PIN is not defined") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
}
function Test-IfVirtualMachine {

    # Below is taken from PSADT
    $hwBios = Get-WmiObject -Class 'Win32_BIOS' -ErrorAction 'Stop' | Select-Object -Property 'Version', 'SerialNumber'
    $hwMakeModel = Get-WmiObject -Class 'Win32_ComputerSystem' -ErrorAction 'Stop' | Select-Object -Property 'Model', 'Manufacturer'

    If ($hwBIOS.Version -match 'VRTUAL') {
        $hwType = 'Virtual:Hyper-V'
    }
    ElseIf ($hwBIOS.Version -match 'A M I') {
        $hwType = 'Virtual:Virtual PC'
    }
    ElseIf ($hwBIOS.Version -like '*Xen*') {
        $hwType = 'Virtual:Xen'
    }
    ElseIf ($hwBIOS.SerialNumber -like '*VMware*') {
        $hwType = 'Virtual:VMWare'
    }
    ElseIf ($hwBIOS.SerialNumber -like '*Parallels*') {
        $hwType = 'Virtual:Parallels'
    }
    ElseIf (($hwMakeModel.Manufacturer -like '*Microsoft*') -and ($hwMakeModel.Model -notlike '*Surface*')) {
        $hwType = 'Virtual:Hyper-V'
    }
    ElseIf ($hwMakeModel.Manufacturer -like '*VMWare*') {
        $hwType = 'Virtual:VMWare'
    }
    ElseIf ($hwMakeModel.Manufacturer -like '*Parallels*') {
        $hwType = 'Virtual:Parallels'
    }
    ElseIf ($hwMakeModel.Model -like '*Virtual*') {
        $hwType = 'Virtual'
    }
    Else {
        $hwType = 'Physical'
    }

    if ($hwType -match "VIRTUAL") {
        Write-Log -LogOutput ("Workstation is on virtual environment, HW type '{0}'" -f $HWBIOS.Version) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $true
    }else {
        Write-Log -LogOutput ("Workstation is on physical environment, HW type '{0}'" -f $HWBIOS.Version) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
}
function Exit-Script {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Provide exit code")][string]$ExitCode
    )
    
    #Exit
    if($ExitCode -eq 0){
        Write-Log -LogOutput ("Exiting script with exitcode '{0}' remediation not required" -f $ExitCode) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        Exit $ExitCode
    }else {
        Write-Log -LogOutput ("Exiting script with exitcode '{0}' remediation required" -f $ExitCode) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName -LogLevel 3
        Exit $ExitCode
    }
}

#Write log
Write-Log -LogOutput ((" ----------------------------------------------------- Checking values ({0}) ----------------------------------------------------- ") -f $UserName) -ComponentName "Main" -Path $LogLocation -Name $LogName

# Test if virtual machine 
if (Test-IfVirtualMachine) {
    Write-Log -LogOutput ("Running on virtual environment, skipping remediation..") -ComponentName "Main" -Path $LogLocation -Name $LogName
    Exit-Script -ExitCode 0
}

# Test to see if OOBE is running
if (Test-RunningOOBE) {
    Write-Log -LogOutput ("OOBE is running, skipping remediation..") -ComponentName "Main" -Path $LogLocation -Name $LogName
    Exit-Script -ExitCode 0
}

# Testing if any logged on users
if (!(Test-IfAnyLoggedOnUsers)) {
    Write-Log -LogOutput (("No interactive user found, exiting.. {0}" -f ($ActiveUserNames))) -ComponentName "Main" -Path $LogLocation -Name $LogName
    Exit-Script -ExitCode 0
}

# Check if machine is PIN ready
if (!(Test-RegistryKey -Key "HKLM:\SOFTWARE\Policies\Microsoft\FVE" -Name "UseTPMPIN" -Value "2")`
  -or !(Test-RegistryKey -Key "HKLM:\SOFTWARE\Policies\Microsoft\FVE" -Name "UseTPM" -Value "2")) {
    Write-Log -LogOutput ("Not configured for PIN enablement, skipping remediation..") -ComponentName "Main" -Path $LogLocation -Name $LogName
    Exit-Script -ExitCode 0
}

#Check if app is already running
$RunningPID = Get-RegistryKey -Key $RegistryKeyPath -Name "PID"
if( -not [string]::IsNullOrEmpty($RunningPID) ) {
    if (Get-Process -PID $RunningPID -ErrorAction SilentlyContinue | Where-Object{$_.ProcessName -eq "powershell"}) {
        Write-Log -LogOutput ("Application with PID '$($RunningPID)' is already running, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName
        Exit-Script -ExitCode 0
    }
}else {
        Write-Log -LogOutput ("Application with PID '$($RunningPID)' is not running..") -ComponentName "Main" -Path $LogLocation -Name $LogName
}

# If bitlocker should be resumed
if (Test-BitLockerShouldBeResumed) {
    Exit-Script -ExitCode 1
}

# Check if PIN has already been configured
if (!(Test-BitLockerPINProtectorSet)) {
    Write-Log -LogOutput ("Missing PIN Protector, remediation required..") -ComponentName "Main" -Path $LogLocation -Name $LogName
    Exit-Script -ExitCode 1
}

#Exit
Exit-Script -ExitCode 0
