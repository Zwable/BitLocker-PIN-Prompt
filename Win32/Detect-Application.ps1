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
        [Parameter(Mandatory=$false, HelpMessage="Provide level, 1 = default, 2 = warning 3 = error")][bool]$WriteToEventViewer = $false,
        [Parameter(Mandatory=$false, HelpMessage="To write to host or not")][switch]$WriteToHost
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
            if ($WriteToHost) {
                Write-Host ("[$($ComponentName):$($Scriptline)]" + $logOutput)
            }
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
function Test-BitLockerTpmAndPINProtectorSet {

    # Determine if PIN is configured
    $Protectors = (Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector
    $HasTPM = $Protectors | Where-Object { $_.KeyProtectorType -eq 'Tpm' }
    $HasTPMPin = $Protectors | Where-Object { $_.KeyProtectorType -eq 'TpmPin' }

    if ($HasTPM -and $HasTPMPin) {
        Write-Log -LogOutput ("BitLocker TPM and TPMPIN is defined, cleanup needed to ensure Pre-Boot authentication") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $true
    }else{
        Write-Log -LogOutput ("BitLocker TPM and TPMPIN is not defined at the same time") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
}
#Write log
Write-Log -LogOutput ((" ----------------------------------------------------- Checking values ({0}) ----------------------------------------------------- ") -f $UserName) -ComponentName "Main" -Path $LogLocation -Name $LogName

# Test to see if OOBE is running
if (Test-RunningOOBE) {
    Write-Log -LogOutput ("OOBE is running, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
    return 0
}

#Check if app is already running
$RunningPID = Get-RegistryKey -Key $RegistryKeyPath -Name "PID"
if( -not [string]::IsNullOrEmpty($RunningPID) ) {
    if (Get-Process -PID $RunningPID -ErrorAction SilentlyContinue | Where-Object{$_.ProcessName -eq "powershell"}) {
        Write-Log -LogOutput ("Application is already running, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
        return 0
    }
}else {
        Write-Log -LogOutput ("Application is not running..") -ComponentName "Main" -Path $LogLocation -Name $LogName
}

# Testing if any logged on users
if (!(Test-IfAnyLoggedOnUsers)) {
    Write-Log -LogOutput ("No interactive user found, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
    return 0
}

# Check if machine is PIN ready
if (!(Test-RegistryKey -Key "HKLM:\SOFTWARE\Policies\Microsoft\FVE" -Name "UseTPMPIN" -Value "2")`
  -or !(Test-RegistryKey -Key "HKLM:\SOFTWARE\Policies\Microsoft\FVE" -Name "UseTPM" -Value "2")) {
    Write-Log -LogOutput ("Not configured for PIN enablement, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
    return 0
}else {
    Write-Log -LogOutput ("Configured for PIN enablement..") -ComponentName "Main" -Path $LogLocation -Name $LogName
}

# If bitlocker should be resumed
if (Test-BitLockerTpmAndPINProtectorSet) {
    Write-Log -LogOutput ("BitLocker TPM and TPMPIN is defined, cleanup needed to ensure Pre-Boot authentication") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
    return 0
}else {
    Write-Log -LogOutput ("BitLocker TPM and TPMPIN is not defined at the same time") -ComponentName "Main" -Path $LogLocation -Name $LogName
}

# Check if PIN has already been configured
if (Test-BitLockerPINProtectorSet) {
    Write-Log -LogOutput ("PIN Key Protector already set, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName -WriteToHost
    return 0
}else {
    Write-Log -LogOutput ("PIN Key Protector not set, triggering prompt..") -ComponentName "Main" -Path $LogLocation -Name $LogName
}

