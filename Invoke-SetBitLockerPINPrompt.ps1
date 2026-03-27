<#

    .SYNOPSIS

    .DESCRIPTION

    .PARAMETER

    .EXAMPLE
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"& { . '.\Start-SetBitLockerPINPrompt.ps1' -UseCustomPSADTBanner `$true -CancelClosingWindow `$false}`""
    .NOTES
    Author: Morten Rønborg
    Date: 2025-06-20
    Email: mr@endpointadmin.com

#>
################################################
param (
    [Parameter(Mandatory=$false, HelpMessage="To ensure script is re-run interactive for the user as SYSTEM")][switch]$InvokeToInteractiveUser,
    [Parameter(Mandatory=$false, HelpMessage="To show custom banner or not")][bool]$UseCustomPSADTBanner = $true,
    [Parameter(Mandatory=$false, HelpMessage="To cancel user closing window or not")][bool]$CancelClosingWindow = $false
)

#Variables
[bool]$RunningAsSystem = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18'

# Global variables (description defined as unique identifier when sharing variables between runspaces)
New-Variable -Name "RegistryKeyPath" -Value "HKLM:\Software\EndpointAdmin\BitlockerPin" -Description "Shared var between runspaces" -Scope Global
New-Variable -Name "LogLocation" -Value "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs" -Description "Shared var between runspaces" -Scope Global
New-Variable -Name "LogName" -Value "Invoke-SetBitLockerPINPrompt" -Description "Shared var between runspaces" -Scope Global
New-Variable -Name "CancelUserClosingWindow" -Value $CancelClosingWindow -Description "Shared var between runspaces" -Scope Global
New-Variable -Name "DryRun" -Value $false -Description "Shared var between runspaces" -Scope Global
New-Variable -Name "DryRunEnforceError" -Value $false -Description "Shared var between runspaces" -Scope Global # when used, DryRun must also be true
New-Variable -Name "UseBanner" -Value $UseCustomPSADTBanner -Description "Shared var between runspaces" -Scope Global

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
function Start-LoadAssemblys {

    #Add assemblys
    Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration
}
function Start-RunAsSystemPresentToInteractiveUser {
    param(
        [string]$Path,
        [bool]$UseCustomPSADTBanner = $false,
        [bool]$CancelClosingWindow = $false
    )

    $ApplicationLoader = @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security;

public class ApplicationLoader
{

    private const uint INVALID_SESSION_ID = 0xFFFFFFFF;
    private static readonly IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;
    public const uint SW_HIDE = 0x00000000;
    private enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,
        WTSConnected,
        WTSConnectQuery,
        WTSShadow,
        WTSDisconnected,
        WTSIdle,
        WTSListen,
        WTSReset,
        WTSDown,
        WTSInit
    }
    private enum WTS_INFO_CLASS
    {
        WTSInitialProgram,
        WTSApplicationName,
        WTSWorkingDirectory,
        WTSOEMId,
        WTSSessionId,
        WTSUserName,
        WTSWinStationName,
        WTSDomainName,
        WTSConnectState,
        WTSClientBuildNumber,
        WTSClientName,
        WTSClientDirectory,
        WTSClientProductId,
        WTSClientHardwareId,
        WTSClientAddress,
        WTSClientDisplay,
        WTSClientProtocolType
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct WTS_SESSION_INFO
    {
        public readonly UInt32 SessionID;

        [MarshalAs(UnmanagedType.LPStr)]
        public readonly String pWinStationName;

        public readonly WTS_CONNECTSTATE_CLASS State;
    }
    [DllImport("wtsapi32.dll", SetLastError = true)]
    private static extern bool WTSQuerySessionInformation(
        System.IntPtr hServer,
        uint sessionId,
        WTS_INFO_CLASS wtsInfoClass,
        out System.IntPtr ppBuffer,
        out uint pBytesReturned);
    [DllImport("wtsapi32.dll", SetLastError = true)]
    private static extern int WTSEnumerateSessions(
        IntPtr hServer,
        int Reserved,
        int Version,
        out IntPtr ppSessionInfo,
        out int pCount);

    [DllImport("wtsapi32.dll")]
    private static extern void WTSFreeMemory(IntPtr pMemory);

    [StructLayout(LayoutKind.Sequential)]
    public struct SECURITY_ATTRIBUTES
    {
        public int Length;
        public IntPtr lpSecurityDescriptor;
        public bool bInheritHandle;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct STARTUPINFO
    {
        public int cb;
        public String lpReserved;
        public String lpDesktop;
        public String lpTitle;
        public uint dwX;
        public uint dwY;
        public uint dwXSize;
        public uint dwYSize;
        public uint dwXCountChars;
        public uint dwYCountChars;
        public uint dwFillAttribute;
        public uint dwFlags;
        public short wShowWindow;
        public short cbReserved2;
        public IntPtr lpReserved2;
        public IntPtr hStdInput;
        public IntPtr hStdOutput;
        public IntPtr hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION
    {
        public IntPtr hProcess;
        public IntPtr hThread;
        public uint dwProcessId;
        public uint dwThreadId;
    }

    enum TOKEN_TYPE : int
    {
        TokenPrimary = 1,
        TokenImpersonation = 2
    }

    enum SECURITY_IMPERSONATION_LEVEL : int
    {
        SecurityAnonymous = 0,
        SecurityIdentification = 1,
        SecurityImpersonation = 2,
        SecurityDelegation = 3,
    }

    public const int TOKEN_DUPLICATE = 0x0002;
    public const uint MAXIMUM_ALLOWED = 0x2000000;
    public const int CREATE_NEW_CONSOLE = 0x00000010;

    public const int IDLE_PRIORITY_CLASS = 0x40;
    public const int NORMAL_PRIORITY_CLASS = 0x20;
    public const int HIGH_PRIORITY_CLASS = 0x80;
    public const int REALTIME_PRIORITY_CLASS = 0x100;

    public const uint INFINITE = 0xFFFFFFFF;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hSnapshot);
    [DllImport("kernel32.dll")]
    static extern uint WTSGetActiveConsoleSessionId();
    [DllImport("kernel32.dll")]
    static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);
    [DllImport("kernel32.dll")]
    static extern bool GetExitCodeProcess(IntPtr hProcess, ref int exitCode);


    [DllImport("advapi32.dll", EntryPoint = "CreateProcessAsUser", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
    public extern static bool CreateProcessAsUser(IntPtr hToken, String lpApplicationName, String lpCommandLine, ref SECURITY_ATTRIBUTES lpProcessAttributes,
        ref SECURITY_ATTRIBUTES lpThreadAttributes, bool bInheritHandle, int dwCreationFlags, IntPtr lpEnvironment,
        String lpCurrentDirectory, ref STARTUPINFO lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll")]
    static extern bool ProcessIdToSessionId(uint dwProcessId, ref uint pSessionId);

    [DllImport("advapi32.dll", EntryPoint = "DuplicateTokenEx")]
    public extern static bool DuplicateTokenEx(IntPtr ExistingTokenHandle, uint dwDesiredAccess,
        ref SECURITY_ATTRIBUTES lpThreadAttributes, int TokenType,
        int ImpersonationLevel, ref IntPtr DuplicateTokenHandle);

    [DllImport("kernel32.dll")]
    static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

    [DllImport("advapi32", SetLastError = true), SuppressUnmanagedCodeSecurity]
    static extern bool OpenProcessToken(IntPtr ProcessHandle, int DesiredAccess, ref IntPtr TokenHandle);

    /// <summary>
    /// Launches the given application with full admin rights, and in addition bypasses the Vista UAC prompt
    /// </summary>
    /// <param name="applicationName">The name of the application to launch</param>
    /// <param name="procInfo">Process information regarding the launched application that gets returned to the caller</param>
    /// <returns></returns>
    public static bool StartProcessAndBypassUAC(String applicationName, string startingDir, out PROCESS_INFORMATION procInfo)
    {
        uint winlogonPid = 0;
        IntPtr hUserTokenDup = IntPtr.Zero, hPToken = IntPtr.Zero, hProcess = IntPtr.Zero;
        procInfo = new PROCESS_INFORMATION();

        // obtain the currently active session id; every logged on user in the system has a unique session id
        uint dwSessionId = GetActiveUserSessionID();

        // obtain the process id of the winlogon process that is running within the currently active session
        Process[] processes = Process.GetProcessesByName("winlogon");
        foreach (Process p in processes)
        {
            if ((uint)p.SessionId == dwSessionId)
            {
                winlogonPid = (uint)p.Id;
            }
        }

        // obtain a handle to the winlogon process
        hProcess = OpenProcess(MAXIMUM_ALLOWED, false, winlogonPid);

        // obtain a handle to the access token of the winlogon process
        if (!OpenProcessToken(hProcess, TOKEN_DUPLICATE, ref hPToken))
        {
            CloseHandle(hProcess);
            return false;
        }

        // Security attibute structure used in DuplicateTokenEx and CreateProcessAsUser
        // I would prefer to not have to use a security attribute variable and to just 
        // simply pass null and inherit (by default) the security attributes
        // of the existing token. However, in C# structures are value types and therefore
        // cannot be assigned the null value.
        SECURITY_ATTRIBUTES sa = new SECURITY_ATTRIBUTES();
        sa.Length = Marshal.SizeOf(sa);

        // copy the access token of the winlogon process; the newly created token will be a primary token
        if (!DuplicateTokenEx(hPToken, MAXIMUM_ALLOWED, ref sa, (int)SECURITY_IMPERSONATION_LEVEL.SecurityIdentification, (int)TOKEN_TYPE.TokenPrimary, ref hUserTokenDup))
        {
            CloseHandle(hProcess);
            CloseHandle(hPToken);
            return false;
        }

        // By default CreateProcessAsUser creates a process on a non-interactive window station, meaning
        // the window station has a desktop that is invisible and the process is incapable of receiving
        // user input. To remedy this we set the lpDesktop parameter to indicate we want to enable user 
        // interaction with the new process.
        STARTUPINFO si = new STARTUPINFO();
        si.cb = (int)Marshal.SizeOf(si);
        si.lpDesktop = @"winsta0\default"; // interactive window station parameter; basically this indicates that the process created can display a GUI on the desktop
        si.dwFlags = SW_HIDE; // Set the window style to hidden

        // flags that specify the priority and creation method of the process
        int dwCreationFlags = NORMAL_PRIORITY_CLASS;

        // create a new process in the current user's logon session
        bool result = CreateProcessAsUser(hUserTokenDup,        // client's access token
                                        null,                   // file to execute
                                        applicationName,        // command line
                                        ref sa,                 // pointer to process SECURITY_ATTRIBUTES
                                        ref sa,                 // pointer to thread SECURITY_ATTRIBUTES
                                        false,                  // handles are not inheritable
                                        dwCreationFlags,        // creation flags
                                        IntPtr.Zero,            // pointer to new environment block 
                                        startingDir,            // name of current directory 
                                        ref si,                 // pointer to STARTUPINFO structure
                                        out procInfo            // receives information about new process
                                        );

        // invalidate the handles
        CloseHandle(hProcess);
        CloseHandle(hPToken);
        CloseHandle(hUserTokenDup);

        return result; // return the result
    }
    public static uint GetActiveUserSessionID(string user_filter = null)
    {

        var activeSessionId = INVALID_SESSION_ID;
        var pSessionInfo = IntPtr.Zero;
        var sessionCount = 0;

        IntPtr userPtr = IntPtr.Zero;
        IntPtr domainPtr = IntPtr.Zero;
        uint bytes = 0;

        // Get a handle to the user access token for the current active session.
        if (WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, out pSessionInfo, out sessionCount) != 0)
        {
            var arrayElementSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
            var current = pSessionInfo;

            for (var i = 0; i < sessionCount; i++)
            {
                var si = (WTS_SESSION_INFO)Marshal.PtrToStructure((IntPtr)current, typeof(WTS_SESSION_INFO));
                current += arrayElementSize;

                WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, si.SessionID, WTS_INFO_CLASS.WTSUserName, out userPtr, out bytes);
                WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, si.SessionID, WTS_INFO_CLASS.WTSDomainName, out domainPtr, out bytes);

                var user = Marshal.PtrToStringAnsi(userPtr);
                var domain = Marshal.PtrToStringAnsi(domainPtr);

                WTSFreeMemory(userPtr);
                WTSFreeMemory(domainPtr);

                if ((user_filter == null && si.State == WTS_CONNECTSTATE_CLASS.WTSActive) || (user == user_filter))
                {
                    activeSessionId = si.SessionID;
                }

            }
        }

        // If enumerating did not work, fall back to the old method
        if (activeSessionId == INVALID_SESSION_ID)
        {
            activeSessionId = WTSGetActiveConsoleSessionId();
        }

        //return bResult;
        return activeSessionId;
    }
    public void Execute(string applicationName)
    {
        ApplicationLoader.PROCESS_INFORMATION procInfo;

        ApplicationLoader.StartProcessAndBypassUAC(applicationName, null, out procInfo);
    }

}
"@

    #Add the code as a type
    Add-Type -TypeDefinition $ApplicationLoader
    
    #Add the application loader
    $ApplicationLoaderObj = New-Object ApplicationLoader

    #Use powershell
    # $CommandLine = (("powershell.exe -ExecutionPolicy Bypass -nologo -noprofile -windowstyle hidden -File `"{0}`" -InvokeToInteractiveUser -UseCustomPSADTBanner:$true") -f $Path)
    $CommandLine = (("powershell.exe -ExecutionPolicy Bypass -nologo -noprofile -windowstyle hidden -Command `"& { . '$($Path)' -InvokeToInteractiveUser -UseCustomPSADTBanner `$$($UseCustomPSADTBanner) -CancelClosingWindow `$$($CancelClosingWindow) }`""))
    
    #Do execution
    Write-Log -LogOutput ("Executing '{0}'" -f ($CommandLine)) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
    $ApplicationLoaderObj.Execute($CommandLine)
}
function Set-PinAcceptedState {

    if (!(Test-String -InputString $syncHash.pbx_passwordfirst.Password -MinimumPINLength $MinimumPIN -EnhancedPIN $UseEnhancedPin)) {
        Write-Log -LogOutput ("Password do not meet criteria") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        $syncHash.tbx_passwordnotaccepted.Text = "Password do not meet criteria"
        $syncHash.stk_passwordnotaccepted.Visibility = "Visible"
        $syncHash.btn_accept.IsEnabled = $false
        return
    }

    if (!($syncHash.pbx_passwordfirst.Password -eq $syncHash.pbx_passwordsecond.Password)) {
        Write-Log -LogOutput ("Passwords do not match") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        $syncHash.tbx_passwordnotaccepted.Text = "Passwords do not match"
        $syncHash.tbx_passwordnoaccepted.Visibility = "Visible"
        $syncHash.stk_passwordnotaccepted.Visibility = "Visible"
        $syncHash.btn_accept.IsEnabled = $false
        return
    }

    $syncHash.stk_passwordnotaccepted.Visibility = "Hidden"
    $syncHash.btn_accept.IsEnabled = $true
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
function Get-UILocalShortName {

    #Map the user hive
    if(!(Test-Path -Path HKU:)){
        New-PSDrive HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
    }

    # Use this method as Get-UICulture does not reflect actual setting when running as system
    $CurrentLoggedOnUserSID = (New-Object -ComObject Microsoft.DiskQuota).TranslateLogonNameToSID((Get-WmiObject -Class Win32_ComputerSystem).Username)
    $UICultureShortname = Get-RegistryKey -Key "HKU:$CurrentLoggedOnUserSID\Control Panel\Desktop" -Name "PreferredUILanguages"
    if (-not [string]::IsNullOrEmpty($UICultureShortname)) {
        Write-Log -LogOutput ("Found PreferredUILanguages '{0}' for logged on user" -f $UICultureShortname) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
    } else {
        $UICultureShortname = (Get-UICulture).Name
        Write-Log -LogOutput ("No PreferredUILanguages defined, user has not changed language from native language '{0}'" -f $UICultureShortname) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
    }

    return $UICultureShortname 
}
function Test-String {
    param (
        [string]$InputString,
        [int]$MinimumPINLength = 4,
        [int]$MaximumPINLength = 20,
        [bool]$EnhancedPIN = $false
    )

    if (-not $InputString) {
        return $false
    }

    # Length check
    if ($InputString.Length -lt $MinimumPINLength -or $InputString.Length -gt $MaximumPINLength) {
        return $false
    }

    # Only digits allowed in standard PIN mode
    if (-not $EnhancedPIN -and $InputString -notmatch '^\d+$') {
        return $false
    }

    # Check for 3+ repeating characters
    if ($InputString -match '(.)\1{2,}') {
        return $false
    }

    # Check for 3+ sequential characters
    $sequences = @(
        "abcdefghijklmnopqrstuvwxyz",
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
        "0123456789"
    )

    foreach ($seq in $sequences) {
        for ($i = 0; $i -le $seq.Length - 3; $i++) {
            $sub = $seq.Substring($i, 3)
            if ($InputString.ToLower().Contains($sub.ToLower())) {
                return $false
            }
        }
    }

    return $true
}

function Set-RegistryKey {
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Provide a Key location")][string]$Key,
        [Parameter(Mandatory = $true, HelpMessage = "Provide the Name of the registry")][string]$Name,
        [Parameter(Mandatory = $true, HelpMessage = "Provide the Data of the registry")][string]$Value,
        [Parameter(Mandatory = $true, HelpMessage = "Provide the Type of the registry DWORD,STRING")][string]$Type
    )
    
    #Set key
    if (!(Test-Path -Path $Key)) {
        New-Item -Path $Key -Force | Out-Null
    }

    #Set value
    New-ItemProperty -Path $Key -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null 
}
function Update-Window {
    Param (
        $Control,
        $Property,
        $Value,
        [switch]$AppendContent
    )

    # This is kind of a hack, there may be a better way to do this
    If ($Property -eq "Close") {

        # The DialogResult only works when process is running as admin, when running as SYSTEM we need Exit()
        $syncHash.Window.Dispatcher.invoke([action] { $syncHash.Window.DialogResult = $true }, "Normal")
        # [System.Windows.Forms.Application]::Exit()
        Stop-Process $pid
    }

    # This updates the control based on the parameters passed to the function
    $syncHash.$Control.Dispatcher.Invoke([action] {
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
                $syncHash.$Control.AppendText($Value)
            }
            Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
}
function Set-WindowWorkArea {
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Middle or Corner")][string]$MiddleOrCorner
    )

    $ScreenWorkArea = [System.Windows.SystemParameters]::WorkArea
    #$ScreenWorkArea = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea

    if ($MiddleOrCorner -eq "Middle") {
        $syncHash.Window.Top = ($ScreenWorkArea.Height / 3) - ($syncHash.Window.Height / 3)
        $syncHash.Window.Left = ($ScreenWorkArea.Width / 2) - ($syncHash.Window.Width / 2)
    }
    else {
        $syncHash.Window.Top = ($ScreenWorkArea.Height - $syncHash.Window.Height)
        $syncHash.Window.Left = ($ScreenWorkArea.Width - ($syncHash.Window.Width))
    } 
}
function Show-NotificationMessage {

    #Load assemblys
    Start-LoadAssemblys

    # Detect current BL config
    $bitlockerPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\FVE'
    $minLength = 6 # the default min length if registry is not configured
    $minLength = Get-ItemPropertyValue -Path $bitlockerPolicyPath -Name 'MinimumPIN' -ErrorAction SilentlyContinue
    $allowSpecial = ($props = Get-ItemProperty -Path $bitlockerPolicyPath -ErrorAction SilentlyContinue).UseEnhancedPin -ne $null -and $props.UseEnhancedPin -ne 0 -and $props.UseEnhancedPin -ne ''
    New-Variable -Name "MinimumPIN" -Value $minLength -Description "Shared var between runspaces" -Scope Global
    New-Variable -Name "UseEnhancedPin" -Value $allowSpecial -Description "Shared var between runspaces" -Scope Global
    if ($allowSpecial) {
        $passwordText = "BitLocker protects this device to prevent unauthorized access to your files and documents, even if someone has physical access. To help keep your data safe, you will need to enter a code each time the computer starts. Choose a code that is <Bold>between $($MinimumPIN) and 20</Bold> characters long. It <Bold>must not be the same as your PC password</Bold> and should only include letters from the <Bold>English alphabet</Bold> (no special or language-specific characters)."
    }else {
        $passwordText = "BitLocker protects this device to prevent unauthorized access to your files and documents, even if someone has physical access. To help keep your data safe, you will need to enter a code each time the computer starts. Choose a code that is <Bold>between $($MinimumPIN) and 20</Bold> numbers long."
    }
    New-Variable -Name "passwordComplexityText" -Value $passwordText -Description "Shared var between runspaces" -Scope Global
    Write-Log "Minimum PIN length is '$($minLength)' and enhanced PIN enabled is '$($allowSpecial)'" -ComponentName "Main" -Path $LogLocation -Name $LogName

    #Define the XAML
    $Global:SyncHash = [hashtable]::Synchronized(@{})
    # Add functions and variables to hashtable
    $InitialSessionState = [initialsessionstate]::CreateDefault()
    Get-ChildItem function:/ | Where-Object Source -like "" | ForEach-Object {
        $functionDefinition = Get-Content "Function:\$($_.Name)"
        $sessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $_.Name, $functionDefinition 
        $InitialSessionState.Commands.Add($sessionStateFunction)
    }

    # Adding all variables will causes some variables not to be added, therefore only add the ones scoped with specified description
    Get-ChildItem variable:/ | ForEach-Object {
        $varDefinition = Get-Content "Variable:\$($_.Name)"
        $sessionStateVariable = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList ("$($_.Name)", $varDefinition,$null)
        if ($($_.Description) -eq "Shared var between runspaces") {
            $InitialSessionState.Variables.Add($sessionStateVariable)
        }
    }
    $SyncHash.Add("SharedSessionState",$InitialSessionState)
    $newRunspace = [runspacefactory]::CreateRunspace($SyncHash["SharedSessionState"])
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Open()

    $newRunspace.SessionStateProxy.SetVariable("syncHash", $SyncHash)

    # Load WPF assembly if necessary
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

    $psCmd = [PowerShell]::Create().AddScript({

            [xml]$xaml = 
            @"
<Window x:Class="BitLockerPIN.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:BitLockerPIN"
        mc:Ignorable="d"
        Title="MainWindow"  x:Name="mainwindow" Height="600" Width="450" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" ShowInTaskbar="False" Topmost="True" AllowsTransparency="True" WindowStyle="None" Background="Transparent">
    <Window.Resources>
        <Style x:Key="ModernRoundedButton" TargetType="Button">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#2F4151"/>
            <Setter Property="BorderBrush" Value="#2F4151"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Width" Value="80"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                            CornerRadius="6"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center"
                                          VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#588BB2"/> 
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#588BB2"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#CCC"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#AAA"/>
                                <Setter Property="Foreground" Value="#777"/>
                                <Setter Property="Cursor" Value="Arrow"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!--Scrollbar Thumbs-->
        <Style x:Key="ScrollThumbs" TargetType="{x:Type Thumb}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Thumb}">
                        <Grid x:Name="Grid">
                            <Rectangle HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Width="Auto" Height="Auto" Fill="Transparent" />
                            <Border x:Name="Rectangle1" CornerRadius="5" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Width="Auto" Height="Auto" Background="{TemplateBinding Background}" />
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="Tag" Value="Horizontal">
                                <Setter TargetName="Rectangle1" Property="Width" Value="Auto" />
                                <Setter TargetName="Rectangle1" Property="Height" Value="7" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!--ScrollBars-->
        <Style x:Key="{x:Type ScrollBar}" TargetType="{x:Type ScrollBar}">
            <Setter Property="Stylus.IsFlicksEnabled" Value="false" />
            <Setter Property="Foreground" Value="#8C8C8C" />
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="Width" Value="8" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                        <Grid x:Name="GridRoot" Width="8" Background="{TemplateBinding Background}">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="0.00001*" />
                            </Grid.RowDefinitions>
                            <Track x:Name="PART_Track" Grid.Row="0" IsDirectionReversed="true" Focusable="false">
                                <Track.Thumb>
                                    <Thumb x:Name="Thumb" Background="{TemplateBinding Foreground}" Style="{DynamicResource ScrollThumbs}" />
                                </Track.Thumb>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton x:Name="PageUp" Command="ScrollBar.PageDownCommand" Opacity="0" Focusable="false" />
                                </Track.IncreaseRepeatButton>
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton x:Name="PageDown" Command="ScrollBar.PageUpCommand" Opacity="0" Focusable="false" />
                                </Track.DecreaseRepeatButton>
                            </Track>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger SourceName="Thumb" Property="IsMouseOver" Value="true">
                                <Setter Value="{DynamicResource ButtonSelectBrush}" TargetName="Thumb" Property="Background" />
                            </Trigger>
                            <Trigger SourceName="Thumb" Property="IsDragging" Value="true">
                                <Setter Value="{DynamicResource DarkBrush}" TargetName="Thumb" Property="Background" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter TargetName="Thumb" Property="Visibility" Value="Collapsed" />
                            </Trigger>
                            <Trigger Property="Orientation" Value="Horizontal">
                                <Setter TargetName="GridRoot" Property="LayoutTransform">
                                    <Setter.Value>
                                        <RotateTransform Angle="-90" />
                                    </Setter.Value>
                                </Setter>
                                <Setter TargetName="PART_Track" Property="LayoutTransform">
                                    <Setter.Value>
                                        <RotateTransform Angle="-90" />
                                    </Setter.Value>
                                </Setter>
                                <Setter Property="Width" Value="Auto" />
                                <Setter Property="Height" Value="8" />
                                <Setter TargetName="Thumb" Property="Tag" Value="Horizontal" />
                                <Setter TargetName="PageDown" Property="Command" Value="ScrollBar.PageLeftCommand" />
                                <Setter TargetName="PageUp" Property="Command" Value="ScrollBar.PageRightCommand" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <DrawingImage x:Key="drw_info">
            <DrawingImage.Drawing>
                <DrawingGroup>
                    <DrawingGroup.ClipGeometry>
                        <RectangleGeometry Rect="0.0,0.0,32.0,32.0"/>
                    </DrawingGroup.ClipGeometry>
                    <DrawingGroup>
                        <GeometryDrawing>
                            <GeometryDrawing.Geometry>
                                <RectangleGeometry Rect="0.0,0.0,32.0,32.0"/>
                            </GeometryDrawing.Geometry>
                        </GeometryDrawing>
                    </DrawingGroup>
                    <GeometryDrawing Brush="#ff000000">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="M 24 11 C 24 6.03 19.971 2 15 2 C 12.515 2 10.271 3.014 8.644 4.644 l 2.833 2.833 C 12.382 6.572 13.619 6 15 6 c 2.76 0 5 2.239 5 5 c 0 2.761 -2.24 5 -5 5 c -1 0 -1 1 -1 1 l 0 7 h 4 V 19.477 C 21.493 18.24 24 14.917 24 11 z m -10 15 l 4 0 l 0 4 l -4 0 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <DrawingGroup/>
                </DrawingGroup>
            </DrawingImage.Drawing>
        </DrawingImage>
        <DrawingImage x:Key="drw_lock">
            <DrawingImage.Drawing>
                <DrawingGroup>
                    <DrawingGroup.ClipGeometry>
                        <RectangleGeometry Rect="0.0,0.0,20.0,20.0"/>
                    </DrawingGroup.ClipGeometry>
                    <GeometryDrawing Brush="#ff000000">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="M 4 8 V 6 a 6 6 0 1 1 12 0 v 2 h 1 a 2 2 0 0 1 2 2 v 8 a 2 2 0 0 1 -2 2 H 3 a 2 2 0 0 1 -2 -2 v -8 c 0 -1.1 0.9 -2 2 -2 h 1 z m 5 6.73 V 17 h 2 v -2.27 a 2 2 0 1 0 -2 0 z M 7 6 v 2 h 6 V 6 a 3 3 0 0 0 -6 0 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                </DrawingGroup>
            </DrawingImage.Drawing>
        </DrawingImage>
        <DrawingImage x:Key="ea_logo">
            <DrawingImage.Drawing>
                <DrawingGroup>
                    <DrawingGroup.ClipGeometry>
                        <RectangleGeometry Rect="0.0,0.0,64.000001,64.0"/>
                    </DrawingGroup.ClipGeometry>
                    <DrawingGroup>
                        <DrawingGroup>
                            <GeometryDrawing Brush="#ff00baff">
                                <GeometryDrawing.Geometry>
                                    <PathGeometry Figures="M 31.2692 62.27 C 18.0245 62.2864 6.32008 53.6572 2.41918 41 c 1.72437 -0.122488 3.37221 -0.760025 4.73 -1.83 c 3.4 -2.68 7 -3.31 11.8 2.85 c 0.414648 0.549572 0.865658 1.07074 1.35 1.56 c 6.10815 6.18587 16.0745 6.24855 22.26 0.14 l 0.15 -0.14 c 6.19004 -6.27345 6.19004 -16.3565 0 -22.63 c -6.11149 -6.20587 -16.101 -6.26861 -22.29 -0.14 l -0.15 0.15 c -0.484814 0.491694 -0.939001 1.01267 -1.36 1.56 c -4.82 6.16 -8.4 5.53 -11.8 2.85 c -1.34487 -1.06578 -2.97871 -1.7033 -4.69 -1.83 C 9.49674 -0.224589 40.6463 -5.74302 55.4582 14.1414 C 70.2701 34.0258 56.0653 62.2919 31.2692 62.27 Z m 0 -36 c 5.3458 0 8.02202 6.46283 4.24242 10.2424 c -3.77959 3.77959 -10.2424 1.10338 -10.2424 -4.24242 c 0 -3.31371 2.68629 -6 6 -6 z"/>
                                </GeometryDrawing.Geometry>
                            </GeometryDrawing>
                            <GeometryDrawing Brush="#ff008eff">
                                <GeometryDrawing.Geometry>
                                    <PathGeometry Figures="M 31.2692 55.27 C 20.7515 55.2806 11.5568 48.179 8.90918 38 c 2.89 -1.52 6 -1.09 10 4 c 0.414648 0.549572 0.865658 1.07074 1.35 1.56 c 6.11159 6.18702 16.0811 6.24969 22.27 0.14 l 0.14 -0.14 c 6.19004 -6.27345 6.19004 -16.3565 0 -22.63 c -6.11302 -6.17637 -16.0762 -6.22563 -22.25 -0.11 l -0.15 0.14 c -0.484814 0.491694 -0.939001 1.01267 -1.36 1.56 c -4 5.11 -7.14 5.54 -10 4 C 13.6222 7.88003 37.7178 2.82903 49.5207 18.0049 C 61.3236 33.1809 50.4957 55.2913 31.2692 55.27 Z"/>
                                </GeometryDrawing.Geometry>
                            </GeometryDrawing>
                        </DrawingGroup>
                    </DrawingGroup>
                </DrawingGroup>
            </DrawingImage.Drawing>
        </DrawingImage>
        <DrawingImage x:Key="drw_locktext">
            <DrawingImage.Drawing>
                <DrawingGroup>
                    <DrawingGroup.ClipGeometry>
                        <RectangleGeometry Rect="0.0,0.0,426.39957,518.73163"/>
                    </DrawingGroup.ClipGeometry>
                    <GeometryDrawing Brush="#ff2f4050">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 202.488 0.325544 c 3.766 -0.3095 7.561 -0.3191 11.339 -0.3254 c 25.109 -0.0418 48.433 9.0195 66.259 26.9259 c 15.21 15.2789 25.395 36.0922 27.523 57.6022 c 1.261 12.741 0.411 28.25 0.392 41.272 c 13.506 -0.131 25.938 -0.209 36.179 9.964 c 11.994 11.916 10.196 27.572 10.136 43.056 c 8.305 0.309 15.137 2.384 21.186 8.339 c 4.186 4.121 6.597 8.577 8.275 14.158 c 2.958 9.844 4.218 20.74 6.098 30.88 l 10.829 58.518 l 16.252 87.295 l 5.669 30.542 c 1.25 6.694 2.784 13.51 3.394 20.293 c 0.514 5.708 0.356 11.494 0.363 17.22 l -0.003 21.347 c -0.024 15.017 -0.418 28.072 -11.762 39.314 c -5.701 5.65 -12.523 9.59 -20.46 11.111 c -6.602 1.266 -13.801 0.744 -20.513 0.74 l -31.31 -0.031 l -117.79 0.01 H 90.975 l -36.993 0.032 c -7.237 0.004 -15.11 0.593 -22.238 -0.646 c -7.156 -1.244 -14.134 -4.909 -19.382 -9.904 c -9.278 -8.831 -11.986 -19.859 -12.256 -32.256 c -0.186 -8.533 -0.071 -17.095 -0.069 -25.631 c 0.002 -7.476 -0.261 -15.072 0.601 -22.507 c 1.629 -14.056 5.037 -28.282 7.62 -42.205 l 15.226 -82.069 l 12.381 -66.977 c 2.11 -11.432 3.589 -23.687 6.775 -34.852 c 1.601 -5.612 3.747 -9.93 7.84 -14.139 c 6.242 -6.418 13.583 -8.42 22.295 -8.477 c -0.077 -15.569 -2.007 -30.915 10.04 -42.91 c 10.413 -10.369 22.486 -10.325 36.23 -10.213 c 0.337 -35.586 -3.525 -61.19 20.95 -91.8451 c 15.472 -19.3794 38.022 -30.8585 62.493 -33.6316 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ffb2b8bf">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 354.325 194.377 c 3.044 0.474 6.052 1.1 8.6 2.935 c 2.195 1.582 4.168 4.24 4.963 6.833 c 1.604 5.227 2.35 10.874 3.342 16.245 l 5.054 27.417 l 16.842 90.988 l 13.551 72.431 c -2.361 -1.227 -4.761 -2.363 -7.285 -3.22 c -4.465 -1.517 -9.034 -2.079 -13.728 -2.268 c -21.015 -0.848 -43.109 -0.085 -64.262 -0.076 l -126.594 0.084 l -104.732 0.146 l -33.347 0.078 c -6.719 0.017 -13.514 -0.175 -20.218 0.245 c -5.989 0.374 -11.05 2.167 -16.395 4.805 c 7.65 -43.656 16.152 -87.195 24.269 -130.768 l 8.862 -47.662 c 1.591 -8.571 2.858 -17.402 4.93 -25.862 c 0.676 -2.762 1.648 -5.475 3.522 -7.668 c 2.835 -3.315 6.449 -4.004 10.55 -4.366 c 2.444 2.33 -1.326 100.926 1.838 115.237 c 1.325 6.206 4.348 11.923 8.732 16.512 c 6.855 7.211 15.821 10.613 25.648 10.861 c 15.636 0.394 31.379 -0.011 47.024 -0.012 l 96.147 -0.011 l 45.733 0.014 c 8.508 0.025 17.764 0.805 26.172 -0.464 c 7.462 -1.125 14.05 -4.332 19.475 -9.591 c 4.591 -4.451 8.439 -10.387 9.905 -16.665 c 2.467 -10.563 1.302 -29.947 1.304 -41.434 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ffb2b8bf">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 37.305 422.615 c 0.033 -0.003 0.065 -0.007 0.098 -0.009 c 8.421 -0.368 16.936 -0.114 25.369 -0.104 l 43.623 0.008 l 135.946 -0.073 l 97.863 -0.083 l 30.963 -0.011 c 5.935 0 12.062 -0.256 17.976 0.241 c 4.527 0.38 9.69 2.62 13.099 5.654 c 4.911 4.371 7.283 10.004 7.663 16.523 c 0.429 7.367 0.702 34.583 -0.806 40.625 c -0.876 3.396 -2.577 6.524 -4.951 9.106 c -4.561 5.026 -10.388 7.351 -17.098 7.649 c -25.913 0.485 -51.896 0.096 -77.818 0.097 l -144.506 0.029 l -84.472 0.038 l -26.504 0.014 c -5.717 0.001 -11.538 0.195 -17.238 -0.258 c -3.956 -0.314 -7.749 -1.87 -10.913 -4.257 c -5.045 -3.807 -8.178 -9.4 -8.915 -15.674 c -0.72 -6.128 -0.731 -36.397 0.394 -41.873 c 0.709 -3.611 2.3 -6.99 4.63 -9.837 c 4.037 -4.869 9.452 -7.149 15.597 -7.805 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff2f4050">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 51.188 456.955 c 5.989 -0.316 23.35 -0.995 28.206 0.909 c 1.863 0.731 3.854 2.057 4.596 4.004 c 0.674 1.768 0.118 3.54 -0.667 5.165 c -1.356 2.804 -2.772 3.969 -5.713 4.957 c -6.096 0.067 -23.576 0.869 -28.475 -1.212 c -1.767 -0.751 -3.261 -1.954 -3.959 -3.785 c -0.703 -1.842 -0.372 -4.203 0.498 -5.937 c 1.141 -2.277 3.182 -3.371 5.514 -4.101 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 379.387 439.871 l 0.616 -0.034 c 1.394 2.252 0.515 18.632 0.452 22.254 l -23.043 -0.023 c -0.78 -0.053 -0.74 0.058 -1.387 -0.489 c -1.127 -5.242 -0.312 -13.975 -0.266 -19.626 c 7.867 -0.798 15.744 -1.492 23.628 -2.082 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 379.952 465.502 l 0.162 0.063 c 1.213 2.016 0.425 18.621 0.398 22.069 c -7.057 -0.965 -14.225 -1.495 -21.306 -2.315 l -3.433 -0.335 l -0.033 -19.494 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 347.188 442.701 c 1.157 -0.159 2.127 -0.341 3.278 -0.016 c 1.117 3.567 0.307 14.979 0.278 19.353 l -16.784 0.009 c -1.014 0.024 -1.654 0.105 -2.568 -0.404 c -0.849 -5.04 -0.098 -12.156 -0.225 -17.515 c 5.346 -0.413 10.686 -0.888 16.021 -1.427 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 337.077 465.452 c 4.555 0.057 9.111 0.052 13.667 -0.013 l -0.002 18.981 l -3.725 -0.426 l -15.755 -1.709 c 0.033 -5.601 0.03 -11.203 -0.01 -16.805 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#fffcfdfd">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 209.977 31.5348 c 0.855 -0.0582 1.712 -0.0865 2.57 -0.085 c 17.359 -0.0096 33.991 6.9684 46.147 19.3607 c 9.338 9.4187 16.222 22.1947 17.726 35.4747 c 0.62 5.476 0.426 11.055 0.422 16.561 l -0.041 22.975 l -59.055 0.011 l -67.222 0.001 l -0.033 -23.711 c -0.006 -19.879 1.217 -32.918 14.834 -48.859 c 11.442 -13.394 27.322 -20.3632 44.652 -21.7284 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 155.945 212.602 c 6.912 -0.637 13.634 2.487 17.603 8.182 c 3.968 5.695 4.573 13.083 1.583 19.347 c -2.991 6.264 -9.115 10.44 -16.039 10.935 c -10.553 0.754 -19.746 -7.13 -20.609 -17.675 c -0.863 -10.545 6.926 -19.818 17.462 -20.789 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 210.857 158.403 c 10.553 -1.32 20.183 6.152 21.525 16.703 c 1.342 10.55 -6.11 20.195 -16.657 21.56 c -10.579 1.368 -20.26 -6.111 -21.606 -16.693 c -1.346 -10.582 6.153 -20.247 16.738 -21.57 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 265.179 159.126 c 10.42 -2.128 20.594 4.589 22.729 15.008 c 2.135 10.418 -4.576 20.596 -14.993 22.738 c -10.427 2.144 -20.616 -4.575 -22.753 -15.003 c -2.137 -10.428 4.588 -20.612 15.017 -22.743 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 156.542 158.969 c 10.596 -0.634 19.708 7.423 20.377 18.017 c 0.668 10.594 -7.359 19.732 -17.951 20.435 c -10.641 0.707 -19.831 -7.366 -20.502 -18.009 c -0.672 -10.643 7.431 -19.806 18.076 -20.443 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 211.159 267.867 c 10.574 -1.162 20.086 6.47 21.244 17.044 c 1.157 10.574 -6.478 20.083 -17.053 21.236 c -10.568 1.153 -20.071 -6.477 -21.228 -17.045 c -1.157 -10.568 6.47 -20.074 17.037 -21.235 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 265.998 213.352 c 10.392 -1.657 20.184 5.35 21.971 15.72 c 1.786 10.371 -5.098 20.25 -15.445 22.166 c -6.81 1.261 -13.771 -1.239 -18.223 -6.543 c -4.453 -5.304 -5.708 -12.592 -3.287 -19.081 c 2.421 -6.488 8.145 -11.172 14.984 -12.262 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 265.952 268.163 c 10.467 -1.694 20.327 5.411 22.033 15.875 c 1.706 10.464 -5.386 20.333 -15.848 22.051 c -10.48 1.722 -20.369 -5.385 -22.078 -15.867 c -1.709 -10.482 5.41 -20.362 15.893 -22.059 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 210.884 213.083 c 10.425 -1.323 19.973 5.989 21.411 16.398 c 1.439 10.409 -5.767 20.038 -16.16 21.591 c -6.826 1.021 -13.673 -1.705 -17.93 -7.138 c -4.257 -5.432 -5.266 -12.733 -2.642 -19.117 c 2.623 -6.383 8.475 -10.864 15.321 -11.734 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                    <GeometryDrawing Brush="#ff598db4">
                        <GeometryDrawing.Geometry>
                            <PathGeometry Figures="m 156.124 268.24 c 10.467 -0.888 19.698 6.817 20.695 17.274 c 0.996 10.457 -6.613 19.767 -17.059 20.871 c -6.854 0.725 -13.569 -2.287 -17.586 -7.888 c -4.016 -5.602 -4.715 -12.929 -1.829 -19.188 c 2.886 -6.259 8.912 -10.486 15.779 -11.069 z" FillRule="Nonzero"/>
                        </GeometryDrawing.Geometry>
                    </GeometryDrawing>
                </DrawingGroup>
            </DrawingImage.Drawing>
        </DrawingImage>
    </Window.Resources>
    <Border CornerRadius="5" BorderBrush="Black" BorderThickness="1" Padding="0" Background="White">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Focusable="False" Foreground="Black">
            <Grid>
                <Grid x:Name="grd_mainwindow" Margin="10">
                    <Grid.Effect>
                        <BlurEffect x:Name="maingridBlur" Radius="0" KernelType="Gaussian"/>
                    </Grid.Effect>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="35"/>
                        <RowDefinition Height="200"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Grid.Row="0" Margin="-10">
                        <Image x:Name="img_logo" Margin="0" Grid.Column="0" HorizontalAlignment="Stretch" VerticalAlignment="Bottom"/>
                        <Image x:Name="img_ealogo" Source="{StaticResource ea_logo}" Margin="10" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Stretch="Uniform"/>
                    </Grid>
                    <Grid VerticalAlignment="Center" Grid.Row="1">
                        <StackPanel Orientation="Vertical" x:Name="tbx_custommessageheader">
                            <Image Source="{StaticResource drw_locktext}" HorizontalAlignment="Center" Height="60" VerticalAlignment="Top" Stretch="Uniform"/>
                            <Label x:Name="lbl_mainheadline" HorizontalAlignment="Center" FontFamily="Segoe UI" FontWeight="Bold" FontSize="20">
                                <TextBlock HorizontalAlignment="Center" FontFamily="Segoe UI" FontWeight="Bold" FontSize="20">
                                    <Run Text="BitLocker " Foreground="#2F4151"/>
                                    <Run Text="PIN" Foreground="#588BB2"/>
                                </TextBlock>
                            </Label>
                        </StackPanel>
                    </Grid>
                    <Grid x:Name="grd_password" VerticalAlignment="Top" Grid.Row="2" Opacity="1">
                        <StackPanel x:Name="stc_pwdobjects" Orientation="Vertical" Margin="100,20,100,0">
                            <TextBlock x:Name="tbx_custommessagebody" Margin="-90,-20,-90,20" HorizontalAlignment="Center" FontFamily="Segoe UI" FontSize="12" TextWrapping="Wrap" TextAlignment="Center">
                            $($passwordComplexityText)
                            </TextBlock>
                            <!-- Password 1 -->
                            <Label Content="PIN" FontFamily="Segoe UI" FontWeight="SemiBold" FontSize="13"/>
                            <Border CornerRadius="6" BorderBrush="#CCC" BorderThickness="1" Background="#FFF" Margin="0,2,0,0">
                                <PasswordBox x:Name="pbx_passwordfirst"
                                             Background="Transparent"
                                             BorderThickness="0"
                                             Padding="8,4"
                                             FontSize="13"
                                             FontFamily="Segoe UI"
                                             TabIndex="1"/>
                            </Border>

                            <!-- Password 2 -->
                            <Label Content="Confirm PIN" FontFamily="Segoe UI" FontWeight="SemiBold" FontSize="13" Margin="0,10,0,0"/>
                            <Border CornerRadius="6" BorderBrush="#CCC" BorderThickness="1" Background="#FFF" Margin="0,2,0,0">
                                <PasswordBox x:Name="pbx_passwordsecond"
                                             Background="Transparent"
                                             BorderThickness="0"
                                             Padding="8,4"
                                             FontSize="13"
                                             FontFamily="Segoe UI"
                                             TabIndex="2"/>
                            </Border>

                            <!-- Error Text -->
                            <StackPanel x:Name="stk_passwordnotaccepted" Orientation="Horizontal" Margin="0,5,0,10" Height="20" Visibility="Hidden">
                                <TextBlock x:Name="tbx_passwordnotaccepted"
                                            Foreground="Red"
                                            FontSize="12"
                                            Text="Error. Please try again."
                                            VerticalAlignment="Center"/>
                            </StackPanel>

                            <!-- Button -->
                            <Button x:Name="btn_accept"
                                    Content="OK"
                                    Style="{StaticResource ModernRoundedButton}"
                                    IsDefault="True"
                                    IsEnabled="False"
                                    TabIndex="3"/>
                        </StackPanel>
                    </Grid>
                    <StackPanel x:Name="stck_error" Orientation="Vertical" Grid.Row="2" HorizontalAlignment="Center" Margin="0,10,0,0" Opacity="0" IsEnabled="False" IsHitTestVisible="False">
                        <StackPanel>
                            <TextBlock x:Name="tbx_savefailed"  HorizontalAlignment="Center" FontFamily="Segoe UI" FontSize="13" TextWrapping="Wrap" TextAlignment="Left">
                                Oops! Something went wrong while trying to save your code. Please try again later or contact the service desk for assistance.<LineBreak />
                            </TextBlock>
                        </StackPanel>
                        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Width="Auto" Height="100">
                            <TextBox x:Name="tbx_stacktrace" IsReadOnly="True" Foreground="#000" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" FontFamily="Consolas" FontSize="13" TextWrapping="Wrap" Margin="0,5,0,0"  IsTabStop="False"></TextBox>
                        </ScrollViewer>
                        <Button x:Name="btn_closeerrorwindow" HorizontalAlignment="Center" IsCancel="True" Padding="3" Margin="10" Style="{StaticResource ModernRoundedButton}">
                            Close
                        </Button>
                    </StackPanel>
                </Grid>
                <Grid x:Name="grd_infobox" >
                    <Border x:Name="InfoBox" CornerRadius="5" BorderBrush="Black" BorderThickness="1" Padding="0" Background="GhostWhite" Width="425" Opacity="0" Height="350">
                        <Border.Style>
                            <Style TargetType="Border">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding IsChecked, ElementName=tgl_info}" Value="True">
                                        <Setter Property="IsHitTestVisible" Value="True"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding IsChecked, ElementName=tgl_info}" Value="False">
                                        <Setter Property="IsHitTestVisible" Value="False"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </Border.Style>
                        <ScrollViewer VerticalScrollBarVisibility="Auto" Focusable="False" Foreground="Black">
                            <Grid Margin="10" Opacity="1">
                                <StackPanel>
                                    <TextBlock x:Name="tbx_infopba" HorizontalAlignment="Center" FontFamily="Segoe UI" FontSize="12" TextWrapping="Wrap" TextAlignment="Left" Margin="10">
                                
                                    <Bold>Understanding BitLocker</Bold><LineBreak />
                                    BitLocker is a full-disk encryption feature designed to protect data by preventing unauthorized access to your drive. It secures the contents of your device by encrypting all files, ensuring that only authorized users can access the information - even if the physical device falls into the wrong hands.<LineBreak /><LineBreak />

                                    <Bold>The Importance of Setting a BitLocker PIN</Bold><LineBreak />
                                    To strengthen security, BitLocker allows you to configure a startup PIN. This serves as an added authentication layer before the operating system loads, significantly reducing the risk of data breaches in the event your device is lost or stolen.<LineBreak /><LineBreak />
                                    Once the BitLocker PIN has been configured, you will be prompted to enter it during each system startup:
                                    </TextBlock>
                                    <Image x:Name="img_pba" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="10"/>
                                    <TextBlock x:Name="tbx_infotext" HorizontalAlignment="Center" FontFamily="Segoe UI" FontSize="12" TextWrapping="Wrap" TextAlignment="Left" Margin="10">

                                    <Bold>PIN Configuration Guidelines</Bold><LineBreak />
                                    When creating your BitLocker PIN, please ensure it meets the following requirements:<LineBreak /><LineBreak />
                                    &#160;• Length must be between 4 and 20 characters, and may include letters, numbers, and supported special symbols.<LineBreak />
                                    &#160;• The PIN must differ from your Windows user account password.<LineBreak />
                                    &#160;• Avoid including any personal identifiers, such as your name or date of birth.<LineBreak />
                                    &#160;• Use only standard English alphabet characters (no regional or accented letters).<LineBreak />
                                    &#160;• Do not include sequences of three or more consecutive letters or numbers (e.g., abc, 123).<LineBreak /><LineBreak /><LineBreak />

                                    <Bold>Additional Resources</Bold><LineBreak />
                                    For further guidance contact your IT department
                                    </TextBlock>
                                </StackPanel>
                            </Grid>
                        </ScrollViewer>
                    </Border>
                </Grid>
                <ToggleButton x:Name="tgl_info" Width="25" Height="25" HorizontalAlignment="Right" VerticalAlignment="Bottom" Background="{x:Null}" BorderThickness="0" IsTabStop="True" TabIndex="4" Margin="5" IsHitTestVisible="True">
                    <ToggleButton.Triggers>
                        <EventTrigger RoutedEvent="ToggleButton.Unchecked">
                            <BeginStoryboard>
                                <Storyboard x:Name="ShowInfoGrid">
                                    <DoubleAnimation BeginTime="00:00:00" Duration="0:0:0.2" Storyboard.TargetName="InfoBox" Storyboard.TargetProperty="Opacity" To="0"/>
                                    <DoubleAnimation BeginTime="00:00:00" Duration="0:0:0.2" Storyboard.TargetName="maingridBlur" Storyboard.TargetProperty="Radius" To="0"/>
                                </Storyboard>
                            </BeginStoryboard>
                        </EventTrigger>
                        <EventTrigger RoutedEvent="ToggleButton.Checked">
                            <BeginStoryboard>
                                <Storyboard x:Name="HideInfoGrid">
                                    <DoubleAnimation BeginTime="00:00:00" Duration="0:0:0.2" Storyboard.TargetName="InfoBox" Storyboard.TargetProperty="Opacity" To="1"/>
                                    <DoubleAnimation BeginTime="00:00:00" Duration="0:0:0.2" Storyboard.TargetName="maingridBlur" Storyboard.TargetProperty="Radius" To="10"/>
                                </Storyboard>
                            </BeginStoryboard>
                        </EventTrigger>
                    </ToggleButton.Triggers>
                    <Image Source="{StaticResource drw_info}" Stretch="Uniform"/>
                </ToggleButton>
            </Grid>
        </ScrollViewer>
    </Border>
</Window>

"@

            # Remove XML attributes that break a couple things.
            #   Without this, you must manually remove the attributes
            #   after pasting from Visual Studio. If more attributes
            #   need to be removed automatically, add them below.
            $AttributesToRemove = @(
                'x:Class',
                'mc:Ignorable'
            )

            foreach ($Attrib in $AttributesToRemove) {
                if ( $xaml.Window.GetAttribute($Attrib) ) {
                    $xaml.Window.RemoveAttribute($Attrib)
                }
            }
    
            $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    
            $syncHash.Window = [Windows.Markup.XamlReader]::Load( $reader )

            [xml]$XAML = $xaml
            $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % {
                #Find all of the form types and add them as members to the synchash
                $syncHash.Add($_.Name, $syncHash.Window.FindName($_.Name) )

            }

            Set-WindowWorkArea -MiddleOrCorner "Corner"

            #Add picture
            Add-Type -AssemblyName PresentationFramework

            if ($UseBanner) {
                $pngFilePath = "$env:ProgramData\EndpointAdmin\InstallationBanner\AppDeployToolkitBanner.png"
                if (Test-Path $pngFilePath) {
                    $imageBytes = [System.IO.File]::ReadAllBytes($pngFilePath)
                    $memoryStream = New-Object System.IO.MemoryStream
                    $memoryStream.Write($imageBytes, 0, $imageBytes.Length)
                    $memoryStream.Position = 0

                    # Convert to BitmapImage for WPF
                    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bitmap.BeginInit()
                    $bitmap.StreamSource = $memoryStream
                    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bitmap.EndInit()
                    $bitmap.Freeze()  # Optional: Make it cross-thread accessible
                    $syncHash.img_logo.Source = $bitmap
                    
                    # Hide Endpoint Admin logo
                    $syncHash.img_ealogo.Visibility = "Hidden"

                }
            }

            $Base64Image = "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"
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit()
            $bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($Base64Image)
            $bitmap.EndInit()
            $syncHash.img_pba.Source = $bitmap

            $Script:JobCleanup = [hashtable]::Synchronized(@{})
            $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

            $jobCleanup.Flag = $True
            $newRunspace = [runspacefactory]::CreateRunspace()
            $newRunspace.ApartmentState = "STA"
            $newRunspace.ThreadOptions = "ReuseThread"          
            $newRunspace.Open()        
            $newRunspace.SessionStateProxy.SetVariable("jobCleanup", $jobCleanup)     
            $newRunspace.SessionStateProxy.SetVariable("jobs", $jobs) 
            $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({

                    #Routine to handle completed runspaces
                    Do {    
                        Foreach ($runspace in $jobs) {            
                            If ($runspace.Runspace.isCompleted) {
                                [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                                $runspace.powershell.dispose()
                                $runspace.Runspace = $null
                                $runspace.powershell = $null               
                            } 
                        }
                        #Clean out unused runspace jobs
                        $temphash = $jobs.clone()
                        $temphash | Where {
                            $_.runspace -eq $Null
                        } | ForEach {
                            $jobs.remove($_)
                        }        
                        Start-Sleep -Seconds 1     
                    } while ($jobCleanup.Flag)
                })
            $jobCleanup.PowerShell.Runspace = $newRunspace
            $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  

            $syncHash.btn_accept.Add_Click({

                    $newRunspace = [runspacefactory]::CreateRunspace($SyncHash["SharedSessionState"])
                    $newRunspace.ApartmentState = "STA"
                    $newRunspace.ThreadOptions = "ReuseThread"          
                    $newRunspace.Open()
                    $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash) 
                    $PowerShell = [PowerShell]::Create().AddScript({

                            try {
                                Update-Window -Control "stc_pwdobjects" -Property "IsEnabled" -Value $False
                                if ($DryRun) {
                                    if ($DryRunEnforceError) {
                                        Write-Log -LogOutput ("DryRunEnforceError was specified") -ComponentName "Main" -Path $LogLocation -Name $LogName
                                        throw 1337
                                    }
                                    Write-Log -LogOutput ("DryRun was specified") -ComponentName "Main" -Path $LogLocation -Name $LogName
                                    Update-Window -Property "Close"

                                }else {

                                    Write-Log -LogOutput ("Setting bitlocker PIN..") -ComponentName "Main" -Path $LogLocation -Name $LogName
                                    Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -Pin $(ConvertTo-SecureString $syncHash.pbx_passwordfirst.Password -AsPlainText -Force) -TpmAndPinProtector -ErrorAction Stop
                                    [string]$DateTimeNow = Get-Date -Format yyyy-MM-ddTHH:mm:ssZ
                                    Set-RegistryKey -Key $RegistryKeyPath -Name "PINLastSet" -Value $DateTimeNow -Type "String"
                                    if (Test-BitLockerTpmAndPINProtectorSet) {

                                        Write-Log -LogOutput ("Removing Tpm protector key from system drive as both Tpm and TpmPin is present, otherwise Pre-Authentication would not occur") -ComponentName "Main" -Path $LogLocation -Name $LogName
                                        $TpmProtector = $(Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector | Where-Object { $_.KeyProtectorType -eq 'Tpm' }
                                        Remove-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $TpmProtector.KeyProtectorId
                                    }
                                    Update-Window -Property "Close"
                                }
                            }
                            catch {
                                Write-Log -LogOutput ("Failed with errorcode: {0}" -f $_) -ComponentName "Main" -Path $LogLocation -Name $LogName
                                Update-Window -Control "tbx_stacktrace" -Property "Text" -Value $_
                                Update-Window -Control "stck_error" -Property "Opacity" -Value "1"
                                Update-Window -Control "grd_password" -Property "Opacity" -Value "0"
                                Update-Window -Control "stck_error" -Property "IsEnabled" -Value $True
                                Update-Window -Control "grd_password" -Property "IsEnabled" -Value $False
                                Update-Window -Control "stck_error" -Property "IsHitTestVisible" -Value $True
                                
                                #Update-Window -Control "tbx_savefailed" -Property "Text" -Value "Failed to set PIN"
                                $syncHash.btn_copyerrormsg.Add_Click({
                                    # Set-Clipboard -Text $syncHash.tbx_stacktrace.Text.Value
                                })
                            }
                        })
                    $PowerShell.Runspace = $newRunspace
                    [void]$Jobs.Add((
                            [pscustomobject]@{
                                PowerShell = $PowerShell
                                Runspace   = $PowerShell.BeginInvoke()
                            }
                        ))
                })

            # Used to set button enablement
            $syncHash.pbx_passwordfirst.Add_PasswordChanged({
                Set-PinAcceptedState
            })
            $syncHash.pbx_passwordsecond.Add_PasswordChanged({
                Set-PinAcceptedState
            })

            $syncHash.btn_closeerrorwindow.Add_Click({
                Write-Log -LogOutput ("Clicking error button") -ComponentName "Clicking_Error" -Path $LogLocation -Name $LogName
                $Global:CancelUserClosingWindow = $false
            })

            $syncHash.Window.Add_Closing({

                    #Cancel if the user tries to close the window
                    Write-Log -LogOutput ("Window closing, cancellation '{0}'" -f $CancelUserClosingWindow) -ComponentName "Closing" -Path $LogLocation -Name $LogName
                    $_.Cancel = $CancelUserClosingWindow

                    if (!($CancelUserClosingWindow)) {
                        #Stop all runspaces
                        Update-Window -Property "Close"
                    }
                })

            #$syncHash.Window.Activate()
            $syncHash.Window.ShowDialog() | Out-Null
            $syncHash.Error = $Error
        })

    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()

    #Set registry key to enable not to run multiple instances of app
    Set-RegistryKey -Key $RegistryKeyPath -Name "PID" -Value $pid -Type "STRING"

    #Create an application context for it to all run within.
    $AppContext = New-Object System.Windows.Forms.ApplicationContext
    [void][System.Windows.Forms.Application]::Run($AppContext)
}
function Test-IfAnyLoggedOnUsers {

    #Using Query session
    $LoggedonUsers = ((quser) -replace '\s{2,}', ',') -replace '>','' | ConvertFrom-Csv
    if ($LoggedonUsers) {
        foreach ($LoggedOnUser in $LoggedonUsers) {
            Write-Log -LogOutput ("Found logged on user '{0}'" -f $LoggedOnUser.Username) -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        }
    }else{
        Write-Log -LogOutput ("No logged on users found..") -ComponentName $($MyInvocation.MyCommand) -Path $LogLocation -Name $LogName
        return $false
    }
    return $true
}
function Test-BitLockerShouldBeResumed {

    # Do not do anything if testing
    if($DryRun) { return $false }

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

    # Always show prompt when testing
    if($DryRun) { return $false }

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
function Test-BitLockerTpmAndPINProtectorSet {

    # Always show prompt when testing
    if($DryRun) { return $false }

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

# Resume bitlocker if falsely suspended
if (Test-BitLockerShouldBeResumed) {
    Resume-BitLocker -MountPoint  $env:SystemDrive | Out-Null
}

# Remove Tpm key protector if present with TpmPin
if (Test-BitLockerTpmAndPINProtectorSet) {

    try {
        Write-Log -LogOutput ("Removing Tpm protector key from system drive as both Tpm and TpmPin is present, otherwise Pre-Authentication would not occur") -ComponentName "Main" -Path $LogLocation -Name $LogName
        $TpmProtector = $(Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector | Where-Object { $_.KeyProtectorType -eq 'Tpm' }
        Remove-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $TpmProtector.KeyProtectorId
    }
    catch {
        Write-Log -LogOutput ("Error: $_") -ComponentName "Main" -Path $LogLocation -Name $LogName
    }
}

# Set BitLocker PIN if not already set
if (!(Test-BitLockerPINProtectorSet)) {

    #Check if app is already running
    $RunningPID = Get-RegistryKey -Key $RegistryKeyPath -Name "PID"
    if( -not [string]::IsNullOrEmpty($RunningPID) ) {
        if (Get-Process -PID $RunningPID -ErrorAction SilentlyContinue | Where-Object{$_.ProcessName -eq "powershell"}) {
            Write-Log -LogOutput ("Application is already running, exiting..") -ComponentName "Main" -Path $LogLocation -Name $LogName
            Exit 1
        }
    }

    #Re-run to show prompt to interactive user running as system
    if (($RunningAsSystem) -and !($InvokeToInteractiveUser)) {
        Write-Log "*********************************** SCRIPT START ***********************************" -ComponentName "Main" -Path $LogLocation -Name $LogName
        # Testing if any logged on users, otherwise the system process would be presented to a hidden desktop.
        if (Test-IfAnyLoggedOnUsers) {
            Write-Log -LogOutput "Running as SYSTEM but not presented to user, re-running launching to interactive user, use banner is set to '$($UseCustomPSADTBanner)'.." -ComponentName "Main" -Path $LogLocation -Name $LogName
            Start-RunAsSystemPresentToInteractiveUser -Path $($MyInvocation.MyCommand.Path) -UseCustomPSADTBanner $UseCustomPSADTBanner -CancelClosingWindow $CancelClosingWindow
            break
        }else {
            Exit 2
        }
    }

    # Show to interactive user 
    if ($InvokeToInteractiveUser) {
        Write-Log -LogOutput ("Showing prompt for enduser.." -f $ExitCode) -ComponentName "Main" -Path $LogLocation -Name $LogName
        Show-NotificationMessage
    }
}
