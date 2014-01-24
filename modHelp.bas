Attribute VB_Name = "modHelp"
Option Explicit

Public Function GetHelpText$(sNodeName$)
    Dim sName$, sHelp$
    sName = sNodeName
    If sName = "System" Then
        GetHelpText = "Click an item to see a description of its function and origin."
        Exit Function
    End If
    If sName = "Users" Then
        GetHelpText = "These are the startup items for other users."
        Exit Function
    End If
    If sName = "Hardware" Then
        GetHelpText = "These are the startup items for other hardware configurations."
        Exit Function
    End If
    '=== save section names ===
    If sName = "Files" Then
        GetHelpText = "This group contains sections that cover files that get loaded just by being there."
        Exit Function
    End If
    If sName = "MSIE" Then
        GetHelpText = "This group contains sections that cover Internet Explorer in some way."
        Exit Function
    End If
    If sName = "Hijack" Then
        GetHelpText = "This group contains sections that cover system components that can be hijacked in some way, though not in a way to launch files."
        Exit Function
    End If
    If sName = "Disabled" Then
        GetHelpText = "This group contains sections that cover disabled or blocked items, either for protection or for system optimization."
        Exit Function
    End If
    If sName = "Registry" Then
        GetHelpText = "This group contains all sections that originate from the Registry."
        Exit Function
    End If
    
    sName = GetSectionFromKey(sName)
    Select Case sName
        Case "RunningProcesses"
            sHelp = "Processes that are currently running in " & _
                    "memory, as well as all dll libraries loaded " & _
                    "by each process." & _
                    vbCrLf & "[All Windows versions]"
        Case "AutoStartFolders", "AutoStartFoldersStartup", "AutoStartFoldersUser Startup", "AutoStartFoldersCommon Startup", "AutoStartFoldersUser Common Startup", "AutoStartFoldersIOSUBSYS folder", "AutoStartFoldersVMM32 folder", "Windows Vista common Startup", "Windows Vista roaming profile Startup", "Windows Vista roaming profile Startup 2"
            sHelp = "Special folders that contain files that are " & _
                    "started when users logon." & _
                    vbCrLf & "[All Windows versions]"
        Case "TaskScheduler", "TaskSchedulerJobs", "TaskSchedulerJobsSystem"
            sHelp = "The folder that contains all jobs the " & _
                    "Windows Task Scheduler can be setup to run " & _
                    "periodically." & _
                    vbCrLf & "[All Windows versions]"
        Case "IniFiles", "IniFilessystem.ini", "IniFileswin.ini"
            sHelp = "Some system settings are stored in *.ini " & _
                    "files instead of in the Registry. Some of " & _
                    "these settings can be used to start a " & _
                    "program on system startup." & _
                    vbCrLf & "[Windows 9x/ME]"
        Case "IniMapping"
            sHelp = "Windows NT4, 2000 and XP (and newer) map " & _
                    "several *.ini files to the Registry. Some " & _
                    "of the settings can be used to start a " & _
                    "program on system startup." & _
                    vbCrLf & "[Windows NT/2k/XP/2003/Vista]"
        Case "AutorunInfs"
            sHelp = "The autorun.inf files on a CD tell Windows " & _
                    "what to start when the CD is inserted. " & _
                    "This technique can be applied to hard " & _
                    "disks too, making Windows executing an " & _
                    "arbitrary file when the drive is opened in " & _
                    "Explorer." & _
                    vbCrLf & "[Windows 95/98/NT4/2000/XP/XPSP2/2003]"
        Case "ScriptPolicies", "ScriptPolicies", "ScriptPolicies"
            sHelp = "Windows NT, 2000 and XP (and newer) can be " & _
                    "setup to execute scripts when the system " & _
                    "starts or shuts down, or when a user logs " & _
                    "on or off." & _
                    vbCrLf & "[Windows 2000/XP/2003/Vista]"
        Case "BatFiles", "BatFileswinstart.bat", "BatFilesdosstart.bat", "BatFilesautoexec.bat", "BatFilesconfig.sys", "BatFilesautoexec.nt", "BatFilesconfig.nt"
            sHelp = "Windows 95, 98, ME (and sometimes newer) " & _
                    "uses *.bat files for starting certain " & _
                    "components of the system. Anything in these " & _
                    "files is executed on system startup." & _
                    vbCrLf & "[winstart.bat,dosstart.bat: Windows 95/98/98SE]" & _
                    vbCrLf & "[autoexec.bat,config.sys: Windows 9x/ME]" & _
                    vbCrLf & "[autoexec.nl,config.nt: Windows NT4/2000/XP/2003/Vista]"
        Case "OnRebootActions", "OnRebootActionsBootExecute", "OnRebootActionsWininit.ini", "OnRebootActionsWininit.bak"
            sHelp = "Windows can be setup to perform actions " & _
                    "(such as deleting or renaming a file) when " & _
                    "the system is restarted. These instructions " & _
                    "can be present in the Registry or in the " & _
                    "wininit.ini file. The wininit.bak file is " & _
                    "a backup of the wininit.ini used last time " & _
                    "and is created by Windows itself." & _
                    vbCrLf & "[wininit.ini: Windows 9x/ME]" & _
                    vbCrLf & "[other: Windows NT4/2000/XP/2003/Vista]"
        Case "ShellCommands", "ShellCommandsbat", "ShellCommandscmd", "ShellCommandscom", "ShellCommandsexe", "ShellCommandshta", "ShellCommandsjs", "ShellCommandsjse", "ShellCommandspif", "ShellCommandsscr", "ShellCommandstxt", "ShellCommandsvbe", "ShellCommandsvbs", "ShellCommandswsf", "ShellCommandswsh"
            sHelp = "These settings control what Windows does " & _
                    "when a certain filetype is opened - if it " & _
                    "is run by itself or if some application is " & _
                    "used to open it. This setting can be " & _
                    "altered to load a different or additional " & _
                    "program." & _
                    vbCrLf & "[All Windows versions]"
        Case "Services", "NTServices", "VxDServices"
            sHelp = "The Windows services play an important role " & _
                    "in Windows NT4, 2000 and XP (and newer), " & _
                    "and a smaller one in Windows 95, 98 and ME. " & _
                    "All listed items are started on system " & _
                    "startup." & _
                    vbCrLf & "[NT Services: Windows NT4/2000/XP/2003/Vista]" & _
                    vbCrLf & "[VxD Services: All Windows versions]"
        Case "DriverFilters", "DriverFiltersClass", "DriverFiltersDevice"
            sHelp = "Device drivers can be setup to insert " & _
                    "themselves into the chain of drivers for a " & _
                    "hardware device, either above it " & _
                    "(UpperFilters) or below it (LowerFilters), " & _
                    "enabling them to filter messages before " & _
                    "passing them on to other drivers. The class " & _
                    "filters apply to generic types of hardware, " & _
                    "the device filters apply to specific " & _
                    "devices only." & _
                    vbCrLf & "[Windows 2000/XP/2003/Vista]"
        Case "WinLogonAutoruns", "WinLogonL", "WinLogonW", "WinLogonNotify", "WinLogonGinaDLL", "WinLogonGPExtensions"
            sHelp = "When a user logs on, winlogon.exe performs " & _
                    "a number of actions. A number of DLL " & _
                    "libraries and programs are started and will " & _
                    "stay loaded until the user logs off. The " & _
                    "Notify subkey contains links to dll " & _
                    "libraries that are loaded when a particular " & _
                    "event occurs, like logging on/off, starting " & _
                    "up, shutting down, or activating the " & _
                    "screensaver." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "BHOs", "BHO"
            sHelp = "Browser Helper Objects are loaded with " & _
                    "Internet Explorer (and sometimes Windows " & _
                    "Explorer) and can view, control and modify " & _
                    "everything that MSIE does." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "ActiveX"
            sHelp = "ActiveX objects. When loaded for the first " & _
                    "time, these objects will run before the " & _
                    "user logs on." & _
                    vbCrLf & "[All Windows versions]"
        Case "IEToolbars", "IEToolbarsUser", "IEToolbarsSystem"
            sHelp = "The toolbars that are loaded and displayed " & _
                    "(if enabled) when Internet Explorer is " & _
                    "loaded." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "IEExtensions"
            sHelp = "The additional buttons and tools that are " & _
                    "loaded and displayed (if enabled) when " & _
                    "Internet Explorer is loaded." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "IEExplBars"
            sHelp = "The additional special bars that are loaded " & _
                    "and displayed (if enabled) when Internet " & _
                    "Explorer is loaded." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "IEMenuExt"
            sHelp = "The additional commands in the right-click " & _
                    "menu of websites that are loaded and " & _
                    "displayed (if enabled) when Internet " & _
                    "Explorer is loaded." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "IEBands"
            sHelp = "All the bands that exist on the system that " & _
                    "are loaded and displayed (if enabled) when " & _
                    "Internet Explorer is loaded. The Search bar " & _
                    "is an example of a band." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "DPFs", "DPF"
            sHelp = "The objects in the 'Downloaded Program " & _
                    "Files' are additional plug-ins to Internet " & _
                    "Explorer and are loaded when it starts." & _
                    "When deleted, these objects are downloaded " & _
                    "and installed again (after prompting)." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "URLSearchHooks"
            sHelp = "Hooks that can intercept and redirect URLs " & _
                    "entered into the Internet Explorer address " & _
                    "bar." & _
                    vbCrLf & "[Internet Explorer 4.0 and newer]"
        Case "ExplorerClones"
            sHelp = "Due to the way Windows looks up paths for " & _
                    "files, putting a file Explorer.exe in " & _
                    "folders other than the Windows folder can " & _
                    "cause them to be executed before the normal " & _
                    "Explorer.exe file. There should be only an " & _
                    "Explorer.exe file in the Windows directory." & _
                    vbCrLf & "[All Windows versions]"
        Case "ImageFileExecution"
            sHelp = "In the 'Image File Execution' Registry key, " & _
                    "a file can be setup to be used as a " & _
                    "debugger for another program. Whenever the " & _
                    "host program is started, the 'debugger' " & _
                    "program is loaded as well." & _
                    vbCrLf & "Note: when a debugger file deleted " & _
                    "but still set, the host program will not start!" & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "ContextMenuHandlers"
            sHelp = "When an item is added to the right-click " & _
                    "menu of certain items, the associated dll " & _
                    "library is loaded into memory when the " & _
                    "system starts." & _
                    vbCrLf & "[All Windows versions]"
        Case "ColumnHandlers"
            sHelp = "Programmers can add custom columns to the " & _
                    "'Detailed' view of Explorer. The dll " & _
                    "library associated with the column is " & _
                    "loaded into memory whenever the 'Detailed' " & _
                    "view is active." & _
                    vbCrLf & "[Windows ME/2000/XP/2003/Vista]"
        Case "ShellExecuteHooks"
            sHelp = "This Registry key lists all COM objects " & _
                    "(dll libraries) that trap execute commands." & _
                    vbCrLf & "[All Windows versions]"
        Case "ShellExts"
            sHelp = "User interface extensions that are approved by " & _
                    "the system or the user." & _
                    vbCrLf & "[All Windows versions]"
        Case "RunRegkeys"
            sHelp = "There are several 'Run' keys in the Registry " & _
                    "that contain values that are executed on system " & _
                    "startup or when the user logs on. This type of " & _
                    "autostart location is very common." & _
                    vbCrLf & "[Run,RunOnce,RunOnceEx: All Windows versions]" & _
                    vbCrLf & "[RunServices,RunServicesOnce, Windows 9x/ME]"
        Case "RunExRegkeys"
            sHelp = "There are several 'Run' keys in the " & _
                    "Registry that have subkeys which " & _
                    "contain values themselves, that are " & _
                    "executed on system startup (mostly only " & _
                    "once, after which they keys are deleted by " & _
                    "Windows)." & _
                    vbCrLf & "[Run\*, RunOnce\*: Windows 2000]" & _
                    vbCrLf & "[Run\Setup: All Windows versions]"
        Case "Policies" '"Policy",
            sHelp = "The Windows Policies can be used to start " & _
                    "applications when the system is started or " & _
                    "a user logs in. Also, it can be used to " & _
                    "setup the work environment (Shell) for a " & _
                    "user, including any restrictions to the " & _
                    "user's actions." & _
                    vbCrLf & "[All Windows versions]"
        Case "Protocols", "ProtocolsFilter", "ProtocolsHandler"
            sHelp = "The pluggable MIME filters and protocol " & _
                    "handlers can see and manipulate everything " & _
                    "that passes through them through Internet " & _
                    "Explorer and the URL Moniker." & _
                    vbCrLf & "[All Windows versions]"
        Case "UtilityManager"
            sHelp = "The Accessibility Utility Manager controls " & _
                    "if programs like Magnifier and Screen " & _
                    "Reader are loaded when Windows starts up. " & _
                    "It can also be configured to start other " & _
                    "programs on system startup." & _
                    vbCrLf & "[Windows 2000/XP/2003/Vista]"
        Case "WOW", "WOWKnownDlls", "WOWKnownDlls32b"
            sHelp = "WOW is the 16-bit compatibility manager of " & _
                    "Windows. Its Registry key contains several " & _
                    "values that can be used to start programs " & _
                    "or dll libraries under certain conditions. " & _
                    "The KnownDlls lists contain all dll " & _
                    "libraries that can be run from the System " & _
                    "folder only. If a dll with the same name is " & _
                    "present somewhere else, it will not be " & _
                    "loaded." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "ShellServiceObjectDelayLoad", "SSODL"
            sHelp = "The shell Service Objects are loaded by " & _
                    "Explorer as soon as it starts." & _
                    vbCrLf & "[All Windows versions]"
        Case "SharedTaskScheduler"
            sHelp = "The shared Task Scheduler items (in the " & _
                    "Registry) are loaded everytime you open " & _
                    "Explorer." & _
                    vbCrLf & "[All Windows versions]"
        Case "MPRServices"
            sHelp = "Similar to the 'Notify' Registry key in " & _
                    "Windows 2000/XP, this Registry key can be " & _
                    "used by Windows 95/98/ME to load a dll." & _
                    vbCrLf & "[Windows 9x/ME]"
        Case "CmdProcAutorun"
            sHelp = "When the Windows 2000/XP command line is " & _
                    "opened (CMD.EXE), it will look for this " & _
                    "Registry value and execute any program in " & _
                    "it before starting." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "WinsockLSP", "WinsockLSPProtocols", "WinsockLSPNamespaces"
            sHelp = "Layered Service Providers (LSP) are small " & _
                    "pieces of software that can be added or " & _
                    "inserted into the Windows TCP/IP handler by " & _
                    "other software. Data outward bound from " & _
                    "your computer to a legitimate destination " & _
                    "on the Internet can be intercepted by an " & _
                    "LSP and sent somewhere other than where " & _
                    "you intend it to go. They are executed " & _
                    "before user login." & _
                    vbCrLf & "[All Windows versions]"
        Case "3rdPartyApps"
            sHelp = "Autostarts from non-Microsoft programs."
        Case "ICQ"
            sHelp = "ICQ contains a component named ICQNET that will " & _
                    "run any applications setup for it when an " & _
                    "Internet connection is detected."
        Case "mIRC", "mIRCmirc.ini", "mIRCrfiles", "mIRCafiles", "mIRCperform.ini"
            sHelp = "mIRC can be setup to load custom scripts that " & _
                    "have malicious purposes. These scripts can " & _
                    "perform virtually any action on a system."
        Case "DisabledEnums"
            sHelp = "All autostart items that are disabled, or " & _
                    "serve only to block websites, domains, " & _
                    "ActiveX objects, etc."
        Case "Hijack"
            sHelp = "The most common places on a Windows system " & _
                    "that a browser hijacker will hook into, " & _
                    "apart from the autostart locations."
        Case "ResetWebSettings"
            sHelp = "The URLs that Internet Explorer is reset to " & _
                    "are stored in a file called IERESET.INF, " & _
                    "and these URLs are used whenever you use " & _
                    "'Reset Web Settings'." & _
                    vbCrLf & "[Internet Explorer 5 and newer]"
        Case "IEURLs"
            sHelp = "All URLs that have been setup for Internet " & _
                    "Explorer to use as start pages, search " & _
                    "pages, search bars, etc. Note: some of " & _
                    "these are only used under certain " & _
                    "conditions." & _
                    vbCrLf & "[All Internet Explorer versions]"
        Case "URLPrefix", "URLDefaultPrefix"
            sHelp = "When an address is typed into the Address " & _
                    "Bar without the usual 'http://' or 'www' " & _
                    "prefix, Internet Explorer will look here " & _
                    "for the default text to prefix." & _
                    vbCrLf & "[Internet Explorer 5 and newer]"
        'Case "PolicyRestrictions"
        '    sHelp = "The Windows Policies can be used to restrict individual users in what they can do on the system. Mostly this involves restricting access to system tools and system functions."
        Case "HostsFilePath"
            sHelp = "The location of the hosts file can be " & _
                    "changed from the default in Windows 2000/XP " & _
                    "(c:\windows\system32\drivers\etc). If a " & _
                    "hosts file is present at any other " & _
                    "location, it is ignored." & _
                    vbCrLf & "[Windows NT4/2000/XP/Vista/2003]"
        Case "HostsFile"
            sHelp = "The hosts file is a file that can be used " & _
                    "somewhat like a local DNS server, mapping " & _
                    "host names to IP addresses. While this " & _
                    "speeds up name resolving, it can also be " & _
                    "used to completely block a domain by " & _
                    "mapping it to a non-routable IP address " & _
                    "like 127.0.0.1 or 0.0.0.0." & _
                    vbCrLf & "[All Windows versions]"
        Case "Killbits"
            sHelp = "Internet Explorer checks this Registry " & _
                    "key before loading and executing an " & _
                    "ActiveX control. If the CLSID of the " & _
                    "control is marked as incompatible (using a " & _
                    "'kill bit'), it will not be loaded. Note " & _
                    "that only blocked ActiveX objects that " & _
                    "actually exist on the system are listed " & _
                    "here." & _
                    vbCrLf & "[Internet Explorer 4 and newer]"
        Case "Zones"
            sHelp = "Internet Explorer uses several zones to " & _
                    "determine the security policy for websites, " & _
                    "some with more permissions (Trusted Zone) " & _
                    "or with less permissions (Restricted Zone). " & _
                    "This can be used to allow a site more " & _
                    "privileges (e.g. ActiveX, Java, scripting) " & _
                    "or defang it." & _
                    vbCrLf & "[Internet Explorer 4 and newer]"
        Case "msconfig9x"
            sHelp = "All the autostart items that are disabled " & _
                    "using the MSConfig tool in Windows 98/ME." & _
                    vbCrLf & "[Windows 98/ME]"
        Case "msconfigxp"
            sHelp = "All the autostart items that are disabled " & _
                    "using the MSConfig tool in Windows XP." & _
                    vbCrLf & "[Windows XP/2000/Vista] "
        Case "StoppedServices", "StoppedOnlyServices", "DisabledServices"
            sHelp = "NT Services that are not currently running, " & _
                    "either because they are stopped, set to run " & _
                    "manually, or are disabled." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "XPSecurity", "XPSecurityCenter"
            sHelp = "Windows XP SP2 has a central controlpoint, " & _
                    "called Security Center, where antivirus, " & _
                    "firewall and update settings are managed " & _
                    "from the Control Panel. However, all three " & _
                    "monitors can be overridden or disabled." & _
                    vbCrLf & "[Windows XP Service Pack 2]"
        Case "XPSecurityRestore"
            sHelp = "Windows XP has a feature called System " & _
                    "Restore which enables the user to save " & _
                    "snapshots of the system configuration, and " & _
                    "rollback to one of those snapshots later. " & _
                    "While restore points can consume a lot of " & _
                    "disk space, the rollbacks are very " & _
                    "effective." & _
                    vbCrLf & "[Windows ME/XP]"
        Case "XPFirewall", "XPFirewallDomain", "XPFirewallStandard", "XPFirewallDomainApps", "XPFirewallDomainPorts", "XPFirewallStandard", "XPFirewallStandardApps", "XPFirewallStandardPorts"
            sHelp = "The Windows Firewall that comes with Windows " & _
                    "XP SP 2 and newer can be setup to allow programs " & _
                    "through based on a filename or a port number " & _
                    "(though not both in one rule). " & _
                    "While it has rudimentary options to set the " & _
                    "remote IP range for a 'rule' as well, all " & _
                    "rules are stored in the Registry unencrypted, " & _
                    "allowing any program to add itself to the " & _
                    "Exceptions list." & _
                    vbCrLf & "[Windows XP Service Pack 2]"
        Case "PrintMonitors"
            sHelp = "This Registry key contains all drivers that " & _
                    "monitor print ports." & _
                    vbCrLf & "[All Windows versions]"
        Case "SecurityProviders"
            sHelp = "This Registry value contains a list of DLL " & _
                    "filenames that are loaded by Windows at " & _
                    "startup." & _
                    vbCrLf & "[All Windows versions]"
        Case "DesktopComponents"
            sHelp = "Desktop Components are ActiveX objects that can " & _
                    "be made part of the desktop whenever Active " & _
                    "Desktop is enabled (introduced in Windows 98), " & _
                    "where it runs as a (small) website widget." & _
                    vbCrLf & "[Windows 98/ME/2000/XP/2003/Vista]"
        Case "AppPaths"
            sHelp = "The 'App Paths' Registry key maps an executable " & _
                    "to its full path, so you can type 'wordpad' " & _
                    "without a path in the Run dialog to start " & _
                    "Wordpad. It is also possible to modify the " & _
                    "setting to start a completely different " & _
                    "program, though." & _
                    vbCrLf & "[All Windows versions]"
        Case "MountPoints", "MountPoints2"
            sHelp = "Much like the autorun.inf file, Windows has a " & _
                    "Registry key that can set a program to " & _
                    "automatically start when a diskette or CD is " & _
                    "inserted into the computer. This MountPoints " & _
                    "key can also be configured to start a program " & _
                    "when a user opens or explores a hard drive." & _
                    vbCrLf & "[All Windows versions]"
        Case "SafeBootMinimal", "SafeBootNetwork", "SafeBootAltShell"
            sHelp = "When the system is started in Safe Mode, " & _
                    "Windows will load these services ONLY, " & _
                    "instead of all of the ones listed above." & _
                    "When using Safe Mode with Network Support, " & _
                    "a few more services are loaded to for " & _
                    "network support." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "SafeBootAlt"
            sHelp = "When the system is started in Safe Mode " & _
                    "with Command Prompt, this value sets " & _
                    "the filename for the command prompt " & _
                    "which replaces the standard shell " & _
                    "(Explorer.exe). For reasons unknown, " & _
                    "this alternate shell is also used (if " & _
                    "enabled) when starting the system in normal " & _
                    "mode (!)." & _
                    vbCrLf & "[Windows NT4/2000/XP/2003/Vista]"
        Case "WindowsDefender", "WindowsDefenderDisabled"
            sHelp = "The Microsoft Window Defender can be disable " & _
                    "with a single Registry value. When this is " & _
                    "to 1, the system is not protected." & _
                    vbCrLf & "[Windows XP/2003/Vista]"
        Case "LsaPackages", "LsaPackagesAuth", "LsaPackagesNoti", "LsaPackagesSecu"
            sHelp = "The packages listed here are part of the " & _
                    "Windows NT Local Security Authority (LSA) and " & _
                    "are called on certain system events." & vbCrLf & _
                    "Authentication packages are called when a user " & _
                    "attempts to log on to Windows." & vbCrLf & _
                    "Notification packages are called when passwords " & _
                    "are set or changed." & vbCrLf & _
                    "Security packages are implementations of security " & _
                    "protocols such as Kerberos, SSL and TLS." & _
                    vbCrLf & "[Windows 2000/XP/2003/Vista]"
        Case "Drivers", "Drivers32RDP"
            sHelp = "The drivers and libraries here are loaded whenever " & _
                    "the application involved requires it." & _
                    vbCrLf & "[Windows NT/2000/XP/2003/Vista]"
        
        Case "System", "Users", "Hardware"
        Case Else
            If IsRunningInIDE Then sHelp = "(not found!) " & sName
    End Select
    
    GetHelpText = sHelp
End Function

Public Function GetSectionFromKey$(sName$)
    Dim i&
    'strip usernames from node name
    For i = 0 To UBound(sUsernames)
        If InStr(sName, sUsernames(i)) > 0 Then
            If Len(sName) = Len("Users" & sUsernames(i)) Then
                GetSectionFromKey = "These are the startup items for the user '" & MapSIDToUsername(sUsernames(i)) & "'"
                Exit Function
            Else
                sName = Mid$(sName, Len(sUsernames(i)) + 1)
            End If
        End If
    Next i
    'strip hardware cfgs from node name
    For i = 1 To UBound(sHardwareCfgs)
        If InStr(sName, sHardwareCfgs(i)) > 0 Then
            If Len(sName) = Len("Hardware" & sHardwareCfgs(i)) Then
                GetSectionFromKey = "These are the startup items for the hardware configuration '" & MapControlSetToHardwareCfg(sHardwareCfgs(i)) & "'"
                Exit Function
            Else
                sName = Mid$(sName, Len(sHardwareCfgs(i)) + 1)
            End If
        End If
    Next i
    
    'strip the numbers from the node name in case it's a child node
    If InStr(sName, "Ticks") > 0 Then
        GetSectionFromKey = "The time it took StartupList to enumerate the items in this section."
        Exit Function
    End If
    If InStr(2, sName, "System") > 0 Then sName = Replace(sName, "System", vbNullString)
    If InStr(2, sName, "User") > 0 Then sName = Replace(sName, "User", vbNullString)
    If InStr(2, sName, "Shell") > 0 Then sName = Replace(sName, "Shell", vbNullString)
    If InStr(2, sName, "Lower") > 0 Then sName = Replace(sName, "Lower", vbNullString)
    If InStr(2, sName, "Upper") > 0 Then sName = Replace(sName, "Upper", vbNullString)
    If InStr(2, sName, "Range") > 0 Then sName = Replace(sName, "Range", vbNullString)
    If InStr(2, sName, "Val") > 0 Then sName = Replace(sName, "Val", vbNullString)
    If InStr(2, sName, "app") > 0 Then sName = Replace(sName, "app", vbNullString)
    If InStr(2, sName, "dde") > 0 Then
        sName = Replace(sName, "app", vbNullString)
    End If
    Do Until Not IsNumeric(Right$(sName, 1)) And _
       Right$(sName, 1) <> "." And _
       Right$(sName, 3) <> "sub" And _
       Right$(sName, 3) <> "sup"
        If IsNumeric(Right$(sName, 1)) Then sName = Left$(sName, Len(sName) - 1)
        If Right$(sName, 1) = "." Then sName = Left$(sName, Len(sName) - 1)
        If Right$(sName, 3) = "sub" Then sName = Left$(sName, Len(sName) - 3)
        If Right$(sName, 3) = "sup" Then sName = Left$(sName, Len(sName) - 3)
    Loop
    If InStr(sName, "IniMing") > 0 Then sName = Replace(sName, "IniMing", "IniMapping")
    If InStr(sName, "AutoStartFolders Startup") > 0 Then sName = Replace(sName, "Folders Startup", "FoldersUser Startup")
    If InStr(sName, "AutoStartFolders Common Startup") > 0 Then sName = Replace(sName, "Folders Common", "FoldersUser Common")
    GetSectionFromKey = sName
End Function
