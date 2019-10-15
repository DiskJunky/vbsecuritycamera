Attribute VB_Name = "modRegistry"
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   General Use
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           11 Januarary 2001
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Contact:        DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This is used to manipulate the registry and store and
'                   retrieve data from it.
'  ---------------------------------------------------------------------------
'   Terminology:
'
'---    Hives
'These are like the different drives that you see in explorer. Each hive has
'it's own purpose, eg, Current_User contains data that only applys to the
'person using the computer at the moment, not the entire windows system,
'whereas Local_Machine applys to the all users, not just the current one.
'
'---    Sub Keys()
'These can be though of as directories, and they work almost exactly the
'same way. However, each Sub key (or "folder"), as it's own default Value
'(see next section). The can be set using CreateRegString, but specifying a
'blank EntryLabel parameter. Obviously, you create sub keys to help group
'your application data
'
'---    Values
'These can be thought of as FileNames (just names, not the actual files),
'or labels. They are just there to put a name to the data your are storing
'in the registry so that you can retrieve it later. In general there are 3
'types of Value, String, Double and Binary. Keep in mind that these types
'can store their data in different ways, eg Double (or numeric if you prefer),
'can store it's most signifigant bits, either in the normal way, or in reverse
'(don't ask - it gets complicated :) ). It's always best to check in the
'registry editor (RegEdit) if you are reading data from the registry for
'another program, so that you are sure you have the right type.
'
'---    Data
'This is the information that you are actually storing in the registry.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------


'all variables must be declared
Option Explicit

'this module cannot be accessed from outside this project
Option Private Module

'text comparisons are not case sensitive
Option Compare Text

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------

'api calls to retereive the system and windows folders
Private Declare Function GetSystemDirectory _
        Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
            (ByVal lpBuffer As String, _
             ByVal nSize As Long) _
             As Long
Private Declare Function GetWindowsDirectory _
        Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
            (ByVal lpBuffer As String, _
             ByVal nSize As Long) _
             As Long

'get the location of the temp directory on the system
Private Declare Function GetTempDirectory _
        Lib "kernel32" _
        Alias "GetTempPathA" _
            (ByVal lBufferLength As Long, _
             ByVal strBuffer As String) _
             As Long

'get information about the current operating system
Private Declare Function GetVersionEx _
        Lib "kernel32" _
        Alias "GetVersionExA" _
            (ByRef lpVersionInformation As OSVERSIONINFO) _
             As Long

'registry api calls

'close an open registry key
Private Declare Function RegCloseKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long
             
'connect with the registry on a remote machine
Private Declare Function RegConnectRegistry _
        Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
            (ByVal lpMachineName As String, _
             ByVal hKey As Long, _
             phkResult As Long) _
             As Long

'create a new registry key
Private Declare Function RegCreateKey _
        Lib "advapi32.dll" _
        Alias "RegCreateKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             phkResult As Long) _
             As Long
'create new - entended
Private Declare Function RegCreateKeyEx _
        Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal Reserved As Long, _
             ByVal lpClass As String, _
             ByVal dwOptions As Long, _
             ByVal samDesired As Long, _
             lpSecurityAttributes As SECURITY_ATTRIBUTES, _
             phkResult As Long, _
             lpdwDisposition As Long) _
             As Long

'delete the specified registry key (also any sub keys
'for non-NT based systems)
Private Declare Function RegDeleteKey _
        Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String) _
             As Long

'delete a registry value
Private Declare Function RegDeleteValue _
        Lib "advapi32.dll" _
        Alias "RegDeleteValueA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String) _
             As Long

'return a list of registry sub keys in the specified key
Private Declare Function RegEnumKey _
        Lib "advapi32.dll" _
        Alias "RegEnumKeyA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpName As String, _
             ByVal cbName As Long) _
             As Long
Private Declare Function RegEnumKeyEx _
        Lib "advapi32.dll" _
        Alias "RegEnumKeyExA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpName As String, _
             lpcbName As Long, _
             ByVal lpReserved As Long, _
             ByVal lpClass As String, _
             lpcbClass As Long, _
             lpftLastWriteTime As FILETIME) _
             As Long

'get a list of registry values in a key
Private Declare Function RegEnumValue _
        Lib "advapi32.dll" _
        Alias "RegEnumValueA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpValueName As String, _
             lpcbValueName As Long, _
             ByVal lpReserved As Long, _
             lpType As Long, _
             lpData As Byte, _
             lpcbData As Long) _
             As Long

'writes all the attributes of the specified open key
'into the registry
Private Declare Function RegFlushKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long

'get the security attributes of the specified key
Private Declare Function RegGetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR, _
             lpcbSecurityDescriptor As Long) _
             As Long

'creates a subkey under HKEY_USER or HKEY_LOCAL_MACHINE
'and stores registration information from a specified
'file into that subkey. This registration information
'is in the form of a hive. A hive is a discrete body of
'keys, subkeys, and values that is rooted at the top of
'the registry hierarchy. A hive is backed by a single
'file and .LOG file
Private Declare Function RegLoadKey _
        Lib "advapi32.dll" _
        Alias "RegLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpFile As String) _
             As Long

'notify a specified procedure (use the AddressOf
'operator), that a key has changed
Private Declare Function RegNotifyChangeKeyValue _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal bWatchSubtree As Long, _
             ByVal dwNotifyFilter As Long, _
             ByVal hEvent As Long, _
             ByVal fAsynchronus As Long) _
             As Long

'open a registry key for access
Private Declare Function RegOpenKey _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             phkResult As Long) _
             As Long
Private Declare Function RegOpenKeyEx _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal ulOptions As Long, _
             ByVal samDesired As Long, _
             phkResult As Long) _
             As Long

'get key information
Private Declare Function RegQueryInfoKey _
        Lib "advapi32.dll" _
        Alias "RegQueryInfoKeyA" _
            (ByVal hKey As Long, _
             ByVal lpClass As String, _
             lpcbClass As Long, _
             ByVal lpReserved As Long, _
             lpcSubKeys As Long, _
             lpcbMaxSubKeyLen As Long, _
             lpcbMaxClassLen As Long, _
             lpcValues As Long, _
             lpcbMaxValueNameLen As Long, _
             lpcbMaxValueLen As Long, _
             lpcbSecurityDescriptor As Long, _
             lpftLastWriteTime As FILETIME) _
             As Long

'get value information. Note that if you declare the
'lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValue _
        Lib "advapi32.dll" _
        Alias "RegQueryValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpValue As String, _
             lpcbValue As Long) _
             As Long
Private Declare Function RegQueryValueEx _
        Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal lpReserved As Long, _
             lpType As Long, _
             lpData As Any, _
             lpcbData As Long) _
             As Long

'replace one key with another
Private Declare Function RegReplaceKey _
        Lib "advapi32.dll" _
        Alias "RegReplaceKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpNewFile As String, _
             ByVal lpOldFile As String) _
             As Long

'reads registry information from a file and enters it
'into the registry
Private Declare Function RegRestoreKey _
        Lib "advapi32.dll" _
        Alias "RegRestoreKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             ByVal dwFlags As Long) _
             As Long

'saves a registry key and all its values to a file
Private Declare Function RegSaveKey _
        Lib "advapi32.dll" _
        Alias "RegSaveKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             lpSecurityAttributes As SECURITY_ATTRIBUTES) _
             As Long

'set the security attributes of the specified registry
'key
Private Declare Function RegSetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR) _
             As Long

'set the information of an existing value. Note that if
'you declare the lpData parameter as String, you must
'pass it By Value.
Private Declare Function RegSetValue _
        Lib "advapi32.dll" _
        Alias "RegSetValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal dwType As Long, _
             ByVal lpData As String, _
             ByVal cbData As Long) _
             As Long
Private Declare Function RegSetValueEx _
        Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal Reserved As Long, _
             ByVal dwType As Long, _
             lpData As Any, _
             ByVal cbData As Long) _
             As Long
             
'unloads a registry key and its values from the registry
Private Declare Function RegUnLoadKey _
        Lib "advapi32.dll" _
        Alias "RegUnLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String) _
             As Long

'system information api calls
Private Declare Sub GlobalMemoryStatus _
        Lib "kernel32" _
            (lpBuffer As MEMORYSTATUS)
            
Private Declare Function GetDiskFreeSpace _
        Lib "kernel32" _
        Alias "GetDiskFreeSpaceA" _
            (ByVal lpRootPathName As String, _
             lpSectorsPerCluster As Long, _
             lpBytesPerSector As Long, _
             lpNumberOfFreeClusters As Long, _
             lpTotalNumberOfClusters As Long) _
             As Long
             
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () As Long

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
Public Enum MemType
    CPUUsage
    MemoryUsage
    TotalPhysical
    AvailablePhysical
    TotalPageFile
    AvailablePageFile
    TotalVirtual
    AvailableVirtual
    TotalDisk
    AvailableDisk
End Enum

Public Enum AccessType
    FileInput = 0
    FileOutPut = 1
    FileRandom = 2
    FileBinary = 3
    FileAppend = 4
End Enum

'registry root directory constants
Public Enum RegistryHives
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

'registry key constants
Public Enum RegistryKeyAccess
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_EVENT = &H1    '  Event contains key event record
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    READ_CONTROL = &H20000
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_REQUIRED = &HF0000
    SYNCHRONIZE = &H100000
    STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL + KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
End Enum

'registry value attributes
Public Enum RegistryKeyValues
    REG_CREATED_NEW_KEY = &H1               ' New Registry Key created
    REG_EXPAND_SZ = 2                       ' Unicode nul terminated string
    REG_FULL_RESOURCE_DESCRIPTOR = 9        ' Resource list in the hardware description
    REG_LINK = 6                            ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7                        ' Multiple Unicode strings
    REG_NONE = 0                            ' No value type
    REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    REG_NOTIFY_CHANGE_LAST_SET = &H4        ' Time stamp
    REG_NOTIFY_CHANGE_NAME = &H1            ' Create or delete (child)
    REG_NOTIFY_CHANGE_SECURITY = &H8
    REG_OPENED_EXISTING_KEY = &H2           ' Existing Key opened
    REG_OPTION_BACKUP_RESTORE = 4           ' open for backup or restore
    REG_OPTION_CREATE_LINK = 2              ' Created key is a symbolic link
    REG_OPTION_NON_VOLATILE = 0             ' Key is preserved when system is rebooted
    REG_OPTION_RESERVED = 0                 ' Parameter is reserved
    REG_OPTION_VOLATILE = 1                 ' Key is not preserved when system is rebooted
    REG_REFRESH_HIVE = &H2                  ' Unwind changes to last flush
    REG_RESOURCE_LIST = 8                   ' Resource list in the resource map
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    REG_SZ = 1                              ' Unicode nul terminated string
    REG_WHOLE_HIVE_VOLATILE = &H1           ' Restore whole hive volatile
    REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
    REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
End Enum

Public Enum RegistryDataTypes
    REG_DT_SZ = 1                  ' string data
    REG_DT_EXPAND_SZ = 2           ' expanded string
    REG_DT_BINARY = 3              ' Free form binary
    REG_DT_DWORD = 4               ' 32-bit number
    REG_DT_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_DT_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
End Enum

Public Enum RegistryLongTypes
    REG_BINARY = 3              ' Free form binary
    REG_DWORD = 4               ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
End Enum

'error codes returned
Public Enum RegistryErrorCodes
    ERROR_ACCESS_DENIED = 5&
    ERROR_INVALID_PARAMETER = 87    '  dderror
    ERROR_MORE_DATA = 234           '  dderror
    ERROR_NO_MORE_ITEMS = 259
    ERROR_SUCCESS = 0&
End Enum

'the shell folders like my documents, recycle bin, temp directory etc.
Public Enum ShellFoldersType
    'registry entry names
    ApplicationDataDir = 0
    TempInetFilesDir = 1
    CookiesDir = 2
    DesktopDir = 3
    FavouritesDir = 4
    FontsDir = 5
    HistoryDir = 6
    LocalAppDataDir = 7
    NetHoodDir = 8
    MyDocumentsDir = 9
    PrintHoodDir = 10
    StartProgramsDir = 11
    RecentDir = 12
    SendToDir = 13
    StartMenuDir = 14
    StartupDir = 15
    TemplatesDir = 16
    
    'these next items are not stored in the registry
    SystemDir = 17
    WindowsDir = 18
    TempDir = 19
End Enum

Public Enum StartLoginType
    RunBeforeLogin
    RunAfterLogin
End Enum

'the different nt privilages that can be set/unset
Public Enum EnumNTSettings
    'items that can be disabled on the Lock Screen
    CHANGE_PASSWORD = 0
    LOCK_WORKSTATION = 1
    REGISTRY_TOOLS = 2
    TASK_MGR = 3
    
    'the tabs on the Display Properties dialog box
    DISP_APPEARANCE_PAGE = 4
    DISP_BACKGROUND_PAGE = 5
    DISP_CPL = 6
    DISP_SCREENSAVER = 7
    DISP_SETTINGS = 8
End Enum

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
'holds information about the current operating system that the program is
'running on
Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128
End Type

'the current status of physical (ram), virtual memory and the page file.
Public Type MEMORYSTATUS
    dwLength                    As Long
    dwMemoryLoad                As Long
    dwTotalPhys                 As Long
    dwAvailPhys                 As Long
    dwTotalPageFile             As Long
    dwAvailPageFile             As Long
    dwTotalVirtual              As Long
    dwAvailVirtual              As Long
End Type

'defined structures needed
Public Type ACL
    AclRevision                 As Byte
    Sbz1                        As Byte
    AclSize                     As Integer
    AceCount                    As Integer
    Sbz2                        As Integer
End Type

Public Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength                     As Long
    lpSecurityDescriptor        As Long
    bInheritHandle              As Long
End Type

Public Type SECURITY_DESCRIPTOR
    Revision                    As Byte
    Sbz1                        As Byte
    Control                     As Long
    gstrOwner                   As Long
    Group                       As Long
    Sacl                        As ACL
    Dacl                        As ACL
End Type

Public Type TypeRegValues
    enmHive                     As RegistryHives        'holds the hive location of the value
    strSubKey                   As String               'holds the sub key location of the value
    strName                     As String               'holds the name of the value
    lngType                     As Long                 'holds the data type of the value
    varData                     As Variant              'holds the data of the value
End Type

'------------------------------------------------
'             MODULE-LEVEL CONSTANTS
'------------------------------------------------

'module constants
Private Const WIN_INFO_SUBKEY       As String = "Software\Microsoft\Windows\CurrentVersion"                 'HKEY_LOCAL_MACHINE
Private Const WIN_NT_INFO_SUBKEY    As String = "Software\Microsoft\Windows NT\CurrentVersion"              'HKEY_LOCAL_MACHINE
Private Const SHELL_FOLDERS_SUBKEY  As String = "Software\Microsoft\Windows\" + _
                                                "CurrentVersion\Explorer\Shell Folders"                     'HKEY_CURRENT_USER
Private Const COUNTRY_SUBKEY        As String = "Control Panel\International"                               'HKEY_CURRENT_USER
Private Const NT_SETTINGS           As String = WIN_INFO_SUBKEY & "\Policies\System"                        'HKEY_CURRENT_USER
Private Const W2K_SETTINGS          As String = WIN_INFO_SUBKEY & "\Group Policy Objects\LocalUser\" + _
                                                "Software\Microsoft\Windows\CurrentVersion\Policies\System" 'HKEY_CURRENT_USER
Private Const STARTUP_AL_SUBKEY     As String = WIN_INFO_SUBKEY & "\Run"                                    'run after login screen
Private Const STARTUP_BL_SUBKEY     As String = WIN_INFO_SUBKEY & "\RunServices"                            'run before login screen

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub CreateFileAssociation(ByVal strFileType As String, _
                                 ByVal strTypeDescription As String, _
                                 Optional ByVal strExeName As String, _
                                 Optional ByVal strExePath As String, _
                                 Optional ByVal strIconPath As String, _
                                 Optional ByVal blnAcceptMultipleArgs As Boolean = False)
    'This procedure will create a new association for a file. For anyone
    'who is unfamiliar with this, this means that if you were to double-
    'click on a file with the specified extention, the specified application
    'would start. eg, if you were to double click on a .txt file, notepad
    'would start and open the file.
    'Please note that if you wish to associate an icon, the icon has to be
    'a .ico file - no other file types are accepted. If you wish to use an
    'icon that is only in your exe (if your distributing you app for
    'example), then you need to save the icon as a file. This can be done
    'by using;
    '
    'Call SavePicture(MyControl.Picture, App.Path & "\MyIcon.ico")
    '
    'Although, please note that the picture must have originally been an
    'icon before you tried to save it as one.
    
    
    Dim lngResult   As Long
    Dim strFullPath As String
    Dim strAppKey   As String
    
    'exit procedure if the file type feild is blank
    If (strFileType = "") Then
        Exit Sub
    Else
        'if the first character is a dot, then remove it
        If Left(strFileType, 1) = "." Then
            strFileType = Right(strFileType, Len(strFileType) - 1)
        End If
        
        'check to see that the file type is only three characters long
        If Len(strFileType) > 3 Then
            strFileType = Left(strFileType, 3)
        End If
    
        'the type description should be no longer than 25 characters
        '(this is not necessary, but it keeps things neat in the registry)
        If Len(strTypeDescription) > 25 Then
            strTypeDescription = Left(strTypeDescription, 25)
        End If
    End If
    
    'set the default paths and exe name is they were not specified
    If strExeName = "" Then
        strExeName = App.EXEName
    End If
    
    If strExePath = "" Then
        strExePath = App.Path
    End If
    
    'make sure that the exename ends in ".exe"
    If LCase(Right(strExeName, 4)) <> ".exe" Then
        strExeName = strExeName & ".exe"
    End If
    
    'get the full path name of the exe
    If Right(strExePath, 1) = "\" Then
        'if the path already contains a trailing backslash (eg "d:\") then
        'don't add one when creating the path
        strFullPath = strExePath & strExeName
    Else
        'insert a backslash to seperate the name from the path
        strFullPath = strExePath & "\" & strExeName
    End If
    
    'check to make sure that the file exists
    If Dir(strFullPath) = "" Then
        'there is no file
        Exit Sub
    End If
    
    strFullPath = """" + strFullPath + """"
    
    'if no icon was specified, then use the icon for the exe
    If (strIconPath = "") Or (Dir(strIconPath) = "") Then
        strIconPath = strFullPath
    End If
    
    'create the file type extention in the registry
    Call CreateSubKey(HKEY_CLASSES_ROOT, "." & strFileType)
    
    'create the registry entry in the above sub key that holds the
    'sub key with the file path
    'eg, "MyApp.Description", "Vb6.Module", "Word.Document"
    'Note that a blank entry lable name means a default value for that key,
    'if any spaces are in the type description, they are replaced with
    'a "." character.
    strAppKey = Replace(Left(strExeName, Len(strExeName) - 4) & "." & _
                        strTypeDescription, " ", ".")
    Call CreateRegString(HKEY_CLASSES_ROOT, _
                         "." & strFileType, _
                         "", _
                         strAppKey)
    
    'create the key that will hold the applications path and type information.
    'additional commands can be put into the "Shell\Open\Command" sub key.
    'This means that when you right click on the file type, a popup menu
    'appears with the Open option. Other options can be inserted into this
    'menu by creating sub keys in the Shell key like; "Print\Command",
    '"Edit\Command", "Assemble\Command", "Split\Command" etc. where
    'the Command sub key contains a [default] entry with a command line
    'parameter to an executable file like "C:\Windows\Notepad.exe /p %1"
    Call CreateSubKey(HKEY_CLASSES_ROOT, _
                      strAppKey & "\Shell\Open\Command")
    
    'create the text that describes the file type
    Call CreateRegString(HKEY_CLASSES_ROOT, _
                         strAppKey, _
                         "", _
                         strTypeDescription)
    
    'can the application accept multiple arguments (files) being passed to it
    If blnAcceptMultipleArgs Then
        'create the command line parameter to open the file type with the
        'application specified
        Call CreateRegString(HKEY_CLASSES_ROOT, _
                             strAppKey & "\Shell\Open\Command", _
                             "", _
                             strFullPath & " ""%L""")
    
    Else
        'create the command line parameter to open the file type with the
        'application specified
        Call CreateRegString(HKEY_CLASSES_ROOT, _
                             strAppKey & "\Shell\Open\Command", _
                             "", _
                             strFullPath & " ""%1""")
    End If  'can the application accept multiple arguments (files) being passed to it
    
    'create the icon sub key
    Call CreateSubKey(HKEY_CLASSES_ROOT, _
                      strAppKey & "\DefaultIcon")
    
    'create the entry that points to the icon.
    If LCase(Right(strIconPath, 3)) = "exe" Then
        'get icon from .exe
        Call CreateRegString(HKEY_CLASSES_ROOT, _
                             strAppKey & "\DefaultIcon", _
                             "", _
                             strIconPath & ",1")
    Else
        'get icon from .ico file
        Call CreateRegString(HKEY_CLASSES_ROOT, _
                             strAppKey & "\DefaultIcon", _
                             "", _
                             strIconPath & ",0")
    End If
End Sub

Public Sub DeleteFileAssociation(ByVal strFileType As String)
    'This procedure will remove a file association. It is recommended that
    'you only remove an association that your application created, as once
    'the association is gone, it cannot be recreated without knowing the
    'file type, application involved and the icon assiciated with the file type.
    'See CreateFileAssociation for further information.
    
    Dim strSubKeyAssociation As String
    
    'validate the parameter
    
    'make sure that the parameter contains something
    If strFileType = "" Then
        Exit Sub
    End If
    
    'make sure that the first character is a dot (.)
    If Left(strFileType, 1) <> "." Then
        'insert dot
        strFileType = "." & strFileType
    End If
    
    'now we check the registry
    
    strSubKeyAssociation = ReadRegString(HKEY_CLASSES_ROOT, _
                                         strFileType, "")
    
    'if there was an error, then exit
    If LCase(Left(strSubKeyAssociation, 5)) = "error" Then
        Exit Sub
    End If
    
    'delete the commands and information about the selected file type
    Call DeleteSubKey(HKEY_CLASSES_ROOT, strSubKeyAssociation)
End Sub

Public Sub PutAppInStartup(ByVal strEntryLabel As String, _
                           Optional ByVal strFilePath As String, _
                           Optional ByVal blnStartup As StartLoginType = RunAfterLogin, _
                           Optional ByVal blnOverwrite As Boolean = False)
    'This will take an applications full path name and put it into the registry
    'to start the program either before or after the login screen in normally
    'loaded. If no app path is specified, then by default, it puts the current
    'project in to startup after the login screen. Existing enteries are not
    'overwritten. You could call this procedure like;
    '
    'Call PutAppInStartup("MyCoolApp", MyAppsFilePath, RunAfterLogin, False)
    '
    'or
    '
    'Call PutAppInStartup("MyCoolApp")
    '
    'See also RemoveAppFromStartup.
    
    
    Dim strSubKey   As String
    Dim strCheck    As String
    
    'check to see if a file path was specified
    If strFilePath = "" Then
        'specifiy the path from the current project
        
        'if the applications path is a root directory, then don't add a
        'backslash to the path
        If Right(App.Path, 1) = "\" Then
            strFilePath = App.Path & App.EXEName & ".exe"
        Else
            strFilePath = App.Path & "\" & App.EXEName & ".exe"
        End If
    End If
    
    'check to see if the file exists
    If (Dir(strFilePath) = "") Or (strEntryLabel = "") Then
        'can't find file. There is no point in making an entry for a file
        'that doesn't exist, so exit
        Exit Sub
    End If
    
    'create the sub key based on the options
    If blnStartup = RunAfterLogin Then
        'set the app to start after the login screen
        strSubKey = STARTUP_AL_SUBKEY
    Else
        'set the app to run before the login screen
        strSubKey = STARTUP_BL_SUBKEY
    End If
    
    'if the entry already exists and we don't want to overwrite, then exit
    strCheck = ReadRegString(HKEY_LOCAL_MACHINE, _
                             strSubKey, _
                             strEntryLabel)
    If (Not blnOverwrite) And (Left(strCheck, 5) <> "Error") Then
        Exit Sub
    End If
    
    'write to the registry
    Call CreateRegString(HKEY_LOCAL_MACHINE, _
                         strSubKey, _
                         strEntryLabel, _
                         strFilePath)
End Sub

Public Sub RemoveAppFromStartup(ByVal strEntryLabel As String, _
                                Optional ByVal blnStartup As StartLoginType = RunAfterLogin)
    'This procedure will remove an app from the startup be specifying
    'it's label and whether or not the app startsup before or after the
    'login screen. Also see the PutInStartup procedure.
    
    Dim strSubKey   As String
    Dim strCheck    As String
    
    'find the sub key depending on the startup gstrMethod
    If blnStartup = RunAfterLogin Then
        'startup after the login screen [default]
        strSubKey = STARTUP_AL_SUBKEY
    Else
        'startup before the login screen
        strSubKey = STARTUP_BL_SUBKEY
    End If
    
    'check to see if the entry exists
    strCheck = ReadRegString(HKEY_LOCAL_MACHINE, _
                             strSubKey, _
                             strEntryLabel)
    If Left(strCheck, 5) = "Error" Then
        'there was a problem accessing the key, so exit (eg, it might not exist)
        Exit Sub
    End If
    
    'delete the entry
    Call DeleteValue(HKEY_LOCAL_MACHINE, _
                     strSubKey, _
                     strEntryLabel)
End Sub

Public Sub CreateSubKey(ByVal enmHive As RegistryHives, _
                        ByVal strSubKey As String)
    'This procedure will create a sub key in the
    'specified header key.
    
    Dim lngResult   As Long
    Dim hKey        As Long
    
    'create the key
    lngResult = RegCreateKey(enmHive, _
                             strSubKey & Chr(0), _
                             hKey)
    
    'close the key
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub DeleteSubKey(ByVal enmHive As RegistryHives, _
                        ByVal strSubKey As String)
    'This procedure will delete a key from the registry. Please note that
    'the procedure will not delete key values.
    
    Dim lngResult   As Long     'holds any returned value from an api call
    Dim hKey        As Long     'holds a handle to the specified key
    
    'open the key
    lngResult = RegOpenKeyEx(enmHive, _
                             strSubKey & Chr(0), _
                             0&, _
                             KEY_ALL_ACCESS, _
                             hKey)
    
    'delete the key
    lngResult = RegDeleteKey(hKey, "")
    
    'close the key
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub DeleteValue(ByVal enmHive As RegistryHives, _
                       ByVal strSubKey As String, _
                       Optional ByVal strEntryLabel As String)
    'This will remove any registry key or entry value
    
    Dim lngResult       As Long
    Dim hKey            As Long
    Dim strTotalSubKey  As String
    
    'create the full registry subkey and entry label
    strTotalSubKey = strSubKey & Chr(0)
    
    'open the subkey/entry
    lngResult = RegOpenKeyEx(enmHive, _
                             strTotalSubKey, _
                             0&, _
                             KEY_ALL_ACCESS, _
                             hKey)
    
    'delete the key/entry from the registry
    lngResult = RegDeleteValue(hKey, strEntryLabel)
    
    'close the handle
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub CreateRegString(ByVal enmHive As RegistryHives, _
                           ByVal strSubKey As String, _
                           ByVal strEntryLabel As String, _
                           ByVal strText As String)
    'This will put some text into the specified key and entry label. This
    'data can be retrieved with the ReadRegString function
    
    Dim lngResult       As Long
    Dim hKey            As Long
    Dim strTotalSubKey  As String
    
    'create a complete sub key and entry path to send to the api call
    strTotalSubKey = strSubKey & Chr(0)
    
    'try to open the key first
    lngResult = RegOpenKeyEx(enmHive, _
                             strTotalSubKey, _
                             0, _
                             KEY_READ + KEY_WRITE, _
                             hKey)
    
    'if we couldn't open the key, then try and create it
    If (hKey = 0) Then
        'now create the sub key entry if it does not exist
        lngResult = RegCreateKey(enmHive, strTotalSubKey, hKey)
        
        'if no handle was returned, then exit
        If hKey = 0 Then
            Exit Sub
        End If
    End If
    
    'write the text into the key with the specified entry name
    lngResult = RegSetValueEx(hKey, _
                              strEntryLabel, _
                              0&, _
                              REG_SZ, _
                              ByVal strText, _
                              Len(strText))
    
    'close the opened key and exit
    lngResult = RegCloseKey(hKey)
End Sub

Public Function GetSubKeys(ByVal enmHive As RegistryHives, _
                           ByVal strSubKey As String, _
                           Optional ByRef lngFound As Long = 0, _
                           Optional ByRef blnSuccess As Boolean = False, _
                           Optional ByVal strPattern As String = "") _
                           As String()
    'This will return a list of sub keys in the specified sub key if able. If
    'the function is not able to retrieve a list of sub keys then the blnSuccess
    'parameter is set to False, otherwise it is set to True (sub keys returned).
    'If there are no sub keys, then the array will be empty and the blnSuccess
    'parameter is set to True. The array is zero based, but the lngFound
    'parameter holds the actual number of elements found. The strPattern
    'parameter is used for the Like parameter (see MSDN documentation for more
    'information on various pattern codes).
    
    Const BUFFER_SIZE   As Long = 255
    
    Dim hSubKey         As Long                 'holds the handle of the sub key if it exists
    Dim strSubKeys()    As String               'holds the list of sub keys retrieved
    Dim lngNumFound     As Long                 'holds the number of sub keys found
    Dim enmResult       As RegistryErrorCodes   'holds any returned value from an api call
    Dim strKeyName      As String * BUFFER_SIZE 'holds the sub key name returned from the api call
    Dim strTempName     As String               'holds the name of the sub key minus buffered characters. Used to test pattern matching if specified
    Dim lngKeyIndex     As Long                 'holds the index of the sub key. Not to be confused with the count of those beind returned in the array (which will be smaller if a pattern is specified)
    
    'initialise the variables/parameters
    blnSuccess = False
    lngFound = 0
    lngNumFound = lngFound
    ReDim strSubKeys(lngNumFound)
    GetSubKeys = strSubKeys()       'default return value is a blank array
    
    hSubKey = GetSubKeyHandle(enmHive, strSubKey)
    If (hSubKey = 0) Then
        'this sub key does not exist
        Exit Function
    End If
    
    'get the sub keys of the specified sub key
    lngKeyIndex = 0
    Do
        strKeyName = String(BUFFER_SIZE, vbNullChar)
        enmResult = RegEnumKey(hSubKey, lngKeyIndex, strKeyName, BUFFER_SIZE)
        
        'was data returned
        If (Left(strKeyName, 1) <> vbNullChar) Then
            strTempName = Mid(strKeyName, 1, InStr(1, strKeyName, vbNullChar) - 1)
            
            'does the sub key match the pattern (if one was specified)
            If (strPattern = "") Or (strTempName Like strPattern) Then
                'store the name
                ReDim Preserve strSubKeys(lngNumFound)
                strSubKeys(lngNumFound) = strTempName
                lngNumFound = lngNumFound + 1
            End If  'does the key match the pattern
        End If  'was a key found
        lngKeyIndex = lngKeyIndex + 1
    Loop Until enmResult <> ERROR_SUCCESS
    
    'return the results
    blnSuccess = True
    lngFound = lngNumFound
    GetSubKeys = strSubKeys()
End Function

Public Function GetValueNames(ByVal enmHive As RegistryHives, _
                              ByVal strSubKey As String, _
                              Optional ByRef lngFound As Long = 0, _
                              Optional ByRef blnSuccess As Boolean = False, _
                              Optional ByVal strPattern As String = "") _
                              As String()
    'This will return a list of values in the specified sub key if able. If
    'the function is not able to retrieve a list of values then the blnSuccess
    'parameter is set to False, otherwise it is set to True (values returned).
    'If there are no values, then the array will be empty and the blnSuccess
    'parameter is set to True. The array is zero based, but the lngFound
    'parameter holds the actual number of elements found. The strPattern
    'parameter is used for the Like parameter (see MSDN documentation for more
    'information on various pattern codes).
    
    Const BUFFER_SIZE   As Long = 255
    
    Dim hSubKey         As Long                 'holds the handle of the sub key if it exists
    Dim strValues()     As String               'holds the list of sub keys retrieved
    Dim lngNumFound     As Long                 'holds the number of sub keys found
    Dim enmResult       As Long                 'holds any returned value from an api call
    Dim strValueName    As String * BUFFER_SIZE 'holds the sub key name returned from the api call
    Dim strTempName     As String               'holds the name of the sub key minus buffered characters. Used to test pattern matching if specified
    Dim lngValueIndex   As Long                 'holds the index of the sub key. Not to be confused with the count of those beind returned in the array (which will be smaller if a pattern is specified)
    Dim enmType         As Long                 'holds the data type of the specified value
    Dim lngBufferSize   As Long                 'holds the size of the value name returned
    
    'initialise the variables/parameters
    blnSuccess = False
    lngFound = 0
    lngNumFound = lngFound
    ReDim strValues(lngNumFound)
    GetValueNames = strValues()       'default return value is a blank array
    
    hSubKey = GetSubKeyHandle(enmHive, strSubKey)
    If (hSubKey = 0) Then
        'this sub key does not exist
        Exit Function
    End If
    
    'get the values of the specified sub key
    lngValueIndex = 0
    Do
        strValueName = String(BUFFER_SIZE, vbNullChar)
        'strValueName = Space(BUFFER_SIZE - 1) + vbNullChar
        lngBufferSize = BUFFER_SIZE
        enmResult = RegEnumValue(hSubKey, _
                                 lngValueIndex, _
                                 strValueName, _
                                 lngBufferSize, _
                                 0&, _
                                 ByVal enmType, _
                                 ByVal CByte(0), _
                                 0&)
        
        'was data returned
        If (Left(strValueName, 1) <> vbNullChar) Then
            strTempName = Mid(strValueName, 1, InStr(1, strValueName, vbNullChar) - 1)
            
            'does the value match the pattern (if one was specified)
            If (strPattern = "") Or (strTempName Like strPattern) Then
                'store the name
                ReDim Preserve strValues(lngNumFound)
                strValues(lngNumFound) = strTempName
                lngNumFound = lngNumFound + 1
            End If  'does the value match the pattern
        End If  'was a key found
        lngValueIndex = lngValueIndex + 1
    Loop Until enmResult = ERROR_NO_MORE_ITEMS
    
    'return the results
    blnSuccess = True
    lngFound = lngNumFound
    GetValueNames = strValues()
End Function

Public Function GetWinDirectories(ByVal enmDirectory As ShellFoldersType) _
                                  As String
    'This function will return the specfied system directory like the desktop
    'directory, windows directory, temp folder, system directory etc.
    
    'registry entry names
    Const ApplicationData   As String = "AppData"
    Const TempInetFiles     As String = "Cache" 'temperory internet files
    Const Cookies           As String = "Cookies"
    Const Desktop           As String = "Desktop"
    Const Favourites        As String = "Favourites"
    Const Fonts             As String = "Fonts"
    Const History           As String = "History"
    Const LocalAppData      As String = "Local AppData"
    Const NetHood           As String = "NetHood"
    Const MyDocuments       As String = "Personal"
    Const PrintHood         As String = "PrintHood"
    Const StartPrograms     As String = "Programs"
    Const Recent            As String = "Recent"
    Const SendTo            As String = "SendTo"
    Const StartMenu         As String = "Start Menu"
    Const StartUp           As String = "Startup"
    Const Templates         As String = "Templates"
    
    
    Dim strResult As String
    Dim errResult As Long
    
    Select Case enmDirectory
        'registry entry names
        Case ApplicationDataDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      ApplicationData)
        
        Case TempInetFilesDir  'temperory internet files
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      TempInetFiles)
        
        Case CookiesDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Cookies)
        
        Case DesktopDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Desktop)
        
        Case FavouritesDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Favourites)
        
        Case FontsDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Fonts)
        
        Case HistoryDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      History)
        
        Case LocalAppDataDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      LocalAppData)
        
        Case NetHoodDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      NetHood)
        
        Case MyDocumentsDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      MyDocuments)
        
        Case PrintHoodDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      PrintHood)
        
        Case StartProgramsDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      StartPrograms)
        
        Case RecentDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Recent)
        
        Case SendToDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      SendTo)
        
        Case StartMenuDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      StartMenu)
        
        Case StartupDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      StartUp)
        
        Case TemplatesDir
            strResult = ReadRegString(HKEY_CURRENT_USER, _
                                      SHELL_FOLDERS_SUBKEY, _
                                      Templates)
        
        
        'these next items are not stored in the registry
        Case SystemDir
            strResult = Space(255)
            errResult = GetSystemDirectory(strResult, 255)
            
            'remove the null character
            If (InStr(1, strResult, vbNullChar) > 0) Then
                strResult = Left(strResult, InStr(1, strResult, vbNullChar) - 1)
            End If
            
        Case WindowsDir
            strResult = Space(255)
            errResult = GetWindowsDirectory(strResult, 255)
            
            'remove the null character
            If (InStr(1, strResult, vbNullChar) > 0) Then
                strResult = Left(strResult, InStr(1, strResult, vbNullChar) - 1)
            End If
            
        Case TempDir 'temperory folder is always in the Windows directory
            strResult = Space(255)
            errResult = GetTempDirectory(255, strResult)
            
            'remove the null character and add the name of the temperory folder
            If (InStr(1, strResult, vbNullChar) > 0) Then
                strResult = Left(strResult, InStr(1, strResult, vbNullChar) - 1)
            End If
            
    End Select
    
    'return strResult
    GetWinDirectories = strResult
End Function

Public Function GetFileAssociation(ByVal strExtention As String) _
                                   As String
    'This function will return the registry value associated with a particular file extention. The
    'value returned is for the Open command for the file type and normally holds the file path and
    'a parameter value. If there is no Open command for the file type or nothing is associated with
    'it, then a blank string is returned
    
    Dim strSubKey       As String       'holds a subkey value
    Dim strData         As String       'holds the string data for the named value
    Dim strLink         As String       'holds the name of the subkey which contains the command data
    
    strExtention = Trim(strExtention)
    
    'was an extention specified
    If (strExtention = "") Then
        Exit Function
    End If
    
    'is there a "." at the end of the extention
    If (InStr(1, strExtention, ".") <> 1) Then
        strExtention = "." + strExtention
    End If
    
    'look for this extention in the registry as a subkey
    strSubKey = ReadRegString(HKEY_CLASSES_ROOT, strExtention)
    
    If (strSubKey = "") Or (InStr(1, strSubKey, "Error", vbTextCompare) > 0) Then
        Exit Function
    End If
    
    'we have our sub key. Now we need to drill further down the tree and get the Open command
    strLink = strSubKey + "\Shell\Open\Command"
    strData = ReadRegString(HKEY_CLASSES_ROOT, strLink)
    
    'could we get any details
    If (strData <> "") And (InStr(1, strData, "Error", vbTextCompare) <> 0) Then
        'return the open command string
        GetFileAssociation = strData
    End If
End Function

Public Function GetRegisteredOwner() As String
    'This function will returned the registered
    'strOwner for the local machine.
    
    Const OwnerKeyLoc   As String = "RegisteredOwner"
    
    Dim strOwner        As String
    
    'get the registered gstrOwner
    If IsWinNT Then
        strOwner = ReadRegString(HKEY_LOCAL_MACHINE, _
                                 WIN_NT_INFO_SUBKEY, _
                                 OwnerKeyLoc)
    Else
        strOwner = ReadRegString(HKEY_LOCAL_MACHINE, _
                                 WIN_INFO_SUBKEY, _
                                 OwnerKeyLoc)
    End If
    
    'return lngResult
    GetRegisteredOwner = strOwner
End Function

Public Function ReadRegString(ByVal enmHive As RegistryHives, _
                              ByVal strSubKey As String, _
                              Optional ByVal strEntry As String) _
                              As String
    'This function will check a registery string entry and
    'return the result.
    
    Dim strText         As String
    Dim lngResult       As Long
    Dim hOpenKey        As Long
    Dim lngBufferSize   As Long
    
    'open the registry key
    hOpenKey = GetSubKeyHandle(enmHive, strSubKey)
    
    'check for error
    If hOpenKey = 0 Then
        'return error message
        ReadRegString = "Error : Cannot Open Key"
        Exit Function
    End If
    
    'setup the string to hold the return value
    strText = String(255, vbNullChar)
    lngBufferSize = Len(strText)
    
    'query the information in the key
    lngResult = RegQueryValueEx(hOpenKey, _
                                strEntry, _
                                0, _
                                REG_SZ, _
                                ByVal strText, _
                                lngBufferSize)
    
    'close access to the key
    lngResult = RegCloseKey(hOpenKey)
    
    'check for no values returned
    If (Left(strText, 1) = vbNullChar) Then
        'return error message
        ReadRegString = "Error : Cannot Retrieve String"
        Exit Function
    Else
        'remove the null character
        If (InStr(1, strText, vbNullChar) > 0) Then
            strText = Left(strText, InStr(1, strText, vbNullChar) - 1)
        End If
    End If
    
    'function successful, return owners name
    ReadRegString = strText
End Function

Public Function ReadRegLong(ByVal enmHive As RegistryHives, _
                            ByVal strSubKey As String, _
                            ByVal strEntry As String, _
                            Optional ByVal enmType As RegistryLongTypes = REG_BINARY) _
                            As Long
    'This function will check a registery string
    'entry and return the lngResult.
    
    Dim lngValue        As Long
    Dim lngResult       As Long
    Dim hOpenKey        As Long
    Dim lngBufferSize   As Long
    
    'open the registry key
    hOpenKey = GetSubKeyHandle(enmHive, strSubKey)
    
    'check for error
    If hOpenKey = 0 Then
        'return error message
        ReadRegLong = 0
        Exit Function
    End If
    
    lngBufferSize = 4
    
    'query the information in the key
    lngResult = RegQueryValueEx(hOpenKey, _
                                strEntry, _
                                ByVal 0&, _
                                REG_BINARY, _
                                lngValue, _
                                lngBufferSize)
    
    'close access to the key
    lngResult = RegCloseKey(hOpenKey)
    
    'function successful, return owners name
    ReadRegLong = lngValue
End Function

Private Function GetSubKeyHandle(ByVal enmHive As RegistryHives, _
                                 ByVal strSubKey As String, _
                                 Optional ByVal enmAccess As RegistryKeyAccess = KEY_READ) _
                                 As Long
    'This function returns a handle to the specified registry key
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim hKey        As Long     'holds the handle to the specified key
    
    'open the registry key
    lngResult = RegOpenKeyEx(enmHive, strSubKey, 0, enmAccess, hKey)
    
    If lngResult <> ERROR_SUCCESS Then
        'could not create key
        hKey = 0
    End If
        
    'return value
    GetSubKeyHandle = hKey
End Function

Public Function GetSpace(enmSpaceType As MemType, _
                         Optional ByVal strDrive As String = "C:\") _
                         As Long
    'This function returns the amount of specified memory, either in total
    'or available depending on what was passed.
    'Keep in mind that the information returned is volitile - if you call
    'the function twice, there is no guarentee that the values returned
    'will be the same.
    'Note also, that physical memory is ram memory and memory usage is
    'the amount of ram used.
    
    Const CpuSubKey As String = "PerfStats\StatData"
    Const CpuName   As String = "KERNEL\CPUUsage"
    
    Dim enmMemStruc         As MEMORYSTATUS
    Dim lngResult           As Long
    Dim SecPerCluster       As Long
    Dim lngBytPerSector     As Long
    Dim lngFreeClusters     As Long
    Dim lngTotalClusters    As Long
    
    'Before calling GlobalMemoryStatus, we have to tell it the length
    'of the structure we are passing it - this is required by the procedure.
    enmMemStruc.dwLength = Len(enmMemStruc)
    Call GlobalMemoryStatus(enmMemStruc)
    
    'get the disk space. The function must be passed the root directory of
    'a drive like "C:\" or "D:\" and must end with a Null character (chr(0) )
    If Len(strDrive) >= 3 Then
        lngResult = GetDiskFreeSpace((Left(strDrive, 3) & Chr(0)), _
                                     SecPerCluster, _
                                     lngBytPerSector, _
                                     lngFreeClusters, _
                                     lngTotalClusters)
    End If
    
    'save the selected lngResult
    Select Case enmSpaceType
    
    Case CPUUsage 'cpu usage
        lngResult = ReadRegLong(HKEY_DYN_DATA, CpuSubKey, CpuName)
    
    Case MemoryUsage 'ram usage
        lngResult = enmMemStruc.dwMemoryLoad
    
    Case TotalPhysical 'total ram
        lngResult = enmMemStruc.dwTotalPhys
    
    Case AvailablePhysical 'available ram
        lngResult = enmMemStruc.dwAvailPhys
    
    Case TotalPageFile 'total page file
        lngResult = enmMemStruc.dwTotalPageFile
    
    Case AvailablePageFile 'available page file
        lngResult = enmMemStruc.dwAvailPageFile
    
    Case TotalVirtual 'total virtual (swap file)
        lngResult = enmMemStruc.dwTotalVirtual
    
    Case AvailableVirtual 'available virtual
        lngResult = enmMemStruc.dwAvailVirtual
    
    Case TotalDisk 'hard drive space
        lngResult = lngTotalClusters * (lngBytPerSector * SecPerCluster)
    
    Case AvailableDisk 'available hard drive space
        lngResult = lngFreeClusters * (lngBytPerSector * SecPerCluster)
    
    Case Else
        'return -1 as an error code
        lngResult = -1
    End Select
    
    GetSpace = lngResult
End Function

Public Function GetCountry() As String
    'This will return the country from
    'the computers' regional settings
    
    Const CountryKey        As String = "sCountry"  'the registry entry that holds the country name
    Const DEFAULT_COUNTRY   As String = "Ireland"   'the default country to return if unable to retrieve from the registry
    
    Dim strCountry          As String       'holds the value of the registry entry
    
    strCountry = ReadRegString(HKEY_USERS, _
                               COUNTRY_SUBKEY, _
                               CountryKey)
    
    'if it could not get the country, then default to
    'the programmers country
    If UCase(Left(strCountry, 5)) = "ERROR" Then
        strCountry = DEFAULT_COUNTRY
    End If
    
    'return the country
    GetCountry = strCountry
End Function

Public Function ShellFile(ByVal strFilePath As String, _
                          Optional enmFocus As VbAppWinStyle = vbNormalFocus) _
                          As Long
    'This will open any file with the appropiate program
    'as long as it is registered in the registry and
    'if the function is successful, it will return the
    'applications ID.
    
    Dim strExtention    As String       'holds the file extention
    Dim lngDotPos       As Long         'the position of the last . character found in the string
    Dim lngAppId        As Long         'the process id for the started application
    Dim strWindowsDir   As String       'the location of the windows directory
    Dim strSubKeyLoc    As String       'the location of the registry sub key to open the file type
    Dim strOpenWith     As String       'the program to open the file with
    Dim strMulti()      As String       'the individual files if more than one is passed (multiple parameters)
    Dim intCounter      As Integer      'used to cycle through the file list
    
    'get the windows directory
    strWindowsDir = GetWinDirectories(WindowsDir)
    
    'strip qutoation marks from the file path
    strFilePath = Replace(strFilePath, """", "")
    
    'see if the file is a directory, if so open in
    'explorer
    If HasFileAttrib(strFilePath, vbDirectory) Then
        'open the directory
        lngAppId = Shell(AddToPath(strWindowsDir, _
                                 "Explorer.exe /n,/e," _
                                 & strFilePath), _
                         enmFocus)
        
        ShellFile = lngAppId
        Exit Function
    End If
    
    'get the file extention if any exists (after the last
    'position of the backslash)
    lngDotPos = InStrRev(strFilePath, ".")
    If (lngDotPos > 0) Then
        If (InStr(lngDotPos, strFilePath, "\") = 0) Then
            'file extention exists
            strExtention = Right(strFilePath, _
                                 Len(strFilePath) - _
                                 lngDotPos + 1)
        End If
        
    Else
        'assume that the file is an executable
        strExtention = ".exe"
        strFilePath = strFilePath + ".exe"
    End If  'was an extention specified
    
    'if the extention marks any executable file, then
    'simple run it
    Select Case LCase(strExtention)
    Case ".exe", ".com", ".bat", ""
    
        'make sure the file exists
        If (Dir(strFilePath) <> "") And (Trim(strFilePath) <> "") Then
            lngAppId = Shell(strFilePath, enmFocus)
            
            'return a pointer to the application instance
            ShellFile = lngAppId
        End If
        
        'if no directory was specified, then try to run from either the Windows or System directory
        If (lngAppId = 0) And (InStr(1, strFilePath, "\") = 0) Then
            
            'try to run from the windows directory
            lngAppId = ShellFile(AddToPath(GetWinDirectories(WindowsDir), strFilePath), enmFocus)
            
            'if that didn't work, then try to run from the system directory
            If (lngAppId = 0) Then
                lngAppId = ShellFile(AddToPath(GetWinDirectories(SystemDir), strFilePath), enmFocus)
            End If  'is that didn't work, then try to run from the system directory
        End If  'if not directory was specified and we couldn't run the program "as is", the try default directories
        
        ShellFile = lngAppId
        Exit Function
    End Select
    
    'we need to check the executable file types that
    'can run on their own
    strSubKeyLoc = ReadRegString(HKEY_CLASSES_ROOT, _
                                 strExtention)
    strOpenWith = ReadRegString(HKEY_CLASSES_ROOT, _
                                AddToPath(strSubKeyLoc, _
                                        "shell\open\command"))
    
    'make sure no error was returned
    If UCase(Left(strOpenWith, 5)) = "ERROR" Then
        'couldn't open file
        ShellFile = 0
        Exit Function
    End If
    
    'process the string returned so that we can send
    'it to the Shell function
    If InStr(strOpenWith, "%1") > 0 Then
        'replace the parameters with the appropiate
        'file names
        If InStr(strOpenWith, ",") = 0 Then
            'process one file
            strOpenWith = Replace(strOpenWith, _
                                  "%1", _
                                  strFilePath)
        Else
            'process multiple files
            strMulti = Split(strFilePath, ",")
            
            For intCounter = LBound(strMulti) To UBound(strMulti)
                'replace each parameter string with the
                'corresponding number of elements found
                strOpenWith = Replace(strOpenWith, _
                                      "%" & intCounter, _
                                      strMulti(intCounter))
            Next intCounter
        End If
    Else
        'insert the file name(s) at the end of the
        'name of the program. Please note, that this
        'might not actually work for some programs as
        'the extra parameter may produce an error or be
        'ignored altogether. However this is unlikley
        'as this program path was found in the "Open"
        'section of the program commands.
        strOpenWith = strOpenWith & " " & _
                      Chr(34) & strFilePath & Chr(34)   'chr(34) is a double quote character (")
    End If
    
    'replace system path codes with the actual paths (typically on an NT
    'based machine) --NOT case sensitive with vbTextCompare--
    strOpenWith = Replace(strOpenWith, _
                          "%SystemDrive%", _
                          Left(GetWinDirectories(WindowsDir), 3), _
                          Compare:=vbTextCompare)
    strOpenWith = Replace(strOpenWith, _
                          "%SystemRoot%", _
                          GetWinDirectories(WindowsDir), _
                          Compare:=vbTextCompare)
    
    'open the file
    lngAppId = Shell(strOpenWith, enmFocus)
    ShellFile = lngAppId
End Function

Private Function AddToPath(ByVal strPath As String, _
                           ByVal strFileName As String) _
                           As String
    
    'This function takes a file name and a path and will
    'put the two together to form a filepath. This is useful
    'for when the applications' path happens to be the root
    'directory.
    
    If (strPath = "") Then
        'no path was passed
        AddToPath = strFileName
        Exit Function
    End If
    
    'check the last character for a backslash
    If Left(strPath, 1) = "\" Then
        'don't insert a backslash
        AddToPath = strPath & strFileName
    Else
        'insert a backslash
        AddToPath = strPath & "\" & strFileName
    End If
End Function

Private Function FileExists(ByVal strFilePath As String, _
                            Optional ByVal enmFlags As VbFileAttribute = vbNormal) _
                            As Boolean
    'returns True if the file exists
    
    If ((strFilePath = "") Or _
        (Dir(strFilePath, enmFlags) = "")) Then
        'invalid path/filename
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Function HasFileAttrib(ByVal strFilePath As String, _
                               Optional ByVal enmFlags As VbFileAttribute) _
                               As Boolean
    'returns True if the file specified has the
    'appropiate type signiture, eg, a directory or is
    'read-only. If testing multiple attributes, then
    'the file MUST have all attributes to return True
    
    Dim lngErrNum As Long   'holds any error that occurred trying to access the file
    
    'make sure the file exists without upsetting any
    'stored values when the Dir function is being used
    'externally by another procedure/function
    On Error Resume Next
        'test file access
        GetAttr strFilePath
        lngErrNum = Err
    On Error GoTo 0
    
    'exit if an error occured ("#53 - File Not Found"
    'usually occurs)
    If lngErrNum > 0 Then
        HasFileAttrib = False
        Exit Function
    End If
    
    'test the file for attributes
    If ((GetAttr(strFilePath) And enmFlags) = enmFlags) Then
        HasFileAttrib = True
    Else
        HasFileAttrib = False
    End If
End Function

Private Function IsWinNT() As Boolean
    'Detect if the program is running under an NT based system (NT, 2000, XP)
    
    Const VER_PLATFORM_WIN32_NT     As Long = 2
    
    Dim osiInfo    As OSVERSIONINFO    'holds the operating system information
    Dim lngResult  As Long             'returned error value from the api call
    
    'get version information
    osiInfo.dwOSVersionInfoSize = Len(osiInfo)
    lngResult = GetVersionEx(osiInfo)
    
    'return True if the test of windows NT is positive
    IsWinNT = (osiInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Sub NTMenus(ByVal enmPrivilage As EnumNTSettings, _
                   ByVal blnEnable As Boolean)
    'This will enable or disable the windows task manager. Please note that
    'this procedure does not work on any Non-NT based system (win 9x)
    
    Const CHANGE_PASS   As String = "DisableChangePassword"
    Const LOCK_WORK_ST  As String = "DisableLockWorkStation"
    Const REG_TOOLS     As String = "DisableRegistryTools"
    Const TASK_MANAGER  As String = "DisableTaskMgr"
    'disable parts of the Display dialog box
    Const DISPLAY_PAGE  As String = "NoDispAppearancePage"
    Const DISPLAY_BPAGE As String = "NoDispBackgroundPage"
    Const DISPLAY_CPL   As String = "NoDispCPL"
    Const DISPLAY_SCRSV As String = "NoDispScrSavPage"
    Const DISPLAY_SETT  As String = "NoDispSettingsPage"
    
    Dim strValueName    As String   'holds the Value to open
    Dim lngFlag         As Long     'holds the value to set the setting
    
    If Not IsWinNT Then
        'cannot change settings unless this is a winnt system
        Exit Sub
    End If
    
    'get the text to for the registry value for the selected setting
    Select Case enmPrivilage
        'items that can be disabled on the Lock Screen
    Case CHANGE_PASSWORD
        strValueName = CHANGE_PASS
        
    Case LOCK_WORKSTATION
            strValueName = LOCK_WORK_ST
            
    Case REGISTRY_TOOLS
        strValueName = REG_TOOLS
        
    Case TASK_MGR
        strValueName = TASK_MANAGER
    
        'the tabs on the Display Properties dialog box
    Case DISP_APPEARANCE_PAGE
        strValueName = DISPLAY_PAGE
        
    Case DISP_BACKGROUND_PAGE
        strValueName = DISPLAY_BPAGE
        
    Case DISP_CPL
        strValueName = DISPLAY_CPL
        
    Case DISP_SCREENSAVER
        strValueName = DISPLAY_SCRSV
        
    Case DISP_SETTINGS
        strValueName = DISPLAY_SETT
        
    Case Else
        'invalid selection
        Exit Sub
    End Select
    
    'get the value settings
    If Not blnEnable Then
        'disable option
        lngFlag = 1
    Else
        'enable option
        lngFlag = 0
    End If
    
    If IsWinNT Then
        'NT registry location
        Call CreateRegLong(HKEY_CURRENT_USER, _
                           NT_SETTINGS, _
                           strValueName, _
                           lngFlag)
        
        If IsW2000 Then
            'windows 2000 needs an additional entry
            Call CreateRegLong(HKEY_CURRENT_USER, _
                               W2K_SETTINGS, _
                               strValueName, _
                               lngFlag)
        End If
    End If
End Sub

Public Sub AutoRestartShell(ByVal blnEnable As Boolean)
    'This will turn on/off whether or not the windows shell restarts if it is
    'shutdown or not. This only works on NT based systems
    
    'in registry hive HKEY_LOCAL_MACHINE
    Const AUTO_RESTART_SUBKEY   As String = "Software\Microsoft\Windows NT\" + _
                                            "CurrentVersion\WinLogon"
    
    Dim lngResult   As Long         'holds any returned error value from an api call
    Dim hKey        As Long         'holds a handle to the opened key
    Dim lngData     As Long         'holds the data going into the registry key
    
    'if this is not an NT machine, this won't work
    If Not IsWinNT Then
        Exit Sub
    End If
    
    'get the value of the data going into the registry key
    lngData = Abs(blnEnable)
    
    'set the value to enable or disable the specified setting
    Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                       AUTO_RESTART_SUBKEY, _
                       "AutoRestartShell", _
                       lngData)
End Sub

Public Function IsW2000() As Boolean
    'This will only return True if the version returned by the registry
    'value CurrentVersion is 5
    
    Dim strVersion     As String       'holds the verion number of the operating system
    
    'the the machine NT based (NT, 2000, XP)
    If Not IsWinNT Then
        IsW2000 = False
        Exit Function
    End If
    
    'check the version
    strVersion = ReadRegString(HKEY_LOCAL_MACHINE, _
                               WIN_NT_INFO_SUBKEY, _
                               "CurrentVersion")
    
    'could we read the registry entry
    If Len(strVersion) < 0 Then
        IsW2000 = False
        Exit Function
    End If
    
    'check the version
    If (strVersion = "") Then
        IsW2000 = False
    
    Else
        If Left(strVersion, 1) = "5" Then
            IsW2000 = True
        Else
            IsW2000 = False
        End If
    End If
End Function

Public Function IsXp() As Boolean
    'This will only return True if the version returned by the registry
    'value CurrentVersion is 5.1
    
    Dim strVersion     As String       'holds the verion number of the operating system
    
    'the the machine NT based (NT, 2000, XP)
    If Not IsWinNT Then
        IsXp = False
        Exit Function
    End If
    
    'check the version
    strVersion = ReadRegString(HKEY_LOCAL_MACHINE, _
                               WIN_NT_INFO_SUBKEY, _
                               "CurrentVersion")
    
    'could we read the registry entry
    If Len(strVersion) < 0 Then
        IsXp = False
        Exit Function
    End If
    
    'check the version
    If (strVersion = "") Then
        IsXp = False
    
    Else
        If Left(strVersion, 3) = "5.1" Then
            IsXp = True
        Else
            IsXp = False
        End If
    End If
End Function

Public Sub OppLocking(ByVal blnEnable As Boolean)
    'This will enable or disable oppertunistic locking on an NT based machine
    
    'in HKEY_LOCAL_MACHINE registry hive
    Const LOCK_OP_SUBKEY    As String = "System\CurrentControlSet\Services"
    Const W2K_lOCK_LOCAL    As String = LOCK_OP_SUBKEY + "\LanManServer\Parameters"
    Const W2K_LOCK_REMOTE   As String = LOCK_OP_SUBKEY + "\MrxSmb\Parameters"
    Const WNT_LOCK_LOCAL    As String = LOCK_OP_SUBKEY + "\LanManWorkStation\Parameters"
    Const WNT_LOCK_REMOTE   As String = LOCK_OP_SUBKEY + "\LanManServer\Parameters"
    
    Dim lngData             As Long     'holds the numeric value to set to
    
    'make sure we are running on an NT based system
    If Not IsWinNT Then
        Exit Sub
    End If
    
    'what kind of NT based system are we running on
    If IsW2000 Then
        'enable/disable opportunistic locking on windows 2000
        lngData = Abs(blnEnable)
        
        'local locking
        Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                           W2K_lOCK_LOCAL, _
                           "EnableOpLocks", _
                           lngData)
        
        'remote locking
        lngData = Abs(Not blnEnable)
        
        Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                           W2K_LOCK_REMOTE, _
                           "OplocksDisabled", _
                           lngData)
    
    Else
        'enable/disable opportunistic locking on windows NT
        
        lngData = Abs(blnEnable)
        
        'local locking
        Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                           WNT_LOCK_LOCAL, _
                           "UseOpportunisticLocking", _
                           lngData)
        
        'remote locking
        Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                           WNT_LOCK_REMOTE, _
                           "EnableOpLocks", _
                           lngData)
    End If
End Sub

Public Sub CreateRegLong(ByVal enmHive As RegistryHives, _
                         ByVal strSubKey As String, _
                         ByVal strValueName As String, _
                         ByVal lngData As Long, _
                         Optional ByVal enmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)
    'This will create a value in the registry of the specified type
    'and value data
    
    Dim hKey        As Long     'holds a pointer to an open registry key
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    'make sure the registry value exists
    Call CreateSubKey(enmHive, strSubKey)
    
    'open the subkey
    hKey = GetSubKeyHandle(enmHive, strSubKey, KEY_ALL_ACCESS) ' KEY_SET_VALUE)
    
    'create the registry value
    lngResult = RegSetValueEx(hKey, _
                              strValueName, _
                              0, _
                              enmType, _
                              lngData, _
                              4)
    
    'close the registry key
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub OpenVbIdeMaximized(ByVal blnEnable As Boolean)
    'This will set the vb ide to open projects maximized by default
    
    'HKEY_CURRENT_USER
    Const VB_IDE_SUB_KEY    As String = "Software\Microsoft\Visual Basic\6.0"
    
    Call CreateRegString(HKEY_CURRENT_USER, _
                         VB_IDE_SUB_KEY, _
                         "MDIMaximized", _
                         Trim(Str(Abs(blnEnable))))
End Sub

Public Sub SaveArray(ByRef varArray() As Variant, _
                     ByVal enmHive As RegistryHives, _
                     ByVal strSubKey As String, _
                     Optional ByVal strArrayName As String = "VB6_Array", _
                     Optional ByVal enmDataType As RegistryDataTypes = REG_DT_SZ)
    'This will save an array of the specified data type to the specified
    'registry sub key. The array must be initialised and valid for the
    'data type specified as there is no checking done to validate the data.
    
    Dim lngCounter      As Long         'used to cycle through the array specified
    Dim lngMin          As Long         'holds the lower bound of the array
    Dim lngMax          As Long         'holds the upper bound of the array
    
    'make sure that a valid subkey was passed
    If (Trim(strSubKey) = "") Then
        Exit Sub
    End If
    
    'make sure that the sub key exists in the registry
    Call CreateSubKey(enmHive, strSubKey)
    
    'get the size of the array
    lngMin = LBound(varArray)
    lngMax = UBound(varArray)
    
    'save the bounds in the specified key
    Call CreateRegLong(enmHive, _
                       strSubKey, _
                       (strArrayName + "LBound"), _
                       lngMin, _
                       REG_BINARY)
    Call CreateRegLong(enmHive, _
                       strSubKey, _
                       (strArrayName + "UBound"), _
                       lngMax, _
                       REG_BINARY)
    
    'save the elements of the array to the registry
    For lngCounter = lngMin To lngMax
        If (enmDataType = REG_DT_SZ) Then
            'save as string
            Call CreateRegString(enmHive, _
                                 strSubKey, _
                                 (strArrayName & lngCounter), _
                                 varArray(lngCounter))
            
        Else
            'save as numeric
            Call CreateRegLong(enmHive, _
                               strSubKey, _
                               (strArrayName & lngCounter), _
                               varArray(lngCounter), _
                               enmDataType)
        End If
    Next lngCounter
End Sub

Public Sub LoadArray(ByRef varArray() As Variant, _
                     ByVal enmHive As RegistryHives, _
                     ByVal strSubKey As String, _
                     Optional ByVal strArrayName As String = "VB6_Array", _
                     Optional ByVal enmDataType As RegistryDataTypes = REG_DT_SZ)
    'This will load an array saved with the SaveArray procedure above. The
    'data must have been saved using the correct data and datatypes. The array
    'passed to this procedure will be wiped, resized and loaded with whatever
    'information can be retrieved from the registry. It is up to the programmer
    'to ensure that the correct data types are passed to the procedure or the
    'information returned may be corrupt if any information is returned at all.
    
    Dim lngCounter      As Long         'used to cycle through the array specified
    Dim lngMin          As Long         'holds the lower bound of the array
    Dim lngMax          As Long         'holds the upper bound of the array
    
    'make sure that the correct sub key was passed
    If (Trim(strSubKey) = "") Then
        Exit Sub
    End If
    
    'get the size of the array
    lngMin = ReadRegLong(enmHive, _
                         strSubKey, _
                         (strArrayName + "LBound"), _
                         REG_BINARY)
    lngMax = ReadRegLong(enmHive, _
                         strSubKey, _
                         (strArrayName + "UBound"), _
                         REG_BINARY)
    
    'resize the array to accomidate the data
    ReDim varArray(lngMin To lngMax)
    
    For lngCounter = lngMin To lngMax
        If (enmDataType = REG_DT_SZ) Then
            'read string data into the array
            varArray(lngCounter) = ReadRegString(enmHive, _
                                                 strSubKey, _
                                                 (strArrayName & lngCounter))
        
        Else
            'read numeric data into the array
            varArray(lngCounter) = ReadRegLong(enmHive, _
                                               strSubKey, _
                                               (strArrayName & lngCounter), _
                                               enmDataType)
        End If
    Next lngCounter
End Sub

Public Function SearchValues(ByVal enmHive As RegistryHives, _
                             ByVal strSubKey As String, _
                             Optional ByVal strPattern As String = "", _
                             Optional ByVal blnSearchSubKeys As Boolean = True, _
                             Optional ByRef lngFound As Long = 0, _
                             Optional ByRef blnSuccessful As Boolean) _
                             As TypeRegValues()
    'This will return the details of any found value names matching the
    'specified pattern. If blnSearchSubKeys is set to True, then sub
    'keys are also searched
    
    Const BUFFER_SIZE   As Long = 255           'holds the size of the buffer in characters which will contain the value name
    
    Dim lngResult       As Long                 'holds the return value of any api call
    Dim udtValueInfo    As TypeRegValues        'holds information about the value
    Dim udtFoundHere()  As TypeRegValues        'holds any values found from a single call
    Dim udtFoundTot()   As TypeRegValues        'holds all the value information found from all the calls so far
    Dim lngNumFound     As Long                 'holds the number of values found from a single sub key
    Dim lngTotFound     As Long                 'holds the total number of values found so far
    Dim strKeys()       As String               'holds the number of keys in this sub key
    Dim hSubKey         As Long                 'holds a handle to the sub key
    Dim strValueName    As String * BUFFER_SIZE 'holds the name of the value that was found
    Dim lngIndex        As Long                 'holds the index of the registry value
    Dim lngType         As Long                 'holds the data type of the value
    Dim strTempName     As String               'holds the value name without being padded out with null characters
    Dim lngCounter      As Long                 'used for cycling through the results when adding to the end of the final results array
    Dim lngKeyCounter   As Long                 'used for cycling through the list of sub keys returned
    Dim lngKeysFound    As Long                 'holds the number of sub keys found in the specified sub key
    Dim lngFoundPrev    As Long                 'holds the number of values found before we add more elements
    Dim varData         As Variant              'holds the data for the value
    
    'reset the parameters specified
    lngFound = 0
    blnSuccessful = 0
    lngNumFound = 0
    lngTotFound = 0
    ReDim udtFoundHere(lngNumFound)
    ReDim udtFoundTot(lngTotFound)
    
    hSubKey = GetSubKeyHandle(enmHive, strSubKey)
    If (hSubKey = 0) Then
        'this sub key does not exist
        Exit Function
    End If
    
    'get the values of the specified sub key
    lngIndex = 0
    Do
        'get the value info
        strValueName = String(BUFFER_SIZE, vbNullChar)
        lngType = 0
        lngResult = RegEnumValue(hSubKey, _
                                 lngIndex, _
                                 strValueName, _
                                 BUFFER_SIZE, _
                                 0&, _
                                 lngType, _
                                 ByVal varData, _
                                 0&)
        
        'was data returned
        If (Left(strValueName, 1) <> vbNullChar) Then
            strTempName = Mid(strValueName, 1, InStr(1, strValueName, vbNullChar) - 1)
            
            'does this value name match the specified pattern
            If (strTempName = "") Or (strTempName Like strPattern) Then
                'store the name
                ReDim Preserve udtFoundHere(lngNumFound)
                With udtFoundHere(lngNumFound)
                    .enmHive = enmHive
                    .strSubKey = strSubKey
                    .strName = strTempName
                    .lngType = lngType
                    .varData = varData
                End With    'udtFoundHere(lngNumFound)
                lngNumFound = lngNumFound + 1
            End If  'does this value name match the specified pattern
        End If  'was data returned
        
        'get the next values' details
        lngIndex = lngIndex + 1
    Loop Until (lngResult = ERROR_NO_MORE_ITEMS)
    
    'add these details to the total results array
    lngTotFound = lngNumFound
    If (lngTotFound > 0) Then
        ReDim udtFoundTot(lngTotFound - 1)
        For lngCounter = 0 To (lngTotFound - 1)
            'copy data
            udtFoundTot(lngCounter) = udtFoundHere(lngCounter)
        Next lngCounter
    End If
    
    'are we searching subkeys
    If blnSearchSubKeys Then
        strKeys = GetSubKeys(enmHive, strSubKey, lngKeysFound)
        
        For lngKeyCounter = 0 To (lngKeysFound - 1)
            'get the results of each sub key
            udtFoundHere() = SearchValues(enmHive, _
                                          AddToPath(strSubKey, strKeys(lngKeyCounter)), _
                                          strPattern, _
                                          True, _
                                          lngNumFound)
            
            'were any values found
            If (lngNumFound > 0) Then
                'copy the results of the search to the list of currently found
                lngFoundPrev = lngTotFound
                lngTotFound = lngTotFound + lngNumFound
                ReDim Preserve udtFoundTot(lngTotFound - 1)
                For lngCounter = (lngFoundPrev) To (lngTotFound - 1)
                    'copy data
                    udtFoundTot(lngCounter) = udtFoundHere(lngCounter - lngFoundPrev)
                Next lngCounter
            End If  'were any values found
        Next lngKeyCounter
    End If
    
    'return the results
    blnSuccessful = True
    lngFound = lngTotFound
    SearchValues = udtFoundTot()
End Function

Public Sub SetNumLock(Optional ByVal blnTurnOn As Boolean = True)
    'This will turn the numlock on or off when logging in to Nt/2000/XP
    
    Const NUMLOCK_SUBKEY    As String = "Control Panel\Keyboard"    'HKEY_CURRENT_USER
    Const NUMLOCK_VALUE     As String = "InitialKeyboardIndicators"
    
    Dim strOnText  As String   'holds the actual string value that turns the numlock on or off
    
    If Not IsWinNT Then
        'this won't work on a non-nt based system
        Exit Sub
    End If
    
    If blnTurnOn Then
        strOnText = "2"    'on
    Else
        strOnText = "0"   'off
    End If
    
    Call CreateRegString(HKEY_CURRENT_USER, _
                         NUMLOCK_SUBKEY, _
                         NUMLOCK_VALUE, _
                         strOnText)
End Sub

Public Property Let CdAutoRun(ByVal blnOn As Boolean)
    'This will turn on/off the cd autorun
    
    Const AUTORUN_SUBKEY    As String = "System\CurrentControlSet\Services\Cdrom"
    Const AUTORUN_VALUE     As String = "AutoRun"
    
    'turn autorun on/off
    Call CreateRegLong(HKEY_LOCAL_MACHINE, _
                       AUTORUN_SUBKEY, _
                       AUTORUN_VALUE, _
                       CInt(Abs(blnOn)))
End Property

Public Property Get CdAutoRun() As Boolean
    'This will return if the cd auto run is turned on
    
    Const AUTORUN_SUBKEY    As String = "System\CurrentControlSet\Services\Cdrom"
    Const AUTORUN_VALUE     As String = "AutoRun"
    
    'is the autorun on
    CdAutoRun = ReadRegLong(HKEY_LOCAL_MACHINE, _
                            AUTORUN_SUBKEY, _
                            AUTORUN_VALUE)
End Property

Public Property Let DisplayVersionOnDesktop(ByVal blnOn As Boolean)
    'This will turn on/off displaying the windows version and build
    'number on the bottom right of the wallpaper.
    
    Const BUILD_SUBKEY      As String = "Control Panel\Desktop"
    Const BUILD_VALUE       As String = "PaintDesktopVersion"
    
    'turn off version number
    Call CreateRegLong(HKEY_CURRENT_USER, _
                       BUILD_SUBKEY, _
                       BUILD_VALUE, _
                       CInt(Abs(blnOn)))
End Property

Public Property Get DisplayVersionOnDesktop() As Boolean
    'This will return True if the Version number is being displayed on
    'the desktop
    
    Const BUILD_SUBKEY      As String = "Control Panel\Desktop"
    Const BUILD_VALUE       As String = "PaintDesktopVersion"
    
    'see if it is currently displayed
    DisplayVersionOnDesktop = ReadRegLong(HKEY_CURRENT_USER, _
                                          BUILD_SUBKEY, _
                                          BUILD_VALUE)
End Property

Public Property Let TaskbarStartSpeed(ByVal lngSpeed As Long)
    'This will set the startup speed of the task bar
    
    Const TASKBAR_SUBKEY    As String = "Control Panel\Desktop"
    Const TASKBAR_VALUE     As String = "MenuShowDelay"
    
    'set the speed
    Call CreateRegString(HKEY_CURRENT_USER, _
                         TASKBAR_SUBKEY, _
                         TASKBAR_VALUE, _
                         Trim(Str(lngSpeed)))
End Property

Public Property Get TaskbarStartSpeed() As Long
    'This will return teh startup speed of the task bar
    
    Const TASKBAR_SUBKEY    As String = "Control Panel\Desktop"
    Const TASKBAR_VALUE     As String = "MenuShowDelay"
    
    'get the speed
    TaskbarStartSpeed = CLng(ReadRegString(HKEY_CURRENT_USER, _
                                           TASKBAR_SUBKEY, _
                                           TASKBAR_VALUE))
End Property

Public Function GetWindowsCdKey() As String
    'This will return the windows cd key entered when the user first
    'started the machine. This is usefull in case you ever need to do a
    'backup.
    
    Const WIN_SUBKEY        As String = "Software\Microsoft"
    Const WIN_VALUE         As String = "ProductID"
    
    Dim strWinSubkey        As String       'holds the complete subkey to the ProuctId value
    
    'build the complete subkey path based on the os
    If IsWinNT Then
        strWinSubkey = AddToPath(WIN_SUBKEY, "Windows NT\CurrentVersion")
    Else
        strWinSubkey = AddToPath(WIN_SUBKEY, "Windows\CurrentVersion")
    End If
    
    'get the product key
    GetWindowsCdKey = ReadRegString(HKEY_LOCAL_MACHINE, _
                                    strWinSubkey, _
                                    WIN_VALUE)
End Function
