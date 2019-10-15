Attribute VB_Name = "modSysTrayIcon"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     4 June 2003
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    System Tray Icon module
' -----------------------------------------------
'COMMENTS : This is used to extend the
'   functionality so that whey the task bar is
'   recreated, the system tray icon will be to.
'   To do this, a windows hook has to be created
'   to process the message, and vb only allows
'   those messages to be processed in modules
'   because of the nature of the AddressOf
'   operator.
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'this will move memory in ram from one place to
'another
Private Declare Function CopyMemory _
        Lib "kernel32" _
        Alias "RtlMoveMemory" _
            (ByVal Dest As Long, _
             ByVal Src As Long, _
             ByVal Length As Long) _
             As Long

'calls the default window procedure to provide default
'processing for any window messages that an application
'does not process. This function ensures that every
'message is processed. DefWindowProc is called with the
'same parameters received by the window procedure.
Private Declare Function DefWindowProc _
        Lib "user32" _
        Alias "DefWindowProcA" _
            (ByVal hwnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             ByVal lParam As Long) _
             As Long

'this will get the property value for the
'specified window
Private Declare Function GetWindowLong _
    Lib "user32" _
    Alias "GetWindowLongA" _
        (ByVal hwnd As Long, _
         ByVal nIndex As Long) _
         As Long

'get the message id of a specific message to look for
Private Declare Function RegisterWindowMessage _
        Lib "user32" _
        Alias "RegisterWindowMessageA" _
            (ByVal lpString As String) _
             As Long

'------------------------------------------------
'            MODULE LEVEL CONSTANTS
'------------------------------------------------
'used to pick up the appropiate events
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_RBUTTONDBLCLK  As Long = &H206
Private Const WM_MBUTTONDOWN    As Long = &H207
Private Const WM_MBUTTONUP      As Long = &H208
Private Const WM_MBUTTONDBLCLK  As Long = &H209
Private Const WM_USER           As Long = &H400
Private Const WM_USER_SYSTRAY   As Long = &H405
Private Const WM_CLOSE          As Long = &H10
Private Const GWL_USERDATA      As Long = (-21)

'------------------------------------------------
'            MODULE LEVEL VARIABLES
'------------------------------------------------
Private mlngTaskbarMsg      As Long             'holds the message value to look for
Private mblnStarted         As Boolean          'flags whether or not the message is being monitored
Private msysIcons()         As clsSysTrayIcon   'holds a list of system tray icons to notify

'------------------------------------------------
'                  PROCEDURES
'------------------------------------------------
'
'Public Sub InitTaskbarMsg(ByVal sysTrayIcon As clsSysTrayIcon)
'    'This will initialise the message value to look for. This varies every
'    'time windows starts, so this needs to be called before the window
'    'hook is created.
'
'    If sysTrayIcon Is Nothing Then
'        'invalid object
'        Debug.Print "Unable to create system tray message notification"
'        Exit Sub
'    End If
'
'    'don't do this more than once
'    If Not mblnStarted Then
'        mlngTaskbarMsg = RegisterWindowMessage("TaskbarCreated" + vbNullChar)
'        mblnStarted = True
'        ReDim msysIcons(0)
'    End If
'
'    'add the system tray icon object to the array
'    If Not msysIcons(0) Is Nothing Then
'        'add to end of array
'        ReDim msysIcons(UBound(msysIcons) + 1)
'        Set msysIcons(UBound(msysIcons)) = sysTrayIcon
'
'    Else
'        'enter at start of array
'        Set msysIcons(0) = sysTrayIcon
'    End If
'End Sub

'-------------------------------------------------------------------------
'Dummy function to allow AddressOf to assign to a variable
Public Function Pass(lngNum As Long) As Long
    'return the value passed
    Pass = lngNum
End Function

Public Function DeRef(ByRef lngPointer As Long) _
                      As clsSysTrayIcon
    'Return a VB object pointed by lngPointer
    
    Dim lngResult   As Long     'holds any error value returned from an api call
    
    'return a reference to the object pionted to by the pointer
    lngResult = CopyMemory(VarPtr(DeRef), VarPtr(lngPointer), 4)
End Function

Public Function CreateRef(ByRef sysObject As clsSysTrayIcon) _
                          As Long
    'Creates a pointer to a VB object
    
    Dim lngResult   As Long     'holds any error value returned from an api call
    
    'return the pointer to the object
    lngResult = CopyMemory(VarPtr(CreateRef), VarPtr(sysObject), 4)
End Function

Public Sub DestroyRef(ByRef sysObject As Long)
    'Destroys a VB object created by DeRef (otherwise the VB's
    'reference count would be incorrect)
    
    Dim lngNum      As Long     'just holds zero - the pointer to this variable is used to wipe an object pointer
    Dim lngResult   As Long     'holds any error value returned from an api call
    
    'overwrite the object pointer with a pointer to zero
    lngResult = CopyMemory(sysObject, VarPtr(lngNum), 4)
End Sub

Public Function InTrayWndProc(ByVal lnghWnd As Long, _
                              ByVal lngMsg As Long, _
                              ByVal lngwParam As Long, _
                              ByVal lnglParam As Long) _
                              As Long
    'The window procedure for the dummy windows that clsSysTrayIcon creates
    
    Dim lngObjPointer   As Long
    Dim sysObject       As clsSysTrayIcon
    
    'only do this once
    If Not mblnStarted Then
        InitMessage
    End If
    
    Select Case lngMsg
        'Pass WM_USER_SYSTRAY to the clsSysTrayIcon object
        Case WM_USER_SYSTRAY
            lngObjPointer = GetWindowLong(lnghWnd, GWL_USERDATA)
            Set sysObject = DeRef(lngObjPointer)
            Call sysObject.ProcessMessage(lngwParam, lnglParam)
            Call DestroyRef(VarPtr(sysObject))
        
        'If the TaskBar restarts, let clsSysTrayIcon know about it
        Case mlngTaskbarMsg
            lngObjPointer = GetWindowLong(lnghWnd, GWL_USERDATA)
            Set sysObject = DeRef(lngObjPointer)
            
            'only by passing this string will the event be triggered
            Call sysObject.TriggerTaskbarEvent("modSysTrayIcon")
            Call DestroyRef(VarPtr(sysObject))
    End Select
    
    InTrayWndProc = DefWindowProc(lnghWnd, lngMsg, lngwParam, lnglParam)
End Function

'Register the windows message TaskbarCreated so we can watch for it
Private Function InitMessage()
    
    mblnStarted = True
    mlngTaskbarMsg = RegisterWindowMessage("TaskbarCreated")
End Function
