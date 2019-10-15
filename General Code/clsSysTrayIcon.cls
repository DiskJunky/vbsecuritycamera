VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Allows a program to set an icon in the system tray and triggers the appropiate events necessary. This class needs a picture box with the icon already loaded into it to work properly"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     2 August 2002
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    System Tray Icon Class
' -----------------------------------------------
'COMMENTS :
'Used to display an icon in the system tray and
'trigger appropiate events in a picture box
'instead of the rather obscure way that is
'demonstrated when using procedural code to
'implement this. Creating an object also has the
'added benifit of cleaning up after itself at the
'appropiate time without the programmer worrying
'about when to implement the clean up.
'
'This module has a dependancey on "stdole.dll" for
'the IPictureDisp object used to display the icon.
' -----------------------------------------------
'NOTE :
'The parts of this code that relate to processing
'the messages from the icon were obtained from
'http://www.scottandmichelle.net/scott/code/about.html
'=================================================

'require variable declaration
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------
'creates a window with the specified parameters
Private Declare Function CreateWindowEx _
        Lib "user32" _
        Alias "CreateWindowExA" _
            (ByVal dwExStyle As Long, _
             ByVal lpClassName As String, _
             ByVal lpWindowName As String, _
             ByVal dwStyle As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hWndParent As Long, _
             ByVal hMenu As Long, _
             ByVal hInstance As Long, _
             lpParam As Any) _
             As Long

'gets the position of the cursor
Private Declare Function GetCursorPos _
        Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'this will register the specified class to the
'system
Private Declare Function RegisterClass _
        Lib "user32" _
        Alias "RegisterClassA" _
            (Class As WNDCLASS) _
             As Long
'get the message id of a specific message to look for
Private Declare Function RegisterWindowMessage _
        Lib "user32" _
        Alias "RegisterWindowMessageA" _
            (ByVal lpString As String) _
             As Long

'this will set a message to the specified window
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
            (ByVal hwnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             lParam As Any) _
             As Long

'this will set the top most window
Private Declare Function SetForegroundWindow _
        Lib "user32" _
            (ByVal hwnd As Long) _
             As Long

'this will set a property of the specified window
Private Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" _
            (ByVal hwnd As Long, _
             ByVal nIndex As Long, _
             ByVal dwNewLong As Long) _
             As Long

'sets some system tray icon data
Private Declare Function Shell_NotifyIcon _
        Lib "shell32" _
        Alias "Shell_NotifyIconA" _
            (ByVal dwMessage As Long, _
             pnid As NotifyIconData) _
             As Long

'this will unregister a registered class
Private Declare Function UnregisterClass _
        Lib "user32" _
        Alias "UnregisterClassA" _
            (ByVal lpClassName As String, _
             ByVal hInstance As Long) _
             As Long

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
'holds the icon's information when in the system tray
Private Type NotifyIconData
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    UCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

'used to get the cursors position on the screen
Private Type PointAPI
    X                   As Long
    Y                   As Long
End Type

'holds information about a window class
Private Type WNDCLASS
    style               As Long
    lpfnwndproc         As Long
    cbClsextra          As Long
    cbWndExtra2         As Long
    hInstance           As Long
    hIcon               As Long
    hCursor             As Long
    hbrBackground       As Long
    lpszMenuName        As Long
    lpszClassName       As String
End Type

'------------------------------------------------
'            MODULE-LEVEL CONSTANTS
'------------------------------------------------
'used to flag certain information about the icon
Private Const NIM_ADD           As Long = &H0
Private Const NIM_MODIFY        As Long = &H1
Private Const NIM_DELETE        As Long = &H2
Private Const NIF_MESSAGE       As Long = &H1
Private Const NIF_ICON          As Long = &H2
Private Const NIF_TIP           As Long = &H4

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
'             MODULE-LEVEL VARIABLES
'------------------------------------------------
Private WithEvents mpicHook     As PictureBox       'used to trigger the appropiate events when necessary
Attribute mpicHook.VB_VarHelpID = -1
Private micoIcon                As IPictureDisp     'holds a reference to the icon displayed
Private mstrToolTip             As String           'holds the tool tip for the icon
Private mudtSysTrayDetails      As NotifyIconData   'holds information about the current system tray icon
Private mblnIconDisplayed       As Boolean          'a flag to tell if the icon is currently in the system tray or not
Private mlnghWndHidden          As Long             'holds a reference to a hidden window that we can use to capture the system tray icon events
Private mstrClass               As String           'holds a string name of the window

'------------------------------------------------
'                     EVENTS
'------------------------------------------------
Public Event Click(ByVal Button As Integer, _
                   ByVal Shift As Integer, _
                   ByVal X As Integer, _
                   ByVal Y As Integer)
Public Event DblClick(ByVal Button As Integer, _
                      ByVal Shift As Integer, _
                      ByVal X As Integer, _
                      ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)
Public Event MouseMove(ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)
Public Event MouseUp(ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer)
Public Event TaskbarCreated()

'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Property Set PictureBox(ByVal picNew As PictureBox)
    'sets the picture box object
    
    Dim lngResult As Long   'holds any error value returned from the api call
    
    Set mpicHook = picNew

    'update the icon if necessary with the new Hook
    If mblnIconDisplayed Then
        mudtSysTrayDetails.hwnd = picNew.hwnd
        lngResult = Shell_NotifyIcon(NIM_MODIFY, _
                                     mudtSysTrayDetails)
    End If
End Property

Public Property Get PictureBox() As PictureBox
    'returns a rference to the picture box object
    Set PictureBox = mpicHook
End Property

Public Property Set Icon(ByRef icoIcon As IPictureDisp)
    'sets the icon to be displayed in the system tray
    
    Dim lngResult As Long   'holds any error value returned from the api call
    
    Set micoIcon = icoIcon
    
    'update the icon if necessary with the new icon
    If mblnIconDisplayed Then
        mudtSysTrayDetails.hIcon = CLng(icoIcon)
        lngResult = Shell_NotifyIcon(NIM_MODIFY, _
                                     mudtSysTrayDetails)
    End If
End Property

Public Property Get Icon() As IPictureDisp
    'returns a reference to the Icon property
    Set Icon = micoIcon
End Property

Public Property Let ToolTip(ByVal strNew As String)
    'set the tool tip for the system tray icon
    
    Dim lngResult As Long   'holds any returned error value from the api call
    
    mstrToolTip = strNew
    
    'update the icon if necessary with the new tooltip
    If mblnIconDisplayed Then
        mudtSysTrayDetails.szTip = strNew & vbNullChar
        lngResult = Shell_NotifyIcon(NIM_MODIFY, _
                                     mudtSysTrayDetails)
    End If
End Property

Public Property Get ToolTip() As String
    'returns the tool tip for the icon
    ToolTip = mstrToolTip
End Property

Public Sub ShowIcon()
Attribute ShowIcon.VB_UserMemId = 0
    'This will display the icon in the system tray if
    'the appropiate properties have been set and are
    'valid
    
    'unload any previous icon
    Call UnloadIcon
    
    'in case the objects have not been initialised.
    If (Not mpicHook Is Nothing) And _
       (Not micoIcon Is Nothing) Then
        'display the icon
        Call LoadIcon(mpicHook.hwnd, _
                      CLng(micoIcon))
    End If
End Sub

Private Sub LoadIcon(ByVal hWndHook As Long, _
                     ByVal lngIconProperty As Long)
    'This will load an icon into the system tray
    
    Dim lngResult       As Long     'holds any returned error value from the api call
    Dim udtWnd          As WNDCLASS 'holds a list of the window details. This will be used to process any messages posted
    Dim lngMePointer    As Long     'holds a pointer to the current instance of this object
    
    'only create a new icon if one is not already in
    'memory
    If (mudtSysTrayDetails.hwnd <> 0) Or _
       (hWndHook = 0) Or _
       (lngIconProperty = 0) Then
        
        Exit Sub
    End If
    
    'create a hidden window so that we can process any messages coming
    'through
    With udtWnd
        'setup the window details
        .style = 0
        .lpfnwndproc = Pass(AddressOf InTrayWndProc)
        .hInstance = App.hInstance
        .lpszClassName = mstrClass
    End With    'udtWnd
    
    'create the window
    lngResult = RegisterClass(udtWnd)
    mlnghWndHidden = CreateWindowEx(0, _
                                    mstrClass, _
                                    "VB SystemTrayIcon Windows Hidden" & _
                                    "n Window", 0, 0, 0, 100, _
                                    100, 0, 0, 0, ByVal 0)
    
    'associate this object with the window that we created
    lngMePointer = CreateRef(Me)
    lngResult = SetWindowLong(mlnghWndHidden, GWL_USERDATA, lngMePointer)
    
    'set systray icon details
    With mudtSysTrayDetails
        .cbSize = Len(mudtSysTrayDetails)
        .hwnd = hWndHook
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .UCallBackMessage = WM_MOUSEMOVE    'all events to be triggered here
        .hIcon = lngIconProperty
        .szTip = mstrToolTip & vbNullChar
    End With
    
    'send details
    lngResult = Shell_NotifyIcon(NIM_ADD, _
                                 mudtSysTrayDetails)
    
    If (lngResult) Then
        mblnIconDisplayed = True
    Else
        mblnIconDisplayed = False
    End If
End Sub

Public Sub UnloadIcon()
    'remove the icon from the systray
    
    Dim lngResult As Long   'holds any returned error value from the api call
    
    If Not mblnIconDisplayed Then
        'there is no icon to remove
        Exit Sub
    End If
    
    'remove the icon from the system tray
    lngResult = Shell_NotifyIcon(NIM_DELETE, _
                                 mudtSysTrayDetails)
    mudtSysTrayDetails.hwnd = 0
    mblnIconDisplayed = False

    'Close the dummy window
    SendMessage mlnghWndHidden, WM_CLOSE, 0, ByVal 0&
    UnregisterClass mstrClass, App.hInstance
End Sub

Public Sub ProcessMessage(wParam As Long, lParam As Long)
    'Function should only be called from modInTray to process windows
    ' messages generated from the System Tray
    
    Dim udtPos      As PointAPI     'holds the position of the cursor
    Dim lngResult   As Long         'holds any returned error value from an api call
    
    'get the position of the mouse
    lngResult = GetCursorPos(udtPos)
    
    Select Case lParam
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove(0, 0, udtPos.X, udtPos.Y)
        
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton, 0, udtPos.X, udtPos.Y)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton, 0, udtPos.X, udtPos.Y)
        
        Case WM_LBUTTONDBLCLK
            RaiseEvent DblClick(vbLeftButton, 0, udtPos.X, udtPos.Y)
        
        Case WM_MBUTTONDBLCLK
            RaiseEvent DblClick(vbMiddleButton, 0, udtPos.X, udtPos.Y)
        
        Case WM_RBUTTONDBLCLK
            RaiseEvent DblClick(vbRightButton, 0, udtPos.X, udtPos.Y)
        
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton, 0, udtPos.X, udtPos.Y)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton, 0, udtPos.X, udtPos.Y)
        
        Case WM_MBUTTONDOWN
            RaiseEvent MouseDown(vbMiddleButton, 0, udtPos.X, udtPos.Y)
        Case WM_MBUTTONUP
            RaiseEvent MouseUp(vbMiddleButton, 0, udtPos.X, udtPos.Y)
    End Select
End Sub

Public Sub TriggerTaskbarEvent(Optional strReserved As String)
    'This will trigger the task bar event. This should only be used by
    'modSysTrayIcon
    If (strReserved = "modSysTrayIcon") Then
        RaiseEvent TaskbarCreated
    End If
End Sub

Private Sub mpicHook_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    'used to trigger the appropiate events rather than
    'letting the programmer squash all the events into
    'a picture boxes MouseMove event
    
    'this code is not used any more to process events
    'from the icon
    'Exit Sub
    
    Dim lngMousePos As Long     'detect the appropiate events
    Dim udtPos      As PointAPI 'get the position of the mouse (screen pixel co-ordinates only)
    Dim lngResult   As Long     'holds any returned error value from the api call

    'find out what event was triggered
    lngMousePos = X / Screen.TwipsPerPixelX

    'get the current position of the mouse in screen
    'co-ordinates
    lngResult = GetCursorPos(udtPos)

    Select Case lngMousePos
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton, _
                                 Shift, _
                                 CInt(udtPos.X), _
                                 CInt(udtPos.Y))

        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton, _
                                 Shift, _
                                 CInt(udtPos.X), _
                                 CInt(udtPos.Y))
        
        Case WM_MBUTTONDOWN
            RaiseEvent MouseDown(vbMiddleButton, _
                                 Shift, _
                                 CInt(udtPos.X), _
                                 CInt(udtPos.Y))
        
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton, _
                               Shift, _
                               CInt(udtPos.X), _
                               CInt(udtPos.Y))
            DoEvents
            RaiseEvent Click(vbLeftButton, _
                             Shift, _
                             CInt(udtPos.X), _
                             CInt(udtPos.Y))

        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton, _
                               Shift, _
                               CInt(udtPos.X), _
                               CInt(udtPos.Y))
            DoEvents
            RaiseEvent Click(vbRightButton, _
                             Shift, _
                             CInt(udtPos.X), _
                             CInt(udtPos.Y))
        
        Case WM_MBUTTONUP
            RaiseEvent MouseUp(vbMiddleButton, _
                               Shift, _
                               CInt(udtPos.X), _
                               CInt(udtPos.Y))
            
            DoEvents
            RaiseEvent Click(vbMiddleButton, _
                             Shift, _
                             CInt(udtPos.X), _
                             CInt(udtPos.Y))
        
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove(Button, _
                                 Shift, _
                                 CInt(udtPos.X), _
                                 CInt(udtPos.Y))

        Case WM_LBUTTONDBLCLK
            'double click
            RaiseEvent DblClick(vbLeftButton, _
                                Shift, _
                                CInt(udtPos.X), _
                                CInt(udtPos.Y))
        
        Case WM_RBUTTONDBLCLK
            'double click
            RaiseEvent DblClick(vbRightButton, _
                                Shift, _
                                CInt(udtPos.X), _
                                CInt(udtPos.Y))
        
        Case WM_MBUTTONDBLCLK
            'double click
            RaiseEvent DblClick(vbMiddleButton, _
                                Shift, _
                                CInt(udtPos.X), _
                                CInt(udtPos.Y))

        Case Else
            Debug.Print lngMousePos
    End Select
End Sub

Private Sub Class_Initialize()
    'make sure the objects are set to Null initially
    Set micoIcon = Nothing
    Set mpicHook = Nothing
    mstrClass = "VB System Tray Icon Class"
    Call ShowIcon
    'Call InitTaskbarMsg(Me)
End Sub

Private Sub Class_Terminate()
    'remove the icon from the system tray if not already
    'done. This ensures that the program cleans up
    'after itself
    Call UnloadIcon
End Sub
