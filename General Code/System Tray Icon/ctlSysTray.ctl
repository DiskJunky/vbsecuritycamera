VERSION 5.00
Begin VB.UserControl ctlSysTray 
   Appearance      =   0  'Flat
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   300
   ToolboxBitmap   =   "ctlSysTray.ctx":0000
   Begin VB.PictureBox picHook 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   0
      Picture         =   "ctlSysTray.ctx":0312
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "ctlSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     4 June 2003
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    System Tray Icon Control
' -----------------------------------------------
'COMMENTS : This will use both the class and the
'   module to create the system tray icon and
'   process any events/messages.
'=================================================

'require variable declaration
Option Explicit

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
'              MODULE LEVEL VARIABLES
'------------------------------------------------
Private WithEvents msysIcon     As clsSysTrayIcon   'holds a reference to the system tray class
Attribute msysIcon.VB_VarHelpID = -1
Private mblnAutoReload          As Boolean          'flags whether or not the control should automatically reload the icon when exploer.exe restarts

'------------------------------------------------
'                  PROPERTIES
'------------------------------------------------
Public Property Let AutoReload(ByVal blnAuto As Boolean)
Attribute AutoReload.VB_Description = "Sets whether or not the icon should reload itself when explorer restarts itself."
    'This will set whether or not the control should automatically reload the
    'system tray icon if explorer restarts the system tray
    mblnAutoReload = blnAuto
End Property

Public Property Get AutoReload() As Boolean
    'This will return wether or not the control should automatically reload
    'the system trayicon if explorer restarts the system tray
    AutoReload = mblnAutoReload
End Property

Public Property Set Icon(ByVal picNew As StdPicture)
Attribute Icon.VB_Description = "Sets the icon to be displayed in the system tray"
    'This will set the icon that is used for a picture
    Set picHook.Picture = picNew
    Set msysIcon.Icon = picNew
End Property

Public Property Get Icon() As StdPicture
    'This will return a reference to the picture used for the icon
    Set Icon = msysIcon.Icon
End Property

Public Property Let ToolTip(ByVal strNewTip As String)
Attribute ToolTip.VB_Description = "Sets the tooltip for the system tray icon"
    'This will set the tooltip for the system tray icon
    msysIcon.ToolTip = strNewTip
End Property

Public Property Get ToolTip() As String
    'This will return the tooltip for the system tray icon
    ToolTip = msysIcon.ToolTip
End Property

'------------------------------------------------
'                  PROCEDURES
'------------------------------------------------
Private Sub msysIcon_Click(ByVal Button As Integer, _
                           ByVal Shift As Integer, _
                           ByVal X As Integer, _
                           ByVal Y As Integer)
    'activate the click event
    RaiseEvent Click(Button, Shift, X, Y)
End Sub

Private Sub msysIcon_DblClick(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer)
    'activate the double click event
    RaiseEvent DblClick(Button, Shift, X, Y)
End Sub

Private Sub msysIcon_MouseDown(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer)
    'activate the mouse down event
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub msysIcon_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer)
    'activate the mouse move event
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub msysIcon_MouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Integer, _
                             ByVal Y As Integer)
    'activate the mouse up event
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub msysIcon_TaskbarCreated()
    'activate the taskbar created event
    RaiseEvent TaskbarCreated
    
    If mblnAutoReload Then
        'reload the icon
        Call Reload
    End If
End Sub

Private Sub picHook_Resize()
    'set the size of the control to be the same as the picture box
    With UserControl
        .Width = picHook.Width
        .Height = picHook.Height
    End With    'UserControl
End Sub

Private Sub UserControl_Initialize()
    'setup any details before starting up the control
    
    'create the system tray icon class
    Set msysIcon = New clsSysTrayIcon
    Set msysIcon.PictureBox = picHook
    mblnAutoReload = True
End Sub

Private Sub UserControl_InitProperties()
    'load up the default properties
    
    'set the size of the control to be the same as the picture box
    With UserControl
        .Width = picHook.Width
        .Height = picHook.Height
    End With    'UserControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'get the properties from the last save
    
    With PropBag
        Set msysIcon.Icon = .ReadProperty("Icon", Nothing)
        Set picHook.Picture = msysIcon.Icon
        msysIcon.ToolTip = .ReadProperty("ToolTip", Ambient.DisplayName)
        mblnAutoReload = .ReadProperty("AutoReload", True)
    End With    'PropBag
End Sub

Private Sub UserControl_Resize()
    'make sure that the control stays a fixed size
    
    With UserControl
        .Width = picHook.Width
        .Height = picHook.Height
    End With    'Me
End Sub

Private Sub UserControl_Terminate()
    'remove the icon from the system tray
    Set msysIcon = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save any property values
    
    With PropBag
        Call .WriteProperty("Icon", msysIcon.Icon, Nothing)
        Call .WriteProperty("ToolTip", msysIcon.ToolTip, Ambient.DisplayName)
        Call .WriteProperty("AutoReload", mblnAutoReload, True)
    End With    'PropBag
End Sub

Public Sub Show()
Attribute Show.VB_Description = "Displays the icon in the system tray"
    'This will display the icon in the system tray
    Call msysIcon.ShowIcon
End Sub

Public Sub Hide()
Attribute Hide.VB_Description = "Removes the icon from the system tray"
    'This will unload the icon from the system tray
    Call msysIcon.UnloadIcon
End Sub

Public Sub Reload()
Attribute Reload.VB_Description = "Reload the icon in the system tray. This will move its' position in the system tray to the most receint"
    'This will redisplay the system tray icon
    Call msysIcon.ShowIcon
End Sub
