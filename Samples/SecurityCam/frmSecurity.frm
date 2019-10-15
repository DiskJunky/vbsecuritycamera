VERSION 5.00
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Camera"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCamera 
      Caption         =   "Camera"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CheckBox chkCamera 
         Height          =   255
         Left            =   885
         TabIndex        =   5
         Top             =   0
         Value           =   1  'Checked
         Width           =   240
      End
      Begin VB.PictureBox picMotion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   5040
         ScaleHeight     =   3225
         ScaleWidth      =   3945
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.PictureBox picSnap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   240
         ScaleHeight     =   3225
         ScaleWidth      =   3945
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lblMotionMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Motion Detected"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Label lblSnapMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Snapshot"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   4080
         Width           =   3975
      End
   End
   Begin VB.Timer timCapture 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   8160
      Top             =   4200
   End
   Begin SecurityCamera.ctlSysTray tryMenu 
      Left            =   8760
      Top             =   4200
      _ExtentX        =   953
      _ExtentY        =   953
      Icon            =   "frmSecurity.frx":0442
      ToolTip         =   ""
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuTrayShowBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayTurn 
         Caption         =   "Turn &Camera On"
      End
      Begin VB.Menu mnuTrayTurnBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuTrayLogs 
         Caption         =   "View &Logs"
      End
      Begin VB.Menu mnuTrayBreakExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileHide 
         Caption         =   "H&ide"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileBreakExit 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewLogs 
         Caption         =   "&Logs"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   SecurityCamera
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           23 May 2006
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Email:          DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This is a program designed to detect motion and record
'                   snapshots when detected.
'  ---------------------------------------------------------------------------
'   Dependancies:   StdOle2.tlb         (standard vb library)
'                   AviCap32.dll        WebCam data capture dll
'                   ScrRun.dll          Microsoft Scripting Runtime
'
'                   frmAboutScreen.frm  modGeneral.bas      modSizeLimit.bas
'                   modSysTrayIcon.bas  modWebCam.bas       clsBitmap.cls
'                   clsSysTrayIcon.cls  ctlSysTray.ctl      cJPEGi.cls
'                   frmCamOptions.frm   modRegistry.bas
'  ---------------------------------------------------------------------------
'   References:     http://ej.bantz.com/video/
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50351&lngWId=1
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50065&lngWId=1
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------


'all variables must be declared
Option Explicit


'-------------------------------------------------------------------------------
'                               API DECLARATIONS
'-------------------------------------------------------------------------------
'returns the number of milliseconds (second/1000) that windows has been active
'for
Private Declare Function GetTickCount Lib "kernel32" () As Long


'-------------------------------------------------------------------------------
'                              USER DEFINED TYPES
'-------------------------------------------------------------------------------
Private Type ColourType
    intRed          As Integer
    intGreen        As Integer
    intBlue         As Integer
End Type


'-------------------------------------------------------------------------------
'                            MODULE LEVEL CONSTANTS
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
'                            MODULE LEVEL VARIABLES
'-------------------------------------------------------------------------------
Private mintMotionTolerance     As Integer          'holds the pixel gap size when checking for motion within the picture
Private mcolLast()              As ColourType       'holds the pixel colour data from the last screenshot
Private mintCamHeight           As Integer          'holds the camera height in pixels
Private mintCamWidth            As Integer          'holds the camera width in pixels
Private mfsFileSys              As FileSystemObject 'used to query the file system
Private mbmpLast                As clsBitmap        'holds the last screenshot taken
Private mbmpMotion              As clsBitmap        'holds the last motion screenshot
Private mlngFocusDelay          As Long             'holds the milliseconds to wait while the camera auto-focuses when it's first turned on
Private mlngCamStarted          As Long             'holds the tick in milliseconds that the camera was started on
Private mintColourTolerance     As Integer          'holds the percentage by which the colour can change before we flag is as 'Motion'
Private msngCurMovement         As Single           'holds the current movement detected
Private mlngFrameRate           As Long             'holds how many ticks to wait before checking the camera
Private mblnStartCameraOn       As Boolean          'holds if the camera is on as soon as the program starts
Private mblnLogMovements        As Boolean          'holds whether or not to log the movements
Private mblnStartMin            As Boolean          'holds whether or not the program starts minized


'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Private Sub chkCamera_Click()
    'enable/disable polling from the camera
    
    
    Dim strErrMsg               As String           'holds the error message returned back from the WebCam functions
    Dim msgResult               As VbMsgBoxResult   'holds the user's response to the message box
    
    
    'should the camera be enabled
    If (chkCamera.Value = vbChecked) Then
        'enable
        Call StartCam(Me.hwnd, strErrMsg)
        If (strErrMsg <> "") Then
            msgResult = MsgBox(strErrMsg, vbInformation + vbInformation, Me.Caption)
        Else
            timCapture.Enabled = True
            mnuTrayTurn.Caption = "Turn Camera &Off"
        End If
        
    Else
        'disable
        timCapture.Enabled = False
        Call StopCam
        mnuTrayTurn.Caption = "Turn Camera &On"
    End If  'should the camera be enabled
End Sub

Private Sub Form_Load()
    'get all the necessary settings and connect to the webcam
    
    
    Dim strErrMsg               As String           'holds the error message returned back from the WebCam functions
    Dim msgResult               As VbMsgBoxResult   'holds the user's response to the message box
    Dim picTemp                 As StdPicture       'used to temporarily hold a screen capture to determine the size of the screenshots
    Dim bmpTemp                 As clsBitmap        'used to accuratly get the size of the bitmap
    
    
    Set mfsFileSys = New FileSystemObject
    
    'display the tray icon
    Call tryMenu.Show
    
    'get the settings
    Call LoadSettings
    mlngCamStarted = GetTickCount
    
    'initialise the array to 2 dimensions
    ReDim mcolLast(0, 0)
    
    'start the camera and get a screenshot to determine the size of the screenshots as
    'we'll need to resize the window to display them
    Call StartCam(Me.hwnd, strErrMsg)
    If (strErrMsg <> "") Then
        'we cannot connect
        msgResult = MsgBox(strErrMsg, vbInformation + vbOKOnly, Me.Caption)
        
    Else
        'get the screenshot size
        Set picTemp = New StdPicture
        Set bmpTemp = New clsBitmap
        Call GetFromCam(picTemp)
        Set bmpTemp.Picture = picTemp
        mintCamWidth = bmpTemp.Width
        mintCamHeight = bmpTemp.Height
        
        'array is zero-based so it's always one less than the actual pixel size
        ReDim mcolLast((mintCamWidth \ mintMotionTolerance) - 1, _
                       (mintCamHeight \ mintMotionTolerance) - 1)
        
        'resize the controls on the form so that the screenshots can be displayed
        picSnap.Width = bmpTemp.Width * Screen.TwipsPerPixelX
        picSnap.Height = bmpTemp.Height * Screen.TwipsPerPixelY
        lblSnapMsg.Top = picSnap.Top + picSnap.Height + 120
        lblSnapMsg.Width = picSnap.Width
        picMotion.Width = picSnap.Width
        picMotion.Height = picSnap.Height
        picMotion.Left = picSnap.Left + picSnap.Width + 600
        lblMotionMsg.Left = picMotion.Left
        lblMotionMsg.Top = lblSnapMsg.Top
        lblMotionMsg.Width = picMotion.Width
        fraCamera.Width = picMotion.Left + picMotion.Width + 480
        fraCamera.Height = lblSnapMsg.Top + lblSnapMsg.Height + 480
        
        'resize the form to match
        '   width = form_border_width + frame_width + area_around_frame
        Me.Width = (Me.Width - Me.ScaleWidth) + fraCamera.Left + fraCamera.Width + 120
        Me.Height = (Me.Height - Me.ScaleHeight) + fraCamera.Top + fraCamera.Height + 120
        
        timCapture.Enabled = True
        timCapture.Interval = mlngFrameRate
    End If  'is the camera started
    
    chkCamera.Value = Abs(mblnStartCameraOn)
    Me.Visible = Not mblnStartMin
    
    If mblnStartCameraOn Then
        mnuTrayTurn.Caption = "Turn Camera &Off"
    Else
        mnuTrayTurn.Caption = "Turn Camera &On"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'perform any necessary clean up
    
    Call StopCam
    Call tryMenu.Hide
    Set mfsFileSys = Nothing       'free up memory
    Set mbmpLast = Nothing
    
    Call SaveSettings
End Sub

Private Sub Form_Resize()
    'hide the form when minimized
    
    If (Me.WindowState = vbMinimized) Then
        Call Me.Hide
    End If
End Sub

Private Sub mniFileExit_Click()
    'exit the program
    Call UnloadAll(Me)
End Sub

Private Sub mnuHelpAbout_Click()
    'display the about screen
    Load frmAboutScreen
    Call frmAboutScreen.Show(vbModal, Me)
    Set frmAboutScreen = Nothing        'free up memory
End Sub

Private Sub mnuTrayExit_Click()
    'exit the form
    Call UnloadAll(Me)
End Sub

Private Sub mnuTrayLogs_Click()
    'display the logs screen
    Call mnuViewLogs_Click
End Sub

Private Sub mnuTrayOptions_Click()
    'display the options screen
    Call mnuViewOptions_Click
End Sub

Private Sub mnuTrayShow_Click()
    'display the form
    
    If (Me.WindowState = vbMinimized) Then
        Me.WindowState = vbNormal
    End If
    Call Me.Show
End Sub

Private Sub mnuTrayTurn_Click()
    'turn the camera on/off
    
    If (chkCamera.Value = vbChecked) Then
        'turn off
        chkCamera.Value = vbUnchecked
        
    Else
        'turn on
        chkCamera.Value = vbChecked
    End If
End Sub

Private Sub mnuViewLogs_Click()
    'display the logs screen
    Load frmViewLogs
    Call frmViewLogs.Show
End Sub

Private Sub mnuViewOptions_Click()
    'display the options screen
    
    Load frmCamOptions
    Call frmCamOptions.GetOptions(mintMotionTolerance, _
                                  mintColourTolerance, _
                                  mlngFocusDelay, _
                                  mlngFrameRate, _
                                  mblnStartCameraOn, _
                                  mblnLogMovements, _
                                  mblnStartMin)
    Set frmCamOptions = Nothing     'free up memory
    
    'apply any visible settings to the relevant control(s)
    timCapture.Interval = mlngFrameRate
End Sub

Private Sub picMotion_Paint()
    'display the last Motion snapshot (if there is any)
    
    If Not mbmpMotion Is Nothing Then
        Call mbmpMotion.Paint(picMotion.hDC)
    End If
End Sub

Private Sub picSnap_Paint()
    'display the last snapshot (if there is any)
    
    If Not mbmpLast Is Nothing Then
        Call mbmpLast.Paint(picSnap.hDC)
    End If
End Sub

Private Sub timCapture_Timer()
    'get another screenshot
    
    
    'have we passed the initial wait while the camera adjusts it's focus
    If ((mlngFocusDelay + mlngCamStarted) < GetTickCount) Then
        Call DetectMotion
        
    Else
        'grab a junk picture or the camera won't auto-focus
        Call DetectMotion(True)
    End If  'have we passed the initial wait while the camera adjusts it's focus
End Sub

Private Sub tryMenu_DblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    'display the form
    Call Me.Show
End Sub

Private Sub tryMenu_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    'display the menu
    
    If (Button = vbRightButton) Then
        Call Me.PopupMenu(mnuTray, , , , mnuTrayShow)
    End If
End Sub

Private Sub DetectMotion(Optional ByVal blnCaptureNoSave As Boolean = False)
    'This procedure will take a screenshot from the camera and check it for motion and if
    'it finds any, it will save the screenshot and display it on the screen
    
    
    Static picTemp          As StdPicture           'holds a screenshot from the WebCam
    Static bmpTemp          As clsBitmap            'used to get data from the screenshot
    Static bmpTempM         As clsBitmap            'holds the screenshot with the motion points displayed (if any)
    Static bmpPrev          As clsBitmap            'holds the last bitmap scanned
    
    Dim intRow              As Integer              'used to cycle through the rows
    Dim intCol              As Integer              'used to cycle through the columns
    Dim intX                As Integer              'holds the pixel X position
    Dim intY                As Integer              'holds the pixel Y position
    Dim blnMotion           As Boolean              'flags if we detected any motion
    Dim colPoint            As ColourType           'holds the colour values from a single point in the picture
    Dim lngMotionCount      As Long                 'holds how many pixels were detected with motion
    Dim sngMotionLevel      As Single               'holds the amount of motion was detected in the picture
    
    
    'initialise the bitmaps
    If mbmpLast Is Nothing Then
        Set mbmpLast = New clsBitmap
        Call mbmpLast.SetBitmap(mintCamWidth, mintCamHeight, vbBlack)
    End If
    If mbmpMotion Is Nothing Then
        Set mbmpMotion = New clsBitmap
        Call mbmpMotion.SetBitmap(mintCamWidth, mintCamHeight, vbBlack)
    End If
    
    'get a screenshot
    If bmpTempM Is Nothing Then
        Set bmpTempM = New clsBitmap
        Call bmpTempM.SetBitmap(mintCamWidth, mintCamHeight, vbBlack)
    End If
    If bmpTemp Is Nothing Then
        Set bmpTemp = New clsBitmap
        Call bmpTemp.SetBitmap(mintCamWidth, mintCamHeight, vbBlack)
    End If
    If picTemp Is Nothing Then
        Set picTemp = New StdPicture
    End If
    Call GetFromCam(picTemp)
    Set bmpTemp.Picture = picTemp
    Call bmpTemp.Paint(bmpTempM.hDC)
    
    blnMotion = False
    lngMotionCount = 0
    
    'check for motion
    For intRow = LBound(mcolLast, 1) To UBound(mcolLast, 1)
        For intCol = LBound(mcolLast, 2) To UBound(mcolLast, 2)
            
            'convert the array indices into pixel co-ordinates
            intX = intRow * mintMotionTolerance
            intY = intCol * mintMotionTolerance
            
            'get the colour of that pixel
            Call bmpTemp.GetRGB(bmpTemp.Pixel(intX, intY), _
                                colPoint.intRed, _
                                colPoint.intGreen, _
                                colPoint.intBlue)
            
            'do we have a variance
            With mcolLast(intRow, intCol)
                If (Abs(.intRed - colPoint.intRed) > mintColourTolerance) Or _
                   (Abs(.intGreen - colPoint.intGreen) > mintColourTolerance) Or _
                   (Abs(.intBlue - colPoint.intBlue) > mintColourTolerance) Then
                    
                    'draw a red dot where the motion was detected
                    Call bmpTempM.DrawEllipse(intX, intY, 3, 3, , vbRed)
                    blnMotion = True
                    lngMotionCount = lngMotionCount + 1
                End If  'do we have a variance
            End With    'mcolLast(intRow, intCol)
            mcolLast(intRow, intCol) = colPoint
        Next intCol
    Next intRow
    
    
    'display the pictures
    If bmpPrev Is Nothing Then
        Set bmpPrev = New clsBitmap
        Call bmpPrev.SetBitmap(bmpTemp.Width, bmpTemp.Height, vbWhite)
    End If
    If blnMotion And (Not blnCaptureNoSave) Then
        Call SaveMotion(bmpPrev, bmpTemp, bmpTempM)
        Set mbmpLast = bmpTemp
        Set mbmpMotion = bmpTempM
        
        Call picSnap_Paint
        Call picMotion_Paint
        
        lblSnapMsg.Caption = "Last Snapshot Taken At: " + Format(Now, "dd/mm/yyyy  hh:nn:ss")
        
        sngMotionLevel = ((lngMotionCount / (intRow * intCol)) * 100)
        lblMotionMsg.Caption = "Motion Level " + Format(sngMotionLevel, "0") + "%"
        msngCurMovement = sngMotionLevel / 100
    Else
        msngCurMovement = 0
    End If
    Call bmpTemp.Paint(bmpPrev.hDC)
End Sub

Private Sub LoadSettings()
    'This will load the settings from the registry
    
    
    Dim strTemp             As String               'holds the value returned from the registry functions
    Dim lngTemp             As String               'holds the numeric value returned from the registry functions
    Dim msgResult           As VbMsgBoxResult       'holds the users response to the message box
    
    
    'set the default values
    mlngFocusDelay = 3000           'three seconds
    mintMotionTolerance = 15        'distance between pixels
    mintColourTolerance = 25        'ignore a (25/255)% colour change
    mlngFrameRate = 500             'have a second
    mblnStartCameraOn = True
    mblnLogMovements = True
    mblnStartMin = False
    
    
    'can we access the registry to get our settings
    strTemp = ReadRegString(HKEY_CURRENT_USER, App.Title, "Movement Detail")
    If (InStr(1, strTemp, "Error", vbTextCompare) > 0) Then
        'just use the defaults but confirm that the user wants to save the screenshots
        'as they take up space on the hard-drive
        msgResult = MsgBox("Do you wish to log detected motion while the camera is on?" + vbCrLf + _
                           "This will take up some hard drive space but how much will" + vbCrLf + _
                           "depend on how long the camera is on and how much motion it" + vbCrLf + _
                           "detects.", _
                           vbQuestion + vbYesNo, _
                           Me.Caption)
        If (msgResult = vbNo) Then
            mblnLogMovements = False
        End If
        Exit Sub
    End If
    mintMotionTolerance = Val(strTemp)
    mintColourTolerance = Val(ReadRegString(HKEY_CURRENT_USER, App.Title, "Movement Tolerance"))
    mlngFocusDelay = Val(ReadRegString(HKEY_CURRENT_USER, App.Title, "AutoFocus Delay"))
    mlngFrameRate = Val(ReadRegString(HKEY_CURRENT_USER, App.Title, "Frame Rate"))
    mblnStartCameraOn = ReadRegLong(HKEY_CURRENT_USER, App.Title, "Start Camera On")
    mblnLogMovements = ReadRegLong(HKEY_CURRENT_USER, App.Title, "Log Movements")
    mblnStartMin = ReadRegLong(HKEY_CURRENT_USER, App.Title, "Start Minimized")
End Sub

Private Sub SaveMotion(ByRef bmpBefore As clsBitmap, _
                       ByRef bmpAfter As clsBitmap, _
                       ByRef bmpMotion As clsBitmap)
    'This procedure will build a composite image showing the before and after images and
    'then save them to a time/stamped file
    
    
    Static bmpSave              As clsBitmap            'holds the composite image
    Static jpgSave              As cJPEGi               'used to save the image as a jpg file
    
    Dim strFullPath             As String               'holds the complete file path to use
    
    
    'don't bother doing anything if we're not set to log the movements
    If Not mblnLogMovements Then
        Exit Sub
    End If
    
    
    'build up the full file path. The path to the pictures is;
    '       <program path>\<year>\<month>\<day>\<time>.bmp
    With mfsFileSys
        strFullPath = .BuildPath(App.Path, Year(Date))
        If Not .FolderExists(strFullPath) Then
            Call .CreateFolder(strFullPath)
        End If
        strFullPath = .BuildPath(strFullPath, Month(Date))
        If Not .FolderExists(strFullPath) Then
            Call .CreateFolder(strFullPath)
        End If
        strFullPath = .BuildPath(strFullPath, Day(Date))
            If Not .FolderExists(strFullPath) Then
            Call .CreateFolder(strFullPath)
        End If
        strFullPath = .BuildPath(strFullPath, Format(Time, "hh-nn-ss-") & GetTickCount & ".jpg")
    End With    'mfsfilesys
    
    'build the composite image
    If bmpSave Is Nothing Then
        Set bmpSave = New clsBitmap
        Call bmpSave.SetBitmap(bmpBefore.Width + bmpMotion.Width, bmpBefore.Height + bmpAfter.Height, vbWhite)
    Else
        Call bmpSave.Cls
    End If
    If jpgSave Is Nothing Then
        Set jpgSave = New cJPEGi
        jpgSave.Quality = 60
    End If
    
    
    Call bmpBefore.Paint(bmpSave.hDC)
    Call bmpAfter.Paint(bmpSave.hDC, , bmpBefore.Height)
    Call bmpMotion.Paint(bmpSave.hDC, bmpBefore.Width, bmpBefore.Height)
    Call bmpSave.DrawString("Before and after images with detected motion", (bmpBefore.Height \ 2) - 15, bmpBefore.Width + 1, 18, bmpMotion.Width, Me.Font)
    Call bmpSave.DrawString(Format(Now, "Long Date") + "  " + Format(Time, "hh:nn:ss"), (bmpBefore.Height \ 2) + 15, bmpBefore.Width + 1, 18, bmpMotion.Width, Me.Font)
    Call jpgSave.SampleHDC(bmpSave.hDC, bmpSave.Width, bmpSave.Height)
    Call jpgSave.SaveFile(strFullPath)
End Sub

Private Sub SaveSettings()
    'This will save all the settings for the program
    
    
    Call CreateRegString(HKEY_CURRENT_USER, App.Title, "Movement Detail", mintMotionTolerance)
    Call CreateRegString(HKEY_CURRENT_USER, App.Title, "Movement Tolerance", mintColourTolerance)
    Call CreateRegString(HKEY_CURRENT_USER, App.Title, "AutoFocus Delay", mlngFocusDelay)
    Call CreateRegString(HKEY_CURRENT_USER, App.Title, "Frame Rate", mlngFrameRate)
    Call CreateRegLong(HKEY_CURRENT_USER, App.Title, "Start Camera On", CLng(mblnStartCameraOn))
    Call CreateRegLong(HKEY_CURRENT_USER, App.Title, "Log Movements", CLng(mblnLogMovements))
    Call CreateRegLong(HKEY_CURRENT_USER, App.Title, "Start Minimized", CLng(mblnStartMin))
End Sub
