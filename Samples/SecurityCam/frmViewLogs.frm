VERSION 5.00
Begin VB.Form frmViewLogs 
   Caption         =   "View Security Logs"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewLogs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDate 
      Height          =   360
      Left            =   9360
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.ComboBox cboEvent 
      Height          =   360
      Left            =   6480
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "<"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   255
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">"
      Height          =   735
      Left            =   10200
      TabIndex        =   8
      Top             =   6480
      Width           =   255
   End
   Begin VB.Timer timPlay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   6480
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   5160
      Picture         =   "frmViewLogs.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CheckBox chkPlay 
      Caption         =   "&Play"
      Height          =   615
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6540
      Width           =   1335
   End
   Begin VB.PictureBox picTimeLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9615
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   781
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   11775
   End
   Begin VB.ComboBox cboDay 
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox cboMonth 
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox cboYear 
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblEventNum 
      Caption         =   "Of 0"
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   900
      Width           =   855
   End
   Begin VB.Line lnFore 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   11905
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lnBack 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Caption         =   $"frmViewLogs.frx":0984
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   11775
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      Caption         =   "Event "
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblDDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Left            =   8760
      TabIndex        =   17
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "23:59:59"
      Height          =   255
      Left            =   9720
      TabIndex        =   16
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblMiddle 
      Alignment       =   2  'Center
      Caption         =   "12:00:00"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Caption         =   "Day"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lblMonth 
      Alignment       =   1  'Right Justify
      Caption         =   "Month"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      Caption         =   "Year"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   495
   End
End
Attribute VB_Name = "frmViewLogs"
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
'   Description:    This is the screen for viewing the logged data
'  ---------------------------------------------------------------------------
'   Dependancies:   StdOle2.tlb         (standard vb library)
'                   AviCap32.dll        WebCam data capture dll
'                   ScrRun.dll          Microsoft Scripting Runtime
'
'                   frmAboutScreen.frm  modGeneral.bas      modSizeLimit.bas
'                   modSysTrayIcon.bas  modWebCam.bas       clsBitmap.cls
'                   clsSysTrayIcon.cls  ctlSysTray.ctl      cJPEGi.cls
'                   frmSecurity.frm     modRegistry.bas
'  ---------------------------------------------------------------------------
'   References:     http://ej.bantz.com/video/
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50351&lngWId=1
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50065&lngWId=1
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------


'all variables must be declared
Option Explicit


'-------------------------------------------------------------------------------
'                            MODULE LEVEL VARIABLES
'-------------------------------------------------------------------------------
Private mfsFileSys              As FileSystemObject 'used to query the file system
Private mstrDayPicPaths()       As String           'holds the complete file paths to the pictures for the current day
Private mintNumPics             As Integer          'holds the number of pictures within the mstrDayPicPaths() array
Private mintSelPic              As Integer          'holds the array index of the currently selected picture within the array
Private mbmpPreview             As clsBitmap        'holds the picture frame to draw in the Preview picture box
Private mbmpTimeLine            As clsBitmap        'holds the picture graph to display as the timeline
Private mblnNoEvents            As Boolean          'a flag to events not to trigger while we adjust the controls/display the current date otherwise they may trigger multiple times
Private mlngStartSec            As Long             'holds the seconds from midnight that the first event is at
Private mlngEndSec              As Long             'holds the seconds from midnight that the last event is at


'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Private Sub cboDate_Change()
    'update the current day
    
    If (cboDate.ListIndex >= 0) And (cboDate.ListCount > 0) And (Not (timPlay.Enabled Or mblnNoEvents)) Then
        mintSelPic = cboDate.ListIndex
        Call DrawPic
    End If
End Sub

Private Sub cboDate_Click()
    'udpate the current day
    Call cboDate_Change
End Sub

Private Sub cboDate_KeyPress(KeyAscii As Integer)
    'only allow the user to select an item in the list
    Call SelCboItem(KeyAscii, cboDate, True)
End Sub

Private Sub cboDay_Change()
    'update the combox boxes with a valid selection of months/days
    Call GetDayList
    Call DrawPic
End Sub

Private Sub cboDay_Click()
    'update the combo box
    Call cboDay_Change
End Sub

Private Sub cboDay_KeyPress(KeyAscii As Integer)
    'only allow the user to select an item in the list
    Call SelCboItem(KeyAscii, cboDay, True)
End Sub

Private Sub cboEvent_Change()
    'update the current day
    
    If (cboEvent.ListIndex >= 0) And (cboEvent.ListCount > 0) And (Not (timPlay.Enabled Or mblnNoEvents)) Then
        mintSelPic = cboEvent.ListIndex
        Call DrawPic
    End If
End Sub

Private Sub cboEvent_Click()
    'udpate the current day
    Call cboEvent_Change
End Sub

Private Sub cboEvent_KeyPress(KeyAscii As Integer)
    'only allow the user to select an item in the list
    Call SelCboItem(KeyAscii, cboEvent, True)
End Sub

Private Sub cboMonth_Change()
    'update the combo box
    Call GetDays
End Sub

Private Sub cboMonth_Click()
    'update the combo box
    Call cboMonth_Change
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
    'only allow the user to select an item in the list
    Call SelCboItem(KeyAscii, cboMonth, True)
End Sub

Private Sub cboYear_Change()
    'update the combo boxes
    Call GetMonths
End Sub

Private Sub cboYear_Click()
    'update the combo boxes
    Call cboYear_Change
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
    'only allow the user to select an item in the list
    Call SelCboItem(KeyAscii, cboYear, True)
End Sub

Private Sub chkPlay_Click()
    'play/pause the movie
    
    If (chkPlay.Value = vbChecked) Then
        'start playing
        timPlay.Enabled = True
        chkPlay.Caption = "&Pause"
        
    Else
        'pause playing
        timPlay.Enabled = False
        chkPlay.Caption = "&Play"
    End If
End Sub

Private Sub cmdEnd_Click()
    'set the current position to the end of the timeline
    mintSelPic = UBound(mstrDayPicPaths)
    Call DrawPic
End Sub

Private Sub cmdRefresh_Click()
    'update the currently displayed day
    Call cboDay_Change
End Sub

Private Sub cmdStart_Click()
    'set the current position to the start of the events
    mintSelPic = 0
    Call DrawPic
End Sub

Private Sub Form_Load()
    'load up the initial settings
    
    Load frmLogsSplash
    Call frmLogsSplash.Show
    Call frmLogsSplash.DispMsg("Loading log files. This may take a few minutes, please wait...")
    
    Set mfsFileSys = New FileSystemObject
    timPlay.Interval = frmSecurity.timCapture.Interval
    timPlay.Enabled = False
    Call GetYears
    
    'hide
    Unload frmLogsSplash
    Set frmLogsSplash = Nothing     'free up memory
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'perform any necessary cleanup
    
    Set mfsFileSys = Nothing        'free up memory
    Set mbmpPreview = Nothing
    Set mbmpTimeLine = Nothing
End Sub

Private Sub Form_Resize()
    'resize the controls based on what the user wants
    
    
    Dim lngWidth                As Long             'holds the width of the form
    Dim lngHeight               As Long             'holds the height of the form
    
    
    'don't do anything if the form was just minimized
    If (Me.WindowState = vbMinimized) Then
        Exit Sub
    End If
    
    'make sure that we don't hit a minimum size
    If (Me.Width < 12000) Then
        lngWidth = 12000
    Else
        lngWidth = Me.Width
    End If
    If (Me.Height < 8000) Then
        lngHeight = 8000
    Else
        lngHeight = Me.Height
    End If
    
    'if we had to adjust the size of the form then apply the new size
    If (lngWidth <> Me.Width) Or (lngHeight <> Me.Height) Then
        Call Me.Move(Me.Left, Me.Top, lngWidth, lngHeight)
    End If
    
    
    'resize the controls to match
    picPreview.Width = Me.ScaleWidth - 240
    picPreview.Height = Me.ScaleHeight - (picPreview.Top + picTimeLine.Height + lblStart.Height + 240)
    
    lblMsg.Width = picPreview.Width
    
    lnBack.X2 = lblMsg.Left + lblMsg.Width
    lnFore.X2 = lnBack.X2 + Screen.TwipsPerPixelX
    
    picTimeLine.Top = picPreview.Top + picPreview.Height + 120
    picTimeLine.Width = picPreview.Width - (chkPlay.Width + lblEnd.Width + 120)
    
    cmdStart.Top = picTimeLine.Top
    cmdEnd.Top = cmdStart.Top
    cmdEnd.Left = picTimeLine.Left + picTimeLine.Width + 120
    
    chkPlay.Left = Me.ScaleWidth - (chkPlay.Width + 120)
    chkPlay.Top = picTimeLine.Top + ((picTimeLine.Height - chkPlay.Height) \ 2)
    
    lblStart.Top = picTimeLine.Top + picTimeLine.Height + 60
    
    lblMiddle.Top = lblStart.Top
    lblMiddle.Left = picTimeLine.Left + ((picTimeLine.Width - lblMiddle.Width) \ 2)
    
    lblEnd.Top = lblStart.Top
    lblEnd.Left = picTimeLine.Left + (picTimeLine.Width - (lblEnd.Width \ 2))
    
    Call DrawPic
End Sub

Private Sub picPreview_Paint()
    'redraw the preview snapshot
    If Not mbmpPreview Is Nothing Then
        Call mbmpPreview.Paint(picPreview.hDC)
    End If
End Sub

Private Sub picTimeLine_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'calculate where the mouse is in the time line and move to the nearist picture and redraw
    
    
    Select Case Button
    Case vbLeftButton
        'if the user clicked left of the half way mark then move down an image, else move up
        If (x < (picTimeLine.ScaleWidth \ 2)) Then
            
            'can we move down a pic
            If (mintSelPic > 0) Then
                mintSelPic = mintSelPic - 1
                Call DrawPic
            End If
            
        Else
            'move up if possible
            If (mintSelPic < (mintNumPics - 1)) Then
                mintSelPic = mintSelPic + 1
                Call DrawPic
            End If
        End If  'did the user click left of the halfway mark
    End Select
End Sub

Private Sub picTimeLine_Paint()
    'redraw the time line image
    If Not mbmpTimeLine Is Nothing Then
        Call mbmpTimeLine.Paint(picTimeLine.hDC)
    End If
End Sub

Private Sub timPlay_Timer()
    'play from the current position within the event log
    
    Select Case mintSelPic
    Case Is < (mintNumPics - 1)
        mintSelPic = mintSelPic + 1
        Call DrawPic
        
    Case (mintNumPics - 1)
        'stop playing
        chkPlay.Value = vbUnchecked
        timPlay.Enabled = False
        Call DrawPic
    End Select
End Sub

Private Sub DrawPic()
    'This will draw the current picture in mintSelPic
    
    mblnNoEvents = True
    
    Call DrawPreview
    Call DrawTimeLine
    
    mblnNoEvents = False
End Sub

Private Sub DrawPreview()
    'This will draw the preview screen shot for the currently selected image
    
    
    Static bmpBackBuffer        As clsBitmap        'holds the bitmap to draw on
    Static bmpSwap              As clsBitmap        'holds a duplicate reference to the object mbmpPreview but is only used for swapping the pointers between bmpBackBuffer
    Static bmpScale             As clsBitmap        'used to copy the loaded picture to a scale that can be viewed in the picture box
    
    Dim lngWidth                As Long             'holds the width to scale the picture at
    Dim lngHeight               As Long             'holds the height to scale the picture at
    Dim lngTop                  As Long             'holds the Top position to paint the stretched picture at
    Dim lngLeft                 As Long             'holds the Left position to paint the stretched picture at
    Dim sngDestRatio            As Single           'holds the destation ratio between it's width/height
    Dim sngSourceRatio          As Single           'holds the source ratio between it's width/height
    
    
    'create the preview object if we don't have one already
    If mbmpPreview Is Nothing Then
        Set mbmpPreview = New clsBitmap
        Call mbmpPreview.SetBitmap(picPreview.ScaleWidth, _
                                   picPreview.ScaleHeight, _
                                   vbBlack)
    End If
    
    'create the back buffer if necessary
    If bmpBackBuffer Is Nothing Then
        Set bmpBackBuffer = New clsBitmap
    End If
    If (bmpBackBuffer.Width <> picPreview.ScaleWidth) Or (bmpBackBuffer.Height <> picPreview.ScaleHeight) Then
        Call bmpBackBuffer.SetBitmap(picPreview.ScaleWidth, _
                                     picPreview.ScaleHeight, _
                                     vbBlack)
    Else
        Call bmpBackBuffer.Cls
    End If
    
    'is there an image to display
    If (mintSelPic >= mintNumPics) Then
        'no preview available
        Call mbmpPreview.Cls
        Call picPreview_Paint
        Exit Sub
    End If
    
    'do
    If bmpScale Is Nothing Then
        Set bmpScale = New clsBitmap
    End If
    
    
    'load the picture from the file
    Set bmpScale.Picture = LoadPicture(mstrDayPicPaths(mintSelPic))
    
    'we only want the bottom left section of the picture so copy it off
    lngWidth = (bmpScale.Width \ 2)
    lngHeight = (bmpScale.Height \ 2)
    Call bmpScale.Paint(bmpBackBuffer.hDC, , , lngHeight, lngWidth, 0, lngHeight)
    
    'resize bmpScale to the exact size of the sub-image
    Call bmpScale.ReSize(lngWidth, lngHeight)
    Call bmpScale.PaintFrom(bmpBackBuffer.hDC)
    Call bmpBackBuffer.Cls
    
    
    'make sure that the largest side of the source image will fit into the smallest side of
    'the preview screen
    sngSourceRatio = (bmpScale.Width / bmpScale.Height)
    sngDestRatio = (bmpBackBuffer.Width / bmpBackBuffer.Height)
    If (sngDestRatio < sngSourceRatio) Then
        'there'll be a vertical gap between the top/bottom
        lngWidth = bmpBackBuffer.Width
        lngHeight = (bmpBackBuffer.Width / sngSourceRatio)
        lngLeft = 0
        lngTop = (bmpBackBuffer.Height - lngHeight) \ 2
        
    Else
        'there'll be a horizontal gap between the left/right
        lngWidth = (bmpBackBuffer.Height * sngSourceRatio)
        lngHeight = bmpBackBuffer.Height
        lngLeft = (bmpBackBuffer.Width - lngWidth) \ 2
        lngTop = 0
    End If
    
    'scale the image
    'Call bmpScale.Paint(bmpBackBuffer.hDC, lngTop, lngLeft, lngHeight, lngWidth)
    Call bmpBackBuffer.PaintFrom(bmpScale.hDC, bmpScale.Width, bmpScale.Height, , , lngLeft, lngTop, lngWidth, lngHeight)
    
    'display the image on the screen
    Set bmpSwap = mbmpPreview
    Set mbmpPreview = bmpBackBuffer
    Set bmpBackBuffer = bmpSwap
    Call picPreview_Paint
End Sub

Private Sub DrawTimeLine()
    'This will draw out the time line based on the selected log day
    
    
    Static bmpBackBuffer        As clsBitmap        'holds the bitmap to draw on
    Static bmpSwap              As clsBitmap        'holds a duplicate reference to the object mbmpTimeLine but is only used for swapping the pointers between bmpBackBuffer
    Static intIndicatorY(2)     As Integer          'holds the Y co-ordinates for the indicator
    
    Dim intIndicatorX(2)        As Integer          'holds the X co-ordinates for the indicator
    Dim intCurDayX              As Integer          'holds the calculated X co-corinate for the day
    Dim intDayWidth             As Integer          'holds the calculated day width
    Dim strFileTime             As String           'holds the time value extracted from the day path
    Dim lngSeconds              As Long             'holds the number of seconds from midnight the pic was taken at
    Dim intDayCounter           As Integer          'used to cycle through the days
    Dim intPrevX                As Integer          'holds the previous line written out to  - this stops a line being over written
    
    
    'create the time line object if we don't have one already
    If mbmpTimeLine Is Nothing Then
        Set mbmpTimeLine = New clsBitmap
        Call mbmpTimeLine.SetBitmap(picTimeLine.ScaleWidth, _
                                    picTimeLine.ScaleHeight)
        
        'the vertical co-ordinates of the indicator never change so fill them in
        intIndicatorY(0) = 0
        intIndicatorY(1) = 0
        intIndicatorY(2) = 10
    End If
    
    
    'create the back buffer that we're going to draw on
    If bmpBackBuffer Is Nothing Then
        Set bmpBackBuffer = New clsBitmap
        Call bmpBackBuffer.SetBitmap(picTimeLine.ScaleWidth, _
                                     picTimeLine.ScaleHeight)
    End If
    
    intPrevX = -1
    With bmpBackBuffer
        'do we need to resize the back buffer
        If (.Width <> picTimeLine.ScaleWidth) Or (.Height <> picTimeLine.ScaleHeight) Then
            'yes
            Call .SetBitmap(picTimeLine.ScaleWidth, picTimeLine.ScaleHeight)
        Else
            'no but wipe clean the current contents
            Call .Cls
        End If
        
        'draw the box showing the time line itself
        Call .DrawRect(vbBlack, 12, 1, .Height - 15)
        
        'draw three markers showing the start, middle and end
        Call .DrawLine(0, .Height - 3, 0, .Height, vbBlack)
        Call .DrawLine((.Width \ 2), .Height - 3, (.Width \ 2), .Height, vbBlack)
        Call .DrawLine(.Width - 1, .Height - 3, .Width - 1, .Height, vbBlack)
        
        'cycle through the list of files and draw a red line for each file representing
        'where in the day the snapshot was taken (00:00:00 left, 23:59:59 right)
        For intDayCounter = 0 To (mintNumPics - 1)
            
            'extract the file time
            strFileTime = Format(FileDateTime(mstrDayPicPaths(intDayCounter)), "hh:nn:ss")
            lngSeconds = DateDiff("s", "00:00:00", strFileTime)
            intCurDayX = (((lngSeconds - mlngStartSec) / (mlngEndSec - mlngStartSec)) * (.Width - 1))
            
            'is this the currently displayed day
            If (intDayCounter = mintSelPic) Then
                
                'draw the indicator showing where we are in the log for the day
                intIndicatorX(0) = intCurDayX - 4
                intIndicatorX(1) = intCurDayX + 4
                intIndicatorX(2) = intCurDayX
                
                Call .DrawPoly(intIndicatorX(), intIndicatorY(), vbGreen)
                Call .DrawLine(intCurDayX, 12, intCurDayX, .Height - 3, vbYellow)
            
            ElseIf (intCurDayX <> intPrevX) Then
                Call .DrawLine(intCurDayX, 12, intCurDayX, .Height - 3, vbRed)
            End If  'is this the currently displayed day
            intPrevX = intCurDayX
        Next intDayCounter
    End With    'bmpBackBuffer
    
    'swap the buffer for the bitmap actually being displayed. This means that we only ever create
    'bitmaps as needed.
    Set bmpSwap = mbmpTimeLine
    Set mbmpTimeLine = bmpBackBuffer
    Set bmpBackBuffer = bmpSwap
    Call picTimeLine_Paint
    
    'update what we're on
    If (mintNumPics > 0) Then
        cboDate.ListIndex = mintSelPic
        lblEventNum.Caption = "of " + Format(mintNumPics, "#,###,##0")
        cboEvent.ListIndex = mintSelPic
    Else
        lblEvent.Caption = "of 0"
    End If
End Sub

Private Sub GetDays()
    'This will load the days into the Day combo box
    
    
    Dim strFilePath             As String           'holds the file path to the Days folder
    Dim fldDays                 As Folder           'used to hold the folder data
    Dim fldCounter              As Folder           'used to cycle through the sub folders
    Dim strDays()               As String           'holds the list of days loaded
    Dim intNumDays              As Integer          'holds the number of days loaded
    Dim intCounter              As Integer          'used to cycle through the array to add the new data in
    
    
    'initialise the combo box
    Call cboDay.Clear
    
    If (cboMonth.ListIndex < 0) Or (cboYear.ListIndex < 0) Then
        Exit Sub
    End If
    
    'get the path to the directory with the days in it and initialise the array
    strFilePath = mfsFileSys.BuildPath(App.Path, cboYear.List(cboYear.ListIndex))
    strFilePath = mfsFileSys.BuildPath(strFilePath, cboMonth.List(cboMonth.ListIndex))
    intNumDays = 0
    ReDim strDays(intNumDays)
    
    'cycle through the sub folders
    Set fldDays = mfsFileSys.GetFolder(strFilePath)
    For Each fldCounter In fldDays.SubFolders
        'is the folder name in the correct format (a 4 digit numeric)
        If IsNumeric(fldCounter.Name) And (InStr(1, fldCounter.Name, ".") = 0) Then
            'save this folder
            ReDim Preserve strDays(intNumDays)
            strDays(intNumDays) = fldCounter.Name
            intNumDays = intNumDays + 1
        End If
    Next fldCounter
    
    'sort the days and cycle through the list in DECENDING order so that the newist
    'is always at the top of the list
    Call QSortStrings(strDays())
    For intCounter = UBound(strDays) To LBound(strDays) Step -1
        Call cboDay.AddItem(strDays(intCounter))
    Next intCounter
    
    'if there are any days listed then select the first one and display the time line
    If (intNumDays > 0) Then
        cboDay.ListIndex = 0
        Call GetDayList
        Call DrawPic
    End If
End Sub

Private Sub GetDayList()
    'This will get the list of pictures for the currently selected day
    
    
    Dim strFilePath             As String           'holds the file path to the Days folder
    Dim fldDays                 As Folder           'used to hold the folder data
    Dim filCounter              As File             'used to cycle through the files
    Dim strDays()               As String           'holds the list of days loaded
    Dim intNumDays              As Integer          'holds the number of days loaded
    Dim intCounter              As Integer          'used to cycle through the array to add the new data in
    
    
    Call cboDate.Clear
    Call cboEvent.Clear
    If (cboMonth.ListIndex < 0) Or (cboYear.ListIndex < 0) Or (cboDay.ListIndex < 0) Then
        Exit Sub
    End If
    
    'get the path to the directory with the days in it and initialise the array
    strFilePath = mfsFileSys.BuildPath(App.Path, cboYear.List(cboYear.ListIndex))
    strFilePath = mfsFileSys.BuildPath(strFilePath, cboMonth.List(cboMonth.ListIndex))
    strFilePath = mfsFileSys.BuildPath(strFilePath, cboDay.List(cboDay.ListIndex))
    intNumDays = 0
    ReDim strDays(intNumDays)
    
    'cycle through the files
    Set fldDays = mfsFileSys.GetFolder(strFilePath)
    For Each filCounter In fldDays.Files
        'is the filename name in the correct format (a 4 digit numeric)
        If (Mid(filCounter.Name, 1, 8) Like "##-##-##") Then
            'save this file name
            ReDim Preserve strDays(intNumDays)
            strDays(intNumDays) = filCounter.Path
            intNumDays = intNumDays + 1
        End If
    Next filCounter
    
    'sort the days and cycle through the list in DECENDING order so that the newist
    'is always at the top of the list
    Call QSortStrings(strDays())
    
    For intCounter = 0 To (intNumDays - 1)
        Call cboDate.AddItem(Format(FileDateTime(strDays(intCounter)), "General Date"))
        Call cboEvent.AddItem((intCounter + 1))
    Next intCounter
    
    'calculate the starting and ending point of the timeline
    If (intNumDays > 0) Then
        mlngStartSec = DateDiff("s", "00:00:00", Format(FileDateTime(strDays(0)), "Long Time"))
        mlngEndSec = DateDiff("s", "00:00:00", Format(FileDateTime(strDays(intNumDays - 1)), "Long Time"))
    Else
        mlngStartSec = 0        '00:00:00
        mlngEndSec = 86399      '23:59:59
    End If
    lblStart.Caption = Format(DateAdd("s", mlngStartSec, "00:00:00"), "Long Time")
    lblMiddle.Caption = Format(DateAdd("s", ((mlngEndSec - mlngStartSec) / 2) + mlngStartSec, "00:00:00"), "Long Time")
    lblEnd.Caption = Format(DateAdd("s", mlngEndSec, "00:00:00"), "Long Time")
    
    'set the current position within the day
    mintSelPic = 0
    mintNumPics = intNumDays
    mstrDayPicPaths() = strDays()
End Sub

Private Sub GetMonths()
    'This will load the months into the Month combo box
    
    
    Dim strFilePath             As String           'holds the file path to the Months folder
    Dim fldMonths               As Folder           'used to hold the folder data
    Dim fldCounter              As Folder           'used to cycle through the sub folders
    Dim strMonths()             As String           'holds the list of months loaded
    Dim intNumMonths            As Integer          'holds the number of months loaded
    Dim intCounter              As Integer          'used to cycle through the array to add the new data in
    
    
    'initialise the combo boxes
    Call cboMonth.Clear
    Call cboDay.Clear
    
    If (cboYear.ListIndex < 0) Then
        Exit Sub
    End If
    
    'get the path to the directory with the months in it and initialise the array
    strFilePath = mfsFileSys.BuildPath(App.Path, cboYear.List(cboYear.ListIndex))
    intNumMonths = 0
    ReDim strMonths(intNumMonths)
    
    'cycle through the sub folders
    Set fldMonths = mfsFileSys.GetFolder(strFilePath)
    For Each fldCounter In fldMonths.SubFolders
        'is the folder name in the correct format (a 4 digit numeric)
        If IsNumeric(fldCounter.Name) And (InStr(1, fldCounter.Name, ".") = 0) Then
            'save this folder
            ReDim Preserve strMonths(intNumMonths)
            strMonths(intNumMonths) = fldCounter.Name
            intNumMonths = intNumMonths + 1
        End If
    Next fldCounter
    
    'sort the months and cycle through the list in DECENDING order so that the newist
    'is always at the top of the list
    Call QSortStrings(strMonths())
    For intCounter = UBound(strMonths) To LBound(strMonths) Step -1
        Call cboMonth.AddItem(strMonths(intCounter))
    Next intCounter
    
    'if there are any months listed then select the first one and display the days for it
    If (intNumMonths > 0) Then
        cboMonth.ListIndex = 0
        Call GetDays
    End If
End Sub

Private Sub GetYears()
    'This will load the years into the Year combo box
    
    
    Dim strFilePath             As String           'holds the file path to the Years folder
    Dim fldYears                As Folder           'used to hold the folder data
    Dim fldCounter              As Folder           'used to cycle through the sub folders
    Dim strYears()              As String           'holds the list of years loaded
    Dim intNumYears             As Integer          'holds the number of years loaded
    Dim intCounter              As Integer          'used to cycle through the array to add the new data in
    
    
    'initialise the combo boxes
    Call cboYear.Clear
    Call cboMonth.Clear
    Call cboDay.Clear
    
    'get the path to the directory with the years in it and initialise the array
    strFilePath = App.Path
    intNumYears = 0
    ReDim strYears(intNumYears)
    
    'cycle through the sub folders
    Set fldYears = mfsFileSys.GetFolder(strFilePath)
    For Each fldCounter In fldYears.SubFolders
        'is the folder name in the correct format (a 4 digit numeric)
        If (fldCounter.Name Like "####") Then
            'save this folder
            ReDim Preserve strYears(intNumYears)
            strYears(intNumYears) = fldCounter.Name
            intNumYears = intNumYears + 1
        End If
    Next fldCounter
    
    'sort the years and cycle through the list in DECENDING order so that the newist
    'is always at the top of the list
    Call QSortStrings(strYears())
    For intCounter = UBound(strYears) To LBound(strYears) Step -1
        Call cboYear.AddItem(strYears(intCounter))
    Next intCounter
    
    'if there are years then select the first one and display the months
    If (intNumYears > 0) Then
        cboYear.ListIndex = 0
        Call GetMonths
    End If
End Sub
