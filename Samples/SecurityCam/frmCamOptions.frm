VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCamOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Camera Options"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCamOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox chkStartMin 
      Alignment       =   1  'Right Justify
      Caption         =   "Start &Mimized To System Tray"
      Height          =   255
      Left            =   2940
      TabIndex        =   6
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CheckBox chkLogMovements 
      Alignment       =   1  'Right Justify
      Caption         =   "&Log Movements"
      Height          =   255
      Left            =   2940
      TabIndex        =   5
      Top             =   4920
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkStartCameraOn 
      Alignment       =   1  'Right Justify
      Caption         =   "Start With Camera &On"
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   4440
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin MSComctlLib.Slider sldAutoFocusDelay 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   2000
      SmallChange     =   500
      Max             =   15000
      SelStart        =   3
      Value           =   3
   End
   Begin MSComctlLib.Slider sldMovementDetail 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      Min             =   3
      Max             =   30
      SelStart        =   3
      Value           =   3
   End
   Begin MSComctlLib.Slider sldMovementTolerance 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   15
      Max             =   80
   End
   Begin MSComctlLib.Slider sldFrameRate 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   2000
      SmallChange     =   500
      Min             =   1
      Max             =   32000
      SelStart        =   3
      Value           =   3
   End
   Begin VB.Label lblFrameRate 
      Alignment       =   2  'Center
      Caption         =   "Frame Rate"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   3120
      Width           =   6375
   End
   Begin VB.Label lblFrameRateMin 
      Alignment       =   1  'Right Justify
      Caption         =   "Fast"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblFrameRateMax 
      Caption         =   "Slow"
      Height          =   255
      Left            =   7680
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.Line lnBreakFore 
      BorderColor     =   &H80000014&
      X1              =   240
      X2              =   8655
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line lnBreakBack 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   8640
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblAutoMax 
      Caption         =   "Long"
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblAutoMin 
      Alignment       =   1  'Right Justify
      Caption         =   "Short"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblAutoFocusDelay 
      Alignment       =   2  'Center
      Caption         =   "Auto-focus Delay"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Label lblMoveDetMax 
      Caption         =   "Low"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblMoveDetMin 
      Alignment       =   1  'Right Justify
      Caption         =   "High"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblMovementDetail 
      Alignment       =   2  'Center
      Caption         =   "Movement Detail"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label lblMoveTolMax 
      Caption         =   "Insensitive"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblMoveTolMin 
      Alignment       =   1  'Right Justify
      Caption         =   "Sensitive"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblMovementTolerance 
      Alignment       =   2  'Center
      Caption         =   "Level Of Movement Before Activating"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmCamOptions"
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
'   Description:    This is the options screen for the security camera
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
Private mintMovementTolerance   As Integer          'holds how sensitive the camera is to colour changes
Private mintMovementDetail      As Integer          'holds how many pixels we check per frame
Private mlngAutoFocusDelay      As Long             'holds how long we wait for the camera to auto-focus
Private mlngFrameRate           As Long             'holds the delay between frames
Private mblnStartCameraOn       As Boolean          'holds if the camera is on as soon as the program starts
Private mblnLogMovements        As Boolean          'holds whether or not to log the movements
Private mblnStartMin            As Boolean          'holds whether or not the program starts minized


'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    'exit the form
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'remember the new settings so that they will be returned
    mintMovementTolerance = sldMovementTolerance.Value
    mintMovementDetail = sldMovementDetail.Value
    mlngAutoFocusDelay = sldAutoFocusDelay.Value
    mlngFrameRate = sldFrameRate.Value
    mblnStartCameraOn = (chkStartCameraOn.Value = vbChecked)    'returns True or False
    mblnLogMovements = (chkLogMovements.Value = vbChecked)      'returns True or False
    mblnStartMin = (chkStartMin.Value = vbChecked)              'returns True or False
End Sub

Public Sub GetOptions(ByRef intMovementTolerance As Integer, _
                      ByRef intMovementDetail As Integer, _
                      ByRef lngAutoFocusDelay As Long, _
                      ByRef lngFrameRate As Long, _
                      ByRef blnStartCameraOn As Boolean, _
                      ByRef blnLogMovements As Boolean, _
                      ByRef blnStartMin As Boolean)
    'set the options and allow the user to change them
    
    
    'remember the values passed in
    mintMovementTolerance = intMovementTolerance
    mintMovementDetail = intMovementDetail
    mlngAutoFocusDelay = lngAutoFocusDelay
    mlngFrameRate = lngFrameRate
    mblnStartCameraOn = blnStartCameraOn
    mblnLogMovements = blnLogMovements
    mblnStartMin = blnStartMin
    
    'set the controls to reflect the settings
    sldMovementTolerance.Value = mintMovementTolerance
    sldMovementDetail.Value = mintMovementDetail
    sldAutoFocusDelay.Value = mlngAutoFocusDelay
    sldFrameRate.Value = mlngFrameRate
    chkStartCameraOn.Value = Abs(mblnStartCameraOn)
    chkLogMovements.Value = Abs(mblnLogMovements)
    chkStartMin.Value = Abs(mblnStartMin)
    
    'display the form
    Call Me.Show(vbModal)
    
    'get the new settings
    intMovementTolerance = mintMovementTolerance
    intMovementDetail = mintMovementDetail
    lngAutoFocusDelay = mlngAutoFocusDelay
    lngFrameRate = mlngFrameRate
    blnStartCameraOn = mblnStartCameraOn
    blnLogMovements = mblnLogMovements
    blnStartMin = mblnStartMin
End Sub
