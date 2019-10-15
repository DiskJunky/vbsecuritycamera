VERSION 5.00
Begin VB.Form frmLogsSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   120
      Picture         =   "frmLogsSplash.frx":0000
      ScaleHeight     =   810
      ScaleWidth      =   810
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading settings, please wait..."
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   405
      Width           =   5535
   End
End
Attribute VB_Name = "frmLogsSplash"
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
'   Description:    This screen is just used as a splash screen while the
'                   View Logs screen is loading up - it can take a bit of time
'                   if there is a lot of snapshots.
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
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Public Sub DispMsg(Optional ByVal strMsg As String)
    'This will display the current message and update the form
    lblMsg.Caption = strMsg
    Call lblMsg.Refresh
End Sub
