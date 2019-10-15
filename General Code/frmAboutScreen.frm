VERSION 5.00
Begin VB.Form frmAboutScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Information"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picText 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line lnSpacer 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   135
      X2              =   4455
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lnSpacer 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAboutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'AUTHOR :   Eric O'Sullivan
' -----------------------------------------------
'DATE :     17 November 2001
' -----------------------------------------------
'CONTACT:   DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE :    About Screen (Deployable Version)
' -----------------------------------------------
'COMMENTS :
'This screen was first created on the and was
'intended for use in several future programs. The
'idea was that I should only have to create this
'screen once and be able to integrate it into any
'other project seemlessly. I wanted to do this
'instead of creating a new about screen for every
'project where I wanted one.
'
'A note on this About Screen :
'This screen requires the class clsBitmap
'operate the display.
'=================================================

'require variable declaration
Option Explicit

Private mstrAllText As String   'the scrolling text displayed
Private mblnStart   As Boolean  'is the text starting to scroll from the bottom

Private Sub cmdOk_Click()
    'exit screen
    Unload Me
End Sub

Private Sub Form_Load()
    'enter the text that we are going to display
    Call SetText
    mblnStart = True
    timText.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'exit the form
    timText.Enabled = False
    DoEvents
    Unload Me
End Sub

Private Sub timText_Timer()
    'This timer will scroll the animated text
    
    Const WAIT      As Integer = 50 'wait 15 ticks before drawing the next frame
    
    Static udtBmp           As clsBitmap    'holds the final bitmap drawn to the screen
    Static udtSurphase      As clsBitmap    'holds the gradients fading from the background colour to the forground colour
    Static udtMask          As clsBitmap    'holds the mask with the text to display on it
    Static udtFore          As clsBitmap    'holds the background colour to fade from
    Static intTextHeight    As Integer      'holds the total height of the text that we are going to display
    Static lngStartingTick  As Long         'used for controling the frame rate
    Static intScroll        As Integer      'the distance in pixels to scroll
    Static intLineNum       As Integer      'the number of lines in the text to scroll
    
    'reset the info in the timer every time the form loads
    If mblnStart Then
        Set udtBmp = Nothing
        Set udtSurphase = Nothing
        Set udtMask = Nothing
        Set udtFore = Nothing
        intTextHeight = 0
        lngStartingTick = 0
        intScroll = 0
        intLineNum = 0
    End If
    
    'instanciate the bitmaps
    If udtBmp Is Nothing Then
        Set udtBmp = New clsBitmap
    
        'set the bitmap dimensions and create them
        Call udtBmp.SetBitmap(picText.ScaleWidth, _
                              picText.ScaleHeight, _
                              picText.BackColor)
    
    Else
        'clear the data
        Call udtBmp.Cls
    End If
    
    'find out how much time it takes to draw a frame
    If lngStartingTick = 0 Then
        lngStartingTick = udtBmp.GetTick
    End If
    
    'get the number of lines of text to display
    If intLineNum = 0 Then
        intLineNum = LineCount(mstrAllText)
        intTextHeight = picText.TextHeight("I") * intLineNum
    End If
    
    'should we start scrolling the text from the bottom again
    intScroll = intScroll - 1
    If (intScroll < (-intTextHeight)) _
       Or (mblnStart) Then
        intScroll = picText.ScaleHeight
        mblnStart = False
    End If
    
    'only create the surphase if necessary
    If udtSurphase Is Nothing Then
        Set udtSurphase = New clsBitmap
        
        With udtSurphase
            Call .SetBitmap(picText.ScaleWidth, _
                            picText.ScaleHeight, _
                            picText.ForeColor)
        
            'create the surphase
            'text fade in
            Call .Gradient(picText.ForeColor, _
                           picText.BackColor, _
                           0, _
                           (udtSurphase.Height - ((intTextHeight / intLineNum) * 2)), _
                           udtSurphase.Width, _
                           (intTextHeight / intLineNum * 2), _
                           GradHorizontal)
            'text fade out
            Call .Gradient(picText.BackColor, _
                           picText.ForeColor, _
                           0, _
                           0, _
                           udtSurphase.Width, _
                           (intTextHeight / intLineNum) * 2, _
                           GradHorizontal)
        End With
        
        'create the foreground containing the colour to fade from
        Set udtFore = New clsBitmap
        With picText
            Call udtFore.SetBitmap(.ScaleWidth, _
                                   .ScaleHeight, _
                                   .BackColor)
        End With
    End If
    
    'we only need to create the text once
    If udtMask Is Nothing Then
        Set udtMask = New clsBitmap
        
        Call udtMask.SetBitmap(picText.ScaleWidth, _
                               intTextHeight + (picText.ScaleHeight * 2), _
                               vbBlack)
    
        'draw the white text onto the mask in black
        Call udtMask.DrawString(mstrAllText, _
                                picText.ScaleHeight, _
                                0, _
                                intTextHeight, _
                                udtBmp.Width, _
                                picText.Font, _
                                vbWhite)
    End If
    
    'copy the result to the screen
    Call udtBmp.MergeBitmaps(udtSurphase.hdc, _
                             udtFore.hdc, _
                             udtMask.hdc, _
                             intMaskY:=Abs(intScroll - picText.ScaleHeight))
    Call udtBmp.Paint(Me.hdc)
    
    'wait X ticks minus the time it took to draw the frame
    With udtSurphase
        'wait for the frame rate minus how long it took to draw the frame
        Call .Pause(WAIT - (.GetTick - lngStartingTick), True)
        lngStartingTick = .GetTick  'remember the point when we completed the frame
    End With
End Sub

Private Sub SetText()
    'This procedure is used to setting the text displayed in the picture box
    
    '" & vbCrLf & "
    
    'please note that ProductName can be set by going to
    'Project, Project Properties,Make tab. You should see a list box about
    'half way down on the left side. Scroll down until you come to
    'Product Name and enter some text into the text box on the right
    'side of the list box.
    mstrAllText = App.ProductName & vbCrLf & _
                  "Version " & App.Major & "." & _
                               App.Minor & "." & _
                               App.Revision & vbCrLf & _
                  "" & vbCrLf & _
                  "This program was made by" & vbCrLf & _
                  "Eric O'Sullivan." & vbCrLf & _
                  "" & vbCrLf & _
                  "Copyright " & Year(Date) & vbCrLf & _
                  "All rights reserved" & vbCrLf & _
                  "" & vbCrLf & _
                  "For more information, email" & vbCrLf & _
                  "DiskJunky@hotmail.com"
End Sub

Private Function LineCount(ByVal strText As String) _
                           As Integer
    'This function will return the number of lines
    'in the strText
    
    Dim intTemp     As Integer      'holds the current position in the string that we are searching
    Dim intCounter  As Integer      'used for counting the lines
    Dim intLastPos  As Integer      'holds the next position in the string to search from
    
    'start searching from the start of the string
    intLastPos = 1
    
    'search the string
    Do
        intTemp = intLastPos
        intLastPos = InStr(intLastPos + Len(vbCrLf), strText, vbCrLf)
        
        If intTemp <> intLastPos Then
            'a line was found
            intCounter = intCounter + 1
        End If
    Loop Until intLastPos = 0 'intLastPos will =0 when InStr cannot find any more occurances of vbCrlf
    
    'return the number of lines we have found
    LineCount = intCounter + 1
End Function
