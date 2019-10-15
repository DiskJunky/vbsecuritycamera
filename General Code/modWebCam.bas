Attribute VB_Name = "modWebCam"
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   WebCam Module - General Use
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           23 May 2006
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Email:          DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This is to be used to help manage capturing data from a
'                   WebCam.
'  ---------------------------------------------------------------------------
'   Dependancies:   StdOle2.tlb         (standard vb library)
'                   AviCap32.dll        WebCam data capture dll
'  ---------------------------------------------------------------------------
'   References:     http://ej.bantz.com/video/
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------


'all variables must be declared
Option Explicit


'-------------------------------------------------------------------------------
'                               API DECLARATIONS
'-------------------------------------------------------------------------------
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
            (ByVal hwnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             lParam As Any) _
             As Long

Private Declare Function capCreateCaptureWindow _
        Lib "avicap32.dll" _
        Alias "capCreateCaptureWindowA" _
            (ByVal lpszWindowName As String, _
             ByVal dwStyle As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hWndParent As Long, _
             ByVal nID As Long) _
             As Long



'-------------------------------------------------------------------------------
'                            MODULE LEVEL CONSTANTS
'-------------------------------------------------------------------------------
Private Type VIDEOHDR
    lpData          As Long         'address of video buffer
    dwBufferLength  As Long         'size, in bytes, of the Data buffer
    dwBytesUsed     As Long         'see below
    dwTimeCaptured  As Long         'see below
    dwUser          As Long         'user-specific data
    dwFlags         As Long         'see below
    dwReserved(3)   As Long         'reserved; do not use
End Type



'-------------------------------------------------------------------------------
'                            MODULE LEVEL CONSTANTS
'-------------------------------------------------------------------------------
Private Const CONNECT       As Long = 1034
Private Const DISCONNECT    As Long = 1035
Private Const GET_FRAME     As Long = 1084
Private Const COPY          As Long = 1054


'-------------------------------------------------------------------------------
'                            MODULE LEVEL VARIABLES
'-------------------------------------------------------------------------------
Private hWndWebCam          As Long         'holds a handle to the WebCam


'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Public Sub StopCam()
    'stop capturing WebCam data
    
    Dim Result          As Long         'holds the value returned by the api call
    
    If (hWndWebCam <> 0) Then
        Result = SendMessage(hWndWebCam, DISCONNECT, 0, 0)
        hWndWebCam = 0
    End If
End Sub

Public Sub StartCam(ByVal ParentWnd As Long, _
                    Optional ByRef ErrMsg As String)
    'start capturing WebCam data
    
    Dim Result          As Long         'holds the value returned by the api call
    
    
    'set the default error value (ie, none)
    ErrMsg = ""
    
    'if we're already connected then don't do anything
    If (hWndWebCam <> 0) Then
        Exit Sub
    End If
    
    hWndWebCam = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, ParentWnd, 0)
    If (hWndWebCam <> 0) Then
        'send the connection message
        Result = SendMessage(hWndWebCam, CONNECT, 0, 0)
        If (Result = 0) Then
            Call StopCam
            ErrMsg = "Unable to connect to a WebCam. Please make sure that it is connected and turned on"
        End If
    Else
        ErrMsg = "Unable to connect to WebCam. Please make sure that it is connected and turned on"""
    End If
End Sub

Public Sub GetFromCam(ByRef PaintOnto As StdPicture, _
                      Optional ByRef ErrMsg As String)
    'This will capture a single screenshot from the webcam and paint it onto the specified
    'picture box
    
    
    Dim Result              As Long         'holds the returned value from an API call
    
    
    'set the default value
    ErrMsg = ""
    
    'is anything connected
    If (hWndWebCam = 0) Then
        Exit Sub
    End If
    
    'prepare the webcam
    Result = SendMessage(hWndWebCam, GET_FRAME, 0, 0)
    If (Result = 0) Then
        Call StopCam
        ErrMsg = "WebCam was disconnected"
        Exit Sub
    End If
    
    'copy the frame (screenshot) onto the specified picture box
    Result = SendMessage(hWndWebCam, COPY, 0, 0)
    Set PaintOnto = Clipboard.GetData
End Sub

Public Function IsWebCamOn() As Boolean
    'This will return whether or not the WebCam is connected and turned on
    
    Dim Result              As Long         'holds the returned value from an API call
    
    
    'set the default value
    IsWebCamOn = False
    
    'is anything connected
    If (hWndWebCam = 0) Then
        Exit Function
    End If
    
    'prepare the webcam
    Result = SendMessage(hWndWebCam, COPY, 0, 0)
    If (Result <> 0) Then
        IsWebCamOn = True
    End If
End Function
