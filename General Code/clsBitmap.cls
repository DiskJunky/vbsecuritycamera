VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Creates and uses a bitmap effiently and also provides access to some graphical api functions and procedures and some higher-level graphical sub-routines"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'==========================================================================
'   AUTHOR   :  Eric O'Sullivan
' ------------------------------------------------------------------------
'   DATE     :  11 November 2001
' ------------------------------------------------------------------------
'   CONTACT  :  DiskJunky@hotmail.com
' ------------------------------------------------------------------------
'   TITLE    :  Bitmap Class
' ------------------------------------------------------------------------
'   COMMENTS :
'               This module was made for using api graphics functions in
'   your programs. With the following api calls and function and
'   procedures written by me, you have to tools to do almost anything.
'   The only api function listed here that is not directly used by any
'   piece of code in this module is BitBlt. You have the tools to create
'   and manipulate graphics, but it is still necessary to display them
'   manually. The functions themselves mostly need hDc or a handle to
'   work. You can find this hDc in both a forms and pictureboxes'
'   properties. I have also set up a data type called BitmapStruc. For my
'   programs, I have used this structure almost exclusivly for the
'   graphics. The structure holds all the information needed to reference
'   a bitmap created using this module (CreateNewBitmap, DeleteBitmap).
'   Please keep in mind that any object (bitmap, brush, pen etc) needs to
'   be deleted after use or else it will stay in memory until the program
'   is finished. Not doing so will eventually cause your program to take
'   up ALL your computers recources. Also for anyone using optional
'   paramters, it is probably better to use a default parameter values to
'   determine whether or not a parameter has been passed than the
'   function "IsMissing()".
'
'                        ---------------------
' 19 July 2002
' The comments above were for the module APIGraphics. This class is a
' direct conversion from that module and is not completely tested as yet.
' For the purposes of this program (whatever this class happens to
' currently be in), the class will work correctly and is tested for the
' program purpose only. Currently known, there are several dependancies,
'   MsImg32.dll,        StdOle2.tlb,        ATL.dll
'                        ---------------------
'
'Thank you,
'Eric
'==========================================================================

'require variable declaration
Option Explicit

'--------------------------------------------------------------------------
'                           API DECLARATIONS
'--------------------------------------------------------------------------
'These procedures/functions are sorted alphabetically.

'this will alphablend two bitmaps by a specified
'blend amount.
Private Declare Function APIAlphaBlend _
       Lib "msimg32" _
       Alias "AlphaBlend" _
            (ByVal hDcDest As Long, _
             ByVal intLeftDest As Integer, _
             ByVal intTopDest As Integer, _
             ByVal intWidthDest As Integer, _
             ByVal intHeightDest As Integer, _
             ByVal hDcSource As Long, _
             ByVal intLeftSource As Integer, _
             ByVal intTopSource As Integer, _
             ByVal intWidthSource As Integer, _
             ByVal intHeightSource As Integer, _
             ByVal lngBlendFunctionStruc As Long) _
             As Long

'converts a HiMetric value to a pixel value
Private Declare Sub AtlHiMetricToPixel _
        Lib "atl" _
            (ByVal lpSizeInHiMetric As Long, _
             ByVal lpSizeInPix As Long)

'This is used to copy bitmaps
Private Declare Function BitBlt _
       Lib "gdi32" _
            (ByVal hDestDC As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal dwRop As Long) _
             As Long

'used to change various windows settings
Private Declare Function ChangeDisplaySettings _
       Lib "user32" _
       Alias "ChangeDisplaySettingsA" _
            (ByRef wef As Any, _
             ByVal i As Long) _
             As Long

'creates a brush object which can be applied to a bitmap
Private Declare Function CreateBrushIndirect _
       Lib "gdi32" _
            (lpLogBrush As LogBrush) _
             As Long

'creates a colourspace object which can be applied to a bitmap
Private Declare Function CreateColorSpace _
       Lib "gdi32" _
       Alias "CreateColorSpaceA" _
            (lplogcolorspace As LogColorSpace) _
             As Long

'the will create a bitmap compatable with the passed hDc
Private Declare Function CreateCompatibleBitmap _
       Lib "gdi32" _
            (ByVal hdc As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long) _
            As Long

'this create a compatable device context with the specified
'windows handle
Private Declare Function CreateCompatibleDC _
       Lib "gdi32" _
            (ByVal hdc As Long) _
             As Long

'creates an elliptical region in a hDc
Private Declare Function CreateEllipticRgn _
       Lib "gdi32" _
            (ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'creates an elliptical region in a hDc
Private Declare Function CreateEllipticRgnIndirect _
       Lib "gdi32" _
            (EllRect As Rect) _
             As Long

'creates a font compatable with the specified device context
Private Declare Function CreateFontIndirect _
       Lib "gdi32" _
       Alias "CreateFontIndirectA" _
            (lpLogFont As LogFont) _
             As Long

'creates a palette from the specified information
Private Declare Function CreatePalette _
        Lib "gdi32" _
            (lpLogPalette As LOGPALETTE) _
             As Long

'creates a pen that can be applied to a hDc
Private Declare Function CreatePen _
       Lib "gdi32" _
            (ByVal nPenStyle As Long, _
             ByVal nWidth As Long, _
             ByVal crColor As Long) _
             As Long

'creates a pen that can be applied to a hDc
Private Declare Function CreatePenIndirect _
       Lib "gdi32" _
            (lpLogPen As LogPen) _
             As Long

'creates a rectangular region on a hDc
Private Declare Function CreateRectRgn _
       Lib "gdi32" _
            (Left As Integer, _
             Top As Integer, _
             Right As Integer, _
             Bottom As Integer) _
             As Long

'creates a rectangular region to a hDc
Private Declare Function CreateRectRgnIndirect _
        Lib "gdi32" _
            (lpRect As Rect) _
             As Long

'creates a solid colour brush to be applied to
'a bitmap
Private Declare Function CreateSolidBrush _
       Lib "gdi32" _
            (ColorRef As Long) _
             As Long

'removes a device context from memory
Private Declare Function DeleteDC _
       Lib "gdi32" _
            (ByVal hdc As Long) _
             As Long

'removes an object such as a brush or bitmap from memory
Private Declare Function DeleteObject _
       Lib "gdi32" _
            (ByVal hObject As Long) _
             As Long

'this draws the animated minimize/maximize rectangeles
Private Declare Function DrawAnimatedRects _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal idAni As Long, _
             lprcFrom As Rect, _
             lprcTo As Rect) _
             As Long

'this will draw a 2D or 3D edge on the specified DC
Private Declare Function DrawEdge _
        Lib "user32" _
            (ByVal hdc As Long, _
             qrc As Rect, _
             ByVal edge As Long, _
             ByVal grfFlags As Long) _
             As Long

'This will draw a frame border in the specified co-ordinates
Private Declare Function DrawFrameControl _
        Lib "user32" _
            (ByVal hdc As Long, _
             lpRect As Rect, _
             ByVal un1 As Long, _
             ByVal un2 As Long) _
             As Long

'this draws an icon onto a surphase (eg, a bitmap)
Private Declare Function DrawIconEx _
       Lib "user32" _
            (ByVal hdc As Long, _
             ByVal xLeft As Long, _
             ByVal yTop As Long, _
             ByVal hIcon As Long, _
             ByVal cxWidth As Long, _
             ByVal cyWidth As Long, _
             ByVal istepIfAniCur As Long, _
             ByVal hbrFlickerFreeDraw As Long, _
             ByVal diFlags As Long) _
             As Long

'this draws text onto the bitmap
Private Declare Function DrawText _
       Lib "user32" _
       Alias "DrawTextA" _
            (ByVal hdc As Long, _
             ByVal lpStr As String, _
             ByVal nCount As Long, _
             lpRect As Rect, _
             ByVal wFormat As Long) _
             As Long

'this draws an ellipse onto the bitmap
Private Declare Function Ellipse _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Boolean

'this will set the display settings
Private Declare Function EnumDisplaySettings _
       Lib "user32" _
       Alias "EnumDisplaySettingsA" _
            (ByVal a As Long, _
             ByVal B As Long, _
             wef As DEVMODE) _
             As Boolean

'this provides more control than the CreatePen function
Private Declare Function ExtCreatePen _
       Lib "gdi32" _
            (ByVal dwPenStyle As Long, _
             ByVal dwWidth As Long, _
             lplb As LogBrush, _
             ByVal dwStyleCount As Long, _
             lpStyle As Long) _
             As Long

'select a clipping region to the specified DC
Private Declare Function ExtSelectClipRgn _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hRgn As Long, _
             ByVal fnMode As Long) _
             As Long

'this will fill the rectangle with the brush applied
'to the hDc
Private Declare Function FillRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             Fill As Rect, _
             hBrush As Long) _
             As Integer

'this will fill a region with the brush specified
Private Declare Function FillRgn _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hRgn As Long, _
             hBrush As Long) _
             As Boolean

'this will find a window based on its class and
'window name
Private Declare Function FindWindow _
       Lib "user32" _
       Alias "FindWindowA" _
            (ByVal lpClassName As String, _
             ByVal lpWindowName As String) _
             As Long

'this will flash the specified window
Private Declare Function FlashWindowApi _
        Lib "user32" _
        Alias "FlashWindow" _
            (ByVal hWnd As Long, _
             ByVal bInvert As Long) _
             As Long

'this will return the handle of the top-most window
Private Declare Function GetActiveWindow _
       Lib "user32" _
            () _
             As Long

'this will get the state of any specified key - even
'the mouse buttons
Private Declare Function GetAsyncKeyState _
       Lib "user32" _
            (ByVal vKey As Long) _
             As Integer

'this will return each bit of the bitmaps colour data
Private Declare Function GetBitmapBits _
        Lib "gdi32" _
            (ByVal hBitmap As Long, _
             ByVal dwCount As Long, _
             lpBits As Any) _
             As Long

'this will get the dimensions of the specified bitmap
Private Declare Function GetBitmapDimensionEx _
       Lib "gdi32" _
            (ByVal hBitmap As Long, _
             lpDimension As SizeType) _
             As Long

'this will return the current brush origin applied to
'the specified device context
Private Declare Function GetBrushOrgEx _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             lpPoint As PointAPI) _
             As Long

'this will get the class name from the handle of the
'window specified
Private Declare Function GetClassName _
       Lib "user32" _
       Alias "GetClassNameA" _
            (ByVal hWnd As Long, _
             ByVal lpClassName As String, _
             ByVal nMaxCount As Long) _
             As Long

'this will capture a screen shot of the specified
'area of the client
Private Declare Function GetClientRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             lpRect As Rect) _
             As Long

'this gets the cursors icon picture
Private Declare Function GetCursor _
       Lib "user32" _
            () _
             As Long

'this gets the position of the cursor on the screen
'(given in pixels)
Private Declare Function GetCursorPos _
       Lib "user32" _
            (lpPoint As PointAPI) _
             As Long

'gets a hDc of the specified window
Private Declare Function GetDC _
       Lib "user32" _
            (ByVal hWnd As Long) _
            As Long

'gets the entire screen area
Private Declare Function GetDesktopWindow _
       Lib "user32" _
            () _
             As Long

'this will get the current devices' capabilities
Private Declare Function GetDeviceCaps _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal nIndex As Long) _
             As Long

'returns a hanlde to the window that is currently
'in the foreground
Private Declare Function GetForegroundWindow _
        Lib "user32" _
            () As Long

'get the last error to occur from within the api
Private Declare Function GetLastError _
       Lib "kernel32" _
            () _
             As Long

'get the handle of the menu bar on a window
Private Declare Function GetMenu _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'Get the sub menu ID number. This is used to
'reference sub menus along with their handle
Private Declare Function GetMenuItemID _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get information about the specified object such as
'its dimensions etc.
Private Declare Function GetObjectAPI _
       Lib "gdi32" _
       Alias "GetObjectA" _
            (ByVal hObject As Long, _
             ByVal nCount As Long, _
             lpObject As Any) _
             As Long

'get the colour of the specified pixel
Private Declare Function GetPixel _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'this will get the stretch mode of the bitmap
Private Declare Function GetStretchBltMode _
        Lib "gdi32" _
            (ByVal hdc As Long) _
             As Long

'this will get the handle of a specified
'sub menu using the handle of the menu
'and the item ID
Private Declare Function GetSubMenu _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
             As Long

'get the specified system colour
Private Declare Function GetSysColor _
        Lib "user32" _
            (ByVal nIndex As Long) _
             As Long

'gets the palette information for the system
Private Declare Function GetSystemPaletteEntries _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal wStartIndex As Long, _
             ByVal wNumEntries As Long, _
             lpPaletteEntries As PALETTEENTRY) _
             As Long

'this will return the size of a string in a SIZE
'structure
Private Declare Function GetTextExtentPoint _
        Lib "gdi32" _
        Alias "GetTextExtentPoint32A" _
        (ByVal hdc As Long, ByVal lpsz As String, _
         ByVal cbString As Long, _
         lpSize As SizeType) _
         As Long

'get the dimensions of the applied text metrics for
'the device context
Private Declare Function GetTextMetrics _
       Lib "gdi32" _
       Alias "GetTextMetricsA" _
            (ByVal hdc As Long, _
             lpMetrics As TEXTMETRIC) _
             As Long

'returns the amount of time windows has been active for
'in milliseconds (sec/1000)
Private Declare Function GetTickCount _
       Lib "kernel32" _
            () _
             As Long  'very usefull timing function ;)

'returns extended information about the version of the
'operating system currently running
Private Declare Function GetVersionEx _
        Lib "kernel32" _
        Alias "GetVersionExA" _
            (ByRef lpVersionInformation As OSVERSIONINFO) _
             As Long

'retrieves the handle of a window that has the
'specified relationship to the specified window.
Private Declare Function GetWindow _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal wCmd As Long) _
             As Long

'takes a snapshot of the entire window area
'including title bar, borders etc
Private Declare Function GetWindowDC _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'gets the area the specified window takes up
Private Declare Function GetWindowRect _
       Lib "user32" _
            (ByVal hWnd As Long, _
             lpRect As Rect) _
             As Long

'increases the size of a rect structure
Private Declare Function InflateRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'gets any intersection of two rectangles
Private Declare Function IntersectRect _
       Lib "user32" _
            (lpDestRect As Rect, _
             lpSrc1Rect As Rect, _
             lpSrc2Rect As Rect) _
             As Long

'This inverts the colours in the specified
'region
Private Declare Function InvertRect _
        Lib "user32" _
            (ByVal hdc As Long, _
             lpRect As Rect) _
             As Long

'This will return non-zero if the rect is
'NOT empty
Private Declare Function IsRectEmptyApi _
        Lib "user32" _
        Alias "IsRectEmpty" _
            (lpRect As Rect) _
             As Long

'draws a line from the current position to the
'specified co-ordinates
Private Declare Function LineTo _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             XEnd As Integer, _
             YEnd As Integer) _
             As Boolean

'this will load a bitmap directly from a file (NT compatable)
Private Declare Function LoadNTBitmap _
       Lib "user32" _
       Alias "LoadBitmapA" _
            (ByVal hInstance As Long, _
             ByVal lpBitmapName As Any) _
             As Long

'this will load a cursor into a device context
Private Declare Function LoadCursor _
       Lib "user32" _
       Alias "LoadCursorA" _
            (ByVal hInstance As Long, _
             ByVal lpCursorName As Any) _
             As Long

'this will load an image into a device context (NOT NT compatable)
Private Declare Function LoadImage _
       Lib "user32" _
       Alias "LoadImageA" _
            (ByVal hInst As Long, _
             ByVal lpsz As String, _
             ByVal un1 As Long, _
             ByVal n1 As Long, _
             ByVal n2 As Long, _
             ByVal un2 As Long) _
             As Long

'This stops the specified window from updating
'its display. This is mainly used to help cut out
'flicker but does not work on controls.
Private Declare Function LockWindowUpdate _
       Lib "user32" _
            (ByVal hwndLock As Long) _
             As Long
'changes some of a menu's properties
Private Declare Function ModifyMenu _
        Lib "user32" _
        Alias "ModifyMenuA" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal wIDNewItem As Long, _
             ByVal lpString As Any) _
             As Long

'moves the current position to the specified
'point
Private Declare Function MoveToEx _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByRef lpPoint As PointAPI) _
             As Boolean

'multiplies two numbers and divides them by a third
'numbers
Private Declare Function MulDiv _
       Lib "kernel32" _
            (ByVal nNumber As Long, _
             ByVal nNumerator As Long, _
             ByVal nDenominator As Long) _
             As Long

'This will increase or decrease a rectangles
'co-ordinates by the specified amount. Usefull
'for moving graphic blocks as rect structures.
Private Declare Function OffsetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'creates a new picture object accoring to a PictDesc structure
Private Declare Function OleCreatePictureIndirect _
        Lib "olepro32.dll" _
            (PicDesc As PicBmp, _
             RefIID As GUID, _
             ByVal fPictureOwnsHandle As Long, _
             IPic As IPicture) _
             As Long

'this will paint the desktop pattern or wallpaper onto the
'specified DC
Private Declare Function PaintDesktop _
        Lib "user32" _
            (ByVal hdc As Long) _
             As Long

'Pattern Blitter. Used to draw a pattern onto
'a device context
Private Declare Function PatBlt _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal dwRop As Long) _
             As Long

'This draws a polygon onto a device context
'useing an array.
Private Declare Function Polygon _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines to the specifed
'points
Private Declare Function Polyline _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             lpPoint As PointAPI, _
             ByVal nCount As Long) _
             As Long

'This will draw a set of lines starting from
'the current position on the device context.
Private Declare Function PolylineTo _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             lppt As PointAPI, _
             ByVal cCount As Long) _
             As Long

'maps the selected palette to the system palette
Private Declare Function RealizePalette _
        Lib "gdi32" _
            (ByVal hdc As Long) _
             As Long

'This draws a rectangle onto the device
'context
Private Declare Function Rectangle _
       Lib "gdi32" _
            (ByVal hWnd As Long, _
             X1 As Integer, _
             Y1 As Integer, _
             X2 As Integer, _
             Y2 As Integer) _
             As Long

'This will release a device context from
'memory. Not to be confused with DeleteDC
Private Declare Function ReleaseDC _
       Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal hdc As Long) _
             As Long

'this will draw a round-cornered rectangle
'onto the specified device context
Private Declare Function RoundRect _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal Left As Long, _
             ByVal Top As Long, _
             ByVal Right As Long, _
             ByVal Bottom As Long, _
             ByVal X3 As Long, _
             ByVal Y3 As Long) _
             As Long

'this can convert entire type structures
'to other types like a Long
Private Declare Sub RtlMoveMemory _
        Lib "kernel32.dll" _
            (Destination As Any, _
             Source As Any, _
             ByVal Length As Long)

'this will select a region for drawing to the specified
'device context
Private Declare Function SelectClipRgn _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hRgn As Long) _
             As Long

'this will select the specified object to
'a window or device context
Private Declare Function SelectObject _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hObject As Long) _
             As Long

'selects the palette to the device context
Private Declare Function SelectPalette _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hPalette As Long, _
             ByVal bForceBackground As Long) _
             As Long

'this will set the bitmaps colour data to the
'new colour information
Private Declare Function SetBitmapBits _
        Lib "gdi32" _
            (ByVal hBitmap As Long, _
             ByVal dwCount As Long, _
             lpBits As Any) _
             As Long

'This sets the background colour on a bitmap
Private Declare Function SetBkColor _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal crColor As Long) _
             As Long

'This sets the background mode on a bitmap
'(eg, transparent, solid etc)
Private Declare Function SetBkMode _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal nBkMode As Long) _
             As Long

'The SetBrushOrgEx function sets the brush origin
'that GDI assigns to the next brush an application
'selects into the specified device context.
Private Declare Function SetBrushOrgEx _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal nXOrg As Integer, _
             ByVal nYOrg As Integer, _
             lppt As PointAPI) _
             As Integer

'The color adjustment values are used to adjust
'the input color of the source bitmap for calls
'to the StretchBlt and StretchDIBits functions
'when HALFTONE mode is set.
Private Declare Function SetColorAdjustment _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             lpca As COLORADJUSTMENT) _
             As Long

'sets the colourspace to a device context
Private Declare Function SetColorSpace _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal hcolorspace As Long) _
             As Long

'sets the position of the cursor on the screen
Private Declare Function SetCursorPos _
        Lib "user32" _
            (ByVal X As Long, _
             ByVal Y As Long) _
             As Long

'sets the specified window to the top of the zorder queue
Private Declare Function SetForegroundWindow _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'sets the bitmap of a menu
Private Declare Function SetMenuItemBitmaps _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long, _
             ByVal hBitmapUnchecked As Long, _
             ByVal hBitmapChecked As Long) _
             As Long

'sets the current information about the selected menu
Private Declare Function SetMenuItemInfo _
        Lib "user32" _
        Alias "SetMenuItemInfoA" _
            (ByVal hMenu As Long, _
             ByVal uItem As Long, _
             ByVal fByPosition As Long, _
             lpmii As MENUITEMINFO) _
             As Long

'sets the colour of the specified pixel
Private Declare Function SetPixel _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal crColor As Long) _
             As Long

'sets the rectangles size and position
Private Declare Function SetRect _
       Lib "user32" _
            (lpRect As Rect, _
             ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'this will set a rectangular region in the specified
'device context
Private Declare Function SetRectRgn _
        Lib "gdi32" _
            (ByVal hRgn As Long, _
             ByVal X1 As Long, _
             ByVal Y1 As Long, _
             ByVal X2 As Long, _
             ByVal Y2 As Long) _
             As Long

'set the stretching mode for the api StretchBlt. Note
'that a call to SetBrushOrgEx must be made after calling
'SetStretchBltMode or a brush misalignment occurs.
Private Declare Function SetStretchBltMode _
        Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal nStretchMode As Integer) _
             As Integer

'set a system colour
Private Declare Function SetSysColors _
        Lib "user32" _
            (ByVal nChanges As Long, _
             lpSysColor As Long, _
             lpColorValues As Long) _
             As Long

'sets the current text colour
Private Declare Function SetTextColor _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal crColor As Long) _
             As Long

'show or hide the cursor
Private Declare Function ShowCursor _
        Lib "user32" _
            (ByVal bShow As Long) _
             As Long

'pauses the execution of the programs thread
'for a specified amount of milliseconds
Private Declare Sub Sleep _
       Lib "kernel32" _
            (ByVal dwMilliseconds As Long)

'used to stretch or shrink a bitmap
Private Declare Function StretchBlt _
       Lib "gdi32" _
            (ByVal hdc As Long, _
             ByVal X As Long, _
             ByVal Y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal nSrcWidth As Long, _
             ByVal nSrcHeight As Long, _
             ByVal dwRop As Long) _
             As Long

'gets some system information
Private Declare Function SystemParametersInfo _
        Lib "user32" _
        Alias "SystemParametersInfoA" _
            (ByVal uAction As Long, _
             ByVal uParam As Long, _
             lpvParam As Any, _
             ByVal fuWinIni As Long) _
             As Long

'copyies/resizes and image using the specified Transparent colour
Private Declare Function TransparentBlt _
        Lib "msimg32.dll" _
            (ByVal hdc As Long, _
             ByVal lngX As Long, _
             ByVal lngY As Long, _
             ByVal lngWidth As Long, _
             ByVal lngHeight As Long, _
             ByVal hSrcDC As Long, _
             ByVal xSrc As Long, _
             ByVal ySrc As Long, _
             ByVal nSrcWidth As Long, _
             ByVal nSrcHeight As Long, _
             ByVal crTransparent As Long) _
             As Boolean

'resets the specified object, eg a palette
Private Declare Function UnrealizeObject _
        Lib "gdi32" _
            (ByVal lngObject As Long) _
             As Long

'--------------------------------------------------------------------------
'                               ENUMERATORS
'--------------------------------------------------------------------------

'the direction of the gradient
Public Enum GradientTo
    GradHorizontal = 0
    GradVertical = 1
End Enum

'in twips or pixels
Public Enum Scaling
    InTwips = 0
    InPixels = 1
End Enum

'The key values of the mouse buttons
Public Enum MouseKeys
    MouseLeft = 1
    MouseRight = 2
    MouseMiddle = 4
End Enum

'text alignment constants
Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

'bitmap flip constants
Public Enum FlipType
    FlipHorizontally = 0
    FlipVertically = 1
End Enum

'image load constants
Public Enum LoadType
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    'IMAGE_ENHMETAFILE  = 3
End Enum

Public Enum EnumWindowHandles
    WND_TOP_WINDOW = 0
    WND_DESKTOP_WINDOW = 1
    WND_ACTIVE_WINDOW = 2
End Enum

'the state of a key
Public Enum EnumKeyState
    AS_KEY_DOWN = 0
    AS_KEY_PRESSED = 1
    AS_KEY_UP = 2
    AS_KEY_TOGGLED = 3
End Enum

'rotate bitmap constants
Public Enum RotateType
    RotateRight = 0
    RotateLeft = 1
    Rotate180 = 2
End Enum

Public Enum PaintMode
    'BitBlt constants
    P_AND = &H8800C6  ' (DWORD) dest = source AND dest
    P_COPY = &HCC0020 ' (DWORD) dest = source
    P_ERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
    P_INVERT = &H660046       ' (DWORD) dest = source XOR dest
    P_PAINT = &HEE0086        ' (DWORD) dest = source OR dest
    P_MERGE_COPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
    P_MERGE_PAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
    P_NOT_COPY = &H330008      ' (DWORD) dest = (NOT source)
    P_NOT_ERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
End Enum

'system colour constants
Public Enum SystemColours
     col_Scrollbar = 0              'The Scrollbar colour
     col_Background = 1             'Colour of the background with no wallpaper
     col_ActiveCaption = 2          'Caption of Active Window
     col_InactiveCaption = 3        'Caption of Inactive window
     col_MenuBar = 4                'Menu
     col_Window = 5                 'Windows background
     col_WindowFrame = 6            'Window frame
     col_MenuText = 7               'Window Text
     col_WindowText = 8             '3D dark shadow (Win95)
     col_CaptionText = 9            'Text in window caption
     col_ActiveBorder = 10          'Border of active window
     col_InactiveBorder = 11        'Border of inactive window
     col_AppWorkSpace = 12          'Background of MDI desktop
     col_Highlight = 13             'Selected item background
     col_HighlightText = 14         'Selected menu item
     col_BtnFace = 15               'Button
     col_BtnSadow = 16              '3D shading of button
     col_GrayText = 17              'Grey text, of zero if dithering is used.
     col_BtnText = 18               'Button text
     col_InactiveCaptionText = 19   'Text of inactive window
     col_BtnHighlight = 20          '3D highlight of button
     col_ActiveCaption2 = 27        'Win98 only: 2nd active window color
     col_InactiveCaption2 = 28      'Win98 only: 2nd inactive window color
End Enum

'colour adjustment values
Public Enum EnumColourNegative
    neg_Normal = 0                  'display picture as normal
    neg_Negative = 1                'display a negative of the picture
End Enum

'colour adjustment light types
Public Enum EnumLightColour
    lit_Tungsten = 1
    lit_DayLightNoon = 2
    lit_DayLightNTSC = 3
    lit_PaperNormal = 4
    lit_PaperBond = 5
    lit_DayLightStandard = 6
    lit_DayLightNorthern = 7
    lit_Fluorescent = 8
End Enum

'SetStretchBltMode constants
Public Enum EnumStretchModes
    BLACKONWHITE = 1
    STRETCH_ANDSCANS = 1
    WHITEONBLACK = 2
    STRETCH_ORSCANS = 2
    COLORONCOLOR = 3
    STRETCH_DELETESCANS = 3
    HALFTONE = 4
    STRETCH_HALFTONE = 4
End Enum

'edge constants
Public Enum EnumEdges
    BDR_INNER = &HC
    BDR_OUTER = &H3
    BDR_RAISED = &H5
    BDR_RAISEDINNER = &H4
    BDR_RAISEDOUTER = &H1
    BDR_SUNKEN = &HA
    BDR_SUNKENINNER = &H8
    BDR_SUNKENOUTER = &H2
    BDR_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    BDR_OUTLINE = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum EnumEdgeSides
    BF_BOTTOM = &H8
    BF_LEFT = &H1
    BF_RIGHT = &H4
    BF_TOP = &H2
    BF_DIAGONAL = &H10
    BF_FLAT = &H4000     ' For flat rather than 3-D borders.
    BF_MIDDLE = &H800    ' Fill in the middle.
    BF_MONO = &H8000     ' For monochrome borders.
    BF_SOFT = &H1000     ' Use for softer buttons.
    
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
End Enum

Public Enum EnumFrameStyle
    DFC_BUTTON = 4
    DFC_CAPTION = 1
    DFC_MENU = 2
    DFC_SCROLL = 3
End Enum

Public Enum EnumFrameState
    DFCS_ADJUSTRECT = &H2000
    DFCS_BUTTON3STATE = &H8
    DFCS_BUTTONCHECK = &H0
    DFCS_BUTTONPUSH = &H10
    DFCS_BUTTONRADIO = &H4
    DFCS_BUTTONRADIOIMAGE = &H1
    DFCS_BUTTONRADIOMASK = &H2
    DFCS_CAPTIONCLOSE = &H0
    DFCS_CAPTIONHELP = &H4
    DFCS_CAPTIONMAX = &H2
    DFCS_CAPTIONMIN = &H1
    DFCS_CAPTIONRESTORE = &H3
    DFCS_CHECKED = &H400
    DFCS_FLAT = &H4000
    DFCS_INACTIVE = &H100
    DFCS_MENUARROW = &H0
    DFCS_MENUARROWRIGHT = &H4
    DFCS_MENUBULLET = &H2
    DFCS_MENUCHECK = &H1
    DFCS_MONO = &H8000
    DFCS_PUSHED = &H200
    DFCS_SCROLLCOMBOBOX = &H5
    DFCS_SCROLLDOWN = &H1
    DFCS_SCROLLLEFT = &H2
    DFCS_SCROLLRIGHT = &H3
    DFCS_SCROLLSIZEGRIP = &H8
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
    DFCS_SCROLLUP = &H0
End Enum


'--------------------------------------------------------------------------
'                           USER DEFINED TYPES
'--------------------------------------------------------------------------

'APIAlphaBlend information for bitmaps
Private Type BLENDFUNCTION
    bytBlendOp              As Byte 'currently the only blend op supported by windows 98+ is AC_SRC_OVER
    bytBlendFlags           As Byte 'must be left blank
    bytSourceConstantAlpha  As Byte 'the amount to blend by. Must be between 0 and 255
    bytAlphaFormat          As Byte 'don't set this. If you wish more infor, go to "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/bitmaps_3b3m.asp"
End Type

'Bitmap structue for menu information
Private Type MENUITEMINFO
    cbSize                  As Long
    fMask                   As Long
    fType                   As Long
    fState                  As Long
    wID                     As Long
    hSubMenu                As Long
    hbmpChecked             As Long
    hbmpUnchecked           As Long
    dwItemData              As Long
    dwTypeData              As String
    cch                     As Long
End Type

'size structure
Private Type SizeType
    cx                      As Long
    cy                      As Long
End Type

'Text metrics
Private Type TEXTMETRIC
    tmHeight                As Long
    tmAscent                As Long
    tmDescent               As Long
    tmInternalLeading       As Long
    tmExternalLeading       As Long
    tmAveCharWidth          As Long
    tmMaxCharWidth          As Long
    tmWeight                As Long
    tmOverhang              As Long
    tmDigitizedAspectX      As Long
    tmDigitizedAspectY      As Long
    tmFirstChar             As Byte
    tmLastChar              As Byte
    tmDefaultChar           As Byte
    tmBreakChar             As Byte
    tmItalic                As Byte
    tmUnderlined            As Byte
    tmStruckOut             As Byte
    tmPitchAndFamily        As Byte
    tmCharSet               As Byte
End Type


Private Type COLORADJUSTMENT
    caSize                  As Integer
    caFlags                 As Integer
    caIlluminantIndex       As Integer
    caRedGamma              As Integer
    caGreenGamma            As Integer
    caBlueGamma             As Integer
    caReferenceBlack        As Integer
    caReferenceWhite        As Integer
    caContrast              As Integer
    caBrightness            As Integer
    caColorfulness          As Integer
    caRedGreenTint          As Integer
End Type

Private Type CIEXYZ
    ciexyzX                 As Long
    ciexyzY                 As Long
    ciexyzZ                 As Long
End Type

Private Type CIEXYZTRIPLE
    ciexyzRed               As CIEXYZ
    ciexyzGreen             As CIEXYZ
    ciexyBlue               As CIEXYZ
End Type

Private Type LogColorSpace
    lcsSignature            As Long
    lcsVersion              As Long
    lcsSize                 As Long
    lcsCSType               As Long
    lcsIntent               As Long
    lcsEndPoints            As CIEXYZTRIPLE
    lcsGammaRed             As Long
    lcsGammaGreen           As Long
    lcsGammaBlue            As Long
    lcsFileName             As String * 26 'MAX_PATH
End Type

'display settings (800x600 etc)
Private Type DEVMODE
    dmDeviceName            As String * 32
    dmSpecVersion           As Integer
    dmDriverVersion         As Integer
    dmSize                  As Integer
    dmDriverExtra           As Integer
    dmFields                As Long
    dmOrientation           As Integer
    dmPaperSize             As Integer
    dmPaperLength           As Integer
    dmPaperWidth            As Integer
    dmScale                 As Integer
    dmCopies                As Integer
    dmDefaultSource         As Integer
    dmPrintQuality          As Integer
    dmColor                 As Integer
    dmDuplex                As Integer
    dmYResolution           As Integer
    dmTTOption              As Integer
    dmCollate               As Integer
    dmFormName              As String * 32
    dmUnusedPadding         As Integer
    dmBitsPerPel            As Long
    dmPelsWidth             As Long
    dmPelsHeight            As Long
    dmDisplayFlags          As Long
    dmDisplayFrequency      As Long
End Type

Private Type Rect
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type BitmapStruc
    lngDc                   As Long
    lngBitmap               As Long
    lngPointer              As Long
    Area                    As Rect
End Type

Private Type PointAPI
    X                       As Long
    Y                       As Long
End Type

Private Type LogPen
    lopnStyle               As Long
    lopnWidth               As PointAPI
    lopnColor               As Long
End Type

Private Type LogBrush
    lbStyle                 As Long
    lbColor                 As Long
    lbHatch                 As Long
End Type

Private Type FontStruc
    Name                    As String
    Alignment               As AlignText
    Bold                    As Boolean
    Italic                  As Boolean
    Underline               As Boolean
    StrikeThru              As Boolean
    PointSize               As Byte
    Colour                  As Long
End Type

Private Type LogFont
    'for the DrawText api call
    lfHeight                As Long
    lfWidth                 As Long
    lfEscapement            As Long
    lfOrientation           As Long
    lfWeight                As Long
    lfItalic                As Byte
    lfUnderline             As Byte
    lfStrikeOut             As Byte
    lfCharSet               As Byte
    lfOutPrecision          As Byte
    lfClipPrecision         As Byte
    lfQuality               As Byte
    lfPitchAndFamily        As Byte
    lfFaceName(1 To 32)     As Byte
End Type

Private Type Point
    'you'll need this to reference a point on the
    'screen'
    X                       As Integer
    Y                       As Integer
End Type

'To hold the RGB value
Private Type RGBVal
    Red                     As Integer
    Green                   As Integer
    Blue                    As Integer
End Type

'bitmap structure for the GetObject api call
Private Type BITMAP '24 bytes
    bmType                  As Long
    bmWidth                 As Long
    bmHeight                As Long
    bmWidthBytes            As Long
    bmPlanes                As Integer
    bmBitsPixel             As Integer
    bmBits                  As Long
End Type

'holds version information about the operating system
Private Type OSVERSIONINFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

'holds palette information
Private Type PALETTEENTRY
    peRed                   As Byte
    peGreen                 As Byte
    peBlue                  As Byte
    peFlags                 As Byte
End Type

'holds entented palette information
Private Type LOGPALETTE
    palVersion              As Integer
    palNumEntries           As Integer
    palPalEntry(255)        As PALETTEENTRY ' Enough for 256 colors
End Type

'holds descriptor information for a GUI
Private Type GUID
    Data1                   As Long
    Data2                   As Integer
    Data3                   As Integer
    Data4(7)                As Byte
End Type

'holds picture information - used for a Picture property
Private Type PicBmp
    Size                    As Long
    Type                    As Long
    hBmp                    As Long
    hPal                    As Long
    Reserved                As Long
End Type

'used mainly by ApplyMatrix for holding colour
'data while processing
Private Type ColourInfoType
    lngColour               As Long
    intRGB(2)               As Integer
End Type

'--------------------------------------------------------------------------
'                         MODULE LEVEL CONSTANTS
'--------------------------------------------------------------------------
'general constants
Private Const GW_CHILD              As Integer = 5
Private Const GW_HWNDFIRST          As Integer = 0
Private Const GW_HWNDLAST           As Integer = 1
Private Const GW_HWNDNEXT           As Integer = 2
Private Const GW_HWNDPREV           As Integer = 3
Private Const GWL_WNDPROC           As Integer = (-4)
Private Const IDANI_OPEN            As Long = &H1
Private Const IDANI_CLOSE           As Long = &H2
Private Const IDANI_CAPTION         As Long = &H3
Private Const WM_USER               As Long = &H400
Private Const RC_PALETTE            As Long = &H100
Private Const SIZEPALETTE           As Long = 104
Private Const RASTERCAPS            As Long = 38

'clipping constants
Private Const RGN_AND               As Long = 1
Private Const RGN_OR                As Long = 2
Private Const RGN_XOR               As Long = 3
Private Const RGN_DIFF              As Long = 4
Private Const RGN_COPY              As Long = 5
Private Const ERRORREGION           As Integer = 0
Private Const NULLREGION            As Integer = 1
Private Const SIMPLEREGION          As Integer = 2
Private Const COMPLEXREGION         As Integer = 3

'colour adjustment constants
Private Const CA_LOG_FILTER         As Long = &H2
Private Const CA_NEGATIVE           As Long = &H1
Private Const ILLUMINANT_A          As Long = 1
Private Const ILLUMINANT_B          As Long = 2
Private Const ILLUMINANT_C          As Long = 3
Private Const ILLUMINANT_D50        As Long = 4
Private Const ILLUMINANT_D55        As Long = 5
Private Const ILLUMINANT_D65        As Long = 6
Private Const ILLUMINANT_D75        As Long = 7
Private Const ILLUMINANT_F2         As Long = 8

'system parameter info constants
Private Const SPI_GETWORKAREA       As Long = 48

'alphablend constants
Private Const AC_SRC_OVER           As Long = &H0
Private Const AC_SRC_ALPHA          As Long = &H0

'Image load constants
Private Const LR_LOADFROMFILE       As Long = &H10
Private Const LR_CREATEDIBSECTION   As Long = &H2000
Private Const LR_DEFAULTSIZE        As Long = &H40

'PatBlt constants
Private Const PATCOPY               As Long = &HF00021  ' (DWORD) dest = pattern
Private Const PATINVERT             As Long = &H5A0049  ' (DWORD) dest = pattern XOR dest
Private Const PATPAINT              As Long = &HFB0A09  ' (DWORD) dest = DPSnoo
Private Const DSTINVERT             As Long = &H550009  ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS             As Long = &H42      ' (DWORD) dest = BLACK
Private Const WHITENESS             As Long = &HFF0062  ' (DWORD) dest = WHITE

'Display constants
Private Const CDS_FULLSCREEN        As Integer = 4
Private Const DM_BITSPERPEL         As Long = &H40000
Private Const DM_PELSWIDTH          As Long = &H80000
Private Const DM_PELSHEIGHT         As Long = &H100000
Private Const DM_DISPLAYFLAGS       As Long = &H200000
Private Const DM_DISPLAYFREQUENCY   As Long = &H400000

'DrawText constants
Private Const DT_CENTER             As Long = &H1
Private Const DT_BOTTOM             As Long = &H8
Private Const DT_CALCRECT           As Long = &H400
Private Const DT_EXPANDTABS         As Long = &H40
Private Const DT_EXTERNALLEADING    As Long = &H200
Private Const DT_LEFT               As Long = &H0
Private Const DT_NOCLIP             As Long = &H100
Private Const DT_NOPREFIX           As Long = &H800
Private Const DT_RIGHT              As Long = &H2
Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_TABSTOP            As Long = &H80
Private Const DT_TOP                As Long = &H0
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORDBREAK          As Long = &H10
Private Const TRANSPARENT           As Integer = 1
Private Const OPAQUE                As Integer = 2

'CreateBrushIndirect constants
Private Const BS_DIBPATTERN         As Integer = 5
Private Const BS_DIBPATTERN8X8      As Integer = 8
Private Const BS_DIBPATTERNPT       As Integer = 6
Private Const BS_HATCHED            As Integer = 2
Private Const BS_HOLLOW             As Integer = 1
Private Const BS_NULL               As Integer = 1
Private Const BS_PATTERN            As Integer = 3
Private Const BS_PATTERN8X8         As Integer = 7
Private Const BS_SOLID              As Integer = 0
Private Const HS_BDIAGONAL          As Integer = 3  '  /////
Private Const HS_CROSS              As Integer = 4  '  +++++
Private Const HS_DIAGCROSS          As Integer = 5  '  xxxxx
Private Const HS_FDIAGONAL          As Integer = 2  '  \\\\\
Private Const HS_HORIZONTAL         As Integer = 0  '  -----
Private Const HS_NOSHADE            As Integer = 17
Private Const HS_SOLID              As Integer = 8
Private Const HS_SOLIDBKCLR         As Integer = 23
Private Const HS_SOLIDCLR           As Integer = 19
Private Const HS_VERTICAL           As Integer = 1  '  |||||

'LogFont constants
Private Const LF_FACESIZE           As Integer = 32
Private Const FW_BOLD               As Integer = 700
Private Const FW_DONTCARE           As Integer = 0
Private Const FW_EXTRABOLD          As Integer = 800
Private Const FW_EXTRALIGHT         As Integer = 200
Private Const FW_HEAVY              As Integer = 900
Private Const FW_LIGHT              As Integer = 300
Private Const FW_MEDIUM             As Integer = 500
Private Const FW_NORMAL             As Integer = 400
Private Const FW_SEMIBOLD           As Integer = 600
Private Const FW_THIN               As Integer = 100
Private Const DEFAULT_CHARSET       As Integer = 1
Private Const OUT_CHARACTER_PRECIS  As Integer = 2
Private Const OUT_DEFAULT_PRECIS    As Integer = 0
Private Const OUT_DEVICE_PRECIS     As Integer = 5
Private Const OUT_OUTLINE_PRECIS    As Integer = 8
Private Const OUT_RASTER_PRECIS     As Integer = 6
Private Const OUT_STRING_PRECIS     As Integer = 1
Private Const OUT_STROKE_PRECIS     As Integer = 3
Private Const OUT_TT_ONLY_PRECIS    As Integer = 7
Private Const OUT_TT_PRECIS         As Integer = 4
Private Const CLIP_CHARACTER_PRECIS As Integer = 1
Private Const CLIP_DEFAULT_PRECIS   As Integer = 0
Private Const CLIP_EMBEDDED         As Integer = 128
Private Const CLIP_LH_ANGLES        As Integer = 16
Private Const CLIP_MASK             As Long = &HF
Private Const CLIP_STROKE_PRECIS    As Integer = 2
Private Const CLIP_TT_ALWAYS        As Integer = 32
Private Const WM_SETFONT            As Long = &H30
Private Const LF_FULLFACESIZE       As Integer = 64
Private Const DEFAULT_PITCH         As Integer = 0
Private Const DEFAULT_QUALITY       As Integer = 0
Private Const PROOF_QUALITY         As Integer = 2

'GetDeviceCaps constants
Private Const LOGPIXELSY            As Integer = 90       '  Logical pixels/inch in Y
Private Const LOGPIXELSX            As Integer = 88       '  Logical pixels/inch in X

'colourspace constants
Private Const MAX_PATH              As Integer = 260

'pen constants
Private Const PS_COSMETIC           As Long = &H0
Private Const PS_DASH               As Integer = 1      '  -------
Private Const PS_DASHDOT            As Integer = 3      '  _._._._
Private Const PS_DASHDOTDOT         As Integer = 4      '  _.._.._
Private Const PS_DOT                As Integer = 2      '  .......
Private Const PS_ENDCAP_ROUND       As Long = &H0
Private Const PS_ENDCAP_SQUARE      As Long = &H100
Private Const PS_ENDCAP_FLAT        As Long = &H200
Private Const PS_GEOMETRIC          As Long = &H10000
Private Const PS_INSIDEFRAME        As Long = 6
Private Const PS_JOIN_BEVEL         As Long = &H1000
Private Const PS_JOIN_MITER         As Long = &H2000
Private Const PS_JOIN_ROUND         As Long = &H0
Private Const PS_SOLID              As Long = 0

'mouse cursor constants
Private Const IDC_APPSTARTING       As Long = 32650&
Private Const IDC_ARROW             As Long = 32512&
Private Const IDC_CROSS             As Long = 32515&
Private Const IDC_IBEAM             As Long = 32513&
Private Const IDC_ICON              As Long = 32641&
Private Const IDC_NO                As Long = 32648&
Private Const IDC_SIZE              As Long = 32640&
Private Const IDC_SIZEALL           As Long = 32646&
Private Const IDC_SIZENESW          As Long = 32643&
Private Const IDC_SIZENS            As Long = 32645&
Private Const IDC_SIZENWSE          As Long = 32642&
Private Const IDC_SIZEWE            As Long = 32644&
Private Const IDC_UPARROW           As Long = 32516&
Private Const IDC_WAIT              As Long = 32514&

'menu constants
Private Const MFT_RADIOCHECK        As Long = &H200&
Private Const MF_BITMAP             As Long = &H4&
Private Const MIIM_TYPE             As Long = &H10
Private Const MIIM_SUBMENU          As Long = &H4

'some key values for GetASyncKeyState
Private Const KLeft                 As Integer = 37
Private Const KUp                   As Integer = 38
Private Const KRight                As Integer = 39
Private Const KDown                 As Integer = 40
Private Const K_KEY_DOWN            As Integer = -32767
Private Const K_PRESSED             As Integer = -32768

'extra system colour constants to be extended from
'the vb collection SystemColorConstants (Win98+)
Private Const vbActiveTitleBar2     As Double = 2147483675#
Private Const vbInactiveTitleBar2   As Double = 2147483676#

'some mathimatical constants
Private Const PI                    As Single = 3.14159265358979
Private Const PIdiv180              As Single = 1.74532925199433E-02    'used to speed calculations when converting from/to degrees/radians for the functions Sin/Cos/Tan etc.

'some class constants
Private Const NO_COLOUR             As Long = -2147483648# '2 ^ 32 --> 4 bytes

'--------------------------------------------------------------------------
'                           MODULE LEVEL VARIABLES
'--------------------------------------------------------------------------

'some general variables used by some api calls
Private mudtRectFrom        As Rect
Private mudtRectTo          As Rect
Private mlnghTray           As Long
Private mlnghStartMenu      As Long
Private mlnghChild          As Long
Private mstrClass           As String * 255
Private mlngClassNameLen    As Long
Private mlngRetVal          As Long
Private mblnResChanged      As Boolean

'--------------------------------------------------------------------------
'                             CLASS VARIABLES
'--------------------------------------------------------------------------
Private mudtBitmap          As BitmapStruc          'the bitmap that this class encapsulates
Private mlngBackColour      As Long                 'the background colour of the bitmap
Private menmType            As PictureTypeConstants 'the type of picture the bitmap is (Usually a bitmap, but can be Icon or metafile if loaded)
Private mlngCompatableDc    As Long                 'holds a device context that the bitmap is compatable with (default: the screen window)

'--------------------------------------------------------------------------
'                              CLASS METHODS
'--------------------------------------------------------------------------
Public Sub AdjustColour(Optional ByVal enmNegative As EnumColourNegative = neg_Normal, _
                        Optional ByVal enmIlluminant As EnumLightColour = lit_DayLightStandard, _
                        Optional ByVal lngRedGamma As Long = 10000, _
                        Optional ByVal lngGreenGamma As Long = 10000, _
                        Optional ByVal lngBlueGamma As Long = 10000, _
                        Optional ByVal intRoundToBlackAt As Integer = 0, _
                        Optional ByVal intRoundToWhiteFrom As Integer = 10000, _
                        Optional ByVal intContrast As Integer = 0, _
                        Optional ByVal intBrightness As Integer = 0, _
                        Optional ByVal intColourfulness As Integer = 0, _
                        Optional ByVal intRedGreenTint As Integer = 0)
    'This will set the colour adjustmust to the bitmap
    
    Dim lngResult       As Long             'holds any returned error value from an api call
    Dim udtColours      As COLORADJUSTMENT  'holds the colour adjustment values for the bitmap
    
    'make sure that a bitmap has been created
    If mudtBitmap.lngDc = 0 Then
        Exit Sub
    End If
    
    'validate the parameters to make sure that they are within the
    'correct ranges
    If lngRedGamma <> 10000 Then
        Select Case lngRedGamma
        Case Is < 25000, Is > 65000
            'set to no gamma correction
            lngRedGamma = 10000
        End Select
    End If
    If lngGreenGamma <> 10000 Then
        Select Case lngGreenGamma
        Case Is < 25000, Is > 65000
            'set to no gamma correction
            lngGreenGamma = 10000
        End Select
    End If
    If lngBlueGamma <> 10000 Then
        Select Case lngBlueGamma
        Case Is < 25000, Is > 65000
            'set to no gamma correction
            lngBlueGamma = 10000
        End Select
    End If

    If intRoundToBlackAt <> 0 Then
        Select Case intRoundToBlackAt
        Case Is < 0, Is > 4000
            'set to no rounding
            intRoundToBlackAt = 0
        End Select
    End If
    If intRoundToWhiteFrom <> 10000 Then
        Select Case intRoundToWhiteFrom
        Case Is < 6000, Is > 10000
            'set to no rounding
            intRoundToWhiteFrom = 10000
        End Select
    End If
    
    Select Case intContrast
    Case Is < -100, Is > 100
        'default to zero
        intContrast = 0
    End Select
    Select Case intBrightness
    Case Is < -100, Is > 100
        'default to zero
        intBrightness = 0
    End Select
    Select Case intColourfulness
    Case Is < -100, Is > 100
        'default to zero
        intColourfulness = 0
    End Select
    Select Case intRedGreenTint
    Case Is < -100, Is > 100
        'default to zero
        intRedGreenTint = 0
    End Select
    
    'setup the datatype before setting the colour values
    With udtColours
        .caSize = Len(udtColours)
        .caFlags = enmNegative
        .caIlluminantIndex = enmIlluminant
        .caRedGamma = lngRedGamma
        .caGreenGamma = lngGreenGamma
        .caBlueGamma = lngBlueGamma
        .caReferenceBlack = intRoundToBlackAt
        .caReferenceWhite = intRoundToWhiteFrom
        .caContrast = intContrast
        .caBrightness = intBrightness
        .caColorfulness = intColourfulness
        .caRedGreenTint = intRedGreenTint
    End With
    
    'set the colour adjustment
    lngResult = SetColorAdjustment(mudtBitmap.lngDc, udtColours)
End Sub

Public Sub AlphaBlend(ByVal lngPic1hDc As Long, _
                    ByVal lngPic2hDc As Long, _
                    Optional ByVal intDestX As Integer = 0, _
                    Optional ByVal intDestY As Integer = 0, _
                    Optional ByVal intWidth As Integer = -1, _
                    Optional ByVal intHeight As Integer = -1, _
                    Optional ByVal lngPic1X As Integer = 0, _
                    Optional ByVal lngPic1Y As Integer = 0, _
                    Optional ByVal lngPic2X As Integer = 0, _
                    Optional ByVal lngPic2Y As Integer = 0, _
                    Optional ByVal sngBlendAmount As Single = 0.5, _
                    Optional ByVal enmMeasurement As Scaling = InPixels)

    'This uses the windows GDI to blend two bitmaps
    'together. If you need to use a blend mask, then
    'please use the MaskBlend procedure. This function
    'is only supported by w98+ and w2000+ using the
    'Msimg32.dll library. This can be downloaded from
    'the Microsoft web site
    
    Dim udtTempBmp    As BitmapStruc    'this temperorily holds the blended pictures before copying to the destination bitmap
    Dim udtBlendInfo  As BLENDFUNCTION  'this sets the blend information for the api call
    Dim lngBlendStruc As Long           'this will hold the converted BLENDFUNCTION structure
    Dim lngResult     As Long           'this holds any error value returned from the api call
    Dim intPxlHeight  As Integer        'the height in twips of a pixel
    Dim intPxlWidth   As Integer        'the width in twips of a pixel
    
    'set the default heights and widths if necessary
    If intWidth < 0 Then
        intWidth = mudtBitmap.Area.Left + _
                   mudtBitmap.Area.Right
    End If
    If intHeight < 0 Then
        intHeight = mudtBitmap.Area.Top + _
                    mudtBitmap.Area.Bottom
    End If
    
    'convert values to pixels if necessary
    If enmMeasurement = InTwips Then
        'get the pixel intHeight and intWidth
        'values per twips in the current
        'screen resolution.
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'start converting the twip values to pixels
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        lngPic1X = lngPic1X / intPxlWidth
        lngPic2X = lngPic2X / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        lngPic1Y = lngPic1Y / intPxlHeight
        lngPic2Y = lngPic2Y / intPxlHeight
    End If
    
    'set the blend information
    With udtBlendInfo
        .bytBlendOp = AC_SRC_OVER
        .bytSourceConstantAlpha = CByte(255 * sngBlendAmount)
    End With
    
    'convert the type to a long
    Call RtlMoveMemory(lngBlendStruc, _
                       udtBlendInfo, _
                       4)
    
    With udtTempBmp
        'set the bitmap dimensions
        With .Area
            .Bottom = intHeight
            .Right = intWidth
        End With
        
        'create the new bitmap
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area, _
                             mudtBitmap.lngDc)
        
        'copy the first picture
        lngResult = BitBlt(.lngDc, _
                           0, _
                           0, _
                           intWidth, _
                           intHeight, _
                           lngPic1hDc, _
                           lngPic1X, _
                           lngPic1Y, _
                           P_COPY)
        
        'blend the second picture with
        'the first picture
        lngResult = APIAlphaBlend(.lngDc, _
                                  0, _
                                  0, _
                                  intWidth, _
                                  intHeight, _
                                  lngPic2hDc, _
                                  lngPic2X, _
                                  lngPic2Y, _
                                  intWidth, _
                                  intHeight, _
                                  lngBlendStruc)
        
        'copy the picture to the destination
        lngResult = BitBlt(mudtBitmap.lngDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           .lngDc, _
                           0, _
                           0, _
                           P_COPY)
        
        'remove the temperory bitmap from memory
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
    End With
End Sub

Public Sub ApplyMatrix(ByRef sngMatrix() As Single, _
                       Optional ByVal intLeft As Integer = -32768, _
                       Optional ByVal intTop As Integer = -32768, _
                       Optional ByVal intWidth As Integer = 0, _
                       Optional ByVal intHeight As Integer = 0)
    'this will apply a X.X matrix that holds mixing values for the colour in
    'the centre of the matrix (2,2 for a zero based array). This is used by
    'several functions including Blur and Sharpen. ALL matrix value must be
    '-255 and 255 or any decimal between them (typically between 0 and 1).
    'The size of the first two dimensions of the array MUST be equal in size.
    'If there are other dimensions of the array the procedure ignores them.
    'The procedure assumes that there are at least two dimensions in the array.
    
    Const Red           As Integer = 0      'used to index Red in the arrays
    Const Green         As Integer = 1      'used to index Green in the arrays
    Const Blue          As Integer = 2      'used to index Blue in the arrays
    
    Dim intX            As Integer          'used to cycle through the rows of the bitmap
    Dim intY            As Integer          'used to cycle through the columns of the bitmap
    Dim intCol()        As Integer          'holds the red, green and blue value of one pixel
    Dim lngColour       As Long             'holds the long colour value of Col1
    Dim intSrc()        As Integer          'holds the red, green and blue value of the source pixel
    Dim lngSource       As Long             'holds the long colour value of the source pixel
    Dim intMix()        As Integer          'holds the blended red, green and blue of a pixel to be set
    Dim lngMix          As Long             'holds the long colour value of the mixed colour
    Dim udtTemp         As BitmapStruc      'holds the details of the temperory bitmap
    Dim lngResult       As Long             'used to hold any returned error value from an api call
    Dim intMatrixX      As Integer          'used to cycle through the matrix rows
    Dim intMatrixY      As Integer          'used to cycle through the matrix columns
    Dim intColI         As Integer          'used to cycle through the array colours during calculations
    Dim intCentreX      As Integer          'holds the centre index of the array
    Dim intCentreY      As Integer          'holds the centre index of the array
    Dim intTestX        As Integer          'used to test the bounds of the matrix within the applied area
    Dim intTestY        As Integer          'used to test the bounds of the matrix within the applied area
    Dim intUpX          As Integer          'holds the upper bound of the X dimension of the array
    Dim intLowX         As Integer          'holds the lower bound of the X dimension of the array
    Dim intUpY          As Integer          'holds the upper bound of the Y dimension of the array
    Dim intLowY         As Integer          'holds the lower bound of the Y dimension of the array
    Dim intSizeX        As Integer          'holds the size of the X dimension of the array
    Dim intSizeY        As Integer          'holds the size of the Y dimension of the array
    Dim intSizeTot      As Integer          'holds the size of the array in total
    Dim lngColTot()     As Long             'holds the "total" colour value to be averaged
    Dim udtColours()    As ColourInfoType   'holds the colour information for the entire applied area of the bitmap
    Dim intApplyWidth   As Integer          'holds the width of the colour matrix
    Dim intApplyHeight  As Integer          'holds the height of the colour matrix
    Dim sngMatrixTot    As Single           'holds the total of all the matrix mix values
    
    'validate the parameters
    With mudtBitmap.Area
        If (intLeft = -32768) Then
            'default to the bitmaps left
            intLeft = .Left
        End If
        If (intTop = -32768) Then
            'default to the bitmaps top
            intTop = .Top
        End If
        If (intWidth <= 0) Then
            'default to the bitamps width
            intWidth = .Right - .Left
        End If
        If (intHeight <= 0) Then
            'default to the bitmaps height
            intHeight = .Bottom - .Top
        End If
    End With    'mudtBitmap.Area
    
    'get the array dimensions and size
    intUpX = UBound(sngMatrix, 1)
    intLowX = LBound(sngMatrix, 1)
    intUpY = UBound(sngMatrix, 2)
    intLowY = LBound(sngMatrix, 2)
    intSizeX = (intUpX - intLowX) + 1
    intSizeY = (intUpY - intLowY) + 1
    intSizeTot = intSizeX * intSizeY
    
    'make sure that the array bounds match
    If (intSizeX <> intSizeY) Then
        'the array bounds are not equal
        Exit Sub
    End If
    
    'get the centre of the array
    intCentreX = Math.ROUND(((intUpX - intLowX) / 2), 0)
    intCentreY = Math.ROUND(((intUpY - intLowY) / 2), 0)
    
    'make sure that the matrix only contains values from -1 to 1
    For intMatrixX = intLowX To intUpX
        For intMatrixY = intLowY To intUpY
            If (sngMatrix(intMatrixX, intMatrixY) < -255) Then
                'adjust to minimum value
                sngMatrix(intMatrixX, intMatrixY) = -255
            ElseIf (sngMatrix(intMatrixX, intMatrixY) > 255) Then
                'adjust to maxmimum value
                sngMatrix(intMatrixX, intMatrixY) = 255
            End If
            sngMatrixTot = sngMatrixTot + sngMatrix(intMatrixX, intMatrixY)
        Next intMatrixY
    Next intMatrixX
    sngMatrixTot = sngMatrixTot / intSizeTot
    
    'get the complete colour information for the applied area of the
    'source bitmap
    intApplyWidth = (intWidth - 1)
    intApplyHeight = (intHeight - 1)
    ReDim udtColours(intApplyWidth, intApplyHeight)
    For intX = 0 To intApplyWidth
        For intY = 0 To intApplyHeight
            With udtColours(intX, intY)
                .lngColour = GetPixel(mudtBitmap.lngDc, _
                                      intLeft + intX, _
                                      intTop + intY)
                Call GetRGB(.lngColour, _
                            .intRGB(Red), _
                            .intRGB(Green), _
                            .intRGB(Blue))
            End With    'udtColours(intX, intY)
        Next intY
    Next intX
    
    'make sure that the colour arrays are the correct size
    ReDim intCol(2)
    ReDim intSrc(2)
    ReDim intMix(2)
    ReDim lngColTot(2)
    
    With udtTemp
        'create a temperory bitmap matching the size of the original bitmap
        .Area = mudtBitmap.Area
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
    
    'apply the matrix to the bitmap
    For intX = 0 To intApplyWidth
        For intY = 0 To intApplyHeight
            'reset the "total" colour for this pixel
            For intColI = Red To Blue
                lngColTot(intColI) = 0
            Next intColI
                
            'get the colour of the source pixel
            With udtColours(intX, intY)
                intSrc = .intRGB
            End With    'udtColours(intX, intY)
            
            'adjust the pixel according to its matrix value
            For intMatrixX = intLowX To intUpX
                For intMatrixY = intLowY To intUpY
                    
                    'get the position of the pixel to test and see if it
                    'is within the applied area of the bitmap
                    intTestX = intX + (intMatrixX - intCentreX)
                    intTestY = intY + (intMatrixY - intCentreY)
                    
                    'should we get the colour value for this pixel and mix
                    If ((intTestX >= 0) And _
                        (intTestX <= intApplyWidth)) And _
                       ((intTestY >= 0) And _
                        (intTestY <= intApplyHeight)) Then
                        'get the colour value, to the degree specified and
                        'add it to the total colour to be averaged later
                        With udtColours(intTestX, intTestY)
                            intCol() = .intRGB()
                        End With    'udtColours(intTestX, intTestY)
                        
                        'adjust the colour according to the matrix value
                        'and add it to the current total
                        For intColI = Red To Blue
'* ---- Working Code
                            intMix(intColI) = intCol(intColI)
                            intMix(intColI) = intCol(intColI) + _
                                              ((intSrc(intColI) - _
                                                intCol(intColI)) * _
                                               sngMatrix(intMatrixX, _
                                                         intMatrixY))
                            lngColTot(intColI) = (lngColTot(intColI) + _
                                                     intMix(intColI))
                        Next intColI
                    
                    Else
                        'ignore this pixel from the colour blending by
                        'putting the source pixel colour value into the
                        'total
                        For intColI = Red To Blue
                            lngColTot(intColI) = (lngColTot(intColI) + _
                                                     intSrc(intColI))
                        Next intColI
                    End If  'if the matrix value is within the apply area
                
                Next intMatrixY 'matrix index
            Next intMatrixX
            
            'get the "average" colour value adjusting the source pixel
            'in relation to the surrounding colours
            For intColI = Red To Blue
                lngColTot(intColI) = lngColTot(intColI) - intSrc(intColI)
                lngColTot(intColI) = lngColTot(intColI) + _
                                     (intSrc(intColI) * _
                                      sngMatrix(intCentreX, _
                                                intCentreY))
                lngColTot(intColI) = lngColTot(intColI) / intSizeTot
                'lngColTot(intColI) = (lngColTot(intColI) / intSizeTot) * sngMatrixTot
            Next intColI
            
            'set the new colour of the pixel to the destination bitmap
            Call SetPixel(udtTemp.lngDc, _
                          intLeft + intX, _
                          intTop + intY, _
                          RGB(Abs(lngColTot(Red)), _
                              Abs(lngColTot(Green)), _
                              Abs(lngColTot(Blue))))
        Next intY   'bitmap index
    Next intX
    
    With udtTemp
        'copy the bitmap from the temperory bitmap to the source
        lngResult = BitBlt(mudtBitmap.lngDc, _
                           intLeft, _
                           intTop, _
                           intWidth, _
                           intHeight, _
                           .lngDc, _
                           intLeft, _
                           intTop, _
                           vbSrcCopy)
        
        'remove the temperory bitmap from memory
        Call DeleteBitmap(.lngDc, .lngBitmap, .lngPointer)
    End With    'udtTemp
End Sub

Public Sub Blur(Optional ByVal sngAmount As Single = 0.5, _
                Optional ByVal intBlurRadius As Integer = 2, _
                Optional ByVal intLeft As Integer = -32768, _
                Optional ByVal intTop As Integer = -32768, _
                Optional ByVal intWidth As Integer = 0, _
                Optional ByVal intHeight As Integer = 0)
    'This will blur the selected area of the bitmap by the specified amount.
    'The blur area is based on a circular algorithm, in that the blur amount
    'decreases towards the centre of the array. The BlurRadius is measured in
    'pixels from the centre. The matrix will represent the pixels in the
    'surrounding area from the centre pixel.
    
    Dim sngMatrix()     As Single       'holds the matrix that we are going to use to blur the picture
    Dim intX            As Integer      'used to cycle through the array
    Dim intY            As Integer      'used to cycle through the array
    Dim intCentre       As Integer      'holds the array index of the centre
    Dim intSize         As Integer      'holds the size of the array
    Dim sngDist         As Single       'holds the distance of the element from the centre of the array
    
    If (intBlurRadius < 0) Then
        'we cannot create a negative sized array
        Exit Sub
    End If
    sngAmount = sngAmount / 2
    
    'resize the array for the blur area
    intSize = (intBlurRadius * 2)
    ReDim sngMatrix(intSize, intSize)
    
    'setup the "Blur" matrix
    For intX = 0 To intSize
        For intY = 0 To intSize
            'calculate the blur amount based on the distance from the centre
            'of the array. The specified blur amount is what the blur is at
            'right angles to the centre (ie, the blur amount is taken at the
            'radius from the centre to the left, right, top or bottom edges).
            
            'The square of the hypothineuse is equal to the sum of the squares
            'on the other two sides. The distance from the centre is the
            'Hypothineuse. Any negative value will be converted to a positive.
            sngDist = Sqr(((intX - intBlurRadius) ^ 2) + _
                          ((intY - intBlurRadius) ^ 2))
            sngMatrix(intX, intY) = 1 - (sngAmount * Abs(sngDist))
        Next intY
    Next intX
    
    'apply the matrix to the bitmap
    Call ApplyMatrix(sngMatrix, _
                     intLeft, _
                     intTop, _
                     intWidth, _
                     intHeight)
End Sub

Public Sub Mosaic(Optional ByVal sngAmount As Single = 0.2, _
                  Optional ByVal intLeft As Integer = -32768, _
                  Optional ByVal intTop As Integer = -32768, _
                  Optional ByVal intWidth As Integer = 0, _
                  Optional ByVal intHeight As Integer = 0)
    'This will convert the bitmap to a mosaic image. A value of 1 will level
    'all colours on the bitmap into a single averaged colour, whereas a value
    'of 0 will have no mosaic effect on the picture
    
    Dim sngMatrix()         As Single       'this holds the blur matrix that will level all colours in the selected area
    Dim intTileSize         As Integer      'holds the size of the tiles in pixels
    Dim intX                As Integer      'used for cycling through the rows of the matrix
    Dim intY                As Integer      'used for cycling through the columns of the matrix
    Dim intTileX            As Integer      'used for cycling through the tiled rows
    Dim intTileY            As Integer      'used for cycling through the tiled columns
    
    'validate the parameters
    With mudtBitmap.Area
        If (intLeft = -32768) Then
            'default to the bitmaps left
            intLeft = .Left
        End If
        If (intTop = -32768) Then
            'default to the bitmaps top
            intTop = .Top
        End If
        If (intWidth <= 0) Then
            'default to the bitamps width
            intWidth = .Right - .Left
        End If
        If (intHeight <= 0) Then
            'default to the bitmaps height
            intHeight = .Bottom - .Top
        End If
        If (sngAmount < 0) Then
            sngAmount = 0
        ElseIf (sngAmount > 1) Then
            sngAmount = 1
        End If
    End With    'mudtBitmap.Area
    
    'get the size of the tiles
    intTileSize = (((intHeight + intWidth) \ 2) * sngAmount) '- 1
    
    'create the blur matrix
    ReDim sngMatrix(intTileSize, intTileSize)
    For intX = 0 To intTileSize
        For intY = 0 To intTileSize
            sngMatrix(intX, intY) = 0.5
        Next intY
    Next intX
    
    For intTileX = intLeft To (intLeft + intWidth) Step intTileSize
        For intTileY = intTop To (intTop + intHeight) Step intTileSize
            Call ApplyMatrix(sngMatrix, _
                             intTileX, _
                             intTileY, _
                             intTileSize, _
                             intTileSize)
        Next intTileY
    Next intTileX
End Sub

'See AlphaBlend for 80% speed increase
Public Sub MaskBlend(ByVal lngDesthDc As Long, _
                     ByVal lngSrc1hDc As Long, _
                     ByVal lngSrc2hDc As Long, _
                     Optional ByVal lngMaskhDc As Long = 0, _
                     Optional ByVal intDestX As Integer = 0, _
                     Optional ByVal intDestY As Integer = 0, _
                     Optional ByVal intWidth As Integer = 0, _
                     Optional ByVal intHeight As Integer = 0, _
                     Optional ByVal lngSrc1X As Integer = 0, _
                     Optional ByVal lngSrc1Y As Integer = 0, _
                     Optional ByVal lngSrc2X As Integer = 0, _
                     Optional ByVal lngSrc2Y As Integer = 0, _
                     Optional ByVal sngBlendAmount As Single = 0.5, _
                     Optional ByVal intMeasurement As Scaling = InPixels)

    'This is a "brute force" alpha blend function. Because it's written in
    'vb, this function is not as fast at it might otherwise be in another
    'language like C++ or Fox.
    'The purpose of the function is to mix the colours of two bitmaps to
    'produce a result that looks like both pictures. Think of it as fading
    'one picture into another. I get the pixel colour of a point in picture1,
    'and the colour of the corresponding pixel in pixture2, find the 'middle'
    'colour and put it into the destination bitmap. There are no calls to
    'other procedures or functions other than api calls. This is to improve
    'speed as all calculations are made internally.
    'The parameter sngBlendAmount MUST be between 1 and 0. If not then
    'the value is rounded to 1 or zero and a simple Blit is used.
    'However, sngBlendAmount is ignored if a Mask bitmap has been specified.
    'Please note that if the mask used only contains black or white pixels,
    'then it is recommended that you use the MergeBitmaps procedure as
    'it will process the bitmaps much faster (by about 15 to 30 times).
    'Keep in mind that using a mask bitmap is about 25% slower than
    'specifying the blend amount.

    Dim TempBmp         As BitmapStruc  'a temperory bitmap
    Dim lngResult       As Long         'any result returned from an api call
    Dim Col1            As RGBVal       'used to store a pixel colour in RGB format
    Dim Col2            As RGBVal       'used to store a pixel colour in RGB format
    Dim BlendCol        As RGBVal       'used to store a pixel colour in RGB format
    Dim MaskCol         As RGBVal       'used to store a pixel colour in RGB format
    Dim lngCounterX     As Long         'scan the rows of the bitmap
    Dim lngCounterY     As Long         'scan the columns of the bitmap
    Dim intPxlHeight    As Integer      'the pixel height in twips
    Dim intPxlWidth     As Integer      'the pixel width in twips
    Dim lngBlendCol     As Long         'the blended colour calculated from the two pixel colours of the bitmaps
    Dim lngCol1         As Long         'used to store a pixel colour in Long format
    Dim lngCol2         As Long         'used to store a pixel colour in Long format
    Dim lngMaskCol      As Long         'used to store a pixel colour in Long format
    
    'first convert the passed values if they
    'need to be converted.
    If intMeasurement = InTwips Then
        'get the pixel intHeight and intWidth
        'values per twips in the current
        'screen resolution.
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'start converting the twip values to pixels
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        lngSrc1X = lngSrc1X / intPxlWidth
        lngSrc2X = lngSrc2X / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        lngSrc1Y = lngSrc1Y / intPxlHeight
        lngSrc2Y = lngSrc2Y / intPxlHeight
    End If
    
    'check the defaults
    If (intWidth <= 0) Then
        intWidth = Me.Width
    End If
    If (intHeight <= 0) Then
        intHeight = Me.Height
    End If
    
    'validate the sngBlendAmount parameter.
    'It must be a values between 0 and
    '1. If the parameter is outside these
    'bounds, then round to nearist
    'bounding value (0 or 1)
    Select Case sngBlendAmount
    Case Is >= 1
        sngBlendAmount = 1
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngSrc2hDc, _
                           lngSrc2X, _
                           lngSrc2Y, _
                           P_COPY)
        Exit Sub
    Case Is <= 0
        sngBlendAmount = 0
        
        'just copy the picture instead
        'of trying to blend it
        lngResult = BitBlt(lngDesthDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           lngSrc1hDc, _
                           lngSrc1X, _
                           lngSrc1Y, _
                           P_COPY)
        Exit Sub
    End Select
    
    'create a temperory destination
    'bitmap
    With TempBmp
        .Area.Right = intWidth
        .Area.Bottom = intHeight
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area, lngDesthDc)
    End With
    
    'start going through the 2 source
    'bitmaps and blending the colours.
    For lngCounterX = 0 To intWidth
        For lngCounterY = 0 To intHeight
            'get the pixel colours
            lngCol1 = GetPixel(lngSrc1hDc, _
                               lngSrc1X + lngCounterX, _
                               lngSrc1Y + lngCounterY)
            lngCol2 = GetPixel(lngSrc2hDc, _
                               lngSrc2X + lngCounterX, _
                               lngSrc2Y + lngCounterY)
            
            'if a blend mask has been specified,
            'then get the blend amount
            'from the bitmap.
            If lngMaskhDc <> 0 Then
                lngMaskCol = GetPixel(lngMaskhDc, _
                                      lngCounterX, _
                                      lngCounterY)
                
                'convert the long value into
                'the blend amount
                With MaskCol
                    Call GetRGB(lngMaskCol, .Red, .Green, .Blue)
                    'MaskCol.Blue = lngMaskCol \ 65536
                    'MaskCol.Green = ((lngMaskCol - (MaskCol.Blue * 65536)) \ 256)
                    'MaskCol.Red = (lngMaskCol - (MaskCol.Blue * 65536) - _
                                   (MaskCol.Green * 256))
                    
                    'now convert rgb value to
                    'value between 0 and 1
                    '(divide by 3 for the average
                    'rgb and 255 to a value between
                    '1 and 0 (3 * 255 = 765) )
                    sngBlendAmount = (.Red + .Green + .Blue) \ 765
                End With
            End If
            
            'convert long values to rgb values
            With Col2
                Call GetRGB(lngCol2, .Red, .Green, .Blue)
                '.Blue = lngCol2 \ 65536
                '.Green = ((lngCol2 - (Col2.Blue * 65536)) \ 256)
                '.Red = (lngCol2 - (Col2.Blue * 65536) - (Col2.Green * 256))
            End With
            With Col1
                Call GetRGB(lngCol1, .Red, .Green, .Blue)
                '.Blue = lngCol1 \ 65536
                '.Green = ((lngCol1 - (Col1.Blue * 65536)) \ 256)
                '.Red = (lngCol1 - (Col1.Blue * 65536) - (Col1.Green * 256))
            
                'average the colours by blend amount
                If (.Red <> Col2.Red) _
                   Or (.Green <> Col2.Green) _
                   Or (.Blue <> Col2.Blue) Then
                    
                    BlendCol.Red = .Red - ((.Red - Col2.Red) * sngBlendAmount)
                    BlendCol.Green = .Green - ((.Green - Col2.Green) * sngBlendAmount)
                    BlendCol.Blue = .Blue - ((.Blue - Col2.Blue) * sngBlendAmount)
                Else
                    'there is no point in blending
                    'colours that are the same
                    BlendCol = Col1
                End If
            End With
            
            'convert the BlendCol RGB values
            'to a long
            With BlendCol
                lngBlendCol = RGB(.Red, .Green, .Blue)
            End With
            'lngBlendCol = (CLng(BlendCol.Blue) * 65536) + _
                          (CLng(BlendCol.Green) * 256) + _
                          BlendCol.Red
            
            'set the corresponding pixel colour
            'on the temperory bitmap
            lngResult = SetPixel(TempBmp.lngDc, _
                                 lngCounterX, _
                                 lngCounterY, _
                                 lngBlendCol)
        Next lngCounterY
    Next lngCounterX
    
    'copy the blended picture to the
    'destination bitmap
    'lngDesthDc = Form1.hDc
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       intWidth, _
                       intHeight, _
                       TempBmp.lngDc, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the temperory bitmap
    'from memory
    Call DeleteBitmap(TempBmp.lngDc, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Public Sub FlashWindow(ByVal hWnd As Long, _
                       Optional ByVal intTimes As Integer = 6, _
                       Optional ByVal lngIntervalTicks As Long = 400, _
                       Optional ByVal blnInvert As Boolean = True)
    'This will flash the specified window
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngInvert       As Long     'holds the invert flag
    Dim intCounter      As Integer  'used to flash the window the specified number of times
    
    'are we providing a proper visual cue
    lngInvert = Abs(blnInvert)
    
    'flash the window the specified number of times
    For intCounter = 1 To intTimes
        lngResult = FlashWindowApi(hWnd, lngInvert)
        
        'wait for the specified tick interval
        Call Pause(lngIntervalTicks)
    Next intCounter
End Sub
                       
Public Sub FlipBitmap(Optional ByVal lngSourcehDc As Long, _
                      Optional ByVal intDestX As Integer, _
                      Optional ByVal intDestY As Integer, _
                      Optional ByVal intWidth As Integer, _
                      Optional ByVal intHeight As Integer, _
                      Optional ByVal intSourceX As Integer, _
                      Optional ByVal intSourceY As Integer, _
                      Optional ByVal Orientation As FlipType = FlipHorizontally, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)

    'This procedure will flip a picture either
    'horizontally or vertically. It copies
    'the bitmap either row by row or column
    'by column to improve speed.
    
    Dim intPxlHeight    As Integer      'the height in twips of a pixel
    Dim intPxlWidth     As Integer      'the width in twips of a pixel
    Dim intCounter      As Integer      'used to scan through each row/column
    Dim TempBmp         As BitmapStruc  'holds a temperory bitmap to copy the rows/columns
    Dim intFinish       As Integer      'the row/column to finish at
    Dim lngResult       As Long         'holds any returned error value from an api call
    
    'if not source dc was specified, then use the current bitmap
    If lngSourcehDc = 0 Then
        lngSourcehDc = mudtBitmap.lngDc
    End If
    
    'convert the twips to pixel values if necessary
    If udtMeasurement = InTwips Then
        intPxlWidth = Screen.TwipsPerPixelX
        intPxlHeight = Screen.TwipsPerPixelY
        
        intDestX = intDestX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        intSourceX = intSourceX / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intHeight = intHeight / intPxlHeight
        intSourceY = intSourceY / intPxlHeight
    End If
    
    'create the temperory bitmap
    TempBmp.Area.Right = intWidth
    TempBmp.Area.Bottom = intHeight
    Call CreateNewBitmap(TempBmp.lngDc, _
                         TempBmp.lngBitmap, _
                         TempBmp.lngPointer, _
                         TempBmp.Area, _
                         lngSourcehDc)
    
    'define the bounds of the loop depending on the orientation (do I scan
    'the bitmap row by row or column by column)
    Select Case Orientation
    Case FlipHorizontally
        'scan column by column
        intFinish = intWidth - 1
    Case FlipVertically
        'scan row by row
        intFinish = intHeight - 1
    End Select
    
    For intCounter = 0 To intFinish
        'copy the row or column into the appropiate section of the bitmap
        If Orientation = FlipHorizontally Then
            'horizontal
            lngResult = BitBlt(TempBmp.lngDc, _
                               intFinish - intCounter, _
                               0, _
                               1, _
                               intHeight, _
                               lngSourcehDc, _
                               intSourceX + intCounter, _
                               intSourceY, _
                               P_COPY)
        Else
            'flip vertically
            lngResult = BitBlt(TempBmp.lngDc, _
                               0, _
                               intFinish - intCounter, _
                               intWidth, _
                               1, _
                               lngSourcehDc, _
                               intSourceX, _
                               intSourceY + intCounter, _
                               P_COPY)
        End If
    Next
    
    'copy the flipped bitmap onto the destination bitmap
    lngResult = BitBlt(mudtBitmap.lngDc, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.lngDc, _
                       0, _
                       0, _
                       P_COPY)
                       
    'delete the temperory bitmap
    Call DeleteBitmap(TempBmp.lngDc, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Public Sub RotateBitmap(ByVal lngDesthDc As Long, _
                        Optional ByVal sngAngle As Single = 90, _
                        Optional ByVal intDestX As Integer = 0, _
                        Optional ByVal intDestY As Integer = 0, _
                        Optional ByVal intSourceX As Integer = 0, _
                        Optional ByVal intSourceY As Integer = 0, _
                        Optional ByVal intWidth As Integer = 0, _
                        Optional ByVal intHeight As Integer = 0, _
                        Optional ByVal lngSourcehDc As Long = 0, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)
    
    'This procedure will rotate a bitmap 90, 180 or 270 degrees by row/column BitBlt's and specific angles
    'through slower calculations
    
    Const HALF_PI       As Long = PI / 2    'used to help speed up the calculations
    
    Dim lngResult       As Long             'holds any returned error value from an api call
    Dim intPxlWidth     As Integer          'holds the width in twips of a pixel
    Dim intPxlHeight    As Integer          'holds the height in twips of a pixel
    Dim intCounterX     As Integer          'used to scan through the horizontal pixels
    Dim intCounterY     As Integer          'used to scan through the vertical pixels
    Dim TempBmp         As BitmapStruc      'holds a bitmap to copy the pixels onto
    Dim lngBitCol       As Long             'holds the colour value of a single pixel
    Dim intCentreSrcX   As Integer          ' Center of pic1.
    Dim intCentreSrcY   As Integer          '   "
    Dim intCentreDestX  As Integer          ' Center of pic2.
    Dim intCentreDestY  As Integer          '   "
    Dim sngTempAngle    As Single           ' Angle of c2 to p2.
    Dim intRadius       As Integer          ' Radius from c2 to p2.
    Dim intPointX       As Integer          ' Position on pic1.
    Dim intPointY       As Integer          '   "
    Dim intMaxDim       As Integer          ' Max width or height of pic2.
    Dim lngCol(3)       As Long             'holds the colour value for four points while rotating
    
    'convert twips values to pixels if necessary
    If udtMeasurement = InTwips Then
        intPxlHeight = Screen.TwipsPerPixelY
        intPxlWidth = Screen.TwipsPerPixelX
        
        'convert values
        intDestX = intDestX / intPxlWidth
        intSourceX = intSourceX / intPxlWidth
        intWidth = intWidth / intPxlWidth
        intDestY = intDestY / intPxlHeight
        intSourceY = intSourceY / intPxlHeight
        intHeight = intHeight / intPxlHeight
    End If
    
    'set any default values
    If (lngSourcehDc = 0) Then
        lngSourcehDc = mudtBitmap.lngDc
    End If
    If (intWidth <= 0) Then
        intWidth = mudtBitmap.Area.Right - mudtBitmap.Area.Left
    End If
    If (intHeight <= 0) Then
        intHeight = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
    End If
    
    'round off the angle
    sngAngle = Math.ROUND(sngAngle, 1)
    
    'create a temperory bitmap to draw on
    If (sngAngle = 180) Then
        'the intWidth and intHeight dimensions are the same
        TempBmp.Area.Bottom = intHeight
        TempBmp.Area.Right = intWidth
    
    ElseIf (sngAngle = 90) Or (sngAngle = 270) Then
        'rotate the dimensions 90 degrees
        TempBmp.Area.Bottom = intWidth
        TempBmp.Area.Right = intHeight
        
    Else
        'calcuate the new width/height to maximum boundry to hold the rotated picture
        intRadius = Sqr((CLng(intWidth) ^ 2) + (CLng(intHeight ^ 2)))
        TempBmp.Area.Bottom = intRadius
        TempBmp.Area.Right = intRadius
    End If
    
    'create a temperory bitmap with white background so the picture can be merged onto the original by
    'the programmer using a mask
    Call CreateNewBitmap(TempBmp.lngDc, _
                         TempBmp.lngBitmap, _
                         TempBmp.lngPointer, _
                         TempBmp.Area, _
                         mudtBitmap.lngDc, , _
                         vbWhite)
    
    Select Case sngAngle
    Case 90 To 270
        'rotate bitmap right or left
        For intCounterX = 0 To intWidth
            For intCounterY = 0 To intHeight
                'get the pixel colour
                lngBitCol = GetPixel(lngSourcehDc, _
                                     intSourceX + intCounterX, _
                                     intSourceY + intCounterY)
                
                'copy to appropiate part of the temperory bitmap
                If sngAngle = RotateRight Then
                    'rotate right
                    lngResult = SetPixel(TempBmp.lngDc, _
                                         intHeight - intCounterY, _
                                         intCounterX, _
                                         lngBitCol)
                Else
                    'rotate left
                    lngResult = SetPixel(TempBmp.lngDc, _
                                         intCounterY, _
                                         intHeight - intCounterX, _
                                         lngBitCol)
                End If
            Next intCounterY
        Next intCounterX
    
    Case 180
        'rotate bitmap 180 degrees
        
        'we rotate the bitmap 180 degrees by flipping it vertically and
        'horizontally. This is done fastest by calling the FlipBitmap procedure
        lngResult = BitBlt(TempBmp.lngDc, _
                           0, _
                           0, _
                           intWidth, _
                           intHeight, _
                           lngSourcehDc, _
                           0, _
                           0, _
                           vbSrcCopy)
        Call FlipBitmap(TempBmp.lngDc, Orientation:=FlipHorizontally)
        Call FlipBitmap(TempBmp.lngDc, Orientation:=FlipVertically)
    
    Case Else
        'rotate by the specified angle
       
        ' Compute the centers.
        intCentreSrcX = intSourceX + (intWidth \ 2)
        intCentreSrcY = intSourceY + (intHeight \ 2)
        intCentreDestX = intRadius \ 2
        intCentreDestY = intRadius \ 2
        
        ' Compute the image size.
        intMaxDim = intWidth
        If (intMaxDim < intHeight) Then
            intMaxDim = intHeight
        End If
        
        ' For each pixel position on source bitmap
        For intCounterX = 0 To intMaxDim
            For intCounterY = 0 To intMaxDim
                ' Compute polar coordinate of source
                If intCounterX = 0 Then
                    sngTempAngle = HALF_PI
                Else
                    sngTempAngle = Atn(intCounterY / intCounterX)
                End If
                intRadius = Sqr((CLng(intCounterX) ^ 2) + (CLng(intCounterY) ^ 2))
                
                ' Compute rotated position of destination
                intPointX = intRadius * Cos(sngTempAngle + sngAngle)
                intPointY = intRadius * Sin(sngTempAngle + sngAngle)
                
                ' Copy pixels, 4 quadrants at once.
                lngCol(0) = GetPixel(lngSourcehDc, intCentreSrcX + intPointX, intCentreSrcY + intPointY)
                lngCol(1) = GetPixel(lngSourcehDc, intCentreSrcX - intPointX, intCentreSrcY - intPointY)
                lngCol(2) = GetPixel(lngSourcehDc, intCentreSrcX + intPointY, intCentreSrcY - intPointX)
                lngCol(3) = GetPixel(lngSourcehDc, intCentreSrcX - intPointY, intCentreSrcY + intPointX)
                
                Call SetPixel(TempBmp.lngDc, intCentreDestX + intCounterX, intCentreDestY + intCounterY, lngCol(0))
                Call SetPixel(TempBmp.lngDc, intCentreDestX - intCounterX, intCentreDestY - intCounterY, lngCol(1))
                Call SetPixel(TempBmp.lngDc, intCentreDestX + intCounterY, intCentreDestY - intCounterX, lngCol(2))
                Call SetPixel(TempBmp.lngDc, intCentreDestX - intCounterY, intCentreDestY + intCounterX, lngCol(3))
            Next intCounterY
        Next intCounterX
        
    End Select
    
    'copy the temperory bitmap to the destination Dc at the specified
    'co-ordinates
    lngResult = BitBlt(lngDesthDc, _
                       intDestX, _
                       intDestY, _
                       TempBmp.Area.Right, _
                       TempBmp.Area.Bottom, _
                       TempBmp.lngDc, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the temperory bitmap from memory and exit
    Call DeleteBitmap(TempBmp.lngDc, _
                      TempBmp.lngBitmap, _
                      TempBmp.lngPointer)
End Sub

Private Function CreateBitmapPicture(ByVal hBmp As Long, _
                                     ByVal hPal As Long, _
                                     Optional ByVal enmType As PictureTypeConstants = vbPicTypeBitmap) _
                                     As IPictureDisp
    'This function actually creates the picture object from a
    'processed bitmap. This must be called by hDcToPicture
    
    Dim lngResult       As Long         'holds any returned value from an api call
    Dim Pic             As PicBmp
    Dim IPic            As IPictureDisp
    Dim IID_IDispatch   As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic)    ' Length of structure
        .Type = enmType     ' Type of Picture (default: bitmap)
        .hBmp = hBmp        ' Handle to bitmap
        .hPal = hPal        ' Handle to palette (may be null)
    End With

    'Create the picture
    lngResult = OleCreatePictureIndirect(Pic, IID_IDispatch, 0, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function

Private Function hDCToPicture(ByVal lngDc As Long, _
                              ByVal lngBitmap As Long, _
                              ByVal lngPointer As Long, _
                              Optional ByVal enmType As PictureTypeConstants = vbPicTypeBitmap) _
                              As IPictureDisp
    'This will convert a hDc to a picture object
    
    Dim lngResult       As Long
    Dim hPal            As Long         'holds a handle to the pictures palette
    Dim hPalPrev        As Long         'holds a hanlde to the picture previous palette
    Dim RasterCapsScrn  As Long         'holds the raster capabilities of the screen
    Dim HasPaletteScrn  As Long         'holds the screen palette information
    Dim PaletteSizeScrn As Long         'holds the palette size for the screen
    Dim LogPal          As LOGPALETTE   'holds palette information

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(lngDc, RASTERCAPS) ' Raster
    
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(lngDc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        
        'Number of palette entries
        LogPal.palNumEntries = 256
        
        'Retrieve the system palette entries
        lngResult = GetSystemPaletteEntries(lngDc, _
                                            0, _
                                            256, _
                                            LogPal.palPalEntry(0))
        
        'Create the palette
        hPal = CreatePalette(LogPal)
        
        'Select the palette
        hPalPrev = SelectPalette(lngDc, hPal, 0)
        
        'Realize the palette
        lngResult = RealizePalette(lngDc)
    End If

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(lngDc, hPalPrev, 0)
    End If
    
    'return the picture object
    Set hDCToPicture = CreateBitmapPicture(lngBitmap, hPal, enmType)
    
    'delete the palette from memory
    lngResult = UnrealizeObject(hPal)
    lngResult = DeleteObject(hPal)
End Function

Private Sub PictureTohDc(ByRef picPicture As IPictureDisp)
    'This will convert a picture object into a bitmap DC
    
    Dim bytBits()       As Byte         'holds the bits of the bitmap
    Dim lngNumBits      As Long         'holds the number of bits in the array
    Dim lngHeight       As Long         'holds the height of the picture
    Dim lngWidth        As Long         'holds the width of the picture
    Dim lngResult       As Long         'holds any returned error value from an api call
    Dim udtInfo         As BITMAP       'holds information about the bitmap
    
    'make sure that an object was passed
    If picPicture Is Nothing Then
        Exit Sub
    End If
    
    With picPicture
        'get information about the picture
        lngResult = GetObjectAPI(.Handle, Len(udtInfo), udtInfo)
        lngWidth = udtInfo.bmWidth
        lngHeight = udtInfo.bmHeight
        'lngWidth = picPicture.Width / Screen.TwipsPerPixelX
        'lngHeight = picPicture.Height / Screen.TwipsPerPixelY
        
        With mudtBitmap
            'recreate the bitmap
            Call DeleteBitmap(.lngDc, .lngBitmap, .lngPointer)
            
            'set the bitmap size
            With .Area
                .Left = 0
                .Top = 0
                .Right = lngWidth
                .Bottom = lngHeight
            End With    '.Area
            
            'apply the picture to the bitmap
            .lngDc = CreateCompatibleDC(GetDC(0))
            .lngBitmap = picPicture.Handle
            lngResult = SelectObject(.lngDc, .lngBitmap)
        End With    'mudtBitmap
        
        'get the bits of the picture passed
        'lngNumBits = (lngHeight * 3 * (((lngWidth * 3) + 3) And &HFFFFFFFC))
        'If lngNumBits < 1 Then
        '    Exit Sub
        'End If
        'ReDim bytBits(1 To lngNumBits)
        'lngResult = GetBitmapBits(.Handle, lngNumBits, bytBits(1))
        'ReDim Preserve bytBits(1 To lngResult)
        
        'recreate the a new bitamp to match the size of the picture
        'Call SetBitmap(lngWidth, lngHeight)
        
        'copy the picture onto the bitmap
        'lngResult = SetBitmapBits(mudtBitmap.lngBitmap, _
                                  lngResult, _
                                  bytBits(1))
    End With
End Sub

Public Sub DrawRect(Optional ByVal lngColour As Long = NO_COLOUR, _
                    Optional ByVal intTop As Integer, _
                    Optional ByVal intLeft As Integer, _
                    Optional ByVal intHeight As Integer = -1, _
                    Optional ByVal intWidth As Integer = -1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels, _
                    Optional ByVal lngStyle As Long = BS_SOLID, _
                    Optional ByVal lngPattern As Long = HS_SOLID, _
                    Optional ByVal lnghDC As Long = 0)
    
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Dim StartRect   As Rect     'holds the dimensions of the rectangle to be drawn
    Dim lngResult   As Long     'holds any returned error value from the api calls
    Dim lngJunk     As Long     'holds any junk information from an api call
    Dim lnghBrush   As Long     'holds a pointer to a brush with the colour and pattern information for the rectangle
    Dim BrushStuff  As LogBrush 'holds the brush information
    
    'make sure a device context handle was passed, if not default to
    'the bitmaps
    If lnghDC = 0 Then
        lnghDC = mudtBitmap.lngDc
    End If
    
    'check the default height and widths
    With mudtBitmap.Area
        If intHeight < 0 Then
            intHeight = .Bottom - .Top
        End If
        If intWidth < 0 Then
            intWidth = .Right - .Left
        End If
    End With
    
    'initalise values
    StartRect.Top = intTop
    StartRect.Left = intLeft
    StartRect.Bottom = intLeft + intWidth
    StartRect.Right = intTop + intHeight
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert to pixels
        With StartRect
            .Left = .Left / Screen.TwipsPerPixelX
            .Top = .Top / Screen.TwipsPerPixelY
            .Right = .Right / Screen.TwipsPerPixelX
            .Bottom = .Bottom / Screen.TwipsPerPixelY
        End With
    End If
    
    'set the colour to the background colour if
    'none specified
    If lngColour = NO_COLOUR Then
        lngColour = mlngBackColour
    End If
    
    'create a brush
    BrushStuff.lbColor = lngColour
    BrushStuff.lbHatch = lngPattern
    BrushStuff.lbStyle = lngStyle
    
    'apply the brush to the device context
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lnghDC, lnghBrush)
    
    'draw a rectangle
    lngResult = PatBlt(lnghDC, _
                       intLeft, _
                       intTop, _
                       intWidth, _
                       intHeight, _
                       PATCOPY)
    
    'A "Brush" object was created. It must be removed from memory.
    lngJunk = SelectObject(lnghDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub DrawRoundRect(ByVal lngColour As Long, _
                         Optional ByVal intLeft As Integer = -32767, _
                         Optional ByVal intTop As Integer = -32767, _
                         Optional ByVal intRight As Integer = -32767, _
                         Optional ByVal intBottom As Integer = -32767, _
                         Optional ByVal intEdgeRadius As Integer = -1, _
                         Optional ByVal udtMeasurement As Scaling = InPixels, _
                         Optional ByVal lngStyle As Long = BS_SOLID, _
                         Optional ByVal lngPattern As Long = HS_SOLID, _
                         Optional ByVal lnghDC As Long = 0)
                         
    'this draws a rectangle using the co-ordinates
    'and lngColour given.
    
    Const Width = 1 'the pixel width of the edge
    
    Dim lnghPen     As Long     'holds a pointer to a Pen object
    Dim PenStuff    As LogPen   'holds the pen information
    Dim lnghBrush   As Long     'holds a pointer to a Brush object
    Dim BrushStuff  As LogBrush 'holds the brush information
    Dim lngJunk     As Long     'holds "Junk" information from an api call
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim OffsetRect  As Rect     'holds how much to grow/shrink the original size to accomidate the border (if any)
    
    'make sure a dc was specified, and if not, default to the bitmaps dc
    If lnghDC = 0 Then
        lnghDC = mudtBitmap.lngDc
    End If
    lngColour = GetSystemColour(lngColour)
    
    'set the default values
    If (intLeft = -32767) Then
        intLeft = Me.Left
    End If
    If (intTop = -32767) Then
        intTop = Me.Top
    End If
    If (intRight = -32767) Then
        intRight = Me.Left + Me.Width
    End If
    If (intBottom = -32767) Then
        intBottom = Me.Top + Me.Height
    End If
    If (intEdgeRadius < 0) Then
        'average out the width and height specified
        intEdgeRadius = ((intBottom - intTop) + (intRight - intLeft)) \ 4
    End If
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intLeft = intLeft / Screen.TwipsPerPixelX
        intTop = intTop / Screen.TwipsPerPixelX
        intRight = intRight / Screen.TwipsPerPixelY
        intBottom = intBottom / Screen.TwipsPerPixelY
        intEdgeRadius = intEdgeRadius / Screen.TwipsPerPixelX
    End If
    
    'Find out if a specific lngColour is
    'to be set. If so set it.
    
    'set lnghBrush settings (similar to
    'the FillColor property)
    BrushStuff.lbColor = lngColour
    BrushStuff.lbStyle = lngStyle
    BrushStuff.lbHatch = lngPattern  'ignored if lngStyle is solid
    
    'set lnghPen settings (similar to the
    'border properties on controls)
    PenStuff.lopnColor = lngColour
    PenStuff.lopnWidth.X = Width
    PenStuff.lopnStyle = PS_SOLID
    
    'apply the settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lnghPen = SelectObject(lnghDC, lnghPen)
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghBrush = SelectObject(lnghDC, lnghBrush)
    
    '---------in debug - 18/02/02
    'deflate the rectangle dimensions by
    'the radius amount
    OffsetRect.Left = intLeft
    OffsetRect.Top = intTop
    OffsetRect.Right = intRight
    OffsetRect.Bottom = intBottom
    'lngResult = InflateRect(OffsetRect, -intEdgeRadius, -intEdgeRadius)
    
    'draw the rounded rectangle
    lngResult = RoundRect(lnghDC, _
                          OffsetRect.Left, _
                          OffsetRect.Top, _
                          OffsetRect.Right, _
                          OffsetRect.Bottom, _
                          intEdgeRadius, _
                          intEdgeRadius)
    
    'delete the objects created (lnghPen and lnghBrush objects)
    lngJunk = SelectObject(lnghDC, lnghPen)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(lnghDC, lnghBrush)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Sub Invert(Optional ByVal intX As Integer = -32768, _
                  Optional ByVal intY As Integer = -32768, _
                  Optional ByVal intWidth As Integer = -1, _
                  Optional ByVal intHeight As Integer = -1)
    'This will invert the colours in the selected area
    
    Dim lngResult       As Long         'holds any returned value from an api call
    Dim udtArea         As Rect         'holds the dimensions of the area to be inverted
    
    'make sure that there is a bitmap
    If mudtBitmap.lngDc = 0 Then
        Exit Sub
    End If
    
    With mudtBitmap
        'make sure that values were passed, if not set defaults
        'to the bitmaps dimensions
        If intX = -32768 Then
            intX = .Area.Left
        End If
        If intY = -32768 Then
            intY = .Area.Top
        End If
        If intWidth < 0 Then
            intWidth = .Area.Right - .Area.Left
        End If
        If intHeight < 0 Then
            intHeight = .Area.Bottom - .Area.Top
        End If
    
        'put the dimensions into the rect structure
        With udtArea
            .Left = intX
            .Top = intY
            .Right = intX + intWidth
            .Bottom = intY + intHeight
        End With
        
        'invert the selected area
        lngResult = InvertRect(.lngDc, udtArea)
    End With
End Sub

Public Sub TitleToTray(ByRef frm As Form)
    'This function will draw the minimize animation from the form to
    'the title tray.
    
    'find the position of the title tray
    mlnghStartMenu = FindWindow("Shell_TrayWnd", vbNullString)
    mlnghChild = GetWindow(mlnghStartMenu, GW_CHILD)
    Do
        mlngClassNameLen = GetClassName(mlnghChild, _
                                       mstrClass, _
                                       Len(mstrClass))
        If InStr(1, mstrClass, "TrayNotifyWnd") Then
            mlnghTray = mlnghChild
            Exit Do
        End If
        mlnghChild = GetWindow(mlnghChild, GW_HWNDNEXT)
    Loop
    
    'animate the title bar to the title tray
    mlngRetVal = GetWindowRect(frm.hWnd, mudtRectFrom)
    mlngRetVal = GetWindowRect(mlnghTray, mudtRectTo)
    mlngRetVal = DrawAnimatedRects(frm.hWnd, _
                                  IDANI_OPEN Or IDANI_CAPTION, _
                                  mudtRectFrom, _
                                  mudtRectTo)
    
    'hide form
    frm.Visible = False
    frm.Hide
End Sub

Public Sub TrayToTitle(ByRef frm As Form)
    'This function draws the restore animation of the forms title bar, from
    'the system tray to the forms' position.
    
    'find the system trays position
    mlnghStartMenu = FindWindow("Shell_TrayWnd", _
                                vbNullString)
    mlnghChild = GetWindow(mlnghStartMenu, GW_CHILD)
    Do
        mlngClassNameLen = GetClassName(mlnghChild, _
                                       mstrClass, _
                                       Len(mstrClass))
        
        If InStr(1, mstrClass, "TrayNotifyWnd") Then
            mlnghTray = mlnghChild
            Exit Do
        End If
        
        mlnghChild = GetWindow(mlnghChild, _
        GW_HWNDNEXT)
    Loop
    
    'draw the animation
    mlngRetVal = GetWindowRect(frm.hWnd, mudtRectFrom)
    mlngRetVal = GetWindowRect(mlnghTray, mudtRectTo)
    mlngRetVal = DrawAnimatedRects(frm.hWnd, _
                                  IDANI_CLOSE Or IDANI_CAPTION, _
                                  mudtRectTo, _
                                  mudtRectFrom)
    
    'show the window
    frm.Visible = True
    frm.Show
End Sub

Public Sub DrawFrame(Optional ByVal intLeft As Integer, _
                     Optional ByVal intTop As Integer, _
                     Optional ByVal intWidth As Integer, _
                     Optional ByVal intHeight As Integer, _
                     Optional ByVal enmFrameStyle As EnumFrameStyle = DFC_BUTTON, _
                     Optional ByVal enmFrameState As EnumFrameState = DFCS_BUTTON3STATE, _
                     Optional ByVal hdc As Long, _
                     Optional ByVal udtMeasurement As Scaling = InPixels)
    'This will draw a frame around the specified co-ordinates in the specified style on the specified device
    'context. If no size measurements are specified, then the size and positions are assumed to be the current
    'settings for the object. If no device context is specified, then the current object is assumed. The other
    'defaults are shown above.
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim udtSize         As Rect     'holds the size of the border to draw
    
    'where do we get the measurements from
    With mudtBitmap.Area
        If (intLeft <= 0) Then
            intLeft = .Left
        End If
        If (intTop <= 0) Then
            intTop = .Top
        End If
        If (intWidth <= 0) Then
            intWidth = .Right - .Left
        End If
        If (intHeight <= 0) Then
            intHeight = .Bottom - .Top
        End If
    End With    'mudtBitmap.Area
    
    'put the measurements into a Rect structure
    With udtSize
        .Left = intLeft
        .Top = intTop
        .Right = intLeft + intWidth
        .Bottom = intTop + intHeight
    End With    'udtSize
    
    'what measurements are being used
    If (udtMeasurement = InTwips) Then
        Call RectToPixels(udtSize)
    End If  'what measurements are being used
    
    'what device was specified
    If (hdc = 0) Then
        hdc = mudtBitmap.lngDc
    End If
    
    'draw the specified border
    lngResult = DrawFrameControl(hdc, udtSize, enmFrameStyle, enmFrameState)
End Sub

Public Sub DrawLine(ByVal intX1 As Integer, _
                    ByVal intY1 As Integer, _
                    ByVal intX2 As Integer, _
                    ByVal intY2 As Integer, _
                    Optional ByVal lngColour As Long = vbBlack, _
                    Optional ByVal intWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This will draw a line from point1 to point2
    
    Const NUM_OF_POINTS         As Integer = 2
    
    Dim lngResult               As Long     'holds any returned error value from an api call
    Dim lnghPen                 As Long     'holds a pointer to a Pen object
    Dim lngPointer              As Long     'holds the pointer
    Dim PenStuff                As LogPen   'holds the Pen object information
    Dim lngJunk                 As Long     'holds any useless information returned from an api call
    Dim Points(NUM_OF_POINTS)   As PointAPI 'holds the co-ordinates of the line to be drawn
    Dim lngOldPtr               As Long     'used to unselect the pen object from the bitmap before deleting it
    
    'check if conversion is necessary
    If udtMeasurement = InTwips Then
        'convert twip values to pixels
        intX1 = intX1 / Screen.TwipsPerPixelX
        intX2 = intX2 / Screen.TwipsPerPixelX
        intY1 = intY1 / Screen.TwipsPerPixelY
        intY2 = intY2 / Screen.TwipsPerPixelY
    End If
    
    'Find out if a specific lngColour is to be set. If so set it.
    PenStuff.lopnColor = lngColour
    PenStuff.lopnStyle = PS_GEOMETRIC
    PenStuff.lopnWidth.X = intWidth
    
    'apply the pen settings to the device context
    lnghPen = CreatePenIndirect(PenStuff)
    lngOldPtr = SelectObject(mudtBitmap.lngDc, lnghPen)
    
    'set the points
    Points(1).X = intX1
    Points(1).Y = intY1
    Points(2).X = intX2
    Points(2).Y = intY2
    
    'draw the line
    lngResult = Polyline(mudtBitmap.lngDc, Points(1), NUM_OF_POINTS)
    
    'A "Pen" object was created. It must be removed from memory.
    lngJunk = SelectObject(mudtBitmap.lngDc, lngOldPtr)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = DeleteObject(lnghPen)
End Sub

Public Sub DrawPoly(ByRef intX() As Integer, _
                    ByRef intY() As Integer, _
                    Optional ByVal lngFillColour As Long = 0, _
                    Optional ByVal lngBorderColour As Long = 0, _
                    Optional ByVal intBorderWidth As Integer = 1, _
                    Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This function will draw a polygon of size and colours specified.
    'Modified to draw and fill all the points specified in the array,
    
    Dim lngSuccessful   As Long         'holds the result of any api call
    Dim intCounter      As Integer      'used to cycle through the points
    Dim Temp()          As PointAPI     'holds the passed points in a PointAPI structure
    Dim intPointNum     As Integer      'holds the upperbound of the Temp array
    Dim lnghBrush       As Long         'holds a pointer to the brush used for the polygon
    Dim BrushStuff      As LogBrush     'holds the brush details
    Dim lnghPen         As Long         'holds a pointer to the pen used for the polygon
    Dim PenStuff        As LogPen       'holds the pen details
    Dim intMax          As Integer      'holds the upperbound of the passed arrays
    Dim intMin          As Integer      'holds the lowerbound of the passed arrays
    Dim lngOldPtr       As Long         'holds the original pointer of the bitmap. This is used to unselect the pen object before deleting it
    Dim lngOldPtr2      As Long         'holds the original pointer of the bitmap
    
    'make sure that the arrays passed are of the same length
    If (UBound(intX) <> UBound(intY)) Or _
       (LBound(intX) <> LBound(intY)) Then
        'we cannot draw these points
        Exit Sub
    End If
    
    'get the upper and lower bounds of the arrays
    intMax = UBound(intX)
    intMin = LBound(intY)
    
    'find the number of points stored
    intPointNum = Abs((intMax) - (intMin)) '- 1  '-1 becase this is a zero based array
    
    'create a temperory array to hold the points
    ReDim Temp(intPointNum)
        
    'convert all points to pixels if necessary, otherwise, don't waste time.
    If udtMeasurement = InTwips Then
        'convert to pixels
        With Screen
            For intCounter = 0 To intPointNum
                Temp(intCounter).X = intX(intCounter + intMin) \ .TwipsPerPixelX
                Temp(intCounter).Y = intY(intCounter + intMin) \ .TwipsPerPixelY
            Next intCounter
        End With
    
    Else
        'move the points into the PointAPI structure
        For intCounter = 0 To intPointNum
            Temp(intCounter).X = intX(intCounter + intMin)
            Temp(intCounter).Y = intY(intCounter + intMin)
        Next intCounter
    End If
    
    'apply the border and background colours
    BrushStuff.lbColor = lngFillColour
    BrushStuff.lbHatch = 0
    BrushStuff.lbStyle = BS_SOLID
    
    PenStuff.lopnColor = lngBorderColour
    PenStuff.lopnWidth.X = intBorderWidth
    PenStuff.lopnStyle = PS_SOLID
    
    'create the objects necessary to apply the colour and draw settings
    lnghBrush = CreateBrushIndirect(BrushStuff)
    lnghPen = CreatePenIndirect(PenStuff)
    lngOldPtr = SelectObject(mudtBitmap.lngDc, lnghBrush)
    lngOldPtr2 = SelectObject(mudtBitmap.lngDc, lnghPen)
    
    'draw the polygon
    lngSuccessful = Polygon(mudtBitmap.lngDc, Temp(0), intPointNum + 1)
    
    'remove the pen and brush objects
    lngSuccessful = SelectObject(mudtBitmap.lngDc, lngOldPtr)
    lngSuccessful = DeleteObject(lngSuccessful)
    lngSuccessful = DeleteObject(lnghBrush)
    lngSuccessful = SelectObject(mudtBitmap.lngDc, lngOldPtr2)
    lngSuccessful = DeleteObject(lngSuccessful)
    lngSuccessful = DeleteObject(lnghPen)
End Sub

Public Sub LockWindow(ByVal FormName As Form)
    'Prevent the form from updating its display
    
    Dim lngResult As Boolean    'holds the returned error value from an api call
    
    lngResult = LockWindowUpdate(FormName.hWnd)
End Sub

Public Sub UnLockWindow()
    'see the procedure LockWindow
    
    Dim lngResult As Long   'holds the returned error value from an api call
    
    'Let the form update its display
    lngResult = LockWindowUpdate(0)
End Sub

Private Sub CreateNewBitmap(ByRef lngDc As Long, _
                            ByRef lngBitmap As Long, _
                            ByRef lngPointer As Long, _
                            ByRef udtBmpArea As Rect, _
                            Optional ByVal lngCompatableWithhDc As Long = 0, _
                            Optional ByVal udtMeasurement As Scaling = InPixels, _
                            Optional ByVal lngColour As Long = NO_COLOUR)
    
    'This procedure will create a new bitmap compatable with a given
    'form (you will also be able to then use this bitmap in a picturebox).
    'The space specified in "udtArea" should be in "Twips" and will be
    'converted into pixels in the following code.
    
    Dim lngResult           As Long             'holds any returned error value from the api calls
    Dim udtArea             As Rect             'holds the size of the bitmap
    Dim udtOldPos           As PointAPI         'holds the previous postion of the bitmap cursor
    
    'scale the bitmap points if necessary
    udtArea = udtBmpArea
    If udtMeasurement = InTwips Then
        Call RectToPixels(udtArea)
    End If
    
    'create the bitmap and its references
    If lngCompatableWithhDc = 0 Then
        'use the Dc for the desktop window
        lngCompatableWithhDc = GetDC(0)
        lngDc = CreateCompatibleDC(lngCompatableWithhDc)
        
        'create the bitmap
        lngBitmap = CreateCompatibleBitmap(lngCompatableWithhDc, _
                                           (udtArea.Right - udtArea.Left), _
                                           (udtArea.Bottom - udtArea.Top))
        
        lngResult = ReleaseDC(0, lngCompatableWithhDc)
    Else
        'make the dc compatable with the specifed window
        lngDc = CreateCompatibleDC(lngCompatableWithhDc)
    
        'create the bitmap
        lngBitmap = CreateCompatibleBitmap(lngCompatableWithhDc, _
                                           (udtArea.Right - udtArea.Left), _
                                           (udtArea.Bottom - udtArea.Top))
    End If
    
    'apply the bitmap to the device context
    lngPointer = SelectObject(lngDc, lngBitmap)
    
    'set the colour to the background colour if
    'none specified
    If lngColour = NO_COLOUR Then
        lngColour = mlngBackColour
    End If
    
    'set default colours and clear bitmap to selected colour
    lngResult = SetBkMode(lngDc, OPAQUE)
    lngResult = SetBkColor(lngDc, lngColour)
    Call DrawRect(lngColour, lnghDC:=lngDc)
    
    'make sure the bitmaps current pixel location is
    'pointing to 0x0
    lngResult = MoveToEx(lngDc, 0, 0, udtOldPos)
    udtOldPos.X = 0
    udtOldPos.Y = 0
    lngResult = SetBrushOrgEx(lngDc, 0, 0, udtOldPos)
End Sub

Private Sub DeleteBitmap(ByRef lngDc As Long, _
                         ByRef lngBitmap As Long, _
                         ByRef lngPointer As Long)
                        
    'This will remove the bitmap that stored what was displayed before
    'the text was written to the screen, from memory.
    
    Dim lngJunk As Long     'holds useless information from an api call
    
    If lngDc = 0 Then
        'there is nothing to delete. Exit the sub-routine
        Exit Sub
    End If
    
    'remove any clipping regions
    lngJunk = SelectClipRgn(lngDc, 0)
    
    'delete the device context
    lngJunk = SelectObject(lngDc, lngPointer)
    lngJunk = DeleteObject(lngBitmap)
    lngJunk = DeleteDC(lngDc)
    
    'show that the device context has been deleted by setting
    'all parameters passed to the procedure to zero
    lngDc = 0
    lngBitmap = 0
    lngPointer = 0
End Sub

Public Sub MergeBitmaps(ByVal hDcTextureBack As Long, _
                        ByVal hDcTextureFore As Long, _
                        ByVal hDcMask As Long, _
                        Optional ByVal intDestX As Integer = -32768, _
                        Optional ByVal intDestY As Integer = -32768, _
                        Optional ByVal intTextureBackX As Integer = -32768, _
                        Optional ByVal intTextureBackY As Integer = -32768, _
                        Optional ByVal intTextureForeX As Integer = -32768, _
                        Optional ByVal intTextureForeY As Integer = -32768, _
                        Optional ByVal intMaskX As Integer = -32768, _
                        Optional ByVal intMaskY As Integer = -32768, _
                        Optional ByVal intMaskWidth As Integer = -32768, _
                        Optional ByVal intMaskHeight As Integer = -32768, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)
                        
    'This procedure takes a monochrome (black & white) bitmap as a mask,
    '2 source texture bitmaps and a destination bitmap. It copies to the
    'destination bitmap a merge of the two source bitmaps, filling in the
    'pixels according to the mask. For each white pixel in the mask, it copies
    'the corresponding pixel from TextureBack. For each black pixel in the mask,
    'it copies the corresponding pixel from TextureFore to the destination
    'bitmap.
    'eg. Say the mask picture was a black square with a white letter "A" on
    'it. The first texture was of clouds and the second picture was of a
    'tartan design. The lngResult would be a cloud picture in the shape of an
    '"A" on a tartan background.
    
    Dim lngResult   As Long         'holds the returned error value from an api call
    Dim TempMaskBmp As BitmapStruc  'holds the mask bitmap
    Dim TempBackBmp As BitmapStruc  'holds the bitmap you wish to use as a background
    Dim intConvertX As Integer      'holds the height in twips of a pixel
    Dim intConvertY As Integer      'holds the width in twips of a pixel
    
    'check if default values were passed. Default to the current
    'size of the bitmap
    With mudtBitmap.Area
        If (intDestX = -32768) Then
            intDestX = .Left
        End If
        If (intDestY = -32768) Then
            intDestY = .Top
        End If
        If (intTextureBackX = -32768) Then
            intTextureBackX = .Left
        End If
        If (intTextureBackY = -32768) Then
            intTextureBackY = .Top
        End If
        If (intTextureForeX = -32768) Then
            intTextureForeX = .Left
        End If
        If (intTextureForeY = -32768) Then
            intTextureForeY = .Top
        End If
        If (intMaskX = -32768) Then
            intMaskX = .Left
        End If
        If (intMaskY = -32768) Then
            intMaskY = .Top
        End If
        If (intMaskWidth = -32768) Then
            intMaskWidth = .Right - .Left
        End If
        If (intMaskHeight = -32768) Then
            intMaskHeight = .Bottom - .Top
        End If
    End With
    
    'convert passed values if necessary
    If udtMeasurement = InTwips Then
        intConvertX = Screen.TwipsPerPixelX
        intConvertY = Screen.TwipsPerPixelY
        
        intDestX = intDestX / intConvertX
        intDestY = intDestY / intConvertY
        intTextureBackX = intTextureBackX / intConvertX
        intTextureBackY = intTextureBackY / intConvertY
        intTextureForeX = intTextureForeX / intConvertX
        intTextureForeY = intTextureForeY / intConvertY
        intMaskWidth = intMaskWidth / intConvertX
        intMaskHeight = intMaskHeight / intConvertY
    End If
    
    'create temperory bitmaps
    TempMaskBmp.Area.Right = intMaskWidth
    TempMaskBmp.Area.Bottom = intMaskHeight
    TempBackBmp.Area = TempMaskBmp.Area
    Call CreateNewBitmap(TempMaskBmp.lngDc, _
                         TempMaskBmp.lngBitmap, _
                         TempMaskBmp.lngPointer, _
                         TempMaskBmp.Area, _
                         hDcMask, _
                         InPixels, _
                         vbBlack)
    Call CreateNewBitmap(TempBackBmp.lngDc, _
                         TempBackBmp.lngBitmap, _
                         TempBackBmp.lngPointer, _
                         TempBackBmp.Area, _
                         hDcMask, _
                         InPixels, _
                         vbBlack)
    
    'create a white bitmap with a mask shaped hole onto the
    'destination background
    lngResult = BitBlt(TempMaskBmp.lngDc, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureFore, _
                       intTextureForeX, _
                       intTextureForeY, _
                       P_COPY)
    lngResult = StretchBlt(TempBackBmp.lngDc, _
                           0, _
                           0, _
                           intMaskWidth, _
                           intMaskHeight, _
                           hDcMask, _
                           intMaskX, _
                           intMaskY, _
                           intMaskWidth, _
                           intMaskHeight, _
                           P_MERGE_PAINT)
    lngResult = BitBlt(TempMaskBmp.lngDc, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.lngDc, _
                       0, _
                       0, _
                       P_MERGE_PAINT)
    
    'draw a white mask shape onto the second texture
    'where the black mask is
    lngResult = BitBlt(TempBackBmp.lngDc, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       hDcTextureBack, _
                       intTextureBackX, _
                       intTextureBackY, _
                       P_COPY)
    lngResult = StretchBlt(TempBackBmp.lngDc, _
                           0, _
                           0, _
                           intMaskWidth, _
                           intMaskHeight, _
                           hDcMask, _
                           intMaskX, _
                           intMaskY, _
                           intMaskWidth, _
                           intMaskHeight, _
                           P_MERGE_PAINT)

    'merge the two masks
    lngResult = BitBlt(TempMaskBmp.lngDc, _
                       0, _
                       0, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempBackBmp.lngDc, _
                       0, _
                       0, _
                       P_AND)
    
    'copy the merged lngResult to the destination
    lngResult = BitBlt(mudtBitmap.lngDc, _
                       intDestX, _
                       intDestY, _
                       intMaskWidth, _
                       intMaskHeight, _
                       TempMaskBmp.lngDc, _
                       0, _
                       0, _
                       P_COPY)
    
    'remove the two temperory bitmaps from memory
    Call DeleteBitmap(TempMaskBmp.lngDc, _
                      TempMaskBmp.lngBitmap, _
                      TempMaskBmp.lngPointer)
    Call DeleteBitmap(TempBackBmp.lngDc, _
                      TempBackBmp.lngBitmap, _
                      TempBackBmp.lngPointer)
End Sub

Private Sub RectToTwips(ByRef udtRect As Rect)
    'converts pixels to twips in a rect structure
    
    With udtRect
        .Left = .Left * Screen.TwipsPerPixelX
        .Right = .Right * Screen.TwipsPerPixelX
        .Top = .Top * Screen.TwipsPerPixelY
        .Bottom = .Bottom * Screen.TwipsPerPixelY
    End With
End Sub

Private Sub RectToPixels(ByRef udtRect As Rect)
    'converts twips to pixels in a rect structure
    
    With udtRect
        .Left = .Left / Screen.TwipsPerPixelX
        .Right = .Right / Screen.TwipsPerPixelX
        .Top = .Top / Screen.TwipsPerPixelY
        .Bottom = .Bottom / Screen.TwipsPerPixelY
    End With
End Sub

Public Function GetTick() As Long
    'returns the current windows tick
    GetTick = GetTickCount
End Function

Public Function IsFormActive(ByVal frmName As Form) _
                             As Boolean
    'This function returns wether or not the window
    'is active
    
    If frmName.hWnd = GetActiveWindow Then
        IsFormActive = True
    Else
        IsFormActive = False
    End If
End Function

Private Function IsRectEmpty(ByRef udtRect As Rect) _
                             As Boolean
    'This will return True if the rect is empty
    
    IsRectEmpty = Not IsRectEmptyApi(udtRect)
End Function

Public Sub GrayScale(Optional ByVal intLeft As Integer = -32768, _
                     Optional ByVal intTop As Integer = -32768, _
                     Optional ByVal intWidth As Integer = -1, _
                     Optional ByVal intHeight As Integer = -1)
    'This will convert the current picture colours to gray for the specifed
    'bitmap area. If no area is specified, then the procedure will gray the
    'entire current bitmap
    
    Dim intX        As Integer      'used to cycle through the rows of the bitmap
    Dim intY        As Integer      'used to cycle through the column of the bitmap
    Dim intRed      As Integer      'holds the current red colour of a pixel
    Dim intGreen    As Integer      'holds the current green colour of a pixel
    Dim intBlue     As Integer      'holds the current blue colour of a pixel
    Dim intGray     As Integer      'holds the balance gray colour of a pixel
    Dim lngColour   As Long         'holds the RGB colour of a pixel
    Dim lngResult   As Long         'holds any returned error value from an api call
    
    'check the parameters passed
    With mudtBitmap.Area
        If (intLeft = -32768) Then
            intLeft = .Left
        End If
        If (intTop = -32768) Then
            intTop = .Top
        End If
        If (intWidth < 0) Then
            intWidth = .Right - intLeft
        End If
        If (intHeight < 0) Then
            intHeight = .Bottom - intTop
        End If
    End With    'mudtbitmap.Area
    
    'go through the bitmap graying the colours
    For intX = intLeft To (intLeft + intWidth)
        For intY = intTop To (intTop + intHeight)
            'get the colour of the specified pixel
            lngColour = GetPixel(mudtBitmap.lngDc, intX, intY)
            Call GetRGB(lngColour, intRed, intGreen, intBlue)
            
            'convert the colour to gray
            intGray = (intRed + intGreen + intBlue) \ 3
            lngColour = RGB(intGray, intGray, intGray)
            
            'set the gray colour
            lngResult = SetPixel(mudtBitmap.lngDc, intX, intY, lngColour)
        Next intY
    Next intX
End Sub

Public Sub Gradient(ByVal lngStartCol As Long, _
                    ByVal lngFinishCol As Long, _
                    ByVal intLeft As Integer, _
                    ByVal intTop As Integer, _
                    ByVal intWidth As Integer, _
                    ByVal intHeight As Integer, _
                    Optional ByVal Direction As GradientTo = GradHorizontal, _
                    Optional ByVal bytLineWidth As Byte = 1)
    'ByVal mudtbitmap.hdcmemory As Long, _

    'draws a gradient from colour udtStart to colour udtFinish, and assums
    'that all measurments passed to it are in pixels unless otherwise
    'specified.
    
    Dim intCounter      As Integer  'cycles through each step in the gradient
    Dim intBiggestDiff  As Integer  'the biggest difference in either R,G or B values for the start and finish (the number of steps the gradient has to go through)
    Dim udtColour       As RGBVal   'temperory colour space for each step in the gradient
    Dim udtStart        As RGBVal   'the starting colour
    Dim udtFinish       As RGBVal   'the finishing colour
    Dim sngAddRed       As Single   'temperorily holds the current Red value
    Dim sngAddGreen     As Single   'temperorily holds the current Green value
    Dim sngAddBlue      As Single   'temperorily holds the current Blue value
    Dim intColRed       As Integer  'used to get the starting and finishing colours
    Dim intColGreen     As Integer  'used to get the starting and finishing colours
    Dim intColBlue      As Integer  'used to get the starting and finishing colours
    
    'perform all necessary calculations before drawing gradient
    'such as converting long to rgb values, and getting the correct
    'scaling for the bitmap.
    With udtStart
        Call GetRGB(lngStartCol, _
                    intColRed, _
                    intColGreen, _
                    intColBlue)
        
        'store the colours
        .Red = intColRed
        .Green = intColGreen
        .Blue = intColBlue
    End With
    With udtFinish
        Call GetRGB(lngFinishCol, _
                    intColRed, _
                    intColGreen, _
                    intColBlue)
        
        'store the colours
        .Red = intColRed
        .Green = intColGreen
        .Blue = intColBlue
    End With
    
    'draw the colour gradient
    Select Case Direction
    Case GradVertical
        intBiggestDiff = intWidth
    Case GradHorizontal
        intBiggestDiff = intHeight
    End Select
    
    'calculate how much to increment/decrement each colour per step
    sngAddRed = (bytLineWidth * ((udtFinish.Red) - udtStart.Red) / intBiggestDiff)
    sngAddGreen = (bytLineWidth * ((udtFinish.Green) - udtStart.Green) / intBiggestDiff)
    sngAddBlue = (bytLineWidth * ((udtFinish.Blue) - udtStart.Blue) / intBiggestDiff)
    udtColour = udtStart
    
    'calculate the colour of each line before drawing it on the bitmap
    For intCounter = 0 To intBiggestDiff Step bytLineWidth
        'find the point between colour udtStart and udtColour udtFinish that
        'corresponds to the point between 0 and intBiggestDiff
        
        'check for overflow
        If udtColour.Red > 255 Then
            udtColour.Red = 255
        Else
            If udtColour.Red < 0 Then
                udtColour.Red = 0
            End If
        End If
        If udtColour.Green > 255 Then
            udtColour.Green = 255
        Else
            If udtColour.Green < 0 Then
                udtColour.Green = 0
            End If
        End If
        If udtColour.Blue > 255 Then
            udtColour.Blue = 255
        Else
            If udtColour.Blue < 0 Then
                udtColour.Blue = 0
            End If
        End If
        
        'draw the gradient in the proper orientation in the calculated colour
        Select Case Direction
        Case GradVertical
            Call DrawLine(intCounter + intLeft, _
                          intTop, _
                          intCounter + intLeft, _
                          intHeight + intTop, _
                          RGB(udtColour.Red, udtColour.Green, udtColour.Blue), _
                          bytLineWidth, _
                          InPixels)
        Case GradHorizontal
            Call DrawLine(intLeft, _
                          intCounter + intTop, _
                          intLeft + intWidth, _
                          intTop + intCounter, _
                          RGB(udtColour.Red, udtColour.Green, udtColour.Blue), _
                          bytLineWidth, _
                          InPixels)
        End Select
        
        'set next colour
        udtColour.Red = udtColour.Red + sngAddRed
        udtColour.Green = udtColour.Green + sngAddGreen
        udtColour.Blue = udtColour.Blue + sngAddBlue
    Next intCounter
End Sub

'not currently in use - 2/June/2002
'Private Sub FadeGradient(ByVal intDestLeft As Integer, _
'                        ByVal intDestTop As Integer, _
'                        ByVal intDestWidth As Integer, _
'                        ByVal intDestHeight As Integer, _
'                        ByVal lngGradhDc As Long, _
'                        ByVal lngStartFromA As Long, _
'                        ByVal lngFinishToA As Long, _
'                        ByVal lngStartFromB As Long, _
'                        ByVal lngFinishToB As Long, _
'                        ByVal intLeft As Integer, _
'                        ByVal intTop As Integer, _
'                        ByVal intWidth As Integer, _
'                        ByVal intHeight As Integer, _
'                        ByVal udtDirection As GradientTo, _
'                        Optional ByVal udtMesurement As Scaling = 1, _
'                        Optional ByVal bytLineWidth As Byte = 1)
'
'    'This procedure will call the Gradient function to fade it into
'    'the udtColours specified.
'    'Note : all mesurements must me of the same scale, ie they must all
'    'be in pixels or all in twips.
'
'    Dim udtColour(2) As RGBVal
'    Dim udtStart(2) As RGBVal
'    Dim udtFinish(2) As RGBVal
'    Dim lngGradCol(2) As Long
'    Dim intCounter As Integer
'    Dim intBiggestDiff As Integer
'    Dim intValue As Integer
'    Dim intIndex As Integer
'    Dim lngResult As Long
'
'    Const A = 0
'    Const B = 1
'
'    'convert to RGB values
'    udtStart(A) = GetRGB(lngStartFromA)
'    udtStart(B) = GetRGB(lngStartFromB)
'    udtFinish(A) = GetRGB(lngFinishToA)
'    udtFinish(B) = GetRGB(lngFinishToB)
'
'    'convert to pixels if necessary
'    If udtMesurement = InTwips Then
'        intDestLeft = intDestLeft / Screen.TwipsPerPixelX
'        intDestTop = intDestTop / Screen.TwipsPerPixelY
'        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
'        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
'        intLeft = intLeft / Screen.TwipsPerPixelX
'        intTop = intTop / Screen.TwipsPerPixelY
'        intWidth = intWidth / Screen.TwipsPerPixelX
'        intHeight = intHeight / Screen.TwipsPerPixelY
'    End If
'
'
'    'Find the largest difference between any two corresponding
'    'udtColours, and use that as the number of steps to take in the loop,
'    '(therefore guarenteing that it will cycle through all necessary
'    'udtColours without jumping)
'    For intIndex = A To B
'        'test red
'        intValue = Abs(udtStart(intIndex).Red - udtFinish(intIndex).Red)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'
'        'test green
'        intValue = Abs(udtStart(intIndex).Green - udtFinish(intIndex).Green)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'
'        'test blue
'        intValue = Abs(udtStart(intIndex).Blue - udtFinish(intIndex).Blue)
'        If intValue > intBiggestDiff Then
'            intBiggestDiff = intValue
'        End If
'    Next intIndex
'
'    'if there is no difference, then just draw one
'    'gradient and exit
'    If intBiggestDiff = 0 Then
'        'Call Gradient(lngGradhDc, _
'                      lngStartFromA, _
'                      lngStartFromB, _
'                      intLeft, _
'                      intTop, _
'                      intWidth, _
'                      intHeight, _
'                      udtDirection, _
'                      InPixels, _
'                      bytLineWidth)
'        Exit Sub
'    End If
'
'    'fade the gradient
'    For intCounter = 0 To intBiggestDiff
'        'find the point between udtColour udtStart and udtColour udtFinish that
'        'corresponds to the point between 0 and intBiggestDiff
'
'        For intIndex = A To B
'            udtColour(intIndex).Red = udtStart(intIndex).Red + (((udtFinish(intIndex).Red - udtStart(intIndex).Red) / intBiggestDiff) * intCounter)
'            udtColour(intIndex).Green = udtStart(intIndex).Green + (((udtFinish(intIndex).Green - udtStart(intIndex).Green) / intBiggestDiff) * intCounter)
'            udtColour(intIndex).Blue = udtStart(intIndex).Blue + (((udtFinish(intIndex).Blue - udtStart(intIndex).Blue) / intBiggestDiff) * intCounter)
'
'            'convert to long intValue and store
'            lngGradCol(intIndex) = RGB(udtColour(intIndex).Red, _
'                                       udtColour(intIndex).Green, _
'                                       udtColour(intIndex).Blue)
'        Next intIndex
'
'        'draw the gradient onto the bitmap
'        'Call Gradient(lngGradhDc, _
'                      lngGradCol(A), _
'                      lngGradCol(B), _
'                      intLeft, _
'                      intTop, _
'                      intWidth, _
'                      intHeight, _
'                      udtDirection, _
'                      InPixels, _
'                      bytLineWidth)
'
'        'blitt the bitmap to the screen
'        lngResult = BitBlt(mudtBitmap.lngDC, _
'                           intDestLeft, _
'                           intDestTop, _
'                           intDestWidth, _
'                           intDestHeight, _
'                           lngGradhDc, _
'                           intLeft, _
'                           intTop, _
'                           P_COPY)
'        DoEvents
'    Next intCounter
'End Sub

Public Function FromRGB(ByVal bytRed As Byte, _
                        ByVal bytGreen As Byte, _
                        ByVal bytBlue As Byte) _
                        As Long
    'Convert RGB to Long
     
    Dim lngMyVal As Long
    
    lngMyVal = (CLng(bytBlue) * 65536) + (CLng(bytGreen) * 256) + bytRed
    FromRGB = lngMyVal
End Function

Public Sub HideMouse()
    'hides the mouse cursor
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = ShowCursor(False)
End Sub

Public Sub ShowMouse()
    'shows the mouse cursor (if hidden)
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = ShowCursor(True)
End Sub

Public Function GetActivehWnd() As Long
    'returns a handle to the active window
    GetActivehWnd = GetActiveWindow
End Function

Public Function GetKeyCode(Optional ByVal enmKeyState As EnumKeyState = AS_KEY_PRESSED) _
                           As Integer
    'This will return the keycode of the first key we find pressed,
    'including the mouse buttons. If no keys are being pressed, then
    '-1 is returned
    
    Const KEY_PRESSED       As Integer = -32768     'a flag used to tell if a key is being pressed

    
    Dim lngCounter          As Long         'used to cycle through all of the possible characters
    Dim lngResult           As Long         'holds any returned value from an api call
    
    'the default retrun value is -1 or no keys being pressed
    GetKeyCode = -1
    
    'check all possible keys
    For lngCounter = 0 To 255
        'get the state of the current key
        lngResult = GetAsyncKeyState(lngCounter)
        
        'is it pressed
        Select Case lngResult
        Case KEY_PRESSED
            If enmKeyState = AS_KEY_PRESSED Then
                'return this key
                GetKeyCode = lngCounter
                Exit For
            End If
        End Select
    Next lngCounter
End Function

Public Function GetKeyState(ByVal intKeyVal As Integer) _
                            As Integer
    'This will return the state of the specified key
    
    GetKeyState = GetAsyncKeyState(intKeyVal)
End Function

Public Function GetColourBetween(ByVal lngFirstCol As Long, _
                                 ByVal lngSecondCol As Long, _
                                 Optional ByVal sngBlendPercent = 50) _
                                 As Long
    'This will blend two specified colours to produce a third that is somewhere
    'between the two colours. The BlendPercent parameter specified how much
    'to favour the first colour over the second colour.
    
    Dim intRed(2)      As Integer       'holds the red amount for all three colours
    Dim intGreen(2)    As Integer       'holds the green amount for all three colours
    Dim intBlue(2)     As Integer       'holds the blue amount for all three colours
    
    'validate the blend amount so that it is now between 0 and 1
    If (sngBlendPercent < 0) Then
        sngBlendPercent = 0
    ElseIf (sngBlendPercent > 100) Then
        sngBlendPercent = 100
    End If
    sngBlendPercent = sngBlendPercent / 100
    
    'get the RGB values for the two specified colours
    Call GetRGB(lngFirstCol, intRed(0), intGreen(0), intBlue(0))
    Call GetRGB(lngSecondCol, intRed(1), intGreen(1), intBlue(1))
    
    'get the colour between these two values
    intRed(2) = intRed(0) + ((intRed(1) - intRed(0)) * sngBlendPercent)
    intGreen(2) = intGreen(0) + ((intGreen(1) - intGreen(0)) * sngBlendPercent)
    intBlue(2) = intBlue(0) + ((intBlue(1) - intBlue(0)) * sngBlendPercent)
    
    'return the blended colour
    GetColourBetween = RGB(intRed(2), intGreen(2), intBlue(2))
End Function

Public Function GetFontPointSize(ByVal fntDetails As StdFont, _
                                 Optional ByVal strText As String = "W", _
                                 Optional ByVal intHeight As Integer = 0, _
                                 Optional ByVal intWidth As Integer = 0, _
                                 Optional ByVal enmMeasurement As Scaling = InPixels) _
                                 As Integer
    'This will return the largest point size of a font that can be
    'contained in the specified height and with for the specified
    'font and text being displayed
    
    Dim intSizeHeight   As Integer          'holds the largest font size for the specified height
    Dim intSizeWidth    As Integer          'holds the largest font size for the specified width
    Dim lngResult       As Long             'holds any returned error value from an api call
    Dim udtMetrics      As SizeType         'holds the metric details for the text being displayed with the font
    Dim hFont           As Long             'holds a handle to the font
    Dim hOldFont        As Long             'holds a handle to the previous font
    Dim udtTemp         As BitmapStruc      'holds a temperory bitmap to apply the font to
    
    'make sure a font was passed
    If fntDetails Is Nothing Then
        Exit Function
    End If
    
    'convert the values from twips to pixels if necessary
    If enmMeasurement = InTwips Then
        intHeight = intHeight \ Screen.TwipsPerPixelY
        intWidth = intWidth \ Screen.TwipsPerPixelX
    End If
    
    'make sure appropiate dimensions were passed
    If (intHeight < 1) And (intWidth < 1) Then
        'return the same size
        GetFontPointSize = fntDetails.Size
        Exit Function
    End If
    
    'if not text was specified, then set to default
    If strText = "" Then
        strText = "W"
    End If
    
    'create a one pixel bitmap to apply the font to
    With udtTemp
        .Area.Bottom = 1
        .Area.Right = 1
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
    
    'apply the font and get the total size of the text in the
    'specified font
    fntDetails.Size = 8        'just a figure used to get an accurate result
    hFont = ApplyFont(fntDetails, hOldFont, udtTemp.lngDc)
    lngResult = GetTextExtentPoint(udtTemp.lngDc, _
                                   strText, _
                                   Len(strText), _
                                   udtMetrics)
    
    'delete the font objects created by ApplyFont
    lngResult = SelectObject(udtTemp.lngDc, hOldFont)
    lngResult = DeleteObject(lngResult)
    lngResult = SelectObject(udtTemp.lngDc, hFont)
    lngResult = DeleteObject(lngResult)
    
    'now that we have our details, delete the bitmap
    With udtTemp
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
    End With
    
    'calculate the maximum point size that can be contained in
    'the specified height
    If intHeight > 0 Then
        If (udtMetrics.cy / fntDetails.Size) <> 0 Then
            intSizeHeight = intHeight / (udtMetrics.cy / fntDetails.Size)
        End If
    End If
    
    'calculate the maximum point size that can be containted in
    'the specified width
    If intWidth > 0 Then
        If (udtMetrics.cx / fntDetails.Size) <> 0 Then
            intSizeWidth = intWidth / (udtMetrics.cx / fntDetails.Size)
        End If
    End If
    
    'return the appropiate size
    If (intHeight > 0) And (intWidth > 0) Then
        'return the smalled font size for the two dimensions
        If intSizeHeight < intSizeWidth Then
            GetFontPointSize = intSizeHeight
        Else
            GetFontPointSize = intSizeWidth
        End If
    
    Else
        'return the largest font size for the two dimensions
        If intSizeHeight > intSizeWidth Then
            GetFontPointSize = intSizeHeight
        Else
            GetFontPointSize = intSizeWidth
        End If
    End If
'
'    'the smallest point size if usually 8
'    If GetFontPointSize < 8 Then
'        GetFontPointSize = 8
'    End If
'
'    'make sure the font is not too large
'    If GetFontPointSize > 300 Then
'        GetFontPointSize = 300
'    End If
End Function

Public Function GetHypLength(ByVal intX1 As Integer, _
                             ByVal intY1 As Integer, _
                             ByVal intX2 As Integer, _
                             ByVal intY2 As Integer) _
                             As Long
    'This will return the length of a line between the specified co-ordinates.
    'It does this by Pythagorus' Theorim which states that "The square on the
    'hypotheneus is equal to the sum of the squares on the other two sides"
    
    Dim intDiffX        As Integer      'holds the distance in the X axis
    Dim intDiffY        As Integer      'holds the distance in the Y axis
    
    'get the distance, ignoring negatives as the length is always positive
    intDiffX = Abs(intX2 - intX1)
    intDiffY = Abs(intY2 - intY1)
    
    'apply Pythagorus' Theorim
    GetHypLength = Sqr((intDiffX ^ 2) + (intDiffY ^ 2))
End Function

Public Function GetPi() As Double
    'This will return Pi
    GetPi = PI
End Function

Public Sub GetRGB(ByVal lngColour As Long, _
                  Optional ByRef intRed As Integer, _
                  Optional ByRef intGreen As Integer, _
                  Optional ByRef intBlue As Integer)

    'Convert Long to RGB:
    
                                                '        Sys Cols,  Blue  , Green  ,  Red
    Const RED_MASK      As Long = 255           'Binary: 00000000,00000000,00000000,11111111
    Const GREEN_MASK    As Long = 65280         'Binary: 00000000,00000000,11111111,00000000
    Const BLUE_MASK     As Long = 16711680      'Binary: 00000000,11111111,00000000,00000000
    Const SYS_MASK      As Double = 2147483648# 'Binary: 01000000,00000000,00000000,00000000
    Const MAX_COLOUR    As Long = 16777215      'Binary: 00000000,11111111,11111111,11111111
    
    If ((SYS_MASK - Abs(lngColour)) <= 28) Then
        'convert a vb tagged system colour into a valid
        'RGB colour value
        
        'convert to index and get the colour
        lngColour = SYS_MASK - Abs(lngColour) 'lngColour Or SYS_MASK
        lngColour = GetSysColor(lngColour)
    Else
    
        'if the colour value is greater than acceptable
        'then exit
        If (lngColour > MAX_COLOUR) Or _
           (lngColour < -MAX_COLOUR) Then
            'invalid colour range
            Exit Sub
        End If
    End If
    
    'get the red, green and blue values
    intRed = (lngColour And RED_MASK)
    intGreen = (lngColour And GREEN_MASK) \ 256 'shift 8 bits to the right
    intBlue = (lngColour And BLUE_MASK) \ 65536 'shift 16 bits to the right
End Sub

Public Sub Pause(ByVal lngTicks As Long, _
                 Optional ByVal blnSleep As Boolean = False)
    'pause execution of the program for a specified
    'number of ticks, or stop execution of the thread
    'entirly for the period of the thread
    
    Dim lngStart As Long    'marks the starting tick
    
    'should we pause the thread of allow events to
    'operate as normal
    If blnSleep Then
        'don't allow events
        If lngTicks < 0 Then
            lngTicks = 0
        End If
        Call Sleep(lngTicks)
    Else
        'allow events
        lngStart = GetTickCount
        Do While GetTickCount < (lngStart + lngTicks)
            DoEvents
        Loop
    End If
End Sub

Public Sub DrawBorder(Optional ByVal intLeft As Integer, _
                      Optional ByVal intTop As Integer, _
                      Optional ByVal intWidth As Integer, _
                      Optional ByVal intHeight As Integer, _
                      Optional ByVal enmEdgeStyle As EnumEdges = BDR_ETCHED, _
                      Optional ByVal enmEdgeSides As EnumEdgeSides = BF_RECT, _
                      Optional ByVal hdc As Long, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)
    'This will draw a border around the specified co-ordinates in the specified style on the specified device
    'context. If no size measurements are specified, then the size and positions are assumed to be the current
    'settings for the object. If no device context is specified, then the current object is assumed. The other
    'defaults are shown above.
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim udtSize         As Rect     'holds the size of the border to draw
    
    'where do we get the measurements from
    With mudtBitmap.Area
        If (intLeft <= 0) Then
            intLeft = .Left
        End If
        If (intTop <= 0) Then
            intTop = .Top
        End If
        If (intWidth <= 0) Then
            intWidth = .Right - .Left
        End If
        If (intHeight <= 0) Then
            intHeight = .Bottom - .Top
        End If
    End With    'mudtBitmap.Area
    
    'put the measurements into a Rect structure
    With udtSize
        .Left = intLeft
        .Top = intTop
        .Right = intLeft + intWidth
        .Bottom = intTop + intHeight
    End With    'udtSize
    
    'what measurements are being used
    If (udtMeasurement = InTwips) Then
        Call RectToPixels(udtSize)
    End If  'what measurements are being used
    
    'what device was specified
    If (hdc = 0) Then
        hdc = mudtBitmap.lngDc
    End If
    
    'draw the specified border
    lngResult = DrawEdge(hdc, udtSize, enmEdgeStyle, enmEdgeStyle)
End Sub

Public Sub DrawEllipse(ByVal intCenterX As Integer, _
                       ByVal intCenterY As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal intWidth As Integer, _
                       Optional ByVal sngTiltAngle As Single = 90, _
                       Optional ByVal lngColour As Long = 0, _
                       Optional ByVal intThickness As Integer = 1, _
                       Optional ByVal blnIsHollow As Boolean = True, _
                       Optional ByVal udtMesurement As Scaling = InPixels)
                       
    'This procedure will draw an ellipse of the specified dimensions and
    'lngColour, by drawing a line between each of the 360 points that make
    'up the ellipse.
    
    Const a     As Integer = 1
    Const B     As Integer = 2
    
    Dim sngMoveCenterX  As Single   'holds the co-ordinates of the centre point of the moving circle
    Dim sngMoveCenterY  As Single   'holds the co-ordinates of the centre point of the moving circle
    Dim sngCircleX      As Single   'holds the outer co-ordinates of the ellipse
    Dim sngCircleY      As Single   'holds the outer co-ordinates of the ellipse
    Dim sngCounter      As Single   'cycles through each dot on the ellipse
    Dim sngTiltX        As Single   'holds the co-ordinates of the tilted circle
    Dim sngTiltY        As Single   'holds the co-ordinates of the tilted circle
    Dim sngNumOfPoints  As Single   'the number of points used to draw the ellipse
    Dim udtEllipse()    As PointAPI 'holds a list of points on the ellipse
    Dim udtBrushStuff   As LogBrush 'holds the Brush information
    Dim lnghBrush       As Long     'holds a pointer to the brush
    Dim udtPenStuff     As LogPen   'holds the Pen information
    Dim lnghPen         As Long     'holds a pointer to the pen
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngJunk         As Long     'holds junk information from an api call
    Dim sngDegPerPoint  As Single   'holds the number of degrees to move to a new point on the ellipse
    Dim lngOldPtr       As Long     'holds the first pointer to the bitmap - this is used later to unselect the brush and pen
    Dim lngOldPtr2      As Long     'holds the second pointer to the bitmap
    
    'set scaling values
    If udtMesurement = InTwips Then
        'convert parameters to pixels
        intCenterX = (intCenterX / Screen.TwipsPerPixelX) '- intThickness
        intCenterY = (intCenterY / Screen.TwipsPerPixelY) '- intThickness
        intHeight = (intHeight / Screen.TwipsPerPixelY) - (intThickness * 2)
        intWidth = (intWidth / Screen.TwipsPerPixelX) - (intThickness * 2)
        
        'values are now in pixels
        udtMesurement = InPixels
    End If
    lngColour = GetSystemColour(lngColour)
    
    'calculate the radius for intWidth and intHeight
    intHeight = intHeight / 2
    intWidth = (intWidth / 2) - intHeight
    
    'calculate the starting point of the ellipse
    'sngTiltAngle = sngTiltAngle Mod 360
    'sngTiltX = Sin(sngTiltAngle * PI / 180) * intWidth
    'sngTiltY = Cos(sngTiltAngle * PI / 180) * intWidth
    sngTiltX = Sin(sngTiltAngle * PIdiv180) * intWidth
    sngTiltY = Cos(sngTiltAngle * PIdiv180) * intWidth
    
    'draw the ellipse using one line for every three pixels
    'This will increase drawing speed on small ellipses and produce
    'detailed ones for large ellipses.
    sngNumOfPoints = (2 * PI * (intWidth + intHeight)) / 4 '2.Pi.r = circumfrence, /4=per 4 pixels
    
    'size the array to match the number of points to be calculated
    ReDim udtEllipse(sngNumOfPoints)
    
    'calculate the number of degrees between points
    sngDegPerPoint = (360 / sngNumOfPoints)
    
    For sngCounter = 0 To 360 Step sngDegPerPoint
        'sngMoveCenterX = intCenterX + (Cos(sngCounter * PI / 180) * sngTiltX)
        'sngMoveCenterY = intCenterY + (Cos(sngCounter * PI / 180) * sngTiltY)
        sngMoveCenterX = intCenterX + (Cos(sngCounter * PIdiv180) * sngTiltX)
        sngMoveCenterY = intCenterY + (Cos(sngCounter * PIdiv180) * sngTiltY)
        
        'calculate the new position
        'sngCircleX = Sin((sngCounter + sngTiltAngle) * PI / 180) * intHeight
        'sngCircleY = Cos((sngCounter + sngTiltAngle) * PI / 180) * intHeight
        sngCircleX = Sin((sngCounter + sngTiltAngle) * PIdiv180) * intHeight
        sngCircleY = Cos((sngCounter + sngTiltAngle) * PIdiv180) * intHeight
        
        'add the points
        udtEllipse(ROUND(sngCounter / sngDegPerPoint)).X = sngMoveCenterX + sngCircleX
        udtEllipse(ROUND(sngCounter / sngDegPerPoint)).Y = sngMoveCenterY + sngCircleY
    Next sngCounter
    
    'draw the ellipse as a polygon
    
    'first create a brush and pen to display the colours
    udtBrushStuff.lbColor = lngColour
    If blnIsHollow Then
        udtBrushStuff.lbHatch = BS_HOLLOW
    Else
        udtBrushStuff.lbHatch = BS_SOLID
    End If
    udtBrushStuff.lbStyle = 0
    lngOldPtr = CreateBrushIndirect(udtBrushStuff)
    
    'apply brush
    lngResult = SelectObject(mudtBitmap.lngDc, lnghBrush)

    'create the pen
    udtPenStuff.lopnColor = lngColour
    udtPenStuff.lopnWidth.X = intThickness
    udtPenStuff.lopnStyle = PS_SOLID And PS_INSIDEFRAME
    lnghPen = CreatePenIndirect(udtPenStuff)

    'apply pen
    lngOldPtr2 = SelectObject(mudtBitmap.lngDc, lnghPen)
    
    'now draw the ellipse onto the mudtbitmap.hdcmemory
    If Not blnIsHollow Then
        lngResult = Polygon(mudtBitmap.lngDc, udtEllipse(0), sngNumOfPoints)
    Else
        'add a point to complete the circle
        'sngNumOfPoints = sngNumOfPoints + 1
        ReDim Preserve udtEllipse(UBound(udtEllipse) + 1)
        udtEllipse(UBound(udtEllipse)) = udtEllipse(0) 'udtEllipse(UBound(udtEllipse) - 1)
        
        'lngResult = Polyline(mudtBitmap.lngDc, _
                             udtEllipse(0), _
                             Int(sngNumOfPoints) + 1)   'UBound(udtEllipse))
        lngResult = Polygon(mudtBitmap.lngDc, _
                            udtEllipse(0), _
                            Int(sngNumOfPoints) + 1)
    End If
    
    'delete the brush and pen objects from memory
    lngJunk = SelectObject(mudtBitmap.lngDc, lngOldPtr)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = DeleteObject(lnghBrush)
    lngJunk = SelectObject(mudtBitmap.lngDc, lngOldPtr2)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = DeleteObject(lnghPen)
End Sub

Public Function GetAngle(ByVal intX1 As Integer, _
                         ByVal intY1 As Integer, _
                         ByVal intX2 As Integer, _
                         ByVal intY2 As Integer) _
                         As Integer
                         
    'returns the angle of point1 in relation to point2
    
    Dim sngTempAngle As Single
    
    'if the values are not over the center point, then calculate the angle
    If ((intY1 - intY2) <> 0) Or ((intX1 - intX2) <> 0) Then
        
        'the following formula will return an angle from -180 to +180
        sngTempAngle = (Atn(Slope(intX1, intY1, intX2, intY2)) * 180 / PI)
        
        'is the angle past 180 degrees
        If (intY2 - intY1) < 0 Then
            sngTempAngle = 180 + sngTempAngle
        End If
        
        'should the angle be between 180-360 degrees
        If sngTempAngle < 0 Then
            'subtract this angle from 360
            sngTempAngle = 360 + sngTempAngle
        End If
        
        If (sngTempAngle <> 0) Then
            If (intY2 - intY1) < 0 Then
                If (intX2 - intX1) < 0 Then
                    sngTempAngle = sngTempAngle Mod 360
                Else
                    sngTempAngle = (sngTempAngle + 180) Mod 360
                End If
            Else
                If (intX2 - intX1) < 0 Then
                    sngTempAngle = (sngTempAngle + 180) Mod 360
                Else
                    sngTempAngle = sngTempAngle Mod 360
                End If
            End If
        End If
            
        'there are four points where the angle will be "0", at
        '0, 90, 180 and 270 degrees. Adjust appropiatly
        If ((intY2 - intY1) > 0) And ((intX2 - intX1) = 0) Then
            'top
            sngTempAngle = 90
        ElseIf ((intY2 - intY1) < 0) And ((intX2 - intX1) = 0) Then
            'bottom
            sngTempAngle = 270
        ElseIf ((intY2 - intY1) = 0) And ((intX2 - intX1) < 0) Then
            'right
            sngTempAngle = 180
        ElseIf ((intY2 - intY1) = 0) And ((intX2 - intX1) > 0) Then
            'left
            sngTempAngle = 0
        End If
        
        'return the angle
        sngTempAngle = sngTempAngle Mod 360
        GetAngle = CInt(sngTempAngle)
    End If
End Function

Public Function Slope(ByVal intX1 As Integer, _
                      ByVal intY1 As Integer, _
                      ByVal intX2 As Integer, _
                      ByVal intY2 As Integer) _
                      As Single
                      
    'This function finds the slope of a line, where the slope m is,
    '     intX2 - intX1
    'm = ---------------
    '     intY2 - intY1
    
    Dim intXVal As Integer      'holds the horizontal distance
    Dim intYVal As Integer      'holds the vertical distance
    
    intXVal = intX2 - intX1
    intYVal = intY2 - intY1
    If (intXVal = 0) And (intYVal = 0) Then
        'if both values were zero, then
        Slope = 0
        Exit Function
    Else
        'if only one value was zero then
        If (intXVal = 0) Or (intYVal = 0) Then
            'the slope = the other value
            Select Case 0
            Case intXVal
                Slope = intXVal
            Case intYVal
                Slope = intYVal
            End Select
            Exit Function
        End If
    End If
    
    If (intXVal <> 0) And (intYVal <> 0) Then
        'otherwise the slope = the formula
        Slope = (intY2 - intY1) / (intX2 - intX1)
    End If
End Function

Private Function ApplyFont(ByVal fntFormat As StdFont, _
                           ByRef lngOldFont As Long, _
                           Optional ByVal lngApplyhDc As Long = 0, _
                           Optional ByVal lngColour As Long = vbBlack, _
                           Optional ByVal enmAlignment As AlignText = vbCentreAlign, _
                           Optional ByVal intAngle As Integer = 0) _
                           As Long
    'This will apply the specified font to the bitmap and return
    'a handle to the api font object created.
    'NOTE: Both the handle of the font returned AND the handle
    '      of the previous font MUST BE DELETED in the calling
    '      procedure (see DrawString and GetFontPointSize)
    
    Dim udtAPIFont      As LogFont      'contains the font information
    Dim udtFont         As FontStruc    'holds some font information
    Dim intCounter      As Integer      'used to convert a string into a character array
    Dim hDcOldFont      As Long         'holds a handle to the old font
    Dim hDcFont         As Long         'holds a handle to the font
    
    'put information into the definted font structure
    With udtFont
        .Alignment = enmAlignment
        .Name = fntFormat.Name
        .Bold = fntFormat.Bold
        .Italic = fntFormat.Italic
        .Underline = fntFormat.Underline
        .StrikeThru = fntFormat.Strikethrough
        .PointSize = fntFormat.Size
        .Colour = lngColour
    End With
    
    'what device context is the font applied to
    If lngApplyhDc = 0 Then
        lngApplyhDc = mudtBitmap.lngDc
    End If
    
    'apply font details
    udtAPIFont.lfHeight = -((udtFont.PointSize * GetDeviceCaps(mudtBitmap.lngDc, LOGPIXELSY)) / 72)
    udtAPIFont.lfCharSet = DEFAULT_CHARSET
    udtAPIFont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    udtAPIFont.lfEscapement = (intAngle Mod 360) * 10
    
    'move the name of the udtFont into the array
    For intCounter = 1 To Len(udtFont.Name)
        udtAPIFont.lfFaceName(intCounter) = Asc(Mid(udtFont.Name, intCounter, 1))
    Next intCounter
    'this has to be a Null terminated string
    udtAPIFont.lfFaceName(intCounter) = 0
    
    udtAPIFont.lfItalic = udtFont.Italic
    udtAPIFont.lfUnderline = udtFont.Underline
    udtAPIFont.lfStrikeOut = udtFont.StrikeThru
    udtAPIFont.lfOrientation = 0
    udtAPIFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    udtAPIFont.lfPitchAndFamily = DEFAULT_PITCH
    udtAPIFont.lfQuality = PROOF_QUALITY
    
    If udtFont.Bold Then
        udtAPIFont.lfWeight = FW_BOLD
    Else
        udtAPIFont.lfWeight = FW_NORMAL
    End If
    
    udtAPIFont.lfWidth = 0
    hDcFont = CreateFontIndirect(udtAPIFont)
    lngOldFont = SelectObject(lngApplyhDc, hDcFont)
    
    'return a handle to the font. This object must be deleted later
    ApplyFont = hDcFont
End Function

Public Sub DrawString(ByVal strText As String, _
                      Optional ByVal intTop As Integer, _
                      Optional ByVal intLeft As Integer, _
                      Optional ByVal intHeight As Integer, _
                      Optional ByVal intWidth As Integer, _
                      Optional ByVal fntFormat As StdFont, _
                      Optional ByVal lngColour As Long = vbBlack, _
                      Optional ByVal enmAlignment As AlignText = vbCentreAlign, _
                      Optional ByVal intAngle As Integer = 0, _
                      Optional ByVal lngTohDc As Long = 0, _
                      Optional ByVal udtMeasurement As Scaling = InPixels)
                    
    'This procedure will draw strText onto the bitmap in the specified udtFont,
    'colour and position.
    
    Dim lngAlignment    As Long         'holds the text alignment
    Dim udtTextRect     As Rect         'holds the area to draw the text in
    Dim lngResult       As Long         'holds any information returned from the api calls
    Dim lngJunk         As Long         'used to temperorily hold useless information
    Dim hDcFont         As Long         'holds a handle to the font
    Dim hDcOldFont      As Long         'holds a handle to the previous font
    Dim lngOldPtr       As Long         'holds the old pointer. This is used to unselect an object before deleting it
    
    'set Measurement values
    udtTextRect.Top = intTop
    udtTextRect.Left = intLeft
    udtTextRect.Right = intLeft + intWidth
    udtTextRect.Bottom = intTop + intHeight
    
    If udtMeasurement = InTwips Then
        'convert to pixels
        Call RectToPixels(udtTextRect)
    End If
    
    'set the default values
    If (intTop = 0) Then
        udtTextRect.Top = mudtBitmap.Area.Top
    End If
    If (intLeft = 0) Then
        udtTextRect.Left = mudtBitmap.Area.Left
    End If
    If (intWidth <= 0) Then
        udtTextRect.Right = udtTextRect.Left + (mudtBitmap.Area.Right - mudtBitmap.Area.Left)
    End If
    If (intHeight <= 0) Then
        udtTextRect.Bottom = udtTextRect.Top + (mudtBitmap.Area.Bottom - mudtBitmap.Area.Top)
    End If
    
    'where are we drawing this text to
    If (lngTohDc = 0) Then
        'draw to this bitmap
        lngTohDc = mudtBitmap.lngDc
    End If
    
    'apply the font
    hDcFont = ApplyFont(fntFormat, _
                        hDcOldFont, _
                        lngTohDc, _
                        lngColour, _
                        enmAlignment, _
                        intAngle)
    
    Select Case enmAlignment
    Case vbLeftAlign
        lngAlignment = DT_LEFT Or DT_VCENTER
    Case vbCentreAlign
        lngAlignment = DT_CENTER Or DT_VCENTER
    Case vbRightAlign
        lngAlignment = DT_RIGHT Or DT_VCENTER
    End Select
    
    'Draw the strText into the off-screen bitmap before copying the
    'new bitmap (with the strText) onto the screen.
    lngResult = SetBkMode(lngTohDc, TRANSPARENT)
    lngResult = SetTextColor(lngTohDc, lngColour)
    lngResult = DrawText(lngTohDc, _
                         strText, _
                         Len(strText), _
                         udtTextRect, _
                         lngAlignment)
    
    'clean up by deleting the off-screen bitmap and udtFont
    lngJunk = SelectObject(lngTohDc, hDcOldFont)
    lngJunk = DeleteObject(lngJunk)
    lngJunk = SelectObject(lngTohDc, hDcFont)
    lngJunk = DeleteObject(lngJunk)
End Sub

Public Function GetTextHeight(ByVal fntFormat As StdFont) _
                              As Integer
    'This function will return the height of the text using the point size
    
    Dim udtMetrics  As TEXTMETRIC   'holds the font metrics for the bitmap
    Dim lngResult   As Long         'holds any returned error value from an api call
    Dim hFont       As Long         'holds a pointer to the font
    Dim hOldFont    As Long         'holds a pointer to the previous font
    Dim udtTemp     As BitmapStruc  'holds a temperory bitmap to apply the font to
    
    'create a one pixel bitmap
    With udtTemp
        .Area.Bottom = 1
        .Area.Right = 1
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
    
    'apply the font
    hFont = ApplyFont(fntFormat, hOldFont, udtTemp.lngDc)
    
    lngResult = GetTextMetrics(udtTemp.lngDc, _
                               udtMetrics)
    
    GetTextHeight = udtMetrics.tmHeight
    
    'delete the font objects
    lngResult = SelectObject(udtTemp.lngDc, hOldFont)
    lngResult = DeleteObject(lngResult)
    lngResult = SelectObject(udtTemp.lngDc, hFont)
    lngResult = DeleteObject(lngResult)
    
    'delete the bitmap
    With udtTemp
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
    End With
End Function

Public Sub GetScreenRes(ByRef intWidth As Integer, _
                        ByRef intHeight As Integer)
    'This procedure sets the variable to the current screen dimensions.
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX
    intHeight = Screen.Height / Screen.TwipsPerPixelY
End Sub

Public Sub ReturnOldDisplay()
    'returns the display to what it was originally
    
    If mblnResChanged Then
        Call ChangeDisplaySettings(Null, 0)
        mblnResChanged = False
    End If
End Sub

Public Sub SetDisplay(ByVal intWidth As Integer, _
                      ByVal intHeight As Integer)
    'changes the resolution of the screen to new size
    
    Dim udtDevM     As DEVMODE  'holds the current device settings for the system
    Dim lngResult   As Long     'holds any returned error value from an api call
    
    lngResult = EnumDisplaySettings(0, 0, udtDevM)
    
    With udtDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = intWidth  'ScreenWidth
            .dmPelsHeight = intHeight 'ScreenHeight
            .dmBitsPerPel = 32 '(could be 8, 16, 32 or even 4)
    End With
    lngResult = ChangeDisplaySettings(udtDevM, 2)
    
    If lngResult = 0 Then
            Call ChangeDisplaySettings(udtDevM, 4)
    End If
    
    mblnResChanged = True
End Sub

Private Function IsWinNT() As Boolean
    'Detect if the program is running under Windows NT. Different
    'graphics api calls may be necessary under NT based systems
    
    Const VER_PLATFORM_WIN32_NT     As Integer = 2
    
    Dim myOS                        As OSVERSIONINFO
    Dim lngResult                   As Long             'returned error value from the api call
    
    'get version information
    myOS.dwOSVersionInfoSize = Len(myOS)
    lngResult = GetVersionEx(myOS)
    
    'return True if the test of windows NT is positive
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Sub LoadBitmap(ByVal strFileName As String, _
                      Optional ByVal enmImageType As LoadType = IMAGE_BITMAP, _
                      Optional ByVal enmMeasurement As Scaling = InPixels)
    'This procedure will load the specified image into a bitmap.
    'Please keep in mind that this function creates a bitmap object. If the
    'handles passed to this procedure already contain a bitmap, then it is
    'automatically deleted. This is assuming that the load operation was
    'successful. If not, then the pointers are left untouched.
    '
    'NOTE: is not currently supported by WinNT based systems
    
    Dim hBitmap         As Long     'holds the pointer to the bitmap itself
    Dim hMemory         As Long     'holds the pointer to the DC
    Dim hPointer        As Long     'holds a pointer to the previously assigned object
    Dim lngResult       As Long     'holds any returned error value of api calls
    Dim udtLoadArea     As Rect     'holds the size of the bitmap
    Dim udtDimensions   As BITMAP   'holds the bitmap dimensions and settings
    
    'make sure we load the correct image type
    Select Case LCase(Right(strFileName, 3))
    Case "bmp"  'file is a bitmap
        enmImageType = IMAGE_BITMAP
    
    Case "ico"  'file is an icon
        enmImageType = IMAGE_ICON
    
    Case "cur"  'file is a cursor
        enmImageType = IMAGE_CURSOR
        
    Case "jpg"
        Set Me.Picture = LoadPicture(strFileName)
        Exit Sub
    End Select
    
    'we have to pass a Null terminated string
    strFileName = strFileName & Chr(0)
    hBitmap = LoadImage(0, _
                        strFileName, _
                        enmImageType, _
                        0, _
                        0, _
                        LR_LOADFROMFILE)
    
    'if the load operation was unsuccessful then exit the procedure.
    If hBitmap = Null Then
        'load unsuccessful
        Exit Sub
    End If
    
    'create a device context and assign the bitmap to it
    hMemory = CreateCompatibleDC(mudtBitmap.lngDc)
    hPointer = SelectObject(hMemory, hBitmap)
    
    'get the size of the bitmap loaded
    lngResult = GetObjectAPI(hBitmap, Len(udtDimensions), udtDimensions)
    
    'assign the udtDimensions to a rect structure. If no values
    'were returned, then use the classes dimensions
    udtLoadArea.Left = 0
    udtLoadArea.Top = 0
    If lngResult <> 0 Then
        'everything ok, get the size of the loaded bitmap
        udtLoadArea.Right = udtDimensions.bmWidth
        udtLoadArea.Bottom = udtDimensions.bmHeight
    Else
        'default size
        udtLoadArea.Right = 600
        udtLoadArea.Bottom = 480
    End If
    
    If enmMeasurement = InTwips Then
        'convert pixel values to twips
        Call RectToTwips(udtLoadArea)
    End If
    
    With mudtBitmap
        'Now that a new bitmap has been created, then we have to check to
        'see if the passed parameters are already in use hlding a bitmap. If so
        'then we need to delete it from memory before replacing it.
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
        
        'return the handles of the bitmap created
        .lngDc = hMemory
        .lngBitmap = hBitmap
        .lngPointer = hPointer
        .Area = udtLoadArea
    End With
End Sub

Public Sub SizeToBitmap(Optional ByVal hBitmap As Long = 0, _
                        Optional blnStretch As Boolean = False)
    'This will automatically resize the bitmap to the size retrieved from
    'the specified bitmap handle. Should no hBitmap be provided or a zero
    'parameter be passed, the bitmap resizes itself to its DC size. This will
    'reset the Left and Top properties of the bitmap to 0. Should the
    'procedure be unable to retrieve any new size, the original size is
    'unchanged.
    
    Dim udtDimensions   As BITMAP   'holds the size of the bitmap
    Dim lngResult       As Long     'holds any returned error value from an api call
    
    'get the dimensions of the bitmap
    lngResult = GetObjectAPI(hBitmap, Len(udtDimensions), udtDimensions)
    If lngResult = 0 Then
        Exit Sub
    End If
    
    'update the new size
    With udtDimensions
        Call ReSize(.bmWidth, .bmHeight, blnStretch)
    End With
End Sub

Public Sub ReSize(Optional ByVal intWidth As Integer = -1, _
                  Optional ByVal intHeight As Integer = -1, _
                  Optional ByVal blnStretch As Boolean = False)
    'This will change the size of the bitmap to the specified dimensions.
    'The bitmap can be stretched into it's new size. If any negative value
    'is passed, then the bitmaps original dimensions are used.
    
    Dim udtTempBitmap   As BitmapStruc  'holds a temperory bitmap
    Dim lngResult       As Long         'holds any returned error value
    
    'get the new dimensions
    With mudtBitmap.Area
        If intWidth < 0 Then
            'use original width
            intWidth = .Right - .Left
        End If
        If intHeight < 0 Then
            'use original height
            intHeight = .Bottom - .Top
        End If
    End With
    
    'create a bitmap of the size
    With udtTempBitmap
        .Area.Right = intWidth
        .Area.Bottom = intHeight
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area)
    End With
    
    'copy any existing picture from the old to the new
    With mudtBitmap
        If blnStretch Then
            'stretch the old picture
            lngResult = StretchBlt(udtTempBitmap.lngDc, _
                                   0, _
                                   0, _
                                   intWidth, _
                                   intHeight, _
                                   .lngDc, _
                                   .Area.Left, _
                                   .Area.Top, _
                                   .Area.Right - .Area.Left, _
                                   .Area.Bottom - .Area.Top, _
                                   vbSrcCopy)
        Else
            'just copy the old picture. Any extra space will be black
            lngResult = BitBlt(udtTempBitmap.lngDc, _
                               0, _
                               0, _
                               intWidth, _
                               intHeight, _
                               .lngDc, _
                               .Area.Left, _
                               .Area.Top, _
                               vbSrcCopy)
        End If
        
        'adjust the dimensions so that they are in relation to the
        'old bitmaps Top and Left properties
        With .Area
            udtTempBitmap.Area.Left = .Left
            udtTempBitmap.Area.Top = .Top
            udtTempBitmap.Area.Right = .Left + intWidth
            udtTempBitmap.Area.Bottom = .Top + intHeight
        End With
        
        'destroy the old bitmap
        Call DeleteBitmap(.lngDc, .lngBitmap, .lngPointer)
    End With    'mudtBitmap
    
    'change the pointers to the new bitmap
    mudtBitmap = udtTempBitmap
End Sub

Private Function GetWorkArea() As Rect
    'Get the area the user is working with (minus the
    'task bar) in PIXELS
    
    Dim lngResult As Long
    Dim udtWorkArea As Rect             'holds information about the current window
    
    lngResult = SystemParametersInfo(SPI_GETWORKAREA, _
                                     0, _
                                     udtWorkArea, _
                                     0)
    GetWorkArea = udtWorkArea
End Function

Public Function PixelColourCount(Optional ByVal lngStartCol As Long = vbBlack, _
                                 Optional ByVal lngFinishCol As Long = NO_COLOUR) _
                                 As Long
    'This will count all pixels that are within the specified colour range.
    'If no lngFinishCol is specified, only pixels of lngStartCol are counted
    
    Dim intStartRed     As Integer      'holds the red value of the start colour
    Dim intStartGreen   As Integer      'holds the green value of the start colour
    Dim intStartBlue    As Integer      'holds the blue value of the start colour
    Dim intFinRed       As Integer      'holds the red value of the finishing colour
    Dim intFinGreen     As Integer      'holds the green value of the finishing colour
    Dim intFinBlue      As Integer      'holds the blue value of the finishing colour
    Dim intPxlRed       As Integer      'holds the red value of the pixel to check
    Dim intPxlGreen     As Integer      'holds the green value of the pixel to check
    Dim intPxlBlue      As Integer      'holds the blue value of the pixel to check
    Dim intRow          As Integer      'used to cycle through the rows of the bitmap
    Dim intCol          As Integer      'used to cycle through the columns of the bitmap
    Dim lngCheckCol     As Long         'holds the colour of the pixel to check
    Dim lngResult       As Long         'holds any returned error value from an api call
    Dim lngCount        As Long         'holds the count of any matchin pixel
    
    'make sure that a bitmap was created
    If mudtBitmap.lngDc = 0 Then
        Exit Function
    End If
    
    'validate the colours passed
    Call GetRGB(lngStartCol, intStartRed, intStartGreen, intStartBlue)
    If (lngFinishCol = NO_COLOUR) Then
        lngFinishCol = lngStartCol
    End If
    Call GetRGB(lngFinishCol, intFinRed, intFinGreen, intFinBlue)
    
    'search through the bitmap
    lngCount = 0
    With mudtBitmap.Area
        For intRow = .Left To .Right
            For intCol = .Top To .Bottom
                'get the colour of the specified pixel
                lngCheckCol = GetPixel(mudtBitmap.lngDc, intRow, intCol)
                Call GetRGB(lngCheckCol, intPxlRed, intPxlGreen, intPxlBlue)
                
                'see if the colour is within the specified range
                If ((intPxlRed >= intStartRed) And _
                    (intPxlRed <= intFinRed)) And _
                   ((intPxlGreen >= intStartGreen) And _
                    (intPxlGreen <= intFinGreen)) And _
                   ((intPxlBlue >= intStartBlue) And _
                    (intPxlBlue <= intFinBlue)) Then
                    'we have found a valid pixel
                    lngCount = lngCount + 1
                End If  'colour is within range
            Next intCol
        Next intRow
    End With    'mudtBitmap.Area
    
    'return the pixel count
    PixelColourCount = lngCount
End Function

Public Sub PutWindowInFore(ByVal hWnd As Long)
    'This will move the specified window to the top of the zorder
    
    Dim lngResult       As Long     'holds any returned value from an api call
    
    lngResult = SetForegroundWindow(hWnd)
End Sub

Private Function RectIntersects(ByRef udtRect1 As Rect, _
                                ByRef udtRect2 As Rect) _
                                As Boolean
    'This function will return True if the two passed Rect structures
    'intersect each other.
    
    Dim udtTempRect As Rect 'holds the co-ordinates of the intersection rectangle
    Dim lngResult   As Long 'holds any returned error value from an api call
    
    lngResult = IntersectRect(udtTempRect, udtRect1, udtRect2)
    RectIntersects = lngResult
End Function

Public Function MouseKeyPressed(ByVal udtKey As MouseButtonConstants) _
                                As Boolean
    'This will return True if the specified udtKey was pressed
    
    Const KeyPressed    As Integer = -32768
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngMyKey        As Long     'holds the code for the mouse key to check
    
    lngMyKey = udtKey
    
    'find the udtKey state of the mouse udtKey specified
    lngResult = GetAsyncKeyState(lngMyKey)
    
    'if the udtKey was pressed, then
    If lngResult = KeyPressed Then
        'the udtKey is pressed
        MouseKeyPressed = True
    End If
End Function

Public Function GethWnd(ByVal enmWindow As EnumWindowHandles) _
                        As Long
    'This will return a handle to the specified window
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim lngWindow       As Long     'holds a handle to the specified window
    
    'get a handle to the specified window
    Select Case enmWindow
    Case WND_TOP_WINDOW
        lngWindow = GetForegroundWindow
        
    Case WND_DESKTOP_WINDOW
        lngWindow = GetDesktopWindow
    
    Case WND_ACTIVE_WINDOW
        lngWindow = GetActiveWindow
    End Select
    
    'return the handle to the specified window
    GethWnd = lngWindow
End Function

Public Sub GetScreenShot(Optional ByVal intScreenX As Integer = 0, _
                         Optional ByVal intScreenY As Integer = 0, _
                         Optional ByVal intDestWidth As Integer = -1, _
                         Optional ByVal intDestHeight As Integer = -1, _
                         Optional ByVal intDestX As Integer = 0, _
                         Optional ByVal intDestY As Integer = 0, _
                         Optional ByVal udtMeasurement As Scaling = InPixels)
                         
    'This will get a screen shot at the specified co-ordinates and copy
    'them into the specified destination co-ordinates.
    
    Dim lngResult   As Long     'holds any returned error value from an api call
    Dim lnghScreen  As Long     'holds a handle to the screen DC
    
    'set the default if necessary
    With mudtBitmap.Area
        If intDestWidth < 0 Then
            intDestWidth = .Right - .Left
        End If
        If intDestHeight < 0 Then
            intDestHeight = .Bottom - .Top
        End If
    End With
    
    'set the scaling mode specified and convert parameters if necessary
    If udtMeasurement = InTwips Then
        intDestX = intDestX / Screen.TwipsPerPixelX
        intDestY = intDestY / Screen.TwipsPerPixelY
        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
        intScreenX = intScreenX / Screen.TwipsPerPixelX
        intScreenY = intScreenY / Screen.TwipsPerPixelY
    End If
    
    'get a handle to the screen
    lnghScreen = GetDC(0)
    
    'copy the screen shot - GetDesktopWindow was previously
    'used to get the handle on the top window.
    lngResult = BitBlt(mudtBitmap.lngDc, _
                       intDestX, _
                       intDestY, _
                       intDestWidth, _
                       intDestHeight, _
                       GetDC(0), _
                       intScreenX, _
                       intScreenY, _
                       P_COPY)
End Sub

Public Sub GetWallpaper(Optional ByVal intScreenX As Integer = 0, _
                        Optional ByVal intScreenY As Integer = 0, _
                        Optional ByVal intDestWidth As Integer = -1, _
                        Optional ByVal intDestHeight As Integer = -1, _
                        Optional ByVal intDestX As Integer = 0, _
                        Optional ByVal intDestY As Integer = 0, _
                        Optional ByVal udtMeasurement As Scaling = InPixels)
                         
    'This will get a screen shot at the specified co-ordinates and copy
    'them into the specified destination co-ordinates.
    
    Dim lngResult   As Long         'holds any returned error value from an api call
    Dim lnghScreen  As Long         'holds a handle to the screen DC
    Dim udtTemp     As BitmapStruc  'temporarily holds the entire desktop wallpaper/pattern
    Dim udtOldPos   As PointAPI     'used to move the current pointer back to 0, 0
    Dim lnghRect    As Long         'holds a pointer to a rect region for setting the clipping region in the bitmap
    Dim lngOldPtr   As Long         'holds the previous pointer to be applied to the bitmap
    
    'set the default if necessary
    If intDestWidth < 0 Then
        intDestWidth = mudtBitmap.Area.Right - _
                       mudtBitmap.Area.Left
    End If
    If intDestHeight < 0 Then
        intDestHeight = mudtBitmap.Area.Bottom - _
                        mudtBitmap.Area.Top
    End If
    
    'set the scaling mode specified and convert parameters if necessary
    If udtMeasurement = InTwips Then
        intDestX = intDestX / Screen.TwipsPerPixelX
        intDestY = intDestY / Screen.TwipsPerPixelY
        intDestWidth = intDestWidth / Screen.TwipsPerPixelX
        intDestHeight = intDestHeight / Screen.TwipsPerPixelY
        intScreenX = intScreenX / Screen.TwipsPerPixelX
        intScreenY = intScreenY / Screen.TwipsPerPixelY
    End If
    
    'create the temporary bitmap
    With udtTemp
        'set the size to the work area
        .Area = GetWorkArea
'
'        Call CreateNewBitmap(.lngDc, _
'                             .lngBitmap, _
'                             .lngPointer, _
'                             .Area, _
'                             frmClock.picTest.hdc) 'mlngCompatableDc)
        
        'copy the wallpaper - GetDesktopWindow was previously
        'used to get the handle on the top window.
        'lngResult = PaintDesktop(.lngDc)
        'lngResult = PaintDesktop(frmClock.picTest.hDc)
        'With mudtBitmap
        '    udtOldPos.X = .Area.Right
        '    udtOldPos.Y = .Area.Bottom
        '    lngResult = MoveToEx(.lngDc, _
        '                         .Area.Right, _
        '                         .Area.Bottom, _
        '                         udtOldPos)
        'End With
        'lngResult = SetBrushOrgEx(mudtBitmap.lngDc, 0, 0, udtOldPos)
        'lnghRect = CreateRectRgnIndirect(.Area)
        'lngResult = ExtSelectClipRgn(.lngDc, lnghRect, RGN_COPY)
        
        Call CreateNewBitmap(.lngDc, .lngBitmap, .lngPointer, .Area)
        lngResult = PaintDesktop(.lngDc)
        
        'lngResult = PaintDesktop(frmClock.picTest.hDc)
        'Debug.Print lngResult
        'lngResult = DeleteObject(lnghRect)
        'lngResult = PaintDesktop(mudtBitmap.lngDc)
        'lngResult = PaintDesktop(frmClock.picTest.hdc)
        'Debug.Print lngResult
        
        'Dim picTest     As StdPicture
        'Dim lngScreenDc As Long             'holds a handle to the screen
        '
        ''create a picture object and assign a Dc to it
        'Set picTest = New StdPicture
        '
        ''picTest.Width = (Screen.Width / Screen.TwipsPerPixelX)
        ''picTest.Height = (Screen.Height / Screen.TwipsPerPixelY)
        '
        'lngScreenDc = GetDC(0)
        '.lngDc = CreateCompatibleDC(lngScreenDc)
        '.lngPointer = SelectObject(.lngDc, picTest.Handle)
        '
        ''realease the screen Dc
        'lngResult = ReleaseDC(0, lngScreenDc)
        '
        'lngResult = PaintDesktop(.lngDc)
        
        'copy the specified section of the wallpaper
        'lngResult = BitBlt(frmClock.picTest.hdc, 0, 0, intDestWidth, intDestHeight, .lngDc, 100, 100, vbSrcCopy)
        lngResult = BitBlt(mudtBitmap.lngDc, _
                           intDestX, _
                           intDestY, _
                           intDestWidth, _
                           intDestHeight, _
                           .lngDc, _
                           intScreenX, _
                           intScreenY, _
                           vbSrcCopy)
'
'        'destroy the bitmap
'        lngResult = SelectObject(.lngDc, .lngPointer)
'        Set picTest = Nothing
'        lngResult = DeleteDC(.lngDc)
        
        
        'delete the bitmap
        Call DeleteBitmap(.lngDc, .lngBitmap, .lngPointer)
    End With
End Sub

Public Sub GetWindowSnapShot(ByVal hWnd As Long)
    'This will capture the entire contents of a window and
    'resize the bitmap accordingly
    
    Dim lngResult       As Long         'holds any returned value from an api call
    Dim lngTempDc       As Long         'temperorily holds a DC pointer to the window bitmap
    Dim udtWindowArea   As Rect         'holds the size of the widnow
    
    'make sure that a widnow was passed
    If hWnd = 0 Then
        Exit Sub
    End If
    
    'get the dc of the snapshot
    lngTempDc = GetWindowDC(hWnd)
    
    'were we successful in getting the snapshot
    If lngTempDc = 0 Then
        Exit Sub
    End If
    
    'get the dimensions of the DC
    lngResult = GetWindowRect(hWnd, udtWindowArea)
    
    'did we get the size of the window?
    If lngResult = 0 Then
        Exit Sub
    End If
    
    With mudtBitmap
        'create a new bitmap of the new size
        .Area = udtWindowArea
        With .Area
            Call SetBitmap(.Bottom, .Right)
        End With
    End With
    
    'copy the contents of the window DC
    Call PaintFrom(lngTempDc)
    
    'release the windows DC
    lngResult = ReleaseDC(hWnd, lngTempDc)
End Sub

Public Function RndNumber(ByVal dblMin As Double, _
                          ByVal dblMax As Double, _
                          Optional blnReturnInteger As Boolean = False) _
                          As Double
    'This function will produce a random number
    'between the specified values
    
    Dim dblTemp As Double
    
    'if the two values are equal then
    If dblMin = dblMax Then
        'return same number and exit
        RndNumber = dblMin
        Exit Function
    End If
    
    'if dblMin is bigger than dblMax then swap values
    If dblMin > dblMax Then
        dblTemp = dblMin
        dblMin = dblMax
        dblMax = dblMin
    End If
    
    'generate a random number
    dblTemp = (dblMax - dblMin + 1) * Rnd + dblMin
    
    'should we return an integer result
    If blnReturnInteger Then
        dblTemp = Int(dblTemp)
    End If
    
    'make sure that the number is not outside the ranges specified
    Select Case dblTemp
    Case Is < dblMin
        dblTemp = dblMin
        
    Case Is > dblMax
        dblTemp = dblMax
    End Select
    
    'return the randomly generated number
    RndNumber = dblTemp
End Function

Public Sub SetBmpToMenu(ByVal hWnd As Long, _
                        ByVal lngTopPos As Long, _
                        ByVal lngSubPos1 As Long, _
                        Optional ByVal lngSubPos2 As Long = -1, _
                        Optional ByVal lngSubPos3 As Long = -1, _
                        Optional ByVal lnghDC As Long = 0)
    'This will set the graphic of a menu item
    'to any image in the device context. The menu
    'item must NOT be a top level menu or have a
    'sub menu. Top-level menu positions are from
    'left to right, starting at 0. Sub-level menu
    'positions are from top down, starting at 0.
    'NOTE : mudtBitmap.lngDC MUST be the Picture property
    'of a control
    
    Const MF_BYPOSITION As Long = &H400&
    Const MF_BYCOMMAND  As Long = &H0&
    Const BMP_SIZE      As Long = 14
    Const BF_BITMAP     As Long = &H4
    
    Dim lngResult       As Long     'any error message returned from an api call
    Dim hMenu           As Long     'the handle of the current menu item
    Dim hSubMenu        As Long     'the handle of the current sub menu
    Dim lngID           As Long     'the sub menu's ID
    
    'Get the handle of the form's menu
    hMenu = GetMenu(hWnd)
    
    'Get the handle of the form's submenu
    hSubMenu = GetSubMenu(hMenu, lngTopPos)
    
    'get any sub menu's
    If lngSubPos2 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos2)
    End If
    If lngSubPos3 >= 0 Then
        hSubMenu = GetSubMenu(hSubMenu, lngSubPos3)
    End If
    
    'see if we are to use the bitmap graphic or if
    'another has been specified for us
    If lnghDC = 0 Then
        'use the object bitmap
        lnghDC = mudtBitmap.lngDc
    End If
    
    'if we were unable to get the sub menu handle
    'then exit
    If (hMenu = 0) Or _
       (hSubMenu = 0) Or _
       (lnghDC = 0) Then
        Exit Sub
    End If
    
    'set the graphic to the sub menu
    lngID = GetMenuItemID(hSubMenu, lngSubPos1)
    
    lngResult = SetMenuItemBitmaps(hSubMenu, _
                                   lngID, _
                                   MF_BYCOMMAND, _
                                   lnghDC, _
                                   lnghDC)
End Sub

Public Sub Save(ByVal strFileName As String, _
                Optional ByVal enmType As PictureTypeConstants = vbPicTypeBitmap)
    'This will save the picture in the selected location
    'in the selected type
    
    Dim strExtention        As String       'holds the proper extention for the selected type
    
    'make sure that a bitmap is loaded
    If mudtBitmap.lngDc = 0 Then
        Exit Sub
    End If
    
    'in debug, only bitmaps <-- 20 Decemeber 2002
    enmType = vbPicTypeBitmap
    
    'make sure that the filename has the correct
    'extention for the type
    Select Case enmType
    Case vbPicTypeBitmap
        strExtention = ".bmp"

    Case vbPicTypeEMetafile
        strExtention = ".emf"

    Case vbPicTypeIcon
        strExtention = ".ico"

    Case vbPicTypeMetafile
        strExtention = ".wmf"

    Case Else
        'default to a bitmap
        strExtention = ".bmp"
    End Select
    
    'is the extention already in the string
    If (UCase(Right(strFileName, 4)) <> UCase(strExtention)) Then
        'add the extention to the end of the file
        strFileName = strFileName & strExtention
    End If
    
    'try and save the bitmap to a file
    With mudtBitmap
        Call SavePicture(hDCToPicture(.lngDc, _
                                      .lngBitmap, _
                                      .lngPointer, _
                                      enmType), _
                         strFileName)
    End With
End Sub

Public Sub SetBitmap(ByVal lngWidth As Long, _
                     ByVal lngHeight As Long, _
                     Optional ByVal lngColour As Long = vbButtonFace, _
                     Optional ByVal lngCompatableWithhDc As Long = 0)
    'this will recreate the bitmaps size and colour
    'to the specified dimensions
    
    'set the colour (makeing sure system colours are
    'converted to RGB values)
    mlngBackColour = GetSystemColour(lngColour)
    
    'create the bitmap
    With mudtBitmap
        'delete the old bitmap if it existed
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
        
        'create the bitmap size (0x0 by default)
        With .Area
            .Bottom = lngHeight
            .Right = lngWidth
        End With
        
        'create the bitmap compatable with the
        'Desktop device context (usually displayable
        'on a visual basic form or picture box)
        Call CreateNewBitmap(.lngDc, _
                             .lngBitmap, _
                             .lngPointer, _
                             .Area, _
                             lngCompatableWithhDc)
        mlngCompatableDc = lngCompatableWithhDc
        
        'the type of the picture is by default a bitmap
        menmType = vbPicTypeBitmap
    
        'set the default stretchmode
        If IsWinNT Then
            'Half_tone doesn't always work on an NT based system
            Call SetStretchMode(STRETCH_DELETESCANS)
        Else
            Call SetStretchMode
        End If
    End With
End Sub

Public Sub Paint(ByVal lngDestDc As Long, _
                 Optional ByVal intDestX As Integer = 0, _
                 Optional ByVal intDestY As Integer = 0, _
                 Optional ByVal intHeight As Integer = -1, _
                 Optional ByVal intWidth As Integer = -1, _
                 Optional ByVal intSourceX As Integer = 0, _
                 Optional ByVal intSourceY As Integer = 0, _
                 Optional ByVal lngPaintMode As PaintMode = P_COPY, _
                 Optional ByVal lngTransparentCol As Long = NO_COLOUR)
    'This will copy the bitmap onto the specified
    'device context and position. If the height or
    'width values are -1, then this is taken as the
    'height and width of the bitmap
    
    Dim lngResult As Long   'holds any error value returned from the api calls
    
    'if the height or width values are default, then
    'use the bitmaps size as the amount to copy
    If intHeight < 0 Then
        intHeight = mudtBitmap.Area.Bottom
    End If
    If intWidth < 0 Then
        intWidth = mudtBitmap.Area.Right
    End If
    
    'if a transparent colour was specified, then use TransparentBlt (98+ only)
    If (lngTransparentCol = NO_COLOUR) Then
        'copy the bitmap onto the device context as normal
        lngResult = BitBlt(lngDestDc, _
                           intDestX, _
                           intDestY, _
                           intWidth, _
                           intHeight, _
                           mudtBitmap.lngDc, _
                           intSourceX, _
                           intSourceY, _
                           lngPaintMode)
    Else
        'copy using a single specified colour as "transparent"
        lngResult = TransparentBlt(lngDestDc, _
                                   intDestX, _
                                   intDestY, _
                                   intWidth, _
                                   intHeight, _
                                   mudtBitmap.lngDc, _
                                   intSourceX, _
                                   intSourceY, _
                                   intWidth, _
                                   intHeight, _
                                   GetSystemColour(lngTransparentCol))
    End If
End Sub

Public Sub Cls()
    'clears the bitmap by drawing a rectangle covering
    'the entire bitmap with the background colour
    Call DrawRect
End Sub

Public Sub PaintFrom(ByVal lngSourceDC As Long, _
                     Optional ByVal intSrcWidth As Integer = -1, _
                     Optional ByVal intSrcHeight As Integer = -1, _
                     Optional ByVal intSrcLeft As Integer = 0, _
                     Optional ByVal intSrcTop As Integer = 0, _
                     Optional ByVal intDestLeft As Integer = 0, _
                     Optional ByVal intDestTop As Integer = 0, _
                     Optional ByVal intDestWidth As Integer = -1, _
                     Optional ByVal intDestHeight As Integer = -1, _
                     Optional ByVal enmPaintMode As PaintMode = P_COPY, _
                     Optional ByVal lngTransparentCol As Long = NO_COLOUR)
    'This will copy from a source device context onto
    'the bitmap at the appropiate position
    
    Dim lngResult As Long   'holds any returned error value from the api calls
    
    'validate the source device context
    If lngSourceDC = 0 Then
        Exit Sub
    End If
    
    'validate the measurements
    Select Case 0
    Case intSrcWidth, intSrcHeight, _
         intDestWidth, intDestHeight
        'nothing will be displayed, exit
        Exit Sub
    End Select
    
    'if the destination measurements are set to the
    'default values, then assume the destination
    'co-rodinates are to be the size of the bitmap
    If (intDestWidth < 0) Or (intDestHeight < 0) Then
        'set the default sizes as the size of the
        'bitmap
        intDestWidth = mudtBitmap.Area.Right - mudtBitmap.Area.Left
        intDestHeight = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
    End If
    
    If (intSrcWidth < 0) Or (intSrcHeight < 0) Then
        'set the default sizes as the size of the
        'bitmap
        intSrcWidth = mudtBitmap.Area.Right - mudtBitmap.Area.Left
        intSrcHeight = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
    End If
    
    'copy the bitmap using fastest api method possible
    If (lngTransparentCol <> NO_COLOUR) Then
        'copy using a transparent colour
        lngResult = TransparentBlt(mudtBitmap.lngDc, _
                                   intDestLeft, _
                                   intDestTop, _
                                   intDestWidth, _
                                   intDestHeight, _
                                   lngSourceDC, _
                                   intSrcLeft, _
                                   intSrcTop, _
                                   intSrcWidth, _
                                   intSrcHeight, _
                                   GetSystemColour(lngTransparentCol))
    
    ElseIf (intSrcWidth = intDestWidth) And _
           (intSrcHeight = intDestHeight) Then
    
        'use BitBlt to copy
        lngResult = BitBlt(mudtBitmap.lngDc, _
                           intDestLeft, _
                           intDestTop, _
                           intSrcWidth, _
                           intSrcHeight, _
                           lngSourceDC, _
                           intSrcLeft, _
                           intSrcTop, _
                           enmPaintMode)
    Else
        'use StretchBlt to copy
        lngResult = StretchBlt(mudtBitmap.lngDc, _
                               intDestLeft, _
                               intDestTop, _
                               intDestWidth, _
                               intDestHeight, _
                               lngSourceDC, _
                               intSrcLeft, _
                               intSrcTop, _
                               intSrcWidth, _
                               intSrcHeight, _
                               enmPaintMode)
    End If
End Sub

Public Sub SetStretchMode(Optional ByVal lngMode As EnumStretchModes = HALFTONE)
    'This will set the stretch mode to HALF_TONE (4). This is
    'important when resizing images so that colour is
    'preserved properly. Other stretching modes can be
    'specified if colour preservation is not an issue or
    'for working with monochrome bitmaps.
    'STRETCH_HALFTONE = 3
    
    Dim lngResult       As Long             'holds any returned error value from an api call
    Dim udtOldOrigin    As PointAPI         'holds the old brush origin to be re-applied after setting the stretch mode
    
    'make sure that there is an active bitmap or we cannot
    'record the old brush origin
    If mudtBitmap.lngDc = 0 Then
        Exit Sub
    End If
    
    'get the old brush origin
    lngResult = GetBrushOrgEx(mudtBitmap.lngDc, udtOldOrigin)
    
    'make sure the old origin was returned
    If lngResult = 0 Then
        Exit Sub
    End If
    
    'apply the specified stretch mode
    If GetStretchBltMode(mudtBitmap.lngDc) <> HALFTONE Then
        lngResult = SetStretchBltMode(mudtBitmap.lngDc, lngMode)
    End If
    
    'reset the colour adjustment to normal
    Call AdjustColour
    
    're-apply the old brush origin
    lngResult = SetBrushOrgEx(mudtBitmap.lngDc, _
                              udtOldOrigin.X, _
                              udtOldOrigin.Y, _
                              udtOldOrigin)
End Sub

Public Function SupportedFormats() As String
    'Returns a string of all the file formats supported
    'for loading/saving. The string returned is intended
    'for direct use with the common dialog control
    '- currently only bitmaps; 22 August 2002
    
    SupportedFormats = "*.bmp;*.jpg" ';*.ico;*.cur;*.wmf;*.emf;*.gif" '- not supported yet -->";*.ico;*.cur"
End Function

Public Function GetSystemColour(ByVal lngIndex As SystemColorConstants) _
                                As Long
    'Convert a vb tagged system colour into a valid
    'RGB colour value. Please note that there are two
    'index values not represented here but can be
    'retrieved; ActiveTitleBar2 and InactiveTitleBar2.
    'These values can be got by using the constants
    '; vbActiveTitleBar2 and vbInactiveTitleBar2. These
    'values are only valid on systems where the title
    'bar on a window can display a gradient from one
    'colour to another
    
    Const SYS_MASK As Double = 2147483648#   'Binary: 01000000,00000000,00000000,00000000
    
    If (lngIndex < 0) Then
        'the colour passed is a system colour index
        lngIndex = lngIndex + SYS_MASK
        
        'use that index to get the RGB Long colour
        GetSystemColour = GetSysColor(lngIndex)
    Else
        'the colour passed is an ordinary colour, just
        'return
        GetSystemColour = lngIndex
    End If
End Function
                    
Public Sub MousePosition(ByRef intX As Integer, _
                         ByRef intY As Integer)
    'sets the parameters to the current screen position of the mouse
    
    Dim udtMousePos     As PointAPI 'holds the position of the mouse
    Dim lngResult       As Long     'holds any returned error value from an api call
    
    lngResult = GetCursorPos(udtMousePos)
    If lngResult <> 0 Then
        'set the parameters to the mouse position
        With udtMousePos
            intX = .X
            intY = .Y
        End With
    End If
End Sub

'--------------------------------------------------------------------------
'                                CLASS EVENTS
'--------------------------------------------------------------------------

Private Sub Class_Initialize()
    'the defaults are as small as possible to take
    'up less memory until absolutly needed
    Const DEFAULT_HEIGHT = 0
    Const DEFAULT_WIDTH = 0
    
    'set the default values
    Call SetBitmap(DEFAULT_HEIGHT, DEFAULT_WIDTH)
    
    'make sure the appropiate stretch mode is applied
    Call SetStretchMode
End Sub

Private Sub Class_Terminate()
    'remove the bitmap from memory before exiting
    'the class
    
    With mudtBitmap
        Call DeleteBitmap(.lngDc, _
                          .lngBitmap, _
                          .lngPointer)
    End With
End Sub

'--------------------------------------------------------------------------
'                               PROPERTIES
'--------------------------------------------------------------------------

Public Property Get BackgroundColour() As Long
    'returns the RGB colour value of the background
    BackgroundColour = mlngBackColour
End Property

Public Property Let BackgroundColour(lngNew As Long)
    'sets the background colour of the bitmap,
    'converting system colour indecaes to RGB values
    'if necessary
    mlngBackColour = GetSystemColour(lngNew)
End Property

Public Property Get Bits(Optional ByVal lngNumBytes As Long = -1) As Byte()
    'This will return the bits of the bitmap in a zero based array of bytes
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim bytBits()       As Byte     'holds all the bits of the bitmap
    
    'make sure that a valid range of bits was specified
    Select Case lngNumBytes
    Case -1
        'default to the entire bitmap
        lngNumBytes = (Me.Height * 3 * (((Me.Width * 3) + 3) And &HFFFFFFFC))
    
    Case 0, Is < -1
        'return a single bit
        lngNumBytes = 0
    End Select
    
    'resize the array to the number of bytes to return
    ReDim bytBits(1 To lngNumBytes)
    
    'make sure there is a bitmap there
    If (mudtBitmap.lngDc = 0) Then
        'return nothing
        Bits = bytBits
        Exit Property
    End If
    
    'get the bitmap bits
    lngResult = GetBitmapBits(mudtBitmap.lngBitmap, _
                              lngNumBytes, _
                              bytBits(1))
    ReDim Preserve bytBits(1 To lngResult)   'resize to the number of bits used
    
    'return the array
    Bits = bytBits
End Property

Public Property Let Bits(Optional ByVal lngNumBytes As Long = -1, _
                         ByRef bytBits() As Byte)
    'This will set the colour information for the bitmap using the
    'array of bytes as colour data
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    
    'make sure there is a bitmap there
    If (mudtBitmap.lngDc = 0) Then
        'set nothing
        Exit Property
    End If
    
    'make sure that a valid range of bits was specified
    Select Case lngNumBytes
    Case LBound(bytBits)
        'return a single bit
        lngNumBytes = LBound(bytBits)
    
    Case (lngNumBytes < LBound(bytBits)) Or _
         (lngNumBytes > UBound(bytBits)), -1
        'default to the entire bitmap
        lngNumBytes = (UBound(bytBits) - LBound(bytBits)) + LBound(bytBits)
    End Select
    
    'get the bitmap bits
    lngResult = SetBitmapBits(mudtBitmap.lngBitmap, _
                              lngNumBytes, _
                              bytBits(LBound(bytBits)))
End Property

Public Property Get hdc() As Long
    'a read-only property that returns a handle to
    'the bitmap in case the user needs to use the
    'bitmap with other functions or procedures not
    'included here or with other instances of this
    'object
    
    hdc = mudtBitmap.lngDc
End Property

Public Property Let Pixel(ByVal intX As Integer, _
                          ByVal intY As Integer, _
                          ByVal lngColour As Long)
    'This will set the colour of the specified pixel provided that the
    'co-ordinates are correct
    
    Dim lngResult       As Long     'holds any returned value from an api call
    
    'this will set the colour of the specified pixel
    If (mudtBitmap.lngDc = 0) Then
        Exit Property
    End If
    
    lngResult = SetPixel(mudtBitmap.lngDc, intX, intY, lngColour)
End Property

Public Property Get Pixel(ByVal intX As Integer, _
                          ByVal intY As Integer) _
                          As Long
    'This will return the colour of the specified pixel
    
    Dim lngResult       As Long     'holds any returned value from an api call
    
    'make sure there is a bitmap created
    If (mudtBitmap.lngDc = 0) Then
        Exit Property
    End If
    
    'return the pixel value
    lngResult = GetPixel(mudtBitmap.lngDc, intX, intY)
    Pixel = lngResult
End Property

Public Property Get Width() As Long
    'returns the width of the bitmap
    Width = mudtBitmap.Area.Right - mudtBitmap.Area.Left
End Property

Public Property Let Width(ByVal intWidth As Long)
    'sets the width of the bitmap
    With mudtBitmap.Area
        Call ReSize(intWidth)
    End With
End Property

Public Property Get Height() As Long
    'returns the height of the bitmap
    Height = mudtBitmap.Area.Bottom - mudtBitmap.Area.Top
End Property

Public Property Let Height(ByVal intHeight As Long)
    'sets the height of the bitmap
    With mudtBitmap.Area
        Call ReSize(, intHeight)
    End With
End Property

Public Property Get Left() As Integer
    'returns the left position of the bitmap
    Left = mudtBitmap.Area.Left
End Property

Public Property Let Left(ByVal intLeft As Integer)
    'sets the left position of the bitmap
    With mudtBitmap.Area
        .Right = .Right + (.Left - intLeft)
        .Left = intLeft
    End With
End Property

Public Property Get Picture() As IPictureDisp
    'This will return a picture object for the bitmap
    With mudtBitmap
        If .lngDc = 0 Then
            Set Picture = Nothing
        Else
            Set Picture = hDCToPicture(.lngDc, _
                                       .lngBitmap, _
                                       .lngPointer, _
                                       vbPicTypeBitmap)
        End If
    End With
End Property

Public Property Set Picture(ByRef picPicture As IPictureDisp)
    'This will set the bitmap to match a picture object
    
    If Not picPicture Is Nothing Then
        Call PictureTohDc(picPicture)
    End If
End Property

Public Property Get Icon() As IPictureDisp
    'This will return a picture object for the bitmap
    With mudtBitmap
        If .lngDc = 0 Then
            Set Icon = Nothing
        Else
            Set Icon = hDCToPicture(.lngDc, _
                                    .lngBitmap, _
                                    .lngPointer, _
                                    vbPicTypeIcon)
        End If
    End With
End Property

Public Property Get Top() As Integer
    'returns the left position of the bitmap
    Top = mudtBitmap.Area.Top
End Property

Public Property Let Top(ByVal intTop As Integer)
    'sets the left position of the bitmap
    With mudtBitmap.Area
        .Bottom = .Bottom + (.Top - intTop)
        .Top = intTop
    End With
End Property
