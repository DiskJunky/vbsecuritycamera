VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJPEGi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 0

'cJPEGI.cls  "Integer Only JPEG Encoder Class"
'Copyright ©2003 Ron van Tilburg for transform code, free use of mods for educational uses
'
'Adapted/Modified/Restructured from cJPEG (©2003 John Korejwa) by Ron van Tilburg 29 Nov 2003
'for commercial interests please approach all authors, licensing conditions as below
'========================================================================================================
'This is a completely Integer Version of JPEG encoder, with a high quality, fast integer transform.
'This integer transform is not an intgere approximation of the FP transform but rather uses a different
'mathematical approackh to approximating the DCT. Its ideally suited to asm if someone has the way-with-all.
'
'I've tried to control roundings as best I can. It's slightly faster in the IDE and 30% or so faster than the
'FP version it emulates when compiled. Various tweaks (inlining,constants) and other restructuring have occurred
'to get this speed. Small errors also corrected and redundant code removed. In particular it pays to break
'routines with a lot of stack variables up to give the compiler the best chance of register optimisation.
'
'It produces jpgs that are smaller than cJPEG on the whole for a given quality. This is because the losses due
'to rounding do count, but its pretty hard to see with the naked eye.
'BTW: Turning on all optimisations in EXE Properties really makes both versions fly.
'
'This would all have come to nought if John below hadn't actually got the perserverence to write it in
'the first place!!. Believe me its a well built piece of code...
'
'     RVT 29 Nov 2003 (rivit@f1.net.au)
'
'========================================================================================================
'The Original was:
'
'Class Name:   cJpeg.cls  "JPEG Encoder Class"
'Author:       John Korejwa  <korejwa@tiac.net>
'Version:      0.9 beta  [26 / November / 2003]
'
'
'Legal:
'        This class is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This product includes JPEG compression code developed by John Korejwa.  <korejwa@tiac.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.
'
'
'Credits:
'        Special thanks to Barry G., a government research scientist who took an interest in my
'        steganography software and research in late 1999.  I never met Barry in person, but he
'        was kind enough to buy and mail me a book with the ISO DIS 10918-1 JPEG standard.
'
'
'Description:  This class contains code for compressing pictures, sampled via hDC, into
'              baseline .JPG files.  Please report any errors or unusual behavior to the email
'              address above.
'
'Dependencies: None
'

'JPEG Marker Constants                (Note: VB compiler does not compile unused constants)

'Non-Differential Huffman Coding
Private Const SOF0    As Long = &HC0& 'Baseline DCT
Private Const SOF1    As Long = &HC1& 'Extended sequential DCT
Private Const SOF2    As Long = &HC2& 'Progressive DCT
Private Const SOF3    As Long = &HC3& 'Spatial (sequential) lossless

'Differential Huffman coding
Private Const SOF5    As Long = &HC5& 'Differential sequential DCT
Private Const SOF6    As Long = &HC6& 'Differential progressive DCT
Private Const SOF7    As Long = &HC7& 'Differential spatial

'Non-Differential arithmetic coding
Private Const JPG     As Long = &HC8& 'Reserved for JPEG extentions
Private Const SOF9    As Long = &HC9& 'Extended sequential DCT
Private Const SOF10   As Long = &HCA& 'Progressive DCT
Private Const SOF11   As Long = &HCB& 'Spatial (sequential) lossless

'Differential arithmetic coding
Private Const SOF13   As Long = &HCD& 'Differential sequential DCT
Private Const SOF14   As Long = &HCE& 'Differential progressive DCT
Private Const SOF15   As Long = &HCF& 'Differential Spatial

'Other Markers
Private Const DHT     As Long = &HC4& 'Define Huffman tables
Private Const DAC     As Long = &HCC& 'Define arithmetic coding conditioning(s)
Private Const RSTm    As Long = &HD0& 'Restart with modulo 8 count "m"
Private Const RSTm2   As Long = &HD7& 'to 'Restart with modulo 8 count "m"
Private Const SOI     As Long = &HD8& 'Start of image
Private Const EOI     As Long = &HD9& 'End of image
Private Const SOS     As Long = &HDA& 'Start of scan
Private Const DQT     As Long = &HDB& 'Define quantization table(s)
Private Const DNL     As Long = &HDC& 'Define number of lines
Private Const DRI     As Long = &HDD& 'Define restart interval
Private Const DHP     As Long = &HDE& 'Define hierarchical progression
Private Const EXP     As Long = &HDF& 'Expand reference components
Private Const APP0    As Long = &HE0& 'Reserved for application segments
Private Const APPF    As Long = &HEF& '  to Reserved for application segments
Private Const JPGn    As Long = &HF0& 'Reserved for JPEG Extentions
Private Const JPGn2   As Long = &HFD& '  to Reserved for JPEG Extentions
Private Const COM     As Long = &HFE& 'Comment
Private Const RESm    As Long = &H2&  'Reserved
Private Const RESm2   As Long = &HBF& '  to Reserved
Private Const TEM     As Long = &H1&  'For temporary use in arithmetic coding

'Consider these arrays of constants.
'They are initialized with the class and do not change.
Private QLumin(63)       As Byte      'Standard Luminance   Quantum (for 50% quality)
Private QChrom(63)       As Byte      'Standard Chrominance Quantum (for 50% quality)
Private TwoP(31)         As Long      'Powers of two  2^P
Private TwoR(31)         As Long      'Powers of two  2^(R-1)
Private TwoB(255)        As Byte      'The number of bits in value
Private EIDCTScale(7, 7) As Long      'Encoding Constants for scaling IntDCT Coefficients
Private DIDCTScale(7, 7) As Long      'Decoding Constants for scaling IntDCT Coefficients
Private ZigZag(7, 7)     As Long      'Zig Zag order of 8X8 block of samples

'API constants
Private Const BLACKONWHITE    As Long = 1 'nStretchMode constants for
Private Const COLORONCOLOR    As Long = 3 '  SetStretchBltMode() API function
Private Const HALFTONE        As Long = 4 'HALFTONE not supported in Win 95, 98, ME

Private Const BI_RGB          As Long = 0
Private Const DIB_RGB_COLORS  As Long = 0

'Variable types needed for DIBSections.
Private Type SAFEARRAYBOUND
  cElements         As Long
  lLbound           As Long
End Type

Private Type SAFEARRAY
  cDims             As Integer
  fFeatures         As Integer
  cbElements        As Long
  cLocks            As Long
  pvData            As Long
  Bounds(0)         As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
  rgbBlue           As Byte
  rgbGreen          As Byte
  rgbRed            As Byte
  rgbReserved       As Byte
End Type

Private Type BITMAPINFOHEADER
  biSize            As Long
  biWidth           As Long
  biHeight          As Long
  biPlanes          As Integer
  biBitCount        As Integer
  biCompression     As Long
  biSizeImage       As Long
  biXPelsPerMeter   As Long
  biYPelsPerMeter   As Long
  biClrUsed         As Long
  biClrImportant    As Long
End Type

Private Type BITMAPINFO
  bmiHeader         As BITMAPINFOHEADER
  bmiColors         As RGBQUAD
End Type

'API needed for creating DIBSections for sampling and pixel access.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection2 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long   'lplpVoid changed to ByRef
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Custom variable types used for this JPEG encoding implementation
Private Type QUANTIZATIONTABLE
  Qk(63)              As Integer 'Quantization Values
  FScale(63)          As Long    'Multiplication values to scale and Quantize   FDCT output
  IScale(63)          As Long    'Multiplication values to scale and DeQuantize IDCT input
End Type

Private Type HUFFMANTABLE
  Bits(15)            As Byte    'Number of huffman codes of length i+1
  HuffVal(255)        As Byte    'Huffman symbol values
  EHufLN(255)         As Byte    'Huffman codelength in Bits for symbol i
  EHufSI(255)         As Long    'Huffman code size for symbol i
  EHufCO(255)         As Long    'Huffman code      for symbol i
  MinCode(15)         As Long    '
  MaxCode(15)         As Long    'Largest code value for length i+1
End Type

Private Type Component
  Ci                  As Long    'Component ID                       [0-255]
  Hi                  As Long    'Horizontal Sampling Factor         [1-4]
  Vi                  As Long    'Vertical   Sampling Factor         [1-4]
  Tqi                 As Long    'Quantization Table Select          [0-3]
  Data()              As Integer 'DCT Coefficients                   [DC(0),AC(1-63)]*nBlocks
End Type

Private PP            As Long    'Sample Precision [8, 12]
Private YY            As Long    'Number of lines             [Image Height] after clipping
Private XX            As Long    'Number of samples per line  [Image Width]  after clipping
Private Nf            As Long    'Number of components in Frame

Private HMax          As Long    'Maximum horizontal sampling frequency
Private VMax          As Long    'Maximum vertical   sampling frequency

Private m_Data()      As Byte    'JPEG File Data
Private m_Chr         As Long    'Current Character in m_Data
Private m_Ptr         As Long    'Byte index in m_Data
Private m_Bit         As Long    'Bit  index in m_Chr

Private m_Block(7, 7) As Long    'Buffer for calculating intDCT  (7,7)

Private QTable(3)     As QUANTIZATIONTABLE  '4 Quantization Tables
Private HuffDC(3)     As HUFFMANTABLE       '4 DC Huffman Tables
Private HuffAC(3)     As HUFFMANTABLE       '4 AC Huffman Tables
Private Comp()        As Component          'Scan Components

Private m_Quality     As Long
Private m_Comment     As String

'========================================================================================
'      I N T E G E R   D I S C R E T E   C O S I N E   T R A N S F O R M A T I O N
'========================================================================================

Private Sub intDCT()    'Integer Lifted, Loeffler factorisation 22 shifts,40 adds by Jie Liang, Trac Tran
                        'see IEEE TRANSACTIONS ON SIGNAL PROCESSING VOL49,#12 DEC 2001
                                
 'This is carefully optimized and its easy to screw its performance up
 'However its ideally suited to asm, so go for it Vlad!

 Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long
 Dim t4 As Long, t5 As Long, t6 As Long, t7 As Long
 Dim t8 As Long, t9 As Long, t10 As Long, t11 As Long
 Dim i  As Long

  For i = 0 To 7                                  'Process 1D intDCT on each row
    t0 = m_Block(i, 0)
    t8 = m_Block(i, 7)
    t7 = t0 - t8
    t0 = t0 + t8

    t1 = m_Block(i, 1)
    t8 = m_Block(i, 6)
    t6 = t1 - t8
    t1 = t1 + t8

    t2 = m_Block(i, 2)
    t8 = m_Block(i, 5)
    t5 = t2 - t8
    t2 = t2 + t8

    t3 = m_Block(i, 3)
    t8 = m_Block(i, 4)
    t4 = t3 - t8
    t3 = t3 + t8

    'top half
    t8 = t0 + t3
    t9 = t1 + t2
    t8 = t8 + t9
    m_Block(i, 0) = t8
    m_Block(i, 4) = t8 \ 2 - t9                             '  1/2 *t8

    t10 = t1 - t2
    t11 = t0 - t3
    t10 = (t11 \ 4 + t11 \ 8 + t11 \ 32) - t10              ' 13/32*t11
    m_Block(i, 6) = t10
    m_Block(i, 2) = t11 - (t10 \ 4 + t10 \ 16 + t10 \ 32)   '-11/32*t10

    'bottom half
    t7 = t7 - (t4 \ 4 + t4 \ 32 + t4 \ 64)                  '-19/64*t4
    t4 = t4 + (t7 \ 2 + t7 \ 16)                            '+ 9/16*t7
    t7 = t7 - (t4 \ 4 + t4 \ 32 + t4 \ 64)                  '-19/64*t4

    t6 = t6 - (t5 \ 16 + t5 \ 32)                           '- 3/32*t5
    t5 = t5 + (t6 \ 8 + t6 \ 16)                            '+ 3/16*t6
    t6 = t6 - (t5 \ 16 + t5 \ 32)                           '- 3/32*t5

    m_Block(i, 3) = t7 - t5
    m_Block(i, 5) = t4 - t6

    t8 = t4 + t6
    t11 = t5 + t7 + t8
    m_Block(i, 1) = t11
    m_Block(i, 7) = t11 \ 2 - t8                            '  1/2
  Next i

  For i = 0 To 7                                  'Process 1D intDCT on each column
    t0 = m_Block(0, i)
    t8 = m_Block(7, i)
    t7 = t0 - t8
    t0 = t0 + t8

    t1 = m_Block(1, i)
    t8 = m_Block(6, i)
    t6 = t1 - t8
    t1 = t1 + t8

    t2 = m_Block(2, i)
    t8 = m_Block(5, i)
    t5 = t2 - t8
    t2 = t2 + t8

    t3 = m_Block(3, i)
    t8 = m_Block(4, i)
    t4 = t3 - t8
    t3 = t3 + t8

    'top half
    t8 = t0 + t3
    t9 = t1 + t2
    t8 = t8 + t9
    m_Block(0, i) = t8
    m_Block(4, i) = t8 \ 2 - t9                               '  1/2 *t8

    t10 = t1 - t2
    t11 = t0 - t3
    t10 = (t11 \ 4 + t11 \ 8 + t11 \ 32) - t10                ' 13/32*t11
    m_Block(6, i) = t10
    m_Block(2, i) = t11 - (t10 \ 4 + t10 \ 16 + t10 \ 32)     '-11/32*t10

    'bottom half
    t7 = t7 - (t4 \ 4 + t4 \ 32 + t4 \ 64)                    '-19/64*t4
    t4 = t4 + (t7 \ 2 + t7 \ 16)                              '+ 9/16*t7
    t7 = t7 - (t4 \ 4 + t4 \ 32 + t4 \ 64)                    '-19/64*t4

    t6 = t6 - (t5 \ 16 + t5 \ 32)                             '- 3/32*t5
    t5 = t5 + (t6 \ 8 + t6 \ 16)                              '+ 3/16*t6
    t6 = t6 - (t5 \ 16 + t5 \ 32)                             '- 3/32*t5

    m_Block(3, i) = t7 - t5
    m_Block(5, i) = t4 - t6

    t8 = t4 + t6
    t11 = t5 + t7 + t8
    m_Block(1, i) = t11
    m_Block(7, i) = t11 \ 2 - t8                              '  1/2
  Next i

End Sub

'========================================================================================
'                                Q U A N T I Z A T I O N
'========================================================================================

Private Sub TransformAndQuantize(ByRef Pixel() As Byte, ByVal CompIndex As Long, ByRef FScale() As Long, _
                                 ByVal Width As Long, ByVal Height As Long, ByVal RowMod As Long)

 Dim i          As Long   'Pixel Index (Horizontal)
 Dim j          As Long   'Pixel Index (Vertical)
 Dim k          As Long   'index of an 8X8 block of pixels
 Dim p          As Long   'DCT Index (horizontal)
 Dim q          As Long   'DCT Index (vertical)
 Dim t          As Long   'COMPONENT in Zig Zag order at index BlkPtr
 Dim n          As Long   'Size of Image
 Dim BlkPtr     As Long   'Index for quantized intDCT value (in component data)
 Dim BlkMod     As Long   'StepWidth for a Block
 Dim RowModS    As Long   'Length of a Sample Row
 Dim BlkModS    As Long   'Length of a Sample Row of Blocks
 Dim fr         As Long
 Dim fg         As Long
 Dim fb         As Long
 Dim fc         As Long

  With Comp(CompIndex)

    Select Case CompIndex
     Case 0: 'Luminance rm = 0.299; gm = 0.587; bm = 0.114; s = 128
      fr = 19595&: fg = 38470: fb = 7471&: fc = 8388608
     Case 1: 'Chrominance [Blue-Yellow] rm = -0.16874; gm = -0.33126; bm = 0.5; s = 0
      fr = -11059&: fg = -21709&: fb = 32768: fc = 0
     Case 2: 'Chrominance [Red-Green] rm = 0.5; gm = -0.41869; bm = -0.08131; s=0
      fr = 32768: fg = -27439&: fb = -5329&: fc = 0
    End Select

    ReDim .Data(Width * Height - 1)
    BlkPtr = 0                      'Reset output Quantized intDCT Coefficient Index

    n = Height * RowMod
    RowModS = Width * 3
    BlkMod = RowMod - 24
    BlkModS = 8 * RowMod - RowModS
    i = 0
    Do While i < n
      k = RowModS
      Do
        j = i
        '======================================================================================================
        For q = 0 To 7                              'Get 8X8 block of level shifted YCbCr values
          For p = 0 To 7
            m_Block(p, q) = fr * Pixel(i + 2) + fg * Pixel(i + 1) + fb * Pixel(i) - fc
            i = i + 3
          Next p
          i = i + BlkMod
        Next q
        '======================================================================================================
        'Take 8X8 block of unscaled DCT coefficients [m_Block(0-7, 0-7)],
        'Scale, Quantize, and store the results in data() array of
        'COMPONENT in Zig Zag order at index BlkPtr
        '======================================================================================================
        Call intDCT            'Calculate the integer DCT (it is scaled by 65536)

        For q = 0 To 7
          For p = 0 To 7
            t = ZigZag(p, q)
            .Data(BlkPtr + t) = (((m_Block(p, q) + 16384&) \ 32768) * FScale(t) + 131072) \ 262144
            '                          keep 1 decimal bit           2* true scale         4*scale
          Next p
        Next q
        BlkPtr = BlkPtr + 64
        '======================================================================================================
        i = j + 24
        k = k - 24                                        'Start Next Block this Row
      Loop Until k = 0
      i = i + BlkModS
    Loop
  End With

End Sub

Public Property Let Quality(ByVal vData As Long)

 'The JPEG compression standard does not have a formal definition for image Quality.
 'This implementation defines Quality as an integer value between 1 and 100, and
 'generates quantization tables based on the value given.
 '
 'Quality < 50  -  Poor image quality with high compression
 'Quality = 75  -  Good quality pictures for displaying on a monitor or web page ... typical for general use
 'Quality = 92  -  High quality with non-optimal compression ... Appropriate for printing ... [typical digital camera "max quality" setting]
 'Quality > 95  -  Wasteful ... very poor compression with little image quality improvement.  Use 24-bit BMP TrueColor if you need quality this high.

 Dim i           As Long
 Dim qvalue      As Long
 Dim maxvalue    As Long
 Dim scalefactor As Long

  If PP = 12 Then
    maxvalue = 65535
   Else
    maxvalue = 255
  End If

  If vData > 0 And vData <= 100 Then
    m_Quality = vData
   Else
    m_Quality = 75    'else default 75
  End If

  If (m_Quality < 50) Then
    If m_Quality = 0 Then
      scalefactor = 5000
     Else
      scalefactor = 5000 \ m_Quality
    End If
   Else
    If m_Quality >= 100 Then
      scalefactor = 0
     Else
      scalefactor = 200 - m_Quality * 2
    End If
  End If

  With QTable(0)
    For i = 0 To 63
      qvalue = (QLumin(i) * scalefactor + 50) \ 100
      If qvalue <= 0 Then
        qvalue = 1
       ElseIf qvalue > maxvalue Then
        qvalue = maxvalue
      End If
      .Qk(i) = qvalue
    Next i
  End With

  With QTable(1)
    For i = 0 To 63
      qvalue = (QChrom(i) * scalefactor + 50) \ 100
      If qvalue <= 0 Then
        qvalue = 1
       ElseIf qvalue > maxvalue Then
        qvalue = maxvalue
      End If
      .Qk(i) = qvalue
    Next i
  End With

  ExpandDQT 0
  ExpandDQT 1

End Property

Public Property Get Quality() As Long

  Quality = m_Quality

End Property

Private Sub ExpandDQT(ByVal Tqi As Long)

 Dim i          As Long
 Dim j          As Long
 Dim k          As Long
 Dim maxvalue   As Long

  With QTable(Tqi)
    If PP = 12 Then
      maxvalue = 65535
     Else
      maxvalue = 255
    End If

    For i = 0 To 7
      For j = 0 To 7
        k = ZigZag(i, j)
        If .Qk(k) < 1 Or .Qk(k) > maxvalue Then Err.Raise 1, , "Bad Quantization Table"
        .FScale(k) = (EIDCTScale(i, j) + .Qk(k) \ 2) \ .Qk(k) '.Q(k)=1..maxval   'FScale always 0..65536/131072
      Next j
    Next i
  End With

End Sub

'================================================================================
'                 H U F F M A N   T A B L E   G E N E R A T I O N
'================================================================================
Private Sub BuildHuffman(ByRef TheHuff As HUFFMANTABLE, ByRef Freq() As Long)

 'Generate optimized values for BITS and HUFFVAL in a HUFFMANTABLE
 'based on symbol frequency counts.  freq must be dimensioned freq(0-256)
 'and contain counts of symbols 0-255.  freq is destroyed in this procedure.

 Dim i              As Long
 Dim j              As Long
 Dim k              As Long
 Dim n              As Long
 Dim w              As Long
 Dim v1             As Long
 Dim v2             As Long
 Dim others(256)    As Long
 Dim codesize(256)  As Long
 Dim Bits(256)      As Long
 Dim swp            As Long
 Dim swp2           As Long

 'Initialization

  For i = 0 To 256  'Initialize others to -1, (this value terminates chain of indicies)
    others(i) = -1
  Next i
  Freq(256) = 1     'Add dummy symbol to guarantee no code will be all '1' bits

  'Generate codesize()   [find huffman code sizes]
  Do                                   'do loop for (#non-zero-frequencies - 1) times
    v1 = -1                            'find highest v1 for      least value of freq(v1)>0
    v2 = -1                            'find highest v2 for next least value of freq(v2)>0
    swp = 2147483647                   'Max Long variable
    swp2 = 2147483647
    For i = 0 To 256
      w = Freq(i)
      If w <> 0 Then
        If w <= swp2 Then
          If w <= swp Then
            swp2 = swp
            v2 = v1
            swp = w
            v1 = i
           Else
            swp2 = w
            v2 = i
          End If
        End If
      End If
    Next i

    If v2 = -1 Then
      Freq(v1) = 0                     'all elements in freq are now set to zero
      Exit Do                          'done
    End If

    Freq(v1) = Freq(v1) + Freq(v2)     'merge the two branches
    Freq(v2) = 0
    codesize(v1) = codesize(v1) + 1    'Increment all codesizes in v1's branch

    Do While others(v1) >= 0
      v1 = others(v1)
      codesize(v1) = codesize(v1) + 1
    Loop

    others(v1) = v2                    'chain v2 onto v1's branch
    codesize(v2) = codesize(v2) + 1    'Increment all codesizes in v2's branch

    Do While others(v2) >= 0
      v2 = others(v2)
      codesize(v2) = codesize(v2) + 1
    Loop
  Loop

  'Count BITS  [find the number of codes of each size]
  n = 0
  For i = 0 To 256
    w = codesize(i)
    If w <> 0 Then
      Bits(w) = Bits(w) + 1
      If n < w Then n = w    'Keep track of largest codesize
    End If
  Next i

  'Adjust BITS  [limit code lengths to 16 bits]
  i = n
  Do While i > 16
    Do While Bits(i) > 0
      For j = i - 2 To 1 Step -1        'Since symbols are paired for the longest Huffman
        If Bits(j) > 0 Then Exit For    'code, the symbols are removed from this length
      Next j                            'category two at a time.  The prefix for the pair
      Bits(i) = Bits(i) - 2             '(which is one bit shorter) is allocated to one
      Bits(i - 1) = Bits(i - 1) + 1     'of the pair;  then, (skipping the BITS entry for
      Bits(j + 1) = Bits(j + 1) + 2     'that prefix length) a code word from the next
      Bits(j) = Bits(j) - 1             'shortest non-zero BITS entry is converted into
    Loop                                'a prefix for two code words one bit longer.
    i = i - 1
  Loop
  Bits(i) = Bits(i) - 1                 'Remove dummy symbol code from the code length count

  'Copy BITS and HUFFVAL to the HUFFMANTABLE  [HUFFVAL sorted by code length, then by value]
  With TheHuff
    For i = 1 To 16
      .Bits(i - 1) = Bits(i)
    Next i
    k = 0
    For i = 1 To n
      For j = 0 To 255
        If codesize(j) = i Then
          .HuffVal(k) = j
          k = k + 1
        End If
      Next j
    Next i
  End With

End Sub

Private Sub ExpandHuffman(ByRef TheHuff As HUFFMANTABLE, Optional ByVal MaxSymbol As Long = 255)

 'Given a HUFFMANTABLE with valid BITS and HUFFVAL, generate tables for
 'EHUFCO, EHUFSI, MAXCODE, and MINCODE so the table may be used for compression
 'and/or decompression.  In JPEG, MaxSymbol is 255 for an AC Huffman Table.  For
 'DC Tables, MaxSymbol is 11 for PP=8 bit precission, or 15 for PP=12 bit precission.

 Dim i          As Long 'Index for BITS
 Dim j          As Long 'Index for HUFFVAL
 Dim k          As Long 'Index for last HUFFVAL of length (i+1)
 Dim si         As Long 'Huffman code size  ( =2^i )
 Dim code       As Long 'Huffman code
 Dim symbol     As Long 'Huffman symbol

  With TheHuff

    For i = 0 To 255
      .EHufLN(i) = 0
      .EHufSI(i) = 0      'Clear existing values so we can
      .EHufCO(i) = -1     'check for duplicate huffman symbols
    Next i

    j = 0
    si = 1
    code = 0
    For i = 0 To 15
      k = j + .Bits(i)
      If k > 256 Then Err.Raise 1, , "Bad Huffman Table" 'more than 256 symbols
      If j = k Then                                      'no codes of length i+1
        .MinCode(i) = j - code
        .MaxCode(i) = -1
       Else
        .MinCode(i) = j - code
        Do While j < k
          symbol = .HuffVal(j)                                            'read symbol, make sure it's valid
          If symbol > MaxSymbol Then Err.Raise 1, , "Bad Huffman Table"   'invalid symbol
          If .EHufCO(symbol) >= 0 Then Err.Raise 1, , "Bad Huffman Table" 'duplicate symbol
          .EHufLN(symbol) = i
          .EHufSI(symbol) = si                                            'assign code for symbol
          .EHufCO(symbol) = code
          code = code + 1
          j = j + 1
        Loop
        .MaxCode(i) = code - 1
      End If
      si = si + si
      If code >= si Then Err.Raise 1, , "Bad Huffman Table" 'code does not fit into available bits
      code = code + code
    Next i
    If j = 0 Then Err.Raise 1, , "Bad Huffman Table"        'No huffman symbols???
  End With

End Sub

'Estimate a worstcase code length for an entire table
Private Function CodeLength(ByRef TheHuff As HUFFMANTABLE, ByRef Freq() As Long) As Long

 Dim j As Long
 Dim k As Long
 Dim sum As Long

  sum = 0
  With TheHuff
    For j = 0 To 255
      If Freq(j) <> 0 Then
        k = .EHufLN(j) + nBits(j)           'Length of size category + value, in bits
        sum = sum + Freq(j) * k             'Sum all occurances of this coefficient, in bits
      End If
    Next j
  End With

  CodeLength = (sum + 7) \ 8

End Function

'=================================================================================================================
'                           E N T R O P Y   C O D I N G
'=================================================================================================================
Private Sub WriteBitsBegin()

  m_Chr = 0
  m_Bit = 128

End Sub

Private Sub WriteBitsEnd()

  If m_Bit <> 128 Then WriteBits m_Bit, -1

End Sub

Private Sub WriteBits(ByVal mask As Long, ByVal code As Long)

  Do While mask > 0

    If (code And mask) <> 0 Then m_Chr = (m_Chr Or m_Bit)

    If m_Bit = 1 Then             'We completed a byte ...
      m_Data(m_Ptr) = m_Chr       '    add it to the stream
      If m_Chr = 255 Then         'Pad a zero byte and advance pointer
        m_Ptr = m_Ptr + 1         'advance pointer
        m_Data(m_Ptr) = 0         'insert zero
      End If
      m_Ptr = m_Ptr + 1           'advance pointer
      m_Chr = 0                   'clear byte buffer and reset bit index
      m_Bit = 128
     Else                          'increment to next bit position to write
      m_Bit = m_Bit \ 2
    End If
    mask = mask \ 2
  Loop

End Sub

Private Sub EncodeCoefficients(ByRef Data() As Integer, ByRef p As Long, ByRef Pred As Long, _
                               ByVal Td As Long, ByVal Ta As Long)

 'Use Huffman tables to compress a block of 64 quantized DCT coefficients to the local
 'm_Data() byte array.  The coefficients are input in the data() array starting at index p.
 'Pred is the predictor for the DC coefficient.  Td and Ta are indexes to the local DC and AC
 'Huffman Tables to use.

 Dim r     As Long
 Dim rs    As Long
 Dim code  As Long
 Dim p2    As Long

  p2 = p + 64

  code = Data(p) - Pred
  Pred = Data(p)
  p = p + 1

  With HuffDC(Td)
    If code = 0 Then
      WriteBits .EHufSI(0), .EHufCO(0)      'append symbol for size category
     Else
      rs = nBits(code)
      If code < 0 Then code = code - 1
      WriteBits .EHufSI(rs), .EHufCO(rs)    'append symbol for size category
      WriteBits TwoR(rs), code              'append diff
    End If
  End With

  With HuffAC(Ta)
    r = 0
    Do
      code = Data(p)
      If code = 0 Then
        r = r + 16
       Else
        Do While r > 240
          WriteBits .EHufSI(240), .EHufCO(240)          'append RUN16 (a run of 16 zeros)
          r = r - 256
        Loop
        rs = nBits(code)
        If code < 0 Then code = code - 1
        WriteBits .EHufSI(r + rs), .EHufCO(r + rs)    'append run length, size category
        WriteBits TwoR(rs), code                      'append AC value
        r = 0
      End If
      p = p + 1
    Loop While p < p2 'should be equal on exit
    If r <> 0 Then WriteBits .EHufSI(0), .EHufCO(0)     'append EOB (end of block)
  End With

End Sub

'========================================================================================
'                      C O L L E C T I N G   S T A T I S T I C S
'========================================================================================
'These procedures collect statistics of run-length and size categories of DCT coefficients
'so optimized Huffman tables can be generated to compress them.

Private Sub CollectStatisticsAC(ByRef Data() As Integer, ByRef FreqAC() As Long)

 Dim code As Long
 Dim n    As Long 'Number of coefficients in data()
 Dim p    As Long 'Index for current data() coefficient
 Dim p2   As Long
 Dim r    As Long 'Run length of zeros
 Dim rs   As Long 'Run-length/Size-category Symbol

  n = UBound(Data) + 1 '+1 not necessary
  p = 0
  Do While p < n
    p = p + 1
    p2 = p + 63
    r = 0
    Do While p < p2 'should be equal on exit
      code = Data(p)                            'is always <=|255| ??
      If code = 0 Then
        r = r + 16                              '16*count of zeroes
       Else
        Do While r > 240
          FreqAC(240) = FreqAC(240) + 1         'runs of 16 zeros
          r = r - 256
        Loop
        rs = r + nBits(code)                    'nr of bits for this code
        FreqAC(rs) = FreqAC(rs) + 1             'freq of 16*nzeroes + nbits(code)+1
        r = 0
      End If
      p = p + 1
    Loop
    If r <> 0 Then FreqAC(0) = FreqAC(0) + 1    'EOB value
  Loop

End Sub

Private Sub CollectStatisticsDCNonInterleaved(ByRef Data() As Integer, ByRef FreqDC() As Long)

 Dim Pred  As Long     'DC Predictor
 Dim n     As Long     'Number of coefficients in data()
 Dim p     As Long     'Index for current data() coefficient
 Dim rs    As Long     'Size category for Diff

  n = UBound(Data) + 1
  p = 0
  Pred = 0
  Do While p < n
    rs = nBits(Data(p) - Pred)    'nr of bits in delta from last value
    Pred = Data(p)
    FreqDC(rs) = FreqDC(rs) + 1   'frequencies of bitsizes
    p = p + 64
  Loop

End Sub

Private Sub CollectStatisticsDCInterleaved(ByRef Data() As Integer, ByRef FreqDC() As Long, _
                                           ByVal Hi As Long, ByVal Vi As Long)

 Dim p()       As Long      'Index to .data in component f for scanline g
 Dim f         As Long      'Index counter  (component)
 Dim g         As Long      'Index counter  (sampling factor, vertical)
 Dim h         As Long      'Index counter  (sampling factor, horizontal)
 Dim i         As Long      'Index counter  (MCU horizontal)
 Dim j         As Long      'Index counter  (MCU vertical)
 Dim n         As Long      'Number of coefficients in data()
 Dim s         As Long      'Size category for Diff
 Dim Pred      As Long      'DC Predictor
 Dim pLF       As Long      'Line Feed for p in data
 Dim MCUr      As Long      'Number of complete 8X8 blocks in rightmost MCU
 Dim MCUx      As Long      'Number of MCUs per scanline
 Dim MCUy      As Long      'Number of MCU scanlines

  n = UBound(Data) + 1
  ReDim p(Vi - 1)

  MCUx = (XX + 8 * HMax - 1) \ (8 * HMax)
  MCUy = (YY + 8 * VMax - 1) \ (8 * VMax)

  h = (-Int(-XX * Hi / HMax) + 7) \ 8  'Width of scanline in data (MCUs)

  For g = 0 To Vi - 1                  'Initialize .data pointers
    p(g) = 64 * h * g
  Next g
  pLF = 64 * h * (Vi - 1)              'Initialize .data pointer advancer

  MCUr = (h Mod Hi)                    'Number of complete 8X8 Blocks in rightmost MCU
  If MCUr = 0 Then MCUr = Hi

  For j = 1 To MCUy - 1

    'MCUs across a scanline
    For i = 1 To MCUx - 1
      For g = 1 To Vi
        For h = 1 To Hi
          s = nBits(Data(p(g - 1)) - Pred)  'nr of bits in diff
          Pred = Data(p(g - 1))
          FreqDC(s) = FreqDC(s) + 1
          p(g - 1) = p(g - 1) + 64
        Next h
      Next g
    Next i

    'Rightmost MCU
    For g = 1 To Vi
      For h = 1 To Hi
        If h > MCUr Then  'Pad with dummy block
          s = 0
         Else
          s = nBits(Data(p(g - 1)) - Pred)  'nr of bits in diff
          Pred = Data(p(g - 1))
          p(g - 1) = p(g - 1) + 64
        End If
        FreqDC(s) = FreqDC(s) + 1
      Next h
    Next g

    'Advance data pointers
    For g = 0 To Vi - 1
      p(g) = p(g) + pLF
    Next g
  Next j

  'Bottommost MCU Scanline
  For i = 1 To MCUx
    For g = 1 To Vi
      For h = 1 To Hi
        If p(g - 1) >= n Or (i = MCUx And h > MCUr) Then 'Pad with dummy block
          s = 0
         Else
          s = nBits(Data(p(g - 1)) - Pred)  'nr of bits in Diff
          Pred = Data(p(g - 1))
          p(g - 1) = p(g - 1) + 64
        End If
        FreqDC(s) = FreqDC(s) + 1
      Next h
    Next g
  Next i

End Sub

'================================================================================
'                           I M A G E   S A M P L I N G
'================================================================================
Public Sub SetSamplingFrequencies(ByVal H1 As Long, ByVal v1 As Long, _
                                  ByVal H2 As Long, ByVal v2 As Long, _
                                  ByVal H3 As Long, ByVal v3 As Long)

 'This class always samples and compresses pictures in YCbCr colorspace.  The first component, Y,
 'represents the Luminance of the pixels.  This is "how bright" a pixel is.  The Cb and Cr
 'components are Chrominance, which is a measure of how far from neutral-white (toward a color)
 'a pixel is.  The human visual sensory system can discriminate Luminance differences about
 'twice as well as it can discriminate Chrominance differences.
 '
 'Virtually all JPEG files are in YCbCr colorspace.  Other JPEG compliant colorspaces exist, but
 'they are used in specialty equipment.  For example, people in the astronomy or medical fields
 'choose colorspaces that best record the information they are interested in, and don't care about
 'how pretty the picture looks to a person when displayed on a computer monitor.
 '[Apple/Machintosh sometimes uses a four component colorspace, but that colorspace is rare and
 'not widely supported]
 '
 'Sampling frequencies define how often each component is sampled.  Higher frequencies store more
 'information, while lower frequencies store less.  Typically, sampling frequencies are set at
 '2,2, 1,1, 1,1.  This corresponds to the human visual sensory system.  The first component,
 'Luminance, is sampled twice as much because our eyes notice differences in Luminance quite easily.
 'The two Chrominance components are sampled half as much as because our eyes can't distinguish
 'the difference in color changes as well.  One Luminance value is sampled for every pixel, and
 'one Chrominance value is sampled for each 2X2 block of pixels.
 '
 'Digital cameras typically record at sampling frequencies of 1,1, 1,1, 1,1.  This samples every
 'pixel for all three components.  The quality of the picture is a little better when viewed by
 'a person, but the compression benefits drop significantly.  If the picture to be compressed
 'is from a Scanner or Digital camera, and you plan on printing it in the future, and storage
 'space is not a problem, then sampling at these frequencies makes sense.  Otherwise, if you only
 'plan on using the picture to display on a monitor or a web page, [2,2, 1,1, 1,1] makes the
 'most sense.
 '
 'The JPEG standard specifies that sampling frequencies may range from 1-4 for each component
 'in both directions.  However, if any component has a sampling frequency of '3', and another
 'component has a coresponding sampling frequency of '2' or '4', the downsampling process
 'will map fractional pixels to sample values.  This is leagal in the JPEG standard, and this
 'class will compress fractional pixel samplings, but this is not widely supported.  It is
 'highly recommended to AVOID SAMPLING FACTORS OF 3 for maximum compatability with JPEG decoders.
 '
 'Some JPEG encoders avoid the fractional pixel problem by only allowing the end user to pick
 'a "sub-sampling" value.  In such "Sub Sampling" schemes, all Chrominance frequencies are set
 'to one, and the (one or two) sub-sampling value(s) specify Luminance frequencies.
 '
 'There should *never* be an error raised if you are using this class correctly.  It should
 'not be possible for the end user to specify illegal sampling frequency values!
 '[For tinkerers - If you delete the error raising code and specify illegal sampling
 'frequencies, this class will procede to create a non-JPEG compliant file with the values
 'specified]

 Dim i As Long

  If H1 < 1 Or H1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
  If v1 < 1 Or v1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"

  If (H2 Or H3 Or v2 Or v3) = 0 Then  'if H2,H3,V2,V3 are all zero ...
    Nf = 1         'Luminance only.
    ReDim Comp(0)
    Comp(0).Hi = 1 'Set up for sampling Greyscale
    Comp(0).Vi = 1 '(Black and White picture)
   Else
    If H2 < 1 Or H2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    If H3 < 1 Or H3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    If v2 < 1 Or v2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    If v3 < 1 Or v3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    Nf = 3         'YCbCr
    ReDim Comp(2)
    Comp(0).Hi = H1
    Comp(0).Vi = v1
    Comp(0).Tqi = 0
    Comp(1).Hi = H2
    Comp(1).Vi = v2
    Comp(1).Tqi = 1
    Comp(2).Hi = H3
    Comp(2).Vi = v3
    Comp(2).Tqi = 1
  End If

  HMax = -1
  VMax = -1
  For i = 0 To Nf - 1 'determine max h, v sampling factors
    If HMax < Comp(i).Hi Then HMax = Comp(i).Hi
    If VMax < Comp(i).Vi Then VMax = Comp(i).Vi
  Next i

End Sub

Public Function SampleHDC(ByVal lHDC As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long, _
                          Optional ByVal lSrcLeft As Long = 0, _
                          Optional ByVal lSrcTop As Long = 0) As Long

 'Given a valid hDC and dimensions, generate component samplings of an Image.
 'A DIBSection is created to hold Sample(s) of the Image, from which the Image is
 'decomposed into YCbCr components.
 'Returns: 0 = Success
 '         1 = API error while generating a DIBSection

 Dim SA         As SAFEARRAY
 Dim BI         As BITMAPINFO 'Type containing the Bitmap information
 Dim hDIb       As Long       'Handle to the DIBSection
 Dim hBmpOld    As Long       'Handle to the old bitmap in the DC, for clear up
 Dim hDC        As Long       'Handle to the Device context holding the DIBSection
 Dim lPtr       As Long       'Address of memory pointing to the DIBSection's bits

 Dim f          As Long       'Index counter for components
 Dim Pixel()    As Byte       'Byte array containing pixel data
 Dim lRowMod    As Long       'Width of A Row in Bytes

 Dim xs         As Long       'Sample width
 Dim ys         As Long       'Sample height
 Dim xs8        As Long       'Sample width  (padded to 8 pixel barrier)
 Dim ys8        As Long       'Sample height (padded to 8 pixel barrier)
 Dim xsLast     As Long       'Sample width  (for previous component)
 Dim ysLast     As Long       'Sample height (for previous component)

  PP = 8
  XX = lWidth                                 'The to be sampled image block
  YY = lHeight                                'The to be sampled image block

  'Create a DIBSection to store Sampling(s) of the Image
  hDC = CreateCompatibleDC(0)
  If hDC = 0 Then
    SampleHDC = 1 'CreateCompatibleDC() API Failure
   Else

    lWidth = (XX + 7) And &HFFFFFFF8 '8 byte barrier for 8X8 data units
    lHeight = (YY + 7) And &HFFFFFFF8
    lRowMod = (((lWidth * 24) + 31&) And Not 31&) \ 8&
    With BI.bmiHeader
      .biSize = Len(BI.bmiHeader)
      .biWidth = lWidth
      .biHeight = -lHeight                        'top to bottom please
      .biPlanes = 1
      .biBitCount = 24
      .biCompression = BI_RGB
      .biSizeImage = lRowMod * lHeight  '4 byte barrier
    End With

    hDIb = CreateDIBSection2(hDC, BI, DIB_RGB_COLORS, lPtr, 0, 0)
    If hDIb = 0 Then
      SampleHDC = 1 'CreateDIBSection2() API Failure
     Else
      With SA                                   'This code copies the pointer of the 2-D bitmap
        .cbElements = 1                         'pixel data to the pointer of the 1D Pixel() array.
        .cDims = 1                              'This allows you to read/modify the pixel data
        .Bounds(0).lLbound = 0                  'as if it were stored in the Pixel() array.
        .Bounds(0).cElements = lRowMod * lHeight
        .pvData = lPtr                          'Note that this is extreamly efficient, since it copies
      End With                                  'a pointer to the data, and not the data itself.
      hBmpOld = SelectObject(hDC, hDIb)         'Select DIBSection into DC

      If SetStretchBltMode(hDC, HALFTONE) = 0 Then SetStretchBltMode hDC, COLORONCOLOR

      For f = 0 To Nf - 1
        With Comp(f)
          .Ci = f + 1 'Assign an ID to this component

          xs = -Int(-XX * .Hi / HMax)        'determine Sample dimensions
          ys = -Int(-YY * .Vi / VMax)
          xs8 = ((xs + 7) And &HFFFFFFF8)    'Sample dimensions with 8X8 barrier
          ys8 = ((ys + 7) And &HFFFFFFF8)

          If xs8 <> xsLast Or ys8 <> ysLast Then                          'We need to Sample the Image

            If xs = XX And ys = YY Then                                   'Just copy the image to our DIBSection
              BitBlt hDC, 0, 0, xs, ys, lHDC, lSrcLeft, lSrcTop, vbSrcCopy
             Else                                                         'Resample/Resize the Image
              StretchBlt hDC, 0, 0, xs, ys, lHDC, lSrcLeft, lSrcTop, lWidth, lHeight, vbSrcCopy
            End If

            If xs < xs8 Then                                              'Pad right of Sample to 8 block barrier
              StretchBlt hDC, xs + 1, 0, xs8 - xs, ys, hDC, xs, 0, 1, ys, vbSrcCopy
            End If

            If ys < ys8 Then                                              'Pad bottom of Sample to 8 block barrier
              StretchBlt hDC, 0, ys + 1, xs8, ys8 - ys, hDC, 0, ys, xs8, 1, vbSrcCopy
            End If
          End If

          CopyMemory ByVal VarPtrArray(Pixel), VarPtr(SA), 4& 'Get Pixel array descriptor
          'Read 8X8 blocks of pixels, convert from RGB->YCbCr colorspace, FDCT and Quantize
          'the data, store the results in .Data() of this component
          Call TransformAndQuantize(Pixel(), f, QTable(.Tqi).FScale(), xs8, ys8, lRowMod)

          CopyMemory ByVal VarPtrArray(Pixel), 0&, 4 'Clear the Pixel array descriptor
          xsLast = xs
          ysLast = ys
        End With
      Next f

      SelectObject hDC, hBmpOld 'Select CompatibleDC  (unselect DIBSection)
      DeleteObject hDIb         'Delete DIBSection
    End If
    DeleteObject hDC              'Delete CompatibleDC
  End If

End Function

Public Property Let Comment(ByVal Value As String)

 'Assigning a value to this property will add the text Comment to the JPEG file.

  If Len(Value) > 65535 Then Err.Raise 1, , "Illegal Comment Length"
  m_Comment = Value

End Property

Public Property Get Comment() As String

  Comment = m_Comment

End Property

'================================================================================
'                         E M I T I N G   M A R K E R S
'================================================================================
Private Sub InsertJFIF()

  If m_Ptr + 17 > UBound(m_Data) Then Err.Raise 9 'Copymemory will write past bounds of m_Data()

  CopyMemory m_Data(m_Ptr + 0), &H1000E0FF, 4&    'APP0 Marker, Length(APP0)=16
  CopyMemory m_Data(m_Ptr + 4), &H4649464A, 4&    '"JFIF"
  CopyMemory m_Data(m_Ptr + 8), &H10100, 4&       '"/0", Version Major=1, Version Minor=1
  'Units=0  [0=pixel, 1=dpi, 2=dots/cm]
  CopyMemory m_Data(m_Ptr + 12), &H1000100, 4&    'Horizontal pixel density = 1 (dot per pixel)
  'Vertical   pixel density = 1 (dot per pixel)
  CopyMemory m_Data(m_Ptr + 16), &H0&, 2&         'Thumbnail horizontal pixel count = 0
  m_Ptr = m_Ptr + 18                              'Thumbnail vertical   pixel count = 0

End Sub

Private Sub InsertSOF(ByVal SOFMarker As Long)

 Dim i   As Long 'Insert a Start Of Frame marker segment
 Dim Lx  As Long 'PP, YY, XX, Nf, and Ci,Hi,Vi,Tqi, must already be set

  Lx = 8 + 3 * Nf
  m_Data(m_Ptr) = 255                    'SOF
  m_Data(m_Ptr + 1) = SOFMarker And 255
  m_Data(m_Ptr + 2) = Lx \ 256           'Frame Header Length
  m_Data(m_Ptr + 3) = Lx And 255
  m_Data(m_Ptr + 4) = PP                 'Sample precision [8, 12]
  m_Data(m_Ptr + 5) = YY \ 256           'Number of Lines
  m_Data(m_Ptr + 6) = YY And 255
  m_Data(m_Ptr + 7) = XX \ 256           'Number of samples per line
  m_Data(m_Ptr + 8) = XX And 255
  m_Data(m_Ptr + 9) = Nf                 'Number of image components in frame
  m_Ptr = m_Ptr + 10
  For i = 0 To Nf - 1                      'For each component ...
    With Comp(i)
      m_Data(m_Ptr) = .Ci                  'Component identifier
      m_Data(m_Ptr + 1) = .Hi * 16 Or .Vi  'Horizontal/Vertical sampling factors
      m_Data(m_Ptr + 2) = .Tqi             'Quantization table selector
    End With
    m_Ptr = m_Ptr + 3
  Next i

End Sub

Private Sub InsertCOM(ByVal TheComment As String)

 Dim i As Long
 Dim Lx As Long

  Lx = Len(TheComment) + 2
  If Lx > 2 Then
    m_Data(m_Ptr) = 255               'COM marker
    m_Data(m_Ptr + 1) = COM
    m_Data(m_Ptr + 2) = Lx \ 256      'COM marker segment length
    m_Data(m_Ptr + 3) = Lx And 255
    m_Ptr = m_Ptr + 4
    For i = 1 To Len(TheComment)      'Comment text
      m_Data(m_Ptr) = Asc(Mid$(TheComment, i, 1))
      m_Ptr = m_Ptr + 1
    Next i
  End If

End Sub

Private Sub InsertDQT(ByVal MarkerPos As Long, ByVal Tqi As Long)

 Dim i As Long 'Call with MarkerPos = m_Ptr to insert a single table with its own DQT marker
               'Call multiple times with the same MarkerPos to include
               'multiple tables under the same DQT marker

  If m_Ptr < MarkerPos + 4 Then 'Insert Marker
    m_Ptr = MarkerPos + 4
    m_Data(m_Ptr - 4) = 255
    m_Data(m_Ptr - 3) = DQT
  End If
  
  With QTable(Tqi)
    For i = 0 To 63
      If .Qk(i) > 255 Then Exit For
    Next i
    
    If i = 64 Then              '8 bit precision
      m_Data(m_Ptr) = Tqi
      m_Ptr = m_Ptr + 1
      For i = 0 To 63
        m_Data(m_Ptr) = .Qk(i)
        m_Ptr = m_Ptr + 1
      Next i
     Else                        '16 bit precision
      If PP <> 12 Then Err.Raise 1, , "Illegal precision in Quantization Table"
      m_Data(m_Ptr) = Tqi Or 16
      m_Ptr = m_Ptr + 1
      For i = 0 To 63
        m_Data(m_Ptr) = .Qk(i) \ 256
        m_Data(m_Ptr + 1) = .Qk(i) And 255
        m_Ptr = m_Ptr + 2
      Next i
    End If
  End With

  m_Data(MarkerPos + 2) = (m_Ptr - MarkerPos - 2) \ 256& 'Insert Marker segment length
  m_Data(MarkerPos + 3) = (m_Ptr - MarkerPos - 2) And 255&

End Sub

Private Sub InsertDHT(ByVal MarkerPos As Long, ByVal HIndex As Long, ByVal isAC As Boolean)

 Dim i As Long 'Call with MarkerPos = m_Ptr to insert a single table with its own DHT marker
 Dim j As Long 'Call multiple times with the same MarkerPos to include
               'multiple tables under the same DHT marker

  If m_Ptr < MarkerPos + 4 Then 'Insert Marker
    m_Ptr = MarkerPos + 4
    m_Data(m_Ptr - 4) = 255
    m_Data(m_Ptr - 3) = DHT
  End If
  
  If isAC Then
    With HuffAC(HIndex)
      m_Data(m_Ptr) = HIndex Or 16
      m_Ptr = m_Ptr + 1
      j = 0
      For i = 0 To 15
        m_Data(m_Ptr) = .Bits(i)
        m_Ptr = m_Ptr + 1
        j = j + .Bits(i)
      Next i
      For i = 0 To j - 1
        m_Data(m_Ptr) = .HuffVal(i)
        m_Ptr = m_Ptr + 1
      Next i
    End With
   Else
    With HuffDC(HIndex)
      m_Data(m_Ptr) = HIndex
      m_Ptr = m_Ptr + 1
      j = 0
      For i = 0 To 15
        m_Data(m_Ptr) = .Bits(i)
        m_Ptr = m_Ptr + 1
        j = j + .Bits(i)
      Next i
      For i = 0 To j - 1
        m_Data(m_Ptr) = .HuffVal(i)
        m_Ptr = m_Ptr + 1
      Next i
    End With
  End If

  m_Data(MarkerPos + 2) = (m_Ptr - MarkerPos - 2) \ 256& 'Insert Marker segment length
  m_Data(MarkerPos + 3) = (m_Ptr - MarkerPos - 2) And 255&

End Sub

Private Sub InsertMarker(ByVal TheMarker As Long)

  m_Data(m_Ptr) = 255
  m_Data(m_Ptr + 1) = TheMarker
  m_Ptr = m_Ptr + 2

End Sub

'================================================================================================================
'                           E M I T I N G   S C A N S
'================================================================================================================
Private Sub InsertSOSNonInterleaved(ByVal CompIndex As Long, ByVal Td As Long, ByVal Ta As Long)

 'Insert an SOS marker and scan data for a non-interleaved Sequential scan.

 Dim p         As Long     'Index to .data in component
 Dim n         As Long
 Dim Pred      As Long     'Predictor for DC coefficient

 'Insert SOS Marker Segment

  m_Data(m_Ptr) = 255                          'SOS Marker
  m_Data(m_Ptr + 1) = SOS
  m_Data(m_Ptr + 2) = 8 \ 256                  'Marker Segment Length
  m_Data(m_Ptr + 3) = 8 And 255
  m_Data(m_Ptr + 4) = 1                        'Ns     - Number of components in Scan [1-4]
  m_Ptr = m_Ptr + 5
  m_Data(m_Ptr) = Comp(CompIndex).Ci           'Csj    - Component ID
  m_Data(m_Ptr + 1) = Td * 16 Or Ta            'Td, Ta - DC, AC entropy coder selector
  m_Ptr = m_Ptr + 2
  m_Data(m_Ptr) = 0                            'Ss     - Start of spectral selection
  m_Data(m_Ptr + 1) = 63                       'Se     - End of spectral selection
  m_Data(m_Ptr + 2) = 0                        'Ah, Al - Successive approximation bit high/low
  m_Ptr = m_Ptr + 3

  'Insert non-interleaved sequential entropy coded data
  With Comp(CompIndex)

    p = 0
    n = UBound(.Data) + 1
    Pred = 0

    WriteBitsBegin
    Do While p <> n
      EncodeCoefficients .Data, p, Pred, Td, Ta
    Loop
    WriteBitsEnd

  End With

End Sub

Private Sub InsertSOSInterleaved(ByRef CompIndex() As Long, ByRef Td() As Long, ByRef Ta() As Long, _
                                 ByVal FirstIndex As Long, ByVal SecondIndex As Long)

 'Insert an SOS marker and scan data for an interleaved Sequential scan.

 Dim f         As Long      'Index counter  (component)
 Dim g         As Long      'Index counter  (sampling factor, vertical)
 Dim h         As Long      'Index counter  (sampling factor, horizontal)
 Dim i         As Long      'Index counter  (MCU horizontal)
 Dim j         As Long      'Index counter  (MCU vertical)
 Dim Lx        As Long      'Marker Segment Length
 Dim Ns        As Long      'Number of components in Scan [1-4]
 Dim MCUx      As Long      'Number of MCUs per scanline
 Dim MCUy      As Long      'Number of MCU scanlines

 Dim p()        As Long     'Index to .data in component f for scanline g
 Dim pLF()      As Long     'Line Feed for p in .data for component f
 Dim Pred()     As Long     'Predictor for DC coefficient in component f
 Dim MCUr()     As Long     'Number of complete 8X8 blocks in rightmost MCU for component f
 Dim Pad64(63)  As Integer  '8X8 padding block for completing MCUs

  Ns = SecondIndex - FirstIndex + 1
  Lx = 6 + 2 * Ns

  'Insert SOS Marker Segment
  m_Data(m_Ptr) = 255                          'SOS Marker
  m_Data(m_Ptr + 1) = SOS
  m_Data(m_Ptr + 2) = Lx \ 256                 'Marker Segment Length
  m_Data(m_Ptr + 3) = Lx And 255
  m_Data(m_Ptr + 4) = Ns                       'Ns     - Number of components in Scan [1-4]
  m_Ptr = m_Ptr + 5
  For i = FirstIndex To SecondIndex
    m_Data(m_Ptr) = Comp(CompIndex(i)).Ci      'Csj
    m_Data(m_Ptr + 1) = Td(i) * 16 Or Ta(i)    'Td, Ta
    m_Ptr = m_Ptr + 2
  Next i
  m_Data(m_Ptr) = 0                            'Ss     - Start of spectral selection
  m_Data(m_Ptr + 1) = 63                       'Se     - End of spectral selection
  m_Data(m_Ptr + 2) = 0                        'Ah, Al - Successive approximation bit high/low
  m_Ptr = m_Ptr + 3

  'Insert interleaved sequential entropy coded data
  ReDim p(FirstIndex To SecondIndex, VMax - 1)
  ReDim Pred(FirstIndex To SecondIndex)
  ReDim pLF(FirstIndex To SecondIndex)
  ReDim MCUr(FirstIndex To SecondIndex)

  MCUx = (XX + 8 * HMax - 1) \ (8 * HMax)
  MCUy = (YY + 8 * VMax - 1) \ (8 * VMax)

  For f = FirstIndex To SecondIndex
    With Comp(CompIndex(f))
      h = (-Int(-XX * .Hi / HMax) + 7) \ 8  'Width of scanline in .data (MCUs)

      For g = 0 To .Vi - 1                  'Initialize .data pointers
        p(f, g) = 64 * h * g
      Next g
      pLF(f) = 64 * h * (.Vi - 1)           'Initialize .data pointer advancer

      MCUr(f) = (h Mod .Hi)                 'Number of complete 8X8 Blocks in rightmost MCU
      If MCUr(f) = 0 Then MCUr(f) = .Hi
    End With
  Next f

  WriteBitsBegin
  For j = 1 To MCUy - 1

    'Encode MCUs across a scanline
    For i = 1 To MCUx - 1
      For f = FirstIndex To SecondIndex '0 To Ns - 1
        With Comp(CompIndex(f))
          For g = 1 To .Vi
            For h = 1 To .Hi
              EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
            Next h
          Next g
        End With
      Next f
    Next i

    'Encode Rightmost MCU
    For f = FirstIndex To SecondIndex '0 To Ns - 1
      With Comp(CompIndex(f))
        For g = 1 To .Vi
          For h = 1 To .Hi
            If h > MCUr(f) Then 'Pad with dummy block
              Pad64(0) = Pred(f)
              EncodeCoefficients Pad64, 0, Pred(f), Td(f), Ta(f)
             Else
              EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
            End If
          Next h
        Next g
      End With
    Next f

    'Advance .data pointers
    For f = FirstIndex To SecondIndex
      For g = 0 To Comp(CompIndex(f)).Vi - 1
        p(f, g) = p(f, g) + pLF(f)
      Next g
    Next f
  Next j

  'Encode Bottommost MCU Scanline
  For i = 1 To MCUx
    For f = FirstIndex To SecondIndex
      With Comp(CompIndex(f))
        For g = 1 To .Vi
          For h = 1 To .Hi
            If p(f, g - 1) > UBound(.Data) Or (i = MCUx And h > MCUr(f)) Then 'Pad with dummy block
              Pad64(0) = Pred(f)
              EncodeCoefficients Pad64, 0, Pred(f), Td(f), Ta(f)
             Else
              EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
            End If
          Next h
        Next g
      End With
    Next f
  Next i

  WriteBitsEnd

End Sub

Private Sub InsertSequentialScans(ByRef CompIndex() As Long, ByRef Td() As Long, ByRef Ta() As Long, _
                                  ByVal FirstIndex As Long, ByVal SecondIndex As Long)

 'Insert scan components CompIndex(FirstIndex) to CompIndex(SecondIndex) sequentially in compliance
 'with JPEG rules.  Components are interleaved whenever possible to emit as few scans as possible.

 Dim f            As Long       'First Index
 Dim g            As Long       'Second Index
 Dim nb           As Long       'Number of 8X8 blocks in MCU
 Const MaxNb      As Long = 10  'Max 8X8 blocks in MCU  (10 for JPEG compliance)
 Dim flag         As Boolean    'True when ready to insert scan(s)

  f = FirstIndex
  g = FirstIndex
  nb = 0
  flag = False
  Do While f <= SecondIndex

    nb = nb + Comp(CompIndex(g)).Hi * Comp(CompIndex(g)).Vi
    g = g + 1

    If nb > MaxNb Then
      flag = True
      If f <> g - 1 Then g = g - 1
     Else
      If (g - f) = 3 Or g > SecondIndex Then flag = True
    End If

    If flag Then
      If f = g - 1 Then
        InsertSOSNonInterleaved CompIndex(f), Td(f), Ta(f)
       Else
        InsertSOSInterleaved CompIndex, Td, Ta, f, g - 1
      End If
      nb = 0
      f = g
      flag = False
    End If
  Loop

End Sub

'========================================================================================
'                               W R I T I N G   F I L E
'========================================================================================
Private Function OptimizeHuffmanTables(ByRef CompIndex() As Long, ByRef Td() As Long, ByRef Ta() As Long, _
                                       ByVal FirstIndex As Long, ByVal SecondIndex As Long) As Long

 'Optimize Huffman tables for the component indexes given.
 'Returns an estimate of the number of bytes needed for entropy coded data.
 'Estimate assumes a single scan, and entropy coded FF bytes are not followed by a zero stuff byte.

 Dim f            As Long      'First Index
 Dim g            As Long      'Second Index
 Dim i            As Long
 Dim j            As Long
 Dim CodeLen      As Long      'Total bytes required for entropy coded data
 Dim nb           As Long      'Number of 8X8 blocks in MCU
 Const MaxNb      As Long = 10 'Max 8X8 blocks in MCU  (10 for JPEG compliance)

 Dim Freq(256)    As Long      'frequency count for optimizing Huffman tables
 Dim Freq2()      As Long      'copy of freq, used for calcultating entropy coded data size
 Dim IsInter()    As Boolean   'True if component i will be interleaved
 Dim TdUsed()     As Boolean   'True if HuffDC(i) is used
 Dim TaUsed()     As Boolean   'True if HuffAC(i) is used
 Dim flag         As Boolean   'True when ready to include scan(s)

  ReDim IsInter(FirstIndex To SecondIndex)
  ReDim TaUsed(3)
  ReDim TdUsed(3)

  'Determine which components will be interleaved by InsertSequentialScans(), which tables are used
  f = FirstIndex
  g = FirstIndex
  nb = 0
  flag = False
  Do While f <= SecondIndex

    nb = nb + Comp(CompIndex(g)).Hi * Comp(CompIndex(g)).Vi
    g = g + 1

    If nb > MaxNb Then
      flag = True
      If f <> g - 1 Then g = g - 1
     Else
      If (g - f) = 3 Or g > SecondIndex Then flag = True
    End If

    If flag Then
      If f = g - 1 Then
        TdUsed(Td(f)) = True
        TaUsed(Ta(f)) = True
        IsInter(f) = False
       Else
        For i = f To g - 1
          TdUsed(Td(i)) = True
          TaUsed(Ta(i)) = True
          IsInter(i) = True
        Next i
      End If
      nb = 0
      f = g
      flag = False
    End If
  Loop

  'Optimize huffman tables for the scan sequence
  CodeLen = 0
  For i = 0 To 3
    If TdUsed(i) Then
      For f = FirstIndex To SecondIndex
        With Comp(CompIndex(f))
          If Td(f) = i Then
            If IsInter(f) Then
              Call CollectStatisticsDCInterleaved(.Data(), Freq(), .Hi, .Vi)
             Else
              Call CollectStatisticsDCNonInterleaved(.Data(), Freq())
            End If
          End If
        End With
      Next f

      'Optimize and create this DC table
      Freq2() = Freq()
      Call BuildHuffman(HuffDC(i), Freq())
      Call ExpandHuffman(HuffDC(i), IIf(PP = 12, 15, 11))
      CodeLen = CodeLen + CodeLength(HuffDC(i), Freq2())
    End If

    If TaUsed(i) Then
      For f = FirstIndex To SecondIndex
        If Td(f) = i Then Call CollectStatisticsAC(Comp(CompIndex(f)).Data(), Freq())
      Next f

      'Optimize and create this AC table
      Freq2() = Freq()
      Call BuildHuffman(HuffAC(i), Freq())
      Call ExpandHuffman(HuffAC(i), 255)
      CodeLen = CodeLen + CodeLength(HuffAC(i), Freq2())
    End If
  Next i

  OptimizeHuffmanTables = CodeLen

End Function

Public Function SaveFile(FileName As String) As Long

 Dim CompIndex()  As Long 'Indexes of Components to be included
 Dim Td()         As Long 'DC Huffman Table Selectors
 Dim Ta()         As Long 'AC Huffman Table Selectors
 Dim FileNum      As Integer
 Dim i            As Long

  If Len(FileName) = 0 Then
    SaveFile = 1          'FileName not given
   Else
    If (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0) Then
      SaveFile = 2        'File already exists
     Else

      ReDim CompIndex(Nf - 1)
      ReDim Td(Nf - 1)
      ReDim Ta(Nf - 1)

      For i = 0 To Nf - 1
        CompIndex(i) = i
        Td(i) = IIf(i = 0, 0, 1)
        Ta(i) = IIf(i = 0, 0, 1)
      Next i

      i = OptimizeHuffmanTables(CompIndex, Td, Ta, 0, Nf - 1)

      'Estimate maximum possible file size needed
      ReDim m_Data(i * 2 + 1000 + Len(m_Comment))
      m_Ptr = 0

      InsertMarker SOI                                   'SOI - Start of Image
      InsertJFIF                                         'JFIF

      If Len(m_Comment) > 0 Then InsertCOM m_Comment     'COM - Comment
      InsertCOM "cJPEGI Encoder Class, ©2003 Ron van Tilburg" & vbCrLf _
                & "Derived from cJPEG Written by ©2003 John Korejwa <korejwa@tiac.net>" & vbCrLf _
                & "Visual Basic sourcecode available at planetsourcecode.com"

      InsertDQT m_Ptr, 0                                 'DQT - Define Quantization Tables
      If Nf > 1 Then InsertDQT m_Ptr, 1

      InsertSOF SOF0                                     'SOF - Start of Frame

      InsertDHT m_Ptr, 0, isAC:=False                          'DHT - Define Huffman Tables
      InsertDHT m_Ptr, 0, isAC:=True
      If Nf > 1 Then
        InsertDHT m_Ptr, 1, isAC:=False
        InsertDHT m_Ptr, 1, isAC:=True
      End If

      InsertSequentialScans CompIndex, Td, Ta, 0, Nf - 1 'SOS - Scan Data
      InsertMarker EOI                                   'EOI - End of Image

      'Size the final byte array and write to file
      ReDim Preserve m_Data(m_Ptr - 1)
      FileNum = FreeFile
      Open FileName For Binary Access Write As FileNum
      Put #FileNum, , m_Data
      Close FileNum
      Erase m_Data

    End If
  End If

End Function

'========================================================================================
'                         C L A S S   I N I T I A L I Z A T I O N
'========================================================================================
Private Sub Class_Initialize()

 Dim i As Long
 Dim j As Long
 Dim dx As Long
 Dim zz As Long

  i = 0                   'Initialize the ZigZag() array, which maps out the
  j = 0                   'zig-zag sequence of quantized DCT coefficients
  dx = 1                  'in approximately low to high spatial frequencies
  For zz = 0 To 63
    ZigZag(i, j) = zz
    i = i + dx
    j = j - dx
    If i > 7 Then            '  0   1   5   6  14  15  27  28
      i = 7                  '  2   4   7  13  16  26  29  42
      j = j + 2              '  3   8  12  17  25  30  41  43
      dx = -1                '  9  11  18  24  31  40  44  53
     ElseIf j > 7 Then       ' 10  19  23  32  39  45  52  54
      j = 7                  ' 20  22  33  38  46  51  55  60
      i = i + 2              ' 21  34  37  47  50  56  59  61
      dx = 1                 ' 35  36  48  49  57  58  62  63
     ElseIf i < 0 Then
      i = 0 'check (j>7) first
      dx = 1
     ElseIf j < 0 Then
      j = 0
      dx = -1
    End If
  Next zz

  'Luminance Quantization table for Quality = 50
  QLumin(0) = 16:   QLumin(1) = 11:   QLumin(2) = 12:   QLumin(3) = 14
  QLumin(4) = 12:   QLumin(5) = 10:   QLumin(6) = 16:   QLumin(7) = 14
  QLumin(8) = 13:   QLumin(9) = 14:   QLumin(10) = 18:  QLumin(11) = 17
  QLumin(12) = 16:  QLumin(13) = 19:  QLumin(14) = 24:  QLumin(15) = 40
  QLumin(16) = 26:  QLumin(17) = 24:  QLumin(18) = 22:  QLumin(19) = 22
  QLumin(20) = 24:  QLumin(21) = 49:  QLumin(22) = 35:  QLumin(23) = 37
  QLumin(24) = 29:  QLumin(25) = 40:  QLumin(26) = 58:  QLumin(27) = 51
  QLumin(28) = 61:  QLumin(29) = 60:  QLumin(30) = 57:  QLumin(31) = 51
  QLumin(32) = 56:  QLumin(33) = 55:  QLumin(34) = 64:  QLumin(35) = 72
  QLumin(36) = 92:  QLumin(37) = 78:  QLumin(38) = 64:  QLumin(39) = 68
  QLumin(40) = 87:  QLumin(41) = 69:  QLumin(42) = 55:  QLumin(43) = 56
  QLumin(44) = 80:  QLumin(45) = 109: QLumin(46) = 81:  QLumin(47) = 87
  QLumin(48) = 95:  QLumin(49) = 98:  QLumin(50) = 103: QLumin(51) = 104
  QLumin(52) = 103: QLumin(53) = 62:  QLumin(54) = 77:  QLumin(55) = 113
  QLumin(56) = 121: QLumin(57) = 112: QLumin(58) = 100: QLumin(59) = 120
  QLumin(60) = 92:  QLumin(61) = 101: QLumin(62) = 103: QLumin(63) = 99

  'Chrominance Quantization table for Quality = 50
  QChrom(0) = 17:   QChrom(1) = 18:   QChrom(2) = 18:   QChrom(3) = 24
  QChrom(4) = 21:   QChrom(5) = 24:   QChrom(6) = 47:   QChrom(7) = 26
  QChrom(8) = 26:   QChrom(9) = 47:   QChrom(10) = 99:  QChrom(11) = 66
  QChrom(12) = 56:  QChrom(13) = 66:  QChrom(14) = 99:  QChrom(15) = 99
  QChrom(16) = 99:  QChrom(17) = 99:  QChrom(18) = 99:  QChrom(19) = 99
  QChrom(20) = 99:  QChrom(21) = 99:  QChrom(22) = 99:  QChrom(23) = 99
  QChrom(24) = 99:  QChrom(25) = 99:  QChrom(26) = 99:  QChrom(27) = 99
  QChrom(28) = 99:  QChrom(29) = 99:  QChrom(30) = 99:  QChrom(31) = 99
  QChrom(32) = 99:  QChrom(33) = 99:  QChrom(34) = 99:  QChrom(35) = 99
  QChrom(36) = 99:  QChrom(37) = 99:  QChrom(38) = 99:  QChrom(39) = 99
  QChrom(40) = 99:  QChrom(41) = 99:  QChrom(42) = 99:  QChrom(43) = 99
  QChrom(44) = 99:  QChrom(45) = 99:  QChrom(46) = 99:  QChrom(47) = 99
  QChrom(48) = 99:  QChrom(49) = 99:  QChrom(50) = 99:  QChrom(51) = 99
  QChrom(52) = 99:  QChrom(53) = 99:  QChrom(54) = 99:  QChrom(55) = 99
  QChrom(56) = 99:  QChrom(57) = 99:  QChrom(58) = 99:  QChrom(59) = 99
  QChrom(60) = 99:  QChrom(61) = 99:  QChrom(62) = 99:  QChrom(63) = 99

  'Encoding Quantization Scalars  (values are 0.. 0.5 scaled by 131072)
  '  Call GenEIDCT
  EIDCTScale(0, 0) = 16384: EIDCTScale(0, 1) = 16384: EIDCTScale(0, 2) = 25080: EIDCTScale(0, 3) = 23170
  EIDCTScale(0, 4) = 32768: EIDCTScale(0, 5) = 23170: EIDCTScale(0, 6) = 21407: EIDCTScale(0, 7) = 32768
  EIDCTScale(1, 0) = 16384: EIDCTScale(1, 1) = 16384: EIDCTScale(1, 2) = 25080: EIDCTScale(1, 3) = 23170
  EIDCTScale(1, 4) = 32768: EIDCTScale(1, 5) = 23170: EIDCTScale(1, 6) = 21407: EIDCTScale(1, 7) = 32768
  EIDCTScale(2, 0) = 25080: EIDCTScale(2, 1) = 25080: EIDCTScale(2, 2) = 38390: EIDCTScale(2, 3) = 35468
  EIDCTScale(2, 4) = 50159: EIDCTScale(2, 5) = 35468: EIDCTScale(2, 6) = 32768: EIDCTScale(2, 7) = 50159
  EIDCTScale(3, 0) = 23170: EIDCTScale(3, 1) = 23170: EIDCTScale(3, 2) = 35468: EIDCTScale(3, 3) = 32768
  EIDCTScale(3, 4) = 46341: EIDCTScale(3, 5) = 32768: EIDCTScale(3, 6) = 30274: EIDCTScale(3, 7) = 46341
  EIDCTScale(4, 0) = 32768: EIDCTScale(4, 1) = 32768: EIDCTScale(4, 2) = 50159: EIDCTScale(4, 3) = 46341
  EIDCTScale(4, 4) = 65536: EIDCTScale(4, 5) = 46341: EIDCTScale(4, 6) = 42813: EIDCTScale(4, 7) = 65536
  EIDCTScale(5, 0) = 23170: EIDCTScale(5, 1) = 23170: EIDCTScale(5, 2) = 35468: EIDCTScale(5, 3) = 32768
  EIDCTScale(5, 4) = 46341: EIDCTScale(5, 5) = 32768: EIDCTScale(5, 6) = 30274: EIDCTScale(5, 7) = 46341
  EIDCTScale(6, 0) = 21407: EIDCTScale(6, 1) = 21407: EIDCTScale(6, 2) = 32768: EIDCTScale(6, 3) = 30274
  EIDCTScale(6, 4) = 42813: EIDCTScale(6, 5) = 30274: EIDCTScale(6, 6) = 27969: EIDCTScale(6, 7) = 42813
  EIDCTScale(7, 0) = 32768: EIDCTScale(7, 1) = 32768: EIDCTScale(7, 2) = 50159: EIDCTScale(7, 3) = 46341
  EIDCTScale(7, 4) = 65536: EIDCTScale(7, 5) = 46341: EIDCTScale(7, 6) = 42813: EIDCTScale(7, 7) = 65536

  'Powers of 2
  Call Init_BitMasks

  'The nr of bits in value
  TwoB(0) = 0: TwoB(1) = 1: TwoB(2) = 2: TwoB(3) = 2
  For i = 4 To 7:     TwoB(i) = 3: Next i
  For i = 8 To 15:    TwoB(i) = 4: Next i
  For i = 16 To 31:   TwoB(i) = 5: Next i
  For i = 32 To 63:   TwoB(i) = 6: Next i
  For i = 64 To 127:  TwoB(i) = 7: Next i
  For i = 128 To 255: TwoB(i) = 8: Next i

  SetSamplingFrequencies 2, 2, 1, 1, 1, 1
  Quality = 75

End Sub

Private Sub Init_BitMasks()

 Dim i As Long, sb() As Variant

  sb = Array(&H1&, &H2&, &H4&, &H8&, _
       &H10&, &H20&, &H40&, &H80&, _
       &H100&, &H200&, &H400&, &H800&, _
       &H1000&, &H2000&, &H4000, &H8000&, _
       &H10000, &H20000, &H40000, &H80000, _
       &H100000, &H200000, &H400000, &H800000, _
       &H1000000, &H2000000, &H4000000, &H8000000, _
       &H10000000, &H20000000, &H40000000, &H80000000)

  For i = 0 To 30
    TwoP(i) = sb(i)
    TwoR(i + 1) = sb(i)
  Next i
  TwoP(31) = sb(31)

End Sub

'Uncomment this to generate the table above
Private Sub GenEIDCT()

 '  Dim t(7) As Double
 '  Dim i As Long, j As Long

 '  Const PI As Double = 3.14159265358979

 '  t(0) = Sin(PI / 4) / 2
 '  t(1) = 1 / Sqr(8)
 '  t(2) = 1 / (2 * Sin(3 * PI / 8))
 '  t(3) = 0.5
 '  t(4) = Sin(PI / 4)
 '  t(5) = 0.5
 '  t(6) = Sin(3 * PI / 8) / 2
 '  t(7) = 1 / Sqr(2)

 '  Open App.Path & "\EIDCT.txt" For Output As #1
 '  For i = 0 To 7
 '    For j = 0 To 7
 '      EIDCTScale(i, j) = 131072 * t(i) * t(j)
 '      Print #1, "EIDCTScale(" & i & "," & j & ")=" & EIDCTScale(i, j)
 '    Next j
 '  Next i
 '  Close #1

End Sub

'the routine below is functionally equivalent to that below but entirely in integers only
' nBits  = Int((Log(Abs(v)) * 1.442695040889))+1   '1/log(2) v<>0, 0 if v=0
Private Function nBits(ByVal v As Long) As Long    'effectively binary search for Log2 -32768<v<32767

 Const TWOE00  As Long = &H1&
 Const TWOE01  As Long = &H2&
 Const TWOE02  As Long = &H4&
 Const TWOE03  As Long = &H8&
 Const TWOE04  As Long = &H10&
 Const TWOE05  As Long = &H20&
 Const TWOE06  As Long = &H40&
 Const TWOE07  As Long = &H80&
 Const TWOE08  As Long = &H100&
 Const TWOE09  As Long = &H200&
 Const TWOE10  As Long = &H400&
 Const TWOE11  As Long = &H800&
 Const TWOE12  As Long = &H1000&
 Const TWOE13  As Long = &H2000&
 Const TWOE14  As Long = &H4000&
 Const TWOE15  As Long = &H8000&

  If v < 0 Then v = -v

  If v < TWOE08 Then                                    'use linear table lookup for values <256
    nBits = TwoB(v)
   Else                                                 'use binary search
    '    If v < TWOE08 Then                             '0..255 bits 7..0
    '      If v < TWOE04 Then                           '0.. 15 bits 3..0
    '        If v < TWOE02 Then                         '0..  1 bits 1..0
    '          If v < TWOE01 Then
    '            nBits = 1
    '          Else
    '            nBits = 2
    '          End If
    '        ElseIf v < TWOE03 Then                     '4..  7 bits 3..2
    '          nBits = 3
    '        Else
    '          nBits = 4
    '        End If
    '      Else                                         '16..255  bits 7..4
    '        If v < TWOE06 Then                         '16.. 63  bits 5..4
    '          If v < TWOE05 Then
    '            nBits = 5
    '          Else
    '            nBits = 6
    '          End If
    '        ElseIf v < TWOE07 Then                     '64..127 bits 7..6
    '          nBits = 7
    '        Else
    '          nBits = 8
    '        End If
    '      End If
    '    Else                                           '256..32768 bits 15..8
    If v < TWOE12 Then                           'bits 11..8
      If v < TWOE10 Then                         'bits  9..8
        If v < TWOE09 Then
          nBits = 9
         Else
          nBits = 10
        End If
       ElseIf v < TWOE11 Then
        nBits = 11
       Else                                       'bits 11..10
        nBits = 12
      End If
     Else                                         'bits 15..12
      If v < TWOE14 Then                         'bits 13..12
        If v < TWOE13 Then
          nBits = 13
         Else
          nBits = 14
        End If
       ElseIf v < TWOE15 Then                     'bits 15..14
        nBits = 15
       Else
        nBits = 16
      End If
    End If
  End If
  '  End If

End Function

':) Ulli's VB Code Formatter V2.16.6 (2003-Dec-06 16:05) 230 + 1731 = 1961 Lines
