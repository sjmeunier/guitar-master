Attribute VB_Name = "mdlCapture"
'**************************************
'Windows API/Global Declarations for :Ca
'     ptureWindows,CaptureForm,CaptureClient,e
'     tc...
'**************************************
Option Explicit
'declares to disable PC
Public Const SPI_SCREENSAVERRUNNING = 97


Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Boolean, ByVal fuWinIni As Long) As Long
    'global variable for capture setting
    Global Setting As Integer


Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
    End Type


Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough For 256 colors
    End Type


Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
    End Type


    #If Win32 Then
        Private Const RASTERCAPS As Long = 38
        Private Const RC_PALETTE As Long = &H100
        Private Const SIZEPALETTE As Long = 104


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type


Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hDC As Long) As Long


Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long _
    ) As Long


Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, _
    ByVal iCapabilitiy As Long) As Long


Private Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
    ByVal hDC As Long, ByVal wStartIndex As Long, _
    ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long


Private Declare Function CreatePalette Lib "gdi32" ( _
    lpLogPalette As LOGPALETTE) As Long


Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
    ByVal hObject As Long) As Long


Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
    ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, _
    ByVal YSrc As Long, ByVal dwRop As Long) As Long


Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long


Private Declare Function GetForegroundWindow Lib "user32" () As Long


Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, _
    ByVal hPalette As Long, ByVal bForceBackground As Long) As Long


Private Declare Function RealizePalette Lib "gdi32" ( _
    ByVal hDC As Long) As Long


Private Declare Function GetWindowDC Lib "user32" ( _
    ByVal hWnd As Long) As Long


Private Declare Function GetDC Lib "user32" ( _
    ByVal hWnd As Long) As Long


Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As Long, lpRect As RECT) As Long


Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
    ByVal hDC As Long) As Long


Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Type PicBmp
    Size As Long


Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
    End Type


Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
    PicDesc As PicBmp, RefIID As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#ElseIf Win16 Then
    Private Const RASTERCAPS As Integer = 38
    Private Const RC_PALETTE As Integer = &H100
    Private Const SIZEPALETTE As Integer = 104


Private Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
    End Type


Private Declare Function CreateCompatibleDC Lib "GDI" ( _
    ByVal hDC As Integer) As Integer


Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
    ByVal hDC As Integer, ByVal nWidth As Integer, _
    ByVal nHeight As Integer) As Integer


Private Declare Function GetDeviceCaps Lib "GDI" ( _
    ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer


Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
    ByVal hDC As Integer, ByVal wStartIndex As Integer, _
    ByVal wNumEntries As Integer, _
    lpPaletteEntries As PALETTEENTRY) As Integer


Private Declare Function CreatePalette Lib "GDI" ( _
    lpLogPalette As LOGPALETTE) As Integer


Private Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, _
    ByVal hObject As Integer) As Integer


Private Declare Function BitBlt Lib "GDI" (ByVal hDCDest As Integer, _
    ByVal XDest As Integer, ByVal YDest As Integer, _
    ByVal nWidth As Integer, ByVal nHeight As Integer, _
    ByVal hDCSrc As Integer, ByVal XSrc As Integer) As Long
                                                        

#End If

                                                
'**************************************
' Name: CaptureWindows,CaptureForm,Captu
'     reClient,etc...
' Description:Screen capture code.
' By: StonePage
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'
'Warranty:
'code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************

' CreateBitmapPicture
' - Creates a bitmap type Picture object
'     from a bitmap and palette
'
' hBmp
' - Handle to a bitmap
'
' hPal
' - Handle to a Palette
' - Can be null if the bitmap doesn't us
'     e a palette
'
' Returns
' - Returns a Picture object containing
'     the bitmap


#If Win32 Then


Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture
    Dim r As Long
#ElseIf Win16 Then


Public Function CreateBitmapPicture(ByVal hBmp As Integer, _
    ByVal hPal As Integer) As Picture
    Dim r As Integer
#End If
Dim Pic As PicBmp
' IPicture requires a reference to "Stan
'     dard OLE Types"
Dim IPic As IPicture
Dim IID_IDispatch As GUID
' Fill in with IDispatch Interface ID


With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With
' Fill Pic with necessary parts


With Pic
    .Size = Len(Pic) ' Length of structure
    .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
    .hBmp = hBmp ' Handle To bitmap
    .hPal = hPal ' Handle To palette (may be null)
End With
' Create Picture object
r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
' Return the new Picture object
Set CreateBitmapPicture = IPic
End Function
' CaptureWindow
' - Captures any portion of a window
'
' hWndSrc
' - Handle to the window to be captured
'
' Client
' - If True CaptureWindow captures from
'     the client area of the window
' - If False CaptureWindow captures from
'     the entire window
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
' - Specify the portion of the window to
'     capture
' - Dimensions need to be specified in p
'     ixels
'
' Returns
' - Returns a Picture object containing
'     a bitmap of the specified
' portion of the window that was capture
'     d


#If Win32 Then


Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal Client As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
#ElseIf Win16 Then


Public Function CaptureWindow(ByVal hWndSrc As Integer, _
    ByVal Client As Boolean, ByVal LeftSrc As Integer, _
    ByVal TopSrc As Integer, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Integer
    Dim hBmp As Integer
    Dim hBmpPrev As Integer
    Dim r As Integer
    Dim hDCSrc As Integer
    Dim hPal As Integer
    Dim hPalPrev As Integer
    Dim RasterCapsScrn As Integer
    Dim HasPaletteScrn As Integer
    Dim PaletteSizeScrn As Integer
#End If
Dim LogPal As LOGPALETTE
' Depending on the value of Client get t
'     he proper device context


If Client Then
    hDCSrc = GetDC(hWndSrc) ' Get device context For client area
Else
    hDCSrc = GetWindowDC(hWndSrc) ' Get device context For entire window
End If
' Create a memory device context for the
'     copy process
hDCMemory = CreateCompatibleDC(hDCSrc)
' Create a bitmap and place it in the me
'     mory DC
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
' Get screen properties
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette
' If the screen has a palette make a cop
'     y and realize it


If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    ' Create a copy of the system palette
    LogPal.palVersion = &H300
    LogPal.palNumEntries = 256
    r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
    hPal = CreatePalette(LogPal)
    ' Select the new palette into the memory
    '     DC and realize it
    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)
End If
' Copy the on-screen image into the memo
'     ry DC
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
LeftSrc, TopSrc, vbSrcCopy)
' Remove the new copy of the the on-scre
'     en image
hBmp = SelectObject(hDCMemory, hBmpPrev)
' If the screen has a palette get back t
'     he palette that was selected
' in previously


If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If
' Release the device context resources b
'     ack to the system
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
' Call CreateBitmapPicture to create a p
'     icture object from the bitmap
' and palette handles. Then return the r
'     esulting picture object.
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
' CaptureScreen
' - Captures the entire screen
'
' Returns
' - Returns a Picture object containing
'     a bitmap of the screen


Public Function CaptureScreen() As Picture


    #If Win32 Then
        Dim hWndScreen As Long
    #ElseIf Win16 Then
        Dim hWndScreen As Integer
    #End If
    ' Get a handle to the desktop window
    hWndScreen = GetDesktopWindow()
    ' Call CaptureWindow to capture the enti
    '     re desktop give the handle and
    ' return the resulting Picture object
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
    Screen.Width \ Screen.TwipsPerPixelX, _
    Screen.Height \ Screen.TwipsPerPixelY)
End Function
' CaptureForm
' - Captures an entire form including ti
'     tle bar and border
'
' frmSrc
' - The Form object to capture
' Returns
' - Returns a Picture object containing
'     a bitmap of the entire form


Public Function CaptureForm(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the enti
    '     re form given it's window
    ' handle and then return the resulting P
    '     icture object
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, _
    frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
    frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function
' CaptureClient
' - Captures the client area of a form
'
' frmSrc
' - The Form object to capture
'
' Returns
' - Returns a Picture object containing
'     a bitmap of the form's client
' area


Public Function CaptureClient(frmSrc As Form) As Picture
    ' Call CaptureWindow to capture the clie
    '     nt area of the form given it's
    ' window handle and return the resulting
    '     Picture object
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, _
    frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
    frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function
' CaptureActiveWindow
' - Captures the currently active window
'     on the screen
'
' Returns
' - Returns a Picture object containing
'     a bitmap of the active window


Public Function CaptureActiveWindow(cX As Long, cY As Long, cWidth As Long, cHeight As Long) As Picture


    #If Win32 Then
        Dim hWndActive As Long
        Dim r As Long
    #ElseIf Win16 Then
        Dim hWndActive As Integer
        Dim r As Integer
    #End If
    Dim RectActive As RECT
    ' Get a handle to the active/foreground
    '     window
    hWndActive = GetForegroundWindow()
    ' Get the dimensions of the window
    r = GetWindowRect(hWndActive, RectActive)
    ' Call CaptureWindow to capture the acti
    '     ve window given it's handle and
    ' return the Resulting Picture object
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, cX, cY, cWidth, cHeight)
End Function
' PrintPictureToFitPage
' - Prints a Picture object as big as po
'     ssible
'
' Prn
' - Destination Printer object
'
' Pic
' - Source Picture object


Public Sub PrintPictureToFitPage(Pic As Picture)
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    ' Determine if picture should be printed
    '     in landscape or portrait and
    ' set the orientation


    If Pic.Height >= Pic.Width Then
        Printer.Orientation = vbPRORPortrait ' Taller than wide
    Else
        Printer.Orientation = vbPRORLandscape ' Wider than tall
    End If
    ' Calculate device independent Width to
    '     Height ratio for picture
    PicRatio = Pic.Width / Pic.Height
    ' Calculate the dimentions of the printa
    '     ble area in HiMetric
    PrnWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHiMetric)
    PrnHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to
    '     Height ratio for printer
    PrnRatio = PrnWidth / PrnHeight
    ' Scale the output to the printable area
    '


    If PicRatio >= PrnRatio Then
        ' Scale picture to fit full width of pri
        '     ntable area
        PrnPicWidth = Printer.ScaleX(PrnWidth, vbHiMetric, Printer.ScaleMode)
        PrnPicHeight = Printer.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
        Printer.ScaleMode)
    Else
        ' Scale picture to fit full height of pr
        '     intable area
        PrnPicHeight = Printer.ScaleY(PrnHeight, vbHiMetric, Printer.ScaleMode)
        PrnPicWidth = Printer.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
        Printer.ScaleMode)
    End If
    ' Print the picture using the PaintPictu
    '     re method
    Printer.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub
                                                        

