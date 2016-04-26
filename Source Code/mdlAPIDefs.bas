Attribute VB_Name = "mdlAPIDefs"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, _
ByVal X As Integer, _
ByVal Y As Integer, _
ByVal nWidth As Integer, _
ByVal nHeight As Integer, _
ByVal hSrcDC As Long, _
ByVal xSrc As Integer, _
ByVal ySrc As Integer, _
ByVal dwRop As Long) As Integer

'Dim MouseEvent As Boolean
'Dim X1 As Single
'Dim Y1 As Single

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046


