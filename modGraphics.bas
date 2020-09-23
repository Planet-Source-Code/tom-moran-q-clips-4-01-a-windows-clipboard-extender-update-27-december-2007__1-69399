Attribute VB_Name = "modGraphics"
Option Explicit

'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const ERRORAPI = 0

Public Declare Function SetWindowPos _
        Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_AND = 1

Public Const RGN_MIN = RGN_AND
Public Const RGN_OR = 2

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Long, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Global Const SRCCOPY = &HCC0020

Public ColorCapType As Integer

Private Function ZHex(lHex As Long, iZeros As Integer) As String
  'Returns a HEX string of specified length (pad zeros on left)
  ZHex = Right$(String$(iZeros - 1, "0") & Hex$(lHex), iZeros)
End Function

Public Function MakeHexRGB(r As Long, G As Long, B As Long) As String
  'Returns hex value for rgb color values
  MakeHexRGB = ZHex(r, 2) & ZHex(G, 2) & ZHex(B, 2)
End Function

Public Function MakeHexLong(lngColor As Long) As String
  Dim r As Long, G As Long, B As Long
  r = rgbRed(lngColor)
  G = rgbGreen(lngColor)
  B = rgbBlue(lngColor)
  'Returns hex value for a long color value
  MakeHexLong = ZHex(r, 2) & ZHex(G, 2) & ZHex(B, 2)
End Function

Public Function rgbRed(RGBCol As Long) As Integer
  'Returns the Red component from an RGB Color
  rgbRed = RGBCol And &HFF
End Function

Public Function rgbGreen(RGBCol As Long) As Integer
  'Returns the Green component from an RGB Color
  rgbGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function rgbBlue(RGBCol As Long) As Integer
  'Returns the Blue component from an RGB Color
  rgbBlue = (RGBCol And &HFF0000) / &H10000
End Function


