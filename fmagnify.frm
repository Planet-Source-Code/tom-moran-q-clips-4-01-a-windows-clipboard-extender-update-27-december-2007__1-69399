VERSION 5.00
Begin VB.Form fMagnify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Capture - Esc to Cancel"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2475
   Icon            =   "fmagnify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   0
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      Begin VB.Line Line2 
         X1              =   51
         X2              =   51
         Y1              =   0
         Y2              =   96
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   100
         Y1              =   47
         Y2              =   47
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         Height          =   180
         Left            =   690
         Top             =   705
         Width           =   165
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1080
      Top             =   720
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1575
      LargeChange     =   5
      Left            =   1560
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   20
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "fMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MagNifier by oigres P. Email oigres@postmaster.co.uk
'Based on the C++ tool Zoomin (Lupe?)
'New features :Resizeable form, new resolution, bug fix 12/sept/99
'All code written by oigres P.
'indented by indenter5 from http://www.BMSLtd.co.uk by Stephen Bullen
'
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPixel& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long)


Private Const HORZRES = 8
Private Const VERTRES = 10


Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Const RDW_ERASE = &H4
Const RDW_INVALIDATE = &H1
Const SRCCOPY = &HCC0020
Const WM_PAINT = &HF

Dim frmH As Long, magnify As Integer, lastcpx, lastcpy

Private Sub Form_Activate()
 Form_Resize
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case Shift And vbCtrlMask
        Select Case KeyCode
        Case vbKeyF1
            VScroll1.Value = 100
            'MsgBox "F1"
        Case vbKeyF2
            VScroll1.Value = 75
        Case vbKeyF3
            VScroll1.Value = 50
        Case vbKeyF4
            VScroll1.Value = 25
        Case vbKeyF5
            VScroll1.Value = 1
        End Select
    End Select

End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    ret = SetWindowPos(fMagnify.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Show
    Call VScroll1_Change
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
            x As Single, y As Single)
    Dim lngReturnValue As Long
    'move the form if we click on in
    If Button = vbKeyLButton Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(fMagnify.hwnd, WM_NCLBUTTONDOWN, _
                HTCAPTION, 0&)
    End If
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandle
    'with setwindow rgn I got error (invalid value on vscroll1.height)
    fMagnify.Cls 'clear rubbish between labels when resize
    VScroll1.Left = fMagnify.ScaleWidth - VScroll1.Width
    VScroll1.Height = fMagnify.ScaleHeight - Label1.Height
    'resize picturebox to fill form
    Picture1.Left = 0: Picture1.Top = 0
    Picture1.ScaleWidth = fMagnify.ScaleWidth '- VScroll1.Width
    Picture1.ScaleHeight = fMagnify.ScaleHeight ' - Label1.Height

    Picture1.Width = fMagnify.ScaleWidth - VScroll1.Width
    Picture1.Height = fMagnify.ScaleHeight - Label1.Height
    
    Label1.Top = fMagnify.ScaleHeight - Label1.Height
    'magnification label
    Label2.Left = fMagnify.ScaleWidth - Label2.Width
    Label2.Top = fMagnify.ScaleHeight - Label2.Height
    'move red crosshair to middle
    'Shape1.Left = (fMagnify.ScaleWidth \ 2) - (Shape1.Width \ 2)
    
    Line1.X2 = fMagnify.ScaleWidth + 15
    Line1.Y1 = (fMagnify.ScaleHeight \ 2 + Line1.BorderWidth) - 8
    Line1.Y2 = Line1.Y1
    
    Line2.Y2 = fMagnify.ScaleHeight
    Line2.X1 = fMagnify.ScaleWidth \ 2
    Line2.X2 = Line2.X1
    
    
    'Shape1.Top = fMagnify.ScaleHeight - Shape1.Height
    'move colour patch to bottom and middle
    Shape2.Left = (Picture1.ScaleWidth \ 2) - (Shape2.Width \ 2)
    Shape2.Top = ((Picture1.ScaleHeight \ 2) - (Shape2.Height \ 2)) + 4
    Exit Sub
errHandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'see if we ant to clear the picture off the clipboard
    'If Clipboard.GetFormat(vbCFBitmap) Then
    '    If MsgBox("Do you want to clear the clipboard?", vbYesNo, "Magnifier") = vbYes Then
    '        Clipboard.Clear
    '    End If
    
  '  End If
    'exit program
    Unload Me
    Set Form1 = Nothing ' delete variable object
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbKeyLButton Then
    Shape2.Visible = Not Shape2.Visible
End If
End Sub

Private Sub Timer1_Timer()
    Dim cp As POINTAPI 'cursor position
    Dim dsDC As Long, lpPT As POINTAPI
    Dim screenColour As Long
    GetCursorPos cp ' cp has cursor position assigned to it
    
    Label1.Caption = "X= " & cp.x & Space(6 - Len(CStr(cp.x))) & ":Y= " & cp.y

    'check the screen size
    'get desktop device context- to copy from
    dsDC = GetDC(0&)
   ' screenColour = &HFFFFFF And GetPixel(dsDC, cp.X, cp.Y)
   ' Shape1.FillColor = screenColour
   ' fMagnify.Caption = "MagNifier - " & Hex(screenColour)
    'get screen width, height
    hr = GetDeviceCaps(dsDC, HORZRES)
    vr = GetDeviceCaps(dsDC, VERTRES)

    dshwnd = GetDesktopWindow()
    '      vscroll1=1..100 so 1/100=.1; 100/100=1;New Resolution
    Percent = VScroll1.Value / 100
    'new zoom size
    lengthx = (fMagnify.ScaleWidth - VScroll1.Width) * Percent
    lengthy = (fMagnify.ScaleHeight - Label1.Height) * Percent
    'center image about mouse
    offsetx = lengthx \ 2
    offsety = lengthy \ 2
    'actual area to blit to
    blitareax = fMagnify.ScaleWidth - VScroll1.Width
    blitareay = fMagnify.ScaleHeight - Label1.Height
    
    'stop copying the screen off the edges <0 and  >horzres
    'Store the last cursor position that were valid
    If cp.x - offsetx >= 0 And cp.x + offsetx < hr Then '800=screen width
        lastcpx = cp.x
        'Debug.Print "X= " & cp.X
     End If
     If cp.y - offsety >= 0 And cp.y + offsety < vr Then '600= screen height
        'Debug.Print "Y= " & cp.Y
          lastcpy = cp.y  '                dest hdc ,destx,desty,width,height, sourceDC, source x,sourcey,sourcewidth,sourceheight,raster operation
            'ret = StretchBlt(fMagnify.hdc, 0, 0, blitareax, blitareay, dsDC, cp.X - offsetx, cp.Y - offsety, lengthx, lengthy, SRCCOPY)
    End If
    '
    '                destination dc                       , source dc
    ret = StretchBlt(Picture1.hdc, 0, 0, blitareax, blitareay, dsDC, lastcpx - offsetx, lastcpy - offsety, lengthx, lengthy, SRCCOPY)
    Picture1.Refresh
    ''Picture1.Line (0, (Picture1.ScaleHeight \ 2))-(Picture1.ScaleWidth, (Picture1.ScaleHeight \ 2)), &H0&
    'cross mark on form
    'fMagnify.Line (0, 0)-(fMagnify.ScaleWidth - VScroll1.Width, fMagnify.ScaleHeight - Label1.Height)
    'fMagnify.Line (fMagnify.ScaleWidth - VScroll1.Width, 0)-(0, fMagnify.ScaleHeight - Label1.Height)
    ReleaseDC dshwnd, dsDC 'previous bug not releasing memory
End Sub

Private Sub VScroll1_Change()
    'magnify = VScroll1.Value ;100 is max vscroll value
    'output 2 decimal places
    Label2.Caption = Format(100 / VScroll1.Value, "FIXED")
    Line1.BorderWidth = Val(Label2)
    Line2.BorderWidth = Val(Label2)
    
    Picture1.SetFocus
End Sub

Private Sub VScroll1_Scroll()
    Label2.Caption = Format(100 / VScroll1.Value, "FIXED")
    Picture1.SetFocus
End Sub
