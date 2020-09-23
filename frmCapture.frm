VERSION 5.00
Begin VB.Form frmCapture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   2  'Dot
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmCapture.frx":0CCA
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   90
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X1 As Single
Dim Y1 As Single

Const INVERSE = 6       '*Characteristic of DrawMode property(XOR).
Const SOLID = 0         '*Characteristic of DrawStyle property.
Const DOT = 2           '*Characteristic of DrawStyle property.

Dim OldX As Single  '* Mouse locations
Dim OldY As Single


Sub DrawLine(X1, Y1, X2, Y2 As Single)
   '* Save the current mode so that you can reset it on
   '* exit from this sub routine. Not needed in the sample
   '* but would need it if you are not sure what the
   '* DrawMode was on entry to this procedure.
   Dim SavedMode As Integer
   SavedMode% = DrawMode

   '* Set to XOR
   Me.DrawMode = INVERSE

   '*Draw a box or line
     Me.Line (X1, Y1)-(X2, Y2), , B

   '* Reset the DrawMode
   Me.DrawMode = SavedMode%
End Sub

Private Sub Form_Load()
  'Capture desktop and make it this forms background picture
  Dim DeskhWnd As Long, DeskDC As Long
  Me.WindowState = vbMaximized
  DeskhWnd& = GetDesktopWindow()
  DeskDC& = GetDC(DeskhWnd&)
  BitBlt Me.hdc, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
  ReleaseDC DeskhWnd&, DeskDC&
  Me.Picture = Me.Image
  MousePointer = 99
  fMagnify.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'User pressed escape so unload
  If KeyCode = vbKeyEscape Then
   Unload Me
   Unload fMagnify
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then

 DrawStyle = DOT
 MousePointer = 99

'* Store the initial start of the line to draw.
   X1 = x
   Y1 = y

   '* Make the last location equal the starting location
   OldX = X1
   OldY = Y1
 End If
  

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Button Then

      '* Erase the previous line.
      Call DrawLine(X1, Y1, OldX, OldY)

      '* Draw the new line.
      Call DrawLine(X1, Y1, x, y)

      '* Save the coordinates for the next call.
      OldX = x
      OldY = y
   End If

End Sub

Private Sub CaptureIt(xStart As Single, xEnd As Single, yStart As Single, yEnd As Single)
  Dim Left As Long, Top As Long, Right As Long, Bottom, xx As Long
  Dim lWidth As Long, lHeight As Long
  Dim FileName, Ftemp As String
  
  On Error Resume Next
  If iQSound Then PlayWaveSound App.Path & "\capture.wav"
  xEnd = xEnd + 1
  yEnd = yEnd + 1
  'Get left, right, top and bottom regarldess of where they started and ended
  Left = IIf(xStart > xEnd, xEnd, xStart)
  Right = IIf(xStart < xEnd, xEnd, xStart)
  Top = IIf(yStart > yEnd, yEnd, yStart)
  Bottom = IIf(yStart < yEnd, yEnd, yStart)
  lWidth = (Right - Left)
  lHeight = (Bottom - Top)
  
  If lWidth <= 0 Or lHeight <= 0 Then GoTo PROC_TOOSMALL  'Nothing to capture
  
  With picTemp
    .Cls  'Clear our picture box that holds the image till copied to clipboar
    .Width = lWidth 'Set it's hight and width
    .Height = lHeight
  End With
  
  Me.Cls  'Clear screen so we don't get the box and dimensions
  BitBlt picTemp.hdc, 0, 0, lWidth, lHeight, Me.hdc, Left, Top, SRCCOPY   'Copy screen to picture box
  
  Clipboard.Clear 'Clear clipboard
  Clipboard.SetData picTemp.Image 'Copy image to clipboard
  Unload fMagnify

  
PROC_EXIT:
  Exit Sub
  
PROC_TOOSMALL:
  GoTo PROC_EXIT
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    CaptureIt X1, x, Y1, y  'Do the capture
    Unload Me 'Unload form
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If CapButton Then
  frmMainClip.Show
  CapButton = False
 End If
End Sub

