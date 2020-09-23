VERSION 5.00
Begin VB.Form frmHotKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize Q-Clips Hot Keys"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Color Capture Hotkey"
      Height          =   2580
      Left            =   4050
      TabIndex        =   19
      Top             =   90
      Width           =   1920
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2160
         Index           =   2
         Left            =   180
         ScaleHeight     =   2160
         ScaleWidth      =   1560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   1560
         Begin VB.CheckBox chkColor 
            Caption         =   "WinKey"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   14
            Top             =   1605
            Width           =   1035
         End
         Begin VB.CheckBox chkColor 
            Caption         =   "Alt"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   13
            Top             =   1245
            Width           =   1035
         End
         Begin VB.CheckBox chkColor 
            Caption         =   "Shift"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   12
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkColor 
            Caption         =   "Control"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   11
            Top             =   525
            Width           =   1035
         End
         Begin VB.ComboBox cmbColorCap 
            Height          =   315
            Left            =   450
            TabIndex        =   10
            Text            =   "Combo1"
            Top             =   0
            Width           =   1035
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   45
            Picture         =   "frmHotKey.frx":0000
            Top             =   45
            Width           =   240
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6165
      TabIndex        =   16
      Top             =   1395
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   375
      Left            =   6165
      TabIndex        =   15
      Top             =   540
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Q-Clipboard Hotkey"
      Height          =   2580
      Left            =   90
      TabIndex        =   18
      Top             =   90
      Width           =   1920
      Begin VB.CheckBox chkClip 
         Caption         =   "WinKey"
         Height          =   255
         Index           =   3
         Left            =   450
         TabIndex        =   4
         Top             =   1980
         Width           =   1035
      End
      Begin VB.CheckBox chkClip 
         Caption         =   "Alt"
         Height          =   255
         Index           =   2
         Left            =   450
         TabIndex        =   3
         Top             =   1620
         Width           =   990
      End
      Begin VB.CheckBox chkClip 
         Caption         =   "Shift"
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   2
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkClip 
         Caption         =   "Control"
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   1
         Top             =   900
         Width           =   1035
      End
      Begin VB.ComboBox cmbClip 
         Height          =   315
         Left            =   495
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   360
         Width           =   1035
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2115
         Index           =   0
         Left            =   135
         ScaleHeight     =   2115
         ScaleWidth      =   1425
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   1425
         Begin VB.Image Image1 
            Height          =   240
            Left            =   0
            Picture         =   "frmHotKey.frx":058A
            Top             =   90
            Width           =   240
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Screen Capture Hotkey"
      Height          =   2580
      Left            =   2070
      TabIndex        =   17
      Top             =   90
      Width           =   1920
      Begin VB.CheckBox chkCapture 
         Caption         =   "WinKey"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   9
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CheckBox chkCapture 
         Caption         =   "Alt"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox chkCapture 
         Caption         =   "Shift"
         Height          =   195
         Index           =   1
         Left            =   585
         TabIndex        =   7
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CheckBox chkCapture 
         Caption         =   "Control"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
      Begin VB.ComboBox cmbCapture 
         Height          =   315
         Left            =   585
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   360
         Width           =   1035
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   1
         Left            =   180
         ScaleHeight     =   2055
         ScaleWidth      =   1470
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   1470
         Begin VB.Image Image2 
            Height          =   240
            Left            =   0
            Picture         =   "frmHotKey.frx":0B24
            Top             =   90
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frmHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DisplayOptions()
Dim PgmHotKey As Long

'First set the capture defaults
'=================================
'Check the extended keys
 If CapCtrlKey = True Then chkCapture(0).Value = 1
 If CapShiftKey = True Then chkCapture(1).Value = 1
 If CapAltKey = True Then chkCapture(2).Value = 1
 If CapWinKey = True Then chkCapture(3).Value = 1
 'now check the hotkey
 For i = 0 To cmbCapture.ListCount - 1
  If CapHotKey = cmbCapture.List(i) Then
   cmbCapture.ListIndex = i
   Exit For
  End If
 Next
 
 'now set the hotkeys
  If CapHotKey = "F1" Then PgmHotKey = 112
  If CapHotKey = "F2" Then PgmHotKey = 113
  If CapHotKey = "F3" Then PgmHotKey = 114
  If CapHotKey = "F4" Then PgmHotKey = 115
  If CapHotKey = "F5" Then PgmHotKey = 116
  If CapHotKey = "F6" Then PgmHotKey = 117
  If CapHotKey = "F7" Then PgmHotKey = 118
  If CapHotKey = "F8" Then PgmHotKey = 119
  If CapHotKey = "F9" Then PgmHotKey = 120
  If CapHotKey = "F10" Then PgmHotKey = 121
  If CapHotKey = "F11" Then PgmHotKey = 122
  If CapHotKey = "F12" Then PgmHotKey = 123
 
 If Len(CapHotKey) = 1 Then PgmHotKey = Asc(CapHotKey)
 
  
 frmMainClip.VBHotKey2.VKey = PgmHotKey
 frmMainClip.VBHotKey2.CtrlKey = CapCtrlKey
 frmMainClip.VBHotKey2.ShiftKey = CapShiftKey
 frmMainClip.VBHotKey2.AltKey = CapAltKey
 frmMainClip.VBHotKey2.WinKey = CapWinKey
 
 'Second we set the QClip Hotkeys
 '=================================================================
 'Check the extended keys
 If ClipCtrlKey = True Then chkClip(0).Value = 1
 If ClipShiftKey = True Then chkClip(1).Value = 1
 If ClipAltKey = True Then chkClip(2).Value = 1
 If ClipWinKey = True Then chkClip(3).Value = 1
 'now check the hotkey
 For i = 0 To cmbClip.ListCount - 1
  If ClipHotKey = cmbClip.List(i) Then
   cmbClip.ListIndex = i
   Exit For
  End If
 Next
 
 'now set the hotkeys
 
  If ClipHotKey = "F1" Then PgmHotKey = 112
  If ClipHotKey = "F2" Then PgmHotKey = 113
  If ClipHotKey = "F3" Then PgmHotKey = 114
  If ClipHotKey = "F4" Then PgmHotKey = 115
  If ClipHotKey = "F5" Then PgmHotKey = 116
  If ClipHotKey = "F6" Then PgmHotKey = 117
  If ClipHotKey = "F7" Then PgmHotKey = 118
  If ClipHotKey = "F8" Then PgmHotKey = 119
  If ClipHotKey = "F9" Then PgmHotKey = 120
  If ClipHotKey = "F10" Then PgmHotKey = 121
  If ClipHotKey = "F11" Then PgmHotKey = 122
  If ClipHotKey = "F12" Then PgmHotKey = 123

If Len(ClipHotKey) = 1 Then PgmHotKey = Asc(ClipHotKey)

 
 frmMainClip.VBHotKey1.VKey = PgmHotKey
 frmMainClip.VBHotKey1.CtrlKey = ClipCtrlKey
 frmMainClip.VBHotKey1.ShiftKey = ClipShiftKey
 frmMainClip.VBHotKey1.AltKey = ClipAltKey
 frmMainClip.VBHotKey1.WinKey = ClipWinKey
 
 'Third we set the ColorCapture Hotkeys
 '=================================================================
 'Check the extended keys
 If ColorCtrlKey = True Then chkColor(0).Value = 1
 If ColorShiftKey = True Then chkColor(1).Value = 1
 If ColorAltKey = True Then chkColor(2).Value = 1
 If ColorWinKey = True Then chkColor(3).Value = 1
 'now check the hotkey
 For i = 0 To cmbColorCap.ListCount - 1
  If ColorHotKey = cmbColorCap.List(i) Then
   cmbColorCap.ListIndex = i
   Exit For
  End If
 Next
 
 'now set the hotkeys
 
  If ColorHotKey = "F1" Then PgmHotKey = 112
  If ColorHotKey = "F2" Then PgmHotKey = 113
  If ColorHotKey = "F3" Then PgmHotKey = 114
  If ColorHotKey = "F4" Then PgmHotKey = 115
  If ColorHotKey = "F5" Then PgmHotKey = 116
  If ColorHotKey = "F6" Then PgmHotKey = 117
  If ColorHotKey = "F7" Then PgmHotKey = 118
  If ColorHotKey = "F8" Then PgmHotKey = 119
  If ColorHotKey = "F9" Then PgmHotKey = 120
  If ColorHotKey = "F10" Then PgmHotKey = 121
  If ColorHotKey = "F11" Then PgmHotKey = 122
  If ColorHotKey = "F12" Then PgmHotKey = 123

If Len(ColorHotKey) = 1 Then PgmHotKey = Asc(ColorHotKey)

 
 frmMainClip.VBHotKey3.VKey = PgmHotKey
 frmMainClip.VBHotKey3.CtrlKey = ColorCtrlKey
 frmMainClip.VBHotKey3.ShiftKey = ColorShiftKey
 frmMainClip.VBHotKey3.AltKey = ColorAltKey
 frmMainClip.VBHotKey3.WinKey = ColorWinKey
 

End Sub


Sub LoadKeys()
 For i = 65 To 90
  cmbCapture.AddItem Chr$(i)
  cmbClip.AddItem Chr$(i)
  cmbColorCap.AddItem Chr$(i)
 Next
 
 For i = 1 To 12
  cmbCapture.AddItem "F" & Trim(Str$(i))
  cmbClip.AddItem "F" & Trim(Str$(i))
  cmbColorCap.AddItem "F" & Trim(Str$(i))
 Next
 
End Sub


Sub SetHotKeys()
 Dim PgmHotKey As Long
 
 'now set the Capture hotkeys
  If CapHotKey = "F1" Then PgmHotKey = 112
  If CapHotKey = "F2" Then PgmHotKey = 113
  If CapHotKey = "F3" Then PgmHotKey = 114
  If CapHotKey = "F4" Then PgmHotKey = 115
  If CapHotKey = "F5" Then PgmHotKey = 116
  If CapHotKey = "F6" Then PgmHotKey = 117
  If CapHotKey = "F7" Then PgmHotKey = 118
  If CapHotKey = "F8" Then PgmHotKey = 119
  If CapHotKey = "F9" Then PgmHotKey = 120
  If CapHotKey = "F10" Then PgmHotKey = 121
  If CapHotKey = "F11" Then PgmHotKey = 122
  If CapHotKey = "F12" Then PgmHotKey = 123
 
 If Len(CapHotKey) = 1 Then PgmHotKey = Asc(CapHotKey)
 
 frmMainClip.VBHotKey2.VKey = PgmHotKey
 frmMainClip.VBHotKey2.CtrlKey = CapCtrlKey
 frmMainClip.VBHotKey2.ShiftKey = CapShiftKey
 frmMainClip.VBHotKey2.AltKey = CapAltKey
 frmMainClip.VBHotKey2.WinKey = CapWinKey
'----------------------------------------------
  'now set the Clip hotkeys
 
  If ClipHotKey = "F1" Then PgmHotKey = 112
  If ClipHotKey = "F2" Then PgmHotKey = 113
  If ClipHotKey = "F3" Then PgmHotKey = 114
  If ClipHotKey = "F4" Then PgmHotKey = 115
  If ClipHotKey = "F5" Then PgmHotKey = 116
  If ClipHotKey = "F6" Then PgmHotKey = 117
  If ClipHotKey = "F7" Then PgmHotKey = 118
  If ClipHotKey = "F8" Then PgmHotKey = 119
  If ClipHotKey = "F9" Then PgmHotKey = 120
  If ClipHotKey = "F10" Then PgmHotKey = 121
  If ClipHotKey = "F11" Then PgmHotKey = 122
  If ClipHotKey = "F12" Then PgmHotKey = 123

If Len(ClipHotKey) = 1 Then PgmHotKey = Asc(ClipHotKey)

 
 frmMainClip.VBHotKey1.VKey = PgmHotKey
 frmMainClip.VBHotKey1.CtrlKey = ClipCtrlKey
 frmMainClip.VBHotKey1.ShiftKey = ClipShiftKey
 frmMainClip.VBHotKey1.AltKey = ClipAltKey
 frmMainClip.VBHotKey1.WinKey = ClipWinKey
 
'----------------------------------------------
  'now set the Color hotkeys
 
  If ColorHotKey = "F1" Then PgmHotKey = 112
  If ColorHotKey = "F2" Then PgmHotKey = 113
  If ColorHotKey = "F3" Then PgmHotKey = 114
  If ColorHotKey = "F4" Then PgmHotKey = 115
  If ColorHotKey = "F5" Then PgmHotKey = 116
  If ColorHotKey = "F6" Then PgmHotKey = 117
  If ColorHotKey = "F7" Then PgmHotKey = 118
  If ColorHotKey = "F8" Then PgmHotKey = 119
  If ColorHotKey = "F9" Then PgmHotKey = 120
  If ColorHotKey = "F10" Then PgmHotKey = 121
  If ColorHotKey = "F11" Then PgmHotKey = 122
  If ColorHotKey = "F12" Then PgmHotKey = 123

If Len(ColorHotKey) = 1 Then PgmHotKey = Asc(ColorHotKey)

 
 frmMainClip.VBHotKey3.VKey = PgmHotKey
 frmMainClip.VBHotKey3.CtrlKey = ColorCtrlKey
 frmMainClip.VBHotKey3.ShiftKey = ColorShiftKey
 frmMainClip.VBHotKey3.AltKey = ColorAltKey
 frmMainClip.VBHotKey3.WinKey = ColorWinKey

 
End Sub

Private Sub Command1_Click()
'Set capture keys
'===================================
 CapHotKey = cmbCapture.Text
 If chkCapture(0).Value = 1 Then CapCtrlKey = True Else CapCtrlKey = False
 If chkCapture(1).Value = 1 Then CapShiftKey = True Else CapShiftKey = False
 If chkCapture(2).Value = 1 Then CapAltKey = True Else CapAltKey = False
 If chkCapture(3).Value = 1 Then CapWinKey = True Else CapWinKey = False
 
 'Set Clip keys
 ClipHotKey = cmbClip.Text
 If chkClip(0).Value = 1 Then ClipCtrlKey = True Else ClipCtrlKey = False
 If chkClip(1).Value = 1 Then ClipShiftKey = True Else ClipShiftKey = False
 If chkClip(2).Value = 1 Then ClipAltKey = True Else ClipAltKey = False
 If chkClip(3).Value = 1 Then ClipWinKey = True Else ClipWinKey = False
 
 'Set Color keys
 ColorHotKey = cmbColorCap.Text
 If chkColor(0).Value = 1 Then ColorCtrlKey = True Else ColorCtrlKey = False
 If chkColor(1).Value = 1 Then ColorShiftKey = True Else ColorShiftKey = False
 If chkColor(2).Value = 1 Then ColorAltKey = True Else ColorAltKey = False
 If chkColor(3).Value = 1 Then ColorWinKey = True Else ColorWinKey = False
 
 SetHotKeys
 
 WriteOptions
 DoEvents
 
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()

 
 'load the combo boxes
 LoadKeys
 
 'Set the defaults on the window
 DisplayOptions
 

End Sub






