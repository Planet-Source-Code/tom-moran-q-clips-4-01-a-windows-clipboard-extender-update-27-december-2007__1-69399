VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainClip 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Q - Clips:  Main Clip List"
   ClientHeight    =   5085
   ClientLeft      =   150
   ClientTop       =   660
   ClientWidth     =   6075
   Icon            =   "frmMainClip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6075
   Begin prjQClips40.VBHotKey VBHotKey3 
      Left            =   6390
      Top             =   3195
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin prjQClips40.VBHotKey VBHotKey2 
      Left            =   6390
      Top             =   2520
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin prjQClips40.VBHotKey VBHotKey1 
      Left            =   6345
      Top             =   1935
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.PictureBox picClipLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   270
      ScaleHeight     =   945
      ScaleWidth      =   2355
      TabIndex        =   20
      Top             =   180
      Width           =   2355
      Begin VB.Label lblClipEmpty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "lblClipEmpty"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   135
         Width           =   2355
      End
   End
   Begin VB.CheckBox chkWarnings 
      Caption         =   "Show warnings"
      Height          =   255
      Left            =   4365
      TabIndex        =   18
      Top             =   3330
      Width           =   1410
   End
   Begin VB.CheckBox chkPasteMinimize 
      Caption         =   "Paste Minimize"
      Height          =   255
      Left            =   4365
      TabIndex        =   17
      Top             =   2970
      Width           =   1365
   End
   Begin VB.CheckBox chkStartMinimized 
      Caption         =   "Start Minimized"
      Height          =   255
      Left            =   4365
      TabIndex        =   16
      Top             =   2610
      Width           =   1395
   End
   Begin prjQClips40.CandyButton cmdCapture 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3780
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "       Capture"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmMainClip.frx":86EA
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "On Top"
      Height          =   255
      Left            =   2940
      TabIndex        =   11
      Top             =   3330
      Width           =   1080
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Sound On"
      Height          =   255
      Left            =   2940
      TabIndex        =   10
      Top             =   2970
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6255
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkAutoHover 
      Caption         =   "Auto Select Clip"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   4620
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkAutoSave 
      Caption         =   "Auto Save"
      Height          =   255
      Left            =   2940
      TabIndex        =   8
      Top             =   2610
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox TextClip 
      Height          =   795
      Index           =   0
      Left            =   7065
      TabIndex        =   7
      Text            =   "TextClip"
      Top             =   90
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picClip 
      Height          =   735
      Index           =   0
      Left            =   7065
      ScaleHeight     =   675
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   1035
      Visible         =   0   'False
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox RTFTextClip 
      Height          =   735
      Index           =   0
      Left            =   7155
      TabIndex        =   4
      Top             =   1980
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMainClip.frx":8C84
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   2940
      ScaleHeight     =   2055
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   125
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   65
         ScaleHeight     =   1935
         ScaleWidth      =   2595
         TabIndex        =   2
         Top             =   45
         Width           =   2595
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   1485
            Left            =   135
            ToolTipText     =   " Left Click to Open - Right Click to Open with... "
            Top             =   225
            Visible         =   0   'False
            Width           =   2190
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmMainClip.frx":8D0E
         ToolTipText     =   " Left Click to Open - Right Click to Open with... "
         Top             =   60
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6255
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClip.frx":8D14
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClip.frx":92AE
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClip.frx":9848
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainClip.frx":9DF2
            Key             =   "Files"
         EndProperty
      EndProperty
   End
   Begin prjQClips40.ucCoolList ucCoolList1 
      Height          =   4380
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7726
      ScrollBarWidth  =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   8388608
      FontSelected    =   0
      BackSelected    =   14737632
      BackSelectedG1  =   14737632
      BackSelectedG2  =   12632256
      HoverSelection  =   -1  'True
      ItemHeight      =   36
      ItemHeightAuto  =   0   'False
      ItemOffset      =   4
      ItemTextLeft    =   25
      SelectModeStyle =   2
      MousePointer    =   99
      MouseIcon       =   "frmMainClip.frx":A38C
   End
   Begin prjQClips40.CandyButton cmdUnHook 
      Height          =   465
      Index           =   0
      Left            =   3285
      TabIndex        =   13
      Top             =   4365
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "       Turn Q-Clips Capture Off"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4194304
      Picture         =   "frmMainClip.frx":AC66
      PictureAlignment=   2
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8421504
      ColorButtonUp   =   12632256
      ColorButtonDown =   14737632
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin prjQClips40.CandyButton cmdPane 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   14
      Top             =   4560
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmMainClip.frx":B210
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin prjQClips40.CandyButton cmdPane 
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Top             =   4560
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmMainClip.frx":B7BA
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin prjQClips40.CandyButton cmdColorCapture 
      Height          =   375
      Left            =   4410
      TabIndex        =   19
      Top             =   3780
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "       Color Pick"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmMainClip.frx":BD64
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin prjQClips40.CandyButton cmdUnHook 
      Height          =   465
      Index           =   1
      Left            =   3285
      TabIndex        =   22
      Top             =   4365
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     Turn Q-Clips Capture On"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmMainClip.frx":C2FE
      PictureAlignment=   2
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   160
      ColorButtonUp   =   128
      ColorButtonDown =   240
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin VB.Image imgPane 
      Height          =   240
      Index           =   1
      Left            =   7860
      Picture         =   "frmMainClip.frx":C8A8
      Top             =   2835
      Width           =   240
   End
   Begin VB.Image imgPane 
      Height          =   240
      Index           =   0
      Left            =   7425
      Picture         =   "frmMainClip.frx":CE42
      Top             =   2835
      Width           =   240
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing on Clipboard"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2940
      TabIndex        =   3
      Top             =   2250
      Width           =   2715
   End
   Begin VB.Image imgHandCursor 
      Height          =   480
      Left            =   6300
      Picture         =   "frmMainClip.frx":D3DC
      Top             =   1305
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Custom Clip List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewClips 
         Caption         =   "Open Main Q-Clip List"
         Shortcut        =   ^Z
      End
      Begin VB.Menu xxxmenuF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveClip 
         Caption         =   "Save Custom Clip List As..."
      End
      Begin VB.Menu xxxmenuF2 
         Caption         =   "-"
      End
      Begin VB.Menu xxmnuRFiles 
         Caption         =   "Recent Custom Clip Lists"
         Begin VB.Menu mnuRecentFiles 
            Caption         =   " Do Not Show"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   " (Empty)"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles8"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu xxxmenuF3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Shutdown Q-Clips"
      End
   End
   Begin VB.Menu mnuEditMenu 
      Caption         =   "Edit"
      Begin VB.Menu mnuPasteDirect 
         Caption         =   "Paste to Program"
      End
      Begin VB.Menu mnuPasteIndirect 
         Caption         =   "Paste to Clipboard"
      End
      Begin VB.Menu xxxEdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKillClip 
         Caption         =   "Delete Selected"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuKillList 
         Caption         =   "Delete All Clips"
      End
      Begin VB.Menu xxxmnuEdit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveClipAs 
         Caption         =   "Save Clip As..."
      End
      Begin VB.Menu xxxEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenClip 
         Caption         =   "Open Clip"
         Index           =   0
      End
      Begin VB.Menu mnuOpenClip 
         Caption         =   "Open with..."
         Index           =   1
      End
      Begin VB.Menu xxxEdit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearWinClip 
         Caption         =   "Clear Windows Clipboard"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOption 
         Caption         =   "Q-Clipboard On"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Multi-Select Clips"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOption 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Set Hot Keys"
         Index           =   3
      End
      Begin VB.Menu mnuOption 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCapOption 
         Caption         =   "Set Screen Capture Type"
         Begin VB.Menu mnuScreenCap 
            Caption         =   "Capture Region"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuScreenCap 
            Caption         =   "Capture Desktop"
            Index           =   1
         End
         Begin VB.Menu mnuScreenCap 
            Caption         =   "Capture Active Window"
            Index           =   2
         End
      End
      Begin VB.Menu mnuColorCapture 
         Caption         =   "Set Color Capture Type"
         Begin VB.Menu mnuColorCapType 
            Caption         =   "Long Color Number"
            Index           =   0
         End
         Begin VB.Menu mnuColorCapType 
            Caption         =   "RGB Colors"
            Index           =   1
         End
         Begin VB.Menu mnuColorCapType 
            Caption         =   "Hex Color"
            Index           =   2
         End
         Begin VB.Menu mnuColorCapType 
            Caption         =   "All Color Numbers"
            Checked         =   -1  'True
            Index           =   3
         End
      End
      Begin VB.Menu mnuIType 
         Caption         =   "Set Image Type"
         Begin VB.Menu mnuImageType 
            Caption         =   "Bitmap (.bmp)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuImageType 
            Caption         =   "JPEG (.jpg)"
            Index           =   1
         End
      End
      Begin VB.Menu xxxmnuSepO2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunOnStart 
         Caption         =   "Run on Windows Start-Up"
      End
      Begin VB.Menu mnuTheme 
         Caption         =   "Set Theme"
         Begin VB.Menu mnuSetTheme 
            Caption         =   "Blue"
            Index           =   0
         End
         Begin VB.Menu mnuSetTheme 
            Caption         =   "Silver"
            Index           =   1
         End
         Begin VB.Menu mnuSetTheme 
            Caption         =   "Black"
            Index           =   2
         End
         Begin VB.Menu mnuSetTheme 
            Caption         =   "Olive"
            Index           =   3
         End
         Begin VB.Menu mnuSetTheme 
            Caption         =   "Windows Default"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuQHelp 
         Caption         =   "Q-Clips Help Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu xxxHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Q-Clips"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuNothing 
         Caption         =   "Show Q-Clips"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNothing1 
         Caption         =   "Capture Selected Region"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNothing2 
         Caption         =   "Capture Color Number"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMainClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'type for system tray
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
    
'Menu API constants for popup menu
Private Const MF_STRING As Long = &H0&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_BYCOMMAND As Long = &H0&
Private Const TPM_RETURNCMD As Long = &H100&

'Internal menu constants for popup menu

Private Const ID_TOGGLEON As Long = &H6005&
Private Const ID_CANCEL As Long = &H6000&
Private Const ID_SEPERATOR As Long = &H6001
Private Const ID_SHOWQ As Long = &H6002&
Private Const ID_CAPIMAGE As Long = &H6003&
Private Const ID_CAPCOLOR As Long = &H6007&
Private Const ID_EXIT As Long = &H6004&
Private Const ID_HOTKEY As Long = &H6006&
Const MF_CHECKED As Long = &H8&
Const MF_UNCHECKED As Long = &H0&

'Popup menu variable
Private m_hPopup As Long

'Functions used for popup menu
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpNewItem$) As Long
Private Declare Function ClientToScreen& Lib "user32" (ByVal hwnd&, lpPoint As POINTAPI)
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu&, ByVal wFlags&, ByVal x&, ByVal y&, ByVal nReserved&, ByVal hwnd&, ByVal lpRect&) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function CheckMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long


'System tray icon constants
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLICK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = 1&

'System tray variable
Private Notify As NOTIFYICONDATA

'Function used for system tray
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Dim Loading As Boolean

Private Declare Function Sleep Lib "kernel32" _
  (ByVal dwMilliseconds As Long) As Long

Private Sub ShowColor(BkColor As String)
 On Error Resume Next
 
 Dim RColor As Long
 Dim GColor As Long
 Dim BColor As Long
 Dim pos As Integer
 Dim pos2 As Integer
 
 
 BkColor = Trim(BkColor)
 
'Check to see if RGB or All Colors
 If Left$(BkColor, 2) = "R:" Then
  
  'this will check if the clip has All Colors
  pos = InStr(1, BkColor, "Long Color:")
   
   If pos Then
    RColor = Val(Mid$(BkColor, pos + 11, Len(BkColor))) ' - (pos + 12)))
    Text1.BackColor = RColor
    Exit Sub
   End If
   
  'this code will parse RGB Colors
  pos2 = InStr(1, BkColor, "G:")
   RColor = Val(Mid$(BkColor, 3, pos2 - 1))
   
  pos = InStr(1, BkColor, "B:")
   GColor = Val(Mid$(BkColor, pos2 + 2, pos - 1))
   BColor = Val(RTrim(Mid$(BkColor, pos + 2, Len(BkColor))))
   Text1.BackColor = RGB(RColor, GColor, BColor)
  
  Exit Sub
  
 End If
   
 'This is to check if color number is Hex
 If Left$(BkColor, 1) = "&" Then
  Text1.BackColor = BkColor
  Exit Sub
 End If
 
 'Finally, check to see if long number
 If Left$(BkColor, 1) < Chr$(48) Or Left$(BkColor, 1) > Chr$(57) Then Exit Sub
 
 Text1.BackColor = Val(BkColor)
 
End Sub


Private Sub CapWindow()

End Sub


Private Sub ClearAll()
 Dim indx As Integer
 Dim TempText As String
 TempText = ListTag(indx)
 Busy = True
 
 On Error GoTo Errhandler
 
 'find out what we are deleting
 For indx = 0 To ucCoolList1.ListCount - 1
 
  j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
  TempText = ListTag(indx)
 
   'If j is 0 then this is parent control and can't remove that
   If j > 0 Then
   'delete the control
     If Left$(TempText, 3) = "PIC" Then
      Unload picClip(j)
     ElseIf Left$(TempText, 3) = "TXT" Then
      Unload TextClip(j)
     ElseIf Left$(TempText, 3) = "RTF" Then
      Unload RTFTextClip(j)
     End If
   End If
    
   'delete the ListTag
   ListTag(indx) = ""
   
 Next indx
 
 ucCoolList1.Clear
 DoEvents
 TextCount = 0
 PicCount = 0
 RTFTextCount = 0
 Text1.Visible = False
 Picture1.Visible = False
 lblInfo.Caption = "Nothing on Clipboard"
 Busy = False
 Exit Sub

Errhandler:
  
  MsgBox Error$, vbApplicationModal + vbOKOnly, "Q-Clip Error!"
  Resume Next
  
End Sub


Private Sub CreateMenu()
  
  'Create popup menu
  m_hPopup = CreatePopupMenu()
  Call AppendMenu(m_hPopup, MF_STRING, ID_SHOWQ, ByVal "Show Q-Clips")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_TOGGLEON, ByVal "Q-Clipboard On")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_CAPIMAGE, ByVal "Capture Selected Region")
  Call AppendMenu(m_hPopup, MF_STRING, ID_CAPCOLOR, ByVal "Capture Color Number")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_HOTKEY, ByVal "Set Hot Keys")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_CANCEL, ByVal "Cancel")
  Call AppendMenu(m_hPopup, MF_STRING Or MF_SEPARATOR, ID_SEPERATOR, ByVal vbNullString)
  Call AppendMenu(m_hPopup, MF_STRING, ID_EXIT, ByVal "Shutdown Q-Clips")
  
  'this API for checkmark on menu item
  If ClipboardOn = True Then
   Call CheckMenuItem(m_hPopup, ID_TOGGLEON, &H8)
  Else
   Call CheckMenuItem(m_hPopup, ID_TOGGLEON, &H0)
  End If
  
  'Bold the first menu item (Show)
  Call SetMenuDefaultItem(m_hPopup, 0, 1&)
  
End Sub
Private Sub CustomSaveAs()
    
    On Error Resume Next
    CMDialog1.CancelError = True
    CMDialog1.Filter = "Q-Clip Files (*.qcl)|*.qcl|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CMDialog1.FilterIndex = 1
    CMDialog1.FLAGS = cdlOFNOverwritePrompt
    CMDialog1.DialogTitle = "Save Custom Q-Clip List As..."
    If Right$(QFileName, 3) <> "qcx" Then CMDialog1.FileName = QFileName
    CMDialog1.ShowSave
    If Err = 32755 Then   ' User chose Cancel.
       Exit Sub
    Else
    
     QFileName = CMDialog1.FileName
     
     If QFileName = App.Path & "\qcliplist.qcx" Then
      iret = MsgBox("This name is reserved. Try Again.", vbOKOnly + vbApplicationModal, "Q-Clip Error")
       Exit Sub
      Else
       Call CustomSave(QFileName)
      End If
     End If
    
End Sub

Public Function LastPart(Text As String) As String
 
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  LastPart = Temp
 Else
  LastPart = Mid$(Temp, i + 1)
 End If
 
End Function
Private Sub AutoLoad()
 Dim TempText As String
 Dim FileNum As Integer
 Dim TempNum As Integer
 Dim There As Boolean
 Dim IsFileList As Boolean
 
 On Error Resume Next
 
 There = Exist(QFileName)
 If Not There Then Exit Sub
 
 'if nothing in clipboard then tell 'em
 If ucCoolList1.ListCount < 1 Then
  picClipLabel.Visible = True
 Else
  picClipLabel.Visible = False
 End If
 
 If FileLen(QFileName) = 0 Then Exit Sub
 
 Screen.MousePointer = 11
 
 FileNum = FreeFile
 PicCount = 0
 
 
 Open QFileName For Input As #FileNum
 
  While Not EOF(FileNum)
 
   Line Input #FileNum, TempText
  
    If UCase(Right$(TempText, 1)) = "P" Or UCase(Right$(TempText, 1)) = "G" Then 'it's a bitmap
  
       ListTag(ucCoolList1.ListCount) = "PIC" & Str(PicCount)
       
       Call ucCoolList1.AddItem("<Graphic Image Clip>", 2, 2)
       
       DoEvents
       
       If PicCount > 0 Then
        Load frmMainClip.picClip(PicCount)
       End If
       
       frmMainClip.Image1.Stretch = False
       frmMainClip.picClip(PicCount).Picture = LoadPicture(TempText)
       frmMainClip.Image1.Picture = frmMainClip.picClip(PicCount).Picture
       PicCount = PicCount + 1
    End If
    
    If UCase(Right$(TempText, 1)) = "T" Then 'it's a text file
        'See if a file list
        IsFileList = InStr(1, TempText, "FIL")
        
        If IsFileList = True Then
         ListTag(ucCoolList1.ListCount) = "FIL" & Str(TextCount)
        Else
         ListTag(ucCoolList1.ListCount) = "TXT" & Str(TextCount)
        End If
        
        'firsttime through use main else load a new control
        If TextCount > 0 Then
         Load frmMainClip.TextClip(TextCount)
        End If
        
        TempNum = FreeFile
        Open TempText For Input As #TempNum
     
        'load the text to the control
        TextClip(TextCount).Text = Input(LOF(TempNum), TempNum)
        Close #TempNum
        
        'parse out the first 40 characters for display in list box
        If Len(TextClip(TextCount).Text) > 40 Then
         TempText = Left$(TextClip(TextCount).Text, 40) & "..."
        Else
         TempText = TextClip(TextCount).Text
        End If
        
        'Add to list
        If IsFileList = True Then
         Call ucCoolList1.AddItem(TempText, 4, 4)
        Else
         Call ucCoolList1.AddItem(TempText, 1, 1)
        End If
        
        'increase the count for next text clip
        TextCount = TextCount + 1
     
     End If 'Text
     
      If UCase(Right$(TempText, 1)) = "F" Then 'it's a RTF file
        ListTag(ucCoolList1.ListCount) = "RTF" & Str(RTFTextCount)
        
        'firsttime through use main control else load a new control
        If RTFTextCount > 0 Then
         Load RTFTextClip(RTFTextCount)
        End If
        
        'load the RTF Text to the control
        RTFTextClip(RTFTextCount).LoadFile TempText, rtfRTF
               
        'parse out the first 40 characters for display in list box
        If Len(RTFTextClip(RTFTextCount).Text) > 40 Then
         TempText = Left$(frmMainClip.RTFTextClip(RTFTextCount).Text, 40) & "..."
        Else
         TempText = frmMainClip.RTFTextClip(RTFTextCount).Text
        End If
        'add to list
        Call ucCoolList1.AddItem(TempText, 3, 3)
        
        'increase the count for next text clip
        RTFTextCount = RTFTextCount + 1
     End If
     
 Wend
 Close FileNum
 
  'if nothing in clipboard then tell 'em
 If ucCoolList1.ListCount < 1 Then
  picClipLabel.Visible = True
 Else
  picClipLabel.Visible = False
 End If
 
 Screen.MousePointer = 0
 ucCoolList1.ListIndex = 0
 Me.Refresh
 DoEvents
 
 'Show max warning if showwarnings is true
  If frmMainClip.ucCoolList1.ListCount = 25 And ShowWarnings = True And Warning25 = True Then
   Warning25 = False
   If frmMainClip.Visible = False Then frmMainClip.Visible = True
   TempText = "Warning: You have 25 clips in this Q-Clips Collection." & vbCrLf
   TempText = TempText & "The next clipboard capture will delete the oldest (1st) clip in this collection."
   MsgBox TempText, vbExclamation + vbOKOnly + vbApplicationModal, "Max Q-Clips In Collection"
  End If
 
End Sub


Private Sub AutoSave()
 Dim indx As Integer
 Dim TempText As String
 Dim FileNum As Integer
 Dim TempNum As Integer
 Dim ClipFilename As String
 Dim nStat As Long
 Dim strcontents As String
 Dim NewFileName As String
 Dim MyPic As StdPicture
 
 iret = InStr(QFileName, ".")
 NewFileName = Left$(QFileName, iret - 1)
 
 On Error Resume Next
 Screen.MousePointer = 11
 FileNum = FreeFile
 
 Open QFileName For Output As #FileNum
 
 For indx = 0 To ucCoolList1.ListCount - 1
   
   TempText = ListTag(indx)
   j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
   
   If Left$(TempText, 3) = "PIC" Then 'save as jpg bitmap
    
   ' picClip(j).Picture = picClip(j).Image
    
      If ImageType = 1 Then 'it's a jpeg
       ClipFilename = NewFileName & "PIC" & LTrim(Str(indx)) & ".JPG"
       SavePicture picClip(j), "$$_temp.bmp"
       DoEvents
       Set MyPic = LoadPicture("$$_temp.bmp")
       Set m_Image = New cImage
       m_Image.CopyStdPicture MyPic
       Call SaveImage(m_Image, ClipFilename)
       Set MyPic = Nothing
       Kill "$$_temp.bmp"
      Else
       ClipFilename = NewFileName & "PIC" & LTrim(Str(indx)) & ".BMP"
       SavePicture picClip(j), ClipFilename
      End If
      
     Print #FileNum, ClipFilename
     
   End If
 
   If Left$(TempText, 3) = "TXT" Then
      ClipFilename = NewFileName & "TXT" & LTrim(Str(indx)) & ".TXT"
      TempNum = FreeFile
       ' Open the file.
      Open ClipFilename For Output As #TempNum
       ' Place the contents into a variable.
       strcontents = TextClip(j).Text
       ' Write the variable contents to a saved file.
       Print #TempNum, strcontents
      Close #TempNum
      Print #FileNum, ClipFilename
    End If
 
    If Left$(TempText, 3) = "RTF" Then 'save as rich text
       ClipFilename = NewFileName & "RTF" & LTrim(Str(indx)) & ".RTF"
       RTFTextClip(j).SaveFile ClipFilename, rtfRTF
       Print #FileNum, ClipFilename
    End If
    
   If Left$(TempText, 3) = "FIL" Then
      ClipFilename = NewFileName & "FIL" & LTrim(Str(indx)) & ".TXT"
      TempNum = FreeFile
       ' Open the file.
      Open ClipFilename For Output As #TempNum
       ' Place the contents into a variable.
       strcontents = TextClip(j).Text
       ' Write the variable contents to a saved file.
       Print #TempNum, strcontents
      Close #TempNum
      Print #FileNum, ClipFilename
    End If

  Next indx
  
  Close #FileNum
  
  Screen.MousePointer = 0

End Sub


Private Sub CustomSave(QFileName As String)
 
 Dim indx As Integer
 Dim TempText As String
 Dim FileNum As Integer
 Dim TempNum As Integer
 Dim ClipFilename As String
 Dim NewFileName As String
 Dim strcontents As String
 Dim nStat As Long
 Dim MyPic As StdPicture
 
 If QFileName = "" Then Exit Sub
 iret = InStr(QFileName, ".")
 NewFileName = Left$(QFileName, iret - 1)
 
 
 On Error Resume Next
 Screen.MousePointer = 11
 FileNum = FreeFile
 
 Open QFileName For Output As #FileNum
 
 For indx = 0 To ucCoolList1.ListCount - 1
   
   TempText = ListTag(indx)
   j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
   
   If Left$(TempText, 3) = "PIC" Then
   
      If ImageType = 1 Then 'it's a jpeg
       ClipFilename = NewFileName & "PIC" & LTrim(Str(indx)) & ".JPG"
       SavePicture picClip(j), "$$_temp.bmp"
       DoEvents
       Set MyPic = LoadPicture("$$_temp.bmp")
       Set m_Image = New cImage
       m_Image.CopyStdPicture MyPic
       Call SaveImage(m_Image, ClipFilename)
       Set MyPic = Nothing
       Kill "$$_temp.bmp"
      Else
       ClipFilename = NewFileName & "PIC" & LTrim(Str(indx)) & ".BMP"
       SavePicture picClip(j), ClipFilename
      End If
      
     Print #FileNum, ClipFilename
     
   End If
 
   If Left$(TempText, 3) = "TXT" Then
      ClipFilename = NewFileName & "TXT" & LTrim(Str(indx)) & ".TXT"
      TempNum = FreeFile
       ' Open the file.
      Open ClipFilename For Output As #TempNum
       ' Place the contents into a variable.
       strcontents = TextClip(j).Text
       ' Write the variable contents to a saved file.
       Print #TempNum, strcontents
      Close #TempNum
      Print #FileNum, ClipFilename
    End If
 
    If Left$(TempText, 3) = "RTF" Then
       ClipFilename = NewFileName & "RTF" & LTrim(Str(indx)) & ".RTF"
       RTFTextClip(j).SaveFile ClipFilename, rtfRTF
       Print #FileNum, ClipFilename
    End If

  Next indx
  
  Close #FileNum
  
  Me.Caption = "Q - Clips:  " & LastPart(QFileName)
  UpDateFileMenu QFileName
  
  Screen.MousePointer = 0

End Sub

Private Sub DeleteAll()
 
 Dim indx As Integer
 Dim TempText As String
 TempText = ListTag(indx)
 Busy = True
 
 On Error Resume Next
 
 If iQSound Then PlayWaveSound App.Path & "\deleteall.wav"
 
 'find out what we are deleting
 For indx = 0 To ucCoolList1.ListCount - 1
 
  j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
  TempText = ListTag(indx)
 
   'If j is 0 then this is parent control and can't remove that
   If j > 0 Then
   'delete the control
     If Left$(TempText, 3) = "PIC" Then
      Unload picClip(j)
     ElseIf Left$(TempText, 3) = "TXT" Then
      Unload TextClip(j)
     ElseIf Left$(TempText, 3) = "RTF" Then
      Unload RTFTextClip(j)
     End If
   End If
    
   'delete the ListTag
   ListTag(indx) = ""
   
 Next indx
 
 'destroy clips
 TempText = Left$(QFileName, Len(QFileName) - 4)
 Kill TempText & "*.*"
 
 'reset counters
 ucCoolList1.Clear
 TextCount = 0
 PicCount = 0
 RTFTextCount = 0
 Text1.Visible = False
 Picture1.Visible = False
 lblInfo.Caption = "Nothing on Clipboard"
 picClipLabel.Visible = True
 Busy = False
 
 Exit Sub

Errhandler:
  MsgBox Error$, vbCritical + vbApplicationModal + vbOKOnly, "QClips Error!"
  Resume Next
End Sub

Sub DeleteClip()
 
 Dim indx As Integer
 Dim TempText As String
 TempText = ListTag(indx)
 Busy = True
 
 On Error GoTo Errhandler
 
 'find out what we are deleting
 indx = ucCoolList1.ListIndex
 If indx < 0 Then Exit Sub
 j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
 TempText = ListTag(indx)
 
  'If j is 0 then this is parent control and can't remove that
 If j > 0 Then
 'delete the control
   If Left$(TempText, 3) = "PIC" Then
    Unload picClip(j)
   ElseIf Left$(TempText, 3) = "TXT" Then
    Unload TextClip(j)
   ElseIf Left$(TempText, 3) = "RTF" Then
    Unload RTFTextClip(j)
   End If
 End If
  
 'now remove listtag from list and item
 For i = indx To ucCoolList1.ListCount - 1
  ListTag(i) = ListTag(i + 1)
 Next
 If indx = 0 Then
   Text1.Visible = False
   Picture1.Visible = False
   lblInfo.Caption = "Nothing on Clipboard"
 End If
 ucCoolList1.RemoveItem indx
 
 If indx > -1 Then
  ucCoolList1.ListIndex = ucCoolList1.ListCount - 1
 End If
 
  'if nothing in clipboard then tell 'em
 If ucCoolList1.ListCount < 1 Then
  picClipLabel.Visible = True
 Else
  picClipLabel.Visible = False
 End If
 
  'if nothing in clipboard then tell 'em
 If ucCoolList1.ListCount < 1 Then
  picClipLabel.Visible = True
 Else
  picClipLabel.Visible = False
 End If
 
 ucCoolList1.SetFocus
 Busy = False
 
 If iQSound Then PlayWaveSound App.Path & "\delete.wav"
 
 Exit Sub

Errhandler:
  MsgBox Error$, vbApplicationModal + vbOKOnly, "QClips Error!"
  Resume Next
  
End Sub


Private Sub PasteDirect()

'This is code used to paste the selected
'Q-Clip into the active program window

 Dim Start As Long
  
  If ucCoolList1.ListCount > 0 Then
   PutClip 'copy selected clip to windows clipboard
  Else
   Exit Sub
  End If
  
  If curWindow = 0 Then Exit Sub 'No hotkey pressed
  
  iret = BringWindowToTop(ParentWindow)
  
  iret = SetForegroundWindow(ParentWindow)
  SetActiveWindow curWindow
  
  Putfocus curWindow
  DoEvents
  SetFocusA curWindow
  
  'slow things down abit
  Start = Timer + 0.2
   While Start > Timer
    DoEvents
   Wend
   
   'using API keybd event because Vista does
   'not interpret SendKeys command correctly
      
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event VK_V, 0, 0, 0   ' press Ctrl-V - Paste keys
    DoEvents

    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0   'Release
    keybd_event VK_V, 0, KEYEVENTF_KEYUP, 0   'Release
    DoEvents

    iret = SetForegroundWindow(ParentWindow)
    Putfocus curWindow

    DoEvents
    
    If PasteMinimized = True Then Me.Hide
    
End Sub

Sub PutClip()
    
'Put selected item in windows clipboard

  Dim indx As Integer
  Dim TempText As String
  
  On Error Resume Next
  
  indx = ucCoolList1.ListIndex
  If indx < 0 Then indx = 0
  TempText = ListTag(indx)
  JustPasted = True 'this is so we don't copy to Q-Clips again
  Busy = True


   If Left$(TempText, 3) = "PIC" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       Clipboard.SetData picClip(j).Picture, vbCFBitmap
   ElseIf Left$(TempText, 3) = "FIL" Then 'its a list of files to set back to clipboard
     Dim next_file_name As String
     Dim file_names() As String
     Dim num_file_names As Integer
     Dim pos As Integer
     Dim txt As String

      ' Make an array of file names.
      txt = Trim$(Text1.Text)
      num_file_names = 0
      Do While Len(txt) > 0
        ' Get the next file name.
        pos = InStr(txt, vbCrLf)
        If pos = 0 Then
            next_file_name = Trim$(txt)
            txt = ""
        Else
            next_file_name = Trim$(Left$(txt, pos - 1))
            txt = Trim$(Mid$(txt, pos + Len(vbCrLf)))
        End If

        If Len(next_file_name) > 0 Then
            ' Make room for the next file name.
            num_file_names = num_file_names + 1
            ReDim Preserve file_names(1 To num_file_names)
            file_names(num_file_names) = next_file_name
        End If
       Loop
       
       ClipboardSetFiles file_names

   ElseIf Left$(TempText, 3) = "TXT" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       Clipboard.SetText TextClip(j).Text, vbCFText
   ElseIf Left$(TempText, 3) = "RTF" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       RTFTextClip(j).SelStart = 0
       RTFTextClip(j).SelLength = Len(RTFTextClip(j).TextRTF)
       SendMessage RTFTextClip(j).hwnd, WM_COPY, 0&, 0& 'Copy
   End If
   
   DoEvents
   
   If iQSound Then PlayWaveSound App.Path & "\copyclip.wav"
   
   Picture3.SetFocus
   JustPasted = False
   Busy = False
   
   
End Sub


Sub Reset_Image()

'This is to create thumbnail picture

  Busy = True
  
    Image1.Visible = False
    Image1.Stretch = False
    
    If Image1.Picture Then
      Image1.Height = Image1.Picture.Height
      Image1.Width = Image1.Picture.Width
        
      'is picture smaller than thumbnail?
        If Image1.Picture.Height < Picture1.Height And Image1.Picture.Width < Picture1.Width Then
         Image1.Top = (Picture1.Height - Image1.Picture.Height) / 2 + 150
         Image1.Left = (Picture1.Width - Image1.Picture.Width) / 2 + 200
         Image1.Visible = True
         Busy = False
         Exit Sub
        End If
    
      Image1.Stretch = True

       If Image1.Picture.Height >= Image1.Picture.Width Then
            Image1.Height = Picture1.Height
            Image1.Width = Image1.Width / (Image1.Picture.Height / Image1.Height)
             If Image1.Width > Picture1.Width Then
                Image1.Width = Picture1.Width
                Image1.Height = Image1.Picture.Height / (Image1.Picture.Width / Image1.Width)
             End If
       End If


        If Image1.Picture.Width > Image1.Picture.Height Then
            Image1.Width = Picture1.Width
            Image1.Height = Image1.Height / (Image1.Picture.Width / Image1.Width)
             If Image1.Height > Picture1.Height Then
                Image1.Height = Picture1.Height
                Image1.Width = Image1.Picture.Width / (Image1.Picture.Height / Image1.Height)
             End If
        End If
        
        Image1.Left = (Picture1.Width / 2) - (Image1.Width / 2)
        Image1.Top = (Picture1.Height / 2) - (Image1.Height / 2)
        Image1.Visible = True
        
   End If
 
 Busy = False
 
End Sub




Private Sub SetHotKeyOption()
   'no capturing while setting hotkey
    Dim OldClipboard As Boolean
    OldClipboard = ClipboardOn
 
    ClipboardOn = False
 
    'can't be on top while we set hotkeys
    If IsOnTop = True Then
     iret = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If

    frmHotKey.Show 1
       
    'okay to capture now
    ClipboardOn = OldClipboard
 
    'if option was to be ontop then reset
    If IsOnTop = True Then
     iret = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Public Sub SetHotKey()
Dim PgmHotKey As Long
 
 'now set the Capture hotkeys
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
 
  'now set the Cap hotkeys
 
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

Private Sub chkAutoHover_Click()
 
  ucCoolList1.HoverSelection = Not ucCoolList1.HoverSelection
 
  HoverOn = ucCoolList1.HoverSelection
  
  If Loading = False Then ucCoolList1.SetFocus
  
End Sub

Private Sub chkAutoSave_Click()

 If chkAutoSave.Value = False Then
  iQSave = False
 Else
  iQSave = True
 End If
 
End Sub



Private Sub chkOnTop_Click()

  If chkOnTop.Value = False Then
  
    iret = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    IsOnTop = False
      
  Else
  
    iret = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    IsOnTop = True
        
  End If
    
End Sub

Private Sub chkPasteMinimize_Click()

 
 If chkPasteMinimize.Value = False Then
  PasteMinimized = False
 Else
  PasteMinimized = True
 End If
End Sub

Private Sub chkSound_Click()

 If chkSound.Value = False Then
  iQSound = False
 Else
  iQSound = True
 End If
 
End Sub

Private Sub chkStartMinimized_Click()
 If chkStartMinimized.Value = False Then
  StartMinimized = False
 Else
  StartMinimized = True
 End If
End Sub

Private Sub chkWarnings_Click()

 If chkWarnings.Value = False Then
  ShowWarnings = False
 Else
  ShowWarnings = True
 End If
 
End Sub

Private Sub cmdCapture_Click()
 Dim Start As Long
 
 Me.Hide
 DoEvents
 
 'Delay to let other windows repaint
 If CaptureOption > 0 Then
  Screen.MousePointer = 11
  Call Sleep(250)
 End If
 
 Start = Timer + 0.25
 While Start > Timer
  DoEvents
 Wend
 
 Select Case CaptureOption
 
  Case 0 'capture region
   CapButton = True
   frmCapture.Show
  
  Case 1 'capture desktop
   keybd_event vbKeySnapshot, 0&, 0&, 0&
  
  Case 2 'capture active window
   keybd_event vbKeySnapshot, &H1&, 0&, 0&
     
 End Select

 If CaptureOption > 0 Then
  Call Sleep(250)
    Start = Timer + 0.5
    While Start > Timer
     DoEvents
    Wend
    Screen.MousePointer = 0
    Me.Show
 End If
 
End Sub



Private Sub cmdColorCapture_Click()
 Dim Start As Long
 Me.Hide
 DoEvents
 
'Delays to let screen repaint
  Screen.MousePointer = 11
  Call Sleep(250)

 Start = Timer + 0.25
 While Start > Timer
  DoEvents
 Wend
 
 Screen.MousePointer = 0
 frmGetColor.Show 1

 Me.Show
 
End Sub


Private Sub cmdPane_Click(Index As Integer)

 Select Case Index
 
  Case 0
   
   cmdPane(1).Visible = True
   cmdPane(0).Visible = False
   frmMainClip.Width = 2960
   ViewPane = False
   
  Case 1
  
   cmdPane(0).Visible = True
   cmdPane(1).Visible = False
   frmMainClip.Width = 6000
   ViewPane = True
   
   
 End Select
 
   
End Sub


Private Sub cmdSetHotKey_Click()


 
End Sub

Private Sub cmdUnhook_Click(Index As Integer)
  
 If iQSound Then PlayWaveSound App.Path & "\clickerx.wav"
 
 Select Case Index
  
  Case 0
  
    cmdUnHook(0).Visible = False
    cmdUnHook(1).Visible = True
    mnuOption(0).Checked = False
    Call CheckMenuItem(m_hPopup, ID_TOGGLEON, &H0)
    ClipboardOn = False
      'Unhook the form
     Call ChangeClipboardChain(frmMainClip.hwnd, m_hWndNext)
     UnHookForm frmMainClip
    If frmMainClip.Visible = True Then ucCoolList1.SetFocus
  Case 1
 
    cmdUnHook(0).Visible = True
    cmdUnHook(1).Visible = False
    mnuOption(0).Checked = True
    Call CheckMenuItem(m_hPopup, ID_TOGGLEON, &H8)
    ClipboardOn = True
   'hook it back up
    HookForm frmMainClip
    'Register this form as a Clipboardviewer
    m_hWndNext = SetClipboardViewer(frmMainClip.hwnd)
    DoEvents
    If frmMainClip.Visible = True Then ucCoolList1.SetFocus
 End Select
 

 
End Sub

Private Sub Form_Initialize()
 
 'for XP Theme
 InitCommonControlsXP

End Sub

Private Sub Form_Load()
 
  Me.Height = 5865
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
 Me.AutoRedraw = True
 Me.DrawWidth = Me.ScaleHeight / 128
  Loading = True

    With ucCoolList1
            '-- Initialize ImageList
        Call .SetImageList(ImageList1)
        
        '-- Load back picture and <Hand> pointer
        'Set list icon
       Set .MouseIcon = imgHandCursor.Picture
    End With
 
 'this is Main List filename - it never changes
 QFileName = App.Path & "\qcliplist.qcx"
 
 'Set Default Options
 Themes = 0 'Blue
 iQSound = True
 iQSave = True
 ViewPane = True
 IsOnTop = False
 HoverOn = True
 ShowWarnings = True
 StartMinimized = False
 PasteMinimized = False
 ClipboardOn = True
 RunOnStartUp = False
 ImageType = 0 '.bmp image
 CaptureOption = 0 'Region
 ColorCapType = 3 'All color types
 
 lblClipEmpty.Caption = "Clipboard is empty." & vbCrLf & "Copy or cut to collect items."
 
 
 'Set the hotkey default - F9
  'set program defaults
 ClipHotKey = "Q"
 ClipCtrlKey = True
 ClipShiftKey = False
 ClipAltKey = False
 ClipWinKey = False
 CapHotKey = "F9"
 CapCtrlKey = False
 CapShiftKey = False
 CapAltKey = False
 CapWinKey = False
 ColorHotKey = "F8"
 ColorCtrlKey = False
 ColorShiftKey = False
 ColorAltKey = False
 CapWinKey = False

  'Get User Options
  GetRecentFiles 'from ini file
  GetOptions
  SetOptions
  SetHotKey
  
  'Set the menu check mark on saved defaults
  mnuScreenCap_Click (CaptureOption)
  mnuColorCapType_Click (ColorCapType)
 
  'Load the default clip list
  QFileName = App.Path & "\qcliplist.qcx"
  Me.Caption = "Q - Clips:  Main Clip List"
  AutoLoad
  
  'Add icon to system tray
  With Notify
    .cbSize = Len(Notify)
    .hwnd = Me.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Left click to show." & vbCrLf & "Right click for options." & vbNullChar
  End With
  Dim lResult As Long
  lResult = Shell_NotifyIcon(NIM_ADD, Notify)
  
  'If user wants to be on top
  If IsOnTop = True Then
   iret = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  End If
  
  'Set the color theme
  mnuSetTheme_Click (Themes)
  
  If ClipboardOn = True Then
    'Subclass this form - Hook into Windows Clipboard Chain
    HookForm frmMainClip
    
    'Register this form as a Clipboardviewer
    Busy = True
    m_hWndNext = SetClipboardViewer(frmMainClip.hwnd)
  Else
    mnuOption(0).Checked = False
  End If
 
  'We're done loading so set to capture clipboard
  Busy = False
  Loading = False
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Get system tray mouse message

  Static Message As Long
  Message = x / Screen.TwipsPerPixelX
  If AboutOn = True Then Exit Sub
  Select Case Message
    Case WM_RBUTTONUP 'Show popup menu
      If m_hPopup = 0 Then CreateMenu 'Create the popup menu if it doesn't exists yet
      
      Dim ptAPI As POINTAPI 'Get our mouse position
      GetCursorPos ptAPI

      Call ClientToScreen(Me.hwnd, ptAPI)
      
      'Select popup menu clicked
      Select Case TrackPopupMenu(m_hPopup, TPM_RETURNCMD, ptAPI.x, ptAPI.y, 0&, Me.hwnd, 0&)
        Case ID_CANCEL
          'Do Nothing
        Case ID_EXIT
          Unload Me
        Case ID_CAPIMAGE
          CapButton = False
          frmCapture.Show
        Case ID_CAPCOLOR
          frmGetColor.Show
        Case ID_SHOWQ
          Me.Show
        Case ID_TOGGLEON
          
          If ClipboardOn = True Then
           cmdUnhook_Click (0)
          Else
           cmdUnhook_Click (1)
          End If
          
        Case ID_HOTKEY
         frmHotKey.Show 1
      End Select
      
    Case WM_LBUTTONUP 'show qclips
      
      If Me.Visible = False Then
       Me.Show
      Else
       Me.Hide
      End If
  
  End Select
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
 If UnloadMode = False Then
  Cancel = True
  Me.Hide
 End If
 
 If iQSave = True Then Call AutoSave
 
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
 If Cancel = True Then Exit Sub
 
 On Error Resume Next
 
  'Destroy popup menu if it was created
  If m_hPopup Then Call DestroyMenu(m_hPopup)
  
  'Remove system tray icon
  Dim lResult As Long
  lResult = Shell_NotifyIcon(NIM_DELETE, Notify)
  DoEvents
  
     'Unhook form from Clipboard
   Call ChangeClipboardChain(frmMainClip.hwnd, m_hWndNext)
   UnHookForm Me

 'Save options to ini file and erase all temp files
  WriteOptions
  Kill App.Path & "\TEMP*.*"
  
  'Unload all forms
  Dim oForm As Form
  For Each oForm In Forms
    If oForm.hwnd <> Me.hwnd Then Unload oForm
  Next oForm


End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'Open clip with associated file program

 Dim TempText As String
 Dim ClipFilename As String
 
   TempText = ListTag(ucCoolList1.ListIndex)
   j = Val(Mid$(ListTag(ucCoolList1.ListIndex), 4, Len(ListTag(ucCoolList1.ListIndex))))
   
   If Left$(TempText, 3) = "PIC" Then
       Screen.MousePointer = 11
       ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".BMP"
       SavePicture picClip(j), ClipFilename
   End If
   
 If Button = vbLeftButton Then
   iret = ShellExecute(0&, vbNullString, ClipFilename, vbNullString, vbNullString, vbNormalFocus)
 ElseIf Button = vbRightButton Then
   Shell "rundll32 shell32.dll,OpenAs_RunDLL " & ClipFilename, vbNormalFocus
 End If
        
 Screen.MousePointer = 0
 
End Sub


Private Sub mnuAbout_Click()
 
 'no capturing while about showing About
 
 Dim OldClipboard As Boolean
 OldClipboard = ClipboardOn
 
   ClipboardOn = False
 
   If IsOnTop = True Then
    iret = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
   End If
 
 frmAbout.Show 1
 
   ClipboardOn = OldClipboard
 
   If IsOnTop = True Then
    iret = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   End If
 
End Sub

Private Sub mnuClearWinClip_Click()
 
 Clipboard.Clear
 
End Sub

Private Sub mnuClose_Click()
 
 If ShowWarnings = True Then
   iret = MsgBox("Are you sure you wish to shutdown Q-Clips?", vbQuestion + vbYesNoCancel + vbApplicationModal, "Shutdown Q-Clips?")
   If iret = vbNo Or iret = vbCancel Then Exit Sub
 End If
 
 'Shut 'er down
 Unload Me
 
End Sub

Private Sub mnuColorCapType_Click(Index As Integer)
 
 For i = 0 To 3
  mnuColorCapType(i).Checked = False
 Next
  
 ColorCapType = Index
 
 mnuColorCapType(Index).Checked = True
 
 
End Sub

Private Sub mnuEditMenu_Click()
 If curWindow = 0 Then
  mnuPasteDirect.Enabled = False
 Else
  mnuPasteDirect.Enabled = True
 End If
End Sub

Private Sub mnuExit_Click()
 
 Me.Hide
 
End Sub


Private Sub mnuImageType_Click(Index As Integer)
 mnuImageType(0).Checked = False
 mnuImageType(1).Checked = False
 
 ImageType = Index
 
 mnuImageType(Index).Checked = True
End Sub

Private Sub mnuKillClip_Click()

  If ShowWarnings = True Then
   iret = MsgBox("Are you sure you wish to delete this clip?", vbQuestion + vbYesNoCancel + vbApplicationModal, "Q-Clips: Delete List?")
   If iret = vbCancel Or iret = vbNo Then Exit Sub
  End If
  
   DeleteClip

  
End Sub

Private Sub mnuKillList_Click()
  
 If ucCoolList1.ListCount = 0 Then Exit Sub
 
 If ShowWarnings = True Then
  iret = MsgBox("Are you sure you wish to delete this Q-Clip list and all clips?", vbExclamation + vbYesNoCancel + vbApplicationModal, "Q-Clips: Delete List?")
  If iret = vbCancel Or iret = vbNo Then Exit Sub
 End If
   
 DeleteAll
 
  
End Sub



Private Sub mnuOpen_Click()
    
    On Error Resume Next
    CMDialog1.CancelError = True
    CMDialog1.Filter = "Q-Clip Files (*.qcl)|*.qcl|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Open Custom Q-Clip List"
    CMDialog1.FileName = ""
    CMDialog1.ShowOpen
    If Err = 32755 Then   ' User chose Cancel.
       Exit Sub
    Else

     If iQSave = True Then AutoSave
     
     ClearAll
     
     QFileName = CMDialog1.FileName
     
     Warning25 = True
     
     AutoLoad
     
      If QFileName = App.Path & "\qcliplist.qcx" Then
       Me.Caption = "Q - Clips:  Main Clip List"
      Else
       Me.Caption = "Q - Clips:  " & LastPart(QFileName)
       UpDateFileMenu QFileName
      End If
      
       'if nothing in clipboard then tell 'em
      If ucCoolList1.ListCount < 1 Then
       picClipLabel.Visible = True
      Else
       picClipLabel.Visible = False
      End If
    
    End If
    
     
End Sub



Private Sub mnuOpenClip_Click(Index As Integer)

 'Open clip with associated file program

 Dim TempText As String
 Dim ClipFilename As String
 Dim TempNum As Integer
 Dim strcontents As String
 
 If ucCoolList1.ListIndex < 0 Then Exit Sub
 
   TempText = ListTag(ucCoolList1.ListIndex)
   j = Val(Mid$(ListTag(ucCoolList1.ListIndex), 4, Len(ListTag(ucCoolList1.ListIndex))))
   
    If Left$(TempText, 3) = "PIC" Then
       Screen.MousePointer = 11
       ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".BMP"
       SavePicture picClip(j), ClipFilename
    End If
   
    If Left$(TempText, 3) = "TXT" Or Left$(TempText, 3) = "FIL" Then
      Screen.MousePointer = 11
      ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".TXT"
      TempNum = FreeFile
       ' Open the file.
      Open ClipFilename For Output As #TempNum
        ' Place the contents into a variable.
        strcontents = TextClip(j).Text
        ' Write the variable contents to a saved file.
        Print #TempNum, strcontents
      Close #TempNum
      strcontents = ""
    End If
 
    If Left$(TempText, 3) = "RTF" Then
       Screen.MousePointer = 11
       ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".RTF"
       RTFTextClip(j).SaveFile ClipFilename, rtfRTF
    End If
    
 If Index = 0 Then
   iret = ShellExecute(0&, vbNullString, ClipFilename, vbNullString, vbNullString, vbNormalFocus)
 ElseIf Index = 1 Then
   Shell "rundll32 shell32.dll,OpenAs_RunDLL " & ClipFilename, vbNormalFocus
 End If
        
 Screen.MousePointer = 0
 
End Sub

Private Sub mnuOption_Click(Index As Integer)

 Select Case Index
 
  Case 0 'Q-Clip capture copy's on or off

    If ClipboardOn = True Then
     cmdUnhook_Click (0)
    Else
     cmdUnhook_Click (1)
    End If
    
  Case 1 'Clip Multi-Select
   'RESERVED for multi-select future code
   'ClipMultiSelect = Not ClipMultiSelect
  ' mnuOption(1).Checked = ClipMultiSelect
   'If ClipMultiSelect = False Then
   ' ucCoolList1.SelectMode = 0
   'Else
   ' ucCoolList1.SelectMode = 1
  ' End If
   
   
  Case 3
   
   SetHotKeyOption
  
 End Select
 
 
End Sub



Private Sub mnuOptions_Click()
 'Reserved for future multi-select
 'mnuOption(1).Enabled = Not HoverOn
 'mnuOption(1).Checked = ClipMultiSelect
 
End Sub


Private Sub mnuPasteDirect_Click()
 
 PasteDirect
 
End Sub

Private Sub mnuPasteIndirect_Click()
 
 If ucCoolList1.ListIndex < 0 Then Exit Sub

 PutClip
 
 If PasteMinimized = True Then Me.Hide

 
End Sub

Private Sub mnuQHelp_Click()
 Dim Ftemp As String
 Dim There As Integer
 On Error Resume Next
 
   Ftemp = App.Path & "\qcliphelp.pdf"
   There = Exist(Ftemp)
   
    If Not There Then
      MsgBox "Q-Clips PDF help manual not found.  Contact iqProPlus", vbOKOnly + vbInformation + vbApplicationModal, "Q-Clips Help Manual"
      Exit Sub
    End If
   
   iret = ShellExecute(0&, vbNullString, Ftemp, vbNullString, vbNullString, vbNormalFocus)
   
   Exit Sub
   
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
 
 If RecentDocs(Index) = QFileName Then Exit Sub 'trying to open current file
 
 Dim There As Integer
 On Error Resume Next
 Screen.MousePointer = 11
 
 If iQSave = True Then AutoSave
      
 QFileName = RecentDocs(Index)
 
 There = Exist(QFileName)
  If There Then
     ClearAll
     DoEvents
     Warning25 = True
     AutoLoad
     Me.Caption = "Q - Clips:  " & LastPart(QFileName)
     UpDateFileMenu QFileName
  Else
     MsgBox QFileName & " no longer available!", vbOKOnly + vbApplicationModal, "Clip List Not Found"
     Screen.MousePointer = 0
     Exit Sub
 End If
 
End Sub

Private Sub mnuRunOnStart_Click()
 Dim Reg As Object
 
 RunOnStartUp = Not RunOnStartUp
 
 Select Case RunOnStartUp
 
  Case True
   mnuRunOnStart.Checked = True
   On Error Resume Next

   Set Reg = CreateObject("Wscript.shell")
   Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
   
   
  Case False
   mnuRunOnStart.Checked = False
   On Error Resume Next

   Set Reg = CreateObject("Wscript.Shell")
   Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
 End Select
 
End Sub

Private Sub mnuSave_Click()

 Call CustomSave(QFileName)

End Sub
Private Sub mnuSaveClip_Click()
 
 '1st save the clip file we're in if autosave is on
 If iQSave = True Then AutoSave
 
 Call CustomSaveAs
 
End Sub

Private Sub mnuSaveClipAs_Click()
  Dim TempText As String
  Dim ClipFilename As String
  Dim TempNum As Integer
  Dim strcontents As String
  Dim MyPic As StdPicture
  
   On Error Resume Next
   
   If ucCoolList1.ListIndex < 0 Then Exit Sub
 
   TempText = ListTag(ucCoolList1.ListIndex)
   j = Val(Mid$(ListTag(ucCoolList1.ListIndex), 4, Len(ListTag(ucCoolList1.ListIndex))))
   
    If Left$(TempText, 3) = "PIC" Then
     CMDialog1.Filter = "JPEG Image (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap Image (*.bmp)|*.bmp"
    ElseIf Left$(TempText, 3) = "TXT" Or Left$(TempText, 3) = "FIL" Then
     CMDialog1.Filter = "Text File (*.txt)|*.txt"
    ElseIf Left$(TempText, 3) = "RTF" Then
     CMDialog1.Filter = "Rich Text (*.rtf)|*.rtf"
    End If
    CMDialog1.FileName = ""
    CMDialog1.CancelError = True
    CMDialog1.FilterIndex = 1
    CMDialog1.FLAGS = cdlOFNOverwritePrompt
    CMDialog1.DialogTitle = "Save Q-Clip As..."
    CMDialog1.ShowSave
    
   If Err = 32755 Then   ' User chose Cancel.
       Exit Sub
   Else
   
     ClipFilename = CMDialog1.FileName
     
     If Left$(TempText, 3) = "PIC" Then
      Screen.MousePointer = 11
      If LCase(Right$(ClipFilename, 3)) = "jpg" Or LCase(Right$(ClipFilename, 3)) = "peg" Then
       SavePicture picClip(j), "$$_temp.bmp"
       DoEvents
       Set MyPic = LoadPicture("$$_temp.bmp")
       Set m_Image = New cImage
       m_Image.CopyStdPicture MyPic
       Call SaveImage(m_Image, ClipFilename)
       Set MyPic = Nothing
       Kill "$$_temp.bmp"
      Else
       SavePicture picClip(j), ClipFilename
      End If
     End If
   
     If Left$(TempText, 3) = "TXT" Or Left$(TempText, 3) = "FIL" Then
       Screen.MousePointer = 11
       TempNum = FreeFile
       ' Open the file.
       Open ClipFilename For Output As #TempNum
        ' Place the contents into a variable.
         strcontents = TextClip(j).Text
         ' Write the variable contents to a saved file.
         Print #TempNum, strcontents
       Close #TempNum
       strcontents = ""
     End If
 
     If Left$(TempText, 3) = "RTF" Then
       Screen.MousePointer = 11
       RTFTextClip(j).SaveFile ClipFilename, rtfRTF
     End If
    
    Screen.MousePointer = 0
    
   End If
    
End Sub

Private Sub mnuScreenCap_Click(Index As Integer)
 For i = 0 To 2
  mnuScreenCap(i).Checked = False
 Next
 
  CaptureOption = Index
  
  mnuScreenCap(Index).Checked = True
  
End Sub

Private Sub mnuSetTheme_Click(Index As Integer)
 For i = 0 To 4
  mnuSetTheme(i).Checked = False
 Next
  
  Select Case Index
  
   Case 0
    DrawBlue

   Case 1
     DrawSilver
   Case 2
     DrawBlack
   Case 3
     DrawOlive
   Case 4
     DrawNone
     
  End Select
  
 mnuSetTheme(Index).Checked = True
 Themes = Index
 
End Sub

Private Sub mnuViewClips_Click()

 If iQSave = True Then AutoSave
 ClearAll
 QFileName = App.Path & "\qcliplist.qcx"
 Warning25 = True
 AutoLoad
 Me.Caption = "Q - Clips:  Main Clip List"
 
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 'Open file with associated program
 Dim TempText As String
 Dim ClipFilename As String
 Dim TempNum As Integer
 Dim strcontents As String
 
 
   TempText = ListTag(ucCoolList1.ListIndex)
   j = Val(Mid$(ListTag(ucCoolList1.ListIndex), 4, Len(ListTag(ucCoolList1.ListIndex))))
   
    If Left$(TempText, 3) = "TXT" Or Left$(TempText, 3) = "FIL" Then
     Screen.MousePointer = 11
      'let's see if it's a URL
     strcontents = TextClip(j).Text
     iret = InStr(1, strcontents, " ")
     If iret = 0 And LTrim(Left$(strcontents, 7)) = "http://" Or iret = 0 And LTrim(Left$(strcontents, 7)) = "<http:/" Or iret = 0 And LTrim(Left$(strcontents, 4)) = "www." Then
      ClipFilename = strcontents
       If Left$(ClipFilename, 1) = "<" Then
        TrimNull ClipFilename
        ClipFilename = Right$(ClipFilename, Len(ClipFilename) - 1)
        ClipFilename = Left$(ClipFilename, Len(ClipFilename) - 1)
       End If
     Else
      ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".TXT"
      TempNum = FreeFile
       ' Open the file.
      Open ClipFilename For Output As #TempNum
       ' Place the contents into a variable.
       strcontents = TextClip(j).Text
       ' Write the variable contents to a saved file.
       Print #TempNum, strcontents
      Close #TempNum
      strcontents = ""
      End If
    End If
 
    If Left$(TempText, 3) = "RTF" Then
       Screen.MousePointer = 11
     'let's see if it's a URL
      strcontents = RTFTextClip(j).Text
      iret = InStr(1, strcontents, " ")
      If iret = 0 And LTrim(Left$(strcontents, 7)) = "http://" Or iret = 0 And LTrim(Left$(strcontents, 7)) = "<http:/" Or iret = 0 And LTrim(Left$(strcontents, 4)) = "www." Then
       ClipFilename = strcontents
       If Left$(ClipFilename, 1) = "<" Then
        TrimNull ClipFilename
        ClipFilename = Right$(ClipFilename, Len(ClipFilename) - 1)
        ClipFilename = Left$(ClipFilename, Len(ClipFilename) - 1)
       End If
      Else
       ClipFilename = App.Path & "\TEMP" & LTrim(Str(ucCoolList1.ListIndex)) & ".RTF"
       RTFTextClip(j).SaveFile ClipFilename, rtfRTF
      End If
    End If
 
 
 If Button = vbLeftButton Then
   iret = ShellExecute(0&, vbNullString, ClipFilename, vbNullString, vbNullString, vbNormalFocus)
 ElseIf Button = vbRightButton Then
   Shell "rundll32 shell32.dll,OpenAs_RunDLL " & ClipFilename, vbNormalFocus
 End If
 
 Screen.MousePointer = 0
 
End Sub


Private Sub ucCoolList1_ListIndexChange()
  
 If Me.Visible = False Then Exit Sub
 
 Dim indx As Integer
 Dim TempText As String
 Dim IsFileList As Boolean
 
 Text1.BackColor = vbWhite
 
   'if nothing in clipboard then tell 'em
 If ucCoolList1.ListCount < 1 Then
  picClipLabel.Visible = True
 Else
  picClipLabel.Visible = False
 End If
 
 indx = ucCoolList1.ListIndex
 If indx < 0 Then indx = 0
 If indx > 0 And Me.WindowState = 0 And Me.Visible = True Then Me.SetFocus
 TempText = ListTag(indx)


 If Left$(TempText, 3) = "PIC" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Image1.Picture = picClip(j)
       Reset_Image
       Picture1.Visible = True
       Text1.Visible = False
       lblInfo = "Picture Clip - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
 End If
 
 If Left$(TempText, 3) = "FIL" Then
   j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
   Text1.Text = TextClip(j).Text
   Picture1.Visible = False
   Text1.Visible = True
   lblInfo = "Copy Files - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
 End If
 
 If Left$(TempText, 3) = "TXT" Then
   j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
   Text1.Text = TextClip(j).Text
   Picture1.Visible = False
   Text1.Visible = True
   If LTrim(Left$(Text1.Text, 7)) = "http://" Or iret = 0 And LTrim(Left$(Text1.Text, 7)) = "<http:/" Or LTrim(Left$(Text1.Text, 4)) = "www." Then
    lblInfo = "Web Address - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
   Else
    lblInfo = "Text Clip - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
   End If
 End If
 
 If Left$(TempText, 3) = "RTF" Then
  j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
  Text1.Text = RTFTextClip(j).Text
  Picture1.Visible = False
  Text1.Visible = True
  If LTrim(Left$(Text1.Text, 7)) = "http://" Or iret = 0 And LTrim(Left$(Text1.Text, 7)) = "<http:/" Or LTrim(Left$(Text1.Text, 4)) = "www." Then
   lblInfo = "Web Address - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
  Else
   lblInfo = "Rich Text Clip - Item" & Str(indx + 1) & " of" & Str(ucCoolList1.ListCount)
  End If
 End If
 

 
End Sub


Private Sub ucCoolList1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 
 If Button = vbLeftButton And Shift = 2 Then
  ShowColor (Text1.Text)
  Exit Sub
 End If
  
 
 If Button = vbLeftButton And ucCoolList1.ListCount > 0 Then

  If ucCoolList1.HoverSelection = True Then PasteDirect

 End If


 '----Show Edit options popup menu
 
 If Button = vbRightButton Then
  
   PopupMenu mnuEditMenu
  
 End If
 
 
End Sub

Private Sub VBHotKey1_HotkeyPressed()

' When hotkey is pressed the program comes here
' to get the handle to the active window.  This
' is so we can paste directly back into the active program

 Dim Start As Long
 Dim sWindowText As String * 100
 Dim sClassName As String * 100
 Dim hWndParent As Long
 Dim sParentClassName As String * 100
 Dim wID As Long
 Dim lWindowStyle As Long
 Dim hInstance As Long
 Dim sParentWindowText As String * 100
 Dim x As Long
 Dim y As Long
 
 On Error Resume Next
 
 'First we need to get all the window handles
 Call GetCursorPos(mousePT) 'curWindow is hWnd when hotkey pressed
  x = mousePT.x
  y = mousePT.y
  curWindow = WindowFromPoint(x, y)
  
  'Get window text/caption
  r = GetWindowText(curWindow, sWindowText, 100)      ' Window text
  WinCaption = Left(sWindowText, r)

  'Get window Classname
  r = GetClassName(curWindow, sClassName, 100)         ' Window Class
  WinClassName = Left(sClassName, r)

'Now we need to get Parent Window Information

         ' Get handle of parent window:
         hWndParent = GetParent(curWindow)

         ' If there is a parent get more info:

         If hWndParent <> 0 Then
            ' Get ID of window:
            'wID = GetWindowWord(curWindow, GWW_ID)
           ' Print "Window ID Number: "; (wID)
          ParentWindow = hWndParent

            ' Get the text of the Parent window:
            r = GetWindowText(hWndParent, sParentWindowText, 100)
            ParentCaption = Left(sParentWindowText, r)

            ' Get the class name of the parent window:
            r = GetClassName(hWndParent, sParentClassName, 100)
            ParentClassName = Left(sParentClassName, r)
        
        End If

         ' Get window instance:
         hInstance = GetWindowWord(curWindow, GWW_HINSTANCE)


    DoEvents

    Me.Show
    Me.SetFocus
    
    If ucCoolList1.ListCount > 0 Then
     ucCoolList1.ListIndex = ucCoolList1.ListCount - 1
     ucCoolList1.SetFocus
    End If

End Sub


Private Sub VBHotKey2_HotkeyPressed()

'Screen Capture hot key was pressed

 Dim Start As Long
 If AboutOn = True Then Exit Sub
 If Me.Visible Then Me.Hide
 
 'Delay to let other windows repaint
 Start = Timer + 0.25
 While Start > Timer
  DoEvents
 Wend
 
 Select Case CaptureOption
 
  Case 0 'capture region
   CapButton = True
   frmCapture.Show
  Case 1 'capture desktop
   keybd_event vbKeySnapshot, 0&, 0&, 0&
  Case 2 'capture active window
   keybd_event vbKeySnapshot, &H1&, 0&, 0&
   
 End Select
 
   
 
End Sub


Private Sub VBHotKey3_HotkeyPressed()

 If AboutOn = True Then Exit Sub
 frmGetColor.Show
 
End Sub
