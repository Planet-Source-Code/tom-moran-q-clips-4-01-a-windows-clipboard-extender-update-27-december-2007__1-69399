VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About iQ WordPad"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   Icon            =   "frmAboutiqClip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAboutiqClip.frx":86EA
   ScaleHeight     =   5205
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjQClips40.CandyButton cmdOkay 
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Top             =   4545
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Okay"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   14704640
      ColorButtonUp   =   13668448
      ColorButtonDown =   11108432
      BorderBrightness=   0
      ColorBright     =   16775930
      DisplayHand     =   0   'False
      ColorScheme     =   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   585
      TabIndex        =   1
      Top             =   4185
      Width           =   3780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   3375
      Width           =   4380
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetUserName Lib "advapi32.dll" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, _
            nSize As Long) As Long

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" _
   (lpBuffer As MEMORYSTATUS)

Private Function UserName() As String

    Dim llReturn As Long
    Dim lsUserName As String
    Dim lsBuffer As String
    
    lsUserName = ""
    lsBuffer = Space$(255)
    llReturn = GetUserName(lsBuffer, 255)
    
    
    If llReturn Then
       lsUserName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
    End If
    
    UserName = lsUserName
End Function

Private Sub cmdOkay_Click()
 Unload Me
End Sub

Private Sub Form_Load()
   Dim MS As MEMORYSTATUS
   
   MS.dwLength = Len(MS)
   GlobalMemoryStatus MS

    Label1.Caption = UserName
    Label2.Caption = Format$(MS.dwAvailVirtual / 1024, "###,###,###,###") & " KB"
    
   AboutOn = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
 AboutOn = False
End Sub
