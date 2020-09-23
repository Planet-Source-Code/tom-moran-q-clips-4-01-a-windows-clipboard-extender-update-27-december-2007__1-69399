Attribute VB_Name = "ModIniSettings"
Option Explicit
' API functions used to read and write to INI.
' Used for handling the recent files list and options.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Long

'Needed to keep full path of recent documents_
'which is different than parsed recent doc menu items
Public RecentDocs(0 To 8) As String

'Used to compact recent file paths-----------------
Private Const MAX_PATH As Long = 260

Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
   Alias "PathCompactPathExA" _
  (ByVal pszOut As String, _
   ByVal pszSrc As String, _
   ByVal cchMax As Long, _
   ByVal dwFlags As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Private TopForm As Long
Private LeftForm As Long

Public ColorHotKey As String
Public ColorCtrlKey As Boolean
Public ColorShiftKey As Boolean
Public ColorAltKey As Boolean
Public ColorWinKey As Boolean
Public ClipHotKey As String
Public ClipCtrlKey As Boolean
Public ClipShiftKey As Boolean
Public ClipAltKey As Boolean
Public ClipWinKey As Boolean
Public CapHotKey As String
Public CapCtrlKey As Boolean
Public CapShiftKey As Boolean
Public CapAltKey As Boolean
Public CapWinKey As Boolean
Public RunOnStartUp As Boolean

Public CaptureOption As Integer
'0 =Region
'1 = Desktop
'2 = Active Window

Private Key As String
Private retval As Long



Public Sub WriteRecentFiles(OpenFileName As String)
'=========== write recent files======================
  Dim IniString As String
  IniString = String(255, 0)

  ' Copy RecentFile1 to RecentFile2, etc.
  For i = 7 To 1 Step -1
    Key = "RecentFile" & Trim(i)
    retval = GetPrivateProfileString("Recent Files", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
    If retval And Left(IniString, 8) <> "Not Used" Then
      Key = "RecentFile" & Trim((i + 1))
      retval = WritePrivateProfileString("Recent Files", Key, IniString, App.Path & "\qclip.ini")
    End If
  Next i
  
  ' Write openfile to first Recent File.
    retval = WritePrivateProfileString("Recent Files", "RecentFile1", OpenFileName, App.Path & "\qclip.ini")

End Sub

Sub UpDateFileMenu(FileName As String)

        ' Check if OpenFileName is already on MRU list.
       
        retval = OnRecentFilesList(FileName)
        If retval = False Then
          ' Write OpenFileName to INI
          WriteRecentFiles (LCase(FileName))
          
        End If
        ' Update menus for most recent file list.
        GetRecentFiles
        
End Sub
Public Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function
Public Function OnRecentFilesList(FileName As String) As Integer

  RecentDocs(0) = FileName
 
  For i = 1 To 8
    'debug MsgBox LCase(RecentDocs(i)) & vbCrLf & LCase(Filename)
    If LCase(RecentDocs(i)) = LCase(FileName) Then
     'move the recent doc to the top
     For j = i To 1 Step -1
      RecentDocs(j) = RecentDocs(j - 1)
       ' Write changed file
      retval = WritePrivateProfileString("Recent Files", "RecentFile" & Trim((j)), RecentDocs(j), App.Path & "\qclip.ini")
     Next j
     OnRecentFilesList = True
     Exit Function
    End If
  Next i
  
    OnRecentFilesList = False
End Function

Private Function MakeCompactedPathChrs(ByVal sPath As String, _
                                       ByVal cchMax As Long) As String

  'Truncates a path to a specified
  'number of characters by replacing
  'path components with ellipses.
   Dim buff As String
   
  'cchMax is the maximum number of characters
  'to be contained in the new string, **including
  'the terminating NULL character**. For example,
  'if cchMax = 8, the resulting string would contain
  'a maximum of 7 characters plus the termnating null.
  '
  'Because of this, we're add 1 to the value passed
  'as cchMax to ensure the resulting string is
  'the size requested.
   cchMax = cchMax + 1
   buff = Space$(MAX_PATH)
   retval = PathCompactPathEx(buff, sPath, cchMax, 0&)
   
   MakeCompactedPathChrs = TrimNull(buff)
   
End Function



Public Sub GetRecentFiles()
'----------GetRecentFiles
  Dim IniString As String
  
  'Clear out the file array
  For i = 1 To 8
   RecentDocs(i) = ""
  Next
  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)

  ' Get recent file strings from qclip.ini
  For i = 1 To 8
    Key = "RecentFile" & i
    retval = GetPrivateProfileString("Recent Files", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
     
    If retval And Left(IniString, 8) <> "Not Used" Then
      ' Update the form's menu.
      RecentDocs(i) = TrimNull(IniString)
    End If
  Next i
  '===========================================================================
   'Now let's write just the filename to the menu with shortened path
    
   For i = 1 To 8
    If RecentDocs(i) = "" Then Exit For
    frmMainClip.mnuRecentFiles(i).Caption = MakeCompactedPathChrs(RecentDocs(i), 30)
    frmMainClip.mnuRecentFiles(i).Visible = True
   Next
    If RecentDocs(1) <> "" Then frmMainClip.mnuRecentFiles(1).Enabled = True

  '===============================================================================
End Sub


'-----------------------------------------------------
Sub GetOptions()
  Dim l As Integer
  Dim IniString As String

  ' This variable must be large enough to hold the return string
  ' from the GetPrivateProfileString API.
  IniString = String(255, 0)
  
  
  'Get the Screen Capture keys
  '==============================================================================
  Key = "CapHotKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      CapHotKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "CapCtrlKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
     CapCtrlKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "CapShiftKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      CapShiftKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "CapAltKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      CapAltKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "CapWinKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      CapWinKey = Left$(IniString, l - 1)
    End If
  End If
  
  'Get the Clipboard hot key defaults
  '==============================================================================
  
  Key = "ClipHotKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ClipHotKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ClipCtrlKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
     ClipCtrlKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ClipShiftKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ClipShiftKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "ClipAltKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ClipAltKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "ClipWinKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ClipWinKey = Left$(IniString, l - 1)
    End If
  End If
  
  ' Get Color Capture HotKey
  '==============================================================================
  
  Key = "ColorHotKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ColorHotKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ColorCtrlKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
     ColorCtrlKey = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ColorShiftKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ColorShiftKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "ColorAltKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ColorAltKey = Left$(IniString, l - 1)
    End If
  End If

  Key = "ColorWinKey"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ColorWinKey = Left$(IniString, l - 1)
    End If
  End If

  
  'Get program options
  '==============================================================================
  Key = "iQSound"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      iQSound = Left$(IniString, l - 1)
    End If
  End If

  Key = "iQSave"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      iQSave = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ViewPane"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ViewPane = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "IsOnTop"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      IsOnTop = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "HoverOn"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      HoverOn = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ShowWarnings"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ShowWarnings = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "StartMinimized"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      StartMinimized = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "PasteMinimized"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      PasteMinimized = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "TopForm"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      frmMainClip.Top = Val(Left$(IniString, l - 1))
    End If
  End If
  
  Key = "LeftForm"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      frmMainClip.Left = Val(Left$(IniString, l - 1))
    End If
  End If
  
  Key = "RunOnStartUp"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      RunOnStartUp = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "ClipboardOn"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ClipboardOn = Left$(IniString, l - 1)
    End If
  End If
  
  Key = "Themes"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      Themes = Val(Left$(IniString, l - 1))
    End If
  End If

  Key = "CaptureOption"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      CaptureOption = Val(Left$(IniString, l - 1))
    End If
  End If
 
  Key = "ColorCapType"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ColorCapType = Val(Left$(IniString, l - 1))
    End If
  End If
  
  Key = "ImageType"
  retval = GetPrivateProfileString("Options", Key, "Not Used", IniString, Len(IniString), App.Path & "\qclip.ini")
  If retval And Left(IniString, 8) <> "Not Used" Then
    l = InStr(IniString, Chr$(0))
    If l Then
      ImageType = Left$(IniString, l - 1)
    End If
  End If
  
 End Sub
Sub SetOptions()

 If iQSound = True Then
  frmMainClip.chkSound.Value = 1
 Else
  frmMainClip.chkSound.Value = 0
 End If
 
 If iQSave = True Then
  frmMainClip.chkAutoSave = 1
 Else
  frmMainClip.chkAutoSave = 0
 End If
 
 If ViewPane = True Then
   frmMainClip.cmdPane(0).Visible = True
   frmMainClip.cmdPane(1).Visible = False
   frmMainClip.Width = 6000
 Else
   frmMainClip.cmdPane(1).Visible = True
   frmMainClip.cmdPane(0).Visible = False
   frmMainClip.Width = 2960
 End If
 
 If IsOnTop = True Then
  frmMainClip.chkOnTop.Value = 1
  iret = SetWindowPos(frmMainClip.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 Else
  frmMainClip.chkOnTop.Value = 0
  iret = SetWindowPos(frmMainClip.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
 End If
 
 If HoverOn = True Then
  frmMainClip.chkAutoHover.Value = 1
  frmMainClip.ucCoolList1.HoverSelection = True
  frmMainClip.ucCoolList1.SelectMode = 0
  'ClipMultiSelect = False
 Else
  frmMainClip.chkAutoHover.Value = 0
  frmMainClip.ucCoolList1.HoverSelection = False
  frmMainClip.ucCoolList1.SelectMode = 0
  'ClipMultiSelect = False
 End If
 
 If ShowWarnings = True Then
  frmMainClip.chkWarnings.Value = 1
 Else
  frmMainClip.chkWarnings.Value = 0
 End If
 
 If StartMinimized = True Then
  frmMainClip.chkStartMinimized.Value = 1
 Else
  frmMainClip.chkStartMinimized.Value = 0
 End If
 
 If PasteMinimized = True Then
  frmMainClip.chkPasteMinimize.Value = 1
 Else
  frmMainClip.chkPasteMinimize.Value = 0
 End If
 
 frmMainClip.mnuRunOnStart.Checked = RunOnStartUp
 
 If ClipboardOn = False Then
    frmMainClip.cmdUnHook(0).Visible = False
    frmMainClip.cmdUnHook(1).Visible = True
    frmMainClip.mnuOption(0).Checked = False
 End If

 frmMainClip.mnuImageType(0).Checked = False
 frmMainClip.mnuImageType(1).Checked = False
 frmMainClip.mnuImageType(ImageType).Checked = True
   
End Sub

Public Sub WriteOptions()
  Dim IniString As String
  
   'get last settings so will start in same position next time
   TopForm = frmMainClip.Top
   LeftForm = frmMainClip.Left


  'Hot Keys Defaults
  '=====================================
  'Screen capture
  '=====================================
  IniString = CapHotKey
  retval = WritePrivateProfileString("Options", "CapHotKey", IniString, App.Path & "\qclip.ini")
  
  IniString = CapCtrlKey
  retval = WritePrivateProfileString("Options", "CapCtrlKey", IniString, App.Path & "\qclip.ini")

  IniString = CapShiftKey
  retval = WritePrivateProfileString("Options", "CapShiftKey", IniString, App.Path & "\qclip.ini")

  IniString = CapAltKey
  retval = WritePrivateProfileString("Options", "CapAltKey", IniString, App.Path & "\qclip.ini")
  
  IniString = CapWinKey
  retval = WritePrivateProfileString("Options", "CapWinKey", IniString, App.Path & "\qclip.ini")
  
  'hotkeys for Clipboard
  '===========================================
  IniString = ClipHotKey
  retval = WritePrivateProfileString("Options", "ClipHotKey", IniString, App.Path & "\qclip.ini")
  
  IniString = ClipCtrlKey
  retval = WritePrivateProfileString("Options", "ClipCtrlKey", IniString, App.Path & "\qclip.ini")

  IniString = ClipShiftKey
  retval = WritePrivateProfileString("Options", "ClipShiftKey", IniString, App.Path & "\qclip.ini")

  IniString = ClipAltKey
  retval = WritePrivateProfileString("Options", "ClipAltKey", IniString, App.Path & "\qclip.ini")
  
  IniString = ClipWinKey
  retval = WritePrivateProfileString("Options", "ClipWinKey", IniString, App.Path & "\qclip.ini")
  
  'hotkeys for Color Capture
  '===========================================
  IniString = ColorHotKey
  retval = WritePrivateProfileString("Options", "ColorHotKey", IniString, App.Path & "\qclip.ini")
  
  IniString = ColorCtrlKey
  retval = WritePrivateProfileString("Options", "ColorCtrlKey", IniString, App.Path & "\qclip.ini")

  IniString = ColorShiftKey
  retval = WritePrivateProfileString("Options", "ColorShiftKey", IniString, App.Path & "\qclip.ini")

  IniString = ColorAltKey
  retval = WritePrivateProfileString("Options", "ColorAltKey", IniString, App.Path & "\qclip.ini")
  
  IniString = ColorWinKey
  retval = WritePrivateProfileString("Options", "ColorWinKey", IniString, App.Path & "\qclip.ini")
 
 
   'Program Defaults
  '=====================================
  
  IniString = iQSound
  retval = WritePrivateProfileString("Options", "iQSound", IniString, App.Path & "\qclip.ini")

  IniString = iQSave
  retval = WritePrivateProfileString("Options", "iQSave", IniString, App.Path & "\qclip.ini")

  IniString = ViewPane
  retval = WritePrivateProfileString("Options", "ViewPane", IniString, App.Path & "\qclip.ini")
  
  IniString = IsOnTop
  retval = WritePrivateProfileString("Options", "IsOnTop", IniString, App.Path & "\qclip.ini")

  IniString = HoverOn
  retval = WritePrivateProfileString("Options", "HoverOn", IniString, App.Path & "\qclip.ini")

  IniString = ShowWarnings
  retval = WritePrivateProfileString("Options", "ShowWarnings", IniString, App.Path & "\qclip.ini")

  IniString = StartMinimized
  retval = WritePrivateProfileString("Options", "StartMinimized", IniString, App.Path & "\qclip.ini")

  IniString = PasteMinimized
  retval = WritePrivateProfileString("Options", "PasteMinimized", IniString, App.Path & "\qclip.ini")
  
  IniString = TopForm
  retval = WritePrivateProfileString("Options", "TopForm", IniString, App.Path & "\qclip.ini")
  
  IniString = LeftForm
  retval = WritePrivateProfileString("Options", "LeftForm", IniString, App.Path & "\qclip.ini")
  
  IniString = RunOnStartUp
  retval = WritePrivateProfileString("Options", "RunOnStartUp", IniString, App.Path & "\qclip.ini")
  
  IniString = ClipboardOn
  retval = WritePrivateProfileString("Options", "ClipboardOn", IniString, App.Path & "\qclip.ini")
  
  IniString = Themes
  retval = WritePrivateProfileString("Options", "Themes", IniString, App.Path & "\qclip.ini")
  
  IniString = CaptureOption
  retval = WritePrivateProfileString("Options", "CaptureOption", IniString, App.Path & "\qclip.ini")

  IniString = ColorCapType
  retval = WritePrivateProfileString("Options", "ColorCapType", IniString, App.Path & "\qclip.ini")
  
 IniString = ImageType
  retval = WritePrivateProfileString("Options", "ImageType", IniString, App.Path & "\qclip.ini")
  
End Sub


