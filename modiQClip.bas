Attribute VB_Name = "modiQClip"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This stuff is to read or set a list of files copied to
' clipboard by windows explorer or similar programs.
' and to set files back to clipboard
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const CF_HDROP As Long = 15

Private Declare Function IsClipboardFormatAvailable Lib "user32" _
  (ByVal uFormat As Long) As Long
  
Private Declare Function OpenClipboard Lib "user32" _
  (ByVal hWndNewOwner As Long) As Long

Private Declare Function GetClipboardData Lib "user32" _
  (ByVal uFormat As Long) As Long
  
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Private Declare Sub CloseClipboard Lib "user32" ()

Private Declare Function DragQueryFile Lib "shell32.dll" _
   Alias "DragQueryFileA" _
  (ByVal hDrop As Long, _
   ByVal iFile As Long, _
   ByVal lpszFile As String, _
   ByVal cch As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long
  
' DROPFILES data structure.
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

' Global memory routines for set files to clipboard
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
  
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  More Clipboard API's and Globals including the very important_
'  Hook and Set iQ Clipboard as part of Windows Clipboard chain
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long

Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_CHANGECBCHAIN = &H30D

' Static handle to next window in clipboard chain
Public m_hWndNext As Long
Public Const GWL_WNDPROC = (-4)

Public PrevProc As Long
Public IsHooked As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' For playing wave files
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    

' To open clips in associated programs
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long


'Some QClip Globals
Public QFileName As String
Public iret As Long
Public i As Integer
Public j As Integer
Public ListTag(0 To 50) As String
Public TextCount As Integer
Public PicCount As Integer
Public RTFTextCount As Integer
Public JustPasted As Boolean
Public Busy As Boolean
Public ClipboardOn As Boolean
Public CapButton As Boolean
Public AboutOn As Boolean
'Public ClipMultiSelect 'reserved for future multi-select

'Option Globals
Public iQSound As Boolean
Public iQSave As Boolean
Public ViewPane As Boolean
Public IsOnTop As Boolean
Public HoverOn As Boolean
Public ShowWarnings As Boolean
Public Warning25 As Boolean
Public StartMinimized As Boolean
Public PasteMinimized As Boolean
Public ImageType As Integer '0 = bmp  1 = jpg

''''''''''''''''''''''''''''''''''''''

'for xp theme
Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

'To Stay on Top
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Copy the file names into the clipboard.
' Return True if we succeed.
Public Function ClipboardSetFiles(file_names() As String) As Boolean
Dim file_string As String
Dim drop_files As DROPFILES
Dim memory_handle As Long
Dim memory_pointer As Long
Dim i As Long

    ' Clear the clipboard.
    Clipboard.Clear

    ' Open the clipboard.
    If OpenClipboard(0) Then
        ' Build a null-terminated list of file names.
        For i = LBound(file_names) To UBound(file_names)
            file_string = file_string & file_names(i) & vbNullChar
        Next
        file_string = file_string & vbNullChar

        ' Initialize the DROPFILES structure.
        drop_files.pFiles = Len(drop_files)
        drop_files.fWide = 0    ' ANSI characters.
        drop_files.fNC = 0      ' Client coordinates.

        ' Get global memory to hold the DROPFILES
        ' structure and the file list string.
        memory_handle = GlobalAlloc(GHND, Len(drop_files) + Len(file_string))
        If memory_handle Then
            ' Lock the memory while we initialize it.
            memory_pointer = GlobalLock(memory_handle)

            ' Copy the DROPFILES structure and the
            ' file string into the global memory.
            CopyMem ByVal memory_pointer, drop_files, Len(drop_files)
            CopyMem ByVal memory_pointer + Len(drop_files), ByVal file_string, Len(file_string)
            GlobalUnlock memory_handle

            ' Copy the data to the clipboard.
            SetClipboardData CF_HDROP, memory_handle
            ClipboardSetFiles = True
        End If

        ' Close the clipboard.
        CloseClipboard
    End If
End Function

Public Function Exist(TF As String)

Dim FileNum1 As Integer
On Local Error GoTo Whoops

FileNum1 = FreeFile
 Open TF For Input As FileNum1
 Close FileNum1
 Exist = -1
 Exit Function

Whoops:
Exist = 0
Exit Function

End Function
Sub Main()
 '-----------------------------
 'Q-Clips program starts here!
 '-----------------------------
 If App.PrevInstance Then
   End
 End If
 
 Load frmMainClip
 
 Warning25 = True
 
 If StartMinimized Then
  frmMainClip.Hide
 Else
  frmMainClip.Show
 End If
 
 'If main list has 25 max clips and showwarnings is true tell user one time
   If frmMainClip.ucCoolList1.ListCount = 25 And ShowWarnings = True And Warning25 = True Then
   Dim TempText
   Warning25 = False
   If frmMainClip.Visible = False Then frmMainClip.Visible = True
   TempText = "Warning: You have 25 clips in this Q-Clips Collection." & vbCrLf
   TempText = TempText & "The next clipboard capture will delete the oldest (1st) clip in this collection."
   MsgBox TempText, vbExclamation + vbOKOnly + vbApplicationModal, "Max Q-Clips In Collection"
  End If
        
End Sub

Private Sub ResetClipboard()
  'put selected item in windows clipboard
  Dim indx As Integer
  Dim TempText As String
  indx = frmMainClip.ucCoolList1.ListIndex
  If indx < 0 Then indx = 0
  TempText = ListTag(indx)
  
  JustPasted = True
  Busy = True
  
      'Unhook the form while we do this
   Call ChangeClipboardChain(frmMainClip.hwnd, m_hWndNext)
   UnHookForm frmMainClip
   
   DoEvents
   If Left$(TempText, 3) = "PIC" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       Clipboard.SetData frmMainClip.picClip(j).Picture, vbCFBitmap
   ElseIf Left$(TempText, 3) = "TXT" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       Clipboard.SetText frmMainClip.TextClip(j).Text, vbCFText
   ElseIf Left$(TempText, 3) = "RTF" Then
       j = Val(Mid$(ListTag(indx), 4, Len(ListTag(indx))))
       Clipboard.Clear
       frmMainClip.RTFTextClip(j).SelStart = 0
       frmMainClip.RTFTextClip(j).SelLength = Len(frmMainClip.RTFTextClip(j).TextRTF)
       SendMessage frmMainClip.RTFTextClip(j).hwnd, WM_COPY, 0&, 0& 'Copy
   End If
   
  'We're Done so hook it back up
    HookForm frmMainClip
    'Register this form as a Clipboardviewer
    m_hWndNext = SetClipboardViewer(frmMainClip.hwnd)
   DoEvents
   
  JustPasted = False
  
End Sub

Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function
Private Sub ClipboardGetFiles()

   Dim hCBData As Long
   Dim numFiles As Long
   Dim cnt As Long
   Dim cbBuff As Long
   Dim buff As String
   Dim TempText As String
   
      TempText = ""
      
     'First get file names and add it to TempText string
     
      If OpenClipboard(0&) <> 0 Then
        
        'GetClipboardData returns a handle
        'that identifies a list of files.
        'The application retrieves this
        'information by passing the handle
        'to the DragQueryFile function.
         hCBData = GetClipboardData(CF_HDROP)
        
         If hCBData <> 0 Then
         
           'the first call to DragQueryFile
           'passes -1 as the iFile value; the
           'api in return indicates the number
           'of files to be dropped.
                       
           'note that once we're done we do not call
           'DragFinish as we did nothing with the
           'data other than list it.
            numFiles = DragQueryFile(hCBData, -1, 0&, 0&)
            
           'loop for each file ..
            For cnt = 0 To numFiles - 1
            
              '.. and each subsequent call to DragQueryFile
              'returns info about the file specified by
              'the iFile value. By passing null for the
              'buffer and size, the required buffer size
              'is returned.
               cbBuff = DragQueryFile(hCBData, cnt, vbNullString, 0&)

               If cbBuff > 0 Then
               
                 'pad the buffer adding one for
                 'the terminating null char
                  buff = Space$(cbBuff + 1)
                  cbBuff = Len(buff)
                  
                 'the final call will return the file;
                 'the return value, not used here, is
                 'the size of the file string returned.
                  If DragQueryFile(hCBData, cnt, buff, cbBuff) > 0 Then
                     
                    'just add to list for this purpose
                     TempText = TempText & TrimNull(buff) & vbCrLf
                     
                     
                  End If
                  
               End If
            
            Next cnt

         End If  'hCbData
        
         Call CloseClipboard
    
      End If  'OpenClipboard
      
      'Now let's Paste TempText Into Clipboard
       Clipboard.Clear
       Clipboard.SetText TempText, vbCFText
   

    
End Sub
Public Sub GetClip()

 '------------------------------------------
 ' This sub is where the Q-Clips are created
 '-------------------------------------------

  'Did user copy list item to clipboard?
   If JustPasted = True Then
    Exit Sub
   End If

   Dim ClpFmt As Integer
   Dim TempText As String
   Dim Start As Long
  
   On Error Resume Next   ' Set up error handling.
   

   'no captures while we're processing
   Busy = True
  
   ClpFmt = 0
   
   'Get the clipboard format
   If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
   If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
   If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
   If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
   
   'if none of the above is it a list of files?
   If ClpFmt = 0 And Clipboard.GetFormat(vbCFFiles) Then
     ClipboardGetFiles
     ClpFmt = 15
   End If
   
   Select Case ClpFmt
      
      Case 1 'Text Only
       ListTag(frmMainClip.ucCoolList1.ListCount) = "TXT" & Str(TextCount)
        
        'firsttime through use main else load a new control
        If TextCount > 0 Then
         Load frmMainClip.TextClip(TextCount)
        End If

        'load the clip to the control
        frmMainClip.TextClip(TextCount).Text = Clipboard.GetText(vbCFText)
        
        'parse out the first 40 characters for display in list box
        If Len(frmMainClip.TextClip(TextCount).Text) > 40 Then
         TempText = Left$(frmMainClip.TextClip(TextCount).Text, 40) & "..."
        Else
         TempText = frmMainClip.TextClip(TextCount).Text
        End If
        
        If Left$(TempText, 2) = " &" Then TempText = "&&" + TempText
        Call frmMainClip.ucCoolList1.AddItem(TempText, 1, 1)
        
        'load the text into the viewer
        frmMainClip.Text1.Text = frmMainClip.TextClip(TextCount).Text
        
        'increase the count for next text clip
        TextCount = TextCount + 1
        
  '-----------------------------------------------------------------
      Case 2, 4, 6 'Bitmap Only
      
       ListTag(frmMainClip.ucCoolList1.ListCount) = "PIC" & Str(PicCount)
       
       Call frmMainClip.ucCoolList1.AddItem("<Graphic Image Clip>", 2, 2)
       
       DoEvents
       
       If PicCount > 0 Then
        Load frmMainClip.picClip(PicCount)
       End If
       
       frmMainClip.Image1.Stretch = False
       frmMainClip.picClip(PicCount).Picture = Clipboard.GetData()
       frmMainClip.picClip(PicCount).Picture = frmMainClip.picClip(PicCount).Picture
       frmMainClip.Image1.Picture = frmMainClip.picClip(PicCount).Picture
       frmMainClip.Reset_Image
       PicCount = PicCount + 1
       
  '---------------------------------------------------------------------
      Case 8, 9 'RichText
        
        ListTag(frmMainClip.ucCoolList1.ListCount) = "RTF" & Str(RTFTextCount)
        
        'firsttime through use main control else load a new control
        If RTFTextCount > 0 Then
         Load frmMainClip.RTFTextClip(RTFTextCount)
        End If
        
        'load the clip to the control
        frmMainClip.RTFTextClip(RTFTextCount).TextRTF = Clipboard.GetText(vbCFRTF)
        
        'parse out the first 40 characters for display in list box
        If Len(frmMainClip.RTFTextClip(RTFTextCount).Text) > 40 Then
         TempText = Left$(frmMainClip.RTFTextClip(RTFTextCount).Text, 40) & "..."
        Else
         TempText = frmMainClip.RTFTextClip(RTFTextCount).Text
        End If
        Call frmMainClip.ucCoolList1.AddItem(TempText, 3, 3)
        
        'load the text into the viewer
        'Text1.TextRTF = TextClip(TextCount).TextRTF
        frmMainClip.Text1.Text = Clipboard.GetText(vbCFText)
        
        'increase the count for next text clip
        RTFTextCount = RTFTextCount + 1
        
'-------------------------------------------------------------
      Case 15 'Files Only
       ListTag(frmMainClip.ucCoolList1.ListCount) = "FIL" & Str(TextCount)
        
        'firsttime through use main else load a new control
        If TextCount > 0 Then
         Load frmMainClip.TextClip(TextCount)
        End If

        'load the clip to the control
        frmMainClip.TextClip(TextCount).Text = Clipboard.GetText(vbCFText)
        
        'parse out the first 40 characters for display in list box
        If Len(frmMainClip.TextClip(TextCount).Text) > 40 Then
         TempText = Left$(frmMainClip.TextClip(TextCount).Text, 40) & "..."
        Else
         TempText = frmMainClip.TextClip(TextCount).Text
        End If
        
        Call frmMainClip.ucCoolList1.AddItem(TempText, 4, 4)
        
        'load the text into the viewer
        frmMainClip.Text1.Text = frmMainClip.TextClip(TextCount).Text
        
        'increase the count for next text clip
        TextCount = TextCount + 1
        
  '-----------------------------------------------------------------
        
      Case Else 'Clipboard is empty or program specific item in Clipboard

          Busy = False
          Exit Sub
           
   End Select
 '--------------------------------------------------------------
 
 'Play sound
  If iQSound = True And CapButton = False Then
   If frmMainClip.ucCoolList1.ListCount = 25 Then
    PlayWaveSound App.Path & "\25clips.wav"
    If frmMainClip.Visible = False Then frmMainClip.Visible = True
   Else
    PlayWaveSound App.Path & "\clipcopied.wav"
   End If
  End If
 
    'check to see if we are over max count.  if so remove oldest item (index 0)
  If frmMainClip.ucCoolList1.ListCount > 25 Then
    RemoveControl
  End If
  
    'set focus to newly added clip
  frmMainClip.picClipLabel.Visible = False
  frmMainClip.ucCoolList1.ListIndex = frmMainClip.ucCoolList1.ListCount - 1

  'clean-up clipboard for double paste problem
  ResetClipboard


  Busy = False
  DoEvents
  
 ' If this is max clip and showwarnings is true tell user one time
  If frmMainClip.ucCoolList1.ListCount = 25 And ShowWarnings = True And Warning25 = True Then
   Warning25 = False
   If frmMainClip.Visible = False Then frmMainClip.Visible = True
   TempText = "Warning: You have 25 clips in this Q-Clips Collection." & vbCrLf
   TempText = TempText & "The next clipboard capture will delete the oldest (1st) clip in this collection."
   MsgBox TempText, vbExclamation + vbOKOnly + vbApplicationModal, "Max Q-Clips In Collection"
  End If
   
End Sub


Public Function InitCommonControlsXP() As Boolean

On Error Resume Next

Dim iccex As tagInitCommonControlsEx


With iccex
  .lngSize = Len(iccex)
  .lngICC = ICC_USEREX_CLASSES
  
End With

InitCommonControlsEx iccex
InitCommonControlsXP = CBool(Err = 0)

End Function

Public Sub HookForm(f As Form)
   PrevProc = SetWindowLong(f.hwnd, GWL_WNDPROC, AddressOf WindowProc)
   IsHooked = True
End Sub
Private Sub RemoveControl()
 
 Dim TempText
 TempText = ListTag(0)
 j = Val(Mid$(ListTag(0), 4, Len(ListTag(0))))
 
 'If j is 0 then this is parent control and can't remove that
 If j > 0 Then
  'delete the control
    If Left$(TempText, 3) = "PIC" Then
     Unload frmMainClip.picClip(j)
    ElseIf Left$(TempText, 3) = "TXT" Then
     Unload frmMainClip.TextClip(j)
    ElseIf Left$(TempText, 3) = "RTF" Then
     Unload frmMainClip.RTFTextClip(j)
    End If
  End If
 
 'now remove listtag from list and item
 For i = 0 To frmMainClip.ucCoolList1.ListCount - 1
  ListTag(i) = ListTag(i + 1)
 Next
 
 frmMainClip.ucCoolList1.RemoveItem 0
 
End Sub

Public Sub UnHookForm(f As Form)
    SetWindowLong f.hwnd, GWL_WNDPROC, PrevProc
    IsHooked = False
    
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

 'This is sub-classed notification of Windows Clipboard activity
 
  WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)

  If uMsg = WM_CHANGECBCHAIN Then
    '
    ' If the window being removed is the next window in the
    ' chain, the window specified by the m_hWndNext parameter
    ' becomes the next window and clipboard messages are
    ' passed on to it.
    '
    If wParam = m_hWndNext Then
       m_hWndNext = lParam
    End If
    Call SendMessage(m_hWndNext, uMsg, wParam, lParam)
  End If

 ' Get the clip just copied to Windows Clipboard
  If uMsg = WM_DRAWCLIPBOARD And Busy = False And ClipboardOn Then

    GetClip
 
    Call SendMessage(m_hWndNext, uMsg, wParam, lParam)
 
  End If
    

End Function


Public Sub PlayWaveSound(WaveSound As String)
    On Error Resume Next
    iret = sndPlaySound(WaveSound, SND_ASYNC Or SND_NODEFAULT)
End Sub





