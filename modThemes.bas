Attribute VB_Name = "modThemes"
Option Explicit

Private r1, g1, b1
Private r2, g2, b2

Public Themes As Integer



Sub DrawBlack()

 r2 = 240
 g2 = 240
 b2 = 240
 r1 = 58
 g1 = 58
 b1 = 58
 frmMainClip.Cls
 Call DrawGradient(frmMainClip)
 frmMainClip.cmdUnHook(0).ColorScheme = Custom
 frmMainClip.cmdUnHook(0).ColorButtonUp = &H404040
 frmMainClip.cmdUnHook(0).ColorButtonHover = &H808080
 frmMainClip.cmdUnHook(0).ColorButtonDown = &HC0C0C0
 
 frmMainClip.chkAutoSave.BackColor = 10790052
 frmMainClip.chkStartMinimized.BackColor = 10790052
 frmMainClip.chkSound.BackColor = 11711154
 frmMainClip.chkPasteMinimize.BackColor = 11711154
 frmMainClip.chkOnTop.BackColor = 12566463
 frmMainClip.chkWarnings.BackColor = 12566463
 frmMainClip.chkAutoHover.BackColor = 15527148

 
 'change List Hover Color
 frmMainClip.ucCoolList1.BackSelected = &HE0E0E0
 frmMainClip.ucCoolList1.BackSelectedG1 = &HE0E0E0
 frmMainClip.ucCoolList1.BackSelectedG2 = &HC0C0C0
 frmMainClip.ucCoolList1.FontNormal = vbBlack

End Sub


Sub DrawBlue()

 r1 = 255
 g1 = 255
 b1 = 255
 r2 = 161
 g2 = 192
 b2 = 236
 frmMainClip.Cls
 Call DrawGradient(frmMainClip)
 
'Change the button Color
  frmMainClip.cmdUnHook(0).ColorScheme = WMP10
 
'Change the checkbox background color
 frmMainClip.chkAutoSave.BackColor = &HF4DAC8
 frmMainClip.chkStartMinimized.BackColor = &HF4DAC8
 frmMainClip.chkSound.BackColor = &HF2D5C0
 frmMainClip.chkPasteMinimize.BackColor = &HF2D5C0
 frmMainClip.chkOnTop.BackColor = &HF1D0B9
 frmMainClip.chkWarnings.BackColor = &HF1D0B9
 frmMainClip.chkAutoHover.BackColor = &HECC4A7

 
'change List Hover Color
 frmMainClip.ucCoolList1.BackSelected = &HFAEFE7
 frmMainClip.ucCoolList1.BackSelectedG1 = &HFAEFE7
 frmMainClip.ucCoolList1.BackSelectedG2 = &HEDC1A2
 frmMainClip.ucCoolList1.FontNormal = 8388608
End Sub


Sub DrawGradient(Mee As Form)
Dim r, G, B

Mee.Cls
For i = 0 To 100
    r = r1 - (((r1 - r2) / 100) * i)
    G = g1 - (((g1 - g2) / 100) * i)
    B = b1 - (((b1 - b2) / 100) * i)
    Mee.Line (0, (Mee.ScaleHeight / 100) * i)-(Mee.ScaleWidth, (Mee.ScaleHeight / 100) * i), RGB(r, G, B)
Next i
End Sub
Sub DrawNone()

 
 frmMainClip.BackColor = &H8000000F
 frmMainClip.Cls

 
'Change the button Color
 frmMainClip.cmdUnHook(0).ColorScheme = Custom
 frmMainClip.cmdUnHook(0).ColorButtonUp = &HC0C0C0
 frmMainClip.cmdUnHook(0).ColorButtonHover = &H808080
 frmMainClip.cmdUnHook(0).ColorButtonDown = &HC0C0C0
 
'Change the checkbox background color
 frmMainClip.chkAutoSave.BackColor = &H8000000F
 frmMainClip.chkStartMinimized.BackColor = &H8000000F
 frmMainClip.chkSound.BackColor = &H8000000F
 frmMainClip.chkPasteMinimize.BackColor = &H8000000F
 frmMainClip.chkOnTop.BackColor = &H8000000F
 frmMainClip.chkWarnings.BackColor = &H8000000F
 frmMainClip.chkAutoHover.BackColor = &H8000000F

 
 'change List Hover Color
 frmMainClip.ucCoolList1.BackSelected = 15718086
 frmMainClip.ucCoolList1.BackSelectedG1 = 15718086
 frmMainClip.ucCoolList1.BackSelectedG2 = 15718086
 frmMainClip.ucCoolList1.FontNormal = vbBlack
End Sub

Sub DrawOlive()
 
 r1 = 255
 g1 = 255
 b1 = 255
 r2 = 188
 g2 = 203
 b2 = 153
 frmMainClip.Cls
 Call DrawGradient(frmMainClip)
 frmMainClip.cmdUnHook(0).ColorScheme = Custom
 frmMainClip.cmdUnHook(0).ColorButtonUp = &H436D5E
 frmMainClip.cmdUnHook(0).ColorButtonHover = &H6AA48F
 frmMainClip.cmdUnHook(0).ColorButtonDown = 10144957
 
 frmMainClip.chkAutoSave.BackColor = 12771543
 frmMainClip.chkStartMinimized.BackColor = 12771543
 frmMainClip.chkSound.BackColor = 12311762
 frmMainClip.chkPasteMinimize.BackColor = 12311762
 frmMainClip.chkOnTop.BackColor = 11852237
 frmMainClip.chkWarnings.BackColor = 11852237
 frmMainClip.chkAutoHover.BackColor = 10144957

 
  'change List Hover Color
 frmMainClip.ucCoolList1.BackSelected = &H6AA48F
 frmMainClip.ucCoolList1.BackSelectedG1 = &H6AA48F
 frmMainClip.ucCoolList1.BackSelectedG2 = 10144957
 frmMainClip.ucCoolList1.FontNormal = &H305045
 
 
End Sub

Sub DrawSilver()

 r1 = 255
 g1 = 255
 b1 = 255
 r2 = 168
 g2 = 167
 b2 = 191
 frmMainClip.Cls
 Call DrawGradient(frmMainClip)
 frmMainClip.cmdUnHook(0).ColorScheme = Custom
 frmMainClip.cmdUnHook(0).ColorButtonUp = &H404040
 frmMainClip.cmdUnHook(0).ColorButtonHover = &H808080
 frmMainClip.cmdUnHook(0).ColorButtonDown = &HC0C0C0
 
 frmMainClip.chkAutoSave.BackColor = 14273484
 frmMainClip.chkStartMinimized.BackColor = 14273484
 frmMainClip.chkSound.BackColor = 13944005
 frmMainClip.chkPasteMinimize.BackColor = 13944005
 frmMainClip.chkOnTop.BackColor = 13680319
 frmMainClip.chkWarnings.BackColor = 13680319
 frmMainClip.chkAutoHover.BackColor = 12626346
 
 'change List Hover Color
 frmMainClip.ucCoolList1.BackSelected = 14273484
 frmMainClip.ucCoolList1.BackSelectedG1 = 15656421
 frmMainClip.ucCoolList1.BackSelectedG2 = 14273484
 frmMainClip.ucCoolList1.FontNormal = vbBlack
End Sub


