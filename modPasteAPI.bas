Attribute VB_Name = "modPasteAPI"
Option Explicit

' Stuff mostly to find and paste to active window

   Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
   Public Const VK_CONTROL = &H11
   Public Const VK_V = &H56
   Public Const VK_LEFT = &H25
   Public Const KEYEVENTF_EXTENDEDKEY = &H1
   Public Const KEYEVENTF_KEYUP = &H2
   
   Public curWindow As Long
   Public ParentWindow As Long

   Public WinClassName As String
   Public ParentClassName As String

   Public WinCaption As String
   Public ParentCaption As String

   Public r As Long
   
   Public Type POINTAPI
        x As Long
        y As Long
   End Type
   
   Public mousePT As POINTAPI
  
   Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
   Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
   Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
   Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
   Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
   Public Declare Function GetActiveWindow% Lib "user32" ()
   Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
   
'Extra API's and constants
Public Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_REPLACESEL = &HC2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETFOCUS = &H7
Public Const GWW_HINSTANCE = (-6)

