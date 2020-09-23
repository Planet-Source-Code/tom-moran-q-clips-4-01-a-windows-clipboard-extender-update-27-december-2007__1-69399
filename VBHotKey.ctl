VERSION 5.00
Begin VB.UserControl VBHotKey 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VBHotKey.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   465
   ToolboxBitmap   =   "VBHotKey.ctx":1872
End
Attribute VB_Name = "VBHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_AltKey = False
Const m_def_ShiftKey = False
Const m_def_CtrlKey = False
Const m_def_WinKey = True
Const m_def_VKey = vbKeyA
Const m_def_Enabled = True

'Property Variables:
Dim m_AltKey As Boolean
Dim m_ShiftKey As Boolean
Dim m_CtrlKey As Boolean
Dim m_VKey As Long
Dim m_WinKey As Boolean
Dim m_Enabled As Boolean

Dim mHwnd As Long
Private mHotkey As ApiHotkey

' ##EVENT_DESCRIPTION HotkeyPressed - Fired when the registerd key combination is pressed
' ##EVENT_REMARKS This event fires whether or not your client application has the focus.
' ##EVENT_REMARKS This event will not bring your application to the foreground, so messages will _
not be shown until your application regains the focus
Event HotkeyPressed()

' ##EVENT_DESCRIPTION ErrorOccured - Fired when an error occurs creating a hotkey
' ##EVENT_REMARKS If you attempt to register an illegal combination or one that is _
already registered an error will be triggered.
Event ErrorOccured(ByVal Message As String, ByVal source As String)


Public Property Get About() As Long

End Property

Public Property Let About(ByVal dummy As Long)

If Not Ambient.UserMode Then
    If dummy <> 0 Then
        frmAbout.Show vbModal
    End If
End If

End Property


Public Property Let Enabled(ByVal newvalue As Boolean)

If newvalue <> m_Enabled Then
    m_Enabled = newvalue
    '\\ Notify the property bag that this member has changed.
    PropertyChanged "Enabled"
    '\\ Set or reset the hotkey
    With mHotkey
        If newvalue Then
            .Register
        Else
            .Unregister
        End If
    End With
End If

End Property


' ##BLOCK_DESCRIPTION Enabled - Returns/sets a value that determines whether an object can respond to user-generated events.
' ##BLOCK_REMARKS When enabled is false, another hotkey may use the same combination but no two ENABLED controls may share a single combination.
Public Property Get Enabled() As Boolean

Enabled = m_Enabled

End Property

Public Property Get hAtom() As Long

    hAtom = mHotkey.hAtom
    
End Property

Public Property Get hKey() As Long

    hKey = mHotkey.hKey
    
End Property


Friend Property Get hwnd() As Long

    hwnd = mHotkey.hwnd
    
End Property

Friend Sub RaiseError(ByVal sMessage As String, ByVal sSource As String)

RaiseEvent ErrorOccured(sMessage, sSource)

End Sub

Public Sub RaiseKeyPressEvent(ByVal VKey As Long, ByVal Modifier As Long)

Dim keyThis As ApiHotkey

Set keyThis = New ApiHotkey

With keyThis
    .VKey = VKey
    .Modifier = Modifier
    If .MatchedKey(mHotkey) And m_Enabled Then
        RaiseEvent HotkeyPressed
    End If
End With

Set keyThis = Nothing

End Sub

Public Sub StopHotkey()

If mHwnd <> 0 Then
    Call FreeSubclassedWindow(mHwnd)
    If Not (mHotkey Is Nothing) Then
        mHotkey.Unregister
        Set mHotkey = Nothing
    End If
End If
    
End Sub

Public Property Get UniqueKey() As String

UniqueKey = mHotkey.UniqueKey

End Property


Private Sub UserControl_Initialize()

Set mHotkey = New ApiHotkey

End Sub
' ##BLOCK_DESCRIPTION AltKey - Specifies whether the ALT key is included in the hotkey combination
Public Property Get AltKey() As Boolean
    AltKey = m_AltKey
End Property

Public Property Let AltKey(ByVal New_AltKey As Boolean)

If m_AltKey <> New_AltKey Then
    m_AltKey = New_AltKey
    PropertyChanged "AltKey"
    '\\ If the hotkey has been registered, unregister it and then reregister it
    If Ambient.UserMode Then
        If Me.hAtom <> 0 Then
            With mHotkey
                .Unregister
                .AltKey = m_AltKey
                .Register
            End With
        End If
    End If
End If

End Property

' ##BLOCK_DESCRIPTION ShiftKey - Specifies whether the SHIFT key is included in the hotkey combination
Public Property Get ShiftKey() As Boolean
    ShiftKey = m_ShiftKey
End Property

Public Property Let ShiftKey(ByVal New_ShiftKey As Boolean)
    
If New_ShiftKey <> m_ShiftKey Then
    m_ShiftKey = New_ShiftKey
    PropertyChanged "ShiftKey"
    '\\ If the hotkey has been registered, unregister it and then reregister it
    If Ambient.UserMode Then
        If Me.hAtom <> 0 Then
            With mHotkey
                .Unregister
                .ShiftKey = m_ShiftKey
                .Register
            End With
        End If
    End If
End If

End Property

' ##BLOCK_DESCRIPTION CtrlKey - Specifies whether the CONTROL key is included in the hotkey combination
Public Property Get CtrlKey() As Boolean
    CtrlKey = m_CtrlKey
End Property


' ##BLOCK_DESCRIPTION WinKey - Specifies whether the WIN key is included in the hotkey combination
Public Property Get WinKey() As Boolean
    WinKey = m_WinKey
End Property
Public Property Let CtrlKey(ByVal New_CtrlKey As Boolean)
    
If New_CtrlKey <> m_CtrlKey Then
    m_CtrlKey = New_CtrlKey
    PropertyChanged "CtrlKey"
    '\\ If the hotkey has been registered, unregister it and then reregister it
    If Ambient.UserMode Then
        If Me.hAtom <> 0 Then
            With mHotkey
                .Unregister
                .ControlKey = m_CtrlKey
                .Register
            End With
        End If
    End If
End If

End Property

Public Property Let WinKey(ByVal New_WinKey As Boolean)
    
If m_WinKey <> New_WinKey Then
    m_WinKey = New_WinKey
    PropertyChanged "WinKey"
    '\\ If the hotkey has been registered, unregister it and then reregister it
    If Ambient.UserMode Then
        If Me.hAtom <> 0 Then
            With mHotkey
                .Unregister
                .WinKey = m_WinKey
                .Register
            End With
        End If
    End If
End If

End Property

' ##BLOCK_DESCRIPTION VKey - Specifies which key fires the HotkeyPressed event
' ##BLOCK_REMARKS The F12 key is reserved for the system debugger in Windows 2000 and Windows XP.
' ##BLOCK_REMARKS It is not advisable to use this (or any other common hotkey combination)
Public Property Get VKey() As KeyCodeConstants
    VKey = m_VKey
End Property

Public Property Let VKey(ByVal New_VKey As KeyCodeConstants)
    
If New_VKey <> m_VKey Then
    m_VKey = New_VKey
    PropertyChanged "VKey"
    '\\ If the hotkey has been registered, unregister it and then reregister it
    If Ambient.UserMode Then
        If Me.hAtom <> 0 Then
            With mHotkey
                .Unregister
                .VKey = m_VKey
                .Register
            End With
        End If
    End If
End If

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()


    m_AltKey = m_def_AltKey
    m_ShiftKey = m_def_ShiftKey
    m_CtrlKey = m_def_CtrlKey
    m_VKey = m_def_VKey
    m_WinKey = m_def_WinKey
    m_Enabled = m_def_Enabled
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AltKey = PropBag.ReadProperty("AltKey", m_def_AltKey)
    m_ShiftKey = PropBag.ReadProperty("ShiftKey", m_def_ShiftKey)
    m_CtrlKey = PropBag.ReadProperty("CtrlKey", m_def_CtrlKey)
    m_VKey = PropBag.ReadProperty("VKey", m_def_VKey)
    m_WinKey = PropBag.ReadProperty("WinKey", m_def_WinKey)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    
    '\\ If we in RUN mode
    If Ambient.UserMode Then
        With UserControl
            If .hwnd > 0 Then
                '\\ Register the hotkey
                mHwnd = .hwnd
                With mHotkey
                    .AltKey = m_AltKey
                    .ShiftKey = m_ShiftKey
                    .ControlKey = m_CtrlKey
                    .WinKey = m_WinKey
                    .VKey = m_VKey
                    .hwnd = mHwnd
                    If m_Enabled Then
                        .Register
                    End If
                End With
                '\\ And subclass the parent to listen to it...
                Call SubclassWindow(mHwnd)
                If colControls Is Nothing Then
                    Set colControls = New Collection
                End If
                '\\ DEJ 18-June-2001 Keys must be unique per window only...
                colControls.Add Me, mHotkey.UniqueKey
            End If
        End With
    End If
End Sub

Private Sub UserControl_Resize()

UserControl.Width = 450
UserControl.Height = 450

End Sub

Private Sub UserControl_Terminate()

Me.Enabled = False
Call Me.StopHotkey

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

With PropBag
    Call .WriteProperty("AltKey", m_AltKey, m_def_AltKey)
    Call .WriteProperty("ShiftKey", m_ShiftKey, m_def_ShiftKey)
    Call .WriteProperty("CtrlKey", m_CtrlKey, m_def_CtrlKey)
    Call .WriteProperty("VKey", m_VKey, m_def_VKey)
    Call .WriteProperty("WinKey", m_WinKey, m_def_WinKey)
    Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End With

End Sub

