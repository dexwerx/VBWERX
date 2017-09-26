VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWMCopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' Subclasses WM_COPYDATA
'   - Send/receive small blocks of data from one window to another
'   - used for inter-process communication via Windows' messaging subsystem
'
' Dependencies: VB6.tlb, VBSubclass.dll or MSubclass.bas/ISubclass.cls
'
Option Explicit

Public Event Receive(ByVal MsgType As Long, _
                     ByRef Message As String, _
                     ByVal hWndSrc As Long _
                     )

Private Type TCOPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Private Const WM_COPYDATA As Long = &H4A&

Implements ISubclass

Private m_hWnd As Long

Private Sub Class_Terminate()
    If m_hWnd Then RemoveSubclass m_hWnd, Me
End Sub

Private Function ISubclass_SubclassProc(ByVal hWnd As Long, _
                                        ByVal uMsg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long, _
                                        ByVal dwRefData As Long _
                                        ) As Long
    Dim cds As TCOPYDATASTRUCT
    Dim Message As String
    
    If uMsg = WM_COPYDATA Then
        CopyBytes LenB(cds), cds, ByVal lParam
        If cds.lpData And cds.cbData Then _
            Message = PtrStrLenB(cds.lpData, cds.cbData)
        RaiseEvent Receive(cds.dwData, Message, wParam)
        ISubclass_SubclassProc = API.True
    Else
        ISubclass_SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Public Function Init(ByVal hWnd As Long) As Boolean
    If SetSubclass(hWnd, Me) Then m_hWnd = hWnd: Init = True
End Function

Public Function Send(ByVal hWnd, _
                     ByVal MsgType As Long, _
                     ByRef Message As String _
                     ) As Boolean
    Dim cds As TCOPYDATASTRUCT
    
    cds.dwData = MsgType
    If LenB(Message) Then
        cds.cbData = LenB(Message)
        cds.lpData = StrPtr(Message)
    End If
    Send = SendMessage(hWnd, WM_COPYDATA, m_hWnd, cds)
End Function
