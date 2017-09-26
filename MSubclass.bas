Attribute VB_Name = "MSubclass"
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MSubclass.bas
'
' Subclassing Routines
'   - Dependencies: ISubclass.cls or VB6.tlb
'   - Adapted from HookXP by Karl E. Peterson - http://vb.mvps.org/samples/HookXP/
'   - No SetProp/GetProp bugs per SubTimer.dll(McKinney), SSubTmr6.dll(McMahon)
'   - No DEP issues or assembly thunks per PushParamThunk(Curland), SelfSub(Caton)
'   - No UserObject leaks, by using a single callback to forward
'       http://stackoverflow.com/questions/1120337/setwindowsubclass-is-leaking-user-objects
'
Option Explicit

Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As ISubclass, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As ISubclass) As Long
Public Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function SetSubclass(ByVal hWnd As Long, ByRef This As ISubclass, Optional ByVal dwRefData As Long) As Long
    SetSubclass = SetWindowSubclass(hWnd, AddressOf MSubclass.SubclassProc, This, dwRefData)
End Function

Public Function RemoveSubclass(ByVal hWnd As Long, ByRef This As ISubclass) As Long
    RemoveSubclass = RemoveWindowSubclass(hWnd, AddressOf MSubclass.SubclassProc, This)
End Function

Private Function SubclassProc(ByVal hWnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long, _
                              ByVal uIdSubclass As ISubclass, _
                              ByVal dwRefData As Long) As Long

    Const WM_NCDESTROY As Long = &H82&

    If uMsg = WM_NCDESTROY Then
        RemoveSubclass hWnd, uIdSubclass
        SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    Else
        SubclassProc = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
    End If

End Function
