VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Dual (IDispatch) based Interface has different IID than VB6.tlb / VBSubclass.dll
' This interface is only needed when not referencing either VB6.tlb / VBSubclass.dll
Option Explicit
Public Function SubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
End Function
                              
                              
