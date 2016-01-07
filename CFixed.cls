VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CFixed.cls
'
' Access Memory as an Array
'   - Dependancies: VBXCore.tlb
'   - Points a locked VB SAFEARRAY to arbitrary Memory locations
'   - to pass in ppsa use VBXCore.ArrPtr() or VarPtrArray/VarPtrStringArray
'   - Implemented as a class for RAII
'
Option Explicit

'Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
'Private Declare Function VarPtrStringArray Lib "msvbvm60" Alias "__vbaRefVarAry" (ByRef StringArray As Variant) As Long
'Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
'Private Const NULL_ As Long = 0&

Private Type SAFEARRAY1D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    cElements       As Long
    lLbound         As Long
End Type

Private m_sa    As SAFEARRAY1D
Private m_ppsa  As Long

Public Function Init(ByVal ppsa As Long, ByVal pdata As Long, Optional ByVal cb As Long, Optional ByVal celt As Long) As CFixed
    If cb Then m_sa.cbElements = cb
    If celt Then m_sa.cElements = celt
    m_sa.pvData = pdata
    m_ppsa = ppsa
    DeRef(m_ppsa) = VarPtr(m_sa)
    Set Init = Me
End Function

Private Sub Class_Initialize()
    m_sa.cLocks = 1
    m_sa.cDims = 1
    m_sa.cElements = 1
End Sub

Private Sub Class_Terminate()
    If m_ppsa Then _
      DeRef(m_ppsa) = NULL_
End Sub

Public Property Let Addr(Value As Long)
    m_sa.pvData = Value
End Property

Public Property Get Addr() As Long
    Addr = m_sa.pvData
End Property

'Public Sub Inc(Optional Count As Long = 1)
'    If Count = 1 Then m_sa.pvData = UAdd(m_sa.pvData, m_sa.cbElements): Exit Sub
'    m_sa.pvData = UAdd(m_sa.pvData, m_sa.cbElements * Count)
'End Sub
'
'Public Sub Dec(Optional Count As Long = 1)
'    If Count = 1 Then m_sa.pvData = UAdd(m_sa.pvData, -m_sa.cbElements): Exit Sub
'    m_sa.pvData = UAdd(m_sa.pvData, -m_sa.cbElements * Count)
'End Sub

'Private Property Get DeRef(Ptr As Long) As Long
'    Call GetMem4(ByVal Ptr, DeRef)
'End Property
'
'Private Property Let DeRef(Ptr As Long, Value As Long)
'    Call GetMem4(ByVal Ptr, Value)
'End Property
'
'Private Function UAdd(ByVal Ptr As Long, ByVal Offset As Long) As Long
'    UAdd = (Ptr Xor &H80000000) + Offset Xor &H80000000
'End Function
