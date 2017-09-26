VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' AryMap.cls
'
' Access Arbitrary Memory locations using a Dynamic Array
'   - Dependencies: VB6.tlb
'   - allows VB to use pointers comparable in speed to C/C++
'   - PreDeclared Global Class, Default method is Bind()
'   - Use BasePtr property to change where the Array Points
'   - "Assume No Aliasing" optimization should be off
'   - to pass in ArrayVarPtr use RefAry/RefVarAry/VarPtrArray/VarPtrStringArray
'
' TODO: multiple dimensions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const FADF_AUTO         As Integer = &H1
Private Const FADF_FIXEDSIZE    As Integer = &H10

Private Type TSAFEARRAYBOUND
    cElements       As Long
    lLbound         As Long
End Type

Private Type TSAFEARRAY
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    rgsabound(0)    As TSAFEARRAYBOUND
End Type

Private m_sa    As TSAFEARRAY
Private m_ppsa  As Long

Implements IDisposable

Private Sub IDisposable_Dispose()
    Dispose True
End Sub

Public Sub Dispose(Optional Disposing As Boolean = True)
    If m_ppsa Then
        DeRef(m_ppsa) = vbNullPtr
        m_ppsa = vbNullPtr
    End If
End Sub

Public Function Bind(ByVal ArrayVarPtr As Long, _
            Optional ByVal BasePtr As Long, _
            Optional ByVal ElementSize As Long, _
            Optional ByVal ElementCount As Long = 1, _
            Optional ByVal LowerBound As Long) As AryMap
Attribute Bind.VB_Description = "Bind a Dynamic Array to a New AryMap"
Attribute Bind.VB_UserMemId = 0

    If ArrayVarPtr = vbNullPtr Then
        Err.Raise 5, "AryMap.Bind()", _
            "ArrayVarPtr must point to a valid Array."
    ElseIf DeRef(ArrayVarPtr) Then
        Err.Raise 5, "AryMap.Bind()", _
            "Array is already mapped or dimensioned."
    ElseIf ElementCount > 1 And ElementSize <= 0 Then
        Err.Raise 5, "AryMap.Bind()", _
            "Can't set ElementCount > 1 without setting ElementSize > 0."
    End If

    If AryMap Is Me Then        ' this is the Predeclared instance
        Set Bind = New AryMap
        Bind.Bind ArrayVarPtr, BasePtr, ElementSize, ElementCount, LowerBound
    Else
        Set Bind = Me
        m_sa.cbElements = ElementSize
        m_sa.rgsabound(0).cElements = ElementCount
        m_sa.rgsabound(0).lLbound = LowerBound
        m_sa.pvData = BasePtr
        m_ppsa = ArrayVarPtr
        If m_ppsa Then DeRef(m_ppsa) = Ref(m_sa)
    End If
End Function

Private Sub Class_Initialize()
    m_sa.cDims = 1
    m_sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
    m_sa.cLocks = 1
End Sub

Private Sub Class_Terminate()
    Dispose False
End Sub

Public Property Get BasePtr() As Long
    BasePtr = m_sa.pvData
End Property

Public Property Let BasePtr(ByVal Value As Long)
Attribute BasePtr.VB_Description = "Map Array to new Memory location"
    m_sa.pvData = Value
End Property


