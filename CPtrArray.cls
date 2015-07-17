VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPtrArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CPtrArray.cls
'
' SafeArray Helper
'   - Point an uninitialized Array to an arbitrary Memory Address
'   - to pass in ArrayPtr use ArrPtr(): Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'   - TODO: multiple dimensions, use SafeArray API for descriptor allocation
'
Option Explicit

Private Type SAFEARRAYBOUND
    cElements       As Long
    lLbound         As Long
End Type

Private Type SAFEARRAY
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    rgsabound(0)    As SAFEARRAYBOUND
End Type

Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Dst As Long, ByVal Src As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Src As Long, ByRef Dst As Long)

Private InternalArr As SAFEARRAY
Private ExternalPtr As Long
'

Private Function PEEKD(Address As Long) As Long
    GetMem4 Address, PEEKD
End Function

Public Function Bind(ByVal ArrayPtr As Long, _
                     Optional ByVal BaseAddress As Long, _
                     Optional ByVal BaseIndex As Long = -1, _
                     Optional ByVal Count As Long = -1, _
                     Optional ByVal ElementSize As Long = -1) As CPtrArray
    If ExternalPtr Then Err.Raise 8, TypeName(Me), "This Instance already bound to an Array."
    If Count > 1 And ElementSize <= 0 Then Err.Raise 5, TypeName(Me), "ElementSize must be > 0 if Count is > 1."
    If PEEKD(ArrayPtr) Then Err.Raise 5, TypeName(Me), "Can't bind to an already Dimensioned Array."
    If BaseAddress Then InternalArr.pvData = BaseAddress
    If BaseIndex >= 0 Then InternalArr.rgsabound(0).lLbound = BaseIndex
    If Count >= 0 Then InternalArr.rgsabound(0).cElements = Count
    If ElementSize >= 0 Then InternalArr.cbElements = ElementSize
    PutMem4 ArrayPtr, VarPtr(InternalArr)
    ExternalPtr = ArrayPtr
    Set Bind = Me
End Function

Public Property Get ArrayPointer() As Long
Attribute ArrayPointer.VB_UserMemId = 0
    ArrayPointer = ExternalPtr
End Property

Public Property Let ArrayPointer(Value As Long)
    If Value = ExternalPtr Then Exit Property
    If Value Then Bind Value Else: UnBind
End Property

Public Property Get BaseAddress() As Long
    BaseAddress = InternalArr.pvData
End Property

Public Property Let BaseAddress(Value As Long)
    InternalArr.pvData = Value
End Property

Public Property Get BaseIndex() As Long
    BaseIndex = InternalArr.rgsabound(0).lLbound
End Property

Public Property Let BaseIndex(Value As Long)
    InternalArr.rgsabound(0).lLbound = BaseIndex
End Property

Public Property Get Count() As Long
    Count = InternalArr.rgsabound(0).cElements
End Property

Public Property Let Count(Value As Long)
    If Value > 1 And ElementSize <= 0 Then Err.Raise 380, TypeName(Me), "Cannot set Count > 1 without setting ElementSize."
    InternalArr.rgsabound(0).cElements = Count
End Property

Public Property Get ElementSize() As Long
    ElementSize = InternalArr.cbElements
End Property

Public Property Let ElementSize(Value As Long)
    InternalArr.cbElements = ElementSize
End Property

Public Function UnBind(Optional ByVal ArrayPtr As Long)
    If ArrayPtr = 0 Then ArrayPtr = ExternalPtr
    InternalArr.rgsabound(0).cElements = 0
    InternalArr.pvData = 0
    PutMem4 ArrayPtr, 0
    ExternalPtr = 0
End Function

Private Sub Class_Initialize()
    InternalArr.cDims = 1
    InternalArr.rgsabound(0).cElements = 1
End Sub

Private Sub Class_Terminate()
    If ExternalPtr Then UnBind
End Sub
