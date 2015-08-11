Attribute VB_Name = "MCollection"
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MCollection.cls
'
' Collection Helper routines
'   - Get Collection Keys() Array
'   - Key Exists() function
'   - Coll() equivelent to Array()
'   - IsColl() equivelent to IsArray()
'   - Merge() appends Collection (does not preserve
'
Option Explicit

Private Declare Sub PutMem4 Lib "msvbvm60" (Dst() As Any, Src As Any)

Private Type VBCOLLECTION
    pVTable     As Long
    Unk0        As Long
    Unk1        As Long
    Unk2        As Long
    Count       As Long
    Unk3        As Long
    First       As Long
    Last        As Long
End Type
 
Private Type VBCOLLECTIONITEM
    Data        As Variant
    Key         As String
    Prev        As Long
    Next        As Long
End Type

Private Type SAFEARRAY1D
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements   As Long
    lLbound     As Long
End Type

Public Function Keys(Col As Collection) As Variant()
    Dim ColInternalSA   As SAFEARRAY1D
    Dim ColInternal()   As VBCOLLECTION
    Dim ColItemSA       As SAFEARRAY1D
    Dim ColItem()       As VBCOLLECTIONITEM
    If Col.Count = 0 Then Exit Function
    ColInternalSA.cDims = 1
    ColInternalSA.cElements = 1
    ColInternalSA.pvData = ObjPtr(Col)
    PutMem4 ColInternal, ColInternalSA
    ReDim RetKeys(0 To ColInternal(0).Count - 1) As Variant
    ColItemSA.cDims = 1
    ColItemSA.cElements = 1
    ColItemSA.pvData = ColInternal(0).First
    PutMem4 ColItem, ColItemSA
    Dim ItemIndex As Long
    For ItemIndex = 0 To ColInternal(0).Count - 1
        If StrPtr(ColItem(0).Key) Then
            RetKeys(ItemIndex) = ColItem(0).Key
        Else
            RetKeys(ItemIndex) = ItemIndex
        End If
        ColItemSA.pvData = ColItem(0).Next
    Next
    PutMem4 ColInternal, ByVal 0&
    PutMem4 ColItem, ByVal 0&
    Keys = RetKeys
End Function

Public Function Coll(ParamArray ArgList() As Variant) As Collection
    If UBound(ArgList) + 1 And 1 Then Err.Raise 5, "Coll()", "Invalid Number of Arguments"
    Set Coll = New Collection
    Dim i As Long
    For i = 0 To UBound(ArgList) - 1 Step 2
        Coll.Add ArgList(i + 1), ArgList(i)
    Next
End Function

Public Function IsColl(VarName As Variant) As Boolean
    If VarName Is Nothing Then Exit Function
    IsColl = TypeOf VarName Is Collection
End Function

Public Function Exists(Key As Variant, InCol As Collection) As Boolean
On Error Resume Next
    InCol.Item Key
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub Merge(Dst As Collection, Src As Collection)
    If Src.Count = 0 Then Exit Sub
    Dim Key As Variant
    For Each Key In Keys(Src)
        Select Case VarType(Key)
        Case vbLong:    Dst.Add Src.Item(Key) 'Ignores Indexes
        Case vbString:  If Exists(Key, Dst) Then Dst.Remove Key
                        Dst.Add Src.Item(Key), Key
        End Select
    Next
End Sub
