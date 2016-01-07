' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MCollection.bas
'
' Collection Helper routines
'   - Dependancy CFixed.cls, VBXCore.tlb
'   - Get Collection Keys() Array
'   - KeyExists() function
'   - Coll() equivelent to Array()
'   - IsColl() equivelent to IsArray()
'   - Merge() appends Collection (does not preserve Indexes)
'   - Clone()
'
Option Explicit

Private Type VBCOLLECTION
    pVTable As Long
    Unk0    As Long
    Unk1    As Long
    Unk2    As Long
    Count   As Long
    Unk3    As Long
    First   As Long
    Last    As Long
End Type
Private Type VBCOLLECTIONITEM
    Data    As Variant
    Key     As String   '+16
    Prev    As Long     '+20
    Next    As Long     '+24
End Type

Public Function Keys(Col As Collection) As Variant()
    If Col.Count = 0 Then Exit Function
    Dim ColInternal As VBCOLLECTION
    CopyMemory ColInternal, ByVal ObjPtr(Col), LenB(ColInternal)
    ReDim RetKeys(1 To ColInternal.Count) As Variant
    Dim ColItem() As VBCOLLECTIONITEM
    With New CFixed: Call .Init(ArrPtr(ColItem), ColInternal.First)
        Dim ItemIndex As Long
        For ItemIndex = 1 To ColInternal.Count
            If (StrPtr(ColItem(0).Key)) Then
                RetKeys(ItemIndex) = ColItem(0).Key
            Else
                RetKeys(ItemIndex) = ItemIndex
            End If
            .Addr = ColItem(0).Next
        Next
    End With
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
    If Not IsObject(VarName) Then Exit Function
    IsColl = TypeOf VarName Is Collection
End Function

Public Function ContainsKey(Col As Collection, Key As Variant) As Boolean
On Error Resume Next
    Col.Item Key
    ContainsKey = (Err.Number = 0&)
    Err.Clear
End Function

Public Sub Merge(Dst As Collection, Src As Collection)
    If Src.Count = 0 Then Exit Sub
    Dim Key As Variant
    For Each Key In Keys(Src)
        Select Case VarType(Key)
        Case vbLong:    Dst.Add Src.Item(Key) 'Ignores Indexes
        Case vbString:  If ContainsKey(Dst, Key) Then Dst.Remove Key
                        Dst.Add Src.Item(Key), Key
        End Select
    Next
End Sub

Public Function Clone(Src As Collection) As Collection
    If Src Is Nothing Then Exit Function
    Set Clone = New Collection
    Dim Key, Keys()
    Keys = MCollection.Keys(Src)
    For Each Key In Keys
        Clone.Add Src.Item(Key), Key
    Next
End Function

