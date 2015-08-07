VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHashtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CHashtable.cls
'
' .NET Hashtable Class
'   - Collection Wrapper for CDictionaryEntry
'   - follows .NET implementation
'
Option Explicit

Public m_Hash As Collection

Private Function HashKey(Key As Variant) As String
    Select Case VarType(Key)
    Case vbString:  HashKey = Key
    Case vbObject:  HashKey = CStr(ObjPtr(Key))
    Case Else:      HashKey = CStr(Key)
    End Select
End Function

Private Property Let Var(Item As Variant, Value As Variant)
    If IsObject(Value) Then Set Item = Value Else Item = Value
End Property

Public Sub Add(Key As Variant, Value As Variant)
    Dim Entry As CDictionaryEntry
    Set Entry = New CDictionaryEntry
    m_Hash.Add Entry.Init(Key, Value), HashKey(Key)
End Sub

Public Sub Clear()
    Do While m_Hash.Count > 0
        m_Hash.Remove 1
    Loop
End Sub

Public Function Contains(Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

Public Function ContainsKey(Key As Variant) As Boolean
On Error Resume Next
    m_Hash.Item HashKey(Key)
    ContainsKey = (Err.Number = 0)
    Err.Clear
End Function

Public Function ContainsValue(Value As Variant) As Boolean
    Dim Entry As CDictionaryEntry
    For Each Entry In m_Hash
        If Value = Entry.Value Then ContainsValue = True: Exit Function
    Next
End Function

Public Property Get GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
   Set GetEnumerator = m_Hash.[_NewEnum]
End Property

Public Sub Remove(Key As Variant)
    m_Hash.Remove HashKey(Key)
End Sub

Public Property Get Count() As Long
    Count = m_Hash.Count
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If Not Contains(Key) Then Set Item = Nothing: Exit Property
    Dim Entry As CDictionaryEntry
    Set Entry = m_Hash(HashKey(Key))
    Var(Item) = Entry.Value
End Property

Public Property Let Item(Key As Variant, Value As Variant)
    If Contains(Key) Then m_Hash(HashKey(Key)).Value = Value Else Add Key, Value
End Property

Public Property Set Item(Key As Variant, Value As Variant)
    If Contains(Key) Then Set m_Hash(HashKey(Key)).Value = Value Else Add Key, Value
End Property

Public Property Get Keys() As Variant()
    Dim Index As Long
    Dim Entry As CDictionaryEntry
    ReDim RetVal(0 To m_Hash.Count - 1) As Variant
    For Each Entry In m_Hash
        Var(RetVal(Index)) = Entry.Key
        Index = Index + 1
    Next
    Keys = RetVal
End Property

Public Property Get Values() As Variant()
    Dim Index As Long
    Dim Entry As CDictionaryEntry
    ReDim RetVal(0 To m_Hash.Count - 1) As Variant
    For Each Entry In m_Hash
        Var(RetVal(Index)) = Entry.Value
        Index = Index + 1
    Next
    Values = RetVal
End Property

Private Sub Class_Initialize()
    Set m_Hash = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Hash = Nothing
End Sub
