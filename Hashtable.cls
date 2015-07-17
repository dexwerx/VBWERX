VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hashtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Hash As Collection

Public Sub Add(Key As String, Value As Variant)
    Dim Entry As DictionaryEntry
    Set Entry = New DictionaryEntry
    Entry.Key = Key
    If IsObject(Value) Then Set Entry.Value = Value Else Entry.Value = Value
    m_Hash.Add Entry, Key
End Sub

Public Sub Clear()
    Do While m_Hash.Count > 0
        m_Hash.Remove 1
    Loop
End Sub

Public Function Contains(Key As String) As Boolean
    Contains = ContainsKey(Key)
End Function

Public Function ContainsKey(Key As String) As Boolean
On Error Resume Next
    m_Hash.Item Key
    ContainsKey = (Err.Number = 0)
    Err.Clear
End Function

Public Function ContainsValue(Value As Variant) As Boolean
    Dim Entry As DictionaryEntry
    For Each Entry In m_Hash
        If Value = Entry.Value Then ContainsValue = True: Exit Function
    Next
End Function

Public Property Get GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
   Set GetEnumerator = m_Hash.[_NewEnum]
End Property

Public Sub Remove(Key As String)
    m_Hash.Remove Key
End Sub

Public Property Get Count() As Long
    Count = m_Hash.Count
End Property

Public Property Get Item(Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Entry As DictionaryEntry
    Set Entry = m_Hash.Item(Key)
    If IsObject(Entry.Value) Then Set Item = Entry.Value Else Item = Entry.Value
End Property

Public Property Let Item(Key As String, Value As Variant)
    If Contains(Key) Then m_Hash.Item(Key).Value = Value Else Add Key, Value
End Property

Public Property Set Item(Key As String, Value As Variant)
    If Contains(Key) Then Set m_Hash.Item(Key).Value = Value Else Add Key, Value
End Property

Public Property Get Keys() As String()
    Dim Index As Long
    Dim Entry As DictionaryEntry
    Dim RetVal() As String
    ReDim RetVal(0 To m_Hash.Count - 1) As String
    For Each Entry In m_Hash
        RetVal(Index) = Entry.Key
        Index = Index + 1
    Next
    Keys = RetVal
End Property

Public Property Get Values() As Variant()
    Dim Index As Long
    Dim Entry As DictionaryEntry
    Dim RetVal() As Variant
    ReDim RetVal(0 To m_Hash.Count - 1) As Variant
    For Each Entry In m_Hash
        If IsObject(Entry.Value) Then Set RetVal(Index) = Entry.Value Else RetVal(Index) = Entry.Value
        Index = Index + 1
    Next
    Values = RetVal
End Property

Private Sub Class_Initialize()
    Set m_Hash = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set m_Hash = Nothing
End Sub
