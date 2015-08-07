VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CDictionaryEntry.cls
'
' .NET DictionaryEntry Class
'   - Used by CHashtable
'
Option Explicit

Private m_Key As Variant
Private m_Value As Variant

Private Property Let Var(Item, Value)
    If IsObject(Value) Then Set Item = Value Else Item = Value
End Property

Friend Function Init(Key As Variant, Value As Variant) As CDictionaryEntry
    Var(m_Key) = Key
    Var(m_Value) = Value
    Set Init = Me
End Function

Public Property Get Key() As Variant
    Var(Key) = m_Key
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Var(Value) = m_Value
End Property

Public Property Let Value(NewValue As Variant)
    m_Value = NewValue
End Property

Public Property Set Value(NewValue As Variant)
    Set m_Value = NewValue
End Property



