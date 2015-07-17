VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String
Private m_Value As Variant
'

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(RHS As Variant)
    m_Value = RHS
End Property

Public Property Set Value(RHS As Variant)
    If IsObject(RHS) Then Set m_Value = RHS
End Property

Friend Property Let Key(RHS As String)
    m_Key = RHS
End Property

Public Property Get Key() As String
    Key = m_Key
End Property


