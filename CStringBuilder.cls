VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CStringBuilder.cls
'
' Implements a StringBuilder class like .NET/Java
'   - based on .NET implementation
'   - Indexes are Zero based
'   - uses exponential growth of 1.5 Like STL/VC++
'   - overloaded default Chars property to access entire String like a Value property
'   - uses BSTR/Native VB6 Unicode strings internally
'
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (ByVal Destination As Long, ByVal Length As Long)

Private Const ERR_INVALIDPROCCALLORARG  As Long = 5     'VB Standard Errors
Private Const ERR_SUBSCRIPTOUTOFRANGE   As Long = 9
Private Const ERR_INVALIDPROPVALUE      As Long = 380

Private Const INITIAL_BUFFER_SIZE       As Long = 16    '.NET uses 16
Private Const GROWTH_FACTOR             As Double = 1.5 'Set to 2 to mimic .NET/Java
Private Const GROWTH_EXTRA              As Long = 0     'Set to 2 to mimic Java

Private m_Buffer As String
Private m_Length As Long
'

Private Sub Class_Initialize()
    m_Buffer = String$(INITIAL_BUFFER_SIZE, vbNullChar)
End Sub

Public Property Get Capacity() As Long
    Capacity = Len(m_Buffer)
End Property

Public Property Let Capacity(Value As Long)
    If Value = Len(m_Buffer) Then Exit Property
    If Value < 0 Then Err.Raise ERR_INVALIDPROPVALUE
    If Value < m_Length Then Err.Raise ERR_INVALIDPROCCALLORARG
    If Value > Len(m_Buffer) Then
        m_Buffer = m_Buffer & String$(Value - Len(m_Buffer), vbNullChar)
    Else
        m_Buffer = Left$(m_Buffer, Value)
    End If
End Property

Public Property Get Chars(Optional Index As Variant) As String
Attribute Chars.VB_UserMemId = 0
    If IsMissing(Index) Then
        Chars = Left$(m_Buffer, m_Length)
    Else
        If Index < 0 Or Index >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
        Chars = Mid$(m_Buffer, Index + 1, 1)
    End If
End Property

Public Property Let Chars(Optional Index As Variant, Value As String)
    If IsMissing(Index) Then
        If LenB(Value) = 0 Then
            m_Length = 0
            Exit Property
        End If
        EnsureCapacity Len(Value)
        m_Length = Len(Value)
        RtlMoveMemory StrPtr(m_Buffer), StrPtr(Value), LenB(Value)
    Else
        If Index < 0 Or Index >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
        RtlMoveMemory StrPtr(m_Buffer) + (Index * 2), StrPtr(Value), 2
    End If
End Property

Public Property Get Length() As Long
    Length = m_Length
End Property

Public Property Let Length(ByVal Value As Long)
    If m_Length = Value Then Exit Property
    If Value < 0 Then Err.Raise ERR_INVALIDPROPVALUE
    If Value > m_Length Then
        Dim PadLength As Long
        PadLength = Value
        If Value > Len(m_Buffer) Then PadLength = Len(m_Buffer)
        RtlZeroMemory StrPtr(m_Buffer) + m_Length * 2, (PadLength - m_Length) * 2
    End If
    EnsureCapacity Value
    m_Length = Value
End Property

Public Property Get MaxCapacity() As Long
Const INT32_MAXVALUE = &H7FFFFFFF
    MaxCapacity = INT32_MAXVALUE
End Property

Public Function Append(ByRef Value As String, Optional StartIndex As Long, Optional ByVal Count As Variant) As CStringBuilder
    If StartIndex < 0 Or StartIndex >= Len(Value) Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If IsMissing(Count) Then Count = Len(Value) - StartIndex
    If Count < 0 Then Err.Raise ERR_INVALIDPROCCALLORARG
    If StartIndex + Count > Len(Value) Then Err.Raise ERR_INVALIDPROCCALLORARG
    If Count > 0 Then
        EnsureCapacity m_Length + Count
        RtlMoveMemory StrPtr(m_Buffer) + m_Length * 2, StrPtr(Value) + StartIndex * 2, Count * 2
        m_Length = m_Length + Count
    End If
    Set Append = Me
End Function

Public Function AppendFormat(ByRef Format As String, ParamArray Args() As Variant) As CStringBuilder
    If Format = vbNullString Then Err.Raise ERR_INVALIDPROCCALLORARG
'TODO: Full implementation of .NET Format...
    If UBound(Args) >= 0 Then
        If IsArray(Args(0)) Then Args = Args(0)
        Dim FormatArg As Long
        For FormatArg = 0 To UBound(Args)
            Format = VBA.Strings.Replace(Format, "{" & CStr(FormatArg) & "}", CStr(Args(FormatArg)))
        Next
    End If
    Set AppendFormat = Append(Format)
End Function

Public Function AppendLine(Optional ByRef Value As String = vbNullString) As CStringBuilder
    Set AppendLine = Append(Value & vbNewLine)
End Function
 
Public Function Clear() As CStringBuilder
    m_Length = 0
    Set Clear = Me
End Function

Public Sub CopyTo(SourceIndex As Long, Destination As String, DestinationIndex As Long, Count As Long)
    If Destination = vbNullString Or Count < 0 Then Err.Raise ERR_INVALIDPROCCALLORARG
    If SourceIndex < 0 Or SourceIndex >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If DestinationIndex < 0 Or DestinationIndex >= Len(Destination) Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If SourceIndex + Count > m_Length Or DestinationIndex + Count > Len(Destination) Then Err.Raise ERR_INVALIDPROCCALLORARG
    RtlMoveMemory StrPtr(Destination) + DestinationIndex * 2, StrPtr(m_Buffer) + SourceIndex * 2, Count * 2
End Sub

Public Function EnsureCapacity(ByVal Capacity As Long) As Long
    If Capacity > Len(m_Buffer) Then
        Dim ExpandMin As Long, ExpandExp As Long
        ExpandMin = Capacity - Len(m_Buffer)                            'Requested Capacity
        ExpandExp = Len(m_Buffer) * (GROWTH_FACTOR - 1#) + GROWTH_EXTRA 'Capacity * GROWTH_FACTOR + EXTRA
        If ExpandExp > ExpandMin Then ExpandMin = ExpandExp             'Use the larger of the 2 (Java/.NET behavior)
        m_Buffer = m_Buffer & String$(ExpandMin, vbNullChar)
    End If
    EnsureCapacity = Len(m_Buffer)
End Function

Public Function Equals(Obj As Object) As Boolean
    If Obj Is Nothing Then
        Equals = False
    ElseIf TypeName(Obj) = TypeName(Me) Then
        Equals = (Capacity = Obj.Capacity) And (StrComp(ToString, Obj.ToString) = 0) And (MaxCapacity = Obj.MaxCapacity)
    Else
        Equals = Obj Is Me
    End If
End Function

Public Function Insert(ByVal Index As Long, ByRef Value As String, Optional ByVal Count As Long = 1) As CStringBuilder
    If Count < 1 Then Err.Raise ERR_INVALIDPROCCALLORARG
    If Index < 0 Or Index >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If LenB(Value) > 0 Then
        EnsureCapacity m_Length + Len(Value) * Count
        Dim StrPtrIndex As Long
        StrPtrIndex = StrPtr(m_Buffer) + Index * 2
        RtlMoveMemory StrPtrIndex + LenB(Value) * Count, StrPtrIndex, (m_Length - Index) * 2
        m_Length = m_Length + Len(Value) * Count
        Do Until Count = 0
            RtlMoveMemory StrPtrIndex + LenB(Value) * (Count - 1), StrPtr(Value), LenB(Value)
            Count = Count - 1
        Loop
    End If
    Set Insert = Me
End Function

Public Function Remove(ByVal StartIndex As Long, ByVal Length As Long) As CStringBuilder
    If StartIndex < 0 Or StartIndex >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If StartIndex + Length > m_Length Then Err.Raise ERR_INVALIDPROCCALLORARG
    Dim StrPtrIndex As Long
    StrPtrIndex = StrPtr(m_Buffer) + StartIndex * 2
    RtlMoveMemory StrPtrIndex, StrPtrIndex + Length * 2, (m_Length - (StartIndex + Length)) * 2
    m_Length = m_Length - Length
    Set Remove = Me
End Function

Public Function Replace(OldValue As String, _
                        NewValue As String, _
                        Optional StartIndex As Long, _
                        Optional Count As Long) As CStringBuilder
    If LenB(OldValue) = 0 Then Err.Raise ERR_INVALIDPROCCALLORARG
    If StartIndex < 0 Or StartIndex >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If Count = 0 Then Count = m_Length - StartIndex
    If StartIndex + Count > m_Length Then Err.Raise ERR_INVALIDPROCCALLORARG
    If Count > 0 And StrComp(OldValue, NewValue, vbBinaryCompare) <> 0 Then
        Dim Found As Long, StrPtrFound As Long, Expand As Long
        Found = InStr(StartIndex + 1, m_Buffer, OldValue, vbBinaryCompare)
        Do While Found > 0 And Found - 1 + Len(OldValue) <= m_Length
            Expand = Len(NewValue) - Len(OldValue)
            If Expand > 0 Then EnsureCapacity m_Length + Expand
            StrPtrFound = StrPtr(m_Buffer) + (Found - 1) * 2
            RtlMoveMemory StrPtrFound + LenB(NewValue), StrPtrFound + LenB(OldValue), _
                          (m_Length - (Found - 1)) * 2 - LenB(OldValue)
            If LenB(NewValue) > 0 Then RtlMoveMemory StrPtrFound, StrPtr(NewValue), LenB(NewValue)
            m_Length = m_Length + Expand
            Found = InStr(StartIndex + Found + Len(OldValue), m_Buffer, OldValue, vbBinaryCompare)
        Loop
    End If
    Set Replace = Me
End Function

Public Function ToString(Optional StartIndex As Variant, Optional ByVal Length As Variant) As String
    If IsMissing(StartIndex) Then
        If IsMissing(Length) Then Length = m_Length
        ToString = Left$(m_Buffer, Length)
        Exit Function
    End If
    If StartIndex < 0 Or StartIndex >= m_Length Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    If IsMissing(Length) Then Length = m_Length - StartIndex
    If Length < 0 Then Err.Raise ERR_INVALIDPROCCALLORARG
    If Length > 0 Then ToString = Mid$(m_Buffer, StartIndex + 1, Length)
End Function
