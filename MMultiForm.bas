Attribute VB_Name = "MMultiForm"
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MMultiForm.bas
'
' Implements multiple duplicate forms for MS Access
'   - adapted from http://allenbrowne.com/ser-35.html
'   - Instatiate new form using 'OpenFormMulti New Form_frmEditRecord, "ID=" & ID'
'   - Call 'CloseFormMulti Me' from Form_Close() event, or you will get weird bugs...
'   - Use generic 'DoCmd.Close' to close the active multi form
'   - Do not use explicit 'DoCmd.Close acForm, Me.Name, acSaveNo'
'
Option Compare Database
Option Explicit

Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Const CASCADEX As Long = 400    'offset in twips
Const CASCADEY As Long = 400

Private MultiForms As New Collection

Function OpenFormMulti(NewForm As Form, _
                       Optional Filter As String = "", _
                       Optional Cascade As Boolean = True, _
                       Optional Caption As String) As Form
    MultiForms.Add NewForm, CStr(NewForm.hWnd)

    With NewForm
        .AllowAdditions = Not (Len(Filter) > 0)
        .DataEntry = .AllowAdditions
        If Not .AllowAdditions Then
            .Filter = Filter
            .FilterOn = True
        End If
        If LenB(Caption) Then .Caption = Caption
        .Visible = True
        .SetFocus
    End With
    
    If Cascade And MultiForms.Count > 1 And IsZoomed(NewForm.hWnd) = 0 And IsIconic(NewForm.hWnd) = 0 Then
        With MultiForms(MultiForms.Count - 1)
            NewForm.Move .WindowLeft + CASCADEX, .WindowTop + CASCADEY
        End With
    End If
    
    Set OpenFormMulti = NewForm
End Function

Function CloseFormMulti(CloseForm As Form) As Boolean
    Dim CurForm As Object
    
    If CloseForm Is Nothing Then Exit Function
    
    For Each CurForm In MultiForms
        If CurForm.hWnd = CloseForm.hWnd Then
            MultiForms.Remove CStr(CloseForm.hWnd)
            Set CloseForm = Nothing
            Set CurForm = Nothing
            Exit Function
        End If
    Next
End Function

Function CloseAllFormMulti(Optional FormTypeName As String)
    Dim CurForm As Object
    
    For Each CurForm In MultiForms
        If Len(FormTypeName) = 0 Then
            MultiForms.Remove CStr(CurForm.hWnd)
            Set CurForm = Nothing
        ElseIf TypeName(CurForm) = FormTypeName Then
            MultiForms.Remove CStr(CurForm.hWnd)
            Set CurForm = Nothing
        End If
    Next
End Function
