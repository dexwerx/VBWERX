Attribute VB_Name = "MStartup"
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MStartup.bas
'
' Startup routine for VB6 Projects
'   - Startup Object: Sub Main fixes crash at startup when using a Manifest
'   - initiate loading shell32 prior to comctl32 to fix crashes at shutdown
'
Option Explicit

Private Declare Function InitShell Lib "shell32" Alias "IsUserAnAdmin" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub Main()
    InitShell
    InitCommonControls
    FMain.Show
End Sub
