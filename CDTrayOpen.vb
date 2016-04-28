sub open_cd_tray()
' create random variable to determine if cd tray should open
Dim jnumber As Integer

On Error Resume Next

If Application.UserName = "Justin Lansdale" Then
    jnumber = Int((100 - 0 + 1) * Rnd + 0)
        If jnumber < 6 Then
            Call OpenCDTray
        End If
End If


Dim jnumber As Integer

If Application.UserName = "Justin Lansdale" Then
jnumber = Int((100 - 0 + 1) * Rnd + 0)

    If jnumber < 30 Then
        Call OpenCDTray
        Worksheets("Users").Range("G" & j) = "Macro Run"
    End If
End If


Declare Sub mciSendStringA Lib "winmm.dll" (ByVal lpstrCommand As String, _
ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, _
ByVal hWndCallback As Long)
Option Private Module
Sub OpenCDTray()
    mciSendStringA "Set CDAudio Door Open", 0&, 0, 0
End Sub