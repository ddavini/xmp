Attribute VB_Name = "Module1"
Option Explicit
Private Sub Main()

On Error Resume Next

Dim cmD As String
Dim OverSubsystem As String
Dim CompileAsConsole As Boolean

cmD = Command$
CompileAsConsole = False

If InStr(1, cmD, "modCompileOptions.OBJ", vbTextCompare) Then
    CompileAsConsole = True
Else
    CompileAsConsole = False
End If

cmD = cmD & " " & GetIni(App.Path & "\vbLINK.INI", "PARAM", "VALUE", "")
OverSubsystem = GetIni(App.Path & "\vbLINK.INI", "MAIN", "SUBSYSTEM", "CONSOLE")

If CompileAsConsole Then
    MsgBox ("Application Console")
    cmD = Replace(cmD, "/SUBSYSTEM:WINDOWS,4.0", "/SUBSYSTEM:" & OverSubsystem)
    Shell "vbLink.exe " & cmD
Else
    Shell "vbLink.exe " & cmD
End If

Call SetIni(App.Path & "\cmdstr.txt", "LINK", "Command$", cmD)

End Sub
