Attribute VB_Name = "Program"
Option Explicit

Public Declare Sub Pause Lib "Kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Public Sub Pause2(t As Long)
    Dim i As Long
    Dim p As Long
    p = t / 100
    For i = 1 To p
        DoEvents
        Pause 100
    Next
End Sub

Public Sub Main()

    On Error GoTo catch
    
    Dim cmd As String
    Dim delay As Long
    Dim exeArg As String, exePath As String
    Dim i As Integer
    
    cmd = Command
'          12345678901234567
'    cmd = "7000 calc" 'TEST
    
    i = InStr(1, cmd, " ")
    
    If i > 0 Then
        delay = Val(Left(cmd, i))
        exeArg = Trim(Mid(cmd, i))
        Pause2 delay
        
        ' Promjena tekuæeg direktorijuma
        If Left(exeArg, 1) = """" Then
            i = InStr(1, exeArg, """ ")
        Else
            i = InStr(1, exeArg, " ")
        End If
        If i = 0 Then i = Len(exeArg)
        exePath = Left(exeArg, i)
        If Left(exePath, 1) = """" Then exePath = Mid(exePath, 2)
        If Right(exePath, 1) = """" Then exePath = Left(exePath, Len(exePath) - 1)
        i = InStrRev(exePath, "\")
        If i > 0 Then
            exePath = Left(exePath, i - 1)
            If Mid(exePath, 2, 1) = ":" Then ChDrive Left(exePath, 2)
            ChDir exePath
        End If
        
        ' Pokretanje izvršnog programa
        If exeArg <> "" Then Shell exeArg, vbHide
    
    Else
        If Dir("Help\delaycmd-en.html") <> "" Then
            Shell ("explorer Help\delaycmd-en.html")
        Else
            ShowMsg "Syntax: delaycmd pause command/program [parameters]" + vbCr + "(pause in milliseconds)"
        End If
    End If
    
Exit Sub
catch:
    If Len(exeArg) > 50 Then exeArg = "..." + Right(Replace(exeArg, """", ""), 50)
    ShowMsg Err.Description + "." + vbCrLf + exeArg
End Sub

Private Sub ShowMsg(msg As String)
    frmMsg.Show
    frmMsg.lblMsg.Caption = msg
End Sub
