Option Explicit

' This is an inert awareness-training macro.
' It must be started manually. It does not execute commands, modify files,
' contact a network service, or configure an autorun event.
Sub TrainingSimulation
    Dim marker As String

    marker = "C:\vycvikove_stredisko.txt"
    If Not FileExistsSafe(marker) Then
        marker = "/vycvikove_stredisko.txt"
    End If

    If FileExistsSafe(marker) Then
        MsgBox "Training simulation condition matched." & Chr(13) & _
               "No command, network, or file-changing action was performed.", _
               64, "Security training"
    Else
        MsgBox "Training simulation is inactive on this machine." & Chr(13) & _
               "No action was performed.", _
               48, "Security training"
    End If
End Sub

Function FileExistsSafe(path As String) As Boolean
    On Error GoTo Missing
    Dim attributes As Integer
    attributes = GetAttr(path)
    FileExistsSafe = True
    Exit Function
Missing:
    FileExistsSafe = False
End Function
