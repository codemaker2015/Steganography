Attribute VB_Name = "Module1"
Public drive As String
Public folder As String
Public file As String
Public folderpath(20) As String

Public Function geticon(extension As String) As String
    Select Case extension
        Case "jpeg", "png", "gif", "jpg": geticon = App.Path & "\images\img.jpg"
        Case "mp4", "avi", "3gp": geticon = App.Path & "\images\video.jpg"
        Case "mp3": geticon = App.Path & "\images\music.jpg"
        Case "docx", "doc": geticon = App.Path & "\images\word.jpg"
        Case "xlsx", "xls": geticon = App.Path & "\images\excel.jpg"
        Case "drive": geticon = App.Path & "\images\drive.jpg"
        Case "folder": geticon = App.Path & "\images\folder.jpg"
        Case Else: geticon = App.Path & "\images\other.jpg"
    End Select
End Function

Public Function getextension(FileName As String) As String
    Dim c As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        c = Mid(FileName, i, 1)
        If c = "." Then
            pos = i + 1
        End If
    Next i
    If pos = 0 Then
        getextension = ""
    Else
        getextension = Mid(FileName, pos, (Len(FileName) + 1 - pos))
    End If
End Function

Public Function getfilename(FileName As String) As String
    Dim c As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        c = Mid(FileName, i, 1)
        If c = "\" Then
            pos = i + 1
            Exit For
        End If
    Next i
    If pos = 0 Then
        getfilename = ""
    Else
        getfilename = Mid(FileName, pos, (Len(FileName) + 1 - pos))
    End If
End Function
