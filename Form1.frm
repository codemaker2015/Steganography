VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
' The Drive property also returns the volume label, so trim it.
Dir1.Path = Left$(Drive1.Drive, 1) & ":\"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub


Private Sub File1_Click()
    Dim FileName As String
    FileName = File1.Path
    If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
    FileName = FileName & File1.FileName
    MsgBox FileName
End Sub
