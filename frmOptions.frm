VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line2 
      X1              =   0
      X2              =   1440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblCryptography 
      BackStyle       =   0  'Transparent
      Caption         =   "Cryptography"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblHide 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide / Show"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblSteganography 
      BackStyle       =   0  'Transparent
      Caption         =   "Steganography"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblHide_Click()
    If GetAttr(file) = 2 And vbHidden Then
        lblHide.Caption = "Hide"
        SetAttr file, vbNormal
    Else
        lblHide.Caption = "Show"
        SetAttr file, vbHidden
    End If
    Unload Me
End Sub

Private Sub lblSteganography_Click()
    Unload Me
    frmOption2.Left = frmOptions.Left + 1000
    frmOption2.Top = frmOptions.Top + 1000
    frmOption2.Show
End Sub
