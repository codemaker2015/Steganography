VERSION 5.00
Begin VB.Form frmOption2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblSteganographyText 
      Caption         =   "Steganography with Text"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblSteganographyImage 
      Caption         =   "Steganography with Image"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmOption2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblSteganographyImage_Click()
    Unload frmOptions
    Unload Me
    frmStegnographyImage.Show
End Sub

Private Sub lblSteganographyText_Click()
    Unload frmOptions
    Unload Me
    frmStegnographyText.Show
End Sub
