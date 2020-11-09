VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStegnographyImage 
   Caption         =   "Steganos - Stegnography"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picRecovered 
      Height          =   2655
      Left            =   7320
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox picCombined 
      Height          =   2655
      Left            =   4920
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox picHidden 
      Height          =   2655
      Left            =   2520
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox picVisible 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtNumBits 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "2"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Recovered Image"
      Height          =   195
      Index           =   3
      Left            =   7320
      TabIndex        =   6
      Top             =   720
      Width           =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Combined Image"
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hidden Image"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hidden Image Bits:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cover Image"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2280
   End
End
Attribute VB_Name = "frmStegnographyImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picVisible.AutoRedraw = True
    picHidden.AutoRedraw = True
    picCombined.AutoRedraw = True
    picRecovered.AutoRedraw = True
    
    picVisible.ScaleMode = vbPixels
    picHidden.ScaleMode = vbPixels
    picCombined.ScaleMode = vbPixels
    picRecovered.ScaleMode = vbPixels
    picVisible.Picture = LoadPicture(file)
End Sub

' Hide and then recover the image.
Private Sub cmdGo_Click()
Dim num_bits As Integer

    MousePointer = vbHourglass

    ' Hide the image.
    num_bits = Val(txtNumBits.Text)
    HideImage picVisible, picHidden, picCombined, num_bits

    ' Recover the hidden image.
    RecoverImage picCombined, picRecovered, num_bits

    MousePointer = vbDefault
End Sub

Private Sub picHidden_Click()
    CommonDialog1.ShowOpen

    picHidden.Picture = LoadPicture(CommonDialog1.FileName)
End Sub
