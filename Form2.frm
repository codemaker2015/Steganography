VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   43
      Left            =   16920
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   43
      Left            =   16935
      TabIndex        =   44
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   42
      Left            =   15360
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   42
      Left            =   15375
      TabIndex        =   43
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   41
      Left            =   13680
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   41
      Left            =   13695
      TabIndex        =   42
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   40
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   40
      Left            =   12015
      TabIndex        =   41
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   39
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   39
      Left            =   10335
      TabIndex        =   40
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   38
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   38
      Left            =   8655
      TabIndex        =   39
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   37
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   37
      Left            =   6975
      TabIndex        =   38
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   36
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   36
      Left            =   5295
      TabIndex        =   37
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   35
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   35
      Left            =   3615
      TabIndex        =   36
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   34
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   34
      Left            =   1935
      TabIndex        =   35
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   33
      Left            =   240
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   33
      Left            =   255
      TabIndex        =   34
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   32
      Left            =   17040
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   32
      Left            =   17055
      TabIndex        =   33
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   31
      Left            =   15360
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   31
      Left            =   15375
      TabIndex        =   32
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   30
      Left            =   13680
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   30
      Left            =   13695
      TabIndex        =   31
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   29
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   29
      Left            =   12015
      TabIndex        =   30
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   28
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   28
      Left            =   10335
      TabIndex        =   29
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   27
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   27
      Left            =   8655
      TabIndex        =   28
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   26
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   26
      Left            =   6975
      TabIndex        =   27
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   25
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   25
      Left            =   5295
      TabIndex        =   26
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   24
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   24
      Left            =   3615
      TabIndex        =   25
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   23
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   23
      Left            =   1935
      TabIndex        =   24
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   22
      Left            =   240
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   22
      Left            =   255
      TabIndex        =   23
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   21
      Left            =   17040
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   21
      Left            =   17055
      TabIndex        =   22
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   20
      Left            =   15360
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   20
      Left            =   15375
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   19
      Left            =   13680
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   19
      Left            =   13695
      TabIndex        =   20
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   18
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   18
      Left            =   12015
      TabIndex        =   19
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   17
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   17
      Left            =   10335
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   16
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   16
      Left            =   8655
      TabIndex        =   17
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   15
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   15
      Left            =   6975
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   13
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   14
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   14
      Left            =   5280
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   13
      Left            =   3615
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   11
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   12
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   12
      Left            =   1920
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   11
      Left            =   255
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   0
      Left            =   230
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   1
      Left            =   1910
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   2
      Left            =   3590
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   3
      Left            =   5270
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   4
      Left            =   6950
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   5
      Left            =   8630
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   6
      Left            =   10310
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   7
      Left            =   11990
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   8
      Left            =   13670
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   9
      Left            =   15350
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   10
      Left            =   17030
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   1910
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   3590
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   5270
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   6950
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   8630
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   10310
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   11990
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   13670
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   15350
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   10
      Left            =   17030
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image imgPrevious 
      Height          =   495
      Left            =   250
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   250
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
