VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStegnographyText 
   Caption         =   "Steganos - Steganography"
   ClientHeight    =   6780
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowPixels 
      Caption         =   "Show Pixels"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgImage 
      Left            =   240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Grpahic Files|*.bmp;*.gif;*.jpg|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "Secret Password"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Copyright 2018, Codemaker - GTec Kothamangalam"
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   1080
      Picture         =   "frmStegnographyText.frx":0000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   412
      TabIndex        =   0
      Top             =   840
      Width           =   6240
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Message"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmStegnographyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ArrangeControls()
Dim wid As Single

    Width = picImage.Left + picImage.Width + Width - ScaleWidth + 120
    Height = picImage.Top + picImage.Height + Height - ScaleHeight + 120
    wid = ScaleWidth - txtMessage.Left - 120
    If wid < 120 Then wid = 120
    txtMessage.Width = wid
    txtPassword.Width = wid
End Sub

' Encode this byte's data.
Private Sub EncodeByte(ByVal Value As Byte, ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByVal show_pixels As Boolean)
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

    byte_mask = 1
    For i = 1 To 8
        ' Pick a random pixel and RGB component.
        PickPosition used_positions, wid, hgt, r, c, pixel

        ' Get the pixel's color components.
        UnRGB picImage.Point(r, c), clrr, clrg, clrb
        If show_pixels Then
            clrr = 255
            clrg = clrg And &H1
            clrb = clrb And &H1
        End If

        ' Get the value we must store.
        If Value And byte_mask Then
            color_mask = 1
        Else
            color_mask = 0
        End If

        ' Update the color.
        Select Case pixel
            Case 0
                clrr = (clrr And &HFE) Or color_mask
            Case 1
                clrg = (clrg And &HFE) Or color_mask
            Case 2
                clrb = (clrb And &HFE) Or color_mask
        End Select

        ' Set the pixel's color.
        picImage.PSet (r, c), RGB(clrr, clrg, clrb)

        byte_mask = byte_mask * 2
    Next i
End Sub
' Decode this byte's data.
Private Function DecodeByte(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByVal show_pixels As Boolean) As Byte
Dim Value As Integer
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

    byte_mask = 1
    For i = 1 To 8
        ' Pick a random pixel and RGB component.
        PickPosition used_positions, wid, hgt, r, c, pixel

        ' Get the pixel's color components.
        UnRGB picImage.Point(r, c), clrr, clrg, clrb

        ' Get the stored value.
        Select Case pixel
            Case 0
                color_mask = (clrr And &H1)
            Case 1
                color_mask = (clrg And &H1)
            Case 2
                color_mask = (clrb And &H1)
        End Select

        If color_mask Then
            Value = Value Or byte_mask
        End If

        If show_pixels Then
            picImage.PSet (r, c), RGB( _
                clrr And &H1, _
                clrg And &H1, _
                clrb And &H1)
        End If

        byte_mask = byte_mask * 2
    Next i

    DecodeByte = CByte(Value)
End Function
' Translate a password into an offset value.
Private Function NumericPassword(ByVal password As String) As Long
Dim Value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    ' Initialize the shift values to different
    ' non-zero values.
    shift1 = 3
    shift2 = 17

    ' Process the message.
    str_len = Len(password)
    For i = 1 To str_len
        ' Add the next letter.
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function

' Pick an unused (r, c, pixel) combination.
Private Sub PickPosition(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByRef r As Integer, ByRef c As Integer, ByRef pixel As Integer)
Dim position_code As String

    On Error Resume Next
    Do
        ' Pick a position.
        r = Int(Rnd * wid)
        c = Int(Rnd * hgt)
        pixel = Int(Rnd * 3)

        ' See if the position is unused.
        position_code = "(" & r & "," & c & "," & pixel & ")"
        used_positions.Add position_code, position_code
        If Err.Number = 0 Then Exit Do
        Err.Clear
    Loop
End Sub
' Return the color's components.
Private Sub UnRGB(ByVal color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    r = color And &HFF&
    g = (color And &HFF00&) \ &H100&
    b = (color And &HFF0000) \ &H10000
End Sub

Private Sub cmdDecode_Click()
Dim msg_length As Byte
Dim msg As String
Dim ch As Byte
Dim i As Integer
Dim used_positions As Collection
Dim wid As Integer
Dim hgt As Integer
Dim show_pixels As Boolean

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Initialize the random number generator.
    Rnd -1
    Randomize NumericPassword(txtPassword.Text)

    wid = picImage.ScaleWidth
    hgt = picImage.ScaleHeight
    show_pixels = chkShowPixels.Value
    Set used_positions = New Collection

    ' Decode the message length.
    msg_length = DecodeByte(used_positions, wid, hgt, show_pixels)

    ' Decode the message.
    For i = 1 To msg_length
        ch = DecodeByte(used_positions, wid, hgt, show_pixels)
        msg = msg & Chr$(ch)
    Next i
    picImage.Picture = picImage.Image

    txtMessage.Text = msg

    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEncode_Click()
Dim msg As String
Dim i As Integer
Dim used_positions As Collection
Dim wid As Integer
Dim hgt As Integer
Dim show_pixels As Boolean

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Initialize the random number generator.
    Rnd -1
    Randomize NumericPassword(txtPassword.Text)

    wid = picImage.ScaleWidth
    hgt = picImage.ScaleHeight
    msg = Left$(txtMessage.Text, 255)
    show_pixels = chkShowPixels.Value
    Set used_positions = New Collection

    ' Encode the message length.
    EncodeByte CByte(Len(msg)), _
        used_positions, wid, hgt, show_pixels

    ' Encode the message.
    For i = 1 To Len(msg)
        EncodeByte Asc(Mid$(msg, i, 1)), _
            used_positions, wid, hgt, show_pixels
    Next i
    picImage.Picture = picImage.Image

    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    picImage.ScaleMode = vbPixels
    picImage.AutoRedraw = True
    dlgImage.InitDir = App.Path
    ArrangeControls
    picImage.Picture = LoadPicture(file)
End Sub


Private Sub mnuFileOpen_Click()
    On Error Resume Next
    dlgImage.CancelError = True
    dlgImage.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgImage.ShowOpen
    If Err.Number <> 0 Then Exit Sub

    picImage.Picture = LoadPicture(dlgImage.FileName)
    ArrangeControls
    If Err.Number <> 0 Then Exit Sub

    dlgImage.InitDir = dlgImage.FileName
    dlgImage.FileName = dlgImage.FileTitle
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error Resume Next
    dlgImage.CancelError = True
    dlgImage.Flags = _
        cdlOFNOverwritePrompt Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgImage.ShowSave
    If Err.Number <> 0 Then Exit Sub

    SavePicture picImage.Picture, dlgImage.FileName
    If Err.Number <> 0 Then Exit Sub

    dlgImage.InitDir = dlgImage.FileName
    dlgImage.FileName = dlgImage.FileTitle
End Sub


