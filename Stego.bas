Attribute VB_Name = "Stego"
Option Explicit

' Hide pic_hidden inside pic_visible and place the result in pic_result.
Public Sub HideImage(ByVal pic_visible As PictureBox, ByVal pic_hidden As PictureBox, ByVal pic_combined As PictureBox, ByVal hidden_bits As Integer)
Dim visible_mask As Integer
Dim hidden_mask As Integer
Dim shift As Integer
Dim x As Integer
Dim y As Integer
Dim r_visible As Byte
Dim g_visible As Byte
Dim b_visible As Byte
Dim r_hidden As Byte
Dim g_hidden As Byte
Dim b_hidden As Byte
Dim r_combined As Byte
Dim g_combined As Byte
Dim b_combined As Byte

    shift = 2 ^ (8 - hidden_bits)
    visible_mask = &HFF * 2 ^ hidden_bits
    hidden_mask = &HFF \ shift
    For x = 0 To pic_visible.ScaleWidth - 1
        For y = 0 To pic_visible.ScaleHeight - 1
            UnRGB pic_visible.Point(x, y), r_visible, g_visible, b_visible
            UnRGB pic_hidden.Point(x, y), r_hidden, g_hidden, b_hidden
            r_combined = (r_visible And visible_mask) + ((r_hidden \ shift) And hidden_mask)
            g_combined = (g_visible And visible_mask) + ((g_hidden \ shift) And hidden_mask)
            b_combined = (b_visible And visible_mask) + ((b_hidden \ shift) And hidden_mask)
            pic_combined.PSet (x, y), RGB(r_combined, g_combined, b_combined)
        Next y
    Next x
End Sub

' Recover a hidden image.
Public Sub RecoverImage(ByVal pic_combined As PictureBox, ByVal pic_recovered As PictureBox, ByVal hidden_bits As Integer)
Dim shift As Integer
Dim hidden_mask As Integer
Dim x As Integer
Dim y As Integer
Dim r_combined As Byte
Dim g_combined As Byte
Dim b_combined As Byte
Dim r_recovered As Byte
Dim g_recovered As Byte
Dim b_recovered As Byte
    
    shift = 2 ^ (8 - hidden_bits)
    hidden_mask = &HFF \ shift
    For x = 0 To pic_combined.ScaleWidth - 1
        For y = 0 To pic_combined.ScaleHeight - 1
            UnRGB pic_combined.Point(x, y), r_combined, g_combined, b_combined
            r_recovered = (r_combined And hidden_mask) * shift
            g_recovered = (g_combined And hidden_mask) * shift
            b_recovered = (b_combined And hidden_mask) * shift
            pic_recovered.PSet (x, y), RGB(r_recovered, g_recovered, b_recovered)
        Next y
    Next x
End Sub

' Break a color into red, green, and blue components.
Private Sub UnRGB(ByRef color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    r = color And &HFF&
    g = (color And &HFF00&) \ &H100&
    b = (color And &HFF0000) \ &H10000
End Sub
