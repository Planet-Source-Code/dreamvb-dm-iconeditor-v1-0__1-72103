Attribute VB_Name = "Tools"
Option Explicit

Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Public Const DrawArea = 321
Public Const ColSize = 10
Public AppData As String

Public Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    bReserved As Byte
End Type

Public Sub LongToRGB(LngColor As Long, RgbType As RGBQUAD)
On Error Resume Next
    'Convert Long Color To RGB
    RgbType.Red = (LngColor Mod 256)
    RgbType.Green = ((LngColor And &HFF00) / 256) Mod 256
    RgbType.Blue = ((LngColor And &HFF0000) / 65536)
End Sub

Public Sub DottedLine(DrawDir As Integer, DrawPos As Long, PicBox As PictureBox)
Dim Count As Long
Static xPos As Boolean
Dim col As Long

    'Draws a dotted line for the grid
    For Count = 0 To (DrawArea - 1)
        xPos = (Not xPos)
        
        If (xPos) Then
            col = RGB(92, 92, 92)
        Else
            col = RGB(192, 192, 192)
        End If
        
        If (DrawDir = 1) Then
            'Draw Accross
            SetPixelV PicBox.hdc, Count, 0, col
        End If
        
        If (DrawDir = 2) Then
            SetPixelV PicBox.hdc, 0, Count, col
        End If
    Next Count
    
End Sub

Public Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function GetFilePath(ByVal lPath As String) As String
Dim sPos As Integer
    'Return file path
    sPos = InStrRev(lPath, "\", Len(lPath), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetFilePath = Left$(lPath, sPos - 1)
    Else
        GetFilePath = lPath
    End If
End Function

Public Sub Mirror(pBox As PictureBox, Index As Integer)
    If (Index = 0) Then
        'Mirror left
        StretchBlt pBox.hdc, -pBox.Width, 0, -pBox.Width, pBox.Height, pBox.hdc, 0, 0, pBox.Width, pBox.Height, vbSrcCopy
        pBox.Picture = pBox.Image
        pBox.PaintPicture pBox, pBox.Width, 0, -pBox.Width / 2, pBox.Height, 0, 0, pBox.Width / 2
    Else
        'Mirror right
        StretchBlt pBox.hdc, pBox.Width, 0, -pBox.Width, pBox.Height, pBox.hdc, 0, 0, pBox.Width, pBox.Height, vbSrcCopy
        pBox.Picture = pBox.Image
        pBox.PaintPicture pBox, pBox.Width, 0, -pBox.Width / 2, pBox.Height, 0, 0, pBox.Width / 2
    End If
End Sub

