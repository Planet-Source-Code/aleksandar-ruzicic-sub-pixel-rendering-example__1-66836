Attribute VB_Name = "modSubPixel"
Option Explicit

Type RGBColor
    Red     As Integer
    Green   As Integer
    Blue    As Integer
End Type

Private Type Pixel
    x       As Long
    y       As Long
    Color   As Long
    Area    As Single
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'

Public Sub DrawPixel(DestHDC As Long, _
                     x As Long, y As Long, _
                     Optional Color As Long = 0)
                     
    SetPixel DestHDC, x, y, Color
    
    
End Sub
'

Public Sub DrawSubPixel(DestHDC As Long, _
                        x As Single, y As Single, _
                        Optional Color As Long = 0)
    
    Dim a   As Pixel
    Dim b   As Pixel
    Dim c   As Pixel
    Dim d   As Pixel
    
    a.x = Int(x)
    a.y = Int(y)
    a.Color = GetPixel(DestHDC, a.x, a.y)
    a.Area = (a.x - x + 1) * (a.y - y + 1)
    
    b.x = a.x + 1
    b.y = a.y
    b.Color = GetPixel(DestHDC, b.x, b.y)
    b.Area = (x - b.x + 1) * (a.y - y + 1)
    
    c.x = a.x
    c.y = a.y + 1
    c.Color = GetPixel(DestHDC, c.x, c.y)
    c.Area = (a.x - x + 1) * (y - c.y + 1)
    
    d.x = a.x + 1
    d.y = a.y + 1
    d.Color = GetPixel(DestHDC, d.x, d.y)
    d.Area = (x - b.x + 1) * (y - c.y + 1)
    
    DrawPixel DestHDC, a.x, a.y, BlendColor(a.Color, Color, a.Area)
    DrawPixel DestHDC, b.x, b.y, BlendColor(b.Color, Color, b.Area)
    DrawPixel DestHDC, c.x, c.y, BlendColor(c.Color, Color, c.Area)
    DrawPixel DestHDC, d.x, d.y, BlendColor(d.Color, Color, d.Area)
    
End Sub
'

Public Function BlendColor(base As Long, blend As Long, alpha As Single) As Long
        
    Dim rgbBase     As RGBColor
    Dim rgbBlend    As RGBColor
    
    rgbBase = CRGB(base)
    rgbBlend = CRGB(blend)
    
    With rgbBase
        
        ' because of this (blending) formula
        ' .Red, .Green and .Blue are integers (instead of Byte)
        '
        ' i'm not realy sure why vb complains when bytes are used...
        
        .Red = alpha * (rgbBlend.Red - .Red) + .Red
        .Green = alpha * (rgbBlend.Green - .Green) + .Green
        .Blue = alpha * (rgbBlend.Blue - .Blue) + .Blue
        
    End With
    
    BlendColor = CColor(rgbBase)
    
End Function
'

Public Function CRGB(Color As Long) As RGBColor
    
    Dim hexcode As String
    
    hexcode = Right$("000000" + Hex$(Color), 6)
    
    CRGB.Red = CByte(Val("&H" + Right$(hexcode, 2)))
    CRGB.Green = CByte(Val("&H" + Mid$(hexcode, 3, 2)))
    CRGB.Blue = CByte(Val("&H" + Left$(hexcode, 2)))
    
End Function
'

Public Function CColor(Color As RGBColor) As Long
    
    CColor = RGB(Color.Red, Color.Green, Color.Blue)

End Function
'
