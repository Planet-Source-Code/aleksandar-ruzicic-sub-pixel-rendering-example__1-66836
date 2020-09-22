VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DrawSubPixel demo"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5370
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   13
      ToolTipText     =   "Click to change drawing color"
      Top             =   1770
      Width           =   255
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   90
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   12
      Top             =   75
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5370
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      ToolTipText     =   "Click to change drawing color"
      Top             =   2130
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4455
      TabIndex        =   8
      Top             =   3900
      Width           =   1770
   End
   Begin VB.OptionButton optLine 
      Caption         =   "Diagonal line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   4515
      TabIndex        =   7
      Top             =   3525
      Width           =   1560
   End
   Begin VB.OptionButton optLine 
      Caption         =   "Sine wave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   4515
      TabIndex        =   6
      Top             =   3210
      Width           =   1560
   End
   Begin VB.OptionButton optLine 
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4515
      TabIndex        =   5
      Top             =   2910
      Value           =   -1  'True
      Width           =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Draw with sub-pixels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4455
      TabIndex        =   3
      ToolTipText     =   "Click to draw line using sub-pixels"
      Top             =   645
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4455
      TabIndex        =   2
      ToolTipText     =   "Click to clear picture"
      Top             =   1185
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4455
      TabIndex        =   1
      ToolTipText     =   "Click to draw line using ""normal"" pixels"
      Top             =   105
      Width           =   1770
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   90
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   0
      Top             =   75
      Width           =   4260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Back color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4470
      TabIndex        =   15
      Top             =   1785
      Width           =   1155
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5685
      TabIndex        =   14
      Top             =   1785
      Width           =   660
   End
   Begin VB.Label lblColor 
      BackStyle       =   0  'Transparent
      Caption         =   "00FF00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5685
      TabIndex        =   11
      Top             =   2145
      Width           =   660
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Draw color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4470
      TabIndex        =   9
      Top             =   2145
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Line:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4470
      TabIndex        =   4
      Top             =   2580
      Width           =   1740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'

Private Sub DrawLine(LineType As Byte, subPixels As Boolean)
    
    Dim x   As Single
    Dim y   As Single
    
    Dim lblX As Long
    Dim lblY As Long
    
    Dim l   As Long
    
    Const waitText = "drawing, please wait..."
    
    lblX = 140 - picCanvas.TextWidth(waitText) / 2
    lblY = 140 - picCanvas.TextHeight(waitText) / 2
    
    copyToBuffer
    
    picCanvas.Cls
    picCanvas.ForeColor = picColor.BackColor
    picCanvas.PSet (lblX, lblY), picBack.BackColor
    picCanvas.Print waitText
    DoEvents
    
    Select Case LineType
        
        Case 0 ' circle
            
            For l = 0 To 360
                
                x = Cos(l * 0.01745329) ' 0.01745329 = pi / 180
                y = Sin(l * 0.01745329)
                
                x = x * 50 + 140 ' circle radius = 50
                y = y * 50 + IIf(subPixels, 200, 70)
                
                If subPixels Then
                    
                    DrawSubPixel picBuffer.hdc, x, y, picColor.BackColor
                    
                Else
                    
                    DrawPixel picBuffer.hdc, CLng(x), CLng(y), picColor.BackColor
                
                End If
                
            Next
        
        Case 1 ' sine wave
            
            For l = 0 To 360
                
                x = (l / 360) * 260 + 10
                y = Sin(l * 0.01745329) * 50 + IIf(subPixels, 210, 70)
                
                If subPixels Then
                    
                    x = 280 - x ' horisontal flip the wave
                    
                    DrawSubPixel picBuffer.hdc, x, y, picColor.BackColor
                    
                Else
                    
                    DrawPixel picBuffer.hdc, CLng(x), CLng(y), picColor.BackColor
                
                End If
                
            Next
        
        
        Case 2 ' diagonal line
        
            For x = 10 To 270
                
                If subPixels Then
                    
                    y = 270 - x * 0.5 ' horisontal flip line
                    
                    DrawSubPixel picBuffer.hdc, x, y, picColor.BackColor
                    
                Else
                    
                    y = x * 0.5 ' 0.5 to make it more jagged
                    
                    DrawPixel picBuffer.hdc, CLng(x), CLng(y), picColor.BackColor
                
                End If
            
            Next
    
    End Select
    
    copyFromBuffer
    
End Sub
'

Private Sub copyToBuffer()
    
    BitBlt picBuffer.hdc, 0, 0, 280, 280, picCanvas.hdc, 0, 0, vbSrcCopy

End Sub
'

Private Sub copyFromBuffer()
    
    BitBlt picCanvas.hdc, 0, 0, 280, 280, picBuffer.hdc, 0, 0, vbSrcCopy
    
    picCanvas.Refresh
    
End Sub
'

Private Function selectedLine() As Byte
    
    selectedLine = IIf(optLine(0).Value, 0, _
                   IIf(optLine(1).Value, 1, 2))

End Function
'

Private Sub Command1_Click()
    
    DrawLine selectedLine, False
    
End Sub
'


Private Sub Command2_Click()
    
    picCanvas.Cls

End Sub
'

Private Sub Command3_Click()
    
    DrawLine selectedLine, True
    
End Sub
'

Private Sub Command4_Click()
    
    Unload Me
    
End Sub
'


Private Sub picBack_Click()
    
    Dim Color   As Long
    
    If VBChooseColor(Color) Then
        
        picBack.BackColor = Color
        picCanvas.BackColor = Color
        picBuffer.BackColor = Color
        lblBack.Caption = Right$("000000" + Hex$(Color), 6)
    
    End If
    
End Sub
'

Private Sub picColor_Click()
    
    Dim Color   As Long
    
    If VBChooseColor(Color) Then
        
        picColor.BackColor = Color
        lblColor.Caption = Right$("000000" + Hex$(Color), 6)
    
    End If

End Sub
'
