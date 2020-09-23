VERSION 5.00
Begin VB.Form Examples 
   BackColor       =   &H80000010&
   Caption         =   "Ticker Examples"
   ClientHeight    =   3510
   ClientLeft      =   3255
   ClientTop       =   3450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5370
   Begin Samples.Ticker Tickers 
      Height          =   915
      Index           =   0
      Left            =   270
      TabIndex        =   7
      Top             =   540
      Width           =   1455
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   915
      Index           =   1
      Left            =   1980
      TabIndex        =   8
      Top             =   540
      Width           =   1455
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   915
      Index           =   2
      Left            =   3690
      TabIndex        =   9
      Top             =   540
      Width           =   1455
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   1275
      Index           =   3
      Left            =   270
      TabIndex        =   10
      Top             =   1980
      Width           =   825
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   1275
      Index           =   4
      Left            =   1620
      TabIndex        =   11
      Top             =   1980
      Width           =   825
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   1275
      Index           =   5
      Left            =   2970
      TabIndex        =   12
      Top             =   1980
      Width           =   825
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Samples.Ticker Tickers 
      Height          =   1275
      Index           =   6
      Left            =   4320
      TabIndex        =   13
      Top             =   1980
      Width           =   825
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Control"
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      Height          =   195
      Index           =   5
      Left            =   2970
      TabIndex        =   5
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Downwards"
      Height          =   195
      Index           =   4
      Left            =   1620
      TabIndex        =   4
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Upwards"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   3
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Charts"
      Height          =   195
      Index           =   2
      Left            =   3690
      TabIndex        =   2
      Top             =   270
      Width           =   915
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Two Color"
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   1
      Top             =   270
      Width           =   915
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   915
   End
End
Attribute VB_Name = "Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Ticker2_OnUpdate(Value As Double)

End Sub

Private Sub Form_Load()

' These properties could be set in the form designer,
' but to show what was changed from the standard control,
' all the changes made are listed here.
  
  ' Standard
  Tickers(0).Automatic = True
  Tickers(0).ChangeInterval = 50
  Tickers(0).Enabled = True
  
  ' Two Color
  Tickers(1).Automatic = True
  Tickers(1).TopColor = RGB(210, 40, 20)
  Tickers(1).BottomColor = RGB(20, 40, 210)
  Tickers(1).LineColor = vbWhite
  Tickers(1).ChangeInterval = 50
  Tickers(1).Alignment = tcALI.[Top & Left Justified]
  Tickers(1).Enabled = True
  
  ' Charts
  Tickers(2).Automatic = True
  Tickers(2).TextStyle = tcTEX.[No Text]
  Tickers(2).ChangeInterval = 50
  Tickers(2).GridLines.Add 25
  Tickers(2).GridLines.Add 50
  Tickers(2).GridLines.Add 75
  Tickers(2).GridLines.Add 99
  Tickers(2).GridColor = vbButtonShadow
  Tickers(2).GridInterval = 15
  Tickers(2).GridZOrder = tcZOR.[Top Color]
  Tickers(2).Enabled = True
  
  ' Upwards
  Tickers(3).Automatic = True
  Tickers(3).Appearance = tcAPP.[3D Raised]
  Tickers(3).TextStyle = tcTEX.[No Text]
  Tickers(3).Direction = tcDIR.North
  Tickers(3).BackColor = vbBlue
  Tickers(3).ChartStyle = tcCHA.[One Line]
  Tickers(3).LineColor = vbWhite
  Tickers(3).LineWidth = 3
  Tickers(3).ChangeInterval = 20
  Tickers(3).ChangeRate = 0.04
  Tickers(3).SlopeRate = 0.6
  Tickers(3).SlopeInterval = 1
  Tickers(3).Enabled = True
  
  ' Downwards
  Tickers(4).Automatic = True
  Tickers(4).Appearance = tcAPP.[3D Sunken]
  Tickers(4).TextStyle = tcTEX.[No Text]
  Tickers(4).Direction = tcDIR.South
  Tickers(4).ChangeInterval = 20
  Tickers(4).ChangeRate = 0.04
  Tickers(4).SlopeInterval = 30
  Tickers(4).SlopeRate = 0.8
  Tickers(4).LineColor = vbWhite
  Tickers(4).BackColor = RGB(210, 40, 20)
  Tickers(4).ChartStyle = tcCHA.[One Line]
  Tickers(4).LineWidth = 3
  Tickers(4).Value = 50
  Tickers(4).Enabled = True
  
  ' Warning - Always set the Appearance after Font changes.
  Tickers(5).Font.Bold = True
  Tickers(5).Font.Size = 12
  Tickers(5).Appearance = tcAPP.[Tool Raised]
  Tickers(5).Automatic = True
  Tickers(5).Alignment = tcALI.[Top & Centered]
  Tickers(5).TextStyle = tcTEX.Percentage
  Tickers(5).Direction = tcDIR.East
  Tickers(5).TopColor = vbWhite
  Tickers(5).LineWidth = 2
  Tickers(5).ChangeInterval = 250
  Tickers(5).ChangeRate = 0.01
  Tickers(5).SlopeInterval = 2
  Tickers(5).SlopeRate = 0.5
  Tickers(5).GridLines.Add 20
  Tickers(5).GridLines.Add 30
  Tickers(5).GridLines.Add 70
  Tickers(5).GridLines.Add 80
  Tickers(5).GridColor = &H808080
  Tickers(5).GridZOrder = tcZOR.[Both Colors]
  Tickers(5).Enabled = True

  ' User controlled
  Tickers(6).Font.Name = "Courier"
  Tickers(6).Font.Size = 8
  Tickers(6).Appearance = tcAPP.[Tool Sunken]
  Tickers(6).Automatic = False
  Tickers(6).Alignment = tcALI.[Bottom & Centered]
  Tickers(6).ChartStyle = tcCHA.[Two Color]
  Tickers(6).TextFormat = tcFOR.Standard
  Tickers(6).ScaleMin = -100
  Tickers(6).ValueMin = -100
  Tickers(6).TextStyle = tcTEX.Actual
  Tickers(6).Direction = tcDIR.East
  Tickers(6).BottomColor = vbButtonFace
  Tickers(6).ChangeInterval = 250
  Tickers(6).SlopeRate = 0
  Tickers(6).Enabled = True


End Sub



Private Sub Tickers_OnUpdate(Index As Integer, Value As Double, ByVal GridSync As Boolean)
Static angle As Single

  Select Case Index
  Case 5 ' Warning - providing visual cues
  
    If Value < 20 Or Value > 80 Then
      Tickers(5).TopColor = vbRed
      Tickers(5).BottomColor = vbRed
      Tickers(5).LineColor = vbBlack
    ElseIf Value < 30 Or Value > 70 Then
      Tickers(5).TopColor = vbYellow
      Tickers(5).BottomColor = vbYellow
      Tickers(5).LineColor = vbBlack
    Else
      Tickers(5).TopColor = &H20C020
      Tickers(5).BottomColor = &H20C020
      Tickers(5).LineColor = vbWhite
    End If
    
  Case 6 ' User control - supplying user values
  
    angle = angle + 0.03
    Value = Sin(angle) * 80
    
    With Tickers(6)
        If Value < -2 Then
          .TopColor = vbWhite
          .BottomColor = vbRed
        ElseIf Value > 2 Then
          .TopColor = vbBlue
          .BottomColor = vbWhite
        Else
          .TopColor = vbBlack
          .BottomColor = vbBlack
        End If
    End With
    
  End Select
  

End Sub
