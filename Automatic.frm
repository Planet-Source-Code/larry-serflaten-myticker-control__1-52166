VERSION 5.00
Begin VB.Form AutoSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   2910
   ClientTop       =   2505
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5010
   Begin Automatic.Ticker Ticker 
      Height          =   915
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   4785
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
   Begin VB.Frame Frames 
      Height          =   1500
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   1080
      Width           =   2265
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shares"
         Height          =   195
         Index           =   2
         Left            =   810
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   990
         Width           =   2085
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   2085
      End
   End
   Begin VB.Frame Frames 
      Height          =   1500
      Index           =   0
      Left            =   2610
      TabIndex        =   0
      Top             =   1080
      Width           =   2265
      Begin VB.TextBox Edits 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   1005
      End
      Begin VB.TextBox Edits 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton Buttons 
         Caption         =   "Sell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Buttons 
         Caption         =   "Buy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1170
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   10
         Top             =   720
         Width           =   2085
      End
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   645
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   180
      Width           =   4335
   End
End
Attribute VB_Name = "AutoSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Market As New Collection
Private Money As Currency
Private Shares As Long
Private Finish As Date


Private Sub Buttons_Click(Index As Integer)
  Select Case Index
  Case 0 ' Buy
    Market.Add Abs(Val(Edits(Index)))
    Buttons(0).Caption = CStr(Market.Count) & " Buy"
  Case 1 ' Sell
    Market.Add -Abs(Val(Edits(Index)))
  End Select
End Sub

Private Sub Edits_GotFocus(Index As Integer)
' Highlights text on focus
  With Edits(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub Edits_KeyPress(Index As Integer, KeyAscii As Integer)
' Allows Enter to accept bid
  If KeyAscii = vbKeyReturn Then
    Buttons(Index).Value = True
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  ' All settings were set here to show what was used
  With Ticker
    ' Appearance
    .Appearance = [Tool Raised]
    .ChartStyle = [Two Color]
    .TextStyle = [No Text]
    .LineWidth = 2
    .GridInterval = 90
    .GridLines.Add 20
        
    ' Automatic mode
    .Automatic = True
    .Enabled = True
    
    ' Slope settings
    .ChangeInterval = 20
    .ChangeRate = 0.02
    .SlopeInterval = 10
    .SlopeRate = 0.2
    
    ' Value limits
    .ScaleMax = 210
    .ScaleMin = -20
    .ValueMax = 200
    .ValueMin = -10
  End With
  
  Edits(0).Text = "30"
  Edits(1).Text = "90"
  
  Finish = DateAdd("s", 182, Now)
  Money = 500
  
End Sub

Private Sub Ticker_OnUpdate(Value As Double, ByVal GridSync As Boolean)
  ' Transactions are processed one per vertical grid line.
  If GridSync Then
    If Market.Count Then
      ProcessTransaction Value
    End If
  End If
  InterfaceUpdate
End Sub


Private Sub InterfaceUpdate(Optional Msg As String = "")
Dim mny As String
Dim shr As String
Static cnt As Long, tim As String
' Adjust lables and captions

  mny = Format$(Money, "Currency")
  shr = Format$(Shares, "###,###,###,##0")
  
  ' If a Buy/Sell msg is passed in then use that for
  ' the timer display
  If Len(Msg) Then
    Labels(5).Font.Bold = True
    cnt = 60
    tim = Msg
  End If
  
  If cnt Then
    cnt = cnt - 1
  Else
    Labels(5).Font.Bold = False
    tim = Format$(Finish - Now, "nn:ss")
  End If
  
  ' Update displays
  If Labels(0).Caption <> shr Then Labels(0) = shr
  If Labels(1).Caption <> mny Then Labels(1) = mny
  If Labels(5).Caption <> tim Then Labels(5) = tim
  Caption = Format$(Ticker.Value, "Standard")

  ' Can't buy when price < 0 but you can SELL for a loss!!
  Buttons(0).Enabled = (Ticker.Value > 0.999)
  Edits(0).Enabled = (Ticker.Value > 0.999)

  ' End of game test
  If (Now >= Finish) Or ((Money < 0) And Shares < 1) Then EndGame
  
End Sub


Private Sub ProcessTransaction(ByVal Price As Double)
Dim shr As Long
Dim pur As Double

  If Market.Count > 0 Then
    
    shr = Market(1)
    
    If (shr > 0) And (Price >= 1) Then  'Buy
      Market.Remove 1
      pur = shr * Price
      ' Limit to actual Money value
      If pur > Money Then
        shr = Int(Money / Price)
        pur = shr * Price
      End If
      ' Adjust user stats
      Money = Money - pur
      Shares = Shares + shr
      ' Show transaction
      If shr Then
        Beep
        InterfaceUpdate "BUY:   " & shr & "  @  " & Format$(Price, "Currency")
      End If
      
    ElseIf shr < 0 Then                 ' Sell
      Market.Remove 1
      shr = Abs(shr)
      If shr > Shares Then shr = Shares
      pur = shr * Price
      Money = Money + pur
      Shares = Shares - shr
      If shr Then
        Beep
        InterfaceUpdate "SELL:  " & shr & "  @  " & Format$(Price, "Currency")
      End If
    End If
  End If
  Buttons(0).Caption = CStr(Market.Count) & " Buy"
End Sub



Private Sub EndGame()

  Ticker.Enabled = False
  Ticker.Visible = False
  ' Sell off user's shares
  Set Market = New Collection
  Market.Add -Shares
  ProcessTransaction Ticker.Value

End Sub
