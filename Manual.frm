VERSION 5.00
Begin VB.Form ManualSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   2910
   ClientTop       =   2070
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5295
   Begin VB.TextBox Edit2 
      Height          =   285
      Left            =   1620
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4680
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   285
      Left            =   3870
      TabIndex        =   4
      Top             =   4230
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   285
      Left            =   2430
      TabIndex        =   3
      Top             =   4230
      Width           =   825
   End
   Begin VB.TextBox Edit1 
      Height          =   285
      Left            =   1620
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4230
      Width           =   645
   End
   Begin VB.VScrollBar Scroll 
      Height          =   4335
      Left            =   4950
      TabIndex        =   1
      Top             =   0
      Width           =   285
   End
   Begin VB.Timer Clock 
      Left            =   0
      Top             =   0
   End
   Begin Manual.Ticker Ticker 
      Height          =   3795
      Left            =   90
      TabIndex        =   0
      Top             =   270
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
   Begin VB.Label Label2 
      Caption         =   "Clock Interval"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   4770
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Change Rate:"
      Height          =   195
      Left            =   450
      TabIndex        =   5
      Top             =   4320
      Width           =   1005
   End
End
Attribute VB_Name = "ManualSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Clock_Timer()

  ' Adjust Caption
  If Ticker.Value = Abs(Scroll.Value) Then
    Caption = "LEVEL"
  Else
    Ticker.Value = Abs(Scroll.Value)
    Caption = "SLOPE"
  End If
  
  ' External clock must set the Value and
  ' call the Update method to display the
  ' the next value
  Ticker.Update
  
End Sub

Private Sub Command1_Click()
  ' Update settings
  Ticker.ChangeRate = Val(Edit1)
  Clock.Interval = Val(Edit2)
  Command1.Font.Bold = False
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Edit1_Change()
  Command1.Font.Bold = True
End Sub

Private Sub Edit2_Change()
  Command1.Font.Bold = True
End Sub

Private Sub Form_Load()
  
  ' All settings were set here to show what was used
  With Ticker
    ' Initial values
    .ChangeRate = 0.01
    .Value = 0
    
    ' Special slope mode settings
    .Automatic = True
    .Enabled = False
    
    ' These slope settings will affect the internal target
    ' value while in special slope mode, so they are set to
    ' provide the least interference
    .SlopeInterval = 200000
    .SlopeRate = 0.000001
    
    ' Over-all appearance
    .Appearance = [3D Sunken]
    .ChartStyle = [One Line]
    .TextStyle = [No Text]
    .LineWidth = 3
    
    ' Adding Grid Lines
    .GridInterval = 100  ' Vertical
    
    .GridLines.Add 50    ' Horizontal
    .GridLines.Add 100
    .GridLines.Add 150
    
    ' Value limits
    .ScaleMax = 202
    .ScaleMin = -2
    .ValueMax = 200
    .ValueMin = 0
  End With
  
  ' External clock
  Clock.Interval = 30
  
  ' Init Scroll bar
  Scroll.Min = -Ticker.ValueMax
  Scroll.Max = Ticker.ValueMin
  Scroll.LargeChange = 20
  Scroll.SmallChange = 2
  Scroll.Value = -100
  
  Edit1.Text = "0.01"
  Edit2.Text = "30"
  Command1.Font.Bold = False
  
End Sub

