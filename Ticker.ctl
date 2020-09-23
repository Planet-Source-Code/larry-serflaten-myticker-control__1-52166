VERSION 5.00
Begin VB.UserControl Ticker 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   Begin VB.Timer Clock 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   450
      Top             =   0
   End
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   90
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "Ticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' My Ticker Control
' © 2003 Larry Serflaten
' NOTICE - The source code for this control has been placed in the
' Public Domain at various websites on the internet.
'
'  I RETAIN ALL COPYRIGHTS TO THIS SOURCE CODE AND ANY DERIVATIVE
'  WORKS FROM IT MADE WITHOUT SIGNIFICANT ENHANCEMENT.
'
'  See the remarks at the end of this module for further details.





' PUBLIC ENUMERATIONS:

' I use very short names in enumerations so that I can get Intellisense
' to kick in after only a few keystrokes.  This makes it easy to always
' qualify the enum values:  tcALI.[Bottom & Centered]

Public Enum tcALI     ' Alignment
  [Bottom & Centered] = 0
  [Bottom & Left Justified] = 1
  [Bottom & Right Justified] = 2
  [Top & Centered] = 3
  [Top & Left Justified] = 4
  [Top & Right Justified] = 5
End Enum

Public Enum tcAPP     ' Appearance
  [Borderless] = 0
  [Flat] = 1
  [3D Raised] = 2
  [3D Sunken] = 3
  [Tool Raised] = 4
  [Tool Sunken] = 5
End Enum

Public Enum tcDIR     ' Direction
  [North] = 0
  [South] = 1
  [West] = 2
  [East] = 3
End Enum

Public Enum tcFOR     ' TextFormat
  [Currency] = 0
  [Fixed] = 1
  [General] = 2
  [Integer] = 3
  [Standard] = 4
End Enum
  
Public Enum tcCHA     ' ChartStyle
  [One Line] = 0
  [Two Color] = 1
  [Two Color & Line] = 2
End Enum

Public Enum tcTEX     ' TextStyle
  [No Text] = 0
  [Actual] = 1
  [Percentage] = 2
End Enum

Public Enum tcZOR     ' GridZOrder
  [All] = 0
  [Top Color] = 1
  [Bottom Color] = 2
  [Both Colors] = 3
End Enum



' Stored grid line values
Public GridLines As New Collection
Attribute GridLines.VB_VarProcData = ";Appearance"
Attribute GridLines.VB_VarDescription = "A collection of values used to draw the continuous grid lines."


'  UDT TYPES:

' To define the Chart and Text areas
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' To save the displayed text / position
Private Type TextValue
  Text As String
  X As Long
  Y As Long
End Type
  

  
' DEFAULT VALUES:

' Apprearance variables
Private Const m_def_Alignment = 2           ' Text position
Private Const m_def_Appearance = 1          ' Control border
Private Const m_def_BottomColor = vbWhite   ' Chart bottom color
Private Const m_def_Direction = 2           ' History travel
Private Const m_def_GridColor = &H80000010  ' Grid color
Private Const m_def_GridInterval = 0        ' Grid vertical line interval
Private Const m_def_GridZOrder = 0          ' Grid line ZOrder
Private Const m_def_LineColor = 0           ' Line color
Private Const m_def_ChartStyle = 2           ' Chart display
Private Const m_def_LineWidth = 1           ' Chart line width
Private Const m_def_TextFormat = 3          ' Text format
Private Const m_def_TextStyle = 1           ' Text display
Private Const m_def_TopColor = &H8000000F   ' Chart top color

' Behavioral variables
Private Const m_def_Automatic = False       ' Clock updates
Private Const m_def_ChangeRate = 0.002      ' Small change
Private Const m_def_SlopeInterval = 5       ' Change counter
Private Const m_def_SlopeRate = 0.2         ' Large change
Private Const m_def_Value = 0               ' Current value
Private Const m_def_ValueMax = 100          ' Value max
Private Const m_def_ValueMin = 0            ' Value min
Private Const m_def_ScaleMax = 100          ' Scale max
Private Const m_def_ScaleMin = 0            ' Scale min



' PROPERTY VARIABLES:

' Apprearance variables
Private m_Alignment     As tcALI
Private m_Appearance    As tcAPP
Private m_BottomColor   As OLE_COLOR
Private m_Direction     As tcDIR
Private m_GridColor     As OLE_COLOR
Private m_GridInterval  As Long
Private m_GridZOrder    As tcZOR
Private m_LineColor     As OLE_COLOR
Private m_ChartStyle     As tcCHA
Private m_LineWidth     As Long
Private m_TextStyle     As tcTEX
Private m_TextFormat    As tcFOR
Private m_TopColor      As OLE_COLOR

' Behavioral variables
Private m_Automatic     As Boolean
Private m_ChangeRate    As Double
Private m_SlopeInterval As Long
Private m_SlopeRate     As Double
Private m_Value         As Double
Private m_ValueMax      As Double
Private m_ValueMin      As Double
Private m_ScaleMax      As Double
Private m_ScaleMin      As Double


' Grid counter
Private m_GridCount As Long

' Image border / positions
Private m_ChartArea As RECT
Private m_TextArea As RECT

' Slew counter
Private m_SlewCount As Long
Private m_Target As Double

' Text variables
Private m_SavedText As TextValue
Private m_Formats As Variant




' API ROUTINES:

Private Declare Function DrawEdge Lib "user32" _
    (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, _
    ByVal grfFlags As Long) As Long

Private Declare Function SetRect Lib "user32" _
    (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long




' EVENT DECLARATONS:

Event OnUpdate(ByRef Value As Double, ByVal GridSync As Boolean)
Attribute OnUpdate.VB_Description = "Occurs when Automatic mode is about to change the ticker display."

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the ticker."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the ticker has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse over the ticker."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the ticker has the focus."
'
'
'
'       E V E N T   H A N D L E R S
'

Private Sub Clock_Timer()

' The clock simply does what the user would do to update the control.
  Update

End Sub



Private Sub UserControl_Resize()
  On Error Resume Next
  ' Remove old text to aviod erasing it at old spot
  m_SavedText.Text = ""
  ' Erase entire text area
  UserControl.Line (m_TextArea.Left - 2, m_TextArea.Top - 2)-(m_TextArea.Right + 2, m_TextArea.Bottom + 2), UserControl.BackColor, BF
  ' Resize chart and text areas
  AdjustAreas
  RefreshImage
End Sub



' User control pass-through events

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
  On Error Resume Next
  RefreshImage
End Sub

Private Sub UserControl_Click()
  On Error Resume Next
  RaiseEvent Click
End Sub



' Control mapped properties

'  WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES
'  AHEAD OF ANY OF THE MAPPED PROPERTIES!!!

'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the text area background color."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  ' If in design mode, make change immediate.
  If Not Ambient.UserMode Then Buffer.BackColor = New_BackColor
  PropertyChanged "BackColor"
  RefreshImage
End Property


'MappingInfo=Clock,Clock,-1,Interval
Public Property Get ChangeInterval() As Long
Attribute ChangeInterval.VB_Description = "Returns/sets the number of milliseconds between updates to the ticker value."
Attribute ChangeInterval.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ChangeInterval = Clock.Interval
End Property

Public Property Let ChangeInterval(ByVal New_ChangeInterval As Long)
  Clock.Interval() = New_ChangeInterval
  PropertyChanged "ChangeInterval"
End Property


'MappingInfo=Clock,Clock,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the ticker updates itself  continuiously."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Enabled = Clock.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  Clock.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property


'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/set the font used for the text value."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
' This Get routine is called when a user tries to set properties of the
' font in code.  [  Font.Bold = True ]   While this routine is executing, the
' changes have not yet taken place.  After they set the font property, the
' text area needs to be recalculated, so the user needs to be aware that
' some other appearance property needs to be set after changing a font
' property in code.  Failure to do so will result in the text not being shown.
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
' The IDE creates a font for use when the user alters the font in the
' Font dialog.  The areas can be recalculated after this change, and will
' work as expected.  For that reason, adjusting the font at design time is
' well advised.
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  ApplyPropertyChange True
End Property


'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display the text value."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
  RefreshImage
End Property


'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over the ticker."
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
  UserControl.MousePointer() = New_MousePointer
  PropertyChanged "MousePointer"
End Property


'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Causes the ticker to refresh its image."
  RefreshImage
End Sub



' More public interface properties


Public Property Get Alignment() As tcALI
Attribute Alignment.VB_Description = "Returns/sets the text position and justification."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As tcALI)
  m_Alignment = New_Alignment
  PropertyChanged "Alignment"
  If m_TextStyle > tcTEX.[No Text] Then
    UserControl.Cls
    ApplyPropertyChange
  End If
End Property



Public Property Get Appearance() As tcAPP
Attribute Appearance.VB_Description = "Returns/sets a value that determines the 3D effects used to display the ticker."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As tcAPP)
  m_Appearance = New_Appearance
  PropertyChanged "Appearance"
  UserControl.Cls
  ApplyPropertyChange
End Property



Public Property Get Automatic() As Boolean
Attribute Automatic.VB_Description = "Returns/sets the use of automatic random value changes."
Attribute Automatic.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Automatic = m_Automatic
End Property

Public Property Let Automatic(ByVal New_Automatic As Boolean)
  m_Automatic = New_Automatic
  PropertyChanged "Automatic"
End Property



Public Property Get BottomColor() As OLE_COLOR
Attribute BottomColor.VB_Description = "Returns/sets the color of the charted area lesser than the value of the ticker line."
Attribute BottomColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  BottomColor = m_BottomColor
End Property

Public Property Let BottomColor(ByVal New_BottomColor As OLE_COLOR)
  m_BottomColor = New_BottomColor
  PropertyChanged "BottomColor"
  RefreshImage
End Property



Public Property Get ChangeRate() As Double
Attribute ChangeRate.VB_Description = "Returns/sets the percent of change used between successive peaks and valleys.  (0 - 1)"
Attribute ChangeRate.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ChangeRate = m_ChangeRate
End Property

Public Property Let ChangeRate(ByVal New_ChangeRate As Double)
  m_ChangeRate = Abs(New_ChangeRate - Fix(New_ChangeRate))
  PropertyChanged "ChangeRate"
End Property



Public Property Get ChartStyle() As tcCHA
Attribute ChartStyle.VB_Description = "Returns/sets the appearance of the charted area of the ticker."
Attribute ChartStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ChartStyle = m_ChartStyle
End Property

Public Property Let ChartStyle(ByVal New_ChartStyle As tcCHA)
  m_ChartStyle = New_ChartStyle
  PropertyChanged "ChartStyle"
  ApplyPropertyChange
End Property



Public Property Get Direction() As tcDIR
Attribute Direction.VB_Description = "Returns/set the orientation and direction of the ticker history."
Attribute Direction.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Direction = m_Direction
End Property

Public Property Let Direction(ByVal New_Direction As tcDIR)
  m_Direction = New_Direction
  PropertyChanged "Direction"
  ApplyPropertyChange
End Property



Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets the color of the grid lines."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  GridColor = m_GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
  m_GridColor = New_GridColor
  PropertyChanged "GridColor"
End Property



Public Property Get GridInterval() As Long
Attribute GridInterval.VB_Description = "Returns/sets the distance between timed grid lines."
Attribute GridInterval.VB_ProcData.VB_Invoke_Property = ";Behavior"
  GridInterval = m_GridInterval
End Property

Public Property Let GridInterval(ByVal New_GridInterval As Long)
  m_GridInterval = Abs(New_GridInterval)
  PropertyChanged "GridInterval"
End Property



Public Property Get GridZOrder() As tcZOR
Attribute GridZOrder.VB_Description = "Returns/sets a value used to determine what parts of the chart the grid will cover."
Attribute GridZOrder.VB_ProcData.VB_Invoke_Property = ";Appearance"
  GridZOrder = m_GridZOrder
End Property

Public Property Let GridZOrder(ByVal New_GridZOrder As tcZOR)
  m_GridZOrder = New_GridZOrder
  PropertyChanged "GridZOrder"
End Property



Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Returns/sets the color of the ticker line."
Attribute LineColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
  m_LineColor = New_LineColor
  PropertyChanged "LineColor"
  RefreshImage
End Property



Public Property Get LineWidth() As Long
Attribute LineWidth.VB_Description = "Returns/sets the width of the ticker line."
Attribute LineWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
  LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal New_LineWidth As Long)
  m_LineWidth = Abs(New_LineWidth)
  PropertyChanged "LineWidth"
End Property



Public Property Get Image() As IPictureDisp
Attribute Image.VB_Description = "Returns the current ticker image."
Attribute Image.VB_ProcData.VB_Invoke_Property = ";Appearance"
  UserControl.AutoRedraw = True
  RefreshImage
  Set Image = UserControl.Image
  UserControl.AutoRedraw = False
End Property



Public Property Get ScaleMax() As Double
Attribute ScaleMax.VB_Description = "Returns/sets the upper boundry of the displayed chart."
Attribute ScaleMax.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ScaleMax = m_ScaleMax
End Property

Public Property Let ScaleMax(ByVal New_ScaleMax As Double)
  If New_ScaleMax <= m_ScaleMin Then Err.Raise 380, "Ticker", "ScaleMax must be greater than ScaleMin"
  m_ScaleMax = New_ScaleMax
  PropertyChanged "ScaleMax"
  ApplyPropertyChange False
End Property



Public Property Get ScaleMin() As Double
Attribute ScaleMin.VB_Description = "Returns/sets the lower boundry value of the displayed chart."
Attribute ScaleMin.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ScaleMin = m_ScaleMin
End Property

Public Property Let ScaleMin(ByVal New_ScaleMin As Double)
  If New_ScaleMin >= m_ScaleMax Then Err.Raise 380, "Ticker", "ScaleMin must be less than ScaleMax"
  m_ScaleMin = New_ScaleMin
  PropertyChanged "ScaleMin"
  ApplyPropertyChange False
End Property



Public Property Get SlopeInterval() As Long
Attribute SlopeInterval.VB_Description = "Returns/sets a value that determines how long a slope will continue before a new target value is selected."
Attribute SlopeInterval.VB_ProcData.VB_Invoke_Property = ";Behavior"
  SlopeInterval = m_SlopeInterval
End Property

Public Property Let SlopeInterval(ByVal New_SlopeInteval As Long)
  m_SlopeInterval = Abs(New_SlopeInteval)
  PropertyChanged "SlopeInterval"
End Property



Public Property Get SlopeRate() As Double
Attribute SlopeRate.VB_Description = "Returns/sets the percent of change used between successive values.  (0 - 1)"
Attribute SlopeRate.VB_ProcData.VB_Invoke_Property = ";Behavior"
  SlopeRate = m_SlopeRate
End Property

Public Property Let SlopeRate(ByVal New_SlopeRate As Double)
  m_SlopeRate = Abs(New_SlopeRate - Fix(New_SlopeRate))
  PropertyChanged "SlopeRate"
End Property



Public Property Get TextFormat() As tcFOR
Attribute TextFormat.VB_Description = "Returns/sets the formatting of the displayed text."
Attribute TextFormat.VB_ProcData.VB_Invoke_Property = ";Appearance"
  TextFormat = m_TextFormat
End Property

Public Property Let TextFormat(ByVal New_TextFormat As tcFOR)
  m_TextFormat = New_TextFormat
  PropertyChanged "TextFormat"
  If m_TextStyle > tcTEX.[No Text] Then
    ApplyPropertyChange
  End If
End Property



Public Property Get TextStyle() As tcTEX
Attribute TextStyle.VB_Description = "Returns/sets a value that determines if the text will be displayed as the actual value, a percentage, or not at all."
Attribute TextStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
  TextStyle = m_TextStyle
End Property

Public Property Let TextStyle(ByVal New_TextStyle As tcTEX)
  m_TextStyle = New_TextStyle
  PropertyChanged "TextStyle"
  UserControl.Cls
  ApplyPropertyChange
End Property



Public Property Get TopColor() As OLE_COLOR
Attribute TopColor.VB_Description = "Returns/sets the color of the charted area greater than the value of the ticker line."
Attribute TopColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  TopColor = m_TopColor
End Property

Public Property Let TopColor(ByVal New_TopColor As OLE_COLOR)
  m_TopColor = New_TopColor
  PropertyChanged "TopColor"
  RefreshImage
End Property



Public Property Get Value() As Double
Attribute Value.VB_Description = "Returns/sets the current value (Default)."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Value.VB_UserMemId = 0
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
  
  ' If out of range, raise an error at design time, but not
  ' at run time.
  If Not Ambient.UserMode Then
    If New_Value < m_ValueMin Or New_Value > m_ValueMax Then
      Err.Raise 380, "Ticker", "Value must be in the range of ValueMin to ValueMax"
    End If
  End If
  
  ' To allow the user to use the Slope rate and interval on values they want to supply,
  ' they need to set Automatic mode but disable the clock.  They must then devise their
  ' own timing method and will then control both the target value and update schedule.
  If (m_Automatic = True) And (Clock.Enabled = False) Then
    ' In the special configuration, changing the value sets the target.
    ' The real value uses the slope mechanism (See IncrementValue)
    m_Target = LimitedValue(New_Value)
  Else
    ' If not the special configuration then the value is updated.
    m_Value = LimitedValue(New_Value)
    PropertyChanged "Value"
  End If

End Property



Public Property Get ValueMax() As Double
Attribute ValueMax.VB_Description = "Returns/sets the maximum allowable value."
Attribute ValueMax.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ValueMax = m_ValueMax
End Property

Public Property Let ValueMax(ByVal New_ValueMax As Double)
  If New_ValueMax <= m_ValueMin Then Err.Raise 380, "Ticker", "ValueMax must be greater than ValueMin"
  m_ValueMax = New_ValueMax
  PropertyChanged "ValueMax"
End Property



Public Property Get ValueMin() As Double
Attribute ValueMin.VB_Description = "Returns/sets the minimum allowable value."
Attribute ValueMin.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ValueMin = m_ValueMin
End Property

Public Property Let ValueMin(ByVal New_ValueMin As Double)
  If New_ValueMin >= m_ValueMax Then Err.Raise 380, "Ticker", "ValueMin must be less than ValueMax"
  m_ValueMin = New_ValueMin
  PropertyChanged "ValueMin"
End Property




'
'     P U B L I C    M E T H O D S
'


Public Sub Clear()
Attribute Clear.VB_Description = "Clears the charted area"

' Removes the history portion of the display

  Set Buffer.Picture = Nothing
  Buffer.BackColor = UserControl.BackColor
  RefreshImage
End Sub



Public Sub Update()  ' Sub Update
Attribute Update.VB_Description = "Causes the chart to increment one moment in time. "
Dim syn As Boolean

' Handles the next value updating process.  When in Automatic mode, the clock calls this
' routine to increment the value and update the display.  When not in Automatic mode,
' the user calls this to update the display.

  On Error Resume Next
  ' Pick new value if needed
  If m_Automatic Then
    IncrementValue
  End If
  ' Raise event when clock generated.
  If Clock.Enabled Then
    If m_GridInterval Then
      syn = (m_GridCount = (m_GridInterval - 1))
    End If
    RaiseEvent OnUpdate(m_Value, syn)
    m_Value = LimitedValue(m_Value)
  End If
  
  ' Render new image
  DrawChart
  RefreshImage

End Sub




'
'     D I S P L A Y    R O U T I N E S
'

Private Sub ApplyPropertyChange(Optional Area As Boolean = True)

' Refeshes control at appearance property changes.  Some
' properties can be changed 'on the fly' (Without causing a new
' history image). Those that can be changed on the fly include
' the Value, the Scale Min/Max, the Value Min/Max, the Slope and
' Change properties, plus the image colors.

  If Area Then AdjustAreas
  DrawChart
  RefreshImage
End Sub


Private Sub AddColors(ByVal V1 As Long, ByVal V2 As Long, _
                      ByVal L1 As Long, ByVal L2 As Long, _
                      ByVal L3 As Long, ByVal L4 As Long, _
                      Optional ForeGround As Boolean = False)

'  ZOrder key                BackGround          ForeGround
'  [All] = 0                 Bot Top Pnt
'  [Top Color] = 1               Top              Bot Pnt
'  [Bottom Color] = 2            Bot              Top Pnt
'  [Both Colors] = 3           Bot Top              Pnt

  ' Top and bottom colors
  If m_ChartStyle > tcCHA.[One Line] Then
    Buffer.DrawWidth = 1
    If ForeGround Then
      If m_GridZOrder = tcZOR.[Bottom Color] Then Buffer.Line (V1, V2)-(L1, L2), m_TopColor
      If m_GridZOrder = tcZOR.[Top Color] Then Buffer.Line (V1, V2)-(L3, L4), m_BottomColor
    Else
      If m_GridZOrder <> tcZOR.[Bottom Color] Then Buffer.Line (V1, V2)-(L1, L2), m_TopColor
      If m_GridZOrder <> tcZOR.[Top Color] Then Buffer.Line (V1, V2)-(L3, L4), m_BottomColor
    End If
  End If
  
  ' Add line
  If m_ChartStyle <> tcCHA.[Two Color] Then
    Buffer.DrawWidth = m_LineWidth
    If ForeGround Then
      If m_GridZOrder <> tcZOR.All Then Buffer.PSet (V1, V2), m_LineColor
    Else
      If m_GridZOrder = tcZOR.All Then Buffer.PSet (V1, V2), m_LineColor
    End If
  End If
End Sub


Private Sub AddGrid(ByVal V1 As Single, ByVal V2 As Single, _
                    ByVal L1 As Long, ByVal L2 As Long, _
                    ByVal L3 As Long, ByVal L4 As Long, _
                    Optional Height As Long = 0)
Dim grd As Variant
  
  Buffer.DrawWidth = 1
  ' Add horz. grid
  If Height Then
    For Each grd In GridLines
      Buffer.PSet (V1, Height - (LimitedScale(grd) * V2)), m_GridColor
    Next
  Else
    For Each grd In GridLines
      Buffer.PSet (LimitedScale(grd) * V1, V2), m_GridColor
    Next
  End If
  
  ' Add vertical grid
  If m_GridInterval > 0 And m_GridCount = 0 Then
    Buffer.Line (L1, L2)-(L3, L4), m_GridColor
  End If
  
End Sub


Private Sub DrawChart()
Dim wid As Long, hgt As Long
Dim vsc As Long, fsc As Long ' Value SCaled, Full SCale
Dim vpt As Single            ' The point where Value will be plotted

' Draws the chart to the chart Buffer.  The Buffer maintains the history image of
' past values.  It is a persistant image of the chart area of the control.


  ' Increment Grid interval counter
  If m_GridInterval > 0 Then
    m_GridCount = (m_GridCount + 1) Mod m_GridInterval
  End If
  
  ' Cache values
  vsc = m_Value - m_ScaleMin
  fsc = m_ScaleMax - m_ScaleMin
  
  With Buffer
      ' Cache values
      wid = .ScaleWidth - 1
      hgt = .ScaleHeight - 1
      .DrawWidth = 1
      Select Case m_Direction
      Case tcDIR.North
          ' Calc Value point
          vpt = (vsc * wid) / fsc
          ' Increment history
          BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 1, vbSrcCopy
          ' Erase last line
          Buffer.Line (0, hgt)-Step(.ScaleWidth, 0), UserControl.BackColor
          ' Add background, grid, ForeGround
          AddColors vpt, hgt, .ScaleWidth, hgt, -1, hgt, False
          AddGrid wid / fsc, hgt, -1, hgt, .ScaleHeight, hgt
          AddColors vpt, hgt, .ScaleWidth, hgt, -1, hgt, True
      
      Case tcDIR.South
          vpt = (vsc * wid) / fsc
          BitBlt .hdc, 0, 1, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, vbSrcCopy
          Buffer.Line (0, 0)-Step(.ScaleWidth, 0), UserControl.BackColor
          AddColors vpt, 0, .ScaleWidth, 0, -1, 0, False
          AddGrid wid / fsc, 0, -1, 0, .ScaleHeight, 0
          AddColors vpt, 0, .ScaleWidth, 0, -1, 0, True
           
      Case tcDIR.West
          vpt = hgt - ((vsc * hgt) / fsc)
          'Set Buffer.Picture = Nothing
          BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 1, 0, vbSrcCopy
          Buffer.Line (wid, 0)-Step(0, .ScaleHeight), UserControl.BackColor
          AddColors wid, vpt, wid, -1, wid, .ScaleHeight, False
          AddGrid wid, hgt / fsc, wid, -1, wid, .ScaleHeight, hgt
          AddColors wid, vpt, wid, -1, wid, .ScaleHeight, True
          
      Case tcDIR.East
          vpt = hgt - ((vsc * hgt) / fsc)
          BitBlt .hdc, 1, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, vbSrcCopy
          Buffer.Line (0, 0)-Step(0, .ScaleHeight), UserControl.BackColor
          AddColors 0, vpt, 0, -1, 0, .ScaleHeight, False
          AddGrid 0, hgt / fsc, 0, -1, 0, .ScaleHeight, hgt
          AddColors 0, vpt, 0, -1, 0, .ScaleHeight, True
           
      End Select
  End With
End Sub



Private Sub ShowBorder()
Dim rct As RECT

' Draws the selected border style when updating


  ' Chart / Text separator line
  If m_TextStyle > tcTEX.[No Text] Then
    If m_Alignment > tcALI.[Bottom & Right Justified] Then
      SetRect rct, 0, m_TextArea.Bottom - 1, UserControl.ScaleWidth + 1, m_TextArea.Bottom + 1
    Else
      SetRect rct, 0, m_TextArea.Top, UserControl.ScaleWidth + 1, m_TextArea.Top + 2
    End If
    DrawEdge UserControl.hdc, rct, 2, 15
  End If


  ' Determine outer border position
  If m_Appearance > tcAPP.[Flat] Then
    SetRect rct, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Else
    SetRect rct, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
  End If
  
  ' Draw border
  Select Case m_Appearance
  Case tcAPP.[Borderless]
  Case tcAPP.[Flat]
    With rct
      UserControl.FillStyle = vbFSTransparent
      UserControl.Line (.Left, .Top)-(.Right, .Bottom), vbBlack, B
    End With
  Case tcAPP.[3D Raised]
    DrawEdge UserControl.hdc, rct, 5, 15
  Case tcAPP.[3D Sunken]
    DrawEdge UserControl.hdc, rct, 10, 15
  Case tcAPP.[Tool Raised]
    DrawEdge UserControl.hdc, rct, 4, 15
  Case tcAPP.[Tool Sunken]
    DrawEdge UserControl.hdc, rct, 2, 15
  End Select

End Sub


Private Sub ShowChart()

' Copies the chart to the control.  Using a Buffer to hold and manipulate the chart allows
' for it to be updated out of view and only copied over to the control, when needed.

  With Buffer
    BitBlt UserControl.hdc, m_ChartArea.Left, m_ChartArea.Top, .Width, .Height, .hdc, 0, 0, vbSrcCopy
  End With
End Sub


Private Sub ShowText()
Dim X As Long, Y As Long, wid As Long, klr As Long
Dim tex As String

' The text portion of the control shows the current value.  It can be displayed above
' or below the chart image, with Left, Center, or Right justification.  To reduce
' flicker when erasing the old value, the old text value is printed right on top of
' itself, using the background color.  The new text is printed in its place and is
' saved for the next time it needs to be erased.

  ' Build text
  Select Case m_TextStyle
  Case tcTEX.Actual
    tex = Format$(m_Value, m_Formats(m_TextFormat))
  Case tcTEX.Percentage
    tex = CStr((m_Value - m_ScaleMin) * 100 \ (m_ScaleMax - m_ScaleMin)) & "%"
  End Select
  
  
  ' Initialize printing variables
  With UserControl
    wid = .TextWidth(tex)    ' New text size
    klr = .ForeColor         ' Save Text color
    .ForeColor = .BackColor  ' Set Erase color
    ' Vertical centering
    Y = ((m_TextArea.Bottom - m_TextArea.Top) - .TextHeight(tex)) \ 2 + m_TextArea.Top
    If m_Alignment < tcALI.[Top & Centered] Then Y = Y + 2
  End With
  
  ' Keep text from overwrting chart area when sized small
  If Y < (m_TextArea.Top - 2) Then tex = ""
  
  
  ' Calc X position - Horizontal justification
  Select Case m_Alignment
  Case tcALI.[Bottom & Left Justified], tcALI.[Top & Left Justified]
    X = m_TextArea.Left + 2
  Case tcALI.[Bottom & Centered], tcALI.[Top & Centered]
    X = (m_TextArea.Right - wid) \ 2 + 1
  Case Else
    X = (m_TextArea.Right - wid) - 2
  End Select
  
  
  With m_SavedText
      ' Erase old
      UserControl.CurrentX = .X
      UserControl.CurrentY = .Y
      UserControl.Print .Text;
      
      ' Draw new
      UserControl.CurrentX = X
      UserControl.CurrentY = Y
      UserControl.ForeColor = klr
      UserControl.Print tex;
      
      ' Save values - allows quick erase with low flicker
      .Text = tex
      .X = X
      .Y = Y
  End With

End Sub





'
'     H E L P E R    R O U T I N E S
'

Private Sub AdjustAreas()
Dim hgt As Long, ofs As Long

' Determines Chart and Text positions when control is resized.
  
  ' Sets the chart area rectangle (Adjusted for border appearance)
  With UserControl
      Select Case m_Appearance
      Case tcAPP.[Borderless]
        SetRect m_ChartArea, 0, 0, .ScaleWidth - 1, .ScaleHeight - 1
      Case tcAPP.[Flat]
        SetRect m_ChartArea, 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
      Case tcAPP.[3D Raised], tcAPP.[3D Sunken]
        SetRect m_ChartArea, 2, 2, .ScaleWidth - 3, .ScaleHeight - 3
      Case tcAPP.[Tool Raised], tcAPP.[Tool Sunken]
        SetRect m_ChartArea, 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
      End Select
  End With
  
  ' Sets the text area rectangle (Re-adjusts chart area)
  If m_TextStyle > tcTEX.[No Text] Then
    hgt = UserControl.TextHeight("X") + 2
    
    With m_ChartArea
      If m_Alignment < tcALI.[Top & Centered] Then
        
        ' Bottom
        If (.Top + hgt) < .Bottom Then
          ' If there is room for both then adjust for both
          ofs = .Bottom
          .Bottom = .Bottom - hgt
        Else
          ' Otherwise Chart takes up the whole area
          ofs = .Bottom + hgt
        End If
        
        SetRect m_TextArea, .Left, .Bottom + 1, .Right, ofs
      
      Else
        ' Top
        If (.Top + hgt) < .Bottom Then
          ofs = .Top
          .Top = .Top + hgt + 1
        Else
          ofs = .Top - hgt
        End If
        
        SetRect m_TextArea, .Left, ofs, .Right, .Top - 1
      
      End If
      
    End With
  
  End If
  
  'Set Buffer to new size
  With m_ChartArea
      Buffer.Move 0, 0, .Right - .Left + 1, .Bottom - .Top + 1
  End With

  ' New size means new history image
  Set Buffer.Picture = Nothing
  UserControl.Cls
  Buffer.BackColor = UserControl.BackColor

End Sub


Private Function LimitedScale(ByVal Number As Double) As Double

' Restricts Number to within Scale range

  If Number > m_ScaleMax Then Number = m_ScaleMax
  If Number < m_ScaleMin Then Number = m_ScaleMin
  LimitedScale = Number

End Function


Private Function LimitedValue(ByVal Number As Double) As Double

' Restricts Number to within Value range

  If Number > m_ValueMax Then Number = m_ValueMax
  If Number < m_ValueMin Then Number = m_ValueMin
  LimitedValue = Number

End Function


Private Sub IncrementValue()
Dim inc As Single, rand As Single, range As Single

' Handles Value changes while in Automatic mode.  At every Change interval a target
' value is picked, and the value creeps up to the target based on the slew rate.
  
  ' Increment slew counter
  m_SlewCount = (m_SlewCount + 1) Mod m_SlopeInterval
  If m_SlewCount = 0 Then
    'Pick new target from Value +/- SlopeRate
    range = Abs((m_ValueMax - m_ValueMin) * m_SlopeRate)
    rand = Rnd * (range + range) + 0.00000001
    m_Target = LimitedValue(m_Target + (rand - range))
  End If
  
  ' Increment Value +/- ChangeRate
  inc = (m_ValueMax - m_ValueMin) * m_ChangeRate
  ' Avoid overshooting target
  If inc > Abs(m_Target - m_Value) Then
    m_Value = m_Target
  Else
    m_Value = m_Value + inc * Sgn(m_Target - m_Value)
  End If

End Sub


Private Sub InitLocal()

  ' Initialization that needs to be done whether for a new control, or for
  ' showing a previously unloaded control.
  
  Randomize
  m_Formats = Array("Currency", "Fixed", "General Number", "0", "Standard")
End Sub


Private Sub RefreshImage()

' Makes the current images visible on the control.
  
  ShowChart
  If m_TextStyle > tcTEX.[No Text] Then ShowText
  ShowBorder

End Sub



Private Sub UserControl_InitProperties()

' Initializes variables to default values for new control

  InitLocal
  m_Alignment = m_def_Alignment
  m_Appearance = m_def_Appearance
  m_Automatic = m_def_Automatic
  m_BottomColor = m_def_BottomColor
  m_ChangeRate = m_def_ChangeRate
  m_Direction = m_def_Direction
  m_GridColor = m_def_GridColor
  m_GridInterval = m_def_GridInterval
  m_GridZOrder = m_def_GridZOrder
  m_LineColor = m_def_LineColor
  m_ChartStyle = m_def_ChartStyle
  m_LineWidth = m_def_LineWidth
  m_ScaleMax = m_def_ScaleMax
  m_ScaleMin = m_def_ScaleMin
  m_SlopeRate = m_def_SlopeRate
  m_SlopeInterval = m_def_SlopeInterval
  m_TextFormat = m_def_TextFormat
  m_TextStyle = m_def_TextStyle
  m_TopColor = m_def_TopColor
  m_Value = m_def_Value
  m_ValueMax = m_def_ValueMax
  m_ValueMin = m_def_ValueMin
  m_Target = m_def_Value
  Set Font = Ambient.Font
  Buffer.DrawWidth = m_LineWidth
  Buffer.BackColor = UserControl.BackColor
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  ' Loads property values from storage for re-loaded control
  
  InitLocal
  Clock.Enabled = PropBag.ReadProperty("Enabled", False)
  Clock.Interval = PropBag.ReadProperty("ChangeInterval", 20)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  
  m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
  m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
  m_Direction = PropBag.ReadProperty("Direction", m_def_Direction)
  m_GridColor = PropBag.ReadProperty("GridColor", m_def_GridColor)
  m_GridInterval = PropBag.ReadProperty("GridInterval", m_def_GridInterval)
  m_GridZOrder = PropBag.ReadProperty("GridZOrder", m_def_GridZOrder)
  m_ChartStyle = PropBag.ReadProperty("ChartStyle", m_def_ChartStyle)
  m_TopColor = PropBag.ReadProperty("TopColor", m_def_TopColor)
  m_LineColor = PropBag.ReadProperty("LineColor", m_def_LineColor)
  m_LineWidth = PropBag.ReadProperty("LineWidth", m_def_LineWidth)
  m_BottomColor = PropBag.ReadProperty("BottomColor", m_def_BottomColor)
  m_TextStyle = PropBag.ReadProperty("TextStyle", m_def_TextStyle)
  m_TextFormat = PropBag.ReadProperty("TextFormat", m_def_TextFormat)
  m_Automatic = PropBag.ReadProperty("Automatic", m_def_Automatic)
  m_ChangeRate = PropBag.ReadProperty("ChangeRate", m_def_ChangeRate)
  m_SlopeRate = PropBag.ReadProperty("SlopeRate", m_def_SlopeRate)
  m_SlopeInterval = PropBag.ReadProperty("SlopeInterval", m_def_SlopeInterval)
  m_ValueMin = PropBag.ReadProperty("ValueMin", m_def_ValueMin)
  m_ValueMax = PropBag.ReadProperty("ValueMax", m_def_ValueMax)
  m_ScaleMin = PropBag.ReadProperty("ScaleMin", m_def_ScaleMin)
  m_ScaleMax = PropBag.ReadProperty("ScaleMax", m_def_ScaleMax)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  Buffer.DrawWidth = m_LineWidth
  Buffer.BackColor = UserControl.BackColor
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  ' Writes property values to storage when un-loading.
  
  Call PropBag.WriteProperty("LineWidth", Buffer.DrawWidth, 1)
  Call PropBag.WriteProperty("Enabled", Clock.Enabled, False)
  Call PropBag.WriteProperty("ChangeInterval", Clock.Interval, 20)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
  
  Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
  Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
  Call PropBag.WriteProperty("Direction", m_Direction, m_def_Direction)
  Call PropBag.WriteProperty("GridColor", m_GridColor, m_def_GridColor)
  Call PropBag.WriteProperty("GridInterval", m_GridInterval, m_def_GridInterval)
  Call PropBag.WriteProperty("GridZOrder", m_GridZOrder, m_def_GridZOrder)
  Call PropBag.WriteProperty("ChartStyle", m_ChartStyle, m_def_ChartStyle)
  Call PropBag.WriteProperty("TopColor", m_TopColor, m_def_TopColor)
  Call PropBag.WriteProperty("LineColor", m_LineColor, m_def_LineColor)
  Call PropBag.WriteProperty("LineWidth", m_LineWidth, m_def_LineWidth)
  Call PropBag.WriteProperty("BottomColor", m_BottomColor, m_def_BottomColor)
  Call PropBag.WriteProperty("TextStyle", m_TextStyle, m_def_TextStyle)
  Call PropBag.WriteProperty("TextFormat", m_TextFormat, m_def_TextFormat)
  Call PropBag.WriteProperty("Automatic", m_Automatic, m_def_Automatic)
  Call PropBag.WriteProperty("ChangeRate", m_ChangeRate, m_def_ChangeRate)
  Call PropBag.WriteProperty("SlopeRate", m_SlopeRate, m_def_SlopeRate)
  Call PropBag.WriteProperty("SlopeInterval", m_SlopeInterval, m_def_SlopeInterval)
  Call PropBag.WriteProperty("ValueMin", m_ValueMin, m_def_ValueMin)
  Call PropBag.WriteProperty("ValueMax", m_ValueMax, m_def_ValueMax)
  Call PropBag.WriteProperty("ScaleMin", m_ScaleMin, m_def_ScaleMin)
  Call PropBag.WriteProperty("ScaleMax", m_ScaleMax, m_def_ScaleMax)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub




' LEGAL INFO:

'  I RETAIN ALL COPYRIGHTS TO THIS SOURCE CODE AND ANY DERIVATIVE
'  WORKS FROM IT MADE WITHOUT SIGNIFICANT ENHANCEMENT.
'
' I do this to allow you to use the above source code for any use
' you see fit, but I do not allow you to claim it as your own work,
' or attempt to sell it, or get other renumerations from it without
' significant enhancement to its function and use.  In other words,
' you can't sell this control as your own work, but you can use it
' in your own application or product (that includes significantly
' more functionality than this simple control), and sell that.

' You do not need to display my name or copyright information to
' users of your applicatons, but please keep the control souce code
' entact, as supplied, including these remarks.  Of course giving
' credit where credit is due, is always a good idea!  :-)



' DESCRIPTION:

' The Ticker is a UserControl that displays changing values at periodic
' intervals while showing those values over time.  The values can be set
' to change and update automatically, or they can be managed under user
' control.  The displayed history of values can be ran in any one of four
' directions, with the text representation of the current value displayed
' above or below the history image.
'
' There are several properties supplied to adjust the appearance of the
' control, the colors used, and the text font.  Other properties control
' how often the value changes, and how much it is allowed to change.  It
' can be adjusted to produce history images that range from very smooth
' to extremely jagged.
'
' The upper and lower (Max/Min) limits for the Value are separate from
' the upper and lower limits of the displayed scale.  This allows for
' wide amplitude variations in the display even for small variations in
' the Value.  Typically, you'll want the range of the display slightly
' larger than the range of the Value.  This keeps the displayed line
' from going out of the displayed image area, (Which is allowed in case
' that is desired).



' USAGE:

' For both Manual and Automatic modes, the Scale and Value ranges (min/max)
' need to be set.  For Automatic mode, the Change Interval and Rate as well
' as the Slope Interval and Rate need to be set.



' The Scale range values determine the upper and lower limits of the charted
' area. These help determine the amplitude of the displayed value.

' The Value range values determine the upper and lower limits of the current
' value.  The current value will not be allowed to extend beyond these limits.



' The ChangeInterval determines how fast the update process is called when
' in automatic mode.  It governs the speed of the history list.  It is mapped
' to the Clock Interval and is limited to the Integer range.

' The ChangeRate determines the amount of change allowed for each successive
' target value.  The target value is not the next value, it is a target that
' the current value will head toward.  How long it takes for the current value
' to reach the target value is based on the slope rate and interval.
' The ChangeRate value is a percentage of the entire value range and is limited
' to 0 - 1.  A value of .1 means the next target value can be incremented or
' decremented by 10% (of the value range) from the old target value.

' The SlopeInterval determines how often a new target is selected.  Each
' update is one unit, and when the number of SlopeInterval units have expired,
' a new target value is selected.  It is limited to the range of a Long data type.

' The SlopeRate determines how much the current value is allowed to change
' from one update to the next.  It is a percentage of the entire value
' range and is limitied to 0 - 1.  A value of .1 means the next value can
' increment or decrement by 10% (of the value range) toward the target value.



' In manual mode, the user sets the value he wants displayed, and calls the
' Update method to display the new value.  The history of value is maintained.
' When the user sets the Value, it becomes the current value, and will be used
' at the next update.  If use of the slope values is desired, then set the
' control into automatic mode, but disable the clock (Set Enabled to False).
' In this configuration, set the Value ahead of every call to the Update method
' to ensure the "Automatic" process does not change the target value from your
' desired value.  (An exceedingly large SlopeInterval would also work, but
' that value may ultimately be reached.  Setting the value ahead of each call
' to Update ensures your desired target value is used.)

' In automatic mode, the Change and Slope settings determine the next value
' and how often the display is updated.  The Enabled property is mapped to
' the Clock's Enabled property and determines if the display will be updated
' or not.  When the Clock is enabled, the OnUpdate event fires to allow the user
' to control the value, if needed.

' The grid is governed by the GridLines collection, and the GridInterval property.
' The two types of grid lines are; continuous, following the flow of the charted area,
' and timed, crossways to the flow of the charted area.

' If the GridLines collection has values added, they will be used to display
' the continuous lines.  If the GridInterval property is set, then the timed
' grid lines will be displayed.  The OnUpdate GridSync parameter will be True
' when the current update will be showing a timed grid line.  The time interval
' between grid lines is dependant on the GridInterval and ChangeInterval settings
' and may not be accurate at time keeping.  (it uses the standard VB Timer)

' The other properties control the control's appearance and are self-evident
' from their use and a little experimentation.  Enjoy!


