VERSION 5.00
Begin VB.UserControl ucUpDown 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   0
      Picture         =   "ucUpDown.ctx":0000
      ScaleHeight     =   120
      ScaleWidth      =   810
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   0
      Picture         =   "ucUpDown.ctx":02A9
      ScaleHeight     =   120
      ScaleWidth      =   810
      TabIndex        =   0
      Top             =   0
      Width           =   810
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   615
      Top             =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   855
      TabIndex        =   2
      Top             =   30
      Width           =   450
   End
End
Attribute VB_Name = "ucUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim coter As Integer  'add counter
Dim upp As Boolean   ' count up
Dim dow As Boolean  'count down
Dim zp As Boolean   ' count by 10's or ones

Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_Value = 0
Const m_def_FontColor = vbBlack

Dim m_Max As Long
Dim m_Min As Long
Dim m_Value As Long
Dim m_FontColor As OLE_COLOR
Event Click()
Event MouseDown()
Event MouseUp()

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   upp = True
   Timer1.Enabled = True
   RaiseEvent MouseDown
End Sub

Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If zp = True Then
      zp = False
      GoTo here
   End If
   coter = coter + 1
   Value = coter
   picPos
here:
   Picture1.Print coter
   Picture2.Print coter
   Timer1.Enabled = False
   upp = False
   RaiseEvent MouseUp
End Sub

Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   dow = True
   Timer1.Enabled = True
   RaiseEvent MouseDown
End Sub

Private Sub picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If zp = True Then
       zp = False
      GoTo here
   End If
   coter = coter - 1
   Value = coter
   picPos
here:
   Picture1.Print coter
    Picture2.Print coter
   Timer1.Enabled = False
   dow = False
   RaiseEvent MouseUp
End Sub

Private Sub picPos()
Dim Lcoter As Integer

   If coter > Max Then coter = Max
   If coter < Min Then coter = Min
   
   'get the number of digits
   Label1.Caption = coter
   Lcoter = Len(Label1.Caption)
   If Left$(Label1.Caption, 1) = "-" Then Lcoter = Lcoter - 1
   
   Picture1.Cls
   Picture2.Cls
   Picture2.CurrentY = -100
 
   Select Case Lcoter
      Case 1:
         Picture1.CurrentX = 370
         Picture2.CurrentX = 370
      Case 2:
         Picture1.CurrentX = 270
         Picture2.CurrentX = 270
      Case 3:
         Picture1.CurrentX = 160
         Picture2.CurrentX = 160
   End Select
End Sub

Private Sub Timer1_Timer()
   If upp = True Then coter = coter + 1
   If dow = True Then coter = coter - 1
   zp = True
   Value = coter
   picPos
   Picture1.Print coter
   Picture2.Print coter
   RaiseEvent MouseDown
   RaiseEvent MouseUp
End Sub

Public Property Get FontColor() As OLE_COLOR
   FontColor = m_FontColor
End Property

Public Property Let FontColor(NewFontColor As OLE_COLOR)
   m_FontColor = NewFontColor
   Picture1.ForeColor = m_FontColor
   Picture2.ForeColor = m_FontColor
   picPos
   Picture1.Print coter
   Picture2.Print coter
   PropertyChanged "FontColor"
End Property

Public Property Get Max() As Long
   Max = m_Max
End Property

Public Property Let Max(NewMax As Long)
   m_Max = NewMax
   PropertyChanged "Max"
End Property

Public Property Get Min() As Long
   Min = m_Min
End Property

Public Property Let Min(NewMin As Long)
   m_Min = NewMin
   PropertyChanged "Min"
End Property

Public Property Get Value() As Long
   Value = m_Value
End Property

Public Property Let Value(NewValue As Long)
   m_Value = NewValue
   If m_Value > m_Max Then m_Value = m_Max
   If m_Value < m_Min Then m_Value = m_Min
   coter = m_Value
   picPos
   Picture1.Print coter
   Picture2.Print coter
   PropertyChanged "Value"
End Property

Private Sub UserControl_InitProperties()
   m_Max = m_def_Max
   m_Min = m_def_Min
   m_Value = m_def_Value
End Sub

Private Sub UserControl_Paint()
   Picture1.BackColor = Ambient.BackColor
   Picture2.BackColor = Ambient.BackColor
   Picture1.FontSize = 8
   Picture1.FontBold = True
   Picture2.FontSize = 8
   Picture2.FontBold = True
   picPos
   Picture1.Print coter
   Picture2.Print coter
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      Max = .ReadProperty("Max", m_def_Max)
      Min = .ReadProperty("Min", m_def_Min)
      Value = .ReadProperty("Value", m_def_Value)
      FontColor = .ReadProperty("FontColor", m_def_FontColor)
   End With
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = Picture1.Width
   UserControl.Height = Picture1.Height + Picture2.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Max", m_Max, m_def_Max
      .WriteProperty "Min", m_Min, m_def_Min
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "FontColor", m_FontColor, m_def_FontColor
   End With
End Sub
