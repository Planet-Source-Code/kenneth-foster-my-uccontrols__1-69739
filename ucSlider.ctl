VERSION 5.00
Begin VB.UserControl ucSlider 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   1410
   ScaleWidth      =   2805
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   225
      Width           =   2250
      Begin VB.Image imgSlider 
         Height          =   210
         Left            =   0
         Picture         =   "ucSlider.ctx":0000
         Stretch         =   -1  'True
         Top             =   390
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   45
         Index           =   0
         Left            =   -30
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   1
         Left            =   60
         Top             =   345
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   150
         Top             =   330
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   90
         Index           =   3
         Left            =   240
         Top             =   315
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   4
         Left            =   330
         Top             =   300
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   5
         Left            =   420
         Top             =   285
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   510
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   7
         Left            =   600
         Top             =   255
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   8
         Left            =   690
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   9
         Left            =   780
         Top             =   225
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   10
         Left            =   870
         Top             =   210
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   210
         Index           =   11
         Left            =   960
         Top             =   195
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   12
         Left            =   1050
         Top             =   180
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   13
         Left            =   1140
         Top             =   165
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   14
         Left            =   1230
         Top             =   150
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   270
         Index           =   15
         Left            =   1320
         Top             =   135
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   16
         Left            =   1410
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   17
         Left            =   1500
         Top             =   105
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   18
         Left            =   1590
         Top             =   90
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   19
         Left            =   1680
         Top             =   75
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   20
         Left            =   1770
         Top             =   60
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   360
         Index           =   21
         Left            =   1860
         Top             =   45
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   22
         Left            =   1950
         Top             =   30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   390
         Index           =   23
         Left            =   2040
         Top             =   15
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   420
         Index           =   24
         Left            =   2130
         Top             =   -15
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Line Line1 
         X1              =   -15
         X2              =   2235
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   60
         Y1              =   450
         Y2              =   600
      End
      Begin VB.Line Line3 
         X1              =   585
         X2              =   585
         Y1              =   465
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   1110
         X2              =   1110
         Y1              =   465
         Y2              =   585
      End
      Begin VB.Line Line4 
         X1              =   1650
         X2              =   1650
         Y1              =   465
         Y2              =   615
      End
      Begin VB.Line Line6 
         X1              =   2175
         X2              =   2175
         Y1              =   450
         Y2              =   600
      End
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   30
      TabIndex        =   2
      Top             =   15
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1980
      Picture         =   "ucSlider.ctx":010A
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1980
      Picture         =   "ucSlider.ctx":0254
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Volume Control"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' some code borrowed and modified from Richard Allsebrooks submission EBSlider
' the rest is mine
Public Enum Bar
 Rainbow = 0
 Red = 1
 Blue = 2
 Green = 3
 Yellow = 4
 Purple = 5
 Turq = 6
 Gray = 7
End Enum

Const m_def_DropDownCtrl = True
Const m_def_Caption = ""
Const m_def_DisableDropDown = False
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
Const m_def_TickColor = vbBlack
Const m_def_BackColor = &HC0C0C0
Const m_def_CapBkgdColor = vbWhite
Const m_def_CapFontColor = vbBlack
Const m_def_BarColor = 0
Const m_def_ValueHide = True

Dim m_ValueHide As Boolean
Dim m_DropDownCtrl As Boolean
Dim m_Caption As String
Dim m_DisableDropDown As Boolean
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
Dim m_TickColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_CapBkgdColor As OLE_COLOR
Dim m_CapFontColor As OLE_COLOR
Dim m_BarColor As Integer

Private SliderWidth      As Long

Private Sub UserControl_Initialize()
   m_Min = m_def_Min
   m_Max = m_def_Max
   m_Value = m_def_Value
   m_TickColor = m_def_TickColor
   m_BackColor = m_def_BackColor
   m_CapBkgdColor = m_def_CapBkgdColor
   m_CapFontColor = m_def_CapFontColor
   m_DisableDropDown = m_def_DisableDropDown
   m_BarColor = m_def_BarColor
   m_ValueHide = m_def_ValueHide
   SliderWidth = 130
End Sub

Private Sub UserControl_InitProperties()
   Caption = Extender.Name
   Min = 0
   Max = 100
   Value = 0
   DropDownCtrl = True
   DisableDropDown = False
   SliderWidth = 130
   SliderPos
End Sub

Private Sub UserControl_Resize()
   If DropDownCtrl = False Then
      UserControl.Height = 240
   Else
      UserControl.Height = 855
   End If
   UserControl.Width = 2250
End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim sldPos  As Long
   Dim sldScale  As Single
   Dim shp As Integer
   
   If Button = 1 Then
      With imgSlider
         sldPos = ((.Left + x - SliderWidth / 2) \ 15) * 15
         If sldPos < 0 Then
            sldPos = 0
         ElseIf sldPos > picForm.Width - SliderWidth Then
            sldPos = picForm.Width - SliderWidth - 15
         End If
         .Left = sldPos
         sldScale = ((picForm.Width - 20) - SliderWidth) / (Max - Min)
         Value = (sldPos / sldScale) + Min
         For shp = 0 To 24
            If .Left > Shape1(shp).Left Then
               Shape1(shp).Visible = True
            Else
               Shape1(shp).Visible = False
               If Value = Min Then Shape1(0).Visible = False
            End If
         Next shp
         If Value > Max Then Value = Max
         If Value = Max Then
            Shape1(24).Visible = True
            Value = Max
         End If
      End With
       lblValue.Caption = Value
   End If
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblValue.Visible = True
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If ValueHide = True Then
      lblValue.Visible = False
   Else
      lblValue.Visible = True
   End If
End Sub

Private Sub Image1_Click()
   DropDownCtrl = False
   Image1.Visible = False
   Image2.Visible = True
End Sub

Private Sub Image2_Click()
   DropDownCtrl = True
   Image1.Visible = True
   Image2.Visible = False
End Sub

Private Sub lblCaption_Click()
If DisableDropDown = True Then Exit Sub
DropDownCtrl = Not DropDownCtrl
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If ValueHide = True Then
      lblValue.Visible = False
   Else
      lblValue.Visible = True
   End If
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim sldPos As Long
   Dim sldScale As Single
   Dim shp As Integer
 
   With imgSlider
      If y > 400 Then
         sldPos = ((x - SliderWidth / 2) \ 15) * 15
         If sldPos < 0 Then
            sldPos = 0
         ElseIf sldPos > picForm.Width - SliderWidth Then
            sldPos = picForm.Width - SliderWidth - 15
         End If
         .Left = sldPos
         sldScale = ((picForm.Width - 20) - .Width) / (Max - Min)
         Value = (sldPos / sldScale) + Min
         For shp = 0 To 24
          If Value = Max Then Shape1(24).Visible = True
            If .Left > Shape1(shp).Left Then
               Shape1(shp).Visible = True
            Else
               Shape1(shp).Visible = False
               If Value = Min Then Shape1(0).Visible = False
            End If
         Next shp
         If Value > Max Then Value = Max
         If Value = Max Then
            Shape1(24).Visible = True
            Value = Max
         End If
         lblValue.Caption = Value
         lblValue.Visible = True
      End If
   End With
End Sub

Public Property Get BarColor() As Bar
Attribute BarColor.VB_Description = "Changes the bar color to any of 8 choices."
    BarColor = m_BarColor
End Property

Public Property Let BarColor(NewBarColor As Bar)
Dim x As Integer
Dim y As Integer

   m_BarColor = NewBarColor
     If BarColor = 0 Then                               'Rainbow
         Shape1(0).FillColor = &HC000&
         Shape1(1).FillColor = &HC000&
         Shape1(2).FillColor = &HC000&
         Shape1(3).FillColor = &HC000&
         Shape1(4).FillColor = &HC000&
         Shape1(5).FillColor = &HFF00&
         Shape1(6).FillColor = &HFF00&
         Shape1(7).FillColor = &HFF00&
         Shape1(8).FillColor = &HFF00&
         Shape1(9).FillColor = &HFF00&
         Shape1(10).FillColor = &HC0FF00
         Shape1(11).FillColor = &HC0FF00
         Shape1(12).FillColor = &HC0FF00
         Shape1(13).FillColor = &HC0FFFF
         Shape1(14).FillColor = &HC0FFFF
         Shape1(15).FillColor = &H80FFFF
         Shape1(16).FillColor = &H80FFFF
         Shape1(17).FillColor = &H80FFFF
         Shape1(18).FillColor = &H80FFFF
         Shape1(19).FillColor = &HC0C0FF
         Shape1(20).FillColor = &HC0C0FF
         Shape1(21).FillColor = &H8080FF
         Shape1(22).FillColor = &H8080FF
         Shape1(23).FillColor = &HFF
         Shape1(24).FillColor = &HFF
      Else
         For x = 0 To 24
            y = y + 5
            If BarColor = 1 Then Shape1(x).FillColor = RGB(255 - y, 50, 50)                'Red
            If BarColor = 2 Then Shape1(x).FillColor = RGB(50, 50, 255 - y)                'Blue
            If BarColor = 3 Then Shape1(x).FillColor = RGB(50, 255 - y, 50)                'Green
            If BarColor = 4 Then Shape1(x).FillColor = RGB(255 - y, 255 - y, 50)         'Yellow
            If BarColor = 5 Then Shape1(x).FillColor = RGB(255 - y, 50, 255 - y)        'Purple
            If BarColor = 6 Then Shape1(x).FillColor = RGB(50, 255 - y, 255 - y)         'Turq
            If BarColor = 7 Then Shape1(x).FillColor = RGB(220 - y, 220 - y, 220 - y)   'Gray
          Next x
       End If
   PropertyChanged "BarColor"
End Property
Public Property Get DropDownCtrl() As Boolean
Attribute DropDownCtrl.VB_Description = "Shows or hides Main control window."
   DropDownCtrl = m_DropDownCtrl
End Property

Public Property Let DropDownCtrl(NewDropDownCtrl As Boolean)
   m_DropDownCtrl = NewDropDownCtrl
   If DropDownCtrl = True Then
      If DisableDropDown = True Then
         Image1.Visible = False
      Else
         Image1.Visible = True
      End If
      Image2.Visible = False
   Else
      Image1.Visible = False
      Image2.Visible = True
   End If
   PropertyChanged "DropDownCtrl"
   UserControl_Resize
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Enter descriptor text here."
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
   m_Caption = NewCaption
   lblCaption.Caption = Caption
   PropertyChanged "Caption"
End Property

Public Property Get DisableDropDown() As Boolean
Attribute DisableDropDown.VB_Description = "When False and DropDownCtrl True, control is visible all the time."
   DisableDropDown = m_DisableDropDown
End Property

Public Property Let DisableDropDown(NewDisableDropDown As Boolean)
   m_DisableDropDown = NewDisableDropDown
   If m_DisableDropDown = True Then
      Image1.Visible = False
      Image2.Visible = False
      DropDownCtrl = True
   Else
      If DropDownCtrl = True Then
         Image1.Visible = True
         Image2.Visible = False
      Else
         Image1.Visible = False
         Image2.Visible = True
      End If
   End If
   PropertyChanged "DisableDropDown"
End Property
Public Property Get Font() As Font
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
     Set UserControl.Font = NewFont
     PropertyChanged "Font"
     ' show changes while in IDE
     With lblValue
        .FontSize = UserControl.FontSize
        .FontBold = UserControl.FontBold
        .Font = UserControl.Font
      End With
      With lblCaption
        .FontSize = UserControl.FontSize
        .FontBold = UserControl.FontBold
        .Font = UserControl.Font
      End With
End Property

Public Property Get Min() As Long
   Min = m_Min
End Property

Public Property Let Min(NewMin As Long)
   If NewMin = Max Then
      MsgBox "Sorry but Min and Max cannot be equal.  " & Caption, vbOKOnly, "Please correct error in " & Caption
      Exit Property
   End If
   If NewMin > Max Then
      MsgBox "Sorry but Min is greater then Max.  " & Caption, vbOKOnly, "Please correct error in " & Caption
      Exit Property
   End If
   m_Min = NewMin
   PropertyChanged "Min"
   SliderPos
End Property

Public Property Get Max() As Long
   Max = m_Max
End Property

Public Property Let Max(NewMax As Long)
   If NewMax = Min Then
      MsgBox "Sorry but Max and Min cannot be equal.  " & Caption, vbOKOnly, "Please correct error in " & Caption
      Exit Property
   End If
   If NewMax < Min Then
      MsgBox "Sorry but Max is less than Min.  " & Caption, vbOKOnly, "Please correct error in " & Caption
      Exit Property
   End If
   m_Max = NewMax
   PropertyChanged "Max"
   SliderPos
End Property

Public Property Get Value() As Long
   Value = m_Value
End Property

Public Property Let Value(NewValue As Long)
   m_Value = NewValue
   If m_Value > Max Then m_Value = Max
   If m_Value < Min Then m_Value = Min
   PropertyChanged "Value"
   SliderPos
End Property

Public Property Get ValueHide() As Boolean
Attribute ValueHide.VB_Description = "If False, value will show all the time, otherwise, only when mouse button is pressed."
   Let ValueHide = m_ValueHide
End Property

Public Property Let ValueHide(ByVal NewValueHide As Boolean)
   Let m_ValueHide = NewValueHide
   If m_ValueHide = True Then
      lblValue.Visible = False
   Else
      lblValue.Visible = True
   End If
   PropertyChanged "ValueHide"
End Property

Public Property Get TickColor() As OLE_COLOR
Attribute TickColor.VB_Description = "Select the color for the marks."
   TickColor = m_TickColor
End Property

Public Property Let TickColor(NewTickColor As OLE_COLOR)
   Dim x As Integer
   
   m_TickColor = NewTickColor
   Line1.BorderColor = m_TickColor
   Line2.BorderColor = m_TickColor
   Line3.BorderColor = m_TickColor
   Line4.BorderColor = m_TickColor
   Line5.BorderColor = m_TickColor
   Line6.BorderColor = m_TickColor
   'if you want bar border color to change then uncomment following lines
  ' For x = 0 To 24
  '     Shape1(x).BorderColor = m_TickColor
   'Next x
   PropertyChanged "TickColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets the main background to any color."
   BackColor = m_BackColor
End Property

Public Property Let BackColor(NewBackColor As OLE_COLOR)
   m_BackColor = NewBackColor
   picForm.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get CapBkgdColor() As OLE_COLOR
Attribute CapBkgdColor.VB_Description = "Change the caption background color."
   CapBkgdColor = m_CapBkgdColor
End Property

Public Property Let CapBkgdColor(NewCapBkgdColor As OLE_COLOR)
   m_CapBkgdColor = NewCapBkgdColor
   lblCaption.BackColor = m_CapBkgdColor
   PropertyChanged "CapBkgdColor"
End Property

Public Property Get CapFontColor() As OLE_COLOR
Attribute CapFontColor.VB_Description = "Set the color of the caption font and the value font"
   CapFontColor = m_CapFontColor
End Property

Public Property Let CapFontColor(NewCapFontColor As OLE_COLOR)
   m_CapFontColor = NewCapFontColor
   lblCaption.ForeColor = m_CapFontColor
   lblValue.ForeColor = m_CapFontColor
   PropertyChanged "CapFontColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
   DropDownCtrl = PropBag.ReadProperty("DropDownCtrl", m_def_DropDownCtrl)
   Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   DisableDropDown = PropBag.ReadProperty("DisableDropDown", m_def_DisableDropDown)
   Min = PropBag.ReadProperty("Min", m_def_Min)
   Max = PropBag.ReadProperty("Max", m_def_Max)
   Value = PropBag.ReadProperty("Value", m_def_Value)
   Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
   TickColor = PropBag.ReadProperty("TickColor", m_def_TickColor)
   BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
   CapBkgdColor = PropBag.ReadProperty("CapBkgdColor", m_def_CapBkgdColor)
   CapFontColor = PropBag.ReadProperty("CapFontColor", m_def_CapFontColor)
   BarColor = PropBag.ReadProperty("BarColor", m_def_BarColor)
   ValueHide = PropBag.ReadProperty("ValueHide", m_def_ValueHide)
     With lblValue
        .FontSize = UserControl.FontSize
        .FontBold = UserControl.FontBold
        .Font = UserControl.Font
      End With
      With lblCaption
        .FontSize = UserControl.FontSize
        .FontBold = UserControl.FontBold
        .Font = UserControl.Font
      End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      Call .WriteProperty("DropDownCtrl", m_DropDownCtrl, m_def_DropDownCtrl)
      Call .WriteProperty("Caption", m_Caption, m_def_Caption)
      Call .WriteProperty("DisableDropDown", m_DisableDropDown, m_def_DisableDropDown)
      Call .WriteProperty("Min", m_Min, m_def_Min)
      Call .WriteProperty("Max", m_Max, m_def_Max)
      Call .WriteProperty("Value", m_Value, m_def_Value)
      Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
      Call .WriteProperty("TickColor", m_TickColor, m_def_TickColor)
      Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
      Call .WriteProperty("CapBkgdColor", m_CapBkgdColor, m_def_CapBkgdColor)
      Call .WriteProperty("CapFontColor", m_CapFontColor, m_def_CapFontColor)
      Call .WriteProperty("BarColor", m_BarColor, m_def_BarColor)
      Call .WriteProperty("ValueHide", m_ValueHide, m_def_ValueHide)
   End With
End Sub

Private Function SliderPos()
Dim sldScale  As Single
Dim shp As Integer
   
    With imgSlider
        If Max - Min <> 0 Then
                sldScale = (picForm.Width - SliderWidth) / (Max - Min)
                .Left = (Value - Min) * sldScale - 3
        End If
          For shp = 0 To 24
            If .Left + 15 > Shape1(shp).Left Then
               Shape1(shp).Visible = True
            Else
               Shape1(shp).Visible = False
               If Value = Min Then Shape1(0).Visible = False
            End If
         Next shp
    End With
    lblValue.Caption = Value
End Function
