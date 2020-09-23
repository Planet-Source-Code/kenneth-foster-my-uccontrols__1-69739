VERSION 5.00
Begin VB.UserControl msDial 
   BackColor       =   &H00545249&
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1725
   ScaleWidth      =   2400
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   1545
      Picture         =   "msDial_KF.ctx":0000
      ScaleHeight     =   8100
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   45
      Width           =   555
   End
   Begin VB.PictureBox knob1 
      BackColor       =   &H00545249&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   45
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   45
      Width           =   555
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   315
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Dial"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   675
      Width           =   600
   End
End
Attribute VB_Name = "msDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'the best I can find, the author of this control is Makis Charalambous
'I've reduced the orig size from 307k to 26k plus I've added a few things

Dim oldX As Single, oldY As Single
Dim i As Integer, a As Single, lastVal As Integer, INN As Integer, OTN As Integer

Event DialChange(nValue As Integer)

'Default Property Values:
Const m_def_Value = 0
Const m_def_Caption = "Dial"
Const m_def_CaptionShow = True
Const m_def_Enabled = True
Const m_def_CaptionStyle = 0

Enum eStyle
   ShowasValue = 0
   ShowasLabel = 1
End Enum

Dim m_Caption As String
Dim m_CaptionShow As Boolean
Dim m_Value As Integer
Dim m_Enabled As Boolean
Dim m_CaptionStyle As Integer

'cursor movement enums--------------
Private Type RECT
Left As Integer
top As Integer
Right As Integer
bottom As Integer
End Type
Private Type POINT
X As Long
Y As Long
End Type
'-----------------------------------

' API
'limit cursor movement to picbox------------------------
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
'-------------------------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub knob1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim client As RECT
Dim upperleft As POINT

   If Button <> vbLeftButton Then Exit Sub
   'contain cursor to picbox-------------
   GetClientRect knob1.hWnd, client
   upperleft.X = client.Left
   upperleft.Y = client.top
   ClientToScreen knob1.hWnd, upperleft
   OffsetRect client, upperleft.X, upperleft.Y
   ClipCursor client
   '--------------------------------------
End Sub

Private Sub Knob1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    
        If Y < knob1.ScaleHeight / 2 Then
            If X > oldX Then m_Value = m_Value + 1
            If X < oldX Then m_Value = m_Value - 1
        Else
            If X > oldX Then m_Value = m_Value - 1
            If X < oldX Then m_Value = m_Value + 1
        End If
        
        If X < knob1.ScaleWidth / 2 Then
            If Y > oldY Then m_Value = m_Value - 1
            If Y < oldY Then m_Value = m_Value + 1
        Else
            If Y > oldY Then m_Value = m_Value + 1
            If Y < oldY Then m_Value = m_Value - 1
        End If
           
        If m_Value > 100 Then m_Value = 100
        If m_Value < 0 Then m_Value = 0
        
        RaiseEvent DialChange(m_Value)
        getpicframe
        oldX = X
        oldY = Y
    End If
     
End Sub

Private Sub knob1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'release cursor from picbox
   ClipCursor ByVal 0&
End Sub

Private Sub UserControl_Initialize()
    
    knob1.Picture = Picture2.Picture
    UserControl.Width = 675
    UserControl.Height = 1005
    knob1.Height = 550
    knob1.Width = 625
    Label1.top = 800
    Label1.Left = 0
    Label1.Width = UserControl.Width
    Label1.BackColor = &H545249
    CaptionShow = True
    DrawDot knob1, 0
    
End Sub

Private Sub UserControl_Resize()
    
    UserControl.Width = 675
    If m_CaptionShow = True Then
       UserControl.Height = 1005
    Else
       UserControl.Height = 1005 - Label1.Height
    End If
    knob1.Height = 690
    knob1.Width = 625
    Label1.Caption = m_Caption
   CaptionStyle = ShowasValue
End Sub
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    If m_Value > 100 Then m_Value = 0
    If m_Value < 0 Then m_Value = 100
    getpicframe
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
   m_Caption = NewCaption
   Label1.Caption = m_Caption
   PropertyChanged "Caption"
End Property

Public Property Get CaptionShow() As Boolean
   CaptionShow = m_CaptionShow
End Property

Public Property Let CaptionShow(NewCaptionShow As Boolean)
   m_CaptionShow = NewCaptionShow
   Label1.Visible = m_CaptionShow
   If m_CaptionShow = True Then
      UserControl.Height = 1005
   Else
      UserControl.Height = 1005 - Label1.Height
  End If
   PropertyChanged "CaptionShow"
End Property

Public Property Get CaptionStyle() As eStyle
   CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal newCaptionStyle As eStyle)
   m_CaptionStyle = newCaptionStyle
   If CaptionStyle = ShowasValue Then
      Label1.Caption = m_Value
   Else
      Label1.Caption = Caption
   End If
   PropertyChanged "CaptionStyle"
End Property
Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(NewEnabled As Boolean)
   m_Enabled = NewEnabled
   UserControl.Enabled = m_Enabled
   PropertyChanged "Enabled"
End Property

Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Caption = Extender.Name
    m_CaptionShow = m_def_CaptionShow
    m_Enabled = m_def_Enabled
    m_CaptionStyle = m_def_CaptionStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", m_def_Value)
    Caption = PropBag.ReadProperty("Caption", Extender.Name)
    CaptionShow = PropBag.ReadProperty("CaptionShow", m_def_CaptionShow)
    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
   Call PropBag.WriteProperty("Caption", m_Caption, Extender.Name)
   Call PropBag.WriteProperty("CaptionShow", m_CaptionShow, m_def_CaptionShow)
   Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
   Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
End Sub

Private Sub getpicframe()
     Dim fam As Integer
     
     If m_Value >= 0 Then fam = m_Value
     If m_Value >= 12 And m_Value < 24 Then fam = m_Value - 12
     If m_Value >= 24 And m_Value < 36 Then fam = m_Value - 24
     If m_Value >= 36 And m_Value < 48 Then fam = m_Value - 36
     If m_Value >= 48 And m_Value < 60 Then fam = m_Value - 48
     If m_Value >= 60 And m_Value < 72 Then fam = m_Value - 60
     If m_Value >= 72 And m_Value < 84 Then fam = m_Value - 72
     If m_Value >= 84 And m_Value < 96 Then fam = m_Value - 84
     If m_Value >= 96 And m_Value <= 100 Then fam = m_Value - 96
    
      If CaptionStyle = ShowasValue Then Label1.Caption = m_Value
      BitBlt knob1.hdc, 0, 0, 37, 45, Picture2.hdc, 0, fam * 45, vbSrcCopy
      DrawDot knob1, m_Value
End Sub

Private Sub DrawDot(Obj As Object, ByVal pos As Single)
Dim X As Long
Dim Y As Long
Dim degree As Double
Dim radiusX As Long
Dim radiusY As Long
Dim convert As Double
Dim CenterX As Integer
Dim CenterY As Integer

    pos = pos - 50  'this sets the 0 or start position of scale and 100 or end of scale position
                             ' this example is -50 to 50 for a scale of 100
    With Obj
       ' .Height = 800
       ' .Width = 800
        .AutoRedraw = True
        .FillStyle = 0
    End With
    
    'the size of the circle
    radiusX = 245
    radiusY = 245
    
    'center of the circle
    CenterX = (285)
    CenterY = (280)
    
    'Obj.Cls
    'Obj.Refresh
    
    degree = pos * 3    'this number (3) also affects the circle . ex.3.5 will yield 360 degrees. the 3 is 270, a 2 about 180
    convert = 3.14159265358979 / 180 'from Radian to Degree
    X = CenterX - (Sin(-degree * convert) * radiusX)
    Y = CenterY - (Sin((90 + (degree)) * convert) * radiusY)
    
    'draw the dot
    Obj.Circle (X, Y), 8, vbWhite
   
End Sub
