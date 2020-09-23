VERSION 5.00
Begin VB.UserControl ThumbWheel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   87
   ToolboxBitmap   =   "ThumbWheel.ctx":0000
   Begin VB.PictureBox picV_MSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   615
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picH_MSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   840
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picV_SND 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   612
      Left            =   465
      Picture         =   "ThumbWheel.ctx":00D0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picWheel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   855
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Width           =   195
   End
   Begin VB.PictureBox picH_SND 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   840
      Picture         =   "ThumbWheel.ctx":05EB
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   570
      Left            =   1155
      Shape           =   4  'Rounded Rectangle
      Top             =   930
      Width           =   255
   End
   Begin VB.Image imgVBack 
      Height          =   885
      Left            =   0
      Picture         =   "ThumbWheel.ctx":0AFD
      Top             =   435
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgHBack 
      Height          =   420
      Left            =   495
      Picture         =   "ThumbWheel.ctx":120D
      Top             =   15
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "ThumbWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'ThumbWheel Control
'
'Author Ben Vonk
'24-08-2005 First version (based on Nero's 'ThumbWheel Control' at http://www.planet-source-code.com/vb/authors/ShowBio.asp?lngAuthorId=663524985&lngWId=1)

Option Explicit

' Public Events
Public Event Change()
Attribute Change.VB_Description = "This event is fired when a change has  occured."

' Public Constants
Public Enum Orientations
   Horizontal
   Vertical
End Enum

' Private Variables
Private Clicked        As Boolean
Private Increment      As Integer
Private m_Max          As Integer
Private m_Min          As Integer
Private m_Orientation  As Integer
Private m_Value        As Integer
Private WheelPosition  As Integer
Private m_ShadeControl As Long
Private m_ShadeWheel   As Long
Private LastX          As Single
Private LastY          As Single

Const m_def_Rollover = False
Const m_def_Custom = False
Const m_def_CustomBackColor = vbWhite
Const m_def_CustomBorder = True
Const m_def_CustomBorderColor = vbBlack
Const m_def_Enabled = True

Dim m_CustomBorder As Boolean
Dim m_CustomBorderColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_Custom As Boolean
Dim m_CustomBackColor As OLE_COLOR

Dim m_Rollover As Boolean   'allows wheel to go from max to min/min to max or stop at max and min

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
Private Declare Function BitBlt Lib "GDI32.dll" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "GDI32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Property Get Custom() As Boolean
Attribute Custom.VB_Description = "Use the gray graphics or custom colors."
    Custom = m_Custom
End Property

Public Property Let Custom(NewCustom As Boolean)
    m_Custom = NewCustom
    PropertyChanged "Custom"
    UserControl_Resize
End Property

Public Property Get CustomBackColor() As OLE_COLOR
Attribute CustomBackColor.VB_Description = "Main background color"
    CustomBackColor = m_CustomBackColor
End Property

Public Property Let CustomBackColor(NewCustomBackColor As OLE_COLOR)
    m_CustomBackColor = NewCustomBackColor
    PropertyChanged "CustomBackColor"
    Shape1.BackColor = m_CustomBackColor
    UserControl_Resize
End Property

Public Property Get CustomBorder() As Boolean
Attribute CustomBorder.VB_Description = "Show a border or not show border."
    CustomBorder = m_CustomBorder
End Property

Public Property Let CustomBorder(NewCustomBorder As Boolean)
    m_CustomBorder = NewCustomBorder
    PropertyChanged "CustomBorder"
    If CustomBorder = True Then
       Shape1.BorderStyle = 1
    Else
       Shape1.BorderStyle = 0
    End If
End Property

Public Property Get CustomBorderColor() As OLE_COLOR
Attribute CustomBorderColor.VB_Description = "Change color of border"
    CustomBorderColor = m_CustomBorderColor
End Property

Public Property Let CustomBorderColor(NewCustomBorderColor As OLE_COLOR)
    m_CustomBorderColor = NewCustomBorderColor
    PropertyChanged "CustomBorderColor"
    Shape1.BorderColor = m_CustomBorderColor
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(NewEnabled As Boolean)
    m_Enabled = NewEnabled
    PropertyChanged "Enabled"
    If Enabled = True Then
       UserControl.Enabled = True
    Else
       UserControl.Enabled = False
    End If
End Property

Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
   Max = m_Max
End Property

Public Property Let Max(ByVal NewMax As Integer)

   If NewMax < m_Min Then NewMax = m_Max
   m_Max = NewMax
   m_Value = m_Max
   PropertyChanged "Max"
End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Returns/sets a scroll bar position's minimum Value property setting."
   Min = m_Min
End Property

Public Property Let Min(ByVal NewMin As Integer)

   If NewMin > m_Max Then NewMin = m_Min
   
   m_Min = NewMin
   m_Value = m_Min
   PropertyChanged "Min"

End Property

Public Property Get Orientation() As Orientations
Attribute Orientation.VB_Description = "Returns/sets a scroll bar orientation, horizontal or vertical."
   Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal NewOrientation As Orientations)

   m_Orientation = NewOrientation
   PropertyChanged "Orientation"
   Call UserControl_Resize
End Property

Public Property Get Rollover() As Boolean
Attribute Rollover.VB_Description = "Lets max value rollover to min value and vise versa, or stop at min and max values."
    Rollover = m_Rollover
End Property

Public Property Let Rollover(NewRollover As Boolean)
    m_Rollover = NewRollover
    PropertyChanged "Rollover"
End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of a object."
   Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Integer)

   If NewValue > m_Max Then
      NewValue = m_Min
   ElseIf NewValue < m_Min Then
      NewValue = m_Max
   End If
  
   If m_Value <> NewValue Then
      Increment = 1 - (2 And NewValue > m_Value)
      Call WheelMove
   End If
     
   m_Value = NewValue
   PropertyChanged "Value"
   RaiseEvent Change
End Property

' draws the thumbwheel
Private Sub DrawWheel()

   With picWheel
      If m_Orientation = Horizontal Then
         StretchBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picH_SND.hDC, 0, WheelPosition, .ScaleWidth, 1, vbSrcCopy
         'BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picH_MSK.hDC, 16, 18, vbSrcAnd
     Else
         StretchBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picV_SND.hDC, WheelPosition, 0, 1, .ScaleHeight, vbSrcCopy
         'BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picV_MSK.hDC, 0, 0, vbSrcAnd
      End If
      
      .Refresh
      DoEvents
   End With

End Sub

' draws the moving thumbwheel effect
Private Sub WheelMove()
Dim intCntr As Integer
   If Increment = 0 Then Exit Sub
   If Not Rollover Then
        Select Case Increment
            Case Is < 0
                If m_Value - 1 < m_Min Then Exit Sub
            Case Is > 0
                If m_Value + 1 > m_Max Then Exit Sub
        End Select
    End If
    For intCntr = Sgn(Increment) To Increment Step Sgn(Increment)
   WheelPosition = WheelPosition + Sgn(Increment)
   
   If WheelPosition > 8 Then WheelPosition = 0
   If WheelPosition < 0 Then WheelPosition = 8
                m_Value = m_Value + Sgn(Increment) 'Increment the VALUE
                If Increment < 0 Then
                    If m_Value < m_Min Then
                        If Rollover Then
                            m_Value = m_Max
                        Else
                            m_Value = m_Min
                            Exit Sub
                        End If
                    End If
                    RaiseEvent Change
                Else
                    If m_Value > m_Max Then
                        If Rollover Then
                            m_Value = m_Min
                        Else
                            m_Value = m_Max
                            Exit Sub
                        End If
                    End If
                    RaiseEvent Change
                End If
        DrawWheel
    Next intCntr
   RaiseEvent Change
End Sub

Private Sub picWheel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim client As RECT
Dim upperleft As POINT

   If Button <> vbLeftButton Then Exit Sub
   'contain cursor to picbox-------------
   GetClientRect picWheel.hWnd, client
   upperleft.X = client.Left
   upperleft.Y = client.top
   ClientToScreen picWheel.hWnd, upperleft
   OffsetRect client, upperleft.X, upperleft.Y
   ClipCursor client
   '--------------------------------------
   Clicked = True
   LastX = X
   LastY = Y
End Sub

Private Sub picWheel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   picWheel.ToolTipText = Extender.ToolTipText
   
   If Button <> vbLeftButton Then Exit Sub
   
   If Clicked = True Then
     
      If m_Orientation = Horizontal Then Increment = X - LastX
      If m_Orientation = Vertical Then Increment = Y - LastY

      Call WheelMove
      
      LastX = X
      LastY = Y
   End If
End Sub

Private Sub picWheel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then Clicked = False
   'release cursor from picbox
   ClipCursor ByVal 0&
End Sub

Private Sub UserControl_InitProperties()

   m_Orientation = Horizontal
   m_ShadeControl = vbWhite
   m_ShadeWheel = vbWhite
   m_Max = 100
   Value = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_Max = .ReadProperty("Max", 100)
      m_Min = .ReadProperty("Min", 0)
      m_Orientation = .ReadProperty("Orientation", Horizontal)
      m_Value = .ReadProperty("Value", 0)
      m_Rollover = .ReadProperty("Rollover", m_def_Rollover)
      Custom = .ReadProperty("Custom", m_def_Custom)
      CustomBackColor = .ReadProperty("CustomBackColor", m_def_CustomBackColor)
      CustomBorder = .ReadProperty("CustomBorder", m_def_CustomBorder)
      CustomBorderColor = .ReadProperty("CustomBorderColor", m_def_CustomBorderColor)
      Enabled = .ReadProperty("Enabled", m_def_Enabled)
   End With
End Sub

Private Sub UserControl_Resize()

Dim intTwipsX As Integer
Dim intTwipsY As Integer

   imgHBack.Visible = False
   imgVBack.Visible = False
   intTwipsX = Screen.TwipsPerPixelX
   intTwipsY = Screen.TwipsPerPixelY
   
   ' calculate wheel position for horizontal
   If m_Orientation = Horizontal Then
      With imgHBack
         .top = 0
         .Left = 0
         If Custom = True Then
            Shape1.top = 0
            Shape1.Left = 0
            Shape1.Width = .Width - 14
            Shape1.Height = .Height - 12
            UserControl.Width = 45 * intTwipsX
            UserControl.Height = 16 * intTwipsY
           .Visible = False
           Shape1.Visible = True
           picWheel.Move 2, 2, .Width - 18, .Height - 16
         Else
            If Height <> .Height * intTwipsY Then Height = .Height * intTwipsY
            If Width <> .Width * intTwipsX Then Width = .Width * intTwipsX
            .Visible = True
            Shape1.Visible = False
            picWheel.Move 8, 8, .Width - 18, .Height - 16
         End If
      End With
   End If
   
   ' calculate wheel position for vertical
   If m_Orientation = Vertical Then
      With imgVBack
         .top = 0
         .Left = 0
         If Custom = True Then
            Shape1.top = 0
            Shape1.Left = 0
            Shape1.Width = .Width - 12
            Shape1.Height = .Height - 14
            UserControl.Width = 16 * intTwipsX
            UserControl.Height = 45 * intTwipsY
           .Visible = False
           Shape1.Visible = True
           picWheel.Move 2, 2, .Width - 16, .Height - 18
         Else
            If Height <> .Height * intTwipsY Then Height = .Height * intTwipsY
            If Width <> .Width * intTwipsX Then Width = .Width * intTwipsX
            .Visible = True
            Shape1.Visible = False
            picWheel.Move 8, 8, .Width - 16, .Height - 18
         End If
      End With
   End If
   Call DrawWheel
End Sub

Private Sub UserControl_Terminate()
   ClipCursor ByVal 0&
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Max", m_Max, 100
      .WriteProperty "Min", m_Min, 0
      .WriteProperty "Orientation", m_Orientation, Horizontal
      .WriteProperty "Value", m_Value, 0
      .WriteProperty "Rollover", m_Rollover, m_def_Rollover
      .WriteProperty "Custom", m_Custom, m_def_Custom
      .WriteProperty "CustomBackColor", m_CustomBackColor, m_def_CustomBackColor
      .WriteProperty "CustomBorder", m_CustomBorder, m_def_CustomBorder
      .WriteProperty "CustomBorderColor", m_CustomBorderColor, m_def_CustomBorderColor
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
   End With
End Sub

