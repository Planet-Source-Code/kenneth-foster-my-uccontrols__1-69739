VERSION 5.00
Begin VB.UserControl GradButton 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ScaleHeight     =   780
   ScaleWidth      =   1275
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1245
   End
End
Attribute VB_Name = "GradButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Control   : GradButton
' DateTime  : 4/25/2005 11:06
' Author    : Ken Foster
' Purpose   : Make Gradient Buttons
' Credits   : Middle-out gradient code by Matthew R. Usner
'---------------------------------------------------------------------------------------
Option Explicit

   ' Constants
   Const def_m_Caption = "GradButt"
   Const def_m_ColorOuter = vbBlack
   Const def_m_ColorMid = vbWhite
   Const def_m_ForeColor = vbBlack
   Const def_m_FontposX = 8
   Const def_m_FontposY = 3
   Const m_def_ClkColor = vbBlack
   Const m_def_ClkForeColor = vbBlack
   
   ' varibles
   Dim m_FontposX As Integer
   Dim m_FontposY As Integer
   Dim m_ForeColor As Long
   Dim m_ColorOuter As OLE_COLOR
   Dim m_ColorMid As OLE_COLOR
   Dim m_ClkColor As OLE_COLOR
   Dim m_ClkForeColor As OLE_COLOR
   Dim m_Caption As String
   Dim lcolor1 As Long
   Dim lcolor2 As Long
   Dim Store As Long
   Dim Store2 As Long
   ' events
   Event Click()
   Event MouseDown()
   Event MouseMove()
   Event Mouseup()
   
   ' Declarations
   Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
   Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
   Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
   Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
   Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
   Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Sub DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
   ByVal lCol1 As Long, ByVal lCol2 As Long)
   
   Dim yStart As Long
   Dim xEnd    As Long, yEnd   As Long
   Dim X1      As Long, Y1     As Long
   Dim X2      As Long, Y2     As Long
   Dim lRange  As Long
   Dim iQ      As Integer
   Dim lPtr    As Long, lInc   As Long
   Dim lCols() As Long, lCols2() As Long
   Dim hPO     As Long, hPN    As Long
   Dim r       As Long
   Dim X       As Long, xUp    As Long
   Dim b1(2)   As Byte, b2(2)  As Byte, b3(2) As Byte
   Dim p       As Single, ip   As Single
   Dim Y As Long
 
   
   lInc = 1
   xEnd = (lWidth - 1) '/ 2
   yEnd = (lHeight - 1) '/ 2 ' /2 for top half only
   
   lRange = lHeight + yStart - 1
   
   X1 = IIf(iQ Mod 2, 0, xEnd)
   X2 = IIf(X1, -1, lWidth)
   '  -------------------------------------------------------------------
   '  Fill in the color array with the interpolated color values.
   '  -------------------------------------------------------------------
   ReDim lCols(lRange)
   ReDim lCols2(lRange)
   
   ' Get the r, g, b components of each color.
   RtlMoveMemory b1(0), lCol1, 3
   RtlMoveMemory b2(0), lCol2, 3
   RtlMoveMemory b3(0), 0, 3
   xUp = UBound(lCols)
   
   '        get the full color array in lCols2.
   For X = 0 To xUp
      ' Get the position and the 1 - position.
      p = X / xUp
      ip = 1 - p
      ' Interpolate the value at the current position.
      lCols2(X) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
   Next X
   '        put the array in first half of lcols1
   Y = 0
   For X = 0 To xUp Step 2
      lCols(Y) = lCols2(X)
      Y = Y + 1
   Next X
   For X = xUp - 1 To 1 Step -2
      lCols(Y) = lCols2(X)
      If Y < xUp Then Y = Y + 1
   Next X
   
   For Y1 = -yStart To yEnd
      hPN = CreatePen(0, 1, lCols(lPtr))
      hPO = SelectObject(hDC, hPN)
      MoveTo hDC, X1, Y1, ByVal 0&
      LineTo hDC, X2, Y2
      r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
      lPtr = lPtr + lInc
      Y2 = Y2 + 1
   Next Y1
   
End Sub

Public Sub UpdateGradient()
   Dim p As Integer
   Dim b As Integer
   
   lcolor1 = ColorOuter
   lcolor2 = ColorMid
   
   DrawGradient picButton.hDC, picButton.ScaleWidth, picButton.ScaleHeight, lcolor1, lcolor2
   picButton.AutoRedraw = True
   picButton.Font = Font
   picButton.ScaleMode = 3
   picButton.CurrentX = FontposX
   picButton.CurrentY = FontposY
   picButton.FontSize = Font.Size
   picButton.ForeColor = ForeColor
   picButton.Print Caption
   
End Sub

Private Sub picButton_Click()

   RaiseEvent Click
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Store = ColorOuter
Store2 = ForeColor
ColorOuter = m_ClkColor
ForeColor = m_ClkForeColor
RaiseEvent MouseDown
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ColorOuter = Store
ForeColor = Store2
RaiseEvent Mouseup
End Sub

Private Sub UserControl_Initialize()

   Caption = def_m_Caption
   ColorOuter = def_m_ColorOuter
   ColorMid = def_m_ColorMid
   ForeColor = def_m_ForeColor
   FontposX = def_m_FontposX
   FontposY = def_m_FontposY
   UpdateGradient
End Sub

Private Sub UserControl_Resize()

   picButton.Height = UserControl.Height
   picButton.Width = UserControl.Width
   UpdateGradient
End Sub

Private Sub Usercontrol_ReadProperties(Propbag As PropertyBag)

   m_Caption = Propbag.ReadProperty("Caption", def_m_Caption)
   Set Font = Propbag.ReadProperty("Font", Ambient.Font)
   m_ColorOuter = Propbag.ReadProperty("ColorOuter", m_ColorOuter)
   m_ColorMid = Propbag.ReadProperty("ColorMid", m_ColorMid)
   m_ClkColor = Propbag.ReadProperty("ClkColor", m_ClkColor)
   m_ClkForeColor = Propbag.ReadProperty("ClkForeColor", m_ClkForeColor)
   m_ForeColor = Propbag.ReadProperty("ForeColor", m_ForeColor)
   m_FontposX = Propbag.ReadProperty("FontposX", m_FontposX)
   m_FontposY = Propbag.ReadProperty("FontposY", m_FontposY)
   UpdateGradient
End Sub

Private Sub Usercontrol_WriteProperties(Propbag As PropertyBag)

   Propbag.WriteProperty "Caption", m_Caption, def_m_Caption
   Propbag.WriteProperty "Font", Font, Ambient.Font
   Propbag.WriteProperty "ColorOuter", m_ColorOuter, def_m_ColorOuter
   Propbag.WriteProperty "ColorMid", m_ColorMid, def_m_ColorMid
   Propbag.WriteProperty "ClkColor", m_ClkColor, m_def_ClkColor
   Propbag.WriteProperty "ClkForeColor", m_ClkForeColor, m_def_ClkForeColor
   Propbag.WriteProperty "ForeColor", m_ForeColor, def_m_ForeColor
   Propbag.WriteProperty "FontposX", m_FontposX, def_m_FontposX
   Propbag.WriteProperty "FontposY", m_FontposY, def_m_FontposY
End Sub

Public Property Get Caption() As String

   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)

   m_Caption = NewCaption
   PropertyChanged "Caption"
   UpdateGradient
End Property

Public Property Get Font() As Font

   Set Font = picButton.Font
End Property

Public Property Set Font(NewFont As Font)

   Set picButton.Font = NewFont
   PropertyChanged ("Font")
   UpdateGradient
End Property

Public Property Get ColorOuter() As OLE_COLOR

   ColorOuter = m_ColorOuter
End Property

Public Property Let ColorOuter(NewColorOuter As OLE_COLOR)

   m_ColorOuter = NewColorOuter
   lcolor1 = m_ColorOuter
   PropertyChanged "ColorOuter"
   UpdateGradient
End Property

Public Property Get ColorMid() As OLE_COLOR

   ColorMid = m_ColorMid
End Property

Public Property Let ColorMid(NewColorMid As OLE_COLOR)

   m_ColorMid = NewColorMid
   lcolor2 = m_ColorMid
   PropertyChanged "ColorMid"
   UpdateGradient
End Property

Public Property Get ClkColor() As OLE_COLOR

   ClkColor = m_ClkColor
End Property

Public Property Let ClkColor(NewClkColor As OLE_COLOR)

   m_ClkColor = NewClkColor
   lcolor1 = m_ClkColor
   PropertyChanged "ClkColor"
   UpdateGradient
End Property
Public Property Get ForeColor() As OLE_COLOR

   ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(NewForeColor As OLE_COLOR)

   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   UpdateGradient
End Property

Public Property Get ClkForeColor() As OLE_COLOR

   ClkForeColor = m_ClkForeColor
End Property

Public Property Let ClkForeColor(NewClkForeColor As OLE_COLOR)

   m_ClkForeColor = NewClkForeColor
   PropertyChanged "ClkForeColor"
   UpdateGradient
End Property
Public Property Get FontposX() As Integer
   FontposX = m_FontposX
End Property

Public Property Let FontposX(NewFontposX As Integer)
   m_FontposX = NewFontposX
   PropertyChanged ("FontposX")
   UpdateGradient
End Property

Public Property Get FontposY() As Integer
   FontposY = m_FontposY
End Property

Public Property Let FontposY(NewFontposY As Integer)
   m_FontposY = NewFontposY
   PropertyChanged ("FontposY")
   UpdateGradient
End Property
