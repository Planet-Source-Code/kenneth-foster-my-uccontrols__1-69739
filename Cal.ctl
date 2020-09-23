VERSION 5.00
Begin VB.UserControl Cal 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ScaleHeight     =   3360
   ScaleWidth      =   3945
   ToolboxBitmap   =   "Cal.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   30
      ScaleHeight     =   2895
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   315
      Width           =   3885
      Begin VB.CommandButton cmdCurrentDate 
         Caption         =   "Today"
         Height          =   300
         Left            =   2775
         TabIndex        =   3
         Top             =   60
         Width           =   900
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "cboYear"
         Top             =   60
         Width           =   1215
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Text            =   "cboMonth"
         Top             =   60
         Width           =   1260
      End
      Begin VB.Line Line11 
         X1              =   3150
         X2              =   3150
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line10 
         X1              =   2655
         X2              =   2655
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line9 
         X1              =   2145
         X2              =   2145
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line8 
         X1              =   1635
         X2              =   1635
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line7 
         X1              =   1125
         X2              =   1125
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line6 
         X1              =   615
         X2              =   615
         Y1              =   735
         Y2              =   2910
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   3855
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   3855
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   3840
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3855
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3810
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblDatesT 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   420
         TabIndex        =   13
         Top             =   3135
         Width           =   360
      End
      Begin VB.Label lblDates 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3225
         TabIndex        =   11
         Top             =   465
         Width           =   555
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2775
         TabIndex        =   10
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Thur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2190
         TabIndex        =   9
         Top             =   465
         Width           =   540
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1635
         TabIndex        =   8
         Top             =   465
         Width           =   525
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1140
         TabIndex        =   7
         Top             =   465
         Width           =   465
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   630
         TabIndex        =   6
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lblDayNames 
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   465
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         TabIndex        =   4
         Top             =   435
         Width           =   3660
      End
   End
   Begin VB.Label lblDateBar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Show"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3060
      TabIndex        =   14
      Top             =   30
      Width           =   735
   End
End
Attribute VB_Name = "Cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*********************************************
   '*  KF Dropdown Calendar                               *
   '*  By Ken Foster                            *
   '*    2005                                   *
   '*  Freeware--use or change any way you want *
   '*********************************************
   'Form animation code by Jim Jose
   '*********************************************
   'Click on Date label to open/close calendar
   'Clicking on empty space in calendar box will close also
   '*********************************************
   Option Explicit
   
   '[APIs]
   Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
   Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
   Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
   Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
   Private Declare Function CreateRectRgn Lib "GDI32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
   Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
   Private Declare Function CombineRgn Lib "GDI32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
   Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
   Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
     
   '[Constants]
   Private Const RGN_AND As Long = 1
   Private Const RGN_OR As Long = 2
   Private Const RGN_XOR As Long = 3
   Private Const RGN_COPY As Long = 5
   Private Const RGN_DIFF As Long = 4
   Private Const CB_SETDROPPEDCONTROLRECT = &H160
   Private Const CB_GETITEMHEIGHT = &H154

'[Event Enum]
Private Enum AnimeEventEnum
   aUnload = 0
   aload = 1
 End Enum

'[Effect Enum]
Private Enum AnimeEffectEnum
   eAppearFromTop = 2
   eAppearFromBottom = 3
End Enum

Private Type RECT
   Left As Long
   top As Long
   Right As Long
   bottom As Long
End Type

Private Type POINTAPI
   X   As Long
   Y   As Long
End Type

Private Enum eExpandBy
   Percent50 = 0
   Percent75 = 1
   DoubleWidth = 2
   TripleWidth = 3
   QuadWidth = 4
   NoExpand = 5
End Enum

Private Enum eExpandType
   WidthOnly = 0
   HeightOnly = 1
   HeightAndWidth = 2
End Enum

Const m_def_BackColor = vbWhite
Const m_def_DateDisplay = True
Const m_def_SHButton = True

Dim TheCaption As String
Dim CurMonth As Single
Dim CurDay As Single
Dim CurYear As Single
Dim LastDay As Single
Dim LastIndex As Single
Dim m_BackColor As OLE_COLOR
Dim m_Value As String
Dim CC As Boolean  ' used to keep track if calendar is visible or not
Dim m_DateDisplay As Boolean
Dim m_SHButton As Boolean
Dim StoreDate As String
Event Click()

Private Sub lblDateBar_Click()
  Dim ZX As Long
  Dim c As Control
  On Error Resume Next
  
     For ZX = 1 To UserControl.ParentControls.Count
       Set c = UserControl.ParentControls.Item(ZX)
       c.ZOrder 1
     Next
   If Label2.Caption = "Show" Then
      AnimateForm Picture1, aload, eAppearFromTop, 11, 33
      Label2.Caption = "Hide"
      CC = False
   Else
      AnimateForm Picture1, aUnload, eAppearFromBottom, 11, 33
      Label2.Caption = "Show"
      CC = True
   End If
   If UserControl.Height >= 3280 Then UserControl.Height = 3270

End Sub

Private Sub UserControl_Initialize()
  
   cboMonth.top = -300
   cboYear.top = -300
   cmdCurrentDate.top = 50
   FindYear CStr(CurYear)
   FillMonths
   Call ExpandCombo(cboMonth, HeightOnly, NoExpand)
   FillYears
   Call ExpandCombo(cboYear, HeightOnly, NoExpand)
   InitDates
   DisplayDates
   HighLightDate
   CC = True
   BackColor = m_def_BackColor
   DateDisplay = m_def_DateDisplay
   SHButton = m_def_SHButton
   
End Sub

Private Sub UserControl_Show()
  
   UserControl.BackColor = Ambient.BackColor
   UserControl.Height = lblDateBar.Height + 50
   BackColor = m_BackColor
   cmdCurrentDate_Click
   StoreDate = Value
   AnimateForm Picture1, aUnload, eAppearFromBottom, 11, 33
   Label2.Caption = "Show"

End Sub

Private Sub UserControl_Resize()
   UserControl.Width = 3800
   
   Picture1.Width = UserControl.Width - 25
End Sub

Private Sub Usercontrol_ReadProperties(Propbag As PropertyBag)
   m_Value = Propbag.ReadProperty("Value", StoreDate)
   m_BackColor = Propbag.ReadProperty("BackColor", m_def_BackColor)
   m_DateDisplay = Propbag.ReadProperty("DateDisplay", m_def_DateDisplay)
   m_SHButton = Propbag.ReadProperty("SHButton", m_def_SHButton)
   SHButton = m_SHButton
End Sub

Private Sub Usercontrol_WriteProperties(Propbag As PropertyBag)
   Call Propbag.WriteProperty("Value", m_Value, StoreDate)
   Call Propbag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
   Call Propbag.WriteProperty("DateDisplay", m_DateDisplay, m_def_DateDisplay)
   Call Propbag.WriteProperty("SHButton", m_SHButton, m_def_SHButton)
End Sub

Private Sub cmdCurrentDate_Click()
    InitDates
    
    If DateDisplay = True Then
        lblDateBar.Caption = Value
    Else
        lblDateBar.Caption = "Calendar"
    End If
    
    InitDates
    StoreDate = Value
    RaiseEvent Click
End Sub

Private Sub Label2_Click()
  Dim ZX As Long
  Dim c As Control
  On Error Resume Next
  
     For ZX = 1 To UserControl.ParentControls.Count
       Set c = UserControl.ParentControls.Item(ZX)
       c.ZOrder 1
     Next
   If CC = True Then
      AnimateForm Picture1, aload, eAppearFromTop, 11, 33
      Label2.Caption = "Hide"
      CC = False
   Else
      AnimateForm Picture1, aUnload, eAppearFromBottom, 11, 33
      Label2.Caption = "Show"
      CC = True
   End If
End Sub

Private Sub Picture1_Click()
   AnimateForm Picture1, aUnload, eAppearFromBottom, 11, 33
   Label2.Caption = "Show"
   CC = True
End Sub

Private Sub cboMonth_Click()
   DisplayDates
   If lblDates(LastIndex).Visible = False Then
      If lblDates(LastIndex).Index > 7 Then
         CurDay = lblDates(LastDay).Caption
      End If
   End If
   HighLightDate
   RaiseEvent Click
End Sub

Private Sub cboYear_Click()
   DisplayDates
   If lblDates(LastIndex).Visible = False Then
      If lblDates(LastIndex).Index > 7 Then
         CurDay = lblDates(LastDay).Caption
      End If
   End If
   HighLightDate
   RaiseEvent Click
End Sub

Private Sub DisplayDates()
   Dim iRow As Single
   Dim iColumn As Single
   Dim iDates As Single
   Dim CellTop As Single
   Dim CellLeft As Single
   
   CellTop = lblDates(0).top
   CellLeft = lblDates(0).Left
   
   For iRow = 1 To 6
      For iColumn = 1 To 7
         If iDates = 38 Then
            ShowDates
            Exit Sub
         End If
         On Error Resume Next
         iDates = iDates + 1
         Load lblDates(iDates)
         lblDates(iDates).Move CellLeft, CellTop
         CellLeft = CellLeft + lblDates(0).Width + 100
         lblDates(iDates).Visible = True
         
         Next
         CellTop = CellTop + lblDates(0).Height + 50
         CellLeft = lblDates(0).Left
         Next
      End Sub

Private Sub ExpandCombo(ByRef Combo As ComboBox, ByVal ExpandType As eExpandType, ByVal ExpandBy As eExpandBy, Optional ByVal hFrame As Long)
   
   Dim lRet As Long
   Dim pt As POINTAPI
   Dim rc As RECT
   Dim lComboWidth As Long
   Dim lNewHeight As Long
   Dim lItemHeight As Long
   
   If ExpandType <> HeightOnly Then
      lComboWidth = (Combo.Width / Screen.TwipsPerPixelX)
      Select Case ExpandBy
         Case 0
            lComboWidth = lComboWidth + (lComboWidth * 0.5)
         Case 1
            lComboWidth = lComboWidth + (lComboWidth * 0.75)
         Case 2
            lComboWidth = lComboWidth * 2
         Case 3
            lComboWidth = lComboWidth * 3
         Case 4
            lComboWidth = lComboWidth * 4
      End Select
      lRet = SendMessageByNum(Combo.hWnd, CB_SETDROPPEDCONTROLRECT, lComboWidth, 0)
      
   End If
   
   If ExpandType <> WidthOnly Then
      lComboWidth = Combo.Width / Screen.TwipsPerPixelX
      lItemHeight = SendMessageByNum(Combo.hWnd, CB_GETITEMHEIGHT, 0, 0)
      lNewHeight = lItemHeight * 30
      Call GetWindowRect(Combo.hWnd, rc)
      pt.X = rc.Left
      pt.Y = rc.top
      Call ScreenToClient(hFrame, pt)
      Call MoveWindow(Combo.hWnd, pt.X, pt.Y, lComboWidth, lNewHeight, True)
   End If
   
End Sub

Private Sub FillMonths()
   With cboMonth
      .AddItem "January"
      .AddItem "February"
      .AddItem "March"
      .AddItem "April"
      .AddItem "May"
      .AddItem "June"
      .AddItem "July"
      .AddItem "August"
      .AddItem "September"
      .AddItem "October"
      .AddItem "November"
      .AddItem "December"
   End With
End Sub

Private Sub FillYears()
   Dim iYear As Long
   
   For iYear = Year(Now) - 10 To Year(Now) + 10
      cboYear.AddItem iYear
      Next
      
   End Sub

Private Sub FindYear(Years As String)
   Dim ctr As Integer
   
   With cboYear
      For ctr = 0 To .ListCount - 1
         If .List(ctr) = Years Then
            .ListIndex = ctr
            Exit For
         End If
         Next
      End With
   End Sub

Private Sub HighLightDate()
   Dim X As Single
   
   For X = 1 To 38
      If CurDay = lblDates(X).Caption Then
         lblDates(X).BackStyle = 1
         lblDates(X).BorderStyle = 1
         LastIndex = lblDates(X).Index
      Else
         lblDates(X).BorderStyle = 0
         lblDates(X).BackStyle = 0
         lblDates(X).FontItalic = False
      End If
      
      Next
         'this is where you can format the date ,in label, to your needs.
         StoreDate = Format(cboMonth.Text & "/" & CurDay & "/" & cboYear.Text, "dddd - mmmm dd, yyyy")
         Value = StoreDate
         If DateDisplay = True Then
            lblDateBar.Caption = Value
         Else
            lblDateBar.Caption = "Calendar"
        End If
        
   End Sub

Private Sub InitDates()
   CurMonth = Month(Now)
   CurYear = Year(Now)
   CurDay = Day(Now)
   
   cboMonth.ListIndex = CurMonth - 1
   FindYear CStr(CurYear)
   HighLightDate
   Value = StoreDate
    If DateDisplay = True Then
       lblDateBar.Caption = Value
    Else
       lblDateBar.Caption = "Calendar"
    End If
End Sub

Private Sub lblDates_Click(Index As Integer)
   Dim tmpIndex As Integer
   
   CurDay = lblDates(Index).Caption
   HighLightDate
   
    If DateDisplay = True Then
       lblDateBar.Caption = Value
    Else
       lblDateBar.Caption = "Calendar"
    End If
    
   'close calendar when date is clicked,uncomment to activate
  ' AnimateForm Picture1, aUnload, eAppearFromBottom, 11, 33
   'Label2.Caption = "Show"
   'CC = True
   
   RaiseEvent Click
End Sub

Private Sub ShowDates()
   Dim StartDay As Single
   Dim ctr As Single
   Dim CheckDates As String
   Dim DateCaption As Single
   
   On Error Resume Next
   StartDay = Weekday(cboMonth.Text & "/1/" & cboYear.Text)
   
   For ctr = 0 To StartDay - 1
      lblDates(ctr).Visible = False
      Next
      
      For ctr = StartDay To 38
         DateCaption = DateCaption + 1
         CheckDates = Format(cboMonth & "/" & DateCaption & "/" & cboYear.Text, "Short Date")
         If Not IsDate(CheckDates) Then
            LastDay = lblDates(ctr - 1).Index
            Exit For
         End If
         
         If Weekday(CheckDates) = 1 Then
            lblDates(ctr).ForeColor = &HFF&
         End If
         
         If Weekday(CheckDates) = 7 Then
            lblDates(ctr).ForeColor = &HC00000
         End If
         lblDates(ctr).Caption = DateCaption
         
         Next
         
         For ctr = DateCaption + StartDay - 1 To 38
            lblDates(ctr).Visible = False
            Next
            
         End Sub

Public Property Get Value() As String
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As String)
   m_Value = New_Value
   PropertyChanged "Value"
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_BackColor = New_BackColor
   lblDateBar.BackColor = m_BackColor
   Label2.BackColor = m_BackColor
   Label1.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get DateDisplay() As Boolean
  DateDisplay = m_DateDisplay
End Property

Public Property Let DateDisplay(ByVal New_DateDisplay As Boolean)
   m_DateDisplay = New_DateDisplay
    If DateDisplay = True Then
        lblDateBar.Caption = Value
    Else
        lblDateBar.Caption = "Calendar"
    End If
   PropertyChanged "DateDisplay"
End Property

Public Property Get SHButton() As Boolean
   SHButton = m_SHButton
End Property

Public Property Let SHButton(ByVal New_SHButton As Boolean)
   m_SHButton = New_SHButton
   If SHButton = True Then
      Label2.Visible = True
      lblDateBar.Width = 3015
   Else
      Label2.Visible = False
      lblDateBar.Width = 3765
   End If
   PropertyChanged "SHButton"
End Property

'This Function code was written by Jim Jose

Private Function AnimateForm(hwndObject As Object, ByVal aEvent As AnimeEventEnum, _
   Optional ByVal aEffect As AnimeEffectEnum = 11, _
   Optional ByVal FrameTime As Long = 1, _
   Optional ByVal FrameCount As Long = 33) As Boolean
   On Error GoTo Handle
   Dim X1 As Long, Y1 As Long
   Dim hrgn As Long, tmpRgn As Long
   Dim XValue As Long, YValue As Long
   Dim XIncr As Double, YIncr As Double
   Dim wHeight As Long, wWidth As Long
   
   wWidth = hwndObject.Width / Screen.TwipsPerPixelX
   wHeight = hwndObject.Height / Screen.TwipsPerPixelY
   hwndObject.Visible = True
   
   Select Case aEffect
         
      Case eAppearFromTop
         
         YIncr = (wHeight / FrameCount)
         For Y1 = 0 To FrameCount
            
            ' Define the size of current frame/Create it
            YValue = Y1 * YIncr
            hrgn = CreateRectRgn(0, 0, wWidth, YValue)
            UserControl.Height = 3300  'UserControl.Height + YValue
            
            ' If unload then take the reverse(anti) region
            If aEvent = aUnload Then
               tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
               CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
               DeleteObject tmpRgn
            End If
            
            ' Set the new region for the window
            SetWindowRgn hwndObject.hWnd, hrgn, True:   DoEvents
            Sleep FrameTime
          
         Next Y1
         
      Case eAppearFromBottom
         
         YIncr = wHeight / FrameCount
         For Y1 = 0 To FrameCount
            
            ' Define the size of current frame/Create it
            YValue = wHeight - Y1 * YIncr
            hrgn = CreateRectRgn(0, YValue, wWidth, wHeight)
            If UserControl.Height <= 330 Then GoTo Here
            UserControl.Height = UserControl.Height - YValue
Here:
            ' If unload then take the reverse(anti) region
            If aEvent = aUnload Then
               tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
               CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
               DeleteObject tmpRgn
            End If
            
            ' Set the new region for the window
            SetWindowRgn hwndObject.hWnd, hrgn, True: DoEvents
            Sleep FrameTime
           
         Next Y1
         
   End Select
   
   AnimateForm = True
   
   Exit Function
Handle:
   AnimateForm = False
End Function
