VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin Project1.GradButton GradButton1 
      Height          =   375
      Left            =   2805
      TabIndex        =   6
      Top             =   1665
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Caption         =   "EXIT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   255
      ClkColor        =   33023
      ClkForeColor    =   8421504
      FontposX        =   32
   End
   Begin Project1.ucTextbox ucTextbox1 
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   2895
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   609
      Caption         =   " Name"
      Text            =   "Hello World"
      PointerRight    =   -1  'True
      BoldText        =   -1  'True
   End
   Begin Project1.ucSlider ucSlider1 
      Height          =   855
      Left            =   2460
      TabIndex        =   4
      Top             =   585
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1508
      Caption         =   "ucSlider1"
      Value           =   75
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarColor        =   5
   End
   Begin Project1.msDial msDial1 
      Height          =   1005
      Left            =   210
      TabIndex        =   0
      Top             =   555
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1773
      Caption         =   "Volume"
   End
   Begin Project1.ThumbWheel TW1 
      Height          =   420
      Left            =   1320
      TabIndex        =   3
      Top             =   585
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   741
      CustomBackColor =   0
      CustomBorder    =   0   'False
   End
   Begin Project1.LEDDisplaySTO LED1 
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   1095
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      DigitCount      =   3
      LeadingZeros    =   -1  'True
      BorderColor     =   0
   End
   Begin Project1.CkBx CkBx1 
      Height          =   270
      Left            =   1425
      TabIndex        =   1
      Top             =   1770
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   476
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.StrokeText StrokeText1 
      Height          =   525
      Left            =   240
      TabIndex        =   7
      Top             =   2190
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   926
      Caption         =   "A Small Test"
      Shadow          =   -1  'True
      TStyle          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ucUpDown ucUpDown1 
      Height          =   240
      Left            =   225
      TabIndex        =   8
      Top             =   1815
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   423
      Min             =   -100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TW1_Change()
   LED1.Value = TW1.Value
End Sub

