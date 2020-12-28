VERSION 4.00
Begin VB.Form ProgrammingEasyScreenSpeedTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoamtic Train Control - Programming Mode Easy Screen - Speed Table"
   ClientHeight    =   5145
   ClientLeft      =   675
   ClientTop       =   2025
   ClientWidth     =   7065
   FillStyle       =   0  'Solid
   Height          =   5550
   Icon            =   "ProgrammingEasyScreenSpeedTable.frx":0000
   Left            =   615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7065
   Top             =   1680
   Width           =   7185
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4440
      TabIndex        =   59
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   94
      Left            =   6720
      TabIndex        =   58
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   93
      Left            =   6480
      TabIndex        =   57
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   92
      Left            =   6240
      TabIndex        =   56
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   91
      Left            =   6000
      TabIndex        =   55
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   90
      Left            =   5760
      TabIndex        =   54
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   89
      Left            =   5520
      TabIndex        =   53
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   88
      Left            =   5280
      TabIndex        =   52
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   87
      Left            =   5040
      TabIndex        =   51
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   86
      Left            =   4800
      TabIndex        =   50
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   85
      Left            =   4560
      TabIndex        =   49
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   84
      Left            =   4320
      TabIndex        =   48
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   83
      Left            =   4080
      TabIndex        =   47
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   82
      Left            =   3840
      TabIndex        =   46
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   81
      Left            =   3600
      TabIndex        =   45
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   80
      Left            =   3360
      TabIndex        =   44
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   79
      Left            =   3120
      TabIndex        =   43
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   78
      Left            =   2880
      TabIndex        =   42
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   77
      Left            =   2640
      TabIndex        =   41
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   76
      Left            =   2400
      TabIndex        =   40
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   75
      Left            =   2160
      TabIndex        =   39
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   74
      Left            =   1920
      TabIndex        =   38
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   73
      Left            =   1680
      TabIndex        =   37
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   72
      Left            =   1440
      TabIndex        =   36
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   71
      Left            =   1200
      TabIndex        =   35
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   70
      Left            =   960
      TabIndex        =   34
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   69
      Left            =   720
      TabIndex        =   33
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   68
      Left            =   480
      TabIndex        =   32
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox TextboxSpeedTableCV 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   67
      Left            =   240
      TabIndex        =   31
      Text            =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "ProgrammingEasyScreenSpeedTable.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7440
      Top             =   180
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7440
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7440
      Top             =   720
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   68
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   67
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   69
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   70
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   71
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   72
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   73
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   74
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   75
      Left            =   2160
      TabIndex        =   9
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   76
      Left            =   2400
      TabIndex        =   10
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   77
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   78
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   79
      Left            =   3120
      TabIndex        =   13
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   80
      Left            =   3360
      TabIndex        =   14
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   81
      Left            =   3600
      TabIndex        =   15
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   82
      Left            =   3840
      TabIndex        =   16
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   83
      Left            =   4080
      TabIndex        =   17
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   84
      Left            =   4320
      TabIndex        =   18
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   85
      Left            =   4560
      TabIndex        =   19
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   86
      Left            =   4800
      TabIndex        =   20
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   87
      Left            =   5040
      TabIndex        =   21
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   88
      Left            =   5280
      TabIndex        =   22
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   89
      Left            =   5520
      TabIndex        =   23
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   90
      Left            =   5760
      TabIndex        =   24
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   91
      Left            =   6000
      TabIndex        =   25
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   92
      Left            =   6240
      TabIndex        =   26
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   93
      Left            =   6480
      TabIndex        =   27
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VslideLib.VSlider SliderSpeedTableCV 
      Height          =   3495
      Index           =   94
      Left            =   6720
      TabIndex        =   28
      Top             =   720
      Width           =   255
      _Version        =   65537
      _ExtentX        =   450
      _ExtentY        =   6165
      _StockProps     =   73
      ForeColor       =   8421376
      BevelInner      =   0
      BevelOuter      =   0
      BevelWidth      =   2
      BorderWidth     =   5
      Gap             =   10
      LargeChange     =   10
      ThumbHeight     =   120
      ThumbStyle      =   0
      ThumbWidth      =   195
      TickColor       =   0
      TickCount       =   0
      TickLength      =   0
      TickMarks       =   0
      TickWidth       =   0
      TrackBevel      =   3
      TrackWidth      =   5
      Value           =   0
      Max             =   0
      Min             =   255
   End
   Begin VB.Label Label1 
      Caption         =   $"ProgrammingEasyScreenSpeedTable.frx":0884
      Height          =   615
      Left            =   720
      TabIndex        =   30
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "ProgrammingEasyScreenSpeedTable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub Command1_Click()

    ProgrammingEasyScreenSpeedTable.PrintForm
    
End Sub


Private Sub ButtonClose_Click()

For t = 67 To 94
    Let ProgrammingSpeedTable.LocomotiveDecoderCVd(t).Text = SliderSpeedTableCV(t).Value
Next t
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Remove from Screen Stack
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Screen Stack"
    Dim TemporaryScreen As String
    Dim TemporaryCounter As Integer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Find Current Screen and Hide
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryCounter = 9 To 0 Step -1
        Let Ini.Parameter = CStr(TemporaryCounter)
        Let TemporaryScreen = Ini.Value
        If TemporaryScreen <> "Unused" Then
            If TemporaryScreen = "Programming Easy Screen Speed Table Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Speed Table Screen, Button Close, current window is not listed in the stack to remove it and hide."
            End If
 
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Open Previous Screen
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Screen Stack"
            Let Ini.Parameter = CStr(TemporaryCounter - 1)
            Let TemporaryScreen = Ini.Value
            If TemporaryScreen = "About Screen" Then
                About.Show vbModeless
            ElseIf TemporaryScreen = "Clock Screen" Then
                ClockScreen.Show vbModeless
            ElseIf TemporaryScreen = "Communication Setting Screen " Then
                CommunicationSetting.Show vbModeless
            ElseIf TemporaryScreen = "Define Block Properties Screen" Then
                DefineBlockProperties.Show vbModeless
            ElseIf TemporaryScreen = "Define Blocks Screen" Then
                DefineBlocks.Show vbModeless
            ElseIf TemporaryScreen = "Define Blocks Spreadsheet Screen" Then
                DefineBlocksSpreadsheet.Show vbModeless
            'ElseIf TemporaryScreen = "Fun Screen" Then
            '    FunScreen.Show vbModeless
            ElseIf TemporaryScreen = "Internet Settings Screen" Then
                InternetSettings.Show vbModeless
            ElseIf TemporaryScreen = "Locomotive CV Spreadsheet Screen" Then
                LocomotiveCVSpreadsheet.Show vbModeless
            ElseIf TemporaryScreen = "Locomotive Spreadsheet Screen" Then
                LocomotiveSpreadsheet.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Consist Screen" Then
                MainlineConsist.Show vbModeless
            ElseIf TemporaryScreen = "Mainline CV Changer Screen" Then
                MainlineCVChanger.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Decoder Screen" Then
                MainlineDecoder.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Diesel Screen" Then
                MainlineDiesel.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Easy Screen Configuration Screen" Then
                MainlineEasyScreenConfiguration.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Easy Screen Consist Functions Screen" Then
                MainlineEasyScreenConsistFunctions.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Easy Screen Functions Screen" Then
                MainlineEasyScreenFunctions.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Easy Screen Specific CVs Screen" Then
                MainlineEasyScreenSpecificCvs.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Easy Screen Speed Table Screen" Then
                MainlineEasyScreenSpeedTable.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Macro Maker Screen" Then
                MainlineMacroMaker.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation ATC Screen" Then
                MainlineOperationATC.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Screen" Then
                MainlineOperationGUI.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel1 Screen" Then
                MainlineOperationGuiDiesel1Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel2 Screen" Then
                MainlineOperationGuiDiesel2Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel3 Screen" Then
                MainlineOperationGuiDiesel3Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel4 Screen" Then
                MainlineOperationGuiDiesel4Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Electric1 Screen" Then
                MainlineOperationGuiElectric1Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Operation GUI Steam1 Screen" Then
                MainlineOperationGuiSteam1Screen.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Other Screen" Then
                MainlineOther.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Prototype Info Screen" Then
                MainlinePrototypeInfo.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Rolling Stock Screen" Then
                MainlineRollingStock.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Scale Speed Operation Screen" Then
                MainlineScaleSpeedOperation.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Scale Speed Setting Screen" Then
                MainlineScaleSpeedSetting.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Speed Table Screen" Then
                MainlineSpeedTable.Show vbModeless
            ElseIf TemporaryScreen = "Mainline Steam Screen" Then
                MainlineSteam.Show vbModeless
            ElseIf TemporaryScreen = "Main Screen" Then
                MainScreen.Show vbModeless
            ElseIf TemporaryScreen = "Opening Screen" Then
                OpeningScreen.Show vbModeless
            ElseIf TemporaryScreen = "Password Screen" Then
                Password.Show vbModeless
            ElseIf TemporaryScreen = "Programming Decoder Screen" Then
                ProgrammingDecoder.Show vbModeless
            ElseIf TemporaryScreen = "Programming Diesel Screen" Then
                ProgrammingDiesel.Show vbModeless
            ElseIf TemporaryScreen = "Programming Easy Screen Configuration Screen" Then
                ProgrammingEasyScreenConfiguration.Show vbModeless
            ElseIf TemporaryScreen = "Programming Easy Screen Consist Functions Screen" Then
                ProgrammingEasyScreenConsistFunctions.Show vbModeless
            ElseIf TemporaryScreen = "Programming Easy Screen Functions Screen" Then
                ProgrammingEasyScreenFunctions.Show vbModeless
            ElseIf TemporaryScreen = "Programming Easy Screen Specific CVs Screen" Then
                ProgrammingEasyScreenSpecificCvs.Show vbModeless
            ElseIf TemporaryScreen = "Programming Easy Screen Speed Table Screen" Then
                ProgrammingEasyScreenSpeedTable.Show vbModeless
            ElseIf TemporaryScreen = "Programming Other Screen" Then
                ProgrammingOther.Show vbModeless
            ElseIf TemporaryScreen = "Programming Prototype Info Screen" Then
                ProgrammingPrototypeInfo.Show vbModeless
            ElseIf TemporaryScreen = "Programming Rolling Stock Screen" Then
                ProgrammingRollingStock.Show vbModeless
            ElseIf TemporaryScreen = "Programming Speed Table Screen" Then
                ProgrammingSpeedTable.Show vbModeless
            ElseIf TemporaryScreen = "Programming Steam Screen" Then
                ProgrammingSteam.Show vbModeless
            ElseIf TemporaryScreen = "Room Lighting Control Screen" Then
                RoomLightingControl.Show vbModeless
            ElseIf TemporaryScreen = "Screen Attribute Setting Screen" Then
                ScreenAttributeSetting.Show vbModeless
            ElseIf TemporaryScreen = "Sound Screen" Then
                SoundScreen.Show vbModeless
            ElseIf TemporaryScreen = "Sound Screen Edit List Screen" Then
                SoundScreenEditList.Show vbModeless
            ElseIf TemporaryScreen = "System Information Screen" Then
                SystemInformation.Show vbModeless
            ElseIf TemporaryScreen = "Update Software Screen" Then
                UpdateSoftware.Show vbModeless
            ElseIf TemporaryScreen = "Utilities for Command Control" Then
                UtilitiesForCommandControl.Show vbModeless
            ElseIf TemporaryScreen = "Video Settings Screen" Then
                VideoSettings.Show vbModeless
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Speed Table Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
            End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Loop Premature
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TemporaryCounter = -2
        End If
    Next TemporaryCounter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Screen Stack is Empty
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryCounter = -1 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Programming Easy Screen Speed Table Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   End Sub





Private Sub Form_Activate()

    DoEvents
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Add to Screen Stack
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Screen Stack"
    Dim TemporaryScreen As String
    Dim TemporaryCounter As Integer
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Lop for Checking Sceen Stack
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryCounter = 0 To 9
        Let Ini.Parameter = CStr(TemporaryCounter)
        Let TemporaryScreen = Ini.Value
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Already Present in INI
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If TemporaryScreen = "Programming Easy Screen Speed Table Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Programming Easy Screen Speed Table Screen"
            Let TemporaryCounter = 11
        End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Check Next Item in Stack
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Next TemporaryCounter
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Screen Stack is Full
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryCounter = 10 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Programming Easy Screen Speed Table Screen, Form Activate, stack is full, overflow."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transparency Screen Delay
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "BackgroundImage"
    Dim TemporaryBackgroundImage As String
    Let TemporaryBackgroundImage = Ini.Value
    If TemporaryBackgroundImage = "On" Then
        Let Ini.Parameter = "Transparency"
        Dim TemporaryTransparency As String
        Let TemporaryTransparency = Ini.Value
        If TemporaryTransparency = "On" Then
            Let AlphaBlend.Enabled = True
            Let Ini.Parameter = "Opacity"
            Dim TemporaryOpacity As String
            Let TemporaryOpacity = Ini.Value
            Dim TemporaryScreenDelay As String
            Let TemporaryScreenDelay = Ini.Value
            Dim OutsideLoop As Integer
            Dim InsideLoop As Integer
            For OutsideLoop = 0 To Val(TemporaryOpacity)
                Let AlphaBlend.Opacity = OutsideLoop
                'For InsideLoop = 0 To Val(TemporaryScreenDelay)
                '    DoEvents
                'Next InsideLoop
            Next OutsideLoop
        ElseIf TemporaryTransparency = "Off" Then
           Let AlphaBlend.Enabled = False
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Speed Table Screen, Form Activate, variable error in ATC.INI file for 'Transparncy' setting."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Programming Easy Screen Speed Table Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Programming Easy screen Speed Table Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenSpeedTable.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenSpeedTable.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenSpeedTable.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenSpeedTable.Height)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transparency Screen Delay
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "BackgroundImage"
    Dim TemporaryBackgroundImage As String
    Let TemporaryBackgroundImage = Ini.Value
    If TemporaryBackgroundImage = "On" Then
        Let Ini.Parameter = "Transparency"
        Dim TemporaryTransparency As String
        Let TemporaryTransparency = Ini.Value
        If TemporaryTransparency = "On" Then
            Let AlphaBlend.Enabled = True
            Let Ini.Parameter = "Opacity"
            Dim TemporaryOpacity As String
            Let TemporaryOpacity = Ini.Value
            Dim TemporaryScreenDelay As String
            Let TemporaryScreenDelay = Ini.Value
            Dim OutsideLoop As Integer
            Dim InsideLoop As Integer
            For OutsideLoop = Val(TemporaryOpacity) To 0 Step -1
                Let AlphaBlend.Opacity = OutsideLoop
                'For InsideLoop = Val(TemporaryScreenDelay) To 0 Step -1
                '    DoEvents
                'Next InsideLoop
            Next OutsideLoop
        ElseIf TemporaryTransparency = "Off" Then
           Let AlphaBlend.Enabled = False
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Speed Table Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Programming Easy Screen Speed Table Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ProgrammingEasyScreenSpeedTable.Hide
    'unload ProgrammingEasyScreenspeedtable

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub


Private Sub Form_Load()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While Screen.Width < Width Or Screen.Height < Height
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & vbCrLf & "' which requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen" & vbCrLf & "resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbRetry Then
            Load ScreenAttributeSetting
            ScreenAttributeSetting.Show vbModeless '  And Again
        ElseIf TemporaryResponse = vbCancel Then
            End
        End If
    Loop
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialization of Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Programming Easy Screen Speed Table Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        ProgrammingEasyScreenSpeedTable.Left = (Screen.Width - Width) / 2
        ProgrammingEasyScreenSpeedTable.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + ProgrammingEasyScreenSpeedTable.Width > Screen.Width Then
            Let ProgrammingEasyScreenSpeedTable.Left = Screen.Width - ProgrammingEasyScreenSpeedTable.Width
        Else
            Let ProgrammingEasyScreenSpeedTable.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + ProgrammingEasyScreenSpeedTable.Height > Screen.Height Then
            Let ProgrammingEasyScreenSpeedTable.Top = Screen.Height - ProgrammingEasyScreenSpeedTable.Height
        Else
            Let ProgrammingEasyScreenSpeedTable.Top = Val(TemporaryValueTop)
        End If
    End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim BalloonFont As New StdFont
         
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
            Ini.Parameter = "BalloonHelpFontName"
            Ini.Value = BalloonFont.Name
            Ini.Parameter = "BalloonHelpFontSize"
            Ini.Value = BalloonFont.Size
            Ini.Parameter = "BalloonHelpColour1"
            Colour1 = Ini.Value
            Ini.Parameter = "BalloonHelpColour2"
            Colour2 = Ini.Value
            Ini.Parameter = "BalloonHelpColour3"
            Colour3 = Ini.Value

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step one of twenty-eight, stored in configuration variable sixty-seven."
        Let TemporaryText2 = "Speed Step One"
        ' i= BalloonHelp.DestroyToolTip(SliderSpeedTableCV(67))
        ' i= BalloonHelp.AddToolTip(SliderSpeedTableCV(67), BalloonText1,balBalloon,BalloonHelpText2, balInfo, RGB(BalloonHelpColour1,BalloonHelpColour2,BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime,BalloonHelpShadow,BalloonHelpCenter,BalloonHelpShowOnDemand,BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 27,"")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step two of twenty-eight, stored in configuration variable sixty-eight."
        Let TemporaryText2 = "Speed Step Two"
        ' i= BalloonHelp.DestroyToolTip(SliderSpeedTableCV(68))
        ' i= BalloonHelp.AddToolTip(SliderSpeedTableCV(68), BalloonText1,balBalloon,BalloonHelpText2, balInfo, RGB(BalloonHelpColour1,BalloonHelpColour2,BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime,BalloonHelpShadow,BalloonHelpCenter,BalloonHelpShowOnDemand,BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 27,"")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step three of twenty-eight, stored in configuration variable sixty-nine."
        Let TemporaryText2 = "Speed Step Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(69))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(69), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step four of twenty-eight, stored in configuration variable seventy."
        Let TemporaryText2 = "Speed Step Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(70))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(70), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step five of twenty-eight, stored in configuration variable seventy-one."
        Let TemporaryText2 = "Speed Step Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(71))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(71), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step six of twenty-eight, stored in configuration variable seventy-two."
        Let TemporaryText2 = "Speed Step Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(72))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(72), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step seven of twenty-eight, stored in configuration variable seventy-three."
        Let TemporaryText2 = "Speed Step Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(73))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(73), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step eight of twenty-eight, stored in configuration variable seventy-four."
        Let TemporaryText2 = "Speed Step Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(74))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(74), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step nine of twenty-eight, stored in configuration variable seventy-five."
        Let TemporaryText2 = "Speed Step Nine"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(75))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(75), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step ten of twenty-eight, stored in configuration variable seventy-six."
        Let TemporaryText2 = "Speed Step Ten"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(76))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(76), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step eleven of twenty-eight, stored in configuration variable seventy-seven."
        Let TemporaryText2 = "Speed Step Eleven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(77))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(77), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twelve of twenty-eight, stored in configuration variable seventy-eight."
        Let TemporaryText2 = "Speed Step Twelve"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(78))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(78), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step thirteen of twenty-eight, stored in configuration variable seventy-nine."
        Let TemporaryText2 = "Speed Step Thirteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(79))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(79), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step fourteen of twenty-eight, stored in configuration variable eighty."
        Let TemporaryText2 = "Speed Step Fourteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(80))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(80), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step fifteen of twenty-eight, stored in configuration variable eighty-one."
        Let TemporaryText2 = "Speed Step Fifteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(81))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(81), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step sixteen of twenty-eight, stored in configuration variable eighty-two."
        Let TemporaryText2 = "Speed Step Sixteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(82))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(82), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step seventeen of twenty-eight, stored in configuration variable eighty-three."
        Let TemporaryText2 = "Speed Step Seventeen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(83))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(83), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step eighteen of twenty-eight, stored in configuration variable eighty-four."
        Let TemporaryText2 = "Speed Step Eighteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(84))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(84), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step nineteen of twenty-eight, stored in configuration variable eighty-five."
        Let TemporaryText2 = "Speed Step Nineteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(85))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(85), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty of twenty-eight, stored in configuration variable eighty-six."
        Let TemporaryText2 = "Speed Step Twenty"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(86))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(86), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-one of twenty-eight, stored in configuration variable eighty-seven."
        Let TemporaryText2 = "Speed Step Twenty-One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(87))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(87), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-two of twenty-eight, stored in configuration variable eighty-eight."
        Let TemporaryText2 = "Speed Step Twenty-Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(88))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(88), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-three of twenty-eight, stored in configuration variable eighty-nine."
        Let TemporaryText2 = "Speed Step Twenty-Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(89))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(89), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-four of twenty-eight, stored in configuration variable ninety."
        Let TemporaryText2 = "Speed Step Twenty-Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(90))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(90), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-five of twenty-eight, stored in configuration variable ninety-one."
        Let TemporaryText2 = "Speed Step Twenty-Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(91))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(91), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-six of twenty-eight, stored in configuration variable ninety-two."
        Let TemporaryText2 = "Speed Step Twenty-Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(92))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(92), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-seven of twenty-eight, stored in configuration variable ninety-three."
        Let TemporaryText2 = "Speed Step Twenty-Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(93))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(93), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These slide bars are used to graphically show the" & vbCrLf & "speed table. This slide bar will change the value for speed" & vbCrLf & "step twenty-eight of twenty-eight, stored in configuration variable ninety-four."
        Let TemporaryText2 = "Speed Step Twenty-Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SliderSpeedTableCV(94))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SliderSpeedTableCV(94), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step one of twenty-eight, stored in configuration variable sixty-seven."
        Let TemporaryText2 = "Speed Step One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(67))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(67), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step two of twenty-eight, stored in configuration variable sixty-eight."
        Let TemporaryText2 = "Speed Step Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(68))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(68), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step three of twenty-eight, stored in configuration variable sixty-nine."
        Let TemporaryText2 = "Speed Step Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(69))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(69), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step four of twenty-eight, stored in configuration variable seventy."
        Let TemporaryText2 = "Speed Step Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(70))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(70), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step five of twenty-eight, stored in configuration variable seventy-one."
        Let TemporaryText2 = "Speed Step Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(71))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(71), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step six of twenty-eight, stored in configuration variable seventy-two."
        Let TemporaryText2 = "Speed Step Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(72))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(72), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step seven of twenty-eight, stored in configuration variable seventy-three."
        Let TemporaryText2 = "Speed Step Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(73))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(73), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step eight of twenty-eight, stored in configuration variable seventy-four."
        Let TemporaryText2 = "Speed Step Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(74))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(74), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step nine of twenty-eight, stored in configuration variable seventy-five."
        Let TemporaryText2 = "Speed Step Nine"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(75))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(75), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step ten of twenty-eight, stored in configuration variable seventy-six."
        Let TemporaryText2 = "Speed Step Ten"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(76))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(76), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step eleven of twenty-eight, stored in configuration variable seventy-seven."
        Let TemporaryText2 = "Speed Step Eleven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(77))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(77), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twelve of twenty-eight, stored in configuration variable seventy-eight."
        Let TemporaryText2 = "Speed Step Twelve"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(78))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(78), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step thirteen of twenty-eight, stored in configuration variable seventy-nine."
        Let TemporaryText2 = "Speed Step Thirteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(79))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(79), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step fourteen of twenty-eight, stored in configuration variable eighty."
        Let TemporaryText2 = "Speed Step Fourteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(80))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(80), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step fifteen of twenty-eight, stored in configuration variable eighty-one."
        Let TemporaryText2 = "Speed Step Fifteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(81))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(81), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step sixteen of twenty-eight, stored in configuration variable eighty-two."
        Let TemporaryText2 = "Speed Step Sixteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(82))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(82), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step seventeen of twenty-eight, stored in configuration variable eighty-three."
        Let TemporaryText2 = "Speed Step Seventeen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(83))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(83), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step eighteen of twenty-eight, stored in configuration variable eighty-four."
        Let TemporaryText2 = "Speed Step Eighteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(84))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(84), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step nineteen of twenty-eight, stored in configuration variable eighty-five."
        Let TemporaryText2 = "Speed Step Nineteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(85))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(85), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty of twenty-eight, stored in configuration variable eighty-six."
        Let TemporaryText2 = "Speed Step Twenty"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(86))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(86), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-one of twenty-eight, stored in configuration variable eighty-seven."
        Let TemporaryText2 = "Speed Step Twenty-One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(87))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(87), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-two of twenty-eight, stored in configuration variable eighty-eight."
        Let TemporaryText2 = "Speed Step Twenty-Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(88))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(88), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-three of twenty-eight, stored in configuration variable eighty-nine."
        Let TemporaryText2 = "Speed Step Twenty-Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(89))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(89), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-four of twenty-eight, stored in configuration variable ninety."
        Let TemporaryText2 = "Speed Step Twenty-Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(90))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(90), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-five of twenty-eight, stored in configuration variable ninety-one."
        Let TemporaryText2 = "Speed Step Twenty-Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(91))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(91), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-six of twenty-eight, stored in configuration variable ninety-two."
        Let TemporaryText2 = "Speed Step Twenty-Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(92))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(92), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-seven of twenty-eight, stored in configuration variable ninety-three."
        Let TemporaryText2 = "Speed Step Twenty-Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(93))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(93), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will show the value for speed" & vbCrLf & "step twenty-eight of twenty-eight, stored in configuration variable ninety-four."
        Let TemporaryText2 = "Speed Step Twenty-Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxSpeedTableCV(94))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSpeedTableCV(94), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button closes this window and returns you" & vbCrLf & "to the 'Programming Mode - Speed Table'. All values from" & vbCrLf & "this window will be saved to the 'Programming Mode - Speed Table' screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Slider Controls
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim TemporaryVariableT As Integer
    
    For TemporaryVariableT = 67 To 94
        Let SliderSpeedTableCV(TemporaryVariableT).Value = Val(ProgrammingSpeedTable.LocomotiveDecoderCVd(TemporaryVariableT).Text)
    Next TemporaryVariableT

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
'
' Ends a procedure or block.
'
' Syntax is in the following format
'
'   End Sub
'
' End Sub Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions
' that apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    If ProgrammingEasyScreenSpeedTable.WindowState = vbMinimized Then
    
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BackgroundImage"
        'Dim TemporaryValue As String
        Let TemporaryValue = Ini.Value
    
        'Let BackGround.ImageBoxBackGround.Width = Screen.Width / 15
        'Let BackGround.ImageBoxBackGround.Height = Screen.Height / 15
    
        If TemporaryValue = "On" Then
            Let BackGround.WindowState = vbMinimized
        ElseIf TemporaryValue = "Off" Then
            Let BackGround.Visible = False
        'BackGround.ZOrder 1
        'BackGround.WindowState = 2
        ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Resize, variable not set correctly for 'BackGround Image' in ATC.INI file."
        End If
        
    ElseIf ProgrammingEasyScreenSpeedTable.WindowState = vbNormal Then
    
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BackgroundImage"
        'Dim TemporaryValue As String
        Let TemporaryValue = Ini.Value
    
        'Let BackGround.ImageBoxBackGround.Width = Screen.Width / 15
        'Let BackGround.ImageBoxBackGround.Height = Screen.Height / 15
    
        If TemporaryValue = "On" Then
            Let BackGround.WindowState = vbMaximized
            BackGround.ZOrder 1
        ElseIf TemporaryValue = "Off" Then
            Let BackGround.Visible = False
        ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Resize, variable not set correctly for 'BackGround Image' in ATC.INI file."
        End If
        
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unloading the Form
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Saving the screen size
'

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Programming Easy Screen Speed Table"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenSpeedTable.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenSpeedTable.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenSpeedTable.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenSpeedTable.Height)
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
'
' Ends a procedure or block.
'
' Syntax is in the following format
'
'   End Sub
'
' End Sub Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions
' that apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.
 
End Sub

Private Sub LabelSpeedTableCV_Click(Index As Integer)

End Sub

Private Sub SliderSpeedTableCV_Change(Index As Integer)

Let TextboxSpeedTableCV(Index).Text = SliderSpeedTableCV(Index).Value

End Sub


