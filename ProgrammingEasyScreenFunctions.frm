VERSION 4.00
Begin VB.Form ProgrammingEasyScreenFunctions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Easy Screen - Functions"
   ClientHeight    =   8250
   ClientLeft      =   1290
   ClientTop       =   3045
   ClientWidth     =   9000
   Height          =   8655
   Icon            =   "ProgrammingEasyScreenFunctions.frx":0000
   Left            =   1230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9000
   Top             =   2700
   Width           =   9120
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   6360
      TabIndex        =   124
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   42
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   41
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   40
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   39
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   38
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   37
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   36
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   35
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   34
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   73
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   74
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   75
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   76
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   77
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   78
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   79
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV42 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   80
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   65
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   66
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   67
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   68
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   69
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   70
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   71
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV41 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   72
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   57
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   58
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   59
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   60
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   61
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   62
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   63
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV40 
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   64
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   49
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   50
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   51
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   52
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   53
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   54
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   55
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV39 
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   56
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   41
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   42
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   43
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   44
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   45
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   46
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   47
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV38 
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   48
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   33
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   34
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   35
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   36
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   37
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   38
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   39
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV37 
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   40
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   25
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   26
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   27
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   28
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   29
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   30
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   31
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV36 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   32
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   17
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   18
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   19
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   20
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   21
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   22
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV35 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   24
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   9
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   11
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   12
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV34 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV33 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   33
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   9420
      Top             =   120
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   9360
      Top             =   1200
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   9360
      Top             =   600
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label LabelFunction8 
      Caption         =   "Function eight controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   123
      Top             =   7560
      Width           =   8655
   End
   Begin VB.Label LabelFunction7 
      Caption         =   "Function seven controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   122
      Top             =   7320
      Width           =   8655
   End
   Begin VB.Label LabelFunction6 
      Caption         =   "Function six controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   121
      Top             =   7080
      Width           =   8655
   End
   Begin VB.Label LabelFunction5 
      Caption         =   "Function five controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   120
      Top             =   6840
      Width           =   8655
   End
   Begin VB.Label LabelFunction4 
      Caption         =   "Function four controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   119
      Top             =   6600
      Width           =   8655
   End
   Begin VB.Label LabelFunction3 
      Caption         =   "Function three controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   118
      Top             =   6360
      Width           =   8655
   End
   Begin VB.Label LabelFunction2 
      Caption         =   "Function two controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   117
      Top             =   6120
      Width           =   8655
   End
   Begin VB.Label Label24 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   240
      TabIndex        =   116
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label LabelFunction1 
      Caption         =   "Function one controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   115
      Top             =   5880
      Width           =   8655
   End
   Begin VB.Label LabelFunctionLR 
      Caption         =   "Function zero (yellow wire) controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   114
      Top             =   5640
      Width           =   8655
   End
   Begin VB.Label LabelFunctionLF 
      Caption         =   "Function zero (white wire) controls no outputs."
      Height          =   255
      Left            =   240
      TabIndex        =   113
      Top             =   5400
      Width           =   8655
   End
   Begin VB.Line Line2 
      X1              =   8880
      X2              =   240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "CV42"
      Height          =   195
      Left            =   7200
      TabIndex        =   112
      Top             =   4680
      Width           =   390
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "CV41"
      Height          =   195
      Left            =   7200
      TabIndex        =   111
      Top             =   4320
      Width           =   390
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "CV40"
      Height          =   195
      Left            =   7200
      TabIndex        =   110
      Top             =   3960
      Width           =   390
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "CV39"
      Height          =   195
      Left            =   7200
      TabIndex        =   109
      Top             =   3600
      Width           =   390
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "CV38"
      Height          =   195
      Left            =   7200
      TabIndex        =   108
      Top             =   3240
      Width           =   390
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F8"
      Height          =   195
      Left            =   360
      TabIndex        =   107
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F7"
      Height          =   195
      Left            =   360
      TabIndex        =   106
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F6"
      Height          =   195
      Left            =   360
      TabIndex        =   105
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F5"
      Height          =   195
      Left            =   360
      TabIndex        =   104
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F4"
      Height          =   195
      Left            =   360
      TabIndex        =   103
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F3"
      Height          =   195
      Left            =   360
      TabIndex        =   102
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F2"
      Height          =   195
      Left            =   360
      TabIndex        =   101
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F1"
      Height          =   195
      Left            =   360
      TabIndex        =   100
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F0 Rear"
      Height          =   195
      Left            =   360
      TabIndex        =   99
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Function F0 Front"
      Height          =   195
      Left            =   315
      TabIndex        =   98
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "CV37"
      Height          =   195
      Left            =   7185
      TabIndex        =   97
      Top             =   2880
      Width           =   390
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "CV36"
      Height          =   195
      Left            =   7185
      TabIndex        =   96
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "CV35"
      Height          =   195
      Left            =   7200
      TabIndex        =   95
      Top             =   2160
      Width           =   390
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "CV34"
      Height          =   195
      Left            =   7185
      TabIndex        =   94
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CV33"
      Height          =   195
      Left            =   7185
      TabIndex        =   93
      Top             =   1440
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   8880
      X2              =   360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Output of decoder    14    13    12    11     10     9       8      7      6      5      4      3      2      1"
      Height          =   195
      Left            =   360
      TabIndex        =   92
      Top             =   960
      Width           =   6375
   End
   Begin VB.Image ImageIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ProgrammingEasyScreenFunctions.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"ProgrammingEasyScreenFunctions.frx":0884
      Height          =   495
      Left            =   840
      TabIndex        =   91
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "ProgrammingEasyScreenFunctions"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub ButtonClose_Click()

Let ProgrammingDecoder!LocomotiveDecoderCVd(33) = textboxcvvalue(33).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(34) = textboxcvvalue(34).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(35) = textboxcvvalue(35).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(36) = textboxcvvalue(36).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(37) = textboxcvvalue(37).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(38) = textboxcvvalue(38).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(39) = textboxcvvalue(39).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(40) = textboxcvvalue(40).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(41) = textboxcvvalue(41).Text
Let ProgrammingDecoder!LocomotiveDecoderCVd(42) = textboxcvvalue(42).Text
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
            If TemporaryScreen = "Programming Easy Screen Functions Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Functions Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Programming Easy Screen Functions Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Programming Easy Screen Functions Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonPrint_Click()
    
    ProgrammingEasyScreenFunctions.PrintForm
    
End Sub


Private Sub CheckBoxCV33_Click(Index As Integer)

If CheckBoxCV33(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 1
    If Index = 2 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 2
    If Index = 3 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 4
    If Index = 4 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 8
    If Index = 5 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 16
    If Index = 6 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 32
    If Index = 7 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 64
    If Index = 8 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) - 128
End If

If CheckBoxCV33(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 1
    If Index = 2 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 2
    If Index = 3 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 4
    If Index = 4 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 8
    If Index = 5 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 16
    If Index = 6 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 32
    If Index = 7 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 64
    If Index = 8 Then textboxcvvalue(33).Text = Val(textboxcvvalue(33).Text) + 128
End If

Call UpdateLabelFunctionLF

End Sub


Private Sub CheckBoxCV34_Click(Index As Integer)

If CheckBoxCV34(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 1
    If Index = 2 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 2
    If Index = 3 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 4
    If Index = 4 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 8
    If Index = 5 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 16
    If Index = 6 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 32
    If Index = 7 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 64
    If Index = 8 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) - 128
End If

If CheckBoxCV34(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 1
    If Index = 2 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 2
    If Index = 3 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 4
    If Index = 4 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 8
    If Index = 5 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 16
    If Index = 6 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 32
    If Index = 7 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 64
    If Index = 8 Then textboxcvvalue(34).Text = Val(textboxcvvalue(34).Text) + 128
End If

Call UpdateLabelFunctionLR

End Sub


Private Sub CheckBoxCV35_Click(Index As Integer)

If CheckBoxCV35(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 1
    If Index = 2 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 2
    If Index = 3 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 4
    If Index = 4 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 8
    If Index = 5 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 16
    If Index = 6 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 32
    If Index = 7 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 64
    If Index = 8 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) - 128
End If

If CheckBoxCV35(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 1
    If Index = 2 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 2
    If Index = 3 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 4
    If Index = 4 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 8
    If Index = 5 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 16
    If Index = 6 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 32
    If Index = 7 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 64
    If Index = 8 Then textboxcvvalue(35).Text = Val(textboxcvvalue(35).Text) + 128
End If

Call UpdateLabelFunction1

End Sub

Private Sub CheckBoxCV36_Click(Index As Integer)

If CheckBoxCV36(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 1
    If Index = 2 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 2
    If Index = 3 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 4
    If Index = 4 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 8
    If Index = 5 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 16
    If Index = 6 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 32
    If Index = 7 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 64
    If Index = 8 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) - 128
End If

If CheckBoxCV36(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 1
    If Index = 2 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 2
    If Index = 3 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 4
    If Index = 4 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 8
    If Index = 5 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 16
    If Index = 6 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 32
    If Index = 7 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 64
    If Index = 8 Then textboxcvvalue(36).Text = Val(textboxcvvalue(36).Text) + 128
End If

Call UpdateLabelFunction2
End Sub

Private Sub CheckBoxCV37_Click(Index As Integer)

If CheckBoxCV37(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 1
    If Index = 2 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 2
    If Index = 3 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 4
    If Index = 4 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 8
    If Index = 5 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 16
    If Index = 6 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 32
    If Index = 7 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 64
    If Index = 8 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) - 128
End If

If CheckBoxCV37(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 1
    If Index = 2 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 2
    If Index = 3 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 4
    If Index = 4 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 8
    If Index = 5 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 16
    If Index = 6 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 32
    If Index = 7 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 64
    If Index = 8 Then textboxcvvalue(37).Text = Val(textboxcvvalue(37).Text) + 128
End If

Call UpdateLabelFunction3

End Sub


Private Sub CheckBoxCV38_Click(Index As Integer)

If CheckBoxCV38(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 1
    If Index = 2 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 2
    If Index = 3 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 4
    If Index = 4 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 8
    If Index = 5 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 16
    If Index = 6 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 32
    If Index = 7 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 64
    If Index = 8 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) - 128
End If

If CheckBoxCV38(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 1
    If Index = 2 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 2
    If Index = 3 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 4
    If Index = 4 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 8
    If Index = 5 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 16
    If Index = 6 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 32
    If Index = 7 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 64
    If Index = 8 Then textboxcvvalue(38).Text = Val(textboxcvvalue(38).Text) + 128
End If

Call UpdateLabelFunction4

End Sub


Private Sub CheckBoxCV39_Click(Index As Integer)

If CheckBoxCV39(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 1
    If Index = 2 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 2
    If Index = 3 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 4
    If Index = 4 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 8
    If Index = 5 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 16
    If Index = 6 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 32
    If Index = 7 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 64
    If Index = 8 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) - 128
End If

If CheckBoxCV39(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 1
    If Index = 2 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 2
    If Index = 3 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 4
    If Index = 4 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 8
    If Index = 5 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 16
    If Index = 6 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 32
    If Index = 7 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 64
    If Index = 8 Then textboxcvvalue(39).Text = Val(textboxcvvalue(39).Text) + 128
End If

Call UpdateLabelFunction5

End Sub


Private Sub CheckBoxCV40_Click(Index As Integer)

If CheckBoxCV40(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 1
    If Index = 2 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 2
    If Index = 3 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 4
    If Index = 4 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 8
    If Index = 5 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 16
    If Index = 6 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 32
    If Index = 7 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 64
    If Index = 8 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) - 128
End If

If CheckBoxCV40(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 1
    If Index = 2 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 2
    If Index = 3 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 4
    If Index = 4 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 8
    If Index = 5 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 16
    If Index = 6 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 32
    If Index = 7 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 64
    If Index = 8 Then textboxcvvalue(40).Text = Val(textboxcvvalue(40).Text) + 128
End If

Call UpdateLabelFunction6

End Sub


Private Sub CheckBoxCV41_Click(Index As Integer)

If CheckBoxCV41(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 1
    If Index = 2 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 2
    If Index = 3 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 4
    If Index = 4 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 8
    If Index = 5 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 16
    If Index = 6 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 32
    If Index = 7 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 64
    If Index = 8 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) - 128
End If

If CheckBoxCV41(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 1
    If Index = 2 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 2
    If Index = 3 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 4
    If Index = 4 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 8
    If Index = 5 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 16
    If Index = 6 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 32
    If Index = 7 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 64
    If Index = 8 Then textboxcvvalue(41).Text = Val(textboxcvvalue(41).Text) + 128
End If

Call UpdateLabelFunction7

End Sub


Private Sub CheckBoxCV42_Click(Index As Integer)

If CheckBoxCV42(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 1
    If Index = 2 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 2
    If Index = 3 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 4
    If Index = 4 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 8
    If Index = 5 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 16
    If Index = 6 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 32
    If Index = 7 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 64
    If Index = 8 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) - 128
End If

If CheckBoxCV42(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 1
    If Index = 2 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 2
    If Index = 3 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 4
    If Index = 4 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 8
    If Index = 5 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 16
    If Index = 6 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 32
    If Index = 7 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 64
    If Index = 8 Then textboxcvvalue(42).Text = Val(textboxcvvalue(42).Text) + 128
End If

Call UpdateLabelFunction8

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
        If TemporaryScreen = "Programming Easy Screen Functions Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Programming Easy Screen Functions Screen"
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
            Let Ini.Value = "Programming Easy Screen Functions Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Programming Easy Screen Functions Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Programming Easy Screen Functions Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Programming Easy Screen Functions Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenFunctions.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenFunctions.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenFunctions.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenFunctions.Height)

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
                Let Ini.Value = "Programming Easy Screen Functions Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Programming Easy Screen Functions Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ProgrammingEasyScreenFunctions.Hide
    'unload ProgrammingEasyScreenfunctions

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
    Let Ini.Application = "Programming Easy Screen Functions Screen"
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
        ProgrammingEasyScreenFunctions.Left = (Screen.Width - Width) / 2
        ProgrammingEasyScreenFunctions.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + ProgrammingEasyScreenFunctions.Width > Screen.Width Then
            Let ProgrammingEasyScreenFunctions.Left = Screen.Width - ProgrammingEasyScreenFunctions.Width
        Else
            Let ProgrammingEasyScreenFunctions.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + ProgrammingEasyScreenFunctions.Height > Screen.Height Then
            Let ProgrammingEasyScreenFunctions.Top = Screen.Height - ProgrammingEasyScreenFunctions.Height
        Else
            Let ProgrammingEasyScreenFunctions.Top = Val(TemporaryValueTop)
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
        Call BalloonHelpPart1
        Call BalloonHelpPart2
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'No databases to declare
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update CheckBoxes
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryVariableT As Integer
    Dim TemporaryVariableY As Integer
    Dim TemporaryVariableZ As Integer
    Dim TemporaryVariableBitWeight As Integer

    For TemporaryVariableT = 33 To 42
    
    Let TemporaryVariableZ = Val(ProgrammingDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text)
        
        For TemporaryVariableY = 8 To 1 Step -1
        
        If TemporaryVariableY = 8 Then TemporaryVariableBitWeight = 128
        If TemporaryVariableY = 7 Then TemporaryVariableBitWeight = 64
        If TemporaryVariableY = 6 Then TemporaryVariableBitWeight = 32
        If TemporaryVariableY = 5 Then TemporaryVariableBitWeight = 16
        If TemporaryVariableY = 4 Then TemporaryVariableBitWeight = 8
        If TemporaryVariableY = 3 Then TemporaryVariableBitWeight = 4
        If TemporaryVariableY = 2 Then TemporaryVariableBitWeight = 2
        If TemporaryVariableY = 1 Then TemporaryVariableBitWeight = 1
        
        
        If Val(TemporaryVariableZ) / TemporaryVariableBitWeight >= 1 Then
            Let TemporaryVariableZ = TemporaryVariableZ - TemporaryVariableBitWeight
            If TemporaryVariableT = 33 Then Let CheckBoxCV33(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 34 Then Let CheckBoxCV34(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 35 Then Let CheckBoxCV35(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 36 Then Let CheckBoxCV36(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 37 Then Let CheckBoxCV37(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 38 Then Let CheckBoxCV38(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 39 Then Let CheckBoxCV39(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 40 Then Let CheckBoxCV40(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 41 Then Let CheckBoxCV41(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 42 Then Let CheckBoxCV42(TemporaryVariableY).Value = vbChecked
        Else
            If TemporaryVariableT = 33 Then Let CheckBoxCV33(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 34 Then Let CheckBoxCV34(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 35 Then Let CheckBoxCV35(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 36 Then Let CheckBoxCV36(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 37 Then Let CheckBoxCV37(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 38 Then Let CheckBoxCV38(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 39 Then Let CheckBoxCV39(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 40 Then Let CheckBoxCV40(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 41 Then Let CheckBoxCV41(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 42 Then Let CheckBoxCV42(TemporaryVariableY).Value = vbUnchecked
        End If
        
        Next TemporaryVariableY
        
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






Private Sub UpdateLabelFunctionLF()

Let Temporary$ = ""
Let LabelFunctionLF.Caption = "Function zero (white wire) controls no outputs."
For j = 1 To 8
    If CheckBoxCV33(j).Value = vbChecked Then
        Let LabelFunctionLF.Caption = "Function zero (white wire) controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j - 1) + " "
    End If
Next j
Let LabelFunctionLF.Caption = LabelFunctionLF.Caption + Temporary$

End Sub


Private Sub UpdateLabelFunction1()

Let Temporary$ = ""
Let LabelFunction1.Caption = "Function one controls no outputs."
For j = 1 To 8
    If CheckBoxCV35(j).Value = vbChecked Then
        Let LabelFunction1.Caption = "Function one controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j - 1) + " "
    End If
Next j
Let LabelFunction1.Caption = LabelFunction1.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunctionLR()

Let Temporary$ = ""
Let LabelFunctionLR.Caption = "Function zero (yellow wire) controls no outputs."
For j = 1 To 8
    If CheckBoxCV34(j).Value = vbChecked Then
        Let LabelFunctionLR.Caption = "Function zero (yellow wire) controls the follow outputs, "
        Let Temporary$ = Temporary$ + Str$(j - 1) + " "
    End If
Next j
Let LabelFunctionLR.Caption = LabelFunctionLR.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction2()

Let Temporary$ = ""
Let LabelFunction2.Caption = "Function two controls no outputs."
For j = 1 To 8
    If CheckBoxCV36(j).Value = vbChecked Then
        Let LabelFunction2.Caption = "Function two controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j - 1) + " "
    End If
Next j
Let LabelFunction2.Caption = LabelFunction2.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction3()

Let Temporary$ = ""
Let LabelFunction3.Caption = "Function three controls no outputs."
For j = 1 To 8
    If CheckBoxCV37(j).Value = vbChecked Then
        Let LabelFunction3.Caption = "Function three controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 1) + " "
    End If
Next j
Let LabelFunction3.Caption = LabelFunction3.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction4()

Let Temporary$ = ""
Let LabelFunction4.Caption = "Function four controls no outputs."
For j = 1 To 8
    If CheckBoxCV38(j).Value = vbChecked Then
        Let LabelFunction4.Caption = "Function four controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 1) + " "
    End If
Next j
Let LabelFunction4.Caption = LabelFunction4.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction5()

Let Temporary$ = ""
Let LabelFunction5.Caption = "Function five controls no outputs."
For j = 1 To 8
    If CheckBoxCV39(j).Value = vbChecked Then
        Let LabelFunction5.Caption = "Function five controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 1) + " "
    End If
Next j
Let LabelFunction5.Caption = LabelFunction5.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction6()

Let Temporary$ = ""
Let LabelFunction6.Caption = "Function six controls no outputs."
For j = 1 To 8
    If CheckBoxCV40(j).Value = vbChecked Then
        Let LabelFunction6.Caption = "Function six controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 1) + " "
    End If
Next j
Let LabelFunction6.Caption = LabelFunction6.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction7()

Let Temporary$ = ""
Let LabelFunction7.Caption = "Function seven controls no outputs."
For j = 1 To 8
    If CheckBoxCV41(j).Value = vbChecked Then
        Let LabelFunction7.Caption = "Function seven controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 5) + " "
    End If
Next j
Let LabelFunction7.Caption = LabelFunction7.Caption + Temporary$

End Sub

Private Sub UpdateLabelFunction8()

Let Temporary$ = ""
Let LabelFunction8.Caption = "Function eight controls no outputs."
For j = 1 To 8
    If CheckBoxCV42(j).Value = vbChecked Then
        Let LabelFunction8.Caption = "Function eight controls the following outputs, "
        Let Temporary$ = Temporary$ + Str$(j + 5) + " "
    End If
Next j
Let LabelFunction8.Caption = LabelFunction8.Caption + Temporary$

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    If ProgrammingEasyScreenFunctions.WindowState = vbMinimized Then
    
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
        
    ElseIf ProgrammingEasyScreenFunctions.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Programming Easy Screen Functions"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenFunctions.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenFunctions.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenFunctions.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenFunctions.Height)
 
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



Private Sub BalloonHelpPart1()
     
     Dim TemporaryText1 As String
     Dim TemporaryText2 As String
     Dim i As Long
     Dim t As Boolean
     Dim f As Boolean
     Let t = True
     Let f = False

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-three, outputs for function zero. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 33 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(33))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(33), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-four, outputs for function zero. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode -Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 34 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(34))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(34), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-five, outputs for function one. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode -Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 35 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(35))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(35), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-six, outputs for function two. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 36 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(36))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(36), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-seven, outputs for function three. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 37 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(37))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(37), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-eight, outputs for function four. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 38 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(38))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(38), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "thrity-nine, outputs for function five. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 39 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(39))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(39), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "fourty, outputs for function six. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 40 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(40))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(40), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "fourty-one, outputs for function seven. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 41 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(41))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(41), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

     Let TemporaryText1 = "This text box displays the value for configuration variable" & vbCrLf & "fourty-two, outputs for function eight. When  you 'Close'" & vbCrLf & "this screen the value of the configuration value will copied" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Configuration Variable 42 Value"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(42))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(42), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV33
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output one, the white" & vbCrLf & "wire (usually the front headlight), for function zero, front."
     Let TemporaryText2 = "Output One for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output two, the yellow" & vbCrLf & "wire (usually the rear headlight), for function zero, front."
     Let TemporaryText2 = "Output Two for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function zero, front."
     Let TemporaryText2 = "Output Three for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for function zero, front."
     Let TemporaryText2 = "Output Four for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function zero, front."
     Let TemporaryText2 = "Output Five for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function zero,front."
     Let TemporaryText2 = "Output Six for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function zero, front."
     Let TemporaryText2 = "Output Seven for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function zero, front."
     Let TemporaryText2 = "Output Eight for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV33(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV33(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV34
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output one, the white," & vbCrLf & "wire (usually the front headlight), for funtion zero, rear."
     Let TemporaryText2 = "Output One for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output two, the yellow" & vbCrLf & "wire (usually the rear headlight), for function zero, rear."
     Let TemporaryText2 = "Output Two for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the greeen" & vbCrLf & "wire, for function zero, rear."
     Let TemporaryText2 = "Output Three for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for function zero, rear."
     Let TemporaryText2 = "Output Four for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function zero, rear."
     Let TemporaryText2 = "Output Five for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function zero, rear."
     Let TemporaryText2 = "Output Six for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function zero, rear."
     Let TemporaryText2 = "Output Seven for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function zero, rear."
     Let TemporaryText2 = "Output Eight for Function Zero"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV34(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV34(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV35
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output one, the white" & vbCrLf & "wire, (usually the front headlight) for function one."
     Let TemporaryText2 = "Output One for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output two, the yellow" & vbCrLf & "wire, (usually the rear headlight) for funtion one."
     Let TemporaryText2 = "Output Two for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function one."
     Let TemporaryText2 = "Output Three for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for function one."
     Let TemporaryText2 = "Output Four for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function one."
     Let TemporaryText2 = "Output Five for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function one."
     Let TemporaryText2 = "Output Six for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function one."
     Let TemporaryText2 = "Output Seven for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function one."
     Let TemporaryText2 = "Output Eight for Function One"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV35(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV35(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV36
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output one, the white" & vbCrLf & "wire, (usually the front headlight) for function two."
     Let TemporaryText2 = "Output One for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output two, the yellow" & vbCrLf & "wire, (usually the rear headlight) for funtion two."
     Let TemporaryText2 = "Output Two for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function two."
     Let TemporaryText2 = "Output Three for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for function two."
     Let TemporaryText2 = "Output Four for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function two."
     Let TemporaryText2 = "Output Five for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function two."
     Let TemporaryText2 = "Output Six for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function two."
     Let TemporaryText2 = "Output Seven for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function two."
     Let TemporaryText2 = "Output Eight for Function Two"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV36(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV36(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV37
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function three."
     Let TemporaryText2 = "Output Three for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for funtion three."
     Let TemporaryText2 = "Output Four for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Five for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Six for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Seven for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Eight Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Nine for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function three."
     Let TemporaryText2 = "Output Ten for Function Three"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV37(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV37(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV38
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function four."
     Let TemporaryText2 = "Output Three for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for funtion four."
     Let TemporaryText2 = "Output Four for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Five for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Six for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Seven for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Eight for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Nine for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function four."
     Let TemporaryText2 = "Output Ten for Function Four"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV38(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV38(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV39
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function five."
     Let TemporaryText2 = "Output Three for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for funtion five."
     Let TemporaryText2 = "Output Four for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Five for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Six for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Seven for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Eight for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Nine for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function five."
     Let TemporaryText2 = "Output Ten for Function Five"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV39(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV39(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV40
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output three, the green" & vbCrLf & "wire, for function six."
     Let TemporaryText2 = "Output Three for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output four, the violet" & vbCrLf & "wire, for funtion six."
     Let TemporaryText2 = "Output Four for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output five" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Five for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output six" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Six for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Seven for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Eight for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Nine for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function six."
     Let TemporaryText2 = "Output Ten for Function Six"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV40(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV40(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV41
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & " for function seven."
     Let TemporaryText2 = "Output Seven for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     'Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for funtion seven."
     'Let TemporaryText2 = "Output Eight for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(2))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     'Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Nine for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(3))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     'Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Ten for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(4))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     'Let TemporaryText1 = "This checkbox enabled or disabled output eleven" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Eleven for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(5))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     'Let TemporaryText1 = "This checkbox enabled or disabled output twelve" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Twelve for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(6))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     'Let TemporaryText1 = "This checkbox enabled or disabled output thirteen" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Thirteen for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(7))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     'Let TemporaryText1 = "This checkbox enabled or disabled output fourteen" & vbCrLf & "for function seven."
     'Let TemporaryText2 = "Output Fourteen for Function Seven"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV41(8))
     'Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV41(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
End Sub

Private Sub BalloonHelpPart2()

     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     ' Check Boxes for CV42
     ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     Let TemporaryText1 = "This checkbox enabled or disabled output seven" & vbCrLf & " for function eight."
     Let TemporaryText2 = "Output Seven for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(1))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output eight" & vbCrLf & "for funtion eight."
     Let TemporaryText2 = "Output Eight for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(2))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output nine" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Nine for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(3))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This checkbox enabled or disabled output ten" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Ten for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(4))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
     Let TemporaryText1 = "This checkbox enabled or disabled output eleven" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Eleven for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(5))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output twelve" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Twelve for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(6))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output thirteen" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Thirteen for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(7))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
     Let TemporaryText1 = "This checkbox enabled or disabled output fourteen" & vbCrLf & "for function eight."
     Let TemporaryText2 = "Output Fourteen for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV42(8))
     Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV42(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
                  
     ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
     Let TemporaryText1 = "This button closes this screen and returns you to" & vbCrLf & "the 'Programming Mode - Decoder' screen."
     Let TemporaryText2 = "Output Fourteen for Function Eight"
     'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
     Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub
