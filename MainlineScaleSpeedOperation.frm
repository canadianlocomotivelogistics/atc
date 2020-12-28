VERSION 4.00
Begin VB.Form MainlineScaleSpeedOperation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Scale Speed Operations"
   ClientHeight    =   8010
   ClientLeft      =   1230
   ClientTop       =   1815
   ClientWidth     =   11400
   Height          =   8415
   Icon            =   "MainlineScaleSpeedOperation.frx":0000
   Left            =   1170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11400
   Top             =   1470
   Width           =   11520
   Begin VB.CommandButton ButonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   5640
      TabIndex        =   105
      Top             =   7390
      Width           =   1215
   End
   Begin VB.CheckBox IndicatorFaster 
      Caption         =   "Faster"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   10440
      TabIndex        =   104
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox IndicatorSlower 
      Caption         =   "Slower or"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   9360
      TabIndex        =   103
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox IndicatorFaster 
      Caption         =   "Faster"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   10440
      TabIndex        =   101
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CheckBox IndicatorFaster 
      Caption         =   "Faster"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   10440
      TabIndex        =   100
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox IndicatorFaster 
      Caption         =   "Faster"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   99
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox IndicatorSlower 
      Caption         =   "Slower or"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   98
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox IndicatorSlower 
      Caption         =   "Slower or"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   97
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox IndicatorSlower 
      Caption         =   "Slower or"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   96
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox Start 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10440
      TabIndex        =   93
      Top             =   120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TextBoxCurrentCV 
      Height          =   285
      Left            =   10440
      MaxLength       =   3
      TabIndex        =   91
      Top             =   2280
      Width           =   855
   End
   Begin VB.Data SpeedStepDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SpeedStepTable"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox CurrentCValue 
      Height          =   285
      Left            =   8640
      MaxLength       =   3
      TabIndex        =   88
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox TimerPerLap 
      Height          =   285
      Left            =   9960
      TabIndex        =   87
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TimerEnd 
      Height          =   285
      Left            =   9960
      TabIndex        =   86
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TimerStart 
      Height          =   285
      Left            =   9960
      TabIndex        =   85
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Locomotive Number or Consist"
      Height          =   735
      Left            =   120
      TabIndex        =   79
      Top             =   1920
      Width           =   6735
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4800
         Top             =   240
      End
      Begin VB.Data LocomotiveDatabaseSearch 
         Connect         =   "Access"
         DatabaseName    =   ""
         Enabled         =   0   'False
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomotiveDataBase"
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox ScaledSpeedLocomotiveNumber 
         Height          =   315
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton ButtonSet 
         Caption         =   "Set"
         Height          =   255
         Left            =   5400
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LocomotiveDatabaseNumberSearch 
         DataField       =   "LocomotiveNumber"
         DataSource      =   "LocomotiveDatabaseSearch"
         Height          =   285
         Left            =   3120
         TabIndex        =   82
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox LocomotiveDatabaseDecoderSearch 
         Caption         =   "Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "LocomotiveDatabaseSearch"
         Height          =   255
         Left            =   4440
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ShortAddress 
         Caption         =   "Short Address"
         Height          =   375
         Left            =   1800
         TabIndex        =   80
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox SpeedMatch 
      DataField       =   "ScaledSpeed"
      DataSource      =   "SpeedStepDatabase"
      Height          =   285
      Left            =   8640
      TabIndex        =   76
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox SpeedScaled 
      Height          =   285
      Left            =   8640
      TabIndex        =   73
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TimePerLap 
      Height          =   285
      Left            =   8640
      TabIndex        =   71
      Text            =   "00:00:00"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TimeEnd 
      Height          =   285
      Left            =   8640
      TabIndex        =   69
      Text            =   "00:00:00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TimeStart 
      Height          =   285
      Left            =   8640
      TabIndex        =   68
      Text            =   "00:00:00"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox LoopLength 
      Height          =   285
      Left            =   8640
      TabIndex        =   67
      Text            =   ".1554731"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton ButtonStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   62
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   5640
      TabIndex        =   47
      Top             =   7090
      Width           =   1215
   End
   Begin VB.Data LocomotiveDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDecoders"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV94D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   94
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   32
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV93D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   93
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   31
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV92D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   92
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   30
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV91D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   91
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   29
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV90D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   90
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV89D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   89
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   27
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV88D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   88
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   26
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV87D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   87
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   25
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV86D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   86
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   24
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV85D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   85
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   23
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV84D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   84
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   22
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV83D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   83
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   21
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton ButtonGraph 
      Caption         =   "Graph Speed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   3170
      Width           =   1215
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV82D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   82
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   19
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV81D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   81
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV80D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   80
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   17
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV79D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   79
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   16
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV78D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   78
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   15
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV77D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   77
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   14
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV76D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   76
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV75D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   75
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV74D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   74
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV73D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   73
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV72D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   72
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV71D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   71
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV70D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   70
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV69D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   69
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV68D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   68
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox RecommendedCVSetting 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV67D"
      DataSource      =   "LocomotiveDatabase"
      Height          =   285
      Index           =   67
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "MainlineScaleSpeedOperation.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonSpeedSetting 
      Caption         =   "Speed Setting"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7560
      Top             =   5940
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7560
      Top             =   5400
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7560
      Top             =   4800
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label49 
      Caption         =   "Previous time it needed to be"
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   102
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label49 
      Caption         =   "Previous time it needed to be"
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   95
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label49 
      Caption         =   "Previous time it needed to be"
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   94
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label50 
      Caption         =   "of CV"
      Height          =   255
      Left            =   9960
      TabIndex        =   92
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label49 
      Caption         =   "Locomotive needs to be"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   90
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label41 
      Caption         =   "Current CV Value"
      Height          =   255
      Left            =   7200
      TabIndex        =   89
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label40 
      Caption         =   "mph"
      Height          =   255
      Left            =   9960
      TabIndex        =   78
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label39 
      Caption         =   "Speed To Match"
      Height          =   255
      Left            =   7200
      TabIndex        =   77
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label38 
      Caption         =   "mph"
      Height          =   255
      Left            =   9960
      TabIndex        =   75
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "miles"
      Height          =   255
      Left            =   9960
      TabIndex        =   74
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label36 
      Caption         =   "Speed Scaled"
      Height          =   375
      Left            =   7200
      TabIndex        =   72
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label35 
      Caption         =   "Time per Lap"
      Height          =   255
      Left            =   7200
      TabIndex        =   70
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "End Time"
      Height          =   255
      Left            =   7200
      TabIndex        =   66
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label33 
      Caption         =   "Start Time"
      Height          =   255
      Left            =   7200
      TabIndex        =   65
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label32 
      Caption         =   "Length of Loop"
      Height          =   255
      Left            =   7200
      TabIndex        =   64
      Top             =   240
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   7920
   End
   Begin VB.Label Label31 
      Caption         =   $"MainlineScaleSpeedOperation.frx":0884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   63
      Top             =   960
      Width           =   6735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "CV67 - Step 1"
      Height          =   195
      Left            =   120
      TabIndex        =   61
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "CV68 - Step 2"
      Height          =   195
      Left            =   120
      TabIndex        =   60
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "CV69 - Step 3"
      Height          =   195
      Left            =   120
      TabIndex        =   59
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CV70 - Step 4"
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "CV71 - Step 5"
      Height          =   195
      Left            =   120
      TabIndex        =   57
      Top             =   4320
      Width           =   990
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "CV72 - Step 6"
      Height          =   195
      Left            =   120
      TabIndex        =   56
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "CV73 - Step 7"
      Height          =   195
      Left            =   120
      TabIndex        =   55
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "CV74 - Step 8"
      Height          =   195
      Left            =   120
      TabIndex        =   54
      Top             =   5400
      Width           =   990
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "CV75 - Step 9"
      Height          =   195
      Left            =   120
      TabIndex        =   53
      Top             =   5760
      Width           =   990
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "CV76 - Step 10"
      Height          =   195
      Left            =   120
      TabIndex        =   52
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "CV77 - Step 11"
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "CV78 - Step 12"
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   6840
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "CV79 - Step 13"
      Height          =   195
      Left            =   120
      TabIndex        =   49
      Top             =   7200
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "CV80 - Step 14"
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   7560
      Width           =   1080
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "CV81 - Step 15"
      Height          =   195
      Left            =   2880
      TabIndex        =   46
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "CV82 - Step 16"
      Height          =   195
      Left            =   2880
      TabIndex        =   45
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "CV83 - Step 17"
      Height          =   195
      Left            =   2880
      TabIndex        =   44
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "CV84 - Step 18"
      Height          =   195
      Left            =   2880
      TabIndex        =   43
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "CV85 - Step 19"
      Height          =   195
      Left            =   2880
      TabIndex        =   42
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "CV86 - Step 20"
      Height          =   195
      Left            =   2880
      TabIndex        =   41
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CV87 - Step 21"
      Height          =   195
      Left            =   2880
      TabIndex        =   40
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "CV88 - Step 22"
      Height          =   195
      Left            =   2880
      TabIndex        =   39
      Top             =   5400
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CV89 - Step 23"
      Height          =   195
      Left            =   2880
      TabIndex        =   38
      Top             =   5760
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CV90 - Step 24"
      Height          =   195
      Left            =   2880
      TabIndex        =   37
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CV91 - Step 25"
      Height          =   195
      Left            =   2880
      TabIndex        =   36
      Top             =   6480
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CV92 - Step 26"
      Height          =   195
      Left            =   2880
      TabIndex        =   35
      Top             =   6840
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CV93 - Step 27"
      Height          =   195
      Left            =   2880
      TabIndex        =   34
      Top             =   7200
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CV94 - Step 28"
      Height          =   195
      Left            =   2880
      TabIndex        =   33
      Top             =   7560
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   $"MainlineScaleSpeedOperation.frx":091F
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "MainlineScaleSpeedOperation"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButonPrint_Click()

    mainlinescalespeedoperaton.PrintForm
    
End Sub


Private Sub ButtonClose_Click()
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
            If TemporaryScreen = "Mainline Scale Speed Operation Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Scale Speed Operation Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Scale Speed Operation Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Scale Speed operation Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub ButtonSet_Click()

If ScaledSpeedLocomotiveNumber.Text <> "" Then
    If ShortAdDress.Value = vbUnchecked Then
        Let onebyted.Text = Int(Val(ScaledSpeedLocomotiveNumber.Text) / 256)
        Let twoByteD.Text = Val(ScaledSpeedLocomotiveNumber.Text) - (Val(onebyted.Text) * 256)
        Let onebyted.Text = Val(onebyted.Text) + 128 + 64
        Let ScaledSpeedComment.Text = "Loco " + ScaledSpeedLocomotiveNumber.Text + "; "
    End If
    If ShortAdDress.Value = vbChecked Then
        Let onebyted.Text = Int(Val(ScaledSpeedLocomotiveNumber.Text))
        Let twoByteD.Text = ""
        Let ScaledSpeedComment.Text = "Consist " + ScaledSpeedLocomotiveNumber.Text + "; "
    End If
    Let ButtonStart.Enabled = True
    Let ButtonGraph.Enabled = True
    
    LocomotiveDatabase.Recordset.MoveFirst
    For t = 1 To (Val(ScaledSpeedLocomotiveNumber.Text))
    LocomotiveDatabase.Recordset.MoveNext
    Next t
    
    
 
Else
    Let ButtonStart.Enabled = False
    Let ButtonGraph.Enabled = False
End If

End Sub


Private Sub ButtonSpeedSetting_Click()
 
Load MainlineScaleSpeedSetting
MainlineScaleSpeedSetting.Show vbModeless

End Sub


Private Sub ButtonStart_Click()

Let TextBoxCurrentCV.Text = 94

'For TextBoxCurrentCV.Text = 67 To 94

Let CurrentCValue.Text = RecommendedCVSetting(Val(TextBoxCurrentCV.Text))

End Sub

Private Sub CurrentCValue_Change()

' ----------------------------------------------------------------------------------------------



Let ScaledSpeedComment.Text = "Speed 0"
Let Temporary = 64
    'If ConsistControlDirectionF.Value = vbChecked Then
            Let Temporary = Temporary + 32 ' add forward direction
    'End If
        'This routine assenmles the byte for speed step mode 28
        Let temp1 = 0 ' adds the speed
        Let temp2 = temp1 Mod 2
        Let newspeedvalue = Int(temp1 / 2)
        Let Temporary = Temporary + newspeedvalue
        If temp2 = 1 Then Let Temporary = Temporary + 16
        Let ThreeByteD.Text = Temporary
        Let FourByteD.Text = ""
        Let FiveByteD.Text = ""
        Let SixByteD.Text = ""
  
Call MainScreen.SendCommandviaTrackQ
DoEvents

' -------------------------------------------------------------------------------------------------------------------------
        Let TemporaryByteOne = 0
        Let TemporaryByteTwo = Val(TextBoxCurrentCV.Text) - 1
        
        If TemporaryByteTwo / 512 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 2
            Let TemporaryByteTwo = TemporaryByteTwo - 512
        End If
        
        If TemporaryByteTwo / 256 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 1
            Let TemporaryByteTwo = TemporaryByteTwo - 256
        End If
        
        Let TemporaryByteOne = TemporaryByteOne + 128
        Let TemporaryByteOne = TemporaryByteOne + 64
        Let TemporaryByteOne = TemporaryByteOne + 32
        
       ' If ConsistControlCVRead = vbChecked Then
       '       Let TemporaryByteOne = TemporaryByteOne + 4
       ' Else
            Let TemporaryByteOne = TemporaryByteOne + 8 + 4
       ' End If
        
    Let ThreeByteD.Text = TemporaryByteOne
    Let FourByteD.Text = TemporaryByteTwo
    Let FiveByteD.Text = Val(CurrentCValue.Text)
    Let SixByteD.Text = ""
    
 Let ScaledSpeedComment.Text = "Change CV" + TextBoxCurrentCV.Text + " to " + CurrentCValue.Text

Call MainScreen.SendCommandviaTrackQ
DoEvents




Let ScaledSpeedComment.Text = "Speed " & Val(TextBoxCurrentCV.Text) - 66

Let Temporary = 64
    'If ConsistControlDirectionF.Value = vbChecked Then
        Let Temporary = Temporary + 32 ' add forward direction
    'End If
    
    'This routine assenmles the byte for speed step mode 28
        Let temp1 = Val(TextBoxCurrentCV.Text) - 94 + 31 ' adds the speed
        Let temp2 = temp1 Mod 2
        Let newspeedvalue = Int(temp1 / 2)
        Let Temporary = Temporary + newspeedvalue
        If temp2 = 1 Then Let Temporary = Temporary + 16
        Let ThreeByteD.Text = Temporary
        Let FourByteD.Text = ""
        Let FiveByteD.Text = ""
        Let SixByteD.Text = ""


Call MainScreen.SendCommandviaTrackQ
DoEvents

End Sub


Private Sub FiveByteD_Change()
 
Rem Hexidecimal Converstion
        
    Let FiveByteH.Text = Hex(Val(FiveByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(FiveByteH.Text) = 1 Then
        Let FiveByteH.Text = "0" + FiveByteH.Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(FiveByteD.Text)
        
        If temp / 128 >= 1 Then
            Let FiveByteB.Text = "1"
            Let temp = temp - 128
            Else: FiveByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 64
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 32
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 16
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 8
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 4
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 2
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let FiveByteB.Text = FiveByteB.Text + "1"
            Let temp = temp - 1
            Else: FiveByteB.Text = FiveByteB.Text + "0"
        End If
  
End Sub

Private Sub Form_Activate()

    DoEvents
    
' =============================================================================================================================================================================
' Add to Screen Stack
' =============================================================================================================================================================================
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
        If TemporaryScreen = "Mainline Scale Speed Operation Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Scale Speed Operation Screen"
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
            Let Ini.Value = "Mainline Scale Speed Operation Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Scale Speed Operation Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Scale Speed Operation Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Scale Speed Operation Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineScaleSpeedOperation.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineScaleSpeedOperation.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineScaleSpeedOperation.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineScaleSpeedOperation.Height)

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
                Let Ini.Value = "Mainline Scale Speed Operation Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Scale Speed Operation Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineScaleSpeedOperation.Hide
    'unload Mainlinescalespeedoperation

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
    Let Ini.Application = "Mainline Scale Speed Operation Screen"
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
        MainlineScaleSpeedOperation.Left = (Screen.Width - Width) / 2
        MainlineScaleSpeedSetting.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineScaleSpeedSetting.Width > Screen.Width Then
            Let MainlineScaleSpeedSetting.Left = Screen.Width - MainlineScaleSpeedSetting.Width
        Else
            Let MainlineScaleSpeedSetting.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineScaleSpeedSetting.Height > Screen.Height Then
            Let MainlineScaleSpeedSetting.Top = Screen.Height - MainlineScaleSpeedSetting.Height
        Else
            Let MainlineScaleSpeedSetting.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This text box is where all information from your" & vbCrLf & "serial port is displayed. Commands given by the" & vbCrLf & "program are displayed here. You can also type your" & vbCrLf & "own commands, providing the port is not busy."
        Let TemporaryText2 = "Communication Window"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowDCC)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowDCC, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
Let LocomotiveDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDecoders.mdb"
Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
Let SpeedStepDatabase.DatabaseName = App.Path$ & "\Databases\SpeedStepDatabase.mdb"

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    If MainlineScaleSpeedOperation.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineScaleSpeedOperation.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Scale Speed Operation Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineScaleSpeedOperation.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineScaleSpeedOperation.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineScaleSpeedOperation.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineScaleSpeedOperation.Height)
 
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

Private Sub FourByteD_Change()
   
Rem Hexidecimal Converstion
        
    Let FourByteH.Text = Hex(Val(FourByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(FourByteH.Text) = 1 Then
        Let FourByteH.Text = "0" + FourByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(FourByteD.Text)
        
        If temp / 128 >= 1 Then
            Let FourByteB.Text = "1"
            Let temp = temp - 128
            Else: FourByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 64
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 32
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 16
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 8
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 4
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 2
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 1
            Else: FourByteB.Text = FourByteB.Text + "0"
        End If

End Sub

Private Sub LocomotiveDatabase_Reposition()

Let LocomotiveDatabase.Caption = LocomotiveDatabase.Recordset.AbsolutePosition

End Sub

Private Sub OneByteD_Change()
  
Rem Hexidecimal Converstion
        
    Let OneByteH.Text = Hex(Val(onebyted.Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(OneByteH.Text) = 1 Then
        Let OneByteH.Text = "0" + OneByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(onebyted.Text)
        
        If temp / 128 >= 1 Then
            Let OneByteB.Text = "1"
            Let temp = temp - 128
            Else: OneByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 64
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 32
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 16
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 8
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 4
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 2
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let OneByteB.Text = OneByteB.Text + "1"
            Let temp = temp - 1
            Else: OneByteB.Text = OneByteB.Text + "0"
        End If

End Sub

Private Sub SevenByteD_Change()
  
Rem Hexidecimal Converstion
        
    Let SevenByteH.Text = Hex(Val(sevenbyted.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SevenByteH.Text) = 1 Then
        Let SevenByteH.Text = "0" + SevenByteH.Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(sevenbyted.Text)
        
        If temp / 128 >= 1 Then
            Let SevenByteB.Text = "1"
            Let temp = temp - 128
            Else: SevenByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 64
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 32
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 16
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 8
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 4
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 2
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 1
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If

End Sub

Private Sub SixByteD_Change()
  
Rem Hexidecimal Converstion
        
    Let SixByteH.Text = Hex(Val(SixByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SixByteH.Text) = 1 Then
        Let SixByteH.Text = "0" + SixByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(SixByteD.Text)
        
        If temp / 128 >= 1 Then
            Let SixByteB.Text = "1"
            Let temp = temp - 128
            Else: SixByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 64
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 32
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 16
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 8
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 4
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 2
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 1
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If

End Sub

Private Sub SpeedScaled_Change()

' -------------------------------------------------------------------------------------------------------------------------
' Speed Changes
'
' -------------------------------------------------------------------------------------------------------------------------
'
' Everytime the locomotive goes around the test track, the spped of the locomotive is measure and compared to the last
' time. The difference in time is used to calculate the speed of the locomotive. This scale spped is compared to my
' database for a required scale speed.
'
' The following code generate a indicator of weather the locomotive should be going faster or slower. The valuse of the
' last three spped changes are stored. TheoretiCally the locomotive should generate a slower-faster-slower or
' faster-slower-faster responce once the correct spped scale is found.
'
' My Code
'
' Reomve the last indicator value be placing the new value there, then place the most recent value in.
   
    Let IndicatorSlower(3).Value = IndicatorSlower(2).Value
    Let IndicatorSlower(2).Value = IndicatorSlower(1).Value
    Let IndicatorSlower(1).Value = IndicatorSlower(0).Value
    Let IndicatorFaster(3).Value = IndicatorFaster(2).Value
    Let IndicatorFaster(2).Value = IndicatorFaster(1).Value
    Let IndicatorFaster(1).Value = IndicatorFaster(0).Value

    If Val(SpeedScaled.Text) < Val(SpeedMatch.Text) Then
        Let IndicatorSlower(0).Value = vbUnchecked
        Let IndicatorFaster(0).Value = vbChecked
    Else
        Let IndicatorFaster(0).Value = vbUnchecked
        Let IndicatorSlower(0).Value = vbChecked
    End If

' -------------------------------------------------------------------------------------------------------------------------
'
' Testing the last three laps.
'
' Now, as the locomotive come closer to the require scale speed, the indicators will show a contunous faster or slower
' response. But When the requires speed is found, the locomorive should generate a slower-faster-slower or
' faster-slower-faster response with the indicator. When this occurs, when can move to the next configuration variable in
' the speed table; and the test continues until the last speed step.
'
' My Code
'
' So, we nee to check the indicators for the response; and make changes to the confiruartion varible being adjusted if we
' are conpleted finding this required scale speed.
'

    If IndicatorSlower(0).Value = vbChecked And IndicatorFaster(0).Value = vbUnchecked And _
       IndicatorFaster(1).Value = vbChecked And IndicatorSlower(1).Value = vbUnchecked And _
       IndicatorSlower(2).Value = vbChecked And IndicatorFaster(2).Value = vbUnchecked And _
       IndicatorFaster(3).Value = vbChecked And IndicatorSlower(3).Value = vbUnchecked Or _
       IndicatorSlower(0).Value = vbUnchecked And IndicatorFaster(0).Value = vbChecked And _
       IndicatorFaster(1).Value = vbUnchecked And IndicatorSlower(1).Value = vbChecked And _
       IndicatorSlower(2).Value = vbUnchecked And IndicatorFaster(2).Value = vbChecked And _
       IndicatorFaster(3).Value = vbUnchecked And IndicatorSlower(3).Value = vbChecked Then
        For i = 0 To 2
            Let IndicatorFaster(i).Value = vbUnchecked
            Let IndicatorSlower(i).Value = vbUnchecked
        Next i
        Let RecommendedCVSetting(TextBoxCurrentCV.Text) = Val(CurrentCValue.Text)
        Let TextBoxCurrentCV.Text = Val(TextBoxCurrentCV.Text) - 1
        If Val(TextBoxCurrentCV.Text) = 66 Then Stop
    End If
    
    

    
' -------------------------------------------------------------------------------------------------------------------------
'
' Adjusting the CV setting
'
' If the locomotive is going to fast, then reduce the value of the configuration variable accordingly. If the locomotive
' is going to slow then increase the valuse of the configuration variable accordingly.


    Let SpeedDifference = 1
    If Abs(Val(SpeedMatch.Text) - Val(SpeedScaled.Text)) > 2 Then
        Let SpeedDifference = 2
        If Abs(Val(SpeedMatch.Text) - Val(SpeedScaled.Text)) > 5 Then
            Let SpeedDifference = 5
            If Abs(Val(SpeedMatch.Text) - Val(SpeedScaled.Text)) > 10 Then
                Let SpeedDifference = 10
            End If
        End If
    End If

        
    If IndicatorFaster(0).Value = vbUnchecked And IndicatorSlower(0).Value = vbChecked Then
        Let CurrentCValue.Text = Val(CurrentCValue.Text) - SpeedDifference
        
    End If
    
    If IndicatorFaster(0).Value = vbChecked And IndicatorSlower(0).Value = vbUnchecked Then
        Let CurrentCValue.Text = Val(CurrentCValue.Text) + SpeedDifference
       
    End If

' -------------------------------------------------------------------------------------------------------------------------

End Sub



Private Sub SpeedStepDatabase_Reposition()

Let SpeedStepDatabase.Caption = SpeedStepDatabase.Recordset.AbsolutePosition

End Sub

Private Sub TextBoxCurrentCV_Change()

If TextBoxCurrentCV.Text = 67 Then RecordNumber = 4
If TextBoxCurrentCV.Text = 68 Then RecordNumber = 9
If TextBoxCurrentCV.Text = 69 Then RecordNumber = 13
If TextBoxCurrentCV.Text = 70 Then RecordNumber = 18
If TextBoxCurrentCV.Text = 71 Then RecordNumber = 22
If TextBoxCurrentCV.Text = 72 Then RecordNumber = 27
If TextBoxCurrentCV.Text = 73 Then RecordNumber = 32
If TextBoxCurrentCV.Text = 74 Then RecordNumber = 36
If TextBoxCurrentCV.Text = 75 Then RecordNumber = 41
If TextBoxCurrentCV.Text = 76 Then RecordNumber = 45
If TextBoxCurrentCV.Text = 77 Then RecordNumber = 50
If TextBoxCurrentCV.Text = 78 Then RecordNumber = 54
If TextBoxCurrentCV.Text = 79 Then RecordNumber = 59
If TextBoxCurrentCV.Text = 80 Then RecordNumber = 63
If TextBoxCurrentCV.Text = 81 Then RecordNumber = 68
If TextBoxCurrentCV.Text = 82 Then RecordNumber = 72
If TextBoxCurrentCV.Text = 83 Then RecordNumber = 77
If TextBoxCurrentCV.Text = 84 Then RecordNumber = 81
If TextBoxCurrentCV.Text = 85 Then RecordNumber = 86
If TextBoxCurrentCV.Text = 86 Then RecordNumber = 90
If TextBoxCurrentCV.Text = 87 Then RecordNumber = 94
If TextBoxCurrentCV.Text = 88 Then RecordNumber = 99
If TextBoxCurrentCV.Text = 89 Then RecordNumber = 103
If TextBoxCurrentCV.Text = 90 Then RecordNumber = 108
If TextBoxCurrentCV.Text = 91 Then RecordNumber = 112
If TextBoxCurrentCV.Text = 92 Then RecordNumber = 117
If TextBoxCurrentCV.Text = 93 Then RecordNumber = 121
If TextBoxCurrentCV.Text = 94 Then RecordNumber = 126

SpeedStepDatabase.Recordset.AbsolutePosition = Val(RecordNumber)

End Sub


Private Sub ThreeByteD_Change()

Rem Hexidecimal Converstion
        
    Let ThreeByteH.Text = Hex(Val(ThreeByteD.Text))
    
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(ThreeByteH.Text) = 1 Then
        Let ThreeByteH.Text = "0" + ThreeByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(ThreeByteD.Text)
        
        If temp / 128 >= 1 Then
            Let ThreeByteB.Text = "1"
            Let temp = temp - 128
            Else: ThreeByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 64
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 32
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 16
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 8
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 4
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 2
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 1
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If

End Sub

Private Sub Timer1_Timer()

    Let ButtonClose.Enabled = False

    Let Timer1.Interval = 0

' Returns or sets a value indicating the type of mouse pointer displayed when the mouse is over a particular part of an object at run time.

    ScaledSpeedLocomotiveNumber.Clear

' Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.
    LocomotiveDatabaseSearch.Refresh
    LocomotiveDatabaseSearch.Recordset.MoveFirst
   
    Do
        If LocomotiveDatabaseDecoderSearch.Value = vbChecked Then
            ScaledSpeedLocomotiveNumber.AddItem Format(Val(LocomotiveDatabaseNumberSearch.Text), "0000")
        End If
        LocomotiveDatabaseSearch.Recordset.MoveNext
    Loop While Not LocomotiveDatabaseSearch.Recordset.EOF

    Let ButtonClose.Enabled = True

End Sub


Private Sub TimerPerLap_Change()
    Let temp = 3600 / Val(TimerPerLap.Text) / (1 / LoopLength.Text)
    Let SpeedScaled.Text = Int(temp * 100) / 100
End Sub

Private Sub TwoByteD_Change()
  
Rem Hexidecimal Converstion
        
    Let TwoByteH.Text = Hex(Val(twoByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(TwoByteH.Text) = 1 Then
        Let TwoByteH.Text = "0" + TwoByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(twoByteD.Text)
        
        If temp / 128 >= 1 Then
            Let TwoByteB.Text = "1"
            Let temp = temp - 128
            Else: TwoByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 64
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 32
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 16
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 8
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 4
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 2
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 1
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If

End Sub





