VERSION 4.00
Begin VB.Form MainlineOperationGUI 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode  - GUI"
   ClientHeight    =   9330
   ClientLeft      =   1515
   ClientTop       =   1455
   ClientWidth     =   9630
   Height          =   9735
   Icon            =   "MainlineOperationGUI.frx":0000
   KeyPreview      =   -1  'True
   Left            =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   Top             =   1110
   Width           =   9750
   Begin VB.Timer TimerWindowState 
      Enabled         =   0   'False
      Left            =   20880
      Top             =   1740
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   12660
      TabIndex        =   127
      Top             =   3360
      Width           =   2535
      Begin VB.CheckBox SoundControlFunction9 
         Caption         =   "Function 9"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction10 
         Caption         =   "Function 10"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox SoundControlFunction11 
         Caption         =   "Function 11"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox SoundControlFunction12 
         Caption         =   "Function 12"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Function Controls - Group Two"
      Height          =   1455
      Left            =   12660
      TabIndex        =   122
      Top             =   1800
      Width           =   2535
      Begin VB.CheckBox SoundControlFunction5 
         Caption         =   "Function 5"
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction6 
         Caption         =   "Function 6"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction7 
         Caption         =   "Function 7"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction8 
         Caption         =   "Function 8"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Function Controls - Group One"
      Height          =   1575
      Left            =   12660
      TabIndex        =   116
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox SoundControlFunction0 
         Caption         =   "Function 0"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction1 
         Caption         =   "Function 1"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction2 
         Caption         =   "Function 2"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction3 
         Caption         =   "Function 3"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox SoundControlFunction4 
         Caption         =   "Function 4"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   8340
      TabIndex        =   115
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   7080
      TabIndex        =   114
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton ButtonInternet 
      Caption         =   "&Internet"
      Height          =   255
      Left            =   60
      TabIndex        =   113
      Top             =   9000
      Width           =   1230
   End
   Begin VB.Frame Frame6 
      Caption         =   "Objects for Sound Decoder Search"
      Height          =   1335
      Left            =   3540
      TabIndex        =   108
      Top             =   11550
      Width           =   2535
      Begin VB.Timer Timer2 
         Interval        =   4
         Left            =   1440
         Top             =   360
      End
      Begin VB.Data SoundDecoderDatabase 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomotiveDatabaseShort"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox SoundDecoderNumber 
         DataField       =   "LocomotiveNumber"
         DataSource      =   "SoundDecoderDatabase"
         Height          =   285
         Left            =   120
         TabIndex        =   110
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox SoundDecoderEquiped 
         Caption         =   "Sound Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "SoundDecoderDatabase"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Objects for Locomotive Search"
      Height          =   1335
      Left            =   6180
      TabIndex        =   105
      Top             =   11550
      Width           =   2535
      Begin VB.Timer Timer1 
         Interval        =   3
         Left            =   1440
         Top             =   360
      End
      Begin VB.CheckBox LocomotiveDecoderEquiped 
         Caption         =   "Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "LocomotiveDatabase"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox LocomotiveNumber 
         DataField       =   "LocomotiveNumber"
         DataSource      =   "LocomotiveDatabase"
         Height          =   285
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data LocomotiveDatabase 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomotiveDataBase"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   9960
      TabIndex        =   90
      Top             =   3360
      Width           =   2535
      Begin VB.CheckBox ConsistControlFunction12 
         Caption         =   "Function 12"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlFunction11 
         Caption         =   "Function 11"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlFunction10 
         Caption         =   "Function 10"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlFunction9 
         Caption         =   "Function 9"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameBrakes 
      Caption         =   "Brakes System"
      Height          =   1215
      Left            =   4860
      TabIndex        =   82
      Top             =   6300
      Width           =   4695
      Begin VB.OptionButton OptionSelfLapingBrakesNo 
         Caption         =   "No, thankyou."
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   85
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton OptionSelfLapingBrakesYes 
         Caption         =   "Yes, please."
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Would you like to implement self laping brakes for this locomoyive?"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Timer TimerVirtualCabSignal 
      Enabled         =   0   'False
      Left            =   14160
      Top             =   11040
   End
   Begin VB.Frame FrameVirtualSignal 
      Caption         =   "Signaling System"
      Height          =   1215
      Left            =   4860
      TabIndex        =   60
      Top             =   7680
      Width           =   4695
      Begin VB.OptionButton OptionVirtualCabSignalNo 
         Caption         =   "No, thankyou."
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   63
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton OptionVirtualCabSignalYes 
         Caption         =   "Yes, please."
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   62
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Would you like to use the virtual cab signaling system with the program?"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame FrameLocomotiveType 
      Caption         =   "Locomotive Type"
      Height          =   2235
      Left            =   4860
      TabIndex        =   56
      Top             =   840
      Width           =   4695
      Begin VB.OptionButton OptionGuiDiesel4 
         Caption         =   "Diesel 4 (SD80MAC)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   132
         Top             =   1380
         Width           =   1800
      End
      Begin VB.OptionButton OptionGuiDiesel3 
         Caption         =   "Diesel 3 (SD40-2)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   112
         Top             =   1140
         Width           =   1575
      End
      Begin VB.OptionButton OptionGuiElectric1 
         Caption         =   "Electric 1 (Acela)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   111
         Top             =   1860
         Width           =   1530
      End
      Begin VB.OptionButton OptionGuiDiesel2 
         Caption         =   "Diesel 2 (GP38-2)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   99
         Top             =   900
         Width           =   1560
      End
      Begin VB.OptionButton OptionGuiDiesel1 
         Caption         =   "Diesel 1 (GE 9-40CW)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   58
         Top             =   660
         Value           =   -1  'True
         Width           =   1920
      End
      Begin VB.OptionButton OptionGuiSteam1 
         Caption         =   "Steam 1 (Flying Scottsman)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   57
         Top             =   1620
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Would you like to use the 'diesel' or the 'steam' graphic user interface?"
         Height          =   495
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Timer TimerRPM 
      Left            =   14160
      Top             =   12000
   End
   Begin VB.Timer TimerAirAutomaticBrake 
      Left            =   14160
      Top             =   7200
   End
   Begin VB.Timer TimerDeadmann 
      Left            =   14160
      Top             =   12480
   End
   Begin VB.Timer TimerRadiatorFans 
      Left            =   8880
      Top             =   11040
   End
   Begin VB.Timer TimerSandConsumption 
      Left            =   8880
      Top             =   11520
   End
   Begin VB.Timer TimerWaterConsumption 
      Left            =   8880
      Top             =   12000
   End
   Begin VB.Timer TimerOilConsumption 
      Left            =   8880
      Top             =   12480
   End
   Begin VB.Timer TimerFuelConsumption 
      Left            =   14160
      Top             =   10560
   End
   Begin VB.Timer TimerWaterTemperature 
      Left            =   14160
      Top             =   10080
   End
   Begin VB.Timer TimerWaterPressure 
      Left            =   14160
      Top             =   9600
   End
   Begin VB.Timer TimerOilPressure 
      Left            =   14160
      Top             =   8640
   End
   Begin VB.Timer TimerOilTemperature 
      Left            =   14160
      Top             =   9120
   End
   Begin VB.Timer TimerThrottleDelay 
      Left            =   14160
      Top             =   8160
   End
   Begin VB.Timer TimerAirIndependentBrake 
      Left            =   14160
      Top             =   7680
   End
   Begin VB.Timer TimerAirMainReservoir 
      Left            =   14160
      Top             =   6720
   End
   Begin VB.Timer TimerSpeedChange 
      Left            =   14160
      Top             =   6240
   End
   Begin VB.Timer TimerSendCommand 
      Left            =   14160
      Top             =   5760
   End
   Begin VB.Frame FrameSoundDecoder 
      Caption         =   "Sound Decoder"
      Height          =   6075
      Left            =   60
      TabIndex        =   35
      Top             =   2820
      Width           =   4695
      Begin VB.CheckBox CheckBoxDigitraxFX 
         Caption         =   "Digitrax FX"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   136
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox CheckboxQSI 
         Caption         =   "Quantum Sound Industries"
         Height          =   195
         Left            =   360
         TabIndex        =   134
         Top             =   1740
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox CheckboxSoundtraxxDSX 
         Caption         =   "Soundtraxx DSX"
         Height          =   195
         Left            =   360
         TabIndex        =   133
         Top             =   1560
         Width           =   1515
      End
      Begin VB.ComboBox ComboBoxSoundDecoderNumber 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainlineOperationGUI.frx":0442
         Left            =   3720
         List            =   "MainlineOperationGUI.frx":0444
         TabIndex        =   101
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox TextboxMappedFunctionExtra4 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   89
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox TextboxMappedFunctionExtra3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   88
         Top             =   5340
         Width           =   615
      End
      Begin VB.TextBox TextboxMappedFunctionExtra2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   87
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox TextboxMappedFunctionExtra1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   86
         Top             =   4740
         Width           =   615
      End
      Begin VB.CheckBox CheckBoxAddressShortSoundDecoder 
         Caption         =   "short address (1 to 127) ?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TextBoxMappedFunctionHorn 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   50
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionBell 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   49
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionCoupler 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   48
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionDynamicBrake 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   47
         Top             =   3540
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionRPMPlus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   46
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionRPMMinus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   45
         Top             =   4140
         Width           =   615
      End
      Begin VB.TextBox TextBoxMappedFunctionMute 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   44
         Top             =   4440
         Width           =   615
      End
      Begin VB.CheckBox CheckBoxSoundDecoderEquipped 
         Caption         =   "Sound Decoder Equipped?"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "What function activates the following sounds? "
         Height          =   195
         Left            =   180
         TabIndex        =   137
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Who is the manufacturer of the sound decoder?"
         Height          =   195
         Left            =   180
         TabIndex        =   135
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Line Line3 
         X1              =   4560
         X2              =   240
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         X1              =   4560
         X2              =   240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label LabelMappedSoundDecoder 
         Caption         =   "Sound decoder address is ?"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   103
         Top             =   540
         Width           =   3135
      End
      Begin VB.Label Label13 
         Caption         =   "Is this a"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   720
         Width           =   615
      End
      Begin VB.Label LabelMappedExtraOption4 
         Caption         =   "What is Extra Option Four mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   98
         Top             =   5640
         Width           =   3495
      End
      Begin VB.Label LabelMappedExtraOption3 
         Caption         =   "What is Extra Option Three mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   97
         Top             =   5340
         Width           =   3495
      End
      Begin VB.Label LabelMappedExtraOption2 
         Caption         =   "What is Extra Option Two mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   96
         Top             =   5040
         Width           =   3495
      End
      Begin VB.Label LabelMappedExtraOption1 
         Caption         =   "What is Extra Option One mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   95
         Top             =   4740
         Width           =   3495
      End
      Begin VB.Label LabelMappedMute 
         Caption         =   "What function is the mute mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   4440
         Width           =   3495
      End
      Begin VB.Label LabelMappedRpmMinus 
         Caption         =   "What function is the rpm minus mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   4200
         Width           =   3495
      End
      Begin VB.Label LabelMappedRpmPlus 
         Caption         =   "What function is the rpm plus mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   3900
         Width           =   3495
      End
      Begin VB.Label LabelMappedDynamicBrake 
         Caption         =   "What function is the dynamic brake mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label LabelMappedCoupler 
         Caption         =   "What function is the coupler mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   3300
         Width           =   3495
      End
      Begin VB.Label LabelMappedBell 
         Caption         =   "What function is the bell mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   2700
         Width           =   3495
      End
      Begin VB.Label LabelMappedHorn 
         Caption         =   "What function is the horn mapped to?"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   3000
         Width           =   3495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "MainlineOperationGUI.frx":0446
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   33
      Top             =   120
      Width           =   510
   End
   Begin VB.CommandButton ButtonGUI 
      Caption         =   "&GUI"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame Frame8 
      Caption         =   "Dynamic Brakes"
      Height          =   615
      Left            =   9960
      TabIndex        =   30
      Top             =   8040
      Width           =   2535
      Begin VB.HScrollBar ConsistControlDynamicBrake 
         Height          =   255
         Left            =   120
         Max             =   9
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Changing Configuration Variables"
      Height          =   1575
      Left            =   9960
      TabIndex        =   20
      Top             =   8760
      Width           =   2535
      Begin VB.CheckBox ConsistControlCVWright 
         Caption         =   "Wright Configuration"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ConsistControlCVRead 
         Caption         =   "Read or"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox ConsistControlCVValue 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox ConsistControlCV 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Configuration Variable"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Function Controls - Group Two"
      Height          =   1455
      Left            =   9960
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
      Begin VB.CheckBox ConsistControlFunction8 
         Caption         =   "Function 8"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction7 
         Caption         =   "Function 7"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction6 
         Caption         =   "Function 6"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction5 
         Caption         =   "Function 5"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameLocomotiveNumber 
      Caption         =   "Locomotive Number or Consist"
      Height          =   1815
      Left            =   60
      TabIndex        =   12
      Top             =   840
      Width           =   4695
      Begin VB.TextBox TextBoxLocomotiveDecoderDim 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   53
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox TextBoxLocomotiveDecoderLight 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   52
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox CheckBoxAddressShortLocomotive 
         Caption         =   "short address (1 to 127) ?"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox ComboBoxLocomotiveNumber 
         Height          =   315
         ItemData        =   "MainlineOperationGUI.frx":0888
         Left            =   3720
         List            =   "MainlineOperationGUI.frx":088A
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   4440
         X2              =   240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label15 
         Caption         =   "Locomotive address is?"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Is this a"
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "What function is the dimmer mapped to?"
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label23 
         Caption         =   "What function is the light mapped to?"
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   1140
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function Controls - Group One"
      Height          =   1575
      Left            =   9960
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox ConsistControlFunction4 
         Caption         =   "Function 4"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction3 
         Caption         =   "Function 3"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction2 
         Caption         =   "Function 2"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction1 
         Caption         =   "Function 1"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction0 
         Caption         =   "Function 0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed && Direction Controls"
      Height          =   2895
      Left            =   9960
      TabIndex        =   0
      Top             =   5040
      Width           =   2535
      Begin VB.CheckBox ConsistControlDirectionN 
         Caption         =   "Neutral or"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1010
      End
      Begin VB.CheckBox ConsistControlDirectionR 
         Caption         =   "Reversed Direction"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox ConsistControlDirectionF 
         Caption         =   "Forward"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlSpeed128 
         Caption         =   "128 Step Sped Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox ConsistControlSpeed28 
         Caption         =   "28 Step or"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlSpeed14 
         Caption         =   "14 Step,"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.HScrollBar ConsistControlSpeed 
         Height          =   255
         Left            =   120
         Max             =   127
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label ConsistControlSpeedValue 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   20880
      Top             =   1140
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   20880
      Top             =   660
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   20880
      Top             =   120
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label44 
      Caption         =   "TimerWaterPressure"
      Height          =   375
      Left            =   14640
      TabIndex        =   81
      Top             =   9600
      Width           =   4455
   End
   Begin VB.Label Label43 
      Caption         =   "TimerWaterTemperature - For calculation the temperature of the water"
      Height          =   375
      Left            =   14640
      TabIndex        =   80
      Top             =   10080
      Width           =   4455
   End
   Begin VB.Label Label42 
      Caption         =   "TimerFuelConcumption - Used to calculate the fuel being used"
      Height          =   375
      Left            =   14760
      TabIndex        =   79
      Top             =   10560
      Width           =   3855
   End
   Begin VB.Label Label41 
      Caption         =   "TimerOilConsumption - Used to calculate the oil consumption"
      Height          =   375
      Left            =   9480
      TabIndex        =   78
      Top             =   12480
      Width           =   4455
   End
   Begin VB.Label Label40 
      Caption         =   "TimerWaterConsumption - Used to calculate the consumption of water"
      Height          =   375
      Left            =   9480
      TabIndex        =   77
      Top             =   12000
      Width           =   4455
   End
   Begin VB.Label Label39 
      Caption         =   "TimerSandConsumption - When snding switch is on, used to calculate the sand used. "
      Height          =   375
      Left            =   9480
      TabIndex        =   76
      Top             =   11520
      Width           =   4455
   End
   Begin VB.Label Label38 
      Caption         =   "TimerRadiatorFans - For updating the radiator fans either on or off based on the temperature"
      Height          =   375
      Left            =   9480
      TabIndex        =   75
      Top             =   11040
      Width           =   4455
   End
   Begin VB.Label Label37 
      Caption         =   "TimerDeadmann - To play wave file and display deadmann label."
      Height          =   375
      Left            =   14760
      TabIndex        =   74
      Top             =   12480
      Width           =   4455
   End
   Begin VB.Label Label36 
      Caption         =   "TimerRPM - For updating the RPM of the locomotive, traction effort and amperage with notching if error."
      Height          =   375
      Left            =   14760
      TabIndex        =   73
      Top             =   12000
      Width           =   4455
   End
   Begin VB.Label Label34 
      Caption         =   "Timer CabControl Signals - For count down distance between signals (always set to 1000) for 1 second intervals"
      Height          =   375
      Left            =   14760
      TabIndex        =   72
      Top             =   11040
      Width           =   4455
   End
   Begin VB.Label Label33 
      Caption         =   "TimerOilTemperature"
      Height          =   195
      Left            =   14640
      TabIndex        =   71
      Top             =   9120
      Width           =   3135
   End
   Begin VB.Label Label32 
      Caption         =   "TimerOilPressure"
      Height          =   195
      Left            =   14640
      TabIndex        =   70
      Top             =   8640
      Width           =   3135
   End
   Begin VB.Label Label30 
      Caption         =   "TimerThrottleDelay"
      Height          =   195
      Left            =   14640
      TabIndex        =   69
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label Label29 
      Caption         =   "TimerAirIndependentBrake"
      Height          =   195
      Left            =   14640
      TabIndex        =   68
      Top             =   7680
      Width           =   3135
   End
   Begin VB.Label Label28 
      Caption         =   "TimerAirAutomaticBrake"
      Height          =   195
      Left            =   14640
      TabIndex        =   67
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label27 
      Caption         =   "TimerAirMainResevoir"
      Height          =   195
      Left            =   14640
      TabIndex        =   66
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Label Label26 
      Caption         =   "TimerSpeedChange - If Speed changes"
      Height          =   195
      Left            =   14640
      TabIndex        =   65
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label25 
      Caption         =   "TimerSendCommand - Sending Command to the Port"
      Height          =   195
      Left            =   14640
      TabIndex        =   64
      Top             =   5760
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   "Before starting the graphics user interface, make sure you have selected a locomotive or a consist with the proper settings."
      Height          =   495
      Left            =   960
      TabIndex        =   34
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "MainlineOperationGUI"
Attribute VB_Creatable = False
Attribute VB_Exposed = False







Private Sub ButtonClose_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Button Close, Click" & vbCrLf
    End If ' Debug Tag

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
            If TemporaryScreen = "Mainline Operation GUI Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
            ElseIf TemporaryScreen = "Communication Setting Screen" Then
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
                Let Ini.Value = "Mainline Operation GUI Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Operation GUI Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Button Close, Click" & vbCrLf
    End If ' Debug Tag
    
    End Sub

Private Sub ButtonGUI_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Button GUI, Click" & vbCrLf
    End If ' Debug Tag

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
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If Screen.Width / Screen.TwipsPerPixelX >= 1024 And Screen.Height / Screen.TwipsPerPixelY >= 768 Then
    MainlineOperationGUI.Hide
    If OptionGuiDiesel1.Value = True Then
        Load MainlineOperationGuiDiesel1Screen
        MainlineOperationGuiDiesel1Screen.Show vbModeless
    ElseIf OptionGuiDiesel2.Value = True Then
        Load MainlineOperationGuiDiesel2Screen
        MainlineOperationGuiDiesel2Screen.Show vbModeless
    ElseIf OptionGuiDiesel3.Value = True Then
        Load MainlineOperationGuiDiesel3Screen
        MainlineOperationGuiDiesel3Screen.Show vbModeless
    ElseIf OptionGuiSteam1.Value = True Then
        Load MainlineOperationGuiSteam1Screen
        MainlineOperationGuiSteam1Screen.Show vbModeless
    ElseIf OptionGuiElectric1.Value = True Then
        Load MainlineOperationGuiElectric1Screen
        MainlineOperationGuiElectric1Screen.Show vbModeless
    End If
End If ' Screen size

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Button GUI, Click" & vbCrLf
    End If ' Debug Tag

End Sub



Private Sub ButtonInternet_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Button Internet, Click" & vbCrLf
    End If ' Debug Tag
 
    Load InternetSettings
    InternetSettings.Show vbModeless
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Button Internet, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub ButtonPrint_Click()

    MainlineOperationGUI.PrintForm
    
End Sub

Private Sub ButtonVideoSettings_Click()

    Load VideoSettings
    VideoSettings.Show vbModeless
    
End Sub

Private Sub CheckBoxAddressShortLocomotive_Click()
    
    If CheckBoxAddressShortLocomotive.Value = vbChecked Then
        Let LocomotiveDatabase.RecordSource = "LocomotiveDatabaseShort"
        LocomotiveDatabase.Refresh
        LocomotiveDatabase.Recordset.MoveFirst
    ElseIf CheckBoxAddressShortLocomotive.Value = vbUnchecked Then
        Let LocomotiveDatabase.RecordSource = "LocomotiveDatabaseLong"
        LocomotiveDatabase.Refresh
        LocomotiveDatabase.Recordset.MoveFirst
    End If
    
    Timer1.Interval = 1
    
End Sub

Private Sub CheckBoxaddressshortsounddecoder_Click()
    
    If CheckBoxAddressShortSoundDecoder.Value = vbChecked Then
        Let SoundDecoderDatabase.RecordSource = "LocomotiveDatabaseShort"
        SoundDecoderDatabase.Refresh
        SoundDecoderDatabase.Recordset.MoveFirst
    Else
        Let SoundDecoderDatabase.RecordSource = "LocomotiveDatabaseLong"
        SoundDecoderDatabase.Refresh
        SoundDecoderDatabase.Recordset.MoveFirst
    End If
    
    Timer2.Interval = 1

End Sub








Private Sub checkboxqsi_Click()

    If checkboxqsi.Value = vbChecked Then
        CheckboxSoundtraxxDSX.Value = vbUnchecked
    Else
        CheckboxSoundtraxxDSX.Value = vbChecked
    End If

End Sub

Private Sub Checkboxsounddecoderequipped_Click()
    
    If CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        Let LabelMappedSoundDecoder.Enabled = True
        ComboBoxSoundDecoderNumber.Enabled = True
        Let Label13.Enabled = True
        Let CheckBoxAddressShortSoundDecoder.Enabled = True
        TextBoxMappedFunctionBell.Enabled = True
        TextBoxMappedFunctionHorn.Enabled = True
        TextBoxMappedFunctionCoupler.Enabled = True
        TextBoxMappedFunctionDynamicBrake.Enabled = True
        TextBoxMappedFunctionRPMPlus.Enabled = True
        TextBoxMappedFunctionRPMminus.Enabled = True
        TextBoxMappedFunctionMute.Enabled = True
        TextboxMappedFunctionExtra1.Enabled = True
        TextboxMappedFunctionExtra2.Enabled = True
        TextboxMappedFunctionExtra3.Enabled = True
        TextboxMappedFunctionExtra4.Enabled = True
        CheckboxSoundtraxxDSX.Enabled = True
        checkboxqsi.Enabled = True
    Else
        Let LabelMappedSoundDecoder.Enabled = False
        ComboBoxSoundDecoderNumber.Enabled = False
        Let Label13.Enabled = False
        Let CheckBoxAddressShortSoundDecoder.Enabled = False
        TextBoxMappedFunctionBell.Enabled = False
        TextBoxMappedFunctionHorn.Enabled = False
        TextBoxMappedFunctionCoupler.Enabled = False
        TextBoxMappedFunctionDynamicBrake.Enabled = False
        TextBoxMappedFunctionRPMPlus.Enabled = False
        TextBoxMappedFunctionRPMminus.Enabled = False
        TextBoxMappedFunctionMute.Enabled = False
        TextboxMappedFunctionExtra1.Enabled = False
        TextboxMappedFunctionExtra2.Enabled = False
        TextboxMappedFunctionExtra3.Enabled = False
        TextboxMappedFunctionExtra4.Enabled = False
        CheckboxSoundtraxxDSX.Enabled = False
        checkboxqsi.Enabled = False
    End If

End Sub

Private Sub CheckboxSoundtraxxDSX_Click()

    If CheckboxSoundtraxxDSX.Value = vbChecked Then
        checkboxqsi.Value = vbUnchecked
    Else
        checkboxqsi.Value = vbChecked
    End If
    
End Sub

Private Sub ComboBoxLocomotiveNumber_Change()
    
    If Val(ComboBoxLocomotiveNumber.Text) < 1 Or _
       (Val(ComboBoxLocomotiveNumber.Text) > 127 And CheckBoxAddressShortLocomotive.Value = vbChecked) Or _
       (Val(ComboBoxLocomotiveNumber.Text) > 9999 And CheckBoxAddressShortLocomotive.Value = vbUnchecked) Then
       
        Let OptionVirtualCabSignalNo.Enabled = False
        Let OptionVirtualcabSignalYes.Enabled = False
        Let OptionSelfLapingBrakesNo.Enabled = False
        Let OptionSelfLapingBrakesYes.Enabled = False
        Let CheckBoxSoundDecoderEquipped.Enabled = False
        Let OptionGuiElectric1.Enabled = False
        Let OptionGuiSteam1.Enabled = False
        Let OptionGuiDiesel3.Enabled = False
        Let OptionGuiDiesel2.Enabled = False
        Let OptionGuiDiesel1.Enabled = False
        Let TextBoxLocomotiveDecoderDim.Enabled = False
        Let TextBoxLocomotiveDecoderLight.Enabled = False
        Let ButtonGUI.Enabled = False
        Let Label15.Enabled = False
        Let Label23.Enabled = False
        Let Label24.Enabled = False
                
        Let ComboBoxLocomotiveNumber.ListIndex = 1

        Let TemporaryPrompt = "You need to select a locomotive number or" & vbCrLf & "a consist number from one to nine thousand" & vbCrLf & "nine hundred and ninety nine. (Consist" & vbCrLf & "values or short addresses are vales between" & vbCrLf & "one and one hundred and twenty-seven)."
        MsgBox TemporaryPrompt, vbOKOnly + vbExclamation, "ATC - User Error"
                
    Else
        Let OptionVirtualCabSignalNo.Enabled = True
        Let OptionVirtualcabSignalYes.Enabled = True
        Let OptionSelfLapingBrakesNo.Enabled = True
        Let OptionSelfLapingBrakesYes.Enabled = True
        Let CheckBoxSoundDecoderEquipped.Enabled = True
        Let OptionGuiElectric1.Enabled = True
        Let OptionGuiSteam1.Enabled = True
        Let OptionGuiDiesel3.Enabled = True
        Let OptionGuiDiesel2.Enabled = True
        Let OptionGuiDiesel1.Enabled = True
        Let TextBoxLocomotiveDecoderDim.Enabled = True
        Let TextBoxLocomotiveDecoderLight.Enabled = True
        Let ButtonGUI.Enabled = True
        Let Label15.Enabled = True
        Let Label23.Enabled = True
        Let Label24.Enabled = True
    End If

End Sub

Private Sub ComboBoxSoundDecoderNumber_Change()

    If Val(ComboBoxSoundDecoderNumber.Text) < 1 Or _
       (Val(ComboBoxSoundDecoderNumber.Text) > 127 And CheckBoxAddressShortSoundDecoder.Value = vbChecked) Or _
       (Val(ComboBoxSoundDecoderNumber.Text) > 9999 And CheckBoxAddressShortSoundDecoder.Value = vbUnchecked) Then
       
        Let OptionVirtualCabSignalNo.Enabled = False
        Let OptionVirtualcabSignalYes.Enabled = False
        Let OptionSelfLapingBrakesNo.Enabled = False
        Let OptionSelfLapingBrakesYes.Enabled = False
        Let CheckBoxSoundDecoderEquipped.Enabled = False
        Let OptionGuiElectric1.Enabled = False
        Let OptionGuiSteam1.Enabled = False
        Let OptionGuiDiesel3.Enabled = False
        Let OptionGuiDiesel2.Enabled = False
        Let OptionGuiDiesel1.Enabled = False
        Let TextBoxLocomotiveDecoderDim.Enabled = False
        Let TextBoxLocomotiveDecoderLight.Enabled = False
        Let ButtonGUI.Enabled = False
        Let Label15.Enabled = False
        Let Label23.Enabled = False
        Let Label24.Enabled = False
                
        Let ComboBoxSoundDecoderNumber.ListIndex = 1

        Let TemporaryPrompt = "You need to select a locomotive number or" & vbCrLf & "a consist number from one to nine thousand" & vbCrLf & "nine hundred and ninety nine. (Consist" & vbCrLf & "values or short addresses are vales between" & vbCrLf & "one and one hundred and twenty-seven)."
        MsgBox TemporaryPrompt, vbOKOnly + vbExclamation, "ATC - User Error"
                
    Else
        Let OptionVirtualCabSignalNo.Enabled = True
        Let OptionVirtualcabSignalYes.Enabled = True
        Let OptionSelfLapingBrakesNo.Enabled = True
        Let OptionSelfLapingBrakesYes.Enabled = True
        Let CheckBoxSoundDecoderEquipped.Enabled = True
        Let OptionGuiElectric1.Enabled = True
        Let OptionGuiSteam1.Enabled = True
        Let OptionGuiDiesel3.Enabled = True
        Let OptionGuiDiesel2.Enabled = True
        Let OptionGuiDiesel1.Enabled = True
        Let TextBoxLocomotiveDecoderDim.Enabled = True
        Let TextBoxLocomotiveDecoderLight.Enabled = True
        Let ButtonGUI.Enabled = True
        Let Label15.Enabled = True
        Let Label23.Enabled = True
        Let Label24.Enabled = True
    End If

End Sub

Private Sub ConsistControlCVRead_Click()

If ConsistControlCVRead.Value = vbChecked Then
    Let ConsistControlCVWright.Value = vbUnchecked
Else
    Let ConsistControlCVWright.Value = vbChecked
End If

End Sub

Private Sub ConsistControlCVWright_Click()

If ConsistControlCVWright.Value = vbChecked Then
    Let ConsistControlCVRead.Value = vbUnchecked
Else
    Let ConsistControlCVRead.Value = vbChecked
End If

End Sub




Private Sub ConsistControlSpeed_Change()

Let ConsistControlSpeedValue.Caption = ConsistControlSpeed.Value

End Sub

Private Sub ConsistControlSpeed128_Click()

If ConsistControlSpeed128.Value = vbChecked Then
    Let ConsistControlSpeed28.Value = vbUnchecked
    Let ConsistControlSpeed14.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 127
Let ConsistControlSpeed.Value = 0

End Sub

Private Sub ConsistControlSpeed14_Click()

If ConsistControlSpeed14.Value = vbChecked Then
    Let ConsistControlSpeed28.Value = vbUnchecked
    Let ConsistControlSpeed128.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 15
Let ConsistControlSpeed.Value = 0



End Sub

Private Sub ConsistControlSpeed28_Click()

If ConsistControlSpeed28.Value = vbChecked Then
    Let ConsistControlSpeed14.Value = vbUnchecked
    Let ConsistControlSpeed128.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 28
Let ConsistControlSpeed.Value = 0

End Sub









Private Sub Form_Activate()

    DoEvents
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Form, Activate" & vbCrLf
    End If ' Debug Tag

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
        If TemporaryScreen = "Mainline Operation GUI Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Operation GUI Screen"
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
            Let Ini.Value = "Mainline Operation GUI Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Operation GUI Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Form, Activate" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub Form_Deactivate()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Form, Deactivate" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationGUI.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationGUI.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationGUI.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationGUI.Height)

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Replace Settings
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Parameter = "ComboBoxLocomotiveNumber"
    Let Ini.Value = ComboBoxLocomotiveNumber.Text
    Let Ini.Parameter = "CheckBoxAddressShortLocomotive"
    Let Ini.Value = CheckBoxAddressShortLocomotive.Value
    Let Ini.Parameter = "TextBoxLocomotiveDecoderLight"
    Let Ini.Value = TextBoxLocomotiveDecoderLight.Text
    Let Ini.Parameter = "TextBoxLocomotiveDecoderDim"
    Let Ini.Value = TextBoxLocomotiveDecoderDim.Text
    Let Ini.Parameter = "CheckBoxSoundDecoderEquipped"
    Let Ini.Value = CheckBoxSoundDecoderEquipped.Value
    Let Ini.Parameter = "ComboBoxSoundDecoderNumber"
    Let Ini.Value = ComboBoxSoundDecoderNumber.Text
    Let Ini.Parameter = "CheckBoxAddressShortSoundDecoder"
    Let Ini.Value = CheckBoxAddressShortSoundDecoder.Value
    
    Let Ini.Parameter = "CheckboxSoundtraxxDSX"
    Let Ini.Value = CheckboxSoundtraxxDSX.Value
    Let Ini.Parameter = "CheckboxQSI"
    Let Ini.Value = checkboxqsi.Value
    Let Ini.Parameter = "CheckboxDigitraxFX"
    Let Ini.Value = CheckBoxDigitraxFX.Value
   
    Let Ini.Parameter = "TextBoxMappedFunctionBell"
    Let Ini.Value = TextBoxMappedFunctionBell.Text
    Let Ini.Parameter = "TextBoxMappedFunctionHorn"
    Let Ini.Value = TextBoxMappedFunctionHorn.Text
    Let Ini.Parameter = "TextBoxMappedFunctionCoupler"
    Let Ini.Value = TextBoxMappedFunctionCoupler.Text
    Let Ini.Parameter = "TextBoxMappedFunctionDynamicBrake"
    Let Ini.Value = TextBoxMappedFunctionDynamicBrake.Text
    Let Ini.Parameter = "TextBoxMappedFunctionRPMPlus"
    Let Ini.Value = TextBoxMappedFunctionRPMPlus.Text
    Let Ini.Parameter = "TextBoxMappedFunctionRPMMinus"
    Let Ini.Value = TextBoxMappedFunctionRPMminus.Text
    Let Ini.Parameter = "TextBoxMappedFunctionMute"
    Let Ini.Value = TextBoxMappedFunctionMute.Text
    Let Ini.Parameter = "TextboxMappedFunctionExtra1"
    Let Ini.Value = TextboxMappedFunctionExtra1.Text
    Let Ini.Parameter = "TextboxMappedFunctionExtra2"
    Let Ini.Value = TextboxMappedFunctionExtra2.Text
    Let Ini.Parameter = "TextboxMappedFunctionExtra3"
    Let Ini.Value = TextboxMappedFunctionExtra3.Text
    Let Ini.Parameter = "TextboxMappedFunctionExtra4"
    Let Ini.Value = TextboxMappedFunctionExtra4.Text
    Let Ini.Parameter = "OptionGuiDiesel1"
    Let Ini.Value = OptionGuiDiesel1.Value
    Let Ini.Parameter = "OptionGuiDiesel2"
    Let Ini.Value = OptionGuiDiesel2.Value
    Let Ini.Parameter = "OptionGuiDiesel3"
    Let Ini.Value = OptionGuiDiesel3.Value
    Let Ini.Parameter = "OptionGuiSteam1"
    Let Ini.Value = OptionGuiSteam1.Value
    Let Ini.Parameter = "OptionGuiSteam1"
    Let Ini.Value = OptionGuiSteam1.Value
    Let Ini.Parameter = "OptionGuiElectric1"
    Let Ini.Value = OptionGuiElectric1.Value
    Let Ini.Parameter = "OptionSelfLapingBrakesYes"
    Let Ini.Value = OptionSelfLapingBrakesYes.Value
    Let Ini.Parameter = "OptionSelfLapingBrakesNo"
    Let Ini.Value = OptionSelfLapingBrakesNo.Value
    Let Ini.Parameter = "OptionVirtualcabSignalYes"
    Let Ini.Value = OptionVirtualcabSignalYes.Value
    Let Ini.Parameter = "OptionVirtualCabSignalNo"
    Let Ini.Value = OptionVirtualCabSignalNo.Value

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
                Let Ini.Value = "Mainline Operation GUI Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineOperationGUI.Hide
    'unload MainlineOperationGUI

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Form, Deactivate" & vbCrLf
    End If ' Debug Tag
  
End Sub


Private Sub Form_Load()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Form, Load" & vbCrLf
    End If ' Debug Tag

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
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LocomotiveDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
    Let SoundDecoderDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialization of Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Recall Settings
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Let Ini.Filename = App.Path$ & "\Atc.ini"
'   Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "CheckBoxAddressShortLocomotive"
    Let CheckBoxAddressShortLocomotive.Value = Val(Ini.Value)
    Let Ini.Parameter = "ComboBoxLocomotiveNumber"
    Let ComboBoxLocomotiveNumber.Text = Ini.Value
    Let Ini.Parameter = "TextBoxLocomotiveDecoderLight"
    Let TextBoxLocomotiveDecoderLight.Text = Ini.Value
    Let Ini.Parameter = "TextBoxLocomotiveDecoderDim"
    Let TextBoxLocomotiveDecoderDim.Text = Ini.Value
    Let Ini.Parameter = "CheckBoxSoundDecoderEquipped"
    Let CheckBoxSoundDecoderEquipped.Value = Val(Ini.Value)
    Let Ini.Parameter = "ComboBoxSoundDecoderNumber"
    Let ComboBoxSoundDecoderNumber.Text = Ini.Value
    Let Ini.Parameter = "CheckBoxAddressShortSoundDecoder"
    Let CheckBoxAddressShortSoundDecoder.Value = Val(Ini.Value)
    
    Let Ini.Parameter = "CheckboxSoundtraxxDSX"
    Let CheckboxSoundtraxxDSX.Value = Val(Ini.Value)
    Let Ini.Parameter = "CheckboxQSI"
    Let checkboxqsi.Value = Val(Ini.Value)
    Let Ini.Parameter = "CheckboxDigitraxFX"
    Let CheckBoxDigitraxFX.Value = Val(Ini.Value)
        
    Let Ini.Parameter = "TextBoxMappedFunctionBell"
    Let TextBoxMappedFunctionBell.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionHorn"
    Let TextBoxMappedFunctionHorn.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionCoupler"
    Let TextBoxMappedFunctionCoupler.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionDynamicBrake"
    Let TextBoxMappedFunctionDynamicBrake.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionRPMPlus"
    Let TextBoxMappedFunctionRPMPlus.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionRPMMinus"
    Let TextBoxMappedFunctionRPMminus.Text = Ini.Value
    Let Ini.Parameter = "TextBoxMappedFunctionMute"
    Let TextBoxMappedFunctionMute.Text = Ini.Value
    Let Ini.Parameter = "TextboxMappedFunctionExtra1"
    Let TextboxMappedFunctionExtra1.Text = Ini.Value
    Let Ini.Parameter = "TextboxMappedFunctionExtra2"
    Let TextboxMappedFunctionExtra2.Text = Ini.Value
    Let Ini.Parameter = "TextboxMappedFunctionExtra3"
    Let TextboxMappedFunctionExtra3.Text = Ini.Value
    Let Ini.Parameter = "TextboxMappedFunctionExtra4"
    Let TextboxMappedFunctionExtra4.Text = Ini.Value
    Let Ini.Parameter = "OptionGuiDiesel1"
    Let OptionGuiDiesel1.Value = Ini.Value
    Let Ini.Parameter = "OptionGuiDiesel2"
    Let OptionGuiDiesel2.Value = Ini.Value
    Let Ini.Parameter = "OptionGuiDiesel3"
    Let OptionGuiDiesel3.Value = Ini.Value
    Let Ini.Parameter = "OptionGuiSteam1"
    Let OptionGuiSteam1.Value = Ini.Value
    Let Ini.Parameter = "OptionGuiSteam1"
    Let OptionGuiSteam1.Value = Ini.Value
    Let Ini.Parameter = "OptionGuiElectric1"
    Let OptionGuiElectric1.Value = Ini.Value
    Let Ini.Parameter = "OptionSelfLapingBrakesYes"
    Let OptionSelfLapingBrakesYes.Value = Ini.Value
    Let Ini.Parameter = "OptionSelfLapingBrakesNo"
    Let OptionSelfLapingBrakesNo.Value = Ini.Value
    Let Ini.Parameter = "OptionVirtualcabSignalYes"
    Let OptionVirtualcabSignalYes.Value = Ini.Value
    Let Ini.Parameter = "OptionVirtualCabSignalNo"
    Let OptionVirtualCabSignalNo.Value = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        MainlineOperationGUI.Left = (Screen.Width - Width) / 2
        MainlineOperationGUI.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineOperationGUI.Width > Screen.Width Then
            Let MainlineOperationGUI.Left = Screen.Width - MainlineOperationGUI.Width
        Else
            Let MainlineOperationGUI.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineOperationGUI.Height > Screen.Height Then
            Let MainlineOperationGUI.Top = Screen.Height - MainlineOperationGUI.Height
        Else
            Let MainlineOperationGUI.Top = Val(TemporaryValueTop)
        End If
    End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Form, Load" & vbCrLf
    End If ' Debug Tag

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub




Private Sub Form_Resize()

    If MainlineOperationGUI.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineOperationGUI.WindowState = vbNormal Then
    
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

Private Sub OptionGuiDiesel1_Click()

Let OptionSelfLapingBrakesYes.Enabled = True
Let OptionSelfLapingBrakesNo.Enabled = True
Let OptionVirtualcabSignalYes.Enabled = True
Let OptionVirtualCabSignalNo.Enabled = True

End Sub

Private Sub OptionGuiDiesel2_Click()

Let OptionSelfLapingBrakesYes.Enabled = False
Let OptionSelfLapingBrakesNo.Enabled = False
Let OptionVirtualcabSignalYes.Enabled = False
Let OptionVirtualCabSignalNo.Enabled = False

End Sub

Private Sub OptionGuiDiesel3_Click()

Let OptionSelfLapingBrakesYes.Enabled = False
Let OptionSelfLapingBrakesNo.Enabled = False
Let OptionVirtualcabSignalYes.Enabled = False
Let OptionVirtualCabSignalNo.Enabled = False

End Sub

Private Sub OptionGuiElectric1_Click()

Let OptionSelfLapingBrakesYes.Enabled = False
Let OptionSelfLapingBrakesNo.Enabled = False
Let OptionVirtualcabSignalYes.Enabled = True
Let OptionVirtualCabSignalNo.Enabled = True

End Sub

Private Sub OptionGuiSteam1_Click()

Let OptionSelfLapingBrakesYes.Enabled = False
Let OptionSelfLapingBrakesNo.Enabled = False
Let OptionVirtualcabSignalYes.Enabled = False
Let OptionVirtualCabSignalNo.Enabled = False

End Sub

Private Sub OptionGuiSteam2_Click()

Let OptionSelfLapingBrakesYes.Enabled = False
Let OptionSelfLapingBrakesNo.Enabled = False
Let OptionVirtualcabSignalYes.Enabled = False
Let OptionVirtualCabSignalNo.Enabled = False

End Sub




Private Sub TextBoxLocomotiveDecoderDim_LostFocus()

If Val(TextBoxLocomotiveDecoderDim.Text) < 0 Or _
   Val(TextBoxLocomotiveDecoderDim.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and nine. Most decoders are programmed for a value of four.", vbOKOnly, "Automatic Train Control - Outdated version"
Let TextBoxLocomotiveDecoderDim.Text = 4

End If


End Sub


Private Sub TextBoxLocomotiveDecoderLight_LostFocus()

If Val(TextBoxLocomotiveDecoderLight.Text) < 0 Or _
   Val(TextBoxLocomotiveDecoderLight.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and nine. Most decoders are programmed for a value of zero.", vbOKOnly, "Automatic Train Control - Outdated version"
    Let TextBoxLocomotiveDecoderLight.Text = 0
End If


End Sub




Private Sub TextBoxMappedFunctionBell_LostFocus()

If Val(TextBoxMappedFunctionBell.Text) < 1 Or _
   Val(TextBoxMappedFunctionBell.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders use a default value of one.", vbOKOnly, "Automatic Train Control - Function Value"
    Let TextBoxMappedFunctionBell.Text = 1
End If


End Sub


Private Sub TextBoxMappedFunctionCoupler_LostFocus()

If Val(TextBoxMappedFunctionCoupler.Text) < 1 Or _
   Val(TextBoxMappedFunctionCoupler.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders use a default value of three.", vbOKOnly, "Automatic Train Control - Function Value"
    Let TextBoxMappedFunctionCoupler.Text = 3
End If

End Sub


Private Sub TextBoxMappedFunctionDynamicBrake_LostFocus()

If Val(TextBoxMappedFunctionDynamicBrake.Text) < 1 Or _
   Val(TextBoxMappedFunctionDynamicBrake.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders default to value of four.", vbOKOnly, "Automatic Train Control - Function Value"
    Let TextBoxMappedFunctionDynamicBrake.Text = 4
End If

End Sub



Private Sub TextBoxMappedFunctionHorn_LostFocus()

If Val(TextBoxMappedFunctionHorn.Text) < 1 Or _
   Val(TextBoxMappedFunctionHorn.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders use a default value of two.", vbOKOnly, "Automatic Train Control - Function Value"
    Let TextBoxMappedFunctionHorn.Text = 2
End If

End Sub


Private Sub TextBoxMappedFunctionMute_LostFocus()

If Val(TextBoxMappedFunctionMute.Text) < 1 Or _
   Val(TextBoxMappedFunctionMute.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders used a value of eight.", vbOKOnly, "Automatic Train Control - Function Value"
    Let TextBoxMappedFunctionMute.Text = 8
End If

End Sub




Private Sub TextBoxMappedFunctionRPMMinus_LostFocus()

If Val(TextBoxMappedFunctionRPMminus.Text) < 1 Or _
   Val(TextBoxMappedFunctionRPMminus.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders use a value of six.", vbOKOnly, "Automatic Train Control - Function Value"

Let TextBoxMappedFunctionRPMminus.Text = 6
End If

End Sub


Private Sub TextBoxMappedFunctionRPMPlus_LostFocus()

If Val(TextBoxMappedFunctionRPMPlus.Text) < 1 Or _
   Val(TextBoxMappedFunctionRPMPlus.Text) > 12 Then
    MsgBox "Please enter a valid function button number" + Chr$(13) + "between one and twelve. Most decoders use a value of five.", vbOKOnly, "Automatic Train Control - Function Value"

    Let TextBoxMappedFunctionRPMPlus.Text = 5

End If

End Sub








Private Sub Timer1_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer1, Activate" & vbCrLf
    End If ' Debug Tag
    
    Let ButtonClose.Enabled = False
    Let Timer1.Interval = 0
    Let MousePointer = vbArrowHourglass
    
    ComboBoxLocomotiveNumber.Clear
    
    LocomotiveDatabase.Recordset.MoveFirst
    Do While Not LocomotiveDatabase.Recordset.EOF
        LocomotiveDatabase.Recordset.MoveNext
        If LocomotiveDecoderEquiped.Value = vbChecked Then
            ComboBoxLocomotiveNumber.AddItem Format(Val(LocomotiveNumber.Text), "0000")
        End If
        DoEvents
    Loop
    
    Let MousePointer = Default
    Let ButtonClose.Enabled = True
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "ComboBoxLocomotiveNumber"
    Let ComboBoxLocomotiveNumber.Text = Ini.Value
        
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer1, Dectivate" & vbCrLf
    End If ' Debug Tag

End Sub



Private Sub Timer2_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer2, Activate" & vbCrLf
    End If ' Debug Tag
    
    Let ButtonClose.Enabled = False
    Let Timer2.Interval = 0
    Let MousePointer = vbArrowHourglass
    
    ComboBoxSoundDecoderNumber.Clear
    
    SoundDecoderDatabase.Recordset.MoveFirst
    Do While Not SoundDecoderDatabase.Recordset.EOF
        SoundDecoderDatabase.Recordset.MoveNext
        If SoundDecoderEquiped.Value = vbChecked Then
            ComboBoxSoundDecoderNumber.AddItem Format(Val(SoundDecoderNumber.Text), "0000")
        End If
        DoEvents
    Loop
    
    Let MousePointer = Default
    Let ButtonClose.Enabled = True
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "ComboBoxSoundDecoderNumber"
    Let ComboBoxSoundDecoderNumber.Text = Ini.Value
       
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer2, Deactivate" & vbCrLf
    End If ' Debug Tag

End Sub


Private Sub TimerAirAutomaticBrake_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Air Autoamtic Brake, Timer" & vbCrLf
    End If ' Debug Tag
    
    Let tempPercentage = 1 - (Val(MainlineOperationGuiDiesel1Screen!PictureBoxAutomaticBrake.Tag) / 16)
    
    If OptionSelfLapingBrakesYes.Value = True Then
        If Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) > Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) * tempPercentage Then
                Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value = Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) - 1
        ElseIf (MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) < Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) * tempPercentage Then
                Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value = Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) + 1
        End If
    Else
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxAutomaticBrake.Tag) = 0 Then
            Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value = Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) + 1
        ElseIf Val(MainlineOperationGuiDiesel1Screen!PictureBoxAutomaticBrake.Tag) > 1 Then
            Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value = Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) - 1
        End If
    End If
    
    Let TimerAirAutomaticBrake.Interval = Val(MainlineOperationGuiDiesel1Screen!TextboxSettingAutomaticAirBrake.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let TimerSendCommand.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingsendcommandInterval.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let TimerSpeedChange.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingSpeedChangeInterval.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)
    
    
    
    If Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) > 50 Then
        Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.BorderColor = 49433
    Else
        Let MainlineOperationGuiDiesel1Screen!barairautomaticbrake.BorderColor = 192
    End If
    
    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusAutomaticBrakePressure.Text = (Val(MainlineOperationGuiDiesel1Screen!barairautomaticbrake.Value) / Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value)) ^ 2
    
    ' check this out later
    
    Let temporary1 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusAutomaticBrakePressure.Text)
    Let temporary2 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusIndependentBrakePressure.Text)
    
    If temporary1 < temporary2 Then
        Let Temporary = temporary1
    Else
        Let Temporary = temporary2
    End If
    
    If Temporary > 1 Then Let Temporary = 1
    If Temporary < 0 Then Let temportay = 0
    
    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified.Text = Int(Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) * Temporary)

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Air Automatic Brake, Timer" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub TimerAirIndependentBrake_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Air Indepednet Brake, Timer" & vbCrLf
    End If ' Debug Tag

    
Let tempPercentage = 1 - (Val(MainlineOperationGuiDiesel1Screen!TransPictureBoxIndependentBrake.Tag) / 22)

If OptionSelfLapingBrakesYes.Value = True Then
    If Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) > Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) * tempPercentage Then
            Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value = Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) - 1
    ElseIf (MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) < Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) * tempPercentage Then
            Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value = Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) + 1
    End If
Else
    If Val(MainlineOperationGuiDiesel1Screen!TransPictureBoxIndependentBrake.Tag) = 0 Then
        Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value = Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) + 1
    ElseIf Val(MainlineOperationGuiDiesel1Screen!TransPictureBoxIndependentBrake.Tag) > 1 Then
        Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value = Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) - 1
    End If
End If

Let TimerAirAutomaticBrake.Interval = Val(MainlineOperationGuiDiesel1Screen!TextboxSettingAutomaticAirBrake.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)
Let TimerSendCommand.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingsendcommandInterval.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)
Let TimerSpeedChange.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingSpeedChangeInterval.Text) * tempPercentage * (Val(ScrollTimeAdjustment) / 100 + 1)


If Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) > 50 Then
    Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.BorderColor = 192
End If

Let MainlineOperationGuiDiesel1Screen!TextBoxStatusIndependentBrakePressure.Text = (Val(MainlineOperationGuiDiesel1Screen!BarAirIndependentBrake.Value) / Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value)) ^ 2


Let temporary1 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusAutomaticBrakePressure.Text)
Let temporary2 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusIndependentBrakePressure.Text)

If temporary1 < temporary2 Then
    Let Temporary = temporary1
Else
    Let Temporary = temporary2
End If

If Temporary > 1 Then Let Temporary = 1
If Temporary < 0 Then Let Temporary = 0

Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified.Text = Int(Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) * Temporary)

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Air Independent Brake, Timer" & vbCrLf
    End If ' Debug Tag
 
End Sub


Private Sub TimerAirMainReservoir_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Air Main Resevoir, Timer" & vbCrLf
    End If ' Debug Tag

    
    Let timerairmainReservoir.Interval = 0
    
    If MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value < 100 Then
        Let MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value = Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) + 1
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Notify = False
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Wait = False
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Shareable = False
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Filename = App.Path$ & "\Sounds\Graphics\compressor.wav"
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Command = "Open"
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Notify = True
        Let MainlineOperationGuiDiesel1Screen!SoundCompressor.Command = "Play"
    End If
    
    Let timerairmainReservoir.Interval = 3000 / (Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1) * (Val(ScrollTimeAdjustment) / 100 + 1)
    
    If Val(MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value) > 50 Then
        Let MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.BorderColor = 49433
    Else
        Let MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.BorderColor = 192
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Main Resevoir, Timer" & vbCrLf
    End If ' Debug Tag

End Sub





Private Sub timerdeadmann_Timer()

    Let MainlineOperationGuiDiesel1Screen!labeldeadmann.Visible = True
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Notify = False
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Wait = False
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Shareable = False
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Filename = App.Path$ & "\Sounds\Graphics\Deadman.wav"
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Command = "Open"
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Notify = True
    Let MainlineOperationGuiDiesel1Screen!sounddeadmann.Command = "Play"
    Let timerdeadmann.Interval = 0

End Sub

Private Sub timerfuelconsumption_Timer()

Let MainlineOperationGuiDiesel1Screen!TextBoxFuelTank.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxFuelTank.Text) - 0.02 * (1 + Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) * (Val(ScrollTimeAdjustment) / 100 + 1))
Let MainlineOperationGuiDiesel1Screen!BarFuelTank.Value = Val(MainlineOperationGuiDiesel1Screen!TextBoxFuelTank.Text)

If Val(MainlineOperationGuiDiesel1Screen!BarFuelTank.Value) > Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusFuelTooLow.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarFuelTank.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarFuelTank.BorderColor = 192
End If

End Sub



Private Sub TimerOilConsumption_Timer()

Let MainlineOperationGuiDiesel1Screen!TextBoxOilTank.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxOilTank.Text) - 0.0001
Let MainlineOperationGuiDiesel1Screen!BarOilTank.Value = Val(MainlineOperationGuiDiesel1Screen!TextBoxOilTank.Text)

If Val(MainlineOperationGuiDiesel1Screen!BarOilTank.Value) > Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusOilTooLow.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarOilTank.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarOilTank.BorderColor = 192
End If

End Sub

Private Sub TimerOilPressure_Timer()
 
If Val(MainlineOperationGuiDiesel1Screen!BarOilPressure.Value) < 75 Then
    Let TimerOilPressure.Interval = Int(((Val(MainlineOperationGuiDiesel1Screen!BarOilPressure.Value) / 75) ^ 2) * (6000 / 1 + Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag))) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarOilPressure.Value = Val(MainlineOperationGuiDiesel1Screen!BarOilPressure.Value) + 1
End If

If Val(MainlineOperationGuiDiesel1Screen!BarOilPressure.Value) > 80 Or _
   Val(MainlineOperationGuiDiesel1Screen!BarOilPressure.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusOilTooLowPressure.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarOilPressure.BorderColor = 192
Else
    Let MainlineOperationGuiDiesel1Screen!BarOilPressure.BorderColor = 49433
End If
 
End Sub

Private Sub TimerOilTemperature_Timer()
 
If Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) < 182 Then
    Let TimerOilTemperature.Interval = Int(((Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) / 182) ^ 2) * (30000 / (1 + Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag)))) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) + 1
Else
    Let TimerOilTemperature.Interval = Int(((Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) / 182) ^ 2) * 30000) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) - 1
End If

If Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) > 220 Or _
   Val(MainlineOperationGuiDiesel1Screen!BarOilTemperature.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusOilTooLowTemperature.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarOilTemperature.BorderColor = 192
Else
    Let MainlineOperationGuiDiesel1Screen!BarOilTemperature.BorderColor = 49433
End If

End Sub
Private Sub TimerRadiatorFans_Timer()

If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) > 190 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans1.Text = "On"
If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) > 200 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans2.Text = "On"
If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) > 210 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans3.Text = "On"

If MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans3.Text = "On" Then
    Let TimerRadiatorFans.Interval = Int(((180 / Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value)) ^ 2) * 600) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) - 1
    If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) < 202 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans3.Text = "Off"
ElseIf MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans2.Text = "On" Then
    Let TimerRadiatorFans.Interval = Int(((180 / Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value)) ^ 2) * 1000) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) - 1
    If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) < 192 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans2.Text = "Off"
ElseIf MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans1.Text = "On" Then
    Let TimerRadiatorFans.Interval = Int(((180 / Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value)) ^ 2) * 1500) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) - 1
    If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) < 172 Then Let MainlineOperationGuiDiesel1Screen!TextBoxRadiatorFans1.Text = "Off"
End If

End Sub

Private Sub TimerRPM_Timer()

If Val(MainlineOperationGuiDiesel1Screen!barrpm.Value) < 200 * Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 200 Then
    Let MainlineOperationGuiDiesel1Screen!barrpm.Value = Val(MainlineOperationGuiDiesel1Screen!barrpm.Value) + 10
Else
    Let MainlineOperationGuiDiesel1Screen!barrpm.Value = Val(MainlineOperationGuiDiesel1Screen!barrpm.Value) - 10
End If

Let temp1 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified.Text)
If temp1 <= 10 Then temp1 = 10
Let temp2 = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep.Text)
If temp2 = 0 Then temp2 = 1
Let temp3 = Val(MainlineOperationGuiDiesel1Screen!barrpm.Value) / 1800
Let MainlineOperationGuiDiesel1Screen!BarAmpere.Value = Int((temp3 * 2500) * ((temp2 / temp1) * 0.25))

Let MainlineOperationGuiDiesel1Screen!BarTractionEffort.Value = Val(MainlineOperationGuiDiesel1Screen!BarAmpere.Value) / 1300 * 114000

If Val(MainlineOperationGuiDiesel1Screen!BarAmpere.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusAmpereTooHigh.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarAmpere.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarAmpere.BorderColor = 192
                 
    If MainlineOperationGuiDiesel1Screen!ButtonHelp.Caption = "Help is On" Then
        Let Message = "You have exceeded the maximum amperage rating of " + Str$(Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusAmpereTooHigh.Text))
        Let Message = Message + " amps alotted by the main gererator." + Chr$(13)
        Let Message = Message + "The throttle will automatiCally be reduced by one notch to correct operator error." + Chr$(13)
        MsgBox Message, vbExclamation, "Engineer Error"
    End If
                 
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
    End If
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
    End If
    
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Wait = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Shareable = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Open"
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = True
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Play"
            
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
    
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) >= 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 7 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
            End If

     
 
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

            
            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbnChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

        End If
            
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) = 0 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctiondynamic.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
    

            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

            
        End If
        
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < -1 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
            End If
            
    

            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If
   
    Let temp = App.Path
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let temp = temp + "\Graphics\Locomotive Diesel1\DynamicBrake"
    Else
        Let temp = temp + "\Graphics\Locomotive Diesel1\Throttle"
    End If
    Let temp = temp + Trim$(Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag))
    Let temp = temp + "(d1).bmp"

    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
  
End If

' -------------------------------------------------------------------------------------------------------------------------

If Val(MainlineOperationGuiDiesel1Screen!BarTractionEffort.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusTractionEffortTooHigh.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarTractionEffort.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarTractionEffort.BorderColor = 192
                 
    If MainlineOperationGuiDiesel1Screen!ButtonHelp.Caption = "&Help is On" Then
        Let Message = "You have exceeded the maximum traction effort of" + Str$(Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusTractionEffortTooHigh.Text))
        Let Message = Message + " pounds alotted by the locomotive and coupler." + Chr$(13)
        Let Message = Message + "The throttle will automatiCally be reduced by one notch to correct operator" + Chr$(13)
        Let Message = Message + "error."
        MsgBox Message, vbExclamation, "Engineer Error"
    End If
                 
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
    ElseIf Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
    End If
   
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Wait = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Shareable = False
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Open"
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = True
    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Play"
             
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) >= 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 7 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
             ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
             ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
            End If
            
    
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

                
            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
        End If
        
    

        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call SetSoundDecoderNumber
                Call SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call SetSoundDecoderNumber
                Call SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If

        
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) = 0 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctiondynamic.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
    

            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
            
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < -1 Then
        
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
    

            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call SetSoundDecoderNumber
                    Call SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
    
            
        End If
    End If
    
' ------------------------------------------------------------------------------------------------------------------------------------------------
' Updating the picture
'
    Let temp = App.Path
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let temp = temp + "\Graphics\Loomotive Diesel1\DynamicBrake"
    Else
        Let temp = temp + "\Graphics\Locomotive Diesel1\Throttle"
    End If
    Let temp = temp + Trim$(Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag))
    Let temp = temp + "(d1).bmp"

    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
        
End If

End Sub
Private Sub TimerSandConsumption_Timer()


Let MainlineOperationGuiDiesel1Screen!TextBoxSandTank.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxSandTank.Text) - 0.01
Let MainlineOperationGuiDiesel1Screen!BarSandTank.Value = Val(MainlineOperationGuiDiesel1Screen!TextBoxSandTank.Text)

If Val(MainlineOperationGuiDiesel1Screen!BarSandTank.Value) > Val(MainlineOperationGuiDiesel1Screen!textboxstatussandtoolow.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarSandTank.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarSandTank.BorderColor = 192
End If

End Sub

Private Sub TimerSendCommand_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Send Command, Timer" & vbCrLf
    End If ' Debug Tag


    If Val(ConsistControlSpeed.Value) <> Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified.Text) Then
        Let ConsistControlSpeed.Value = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified.Text)
        If Val(ConsistControlSpeed.Value) <> 1 Then
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------
                   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
                        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Middle Mainscreen, General, Send Command Via TempQ, sending command in host mode."
                    End If ' Debug Tag
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    Call SetLocomotiveNumber
                    Call SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
                        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainscreen, General, Send Command Via TempQ, sending command in stand-alone mode."
                    End If ' Debug Tag
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    Call SetLocomotiveNumber
                    Call SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Send Command, Timer" & vbCrLf
    End If ' Debug Tag

End Sub





Private Sub TimerSpeedChange_Timer()
    

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Speed Change, Timer" & vbCrLf
    End If ' Debug Tag
    
    If Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep.Text) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) Then
        Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) - 1
    Else
        If Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep.Text) > Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) Then
            Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow.Text) + 1
        End If
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Speed Change, Timer" & vbCrLf
    End If ' Debug Tag
    
End Sub



Private Sub timerthrottledelay_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Throttle Delay, Timer" & vbCrLf
    End If ' Debug Tag
     
    Let timerthrottledelay.Interval = 0
    
    If Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusThrottleDelay.Text) > 0 Then
        Let MainlineOperationGuiDiesel1Screen!TextBoxStatusThrottleDelay.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusThrottleDelay.Text) - 1
    End If
    
    Let timerthrottledelay.Interval = 1000 * (Val(ScrollTimeAdjustment) / 100 + 1)
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI, Timer Throttle Delay, Timer" & vbCrLf
    End If ' Debug Tag

End Sub


Private Sub TimerVirtualCabSignal_Timer()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Virtual Signalling Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If OptionVirtualcabSignalYes.Value = True Then
    If Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text) <> 0 Then
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxReverser.Tag) = 0 Then
            Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text) - (102.66 * Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text) / 126)
        Else
            If MainlineOperationGuiDiesel1Screen!PictureBoxReverser.Tag = 2 Then
                Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text) + (102.66 * Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text) / 126)
            End If
        End If
    End If

    If Val(MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text) < 0 Then
    
        If Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text) <> 0 Then
            If MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching = "Stop" Then
            
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Routine to Stop Locomotive
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Move throttle to Notch Zero
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
ONEMORETOESTOP:

                Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep.Text = Int(Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) / 8 * 126)

                If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
                End If
    
                If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
                End If
    
                If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) <> 0 Then
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = False
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Wait = False
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Shareable = False
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Open"
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = True
                    Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Play"
                End If
        
                If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                    
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) >= 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 7 Then
                    
                        If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
                        End If
                            
                        If InternetSettings!Winsock.State = sckConnected Then
                            If InternetSettings.optionHost.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        ElseIf InternetSettings!Winsock.State = sckClosed Then
                            If InternetSettings.OptionStandAlone.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        End If
                        
                        If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                            Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                            Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                            Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                            Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                            Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                            Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                            Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                            Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                            Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                            Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                            Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                            Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                            Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
                        End If
                        
                        If InternetSettings!Winsock.State = sckConnected Then
                            If InternetSettings.optionHost.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        ElseIf InternetSettings!Winsock.State = sckClosed Then
                            If InternetSettings.OptionStandAlone.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        End If
                    End If
                
                    '------------------------------------------------------------------------------------------------------
                    ' Turn the Dynamic Brake Sound Off
                
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) = 0 Then
                    
                        If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                            Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                            Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                            Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                            Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                            Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                            Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                            Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                            Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                            Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                            Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                            Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                            Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                            Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
                        End If

                        If InternetSettings!Winsock.State = sckConnected Then
                            If InternetSettings.optionHost.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        ElseIf InternetSettings!Winsock.State = sckClosed Then
                            If InternetSettings.OptionStandAlone.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        End If
                        
                    End If
                    
                    ' -----------------------------------------------------------------------------------------------------
                    ' End of Turning Dynamic Brake Sound Off
                
                    ' -----------------------------------------------------------------------------------------------------
                    ' Turning Dynamic Brake Sound On
                
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
                    
                        If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
                        ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
                        End If

                        If InternetSettings!Winsock.State = sckConnected Then
                            If InternetSettings.optionHost.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        ElseIf InternetSettings!Winsock.State = sckClosed Then
                            If InternetSettings.OptionStandAlone.Value = True Then
                                Call SetSoundDecoderNumber
                                Call SetSoundControlFunction01234
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction5678
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                                Call MainlineOperationGUI.SetSoundDecoderNumber
                                Call SetSoundControlFunction9101112
                                Call MainScreen.SendCommandviaTrackQ
                                DoEvents
                            End If
                        End If
                    End If
                End If
            
                ' ---------------------------------------------------------------------------------------------------------
                ' End of Turning Dynamic Braking Sound On
   
                Let temp = App.Path
                If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
                    Let temp = temp + "\Graphics\DynamicBrake"
                Else
                    Let temp = temp + "\Graphics\Locomotive Diesel1\Throttle"
                End If
                Let temp = temp + MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag
                Let temp = temp + "(d1).bmp"
    
                Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
                
                For t = 1 To 150000
                        DoEvents
                Next t
            
                If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then GoTo ONEMORETOESTOP


' -------------------------------------------------------------------------------------------------------------------------
' Deadmann Switch Setting
'

                Let MainlineOperationGUI!timerdeadmann.Enabled = False
                ' Let MainlineOperationGuiDiesel1Screen.PictureBoxThrottle.Enabled = False
                ' Locks out the engineer if he/she passes into a stop block.


                ' -------------------------------------------------------------------------------------------------------------------------
                ' Routine to Stop Locomotive Ends
            
            Else
                
                If MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching = "Restricted" Then
                
                ' -----------------------------------------------------------------------------------------------------------------------
                ' Restricted Block Routine
                '
    
                    Let MainlineOperationGUI.ConsistControlSpeed.Value = "54"
                            
                    '----------------------------------------------------------------------------------------------------------------------
                    ' Move throttle to Notch 5
        
ONEMORETORESTRICTED:

                    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep.Text = Int(Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) / 8 * 126)

                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
                    End If
    
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
                    End If
    
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) <> 0 Then
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = False
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Wait = False
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Shareable = False
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Open"
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Notify = True
                        Let MainlineOperationGuiDiesel1Screen!SoundThrottle.Command = "Play"
                    End If
        
                    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) >= 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 7 Then
                            
                            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                                Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                                Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                                Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
                            End If
                            
    

                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If

                            
                            If MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "0" Then
                                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "1" Then
                                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "2" Then
                                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "3" Then
                                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "8" Then
                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "9" Then
                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                                Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                                Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                                Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
                            End If
                            
    

                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                        End If
                
                        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) = 0 Then
                        
                            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                                Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                                Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                                Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
                            End If
    
                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                        End If
                    
                        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < -1 Then
                                                   
                            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "8" Then
                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "9" Then
                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "10" Then
                                Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "11" Then
                                Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "12" Then
                                Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
                            
                            End If
                        
                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call SetSoundDecoderNumber
                                    Call SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                            
                        End If
                    End If
                
                    Let temp = App.Path
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
                        Let temp = temp + "\Graphics\DynamicBrake"
                    Else
                        Let temp = temp + "\Graphics\Locomotive Diesel1\Throttle"
                    End If
                    Let temp = temp + MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag
                    Let temp = temp + "(d1).bmp"
    
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
                    
                    For t = 1 To 150000
                            DoEvents
                    Next t
            
                    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 6 Then GoTo ONEMORETORESTRICTED

'
' End of Routine to Restrict Locomotive
' -------------------------------------------------------------------------------------------------------------------------
                End If
            End If
        End If
    
' -------------------------------------------------------------------------------------------------------------------------
' Make the Approaching Signal Current Signal
'

        Let MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text
        
        If MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = "Stop" Then
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
        Else
            If MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = "Restricted" Then
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Else
                If MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = "Reduced" Then
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                Else
                    If MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = "Clear" Then
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                    End If
                End If
            End If
        End If
 
 '
 ' End of Updating Signal Display
 ' ------------------------------------------------------------------------------------------------------------------------
 
 ' ------------------------------------------------------------------------------------------------------------------------
 ' Make a new Signal (random)
 
        Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalDistance.Text = Int((6000 - 3000 + 1) * Rnd(Timer) + 3000)
        
        Let Temporary = Int(100 * Rnd(Timer))
        
        If Temporary <= 16 Then
            Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text = "Stop"
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
        Else
            If Temporary > 16 And Temporary < 41 Then
                Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text = "Restricted"
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Else
                If Temporary >= 41 And Temporary < 56 Then
                    Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text = "Reduced"
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
                    Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                Else
                    If Temporary >= 56 Then
                        Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text = "Clear"
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
                        Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
                    End If
                End If
            End If
        End If
'
' End of Making a new Signal (random)
' -------------------------------------------------------------------------------------------------------------------------
    
    End If
End If

End Sub

Private Sub TimerWaterConsumption_Timer()

Let MainlineOperationGuiDiesel1Screen!TextBoxWaterTank.Text = Val(MainlineOperationGuiDiesel1Screen!TextBoxWaterTank.Text) - 0.01
Let MainlineOperationGuiDiesel1Screen!BarWaterTank.Value = Val(MainlineOperationGuiDiesel1Screen!TextBoxWaterTank.Text)

If Val(MainlineOperationGuiDiesel1Screen!BarWaterTank.Value) > Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusWaterTooLow.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarWaterTank.BorderColor = 49433
Else
    Let MainlineOperationGuiDiesel1Screen!BarWaterTank.BorderColor = 192
End If

End Sub

Private Sub TimerWaterPressure_Timer()

If Val(MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value) < 75 Then
    Let TimerWaterPressure.Interval = Int(((Val(MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value) / 75) ^ 2) * (6555 / 1 + Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag))) * (Val(ScrollTimeAdjustment) / 100 + 1)
    Let MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value) + 1
End If

If Val(MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value) > 80 Or _
   Val(MainlineOperationGuiDiesel1Screen!BarWaterPressure.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusWaterTooLowPressure.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarWaterPressure.BorderColor = 192
Else
    Let MainlineOperationGuiDiesel1Screen!BarWaterPressure.BorderColor = 49433
End If

End Sub

Private Sub TimerWaterTemperature_Timer()


Let TimerWaterTemperature.Interval = Int(((Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) / 180) ^ 2) * (30000 / (1 + Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag)))) * (Val(ScrollTimeAdjustment) / 100 + 1)

If MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag > -1 Then
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) + Int(Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 2 / 2)
Else
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value = Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) + 1
End If

If Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) > 220 Or _
   Val(MainlineOperationGuiDiesel1Screen!BarWaterTemperature.Value) < Val(MainlineOperationGuiDiesel1Screen!TextBoxStatusWaterTooLowTemperature.Text) Then
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.BorderColor = 192
Else
    Let MainlineOperationGuiDiesel1Screen!BarWaterTemperature.BorderColor = 49433
End If

End Sub


Public Sub SetLocomotiveNumber()

    Let MainScreen!onebyted.Text = Int(Val(ComboBoxLocomotiveNumber.Text) / 256)
    Let MainScreen!twoByteD.Text = Val(ComboBoxLocomotiveNumber.Text) - (Val(MainScreen!onebyted.Text) * 256)
    If CheckBoxAddressShortLocomotive.Value = vbUnchecked Then
        ' Add these values to the first byte to make it a long address
        Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64
    Else
        ' Blank out the first text box, so it not sent to the command control unit.
        Let MainScreen!onebyted.Text = ""
    End If
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Locomotive " & MainlineOperationGUI!ComboBoxLocomotiveNumber.Text & " "
    

End Sub

Public Sub SetConsistFunction01234()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128
    If MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 16
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 is Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""

End Sub

Public Sub SetConsistFunction5678()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128 + 32 + 16
    If MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 is Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""

End Sub

Public Sub SetChangeCV()
    
    Dim TemporaryByteOne As Integer
    
    Let TemporaryByteOne = 0
    Let TemporaryByteTwo = Val(ConsistControlCV.Text) - 1
    If TemporaryByteTwo / 512 >= 1 Then
        Let TemporaryByteOne = TemporaryByteOne + 2
        Let TemporaryByteTwo = TemporaryByteTwo - 512
    End If
    If TemporaryByteTwo / 256 >= 1 Then
        Let TemporaryByteOne = TemporaryByteOne + 1
        Let TemporaryByteTwo = TemporaryByteTwo - 256
    End If
    Let TemporaryByteOne = TemporaryByteOne + 128 + 64 + 32
    If MainlineOperationGUI!ConsistControlCVRead = vbChecked Then
            Let TemporaryByteOne = TemporaryByteOne + 4
    Else
        Let TemporaryByteOne = TemporaryByteOne + 8 + 4
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByteOne
    Let MainScreen!FourByteD.Text = TemporaryByteTwo
    Let MainScreen!FiveByteD.Text = Val(ConsistControlCVValue.Text)
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!SixbtyeD.Text = ""
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Change CV" & MainlineOperationGUI!ConsistControlCV.Text & " to " & MainlineOperationGUI!ConsistControlCVValue.Text & vbCrLf

End Sub




Public Sub SetSoundDecoderNumber()

    If CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        Let MainScreen!onebyted.Text = Int(Val(MainlineOperationGUI.ComboBoxSoundDecoderNumber.Text) / 256)
        Let MainScreen!twoByteD.Text = Val(ComboBoxSoundDecoderNumber.Text) - (Val(MainScreen!onebyted.Text) * 256)
        If CheckBoxAddressShortSoundDecoder.Value = vbUnchecked Then
             Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64
        Else
            Let MainScreen!onebyted.Text = ""
        End If
    End If
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Sound Decoder " & MainlineOperationGUI!ComboBoxSoundDecoderNumber.Text & " "

End Sub



Public Sub SetConsistFunction9101112()
    
    Dim TemporaryByte As Integer
    
    Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128 + 32
    If MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 is Off;"
    End If
    If MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 is Off; "
    End If
    If MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 is Off; "
    End If
    
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    
End Sub

Public Sub SetSpeed()

    Dim TemporarySpeedStep As Integer
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up speed in 128 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If ConsistControlSpeed128.Value = vbChecked Then
        Let TemporarySpeedStep = Val(ConsistControlSpeed.Value)
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Speed " & ConsistControlSpeed.Value & " of 128 "
        If ConsistControlDirectionF.Value = vbChecked Then
            TemporarySpeedStep = TemporarySpeedStep + 128 ' add forward direction
            Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "in forward direction."
        Else
            Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "in reverse direction."
        End If
        
        Let MainScreen!ThreeByteD.Text = 63
        Let MainScreen!FourByteD.Text = TemporarySpeedStep
        Let MainScreen!FiveByteD.Text = ""
        Let MainScreen!SixByteD.Text = ""
        Let MainScreen!sevenbyted.Text = ""
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up speed in 28 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
        Let TemporarySpeedStep = 64
        If ConsistControlDirectionF.Value = vbChecked Then
                Let TemporarySpeedStep = TemporarySpeedStep + 32 ' add forward direction
        End If
        
       If ConsistControlSpeed28.Value = vbChecked Then
            Let temp1 = Val(ConsistControlSpeed.Value) ' adds the speed
            Let temp2 = temp1 Mod 2
            Let newspeedvalue = Int(temp1 / 2)
            Let TemporarySpeedStep = TemporarySpeedStep + newspeedvalue
            If temp2 = 1 Then Let TemporarySpeedStep = TemporarySpeedStep + 16
            Let MainScreen!ThreeByteD.Text = TemporarySpeedStep
            Let MainScreen!FourByteD.Text = ""
            Let MainScreen!FiveByteD.Text = ""
            Let MainScreen!SixByteD.Text = ""
            Let MainScreen!sevenbyted.Text = ""
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up spped for 14 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Else
            ' This routing assembles the byte for speed step mode 14
            
            Let TemporarySpeedStep = TemporarySpeedStep + Val(ConsistControlSpeed.Value) ' add the speed
            Let MainScreen!ThreeByteD.Text = TemporarySpeedStep
            Let MainScreen!FourByteD.Text = ""
            Let MainScreen!FiveByteD.Text = ""
            Let MainScreen!SixByteD.Text = ""
            Let MainScreen!sevenbyted.Text = ""
        End If
    End If

End Sub

Public Sub SetSoundControlFunction01234()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128
    If MainlineOperationGUI!SoundControlFunction0.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 16
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction1.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction2.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction3.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction4.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 is Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""

End Sub

Public Sub SetSoundControlFunction5678()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128 + 32 + 16
    If MainlineOperationGUI!SoundControlFunction5.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 of Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction6.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction7.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction8.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 is On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 is Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    
End Sub

Public Sub SetSoundControlFunction9101112()

    Dim TemporaryByte As Integer
    
    Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text & "Function "
    Let TemporaryByte = 128 + 32
    If MainlineOperationGUI!SoundControlFunction9.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction10.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction11.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 is Off; "
    End If
    If MainlineOperationGUI!SoundControlFunction12.Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 is On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 is Off; "
    End If
    
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    
End Sub

Private Sub TimerWindowState_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI, Timer Window State, Timer" & vbCrLf
    End If ' Debug Tag

    Let TimerWindowState.Enabled = False
    Let TimerWindowState.Interval = 0

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Screen"
    Let Ini.Parameter = "WindowStateHostMode"
    If Ini.Value = "Minimized" Then
        Let Screen.ActiveForm.WindowState = vbMinimized
    ElseIf Ini.Value = "Normal" Then
        Let Screen.ActiveForm.WindowState = vbNormal
    ElseIf Ini.Value <> "Maximized" Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Mainline Operation GUI Screen, Timer Window State, wrong variable in ATC.INI file for 'WindowStateHostMode' setting."
        End If
    End If
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI, Timer Window State, Timer" & vbCrLf
    End If ' Debug Tag

    
End Sub


