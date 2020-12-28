VERSION 4.00
Begin VB.Form MainlineOperationGUIScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Height          =   11670
   Left            =   0
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleMode       =   0  'User
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Top             =   0
   Width           =   15450
   Begin VB.PictureBox PictureBoxLocomotiveCab 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11265
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   529.145
      ScaleMode       =   0  'User
      ScaleWidth      =   1022
      TabIndex        =   0
      Top             =   0
      Width           =   15330
      Begin VB.TextBox TextBoxStatusSpeedStepNow 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Text            =   "0"
         Top             =   840
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxStatusSpeedStepModified 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Text            =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxStatusAutomaticBrakePressure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxStatusIndependentBrakePressure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   12
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton ButtonVideoSettings 
         Caption         =   "Video Settings"
         Enabled         =   0   'False
         Height          =   255
         Left            =   13920
         TabIndex        =   189
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton ButtonVideo 
         Caption         =   "Video is Off"
         Enabled         =   0   'False
         Height          =   255
         Left            =   13920
         TabIndex        =   186
         Top             =   7800
         Width           =   1335
      End
      Begin VB.TextBox LabelCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox TextBoxStatusTractionEffortTooHigh 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   149
         Text            =   "98000"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TextBoxStatusAmpereTooHigh 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   148
         Text            =   "1200"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton ButtonHelp 
         Caption         =   "Help is Off"
         Height          =   255
         Left            =   13920
         TabIndex        =   147
         Top             =   8160
         Width           =   1335
      End
      Begin VB.TextBox TextBoxStatusFuelTooLow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   145
         Text            =   "1000"
         Top             =   3840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusOilTooLow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   144
         Text            =   "60"
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusWaterTooLow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   143
         Text            =   "600"
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusSandTooLow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   142
         Text            =   "70"
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusWaterTooLowPressure 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   141
         Text            =   "60"
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusWaterTooLowTemperature 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   140
         Text            =   "60"
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusOilTooLowPressure 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   139
         Text            =   "60"
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusOilTooLowTemperature 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   138
         Text            =   "60"
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   8
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   118
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label VideoCaptureTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Video Capture Notes"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1245
            TabIndex        =   187
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label VideoCaptureNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "No information available"
            ForeColor       =   &H0000C000&
            Height          =   2415
            Left            =   120
            TabIndex        =   188
            Top             =   240
            Width           =   3765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 8"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   16
            Left            =   3240
            TabIndex        =   119
            Top             =   2670
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   7
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   116
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 7"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   15
            Left            =   3240
            TabIndex        =   117
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   6
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   114
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 6"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   14
            Left            =   3240
            TabIndex        =   115
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   5
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   112
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 5"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   13
            Left            =   3240
            TabIndex        =   113
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   4
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   110
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 4"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   12
            Left            =   3240
            TabIndex        =   111
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   3
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   108
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 3"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   11
            Left            =   3240
            TabIndex        =   109
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   2
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   106
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 2"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   10
            Left            =   3240
            TabIndex        =   107
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   1
         Left            =   3195
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   104
         Top             =   5715
         Width           =   3990
         Begin VB.PictureBox PictureBoxSpeedometer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2460
            Left            =   -15
            Picture         =   "MainlineOperationGUIScreenDiesel.frx":0000
            ScaleHeight     =   2460
            ScaleWidth      =   2220
            TabIndex        =   120
            Top             =   15
            Width           =   2220
            Begin VB.TextBox TextBoxDigitalSpeed 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H0000C000&
               Height          =   285
               Left            =   780
               TabIndex        =   121
               Text            =   "0.0"
               Top             =   1920
               Width           =   705
            End
            Begin VB.Label LabelSpeed10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "10.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   131
               Top             =   1200
               Width           =   345
            End
            Begin VB.Label LabelSpeed9 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "9.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1680
               TabIndex        =   130
               Top             =   840
               Width           =   255
            End
            Begin VB.Label LabelSpeed8 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "8.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1560
               TabIndex        =   129
               Top             =   585
               Width           =   255
            End
            Begin VB.Label LabelSpeed7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "7.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   1320
               TabIndex        =   128
               Top             =   360
               Width           =   255
            End
            Begin VB.Label LabelSpeed6 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "6.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   990
               TabIndex        =   127
               Top             =   240
               Width           =   255
            End
            Begin VB.Label LabelSpeed5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "5.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   645
               TabIndex        =   126
               Top             =   390
               Width           =   255
            End
            Begin VB.Label LabelSpeed4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "4.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   390
               TabIndex        =   125
               Top             =   645
               Width           =   255
            End
            Begin VB.Label LabelSpeed3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "3.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   255
               TabIndex        =   124
               Top             =   990
               Width           =   255
            End
            Begin VB.Label LabelSpeed2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "2.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   360
               TabIndex        =   123
               Top             =   1320
               Width           =   255
            End
            Begin VB.Label LabelSpeed1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "1.0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   510
               TabIndex        =   122
               Top             =   1575
               Width           =   255
            End
         End
         Begin VB.Label LabelDeadmann 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Deadmann"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   133
            Top             =   2520
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label LabelCombinedPower 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Combined Power  Neutral - Notch Zero"
            ForeColor       =   &H0000C000&
            Height          =   690
            Left            =   2205
            TabIndex        =   132
            Top             =   15
            Width           =   1785
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 1"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   9
            Left            =   3240
            TabIndex        =   105
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   6600
         TabIndex        =   103
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   102
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5640
         TabIndex        =   101
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   100
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   99
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   98
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   97
         Top             =   8760
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   8
         Left            =   9090
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   88
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 8"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   89
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   7
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   87
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 7"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   90
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   6
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   86
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 6"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   91
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   5
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   85
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 5"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   92
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   4
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   84
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 4"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   4
            Left            =   3240
            TabIndex        =   93
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   3
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   68
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 3"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   5
            Left            =   3240
            TabIndex        =   94
            Top             =   2660
            Width           =   645
         End
         Begin GBAR1.GBar BarTractionEffort 
            Height          =   165
            Left            =   1080
            TabIndex        =   75
            Top             =   960
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Max             =   130000
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 113,000 lbs"
         End
         Begin VB.Label LabelTractionEffort 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Traction Effort"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   0
            TabIndex        =   74
            Top             =   960
            Width           =   1005
         End
         Begin GBAR1.GBar BarAmpere 
            Height          =   165
            Left            =   1080
            TabIndex        =   73
            Top             =   720
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            ForeColor       =   192
            BackColor       =   0
            Max             =   1300
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 1200 amps"
         End
         Begin VB.Label LabelAmpere 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Ampere"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   480
            TabIndex        =   72
            Top             =   720
            Width           =   540
         End
         Begin GBAR1.GBar BarRPM 
            Height          =   165
            Left            =   1080
            TabIndex        =   71
            Top             =   480
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            ForeColor       =   192
            BackColor       =   0
            Max             =   2000
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 2000 rpm"
         End
         Begin VB.Label LabelRPM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Engine RPM"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   900
         End
         Begin VB.Label LabelPowerDistribution 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Power Distribution"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1080
            TabIndex        =   69
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   2
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   66
         Top             =   5715
         Visible         =   0   'False
         Width           =   3990
         Begin VB.CommandButton ButtonFillOil 
            Height          =   195
            Left            =   240
            TabIndex        =   137
            Top             =   1200
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillWater 
            Height          =   195
            Left            =   240
            TabIndex        =   136
            Top             =   960
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillSand 
            Height          =   195
            Left            =   240
            TabIndex        =   135
            Top             =   720
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillFuel 
            Caption         =   "Command1"
            Height          =   195
            Left            =   240
            TabIndex        =   134
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 2"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   6
            Left            =   3240
            TabIndex        =   95
            Top             =   2660
            Width           =   645
         End
         Begin GBAR1.GBar BarOilTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   83
            Top             =   1200
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Max             =   80
            Value           =   63
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 80 gal"
         End
         Begin VB.Label LabelOilTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Oil"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   600
            TabIndex        =   82
            Top             =   1200
            Width           =   420
         End
         Begin GBAR1.GBar BarWaterTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   81
            Top             =   960
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Max             =   750
            Value           =   660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   "of 750 gal"
         End
         Begin VB.Label LabelWaterTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Water"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   480
            TabIndex        =   80
            Top             =   960
            Width           =   555
         End
         Begin GBAR1.GBar BarSandTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   79
            Top             =   720
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Max             =   200
            Value           =   150
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 200 lbs"
         End
         Begin VB.Label LabelSandTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Sand"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   480
            TabIndex        =   78
            Top             =   720
            Width           =   495
         End
         Begin GBAR1.GBar BarFuelTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   77
            Top             =   480
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Max             =   4500
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   192
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 4500 gal"
         End
         Begin VB.Label LabelFuelTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Fuel"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   600
            TabIndex        =   76
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Resources"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1680
            TabIndex        =   67
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9120
         TabIndex        =   65
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   64
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10080
         TabIndex        =   63
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   10560
         TabIndex        =   62
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   11040
         TabIndex        =   61
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   11520
         TabIndex        =   60
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   12000
         TabIndex        =   59
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton ButtonScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   12480
         TabIndex        =   58
         Top             =   8760
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   1
         Left            =   9075
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   40
         Top             =   5715
         Width           =   3990
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 1"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   7
            Left            =   3240
            TabIndex        =   96
            Top             =   2660
            Width           =   645
         End
         Begin VB.Label LabelWaterTemperature 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Temperature"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   2355
            Width           =   900
         End
         Begin GBAR1.GBar BarWaterTemperature 
            Height          =   165
            Left            =   1080
            TabIndex        =   56
            Top             =   2370
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            ForeColor       =   192
            BackColor       =   0
            Min             =   40
            Max             =   230
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 230 degrees"
         End
         Begin VB.Label LabelWaterPressure 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Pressure"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   360
            TabIndex        =   55
            Top             =   2130
            Width           =   615
         End
         Begin GBAR1.GBar BarWaterPressure 
            Height          =   165
            Left            =   1080
            TabIndex        =   54
            Top             =   2130
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            ForeColor       =   192
            BackColor       =   0
            Min             =   40
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 100 lbs"
         End
         Begin VB.Label LabelWater 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Water"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1800
            TabIndex        =   53
            Top             =   1875
            Width           =   450
         End
         Begin GBAR1.GBar BarOilTemperature 
            Height          =   165
            Left            =   1080
            TabIndex        =   52
            Top             =   1620
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            ForeColor       =   192
            BackColor       =   0
            Min             =   40
            Max             =   230
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 230 degrees"
         End
         Begin GBAR1.GBar BarOilPressure 
            Height          =   165
            Left            =   1080
            TabIndex        =   51
            Top             =   1395
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Min             =   40
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   " of 100 lbs"
         End
         Begin VB.Label LabelOilTemperature 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Temperature"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   0
            TabIndex        =   50
            Top             =   1635
            Width           =   1020
         End
         Begin VB.Label LabelOilPressure 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Pressure"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   1380
            Width           =   825
         End
         Begin VB.Label LabelOil 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Oil"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1845
            TabIndex        =   48
            Top             =   1050
            Width           =   360
         End
         Begin GBAR1.GBar BarAirIndependentBrake 
            Height          =   165
            Left            =   1080
            TabIndex        =   47
            Top             =   765
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Min             =   40
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   "of 100 lbs"
         End
         Begin VB.Label LabelIndependent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Independent"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   765
            Width           =   1020
         End
         Begin GBAR1.GBar BarAirAutomaticBrake 
            Height          =   165
            Left            =   1080
            TabIndex        =   45
            Top             =   540
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Min             =   40
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   "of 100 lbs"
         End
         Begin VB.Label LabelAutomatic 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Automatic"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   525
            Width           =   825
         End
         Begin VB.Label LabelPressure 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Air Pressure"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1590
            TabIndex        =   43
            Top             =   15
            Width           =   840
         End
         Begin GBAR1.GBar BarAirMainReservoir 
            Height          =   165
            Left            =   1080
            TabIndex        =   42
            Top             =   300
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   291
            BackColor       =   0
            Min             =   40
            Value           =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
            FillStyle       =   0
            FontColor       =   16777215
            Units           =   "of 100 lbs"
         End
         Begin VB.Label LabelMainReservoir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Main Reservoir"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   0
            TabIndex        =   41
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.PictureBox PictureBoxResetLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   1425
         ScaleHeight     =   705
         ScaleWidth      =   870
         TabIndex        =   39
         Top             =   10170
         Width           =   870
      End
      Begin VB.PictureBox PictureBoxBell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2520
         ScaleHeight     =   540
         ScaleWidth      =   765
         TabIndex        =   38
         Top             =   10191
         Width           =   765
      End
      Begin VB.TextBox TextBoxStatusPhone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Text            =   "1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6270
         Left            =   0
         ScaleHeight     =   6270
         ScaleWidth      =   1950
         TabIndex        =   36
         Top             =   5250
         Width           =   1950
      End
      Begin VB.CommandButton ButtonDetail 
         Caption         =   "& Data is Off"
         Height          =   285
         Left            =   13920
         TabIndex        =   35
         Top             =   8880
         Width           =   1305
      End
      Begin VB.CommandButton ButtonCaption 
         Caption         =   "&Caption is Off"
         Height          =   255
         Left            =   13920
         TabIndex        =   34
         Top             =   8520
         Width           =   1335
      End
      Begin VB.PictureBox PictureBoxThrottle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2038
         Left            =   5493
         ScaleHeight     =   2040
         ScaleWidth      =   660
         TabIndex        =   33
         Top             =   9034
         Width           =   664
      End
      Begin VB.TextBox TextBoxStatusReverser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   32
         Text            =   "1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox PictureBoxReverser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1830
         Left            =   4185
         ScaleHeight     =   1830
         ScaleWidth      =   765
         TabIndex        =   31
         Top             =   8985
         Width           =   765
      End
      Begin VB.CommandButton ButtonCloseGUI 
         Caption         =   "&Close"
         Height          =   255
         Left            =   13920
         TabIndex        =   30
         Top             =   9240
         Width           =   1290
      End
      Begin VB.TextBox TextBoxStatusResetRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   29
         Text            =   "Off"
         Top             =   3720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxStatusLight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   28
         Text            =   "0"
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxStatusThrottle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   27
         Text            =   "0"
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusSand 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   26
         Text            =   "Off"
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TextBoxStatusHorn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   25
         Text            =   "Off"
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox PictureBoxLight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   13890
         ScaleHeight     =   1485
         ScaleWidth      =   1440
         TabIndex        =   24
         Top             =   5715
         Width           =   1440
      End
      Begin VB.PictureBox PictureBoxResetRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   14130
         ScaleHeight     =   720
         ScaleWidth      =   915
         TabIndex        =   23
         Top             =   10230
         Width           =   915
      End
      Begin VB.PictureBox PictureBoxHorn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2085
         ScaleHeight     =   480
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   10725
         Width           =   795
      End
      Begin VB.PictureBox PictureBoxIndependentBrake 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   12615
         ScaleHeight     =   2115
         ScaleWidth      =   1395
         TabIndex        =   20
         Top             =   8940
         Width           =   1395
      End
      Begin VB.TextBox TextBoxStatusIndependentBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   19
         Text            =   "9"
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxAutomaticBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   7710
         ScaleHeight     =   1860
         ScaleWidth      =   1905
         TabIndex        =   18
         Top             =   9090
         Width           =   1905
      End
      Begin VB.TextBox TextBoxStatusBell 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Text            =   "Off"
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TextBoxStatusSpeedStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Text            =   "0"
         Top             =   1080
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxStatusAutomaticBrake 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TextBoxRadiatorFans1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   10
         Text            =   "Off"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxRadiatorFans2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   9
         Text            =   "Off"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxRadiatorFans3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   8
         Text            =   "Off"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxDynamicBrakeFan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   7
         Text            =   "Off"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxStatusThrottleDelay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   3360
         TabIndex        =   6
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox TextBoxFuelTank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   5
         Text            =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxOilTank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   4
         Text            =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxWaterTank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   3
         Text            =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextBoxSandTank 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   2
         Text            =   "0"
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton ButtonScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   1
         Top             =   8760
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxSand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   440
         Left            =   2865
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   22
         Top             =   9720
         Width           =   675
      End
      Begin VB.TextBox TextBoxStatusResetLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3360
         TabIndex        =   146
         Text            =   "Off"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LabelStatusResetLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TextBoxStatusResetLeft"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   190
         Top             =   3960
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSpeedStepNow"
         Height          =   180
         Left            =   360
         TabIndex        =   164
         Top             =   840
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusIndependentBrakePressure"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   167
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSpeedStepModified"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   165
         Top             =   600
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusAutomaticBrakePressure"
         Height          =   180
         Left            =   240
         TabIndex        =   166
         Top             =   360
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Label LabelRadiatorFans1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxRadiator Fans 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   185
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelRadiatorFans2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxRadiator Fans 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   184
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelRadiatorFans3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxRadiator Fans 3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   183
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelDynamicBrakeFan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxDynamic Brake Fan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   182
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxFuelTank"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   181
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxOilTank"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   180
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxWaterTank"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   179
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label48 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxSandTank"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   178
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelOilTooLowTemperature 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusOilTooLowTemerature"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   177
         Top             =   2160
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label LabelOilTooLowPressure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusOilTooLowPressure"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   176
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label LabelWaterTooLowTemperature 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusWaterTooLowTemperature"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   175
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LabelWaterTooLowPressure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusWaterTooLowPressure"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   174
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label LabelSandTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusSandTooLow"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   173
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelWaterTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusWaterTooLow"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   172
         Top             =   3360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelOilTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusOilTooLow"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   171
         Top             =   3600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelFuelTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusFuelTooLow"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   170
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelStatusAmpereTooHigh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusAmpereTooHigh"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   169
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LabelStatusTractionEffortTooHigh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TextBoxStatusTractionEffortTooHigh"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   168
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSpeedStep"
         Height          =   180
         Left            =   1320
         TabIndex        =   163
         Top             =   1080
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusAutomaticBrake"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   960
         TabIndex        =   162
         Top             =   1320
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusIndependentBrake"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   161
         Top             =   1560
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusReverser"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   160
         Top             =   1800
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusBell"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   159
         Top             =   2040
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label LabelStatusHorn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusHorn"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   158
         Top             =   2280
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label LabelStatusSand 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSand"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   157
         Top             =   2520
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label LabelStatusLight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusLight"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   156
         Top             =   2760
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusThrottle"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1680
         TabIndex        =   155
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusThrottleDelay"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   154
         Top             =   3240
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label LabelStatusPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusPhone"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   153
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelStatusResetRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "TextBoxStatusResetRight"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   152
         Top             =   3720
         Visible         =   0   'False
         Width           =   1830
      End
      Begin vbVidCap.ezVidCap VideoCapture 
         Height          =   5175
         Left            =   0
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   9128
         AutoSize        =   0   'False
         BorderStyle     =   0
         VideoBorder     =   0
         StretchPreview  =   -1  'True
         MakeUserConfirmCapture=   0
         AbortLeftMouse  =   0
         AbortRightMouse =   0
         StreamMaster    =   1
      End
   End
   Begin VB.Menu menuCaptureDevice 
      Caption         =   "Capture Device"
      Visible         =   0   'False
      Begin VB.Menu menuCaptureDeviceVideoSource 
         Caption         =   "Video Source"
      End
      Begin VB.Menu menuCaptureDeviceAudioSetting 
         Caption         =   "Audio Settings"
      End
      Begin VB.Menu menuCaptureDeviceVideoFormat 
         Caption         =   "Video Format"
      End
      Begin VB.Menu menuCaptureDeviceVideoCompression 
         Caption         =   "Video Compression"
      End
      Begin VB.Menu menuCaptureDeviceVideoDisplay 
         Caption         =   "Video Display"
      End
   End
End
Attribute VB_Name = "MainlineOperationGUIScreen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False







Private Sub ButtonCaption_Click()

If ButtonCaption.Caption = "&Caption is On" Then
    Let ButtonCaption.Caption = "&Caption is Off"
Else
    Let ButtonCaption.Caption = "&Caption is On"
End If

End Sub

Private Sub ButtonCloseGUI_Click()

Let MainlineOperationGUI.TimerSendCommand.Interval = 6000
Let MainlineOperationGUI.TimerSendCommand.Enabled = False
Let MainlineOperationGUI.TimerSpeedChange.Interval = 1000
Let MainlineOperationGUI.TimerSpeedChange.Enabled = False
Let MainlineOperationGUI.timerairmainReservoir.Interval = 3000
Let MainlineOperationGUI.timerairmainReservoir.Enabled = False
Let MainlineOperationGUI.TimerAirAutomaticBrake.Interval = 1250
Let MainlineOperationGUI.TimerAirAutomaticBrake.Enabled = False
Let MainlineOperationGUI.TimerAirIndependentBrake.Interval = 1000
Let MainlineOperationGUI.TimerAirIndependentBrake.Enabled = False
Let MainlineOperationGUI.timerthrottledelay.Interval = 1000
Let MainlineOperationGUI.timerthrottledelay.Enabled = False
Let MainlineOperationGUI.TimerOilPressure.Interval = 250
Let MainlineOperationGUI.TimerOilPressure.Enabled = False
Let MainlineOperationGUI.TimerOilTemperature.Interval = 250
Let MainlineOperationGUI.TimerOilTemperature.Enabled = False
Let MainlineOperationGUI.TimerWaterPressure.Interval = 250
Let MainlineOperationGUI.TimerWaterPressure.Enabled = False
Let MainlineOperationGUI.TimerWaterTemperature.Interval = 250
Let MainlineOperationGUI.TimerWaterTemperature.Enabled = False
Let MainlineOperationGUI.timerfuelconsumption.Interval = 1000
Let MainlineOperationGUI.timerfuelconsumption.Enabled = False
Let MainlineOperationGUI.TimerOilConsumption.Interval = 65535
Let MainlineOperationGUI.TimerOilConsumption.Enabled = False
Let MainlineOperationGUI.TimerWaterConsumption.Interval = 65535
Let MainlineOperationGUI.TimerWaterConsumption.Enabled = False
Let MainlineOperationGUI.TimerRadiatorFans.Interval = 2000
Let MainlineOperationGUI.TimerRadiatorFans.Enabled = False
Let MainlineOperationGUI.TimerRPM.Interval = 125
Let MainlineOperationGUI.TimerRPM.Enabled = False


Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
Call SetLocomotiveNumber
Call SetSpeed
Call SendCommand
            
If MainlineOperationGUI.CheckBoxSoundDecoderEquipped.Value = vbChecked Then
    Call SetSoundDecoderNumber
    Let MainlineOperationGUI.ConsistControlFunction6.Value = vbChecked
    Call SetFunction5678
    Call SendCommand
    DoEvents
    Call SetSoundDecoderNumber
    Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
    Call SetFunction5678
    Call SendCommand
End If
 
MainlineOperationGUIScreen.Hide
Unload MainlineOperationGUIScreen
MainlineOperationGUI.Show vbmodelless

End Sub

Private Sub ButtonDetail_Click()

If ButtonDetail.Caption = "&Data is On" Then

    Let ButtonDetail.Caption = "&Data on Off"
    
    Let Label36.Visible = False
    Let Label38.Visible = False
    Let Label30.Visible = False
    Let Label37.Visible = False
    Let Label29.Visible = False
    Let Label27.Visible = False
    Let Label25.Visible = False
    Let Label21.Visible = False
    Let Label20.Visible = False
    Let Label26.Visible = False
    Let LabelStatusHorn.Visible = False
    Let LabelStatusSand.Visible = False
    Let LabelStatusLight.Visible = False
    Let LabelStatusResetRight.Visible = False
    Let LabelStatusResetLeft.Visible = False
    Let Label28.Visible = False
    Let Label38.Visible = False
    Let LabelRadiatorFans1.Visible = False
    Let LabelRadiatorFans2.Visible = False
    Let LabelRadiatorFans3.Visible = False
    Let LabelDynamicBrakeFan.Visible = False
    Let Label39.Visible = False
    Let Label46.Visible = False
    Let Label47.Visible = False
    Let Label48.Visible = False
    Let LabelStatusPhone.Visible = False
    Let LabelOilTooLowPressure.Visible = False
    Let LabelOilTooLowTemperature.Visible = False
    Let LabelWaterTooLowPressure.Visible = False
    Let LabelWaterTooLowTemperature.Visible = False
    Let LabelSandTooLow.Visible = False
    Let LabelWaterTooLow.Visible = False
    Let LabelOilTooLow.Visible = False
    Let LabelFuelTooLow.Visible = False
    Let LabelStatusAmpereTooHigh.Visible = False
    Let LabelStatusTractionEffortTooHigh.Visible = False
    
    Let TextBoxStatusIndependentBrakePressure.Visible = False
    Let TextBoxStatusAutomaticBrakePressure.Visible = False
    Let TextBoxStatusSpeedStepModified.Visible = False
    Let TextBoxStatusSpeedStepNow.Visible = False
    Let TextBoxStatusSpeedStep.Visible = False
    Let TextBoxStatusAutomaticBrake.Visible = False
    Let TextBoxStatusIndependentBrake.Visible = False
    Let TextBoxStatusReverser.Visible = False
    Let TextBoxStatusBell.Visible = False
    Let TextBoxStatusHorn.Visible = False
    Let TextBoxStatusSand.Visible = False
    Let TextBoxStatusLight.Visible = False
    Let TextBoxStatusResetRight.Visible = False
    Let TextBoxStatusResetLeft.Visible = False
    Let TextBoxStatusThrottle.Visible = False
    Let TextBoxStatusThrottleDelay.Visible = False
    Let TextBoxRadiatorFans1.Visible = False
    Let TextBoxRadiatorFans2.Visible = False
    Let TextBoxRadiatorFans3.Visible = False
    Let TextBoxDynamicBrakeFan.Visible = False
    Let TextBoxFuelTank.Visible = False
    Let TextBoxOilTank.Visible = False
    Let TextBoxWaterTank.Visible = False
    Let TextBoxSandTank.Visible = False
    Let TextBoxStatusPhone.Visible = False
    Let TextBoxStatusOilTooLowTemperature.Visible = False
    Let TextBoxStatusOilTooLowTemperature.Visible = False
    Let TextBoxStatusOilTooLowPressure.Visible = False
    Let TextBoxStatusWaterTooLowTemperature.Visible = False
    Let TextBoxStatusWaterTooLowPressure.Visible = False
    Let textboxstatussandtoolow.Visible = False
    Let TextBoxStatusWaterTooLow.Visible = False
    Let TextBoxStatusOilTooLow.Visible = False
    Let TextBoxStatusFuelTooLow.Visible = False
    Let TextBoxStatusAmpereTooHigh.Visible = False
    Let TextBoxStatusTractionEffortTooHigh.Visible = False
        
Else
    
    Let ButtonDetail.Caption = "&Data is On"
    
    Let Label36.Visible = True
    Let Label38.Visible = True
    Let Label30.Visible = True
    Let Label37.Visible = True
    Let Label29.Visible = True
    Let Label27.Visible = True
    Let Label25.Visible = True
    Let Label21.Visible = True
    Let Label20.Visible = True
    Let Label26.Visible = True
    Let LabelStatusHorn.Visible = True
    Let LabelStatusSand.Visible = True
    Let LabelStatusLight.Visible = True
    Let LabelStatusResetRight.Visible = True
    Let LabelStatusResetLeft.Visible = True
    Let Label28.Visible = True
    Let Label38.Visible = True
    Let LabelRadiatorFans1.Visible = True
    Let LabelRadiatorFans2.Visible = True
    Let LabelRadiatorFans3.Visible = True
    Let LabelDynamicBrakeFan.Visible = True
    Let Label39.Visible = True
    Let Label46.Visible = True
    Let Label47.Visible = True
    Let Label48.Visible = True
    Let LabelStatusPhone.Visible = True
    Let LabelOilTooLowPressure.Visible = True
    Let LabelOilTooLowTemperature.Visible = True
    Let LabelWaterTooLowPressure.Visible = True
    Let LabelWaterTooLowTemperature.Visible = True
    Let LabelSandTooLow.Visible = True
    Let LabelWaterTooLow.Visible = True
    Let LabelOilTooLow.Visible = True
    Let LabelFuelTooLow.Visible = True
    Let LabelStatusAmpereTooHigh.Visible = True
    Let LabelStatusTractionEffortTooHigh.Visible = True
    
    
    Let TextBoxStatusIndependentBrakePressure.Visible = True
    Let TextBoxStatusAutomaticBrakePressure.Visible = True
    Let TextBoxStatusSpeedStepModified.Visible = True
    Let TextBoxStatusSpeedStepNow.Visible = True
    Let TextBoxStatusSpeedStep.Visible = True
    Let TextBoxStatusAutomaticBrake.Visible = True
    Let TextBoxStatusIndependentBrake.Visible = True
    Let TextBoxStatusReverser.Visible = True
    Let TextBoxStatusBell.Visible = True
    Let TextBoxStatusHorn.Visible = True
    Let TextBoxStatusSand.Visible = True
    Let TextBoxStatusLight.Visible = True
    Let TextBoxStatusResetRight.Visible = True
    Let TextBoxStatusResetLeft.Visible = True
    Let TextBoxStatusThrottle.Visible = True
    Let TextBoxStatusThrottleDelay.Visible = True
    Let TextBoxStatusThrottleDelay.Visible = True
    Let TextBoxRadiatorFans1.Visible = True
    Let TextBoxRadiatorFans2.Visible = True
    Let TextBoxRadiatorFans3.Visible = True
    Let TextBoxDynamicBrakeFan.Visible = True
    Let TextBoxFuelTank.Visible = True
    Let TextBoxOilTank.Visible = True
    Let TextBoxWaterTank.Visible = True
    Let TextBoxSandTank.Visible = True
    Let TextBoxStatusPhone.Visible = True
    Let TextBoxStatusOilTooLowTemperature.Visible = True
    Let TextBoxStatusOilTooLowPressure.Visible = True
    Let TextBoxStatusWaterTooLowTemperature.Visible = True
    Let TextBoxStatusWaterTooLowPressure.Visible = True
    Let textboxstatussandtoolow.Visible = True
    Let TextBoxStatusWaterTooLow.Visible = True
    Let TextBoxStatusOilTooLow.Visible = True
    Let TextBoxStatusFuelTooLow.Visible = True
    Let TextBoxStatusAmpereTooHigh.Visible = True
    Let TextBoxStatusTractionEffortTooHigh.Visible = True
    
    
End If

End Sub

Private Sub ButtonFillFuel_Click()
   
    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
 
Let TextBoxFuelTank.Text = Val(BarFuelTank.Max) - 100

End Sub

Private Sub ButtonFillFuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Fuel Refilling Button"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenRight(2).Top) + (Val(ButtonFillFuel.Top) / 15) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenRight(2).Left) + (Val(ButtonFillFuel.Left) / 15) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub ButtonFillOil_Click()
   
    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
 
Let TextBoxOilTank.Text = Val(BarOilTank.Max) - 10

End Sub

Private Sub ButtonFillOil_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Oil Refilling Button"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenRight(2).Top) + (Val(ButtonFillOil.Top) / 15) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenRight(2).Left) + (Val(ButtonFillOil.Left) / 15) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub ButtonFillSand_Click()
   
    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
 
Let TextBoxSandTank.Text = Val(BarSandTank.Max) - 10

End Sub

Private Sub ButtonFillSand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Sand Refilling Button"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenRight(2).Top) + (Val(ButtonFillSand.Top) / 15) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenRight(2).Left) + (Val(ButtonFillSand.Left) / 15) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub ButtonFillWater_Click()
   
    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
 
Let TextBoxWaterTank.Text = Val(BarWaterTank.Max) - 43

End Sub

Private Sub ButtonFillWater_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Water Refilling Button"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenRight(2).Top) + (Val(ButtonFillWater.Top) / 15) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenRight(2).Left) + (Val(ButtonFillWater.Left) / 15) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub ButtonHelp_Click()

If ButtonHelp.Caption = "Help is Off" Then
    Let ButtonHelp.Caption = "Help is On"
Else
    Let ButtonHelp.Caption = "Help is Off"
End If

End Sub

Private Sub ButtonScreenLeft_Click(Index As Integer)

    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1

For X = 1 To 8
    Let PictureBoxComputerScreenLeft(X).Visible = False
Next X

    Let PictureBoxComputerScreenLeft(Index).Visible = True

End Sub

Private Sub ButtonScreenLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Function" + Str$(Index) + " Computer Screen"
    Let LabelCaption.Top = Val(ButtonScreenLeft(Index).Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(ButtonScreenLeft(Index).Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub ButtonScreenRight_Click(Index As Integer)
    
    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1

For X = 1 To 8
    Let PictureBoxComputerScreenRight(X).Visible = False
Next X

    Let PictureBoxComputerScreenRight(Index).Visible = True

End Sub

Private Sub ButtonScreenRight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Function" + Str$(Index) + " Computer Screen"
    Let LabelCaption.Top = Val(ButtonScreenRight(Index).Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(ButtonScreenRight(Index).Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub



Private Sub ButtonVideo_Click()

If ButtonVideo.Caption = "Video is Off" Then
    Let ButtonVideo.Caption = "Video is On"
    Let VideoCapture.Visible = True
Else
    Let ButtonVideo.Caption = "Video is Off"
    Let VideoCapture.Visible = False
End If

End Sub


Private Sub ButtonVideoSettings_Click()

With VideoCapture
    If .HasAudio Then
        Let menuCaptureDeviceAudioSetting.Enabled = False
    Else
        Let menuCaptureDeviceAudioSetting.Enabled = True
    End If
    If .HasDlgFormat Then
        Let menuCaptureDeviceVideoFormat.Enabled = True
    Else
        Let menuCaptureDeviceVideoFormat.Enabled = False
    End If
    If .HasDlgDisplay Then
        Let menuCaptureDeviceVideoDisplay.Enabled = True
    Else
        Let menuCaptureDeviceVideoDisplay.Enabled = False
    End If
    If .HasDlgSource Then
        Let menuCaptureDeviceVideoSource.Enabled = True
    Else
        Let menuCaptureDeviceVideoSource.Enabled = False
    End If
End With

MainlineOperationGUIScreen.PopupMenu menuCaptureDevice

End Sub

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.

Let PictureBoxLocomotiveCab.Picture = LoadPicture(App.Path + "\Gui\CabScreen.bmp")
Let PictureBoxResetLeft.Picture = LoadPicture(App.Path + "\Gui\ResetLeftOff.bmp")
Let PictureBoxSand.Picture = LoadPicture(App.Path + "\Gui\SandOff.bmp")
Let PictureBoxBell.Picture = LoadPicture(App.Path + "\Gui\BellOff.bmp")
Let PictureBoxHorn.Picture = LoadPicture(App.Path + "\Gui\HornOff.bmp")
Let PictureBoxReverser.Picture = LoadPicture(App.Path + "\Gui\reverser1.bmp")
Let PictureBoxThrottle.Picture = LoadPicture(App.Path + "\Gui\Throttle0.bmp")
Let PictureBoxAutomaticBrake.Picture = LoadPicture(App.Path + "\Gui\AutomaticBrake 0.bmp")
Let PictureBoxIndependentBrake.Picture = LoadPicture(App.Path + "\Gui\IndependentBrake9.bmp")
Let PictureBoxResetRight.Picture = LoadPicture(App.Path + "\Gui\ResetRightOff.bmp")
Let PictureBoxLight.Picture = LoadPicture(App.Path + "\Gui\Light0.bmp")
Let PictureBoxPhone.Picture = LoadPicture(App.Path + "\Gui\PhoneOnHook.bmp")

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



Private Sub LabelCombinedPower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Combined Power Message"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenLeft(1).Top) + (Val(LabelCombinedPower.Top) / 15) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenLeft(1).Left) + (Val(LabelCombinedPower.Left) / 15) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub menuVideoOne_Click()

End Sub

Private Sub menuCaptureDeviceAudioSetting_Click()
    VideoCapture.ShowDlgAudioFormat
End Sub


Private Sub menuVideoTwo_Click()
VideoCapture.ShowDlgVideoDisplay
End Sub


Private Sub menuCaptureDeviceVideoCompression_Click()
VideoCapture.ShowDlgCompressionOptions
End Sub

Private Sub menuCaptureDeviceVideoDisplay_Click()
VideoCapture.ShowDlgVideoDisplay
End Sub

Private Sub menuCaptureDeviceVideoFormat_Click()
VideoCapture.ShowDlgVideoFormat
End Sub

Private Sub menuCaptureDeviceVideoSource_Click()
VideoCapture.ShowDlgVideoSource
End Sub

Private Sub PictureBoxAutomaticBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(TextBoxStatusAutomaticBrake.Text) < 12 Then
        Let TextBoxStatusAutomaticBrake.Text = Val(TextBoxStatusAutomaticBrake.Text) + 1
        If TextBoxStatusAutomaticBrake.Text = 1 Then
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\automatic.wav"
        Else
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = 1
    End If
End If

If Button = vbRightButton Then
    If Val(TextBoxStatusAutomaticBrake.Text) > 0 Then
        Let TextBoxStatusAutomaticBrake.Text = Val(TextBoxStatusAutomaticBrake.Text) - 1
        If TextBoxStatusAutomaticBrake.Text = 0 Then
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\automatic_close.wav"
        Else
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = 1
    End If
End If
Let temp = App.Path$
Let temp = temp + "\Gui\AutomaticBrake"
Let temp = temp + Right$(Str$(TextBoxStatusAutomaticBrake.Text), 2)
Let temp = temp + ".bmp"

Let PictureBoxAutomaticBrake.Picture = LoadPicture(temp)




If TextBoxStatusThrottle.Text <> 0 And _
   TextBoxStatusReverser.Text <> 1 Then
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
    Let MainlineOperationGUI!TimerDeadman.Interval = 32000
    Let MainlineOperationGUI!TimerDeadman.Enabled = True
Else
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
End If

End Sub

Private Sub PictureBoxAutomaticBrake_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Automatic Brakes"
    Let LabelCaption.Top = Val(PictureBoxAutomaticBrake.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxAutomaticBrake.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub PictureBoxBell_Click()

If TextBoxStatusBell.Text = "Off" Then
    Let TextBoxStatusBell.Text = "On"
    Let temp = App.Path$
    Let temp = temp + "/gui/BellOn.bmp"
    Let PictureBoxBell.Picture = LoadPicture(temp)
    
    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
        Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
    Else
        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        Else
            If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "3" Then
                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
            Else
                If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
                    Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                Else
                    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
                        Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                    Else
                        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
                            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                        Else
                            If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                            Else
                                If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
                                    Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                                Else
                                    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
                                        Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
   
Else
    Let TextBoxStatusBell.Text = "Off"
    Let temp = App.Path$
    Let temp = temp + "/gui/BellOff.bmp"
    Let PictureBoxBell.Picture = LoadPicture(temp)
    
    
    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
        Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
    Else
        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        Else
            If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "3" Then
                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
            Else
                If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
                    Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                Else
                    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
                        Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                    Else
                        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
                            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                        Else
                            If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                            Else
                                If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
                                    Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                                Else
                                    If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
                                        Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End If


    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
    

Call SetSoundDecoderNumber
Call SetFunction01234
Call SendCommand
DoEvents
Call SetSoundDecoderNumber
Call SetFunction5678
Call SendCommand

End Sub

Private Sub PictureBoxBell_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = Asc("W") Then Let PictureBoxBell.Top = Val(PictureBoxBell.Top) - 1
If KeyCode = Asc("A") Then Let PictureBoxBell.Left = Val(PictureBoxBell.Left) - 1
If KeyCode = Asc("S") Then Let PictureBoxBell.Left = Val(PictureBoxBell.Left) + 1
If KeyCode = Asc("Z") Then Let PictureBoxBell.Top = Val(PictureBoxBell.Top) + 1

End Sub

Private Sub PictureBoxBell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Bell Switch"
    Let LabelCaption.Top = Val(PictureBoxBell.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxBell.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub PictureBoxComputerScreenLeft_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxComputerScreenLeft(Index).Left = Val(PictureBoxComputerScreenLeft(Index).Left) - 1
If KeyAscii = Asc("S") Then PictureBoxComputerScreenLeft(Index).Left = Val(PictureBoxComputerScreenLeft(Index).Left) + 1
If KeyAscii = Asc("W") Then PictureBoxComputerScreenLeft(Index).Top = Val(PictureBoxComputerScreenLeft(Index).Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxComputerScreenLeft(Index).Top = Val(PictureBoxComputerScreenLeft(Index).Top) + 1

Debug.Print "Left is "; PictureBoxComputerScreenLeft(Index).Left
Debug.Print "Top is "; PictureBoxComputerScreenLeft(Index).Top

End Sub

Private Sub PictureBoxComputerScreenLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Left Computer Screen" + Str$(Index)
    Let LabelCaption.Top = Val(PictureBoxComputerScreenLeft(Index).Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenLeft(Index).Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub PictureBoxComputerScreenRight_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxComputerScreenRight(Index).Left = Val(PictureBoxComputerScreenRight(Index).Left) - 1
If KeyAscii = Asc("S") Then PictureBoxComputerScreenRight(Index).Left = Val(PictureBoxComputerScreenRight(Index).Left) + 1
If KeyAscii = Asc("W") Then PictureBoxComputerScreenRight(Index).Top = Val(PictureBoxComputerScreenRight(Index).Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxComputerScreenRight(Index).Top = Val(PictureBoxComputerScreenRight(Index).Top) + 1

Debug.Print "Left is "; PictureBoxComputerScreenRight(Index).Left
Debug.Print "Top is "; PictureBoxComputerScreenRight(Index).Top

End Sub

Private Sub PictureBoxComputerScreenRight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Right Computer Screen" + Str$(Index)
    Let LabelCaption.Top = Val(PictureBoxComputerScreenRight(Index).Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxComputerScreenRight(Index).Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If
End Sub


Private Sub PictureBoxHorn_Click()

If TextBoxStatusHorn.Text = "Off" Then
    Let TextBoxStatusHorn.Text = "On"
    Let temp = App.Path
    Let temp = temp + "/gui/HornOn.bmp"
    Let PictureBoxHorn.Picture = LoadPicture(temp)
    
    If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "1" Then
        Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
    Else
        If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        Else
            If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "3" Then
                Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
            Else
                If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "4" Then
                    Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                Else
                    If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "5" Then
                        Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                    Else
                        If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "6" Then
                            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                        Else
                            If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "7" Then
                                Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                            Else
                                If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "8" Then
                                    Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                                Else
                                    If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "9" Then
                                        Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
   
Else
    Let TextBoxStatusHorn.Text = "Off"
    Let temp = App.Path
    Let temp = temp + "/gui/HornOff.bmp"
    Let PictureBoxHorn.Picture = LoadPicture(temp)
    
    
    If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "1" Then
        Let MainlineOperationGUI.ConsistControlFunction1.Value = vbUnchecked
    Else
        If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI.ConsistControlFunction2.Value = vbUnchecked
        Else
            If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "3" Then
                Let MainlineOperationGUI.ConsistControlFunction3.Value = vbUnchecked
            Else
                If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "4" Then
                    Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
                Else
                    If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "5" Then
                        Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
                    Else
                        If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "6" Then
                            Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
                        Else
                            If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "7" Then
                                Let MainlineOperationGUI.ConsistControlFunction7.Value = vbUnchecked
                            Else
                                If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "8" Then
                                    Let MainlineOperationGUI.ConsistControlFunction8.Value = vbUnchecked
                                'Else
                                '    If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "9" Then
                                '        Let MainlineOperationGUI.ConsistControlFunction9.Value = vbUnchecked
                                '    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End If


    Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
    Let MainlineOperationGUI!Wave1.Action = 1
    

Call SetSoundDecoderNumber
Call SetFunction01234
Call SendCommand
DoEvents
Call SetFunction5678
Call SendCommand

End Sub

Private Sub PictureBoxHorn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Horn Switch"
    Let LabelCaption.Top = Val(PictureBoxHorn.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxHorn.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub PictureBoxIndependentBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(TextBoxStatusIndependentBrake.Text) < 9 Then
        Let TextBoxStatusIndependentBrake.Text = Val(TextBoxStatusIndependentBrake.Text) + 1
        If TextBoxStatusIndependentBrake.Text = 1 Then
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\independent.wav"
        Else
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = 1
    End If
End If

If Button = vbRightButton Then
    If Val(TextBoxStatusIndependentBrake.Text) > 0 Then
        Let TextBoxStatusIndependentBrake.Text = Val(TextBoxStatusIndependentBrake.Text) - 1
        If TextBoxStatusIndependentBrake.Text = 0 Then
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\independent_close.wav"
        Else
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = 1
    End If
End If

Let temp = App.Path$
Let temp = temp + "/Gui/IndependentBrake"
Let temp = temp + Right$(Str$(TextBoxStatusIndependentBrake.Text), 1)
Let temp = temp + ".bmp"

Let PictureBoxIndependentBrake.Picture = LoadPicture(temp)





If TextBoxStatusThrottle.Text <> 0 And _
   TextBoxStatusReverser.Text <> 1 Then
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
    Let MainlineOperationGUI!TimerDeadman.Interval = 32000
    Let MainlineOperationGUI!TimerDeadman.Enabled = True
Else
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
End If

End Sub

Private Sub PictureBoxIndependentBrake_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Independent Brakes"
    Let LabelCaption.Top = Val(PictureBoxIndependentBrake.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxIndependentBrake.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub

Private Sub PictureBoxLight_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = Asc("W") Then Let PictureBoxLight.Top = Val(PictureBoxLight.Top) - 1
If KeyCode = Asc("A") Then Let PictureBoxLight.Left = Val(PictureBoxLight.Left) - 1
If KeyCode = Asc("S") Then Let PictureBoxLight.Left = Val(PictureBoxLight.Left) + 1
If KeyCode = Asc("Z") Then Let PictureBoxLight.Top = Val(PictureBoxLight.Top) + 1

End Sub

Private Sub PictureBoxLight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(TextBoxStatusLight.Text) < 2 Then
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        Let MainlineOperationGUI!Wave1.Action = 1
        Let TextBoxStatusLight.Text = Val(TextBoxStatusLight.Text) + 1
    End If
End If

If Button = vbLeftButton Then
    If Val(TextBoxStatusLight.Text) > 0 Then
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
        Let MainlineOperationGUI!Wave1.Action = 1
        Let TextBoxStatusLight.Text = Val(TextBoxStatusLight.Text) - 1
    End If
End If

Let temp = App.Path
Let temp = temp + "/Gui/Light"
Let temp = temp + Right$(Str$(TextBoxStatusLight.Text), 1)
Let temp = temp + ".bmp"

Let PictureBoxLight.Picture = LoadPicture(temp)

If TextBoxStatusLight.Text = "0" Then

    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
        Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
    Else
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        Else
            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
                    Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
                        Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
                            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
                                    Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
                                        Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
                                            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
             End If
        End If
    End If

    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
        Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
    Else
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        Else
            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
                Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
                    Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
                        Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
                            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
                                Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
                                    Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
                                        Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
                                            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                 End If
            End If
        End If
    End If
Else
    If TextBoxStatusLight.Text = "1" Then
            
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        Else
            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
                    Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
                        Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
                            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
                                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
                                        Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
                                            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                                        Else
                                            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
                                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                 End If
            End If
        End If
    
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        Else
            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
                Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
                    Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
                        Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
                            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
                                Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
                                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
                                        Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
                                            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                                        Else
                                            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
                                                Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                     End If
                End If
            End If
        End If
    
    Else
        
        If TextBoxStatusLight.Text = "2" Then
        
            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
                    Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
                        Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
                            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
                                    Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
                                        Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
                                            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
                                        Else
                                            If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
                                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
                                            Else
                                                If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
                                                    Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                     End If
                End If
            End If
    
            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
                Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
            Else
                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
                    Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
                Else
                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
                        Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
                    Else
                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
                            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
                        Else
                            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
                                Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                            Else
                                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
                                    Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
                                Else
                                    If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
                                        Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                                    Else
                                        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
                                            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
                                        Else
                                            If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
                                                Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
                                            Else
                                                If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
                                                    Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                         End If
                    End If
                End If
            End If
        End If
    End If
End If

Call SetLocomotiveNumber
Call SetFunction01234
Call SendCommand
DoEvents
Call SetFunction5678
Call SendCommand

End Sub


Private Sub PictureBoxLight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Light Switch"
    Let LabelCaption.Top = Val(PictureBoxLight.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxLight.Left) - Val(LabelCaption.Width) + Int(X / 15) - 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxPhone_Click()

If Val(TextBoxStatusPhone.Text) = 0 Then
    Let TextBoxStatusPhone.Text = 1
    Let PictureBoxPhone.Picture = LoadPicture(App.Path + "\Gui\PhoneOnHook.bmp")
Else
    Let TextBoxStatusPhone.Text = 0
    Let PictureBoxPhone.Picture = LoadPicture(App.Path + "\Gui\PhoneOffHook.bmp")
End If

End Sub

Private Sub PictureBoxPhone_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = Asc("W") Then Let PictureBoxPhone.Top = Val(PictureBoxPhone.Top) - 1
If KeyCode = Asc("A") Then Let PictureBoxPhone.Left = Val(PictureBoxPhone.Left) - 1
If KeyCode = Asc("S") Then Let PictureBoxPhone.Left = Val(PictureBoxPhone.Left) + 1
If KeyCode = Asc("Z") Then Let PictureBoxPhone.Top = Val(PictureBoxPhone.Top) + 1

End Sub


Private Sub PictureBoxPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Radio Phone"
    Let LabelCaption.Top = Val(PictureBoxPhone.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxPhone.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxResetLeft_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxResetLeft.Left = Val(PictureBoxResetLeft.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxResetLeft.Left = Val(PictureBoxResetLeft.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxResetLeft.Top = Val(PictureBoxResetLeft.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxResetLeft.Top = Val(PictureBoxResetLeft.Top) + 1

End Sub

Private Sub PictureBoxResetLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Let TextBoxStatusResetLeft.Text = "On"

Let MainlineOperationGUI!wave2.Action = 4

' =========================================================================================================================
' Clicking the Reset button with sound

Let temp = App.Path
Let temp = temp + "/Gui/ResetLeftOn.bmp"
Let PictureBoxResetLeft.Picture = LoadPicture(temp)
    
Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
Let MainlineOperationGUI!Wave1.Action = 1
    
' =========================================================================================================================
' End of Routin to Reset Button
 
Let labeldeadmann.Visible = False
      
If TextBoxStatusThrottle.Text <> 0 Then
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
    Let MainlineOperationGUI!TimerDeadman.Interval = 32000
    Let MainlineOperationGUI!TimerDeadman.Enabled = True
Else
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
End If
    
End Sub

Private Sub PictureBoxResetLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Deadmann Switch"
    Let LabelCaption.Top = Val(PictureBoxResetLeft.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxResetLeft.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxResetLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Let temp = App.Path
Let temp = temp + "/Gui/ResetLeftOff.bmp"
Let PictureBoxResetLeft.Picture = LoadPicture(temp)

Let TextBoxStatusResetLeft.Text = "Off"

End Sub

Private Sub PictureBoxResetRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let TextBoxStatusResetRight.Text = "On"
    Let PictureBoxResetRight.Picture = LoadPicture("c:/automatic train Control/gui/ResetRightOn.bmp")
    
    Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
    ' Set Speed to Emergency Stop
    
            Call SetLocomotiveNumber
            Call SetSpeed
            Call SendCommand
            
            Let TextBoxStatusSpeedStep = 0
            Let TextBoxStatusSpeedStepNow = 0
            Let TextBoxStatusSpeedStepModified = 0
            
             
ONEMORE:
        If Val(TextBoxStatusThrottle.Text) > 0 Then
            Let TextBoxStatusThrottle.Text = Val(TextBoxStatusThrottle.Text) - 1
        End If
        If Val(TextBoxStatusThrottle.Text) < 0 Then
            Let TextBoxStatusThrottle.Text = Val(TextBoxStatusThrottle.Text) + 1
        End If
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
            Let MainlineOperationGUI!Wave1.Action = 1
            
            If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                If Val(TextBoxStatusThrottle.Text) >= 0 And _
                   Val(TextBoxStatusThrottle.Text) < 7 Then
                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                    DoEvents
                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                End If
            
                If Val(TextBoxStatusThrottle.Text) = 0 Then
                    Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
                    Call SetSoundDecoderNumber
                    Call SetFunction01234
                    Call SendCommand
                End If
            End If
            
   
    Let temp = App.Path
    If Val(TextBoxStatusThrottle.Text) < 0 Then
        Let temp = temp + "\Gui\DynamicBrake"
    Else
        Let temp = temp + "\Gui\Throttle"
    End If
    Let temp = temp + Right$(Str$(TextBoxStatusThrottle.Text), 1)
    Let temp = temp + ".bmp"

    Let PictureBoxThrottle.Picture = LoadPicture(temp)
    
    DoEvents
            
            
            
If Val(TextBoxStatusThrottle.Text) <> 0 Then GoTo ONEMORE

End Sub


Private Sub PictureBoxResetRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "DCC Emergency Stop"
    Let LabelCaption.Top = Val(PictureBoxResetRight.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxResetRight.Left) - Val(LabelCaption.Width) + Int(X / 15) - 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxResetRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let TextBoxStatusResetRight.Text = "Off"
    Let PictureBoxResetRight.Picture = LoadPicture("c:/automatic train Control/gui/ResetRightOff.bmp")

End Sub

Private Sub PictureBoxReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Val(TextBoxStatusThrottle.Text) = 0 Then

    If Val(TextBoxStatusSpeedStepModified.Text) = 0 Then

        If Button = vbLeftButton Then
            If Val(TextBoxStatusReverser.Text) < 2 Then
                Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\reverser.wav"
                Let MainlineOperationGUI!Wave1.Action = 1
                Let TextBoxStatusReverser.Text = Val(TextBoxStatusReverser.Text) + 1
            End If
        End If

        If Button = vbRightButton Then
            If Val(TextBoxStatusReverser.Text) > 0 Then
                Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\reverser.wav"
                Let MainlineOperationGUI!Wave1.Action = 1
                Let TextBoxStatusReverser.Text = Val(TextBoxStatusReverser.Text) - 1
            End If
        End If

        Let temp = App.Path$
        Let temp = temp + "/gui/Reverser"
        Let temp = temp + Right$(Str$(TextBoxStatusReverser.Text), 1)
        Let temp = temp + ".bmp"

        Let PictureBoxReverser.Picture = LoadPicture(temp)

        If Val(TextBoxStatusReverser.Text) = 2 Then
            Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbChecked
            Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
            Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
        End If
        If Val(TextBoxStatusReverser.Text) = 1 Then
            Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
            Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbChecked
            Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
        End If
        If Val(TextBoxStatusReverser.Text) = 0 Then
            Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
            Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
            Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked
        End If
    Else
        If MainlineOperationGUIScreen.ButtonHelp.Caption = "Help is On" Then
            Let message = "You cannot move the reverser handle unless the speed of the locomotive is nill."
            MsgBox message, vbExclamation, "Engineer Error - Control Interlock"
        End If
    End If
Else
    If MainlineOperationGUIScreen.ButtonHelp.Caption = "Help is On" Then
        Let message = "You cannot move the reverser handle onless the trottle is return to notch zero."
        MsgBox message, vbExclamation, "Engineer Error - Control Interlock"
    End If
End If

End Sub

Private Sub PictureBoxReverser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Reverser Handle"
    Let LabelCaption.Top = Val(PictureBoxReverser.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxReverser.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxSand_Click()

If TextBoxStatusSand.Text = "Off" Then
    Let TextBoxStatusSand.Text = "On"
    Let temp = App.Path$
    Let temp = temp + "/gui/SandOn.bmp"
    Let PictureBoxSand.Picture = LoadPicture(temp)
    Let MainlineOperationGUI!TimerSandConsumption.Interval = 2500
Else
    Let TextBoxStatusSand.Text = "Off"
    Let temp = App.Path$
    Let temp = temp + "/gui/SandOff.bmp"
    Let PictureBoxSand.Picture = LoadPicture(temp)
    Let MainlineOperationGUI!TimerSandConsumption.Interval = 0
End If

Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\control.wav"
Let MainlineOperationGUI!Wave1.Action = 1
    
End Sub

Private Sub PictureBoxSand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Sanding Switch"
    Let LabelCaption.Top = Val(PictureBoxSand.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxSand.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxSpeedometer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Speedometer"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenLeft(1).Top) + Val(PictureBoxSpeedometer.Top) + Int(Y / 15) - 15
    Let LabelCaption.Left = Val(PictureBoxComputerScreenLeft(1).Left) + Val(PictureBoxSpeedometer.Left) + Int(X / 15) + 20 + 15
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub PictureBoxThrottle_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxThrottle.Left = Val(PictureBoxThrottle.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxThrottle.Left = Val(PictureBoxThrottle.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxThrottle.Top = Val(PictureBoxThrottle.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxThrottle.Top = Val(PictureBoxThrottle.Top) + 1

End Sub

Private Sub PictureBoxThrottle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If labeldeadmann.Visible = False Then

If Val(TextBoxStatusThrottleDelay.Text) = 0 Then
' Only allow the throttle handle to be moved if the delay is zero, this makes the engineer stop from moving the
' throttle handle too quickly.

    If Button = vbLeftButton Then

        Let OkToNotchUp = True
    
        If Val(TextBoxStatusThrottle.Text) > 7 Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "The maximum engine revolution has been reached with the throttle in notch eight."
                MsgBox message, vbExclamation, "Engineer Error - Maximum Throttle"
            End If
        End If
        
        If Val(BarFuelTank.Value) < Val(TextBoxStatusFuelTooLow.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Please call your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                Let message = message + "of diesel fuel. You currently have less then " + Str$(Val(TextBoxStatusFuelTooLow.Text))
                Let message = message + " US gallons left."
                MsgBox message, vbExclamation, "Engineer Error - Fuel Reserve"
            End If
        End If
        
        If Val(BarSandTank.Value) < Val(textboxstatussandtoolow.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Please call your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                Let message = message + "of sand. You currently have " + Str$(Val(BarSandTank.Value)) + " pounds of sand left. It is recommended that you have at" + Chr$(13)
                Let message = message + "least " + Str$(Val(textboxstatussandtoolow.Text)) + " pounds of sand."
                MsgBox message, vbExclamation, "Engineer Error - Sand Reserve"
            End If
        End If
        
        If Val(BarWaterTank.Value) < Val(TextBoxStatusWaterTooLow.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Please call your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                Let message = message + "of water. You currently have " + Str$(Val(BarWaterTank.Value)) + " US gallons of water left. It is recommended that you have at" + Chr$(13)
                Let message = message + "least " + Str$(Val(TextBoxStatusWaterTooLow.Text)) + " US gallons of water."
                MsgBox message, vbExclamation, "Engineer Error - Water Reserve"
            End If
        End If
        
        If Val(BarOilTank.Value) < Val(TextBoxStatusOilTooLow.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Please call your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                Let message = message + "of engine lubricating oil. You currently have " + Str$(Val(BarOilTank.Value)) + " US gallons of lubricating left. It is" + Chr$(13)
                Let message = message + "recommended that you have at least " + Str$(Val(TextBoxStatusOilTooLow.Text)) + " US gallons of oil. "
                MsgBox message, vbExclamation, "Engineer Error - Oil Reserve"
            End If
        End If
        
        If Val(BarOilPressure.Value) < Val(TextBoxStatusOilTooLowPressure.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "The engine oil pressure has not reached a pressure to operate the" + Chr$(13)
                Let message = message + "locomotive safely. The current oil pressure is" + Str$(Val(BarOilPressure.Value)) + " pounds per square" + Chr$(13)
                Let message = message + "inch. We recommend that you wait until the pressure reaches at" + Chr$(13)
                Let message = message + "least" + Str$(Val(TextBoxStatusOilTooLowPressure.Text)) + " pounds per square inch."
                MsgBox message, vbExclamation, "Engineer Error - Oil Pressure"
            End If
        End If
        
        If Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text) * 2 / 3 Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "The engine oil temperature has not reached a temperature to operate the" + Chr$(13)
                Let message = message + "locomotive safely. The current oil temperature is" + Str$(Val(BarOilTemperature.Value)) + " degrees. We recommend" + Chr$(13)
                Let message = message + "that you wait until the temperature reaches at least" + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees."
                MsgBox message, vbExclamation, "Engineer Error- Oil Temperature"
            End If
        End If
        
        If Val(BarWaterPressure.Value) < Val(TextBoxStatusWaterTooLowPressure.Text) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "The engine water pressure has not reached a pressure to operate the " + Chr$(13)
                Let message = message + "locomotive safely. The current water pressure is " + Str$(Val(BarWaterPressure.Value)) + " pounds per square inch. We" + Chr$(13)
                Let message = message + "recommend that you wait until the pressure reaches" + Str$(Val(TextBoxStatusWaterTooLowPressure.Text)) + " pounds per square inch."
                MsgBox message, vbExclamation, "Engineer Error - Water Pressure"
            End If
        End If
        
        If Val(BarWaterTemperature.Value) < Val(TextBoxStatusWaterTooLowTemperature.Text) * 2 / 3 Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "The engine water temperature has not reached a temperature to operate the " + Chr$(13)
                Let message = message + "locomotive safely. The current water temperature is " + Str$(Val(BarWaterTemperature.Value)) + " degrees. We" + Chr$(13)
                Let message = message + "recommend that you wait until the temperature reaches" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees."
                MsgBox message, vbExclamation, "Engineer Error - Water Temperature"
            End If
        End If
        
        If ((Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text)) And _
            (Val(TextBoxStatusThrottle.Text) >= 1)) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Although the engine oil temperature is greater than" + Str$(Val(BarOilTemperature.Value) - 30) + "degrees, it has not reached" + Chr$(13)
                Let message = message + "a temperature to operate above notch one. It is recommended that the oil temerature be greater than" + Chr$(13)
                Let message = message + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                MsgBox message, vbExclamation, "Engineer Error- Oil Temperature"
            End If
        End If
        
        If ((Val(BarWaterTemperature.Value) <= Val(TextBoxStatusWaterTooLowTemperature.Text)) And _
            (Val(TextBoxStatusThrottle.Text) >= 2)) Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Although the engine water temperature is greater than" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text) - 30) + " degrees, it has not reached" + Chr$(13)
                Let message = message + "a temperature to operate above notch two. It is recommended that the water engine temperature be greater than" + Chr$(13)
                Let message = message + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                MsgBox message, vbExclamation, "Engineer Error - Water Temperature"
            End If
        End If
        
        If Val(TextBoxStatusLight.Text) = 0 And Val(TextBoxStatusReverser.Text) <> 1 Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Just a little reminder that you should put the headlight on before moving the locomotive" + Chr$(13)
                Let message = message + "in either direction. Please remeber Rule 17 the requires to to dim the lights in the yard and " + Chr$(13)
                Let message = message + "and when approaching another locomotive."
                MsgBox message, vbExclamation, "Engineer Error - Lights"
            End If
        End If
        
        If TextBoxStatusBell.Text = "Off" And Val(TextBoxStatusThrottle.Text) = 0 And Val(TextBoxStatusReverser.Text) <> 1 Then
            Let OkToNotchUp = False
            If ButtonHelp.Caption = "Help is On" Then
                Let message = "Just another little reminder that as an engineer, you should be activating the bell prior" + Chr$(13)
                Let message = message + " moving in any direction."
                MsgBox message, vbExclamation, "Engineer Error - Bell Activation"
            End If
        End If
        
        ' =======================================================================================================================
        '
            
        If OkToNotchUp = True Then
       
            Let TextBoxStatusThrottle.Text = Val(TextBoxStatusThrottle.Text) + 1
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
            Let MainlineOperationGUI!Wave1.Action = 1
            If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                If Val(TextBoxStatusThrottle.Text) > 0 And _
                   Val(TextBoxStatusThrottle.Text) < 8 Then
                    Let MainlineOperationGUI.ConsistControlFunction5.Value = vbChecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                    DoEvents
                    Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                Else
                    If Val(TextBoxStatusThrottle.Text) = 0 Then
                        Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
                        Call SetSoundDecoderNumber
                        Call SetFunction01234
                        Call SendCommand
                    End If
                End If
            End If
        End If
    End If
    ' =====================================================================================================================
    '
    
    If Button = vbRightButton Then

        Let OKToNotchDown = True
        
        If Val(TextBoxStatusThrottle.Text) < -7 Then
            Let OKToNotchDown = False
            If ButtonHelp.Caption = "Help is On" Then
                MsgBox "The maximum application has been reached with the dynamic brake in notch eight.", vbOKOnly, "Engineer Error"
            End If
        End If
        
        'If Val(TextBoxStatusThrottle.Text) < 1 Then
        '    If Val(TextBoxStatusReverser.Text) <> 1 Then
        '        Let OKToNotchDown = False
        '        If ButtonHelp.Caption = "Help is On" Then
        '            Let Message = "The dynamic brake can only be engaged when the reverser handle in in nuetral."
        '            MsgBox Message, vbExclamation, "Engineer Error - Reverser Handle and Dynamic Brake"
        '        End If
        '    End If
        'End If
        
        ' =================================================================================================================
        '
        
        If OKToNotchDown = True Then
        
            Let TextBoxStatusThrottle.Text = Val(TextBoxStatusThrottle.Text) - 1
            Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Gui\throttle.wav"
            Let MainlineOperationGUI!Wave1.Action = 1
            
            If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                If Val(TextBoxStatusThrottle.Text) >= 0 And _
                   Val(TextBoxStatusThrottle.Text) < 7 Then
                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                    DoEvents
                    Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
                    Call SetSoundDecoderNumber
                    Call SetFunction5678
                    Call SendCommand
                End If
            Else
                If Val(TextBoxStatusThrottle.Text) = -1 Then
                    Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
                    Call SetSoundDecoderNumber
                    Call SetFunction01234
                    Call SendCommand
                End If
            End If
        End If
    End If
        
    ' =====================================================================================================================
    '
        
    Let temp = App.Path
    If Val(TextBoxStatusThrottle.Text) < 0 Then
        Let temp = temp + "\Gui\DynamicBrake"
    Else
        Let temp = temp + "\Gui\Throttle"
    End If
    Let temp = temp + Right$(Str$(TextBoxStatusThrottle.Text), 1)
    Let temp = temp + ".bmp"

    Let PictureBoxThrottle.Picture = LoadPicture(temp)

' =========================================================================================================================
' Checking the reverser handle
' In Idle Positions
'
    If TextBoxStatusReverser.Text <> "1" Then
    ' Only send a speed change if the reverser handle is not in nuetral.
        If Val(TextBoxStatusThrottle.Text) >= 0 Then
            If MainlineOperationGUI!ConsistControlSpeed128 = vbChecked Then
                Let TextBoxStatusSpeedStep.Text = Int(Val(TextBoxStatusThrottle.Text) / 8 * 126)
            Else
                If MainlineOperationGUI!ConsistControlSpeed28 = vbChecked Then
                    Let TextBoxStatusSpeedStep.Text = Int(Val(TextBoxStatusThrottle.Text) / 8 * 28)
                Else
                    Let TextBoxStatusSpeedStep.Text = Int(Val(TextBoxStatusThrottle.Text) / 8 * 14)
                End If
            End If
        End If
    End If

    Let TextBoxStatusThrottleDelay.Text = 5
Else
    If ButtonHelp.Caption = "Help is On" Then
        MsgBox "You are moving the throttle handle too quickly between" + Chr$(13) + "notches. General Motors, Electromotive Division" + Chr$(13) + "recommends ten seconds between movements of the throttle", vbOKOnly, "Engineer Error"
    End If
End If

' =========================================================================================================================
' Deadmann Switch Setting
'

If TextBoxStatusThrottle.Text <> 0 And _
   TextBoxStatusReverser.Text <> 1 Then
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
    Let MainlineOperationGUI!TimerDeadman.Interval = 32000
    Let MainlineOperationGUI!TimerDeadman.Enabled = True
Else
    Let MainlineOperationGUI!TimerDeadman.Enabled = False
End If
Else
    If ButtonHelp.Caption = "Help is On" Then
        Let message = "Your locomotive was shut down becuase of the elapsed time of the deadman alarm. Please reset" + Chr$(13)
        Let message = message + "the dead man switch before proceeding."
        MsgBox message, vbExclamation, "Engineer Erorr - Dead man Switch"
    End If
End If

End Sub
Private Sub PictureBoxThrottle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "Throttle"
    Let LabelCaption.Top = Val(PictureBoxThrottle.Top) + Int(Y / 15)
    Let LabelCaption.Left = Val(PictureBoxThrottle.Left) + Int(X / 15) + 20
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub TextBoxDigitalSpeed_Change()

' Draws a circle, ellipse, or arc on an object.  Doesn't support named arguments.
'
'Syntax'
'
'object.Circle Step (x, y), radius, color, start, end, aspect
'
'The Circle method syntax has the following parts.
'
'Part Description
'
'object  Optional.  An object expression that evaluates to an object in the Applies To list.  If object is omitted, the Form with the focus is assumed to be object.
'Step    Optional.  A keyword specifying that the center of the circle, ellipse, or arc is relative to the current coordinates given by the CurrentX and CurrentY properties of object.
'(x, y)  Required.  Single-precision values indicating the coordinates for the center point of the circle, ellipse, or arc.  The ScaleMode property of object determines the units of measure used.
'radius  Required.  Single-precision value indicating the radius of the circle, ellipse, or arc.  The ScaleMode property of object determines the unit of measure used.
'
'color   Optional.  Long integer value indicating the RGB color of the circle's outline.  If omitted, the value of the
'ForeColor property is used.  You can use the RGB function or QBColor function to specify the color.
'start, end  Optional.  Single-precision values.  When an arc or a partial circle or ellipse is drawn, start and end specify (in radians) the beginning and end positions of the arc.  The range for both is -2 pi radians to 2 pi radians.  The default value for start is 0 radians; the default for end is 2 * pi radians.
'aspect  Optional.  Single-precision value indicating the aspect ratio of the circle .  The default value is 1.0, which yields a perfect (non-elliptical) circle on any screen.
'
'Remarks
'
'To fill a circle, set the FillColor and FillStyle properties of the object on which the circle or ellipse is drawn.  Only a closed figure can be filled.  Closed figures include circles, ellipses, or pie slices (arcs with radius lines drawn at both ends).
'When drawing a partial circle or ellipse, if start is negative, Circle draws a radius to start, and treats the angle as positive; if end is negative, Circle draws a radius to end and treats the angle as positive.  The Circle method always draws in a counter-clockwise (positive) direction.
'
'The width of the line used to draw the circle, ellipse, or arc depends on the setting of the DrawWidth property.  The way the circle is drawn on the background depends on the setting of the DrawMode and DrawStyle properties.
'When drawing pie slices, to draw a radius to angle 0 (giving a horizontal line segment to the right), specify a very small negative value for start, rather than zero.
'You can omit an argument in the middle of the syntax, but you must include the argument's comma before including the next argument.  If you omit an optional trailing argument, don't use any commas following the last argument you specify.
'When Circle executes, the CurrentX and CurrentY properties are set to the center point specified by the arguments.'

PictureBoxSpeedometer.DrawWidth = 5

If Val(TextBoxDigitalSpeed.Text) <> 0 Then
    If Val(TextBoxDigitalSpeed.Text) < 10 Then
        Let startpoint = Int(234 - (Val(TextBoxDigitalSpeed.Text) / 12.8 * 288))
    Else
        Let startpoint = Int(234 - (Val(TextBoxDigitalSpeed.Text) / 128 * 288))
    End If
    
    If startpoint < 0 Then
        Let startpoint = 360 + startpoint
    End If
    
    Let EndPoint = 234
    
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, , (320 / 57.3), (234 / 57.3), 1
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, &HC0&, (startpoint / 57.3), (EndPoint / 57.3), 1
Else
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, , (320 / 57.3), (234 / 57.3), 1
End If

End Sub


Private Sub TextBoxDigitalSpeed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ButtonCaption.Caption = "&Caption is On" Then
    If MainlineOperationGUI!timerlabel.Enabled = False Then
        Let MainlineOperationGUI!timerlabel.Interval = 1000
        Let MainlineOperationGUI!timerlabel.Enabled = True
    End If
    Let LabelCaption.Text = "DCC Speed Step"
    Let LabelCaption.Top = Val(PictureBoxComputerScreenLeft(1).Top) + Val(PictureBoxSpeedometer.Top) + (Val(TextBoxDigitalSpeed.Top) / 15) + Int(Y / 15) - 15
    Let LabelCaption.Left = Val(PictureBoxComputerScreenLeft(1).Left) + Val(PictureBoxSpeedometer.Left) + (Val(TextBoxDigitalSpeed.Left) / 15) + Int(X / 15) + 20 + 15
    Let LabelCaption.Visible = True
End If

End Sub


Private Sub TextBoxStatusReverser_Change()

Let LabelCombinedPower.Caption = "Combined Power" + Chr$(13)

If Val(TextBoxStatusReverser.Text) = 0 Then
    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Forward - "
Else
    If Val(TextBoxStatusReverser.Text) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Reverse - "
    End If
End If

If Val(TextBoxStatusThrottle.Text) < 0 Then
    Let lebelcombinedPower.Caption = LabelCombinedPower.Caption + "Dynamic Brake" + Chr(13)
End If

If Abs(TextBoxStatusThrottle.Text) = 0 Then
    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Zero"
Else
    If Abs(TextBoxStatusThrottle.Text) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch One"
    Else
        If Abs(TextBoxStatusThrottle.Text) = 2 Then
            Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Two"
        Else
            If Abs(TextBoxStatusThrottle.Text) = 3 Then
                Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Three"
            Else
                If Abs(TextBoxStatusThrottle.Text) = 4 Then
                    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Four"
                Else
                    If Abs(TextBoxStatusThrottle.Text) = 5 Then
                        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Five"
                    Else
                        If Abs(TextBoxStatusThrottle.Text) = 6 Then
                            Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Six"
                        Else
                            If Abs(TextBoxStatusThrottle.Text) = 7 Then
                                Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Seven"
                            Else
                                If Abs(TextBoxStatusThrottle.Text) = 8 Then
                                    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Eight"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub TextBoxStatusSpeedStepModified_Change()

If TextBoxStatusSpeedStepModified.Text < 10 Then

    Let TextBoxDigitalSpeed.Text = TextBoxStatusSpeedStepModified.Text + ".0"

    Let LabelSpeed1.Caption = "1.0"
    Let LabelSpeed2.Caption = "2.0"
    Let LabelSpeed3.Caption = "3.0"
    Let LabelSpeed4.Caption = "4.0"
    Let LabelSpeed5.Caption = "5.0"
    Let LabelSpeed6.Caption = "6.0"
    Let LabelSpeed7.Caption = "7.0"
    Let LabelSpeed8.Caption = "8.0"
    Let LabelSpeed9.Caption = "9.0"
    Let LabelSpeed10.Caption = "10.0"
    
Else

    Let TextBoxDigitalSpeed.Text = TextBoxStatusSpeedStepModified.Text

    Let LabelSpeed1.Caption = "10"
    Let LabelSpeed2.Caption = "20"
    Let LabelSpeed3.Caption = "30"
    Let LabelSpeed4.Caption = "40"
    Let LabelSpeed5.Caption = "50"
    Let LabelSpeed6.Caption = "60"
    Let LabelSpeed7.Caption = "70"
    Let LabelSpeed8.Caption = "80"
    Let LabelSpeed9.Caption = "90"
    Let LabelSpeed10.Caption = "100"

End If

End Sub
Private Sub TextBoxStatusSpeedStepNow_Change()

Let temporary1 = Val(TextBoxStatusAutomaticBrakePressure.Text)
Let temporary2 = Val(TextBoxStatusIndependentBrakePressure.Text)

If temporary1 < temporary2 Then
    Let temporary = temporary1
Else
    Let temporary = temporary2
End If

If temporary > 1 Then
    Let temporary = 1
End If

Let TextBoxStatusSpeedStepModified.Text = Int(Val(TextBoxStatusSpeedStepNow.Text) * temporary)

End Sub

Private Sub TextBoxStatusThrottle_Change()

Let LabelCombinedPower.Caption = "Combined Power" + Chr$(13)

If Val(TextBoxStatusReverser.Text) = 0 Then
    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Forward - "
Else
    If Val(TextBoxStatusReverser.Text) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Reverse - "
    End If
End If

If Val(TextBoxStatusThrottle.Text) < 0 Then
    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Dynamic Brake" + Chr$(13)
End If

If Abs(TextBoxStatusThrottle.Text) = 0 Then
    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Zero"
Else
    If Abs(TextBoxStatusThrottle.Text) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch One"
    Else
        If Abs(TextBoxStatusThrottle.Text) = 2 Then
            Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Two"
        Else
            If Abs(TextBoxStatusThrottle.Text) = 3 Then
                Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Three"
            Else
                If Abs(TextBoxStatusThrottle.Text) = 4 Then
                    Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Four"
                Else
                    If Abs(TextBoxStatusThrottle.Text) = 5 Then
                        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Five"
                    Else
                        If Abs(TextBoxStatusThrottle.Text) = 6 Then
                            Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Six"
                        Else
                            If Abs(TextBoxStatusThrottle.Text) = 7 Then
                                Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Seven"
                            Else
                                Let LabelCombinedPower.Caption = LabelCombinedPower.Caption + "Notch Eight"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub






Private Sub SendCommand()

    Let MainlineOperationGUI!SevenByteD.Text = "0"

' For Next Statement
'
' Repeats a group of statements a specified number of times.
'
'
' The step argument can be either positive or negative.
' The value of the step argument determines loop processing as follows:
'
' Once the loop starts and all statements in the loop have executed, step is added to counter.
' At this point, either the statements in the loop execute again (based on the same test that caused the loop to execute
' initially), or the loop is exited and execution continues with the statement following the Next statement.
' Tip, changing the value of counter while inside a loop can make it more difficult to read and debug your code.
'   The Exit For can only be used within a For Each...Next or For...Next control structure to provide an alternate way to exit.
'   Any number of Exit For statements may be placed anywhere in the loop.
'   The Exit For is often used with the evaluation of some condition (for example, If...Then), and transfers control to the statement immediately following Next.
'   You can nest For...Next loops by placing one For...Next loop within another.
'   Give each loop a unique variable name as its counter.
'
' My Notes:
'
' For Next statement is used to set up a loop for each of the bits in a bit. I'm trying to calculate the error byte; hence,
' I need to look at each byte of the packet. Eight bits to a byte so...

    For X = 1 To 8

' Temporary Counter

' My Notes:
'
' I needed to use a temporary counter to add up all the bits. In each one of the bytes to be sent to the communication port,
' i examine the bits to see if it is one or zero. At the end of this routine, it is used to calculate the error btye. This eror byte is needed to conplete the packet.

    Let temp = 0

' My Notes:
'
' For each one of these 'if statements', we are checking to see if the byte should be sent to the communication port.
' For example, the first byte is the first byte in the locomotive address. The second byte is the second of  the locomotives
' address; which may not always be needed. There for the check is omitted.
'   Once inside the first 'if statment' we preform another 'if statement'. This statement is used to determine if the
'   bit of the byte is equal to one or zero. We are counting the number of one bits to determin the rror code.
'   If the bit is equal to one, then our temporary vaiiable is incremented by one.

    If MainlineOperationGUI!OneByteD.Text <> "" Then
        If MainlineOperationGUI!OneByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!OneByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!TwoByteD.Text <> "" Then
        If MainlineOperationGUI!TwoByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!TwoByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!ThreeByteD.Text <> "" Then
        If MainlineOperationGUI!ThreeByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!ThreeByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!FourByteD.Text <> "" Then
        If MainlineOperationGUI!FourByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!FourByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!FiveByteD.Text <> "" Then
        If MainlineOperationGUI!FiveByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!FiveByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!SixByteD.Text <> "" Then
        If MainlineOperationGUI!SixByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!SixByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    
' Which Bit?
'
' My Notes:
'
' Since our fornext loop starts at a value of one and continues throu to value of eight, the value of the bit we are
' checking on is placed into a temporary spot. When calculating the error byte we need tuen on the appropriate bit.
        
    If X = 1 Then bitvalue = 128
    If X = 2 Then bitvalue = 64
    If X = 3 Then bitvalue = 32
    If X = 4 Then bitvalue = 16
    If X = 5 Then bitvalue = 8
    If X = 6 Then bitvalue = 4
    If X = 7 Then bitvalue = 2
    If X = 8 Then bitvalue = 1
    
' My Notes:
'
' The last step of the loop is to find out if the total number of ones, is even or odd. This is used in calculating the
' error byte. On the first loop, x =1, and the bitvalue = 128 (most significant bit) and there for if the number of ones,
' is odd then the error bit will be one. This is the 'exclusive or' operation or 'xor'.

    If Int(temp / 2) <> (temp / 2) Then MainlineOperationGUI!SevenByteD.Text = Val(MainlineOperationGUI!SevenByteD.Text) + bitvalue

' My Notes:
'
' This is where we need to return to the top of the 'for next' loop. Again the loop is preformed eight times, once for
' each bit inthe varible.
            
Next X
  

    
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.
'
' Before setting the communication port, I used this let statement to set the visual status on the screen. Nost of the
' Screen contain this lable to help notify the user waht is happening with the program.
'
' Let Statements
'
' Two Visual Basic statements are used in combination with the assignment operator (=).
' The Let statement, although usually implicit, is used for assigning values.
' The Set statement, which must always be explicit, is used for assigning object references.
' If you use Let instead of Set when assigning an object reference, you will generally end up assigning the value of the object's default property.
' Attempting to use the resulting variable as an object reference will usually result in an error, such as  Error 424 Object required.

Let MainlineOperationGUI!LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

' As well, I initially set the 'commandcontrol' string to the North Coast Engineering command for sending a command to the
' decoder. The following format is used in sending a packet:
'       's cxx yy yy..'
'   where 's' repersent the command to send a packet
'   where 'c' represent the nottation of number of times to repeat this packet.
'   where 'xx' is the number of times to send this packet in hexidecimal. I've hardcoded this to four.
'   where 'yy' is the data to be sent to the command station, and repeated as often as necessary.
' The last hexidecimal should be the error byte.

        Let CommandControl = "q"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
        If MainlineOperationGUI!OneByteD.Text <> "" Then
            If MainlineOperationGUI!OneByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!OneByteH.Text
            End If
        End If
        If MainlineOperationGUI!TwoByteD.Text <> "" Then
            If MainlineOperationGUI!TwoByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!TwoByteH.Text
            End If
        End If
        If MainlineOperationGUI!ThreeByteD.Text <> "" Then
            If MainlineOperationGUI!ThreeByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!ThreeByteH.Text
            End If
        End If
        If MainlineOperationGUI!FourByteD.Text <> "" Then
            If MainlineOperationGUI!FourByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!FourByteH.Text
            End If
        End If
        If MainlineOperationGUI!FiveByteD.Text <> "" Then
            If MainlineOperationGUI!FiveByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!FiveByteH.Text
            End If
        End If
        If MainlineOperationGUI!SixByteD.Text <> "" Then
            If MainlineOperationGUI!SixByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!SixByteH.Text
            End If
        End If

        If MainlineOperationGUI!SevenByteD.Text <> "" Then
            If MainlineOperationGUI!SevenByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!SevenByteH.Text
            End If
        End If
            
' We finish the string by adding a carriage return to it. The command station will then recognize the command when sent.
            
            Let CommandControl = CommandControl + Chr$(13)
            
' Start Sending the information
'
' The first order of business before sending the command to the communication port is to add the command string to the
' communication window. This communication window is located the the Automatic Train Control Form, and controls which
' controls all the characters going in and out of the communication port.
' The following statement, i believe, set the cursor to the end of the new text being ddisplayed in the communication window.

    Let MainScreen.CommunicationWindow.Text = MainScreen.CommunicationWindow.Text + CommandControl + Chr$(10)
    Let MainScreen.CommunicationWindow.SelStart = Len(MainScreen.CommunicationWindow.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

    MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent paramters as per National Model Railroader Association specification.
    
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

    While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
        Let temp = DoEvents
    Wend
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...
    
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.
'
' Before setting the communication port, I used this let statement to set the visual status on the screen. Nost of the
' Screen contain this lable to help notify the user waht is happening with the program.
'
' Let Statements
'
' Two Visual Basic statements are used in combination with the assignment operator (=).
' The Let statement, although usually implicit, is used for assigning values.
' The Set statement, which must always be explicit, is used for assigning object references.
' If you use Let instead of Set when assigning an object reference, you will generally end up assigning the value of the object's default property.
' Attempting to use the resulting variable as an object reference will usually result in an error, such as  Error 424 Object required.

Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

' As well, I initially set the 'commandcontrol' string to the North Coast Engineering command for sending a command to the
' decoder. The following format is used in sending a packet:
'       's cxx yy yy..'
'   where 's' repersent the command to send a packet
'   where 'c' represent the nottation of number of times to repeat this packet.
'   where 'xx' is the number of times to send this packet in hexidecimal. I've hardcoded this to four.
'   where 'yy' is the data to be sent to the command station, and repeated as often as necessary.
' The last hexidecimal should be the error byte.

If MainScreen!checkboxdequeuepacket.Value = vbChecked Then

        Let CommandControl = "d"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
        If MainlineOperationGUI!OneByteD.Text <> "" Then
            If MainlineOperationGUI!OneByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!OneByteH.Text
            End If
        End If
        If MainlineOperationGUI!TwoByteD.Text <> "" Then
            If MainlineOperationGUI!TwoByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!TwoByteH.Text
            End If
        End If
        'If mainlineoperationGUI!ThreeByteD.Text <> "" Then
        '    If mainlineoperationGUI!ThreeByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!ThreeByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!FourByteD.Text <> "" Then
        '    If mainlineoperationGUI!FourByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!FourByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!FiveByteD.Text <> "" Then
        '    If mainlineoperationGUI!FiveByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!FiveByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!SixByteD.Text <> "" Then
        '    If mainlineoperationGUI!SixByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!SixByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!SevenByteD.Text <> "" Then
        '    If mainlineoperationGUI!SevenByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!SevenByteH.Text
        '    End If
        'End If
            
' We finish the string by adding a carriage return to it. The command station will then recognize the command when sent.
            
            Let CommandControl = CommandControl + Chr$(13)
            
' Start Sending the information
'
' The first order of business before sending the command to the communication port is to add the command string to the
' communication window. This communication window is located the the Automatic Train Control Form, and controls which
' controls all the characters going in and out of the communication port.
' The following statement, i believe, set the cursor to the end of the new text being ddisplayed in the communication window.

    Let MainScreen.CommunicationWindow.Text = MainScreen.CommunicationWindow.Text + CommandControl + Chr$(10)
    Let MainScreen.CommunicationWindow.SelStart = Len(MainScreen.CommunicationWindow.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

    MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent paramters as per National Model Railroader Association specification.
    
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

    While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
        Let temp = DoEvents
    Wend
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

End If

' =========================================================================================================================
' Automatic Addtion to Comments line
'


End Sub

Public Sub SetLocomotiveNumber()

If MainlineOperationGUI!ShortAdDress.Value = unvbChecked Then
    Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text) / 256)
    Let MainlineOperationGUI!TwoByteD.Text = Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text) - (Val(MainlineOperationGUI!OneByteD.Text) * 256)
    Let MainlineOperationGUI!OneByteD.Text = Val(MainlineOperationGUI!OneByteD.Text) + 128 + 64
    Let MainlineOperationGUI!ConsistControlComment.Text = "Loco " + MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text + "; "
End If

If MainlineOperationGUI!ShortAdDress.Value = vbChecked Then
    Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text))
    Let MainlineOperationGUI!TwoByteD.Text = ""
    Let MainlineOperationGUI!ConsistControlComment.Text = "Consist " + MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text + "; "
End If

End Sub

Private Sub SetFunction01234()

Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Function "

Let temporarybyte = 128

If MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 16
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "0 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "0 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 1
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "1 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "1 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 2
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "2 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "2 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 4
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "3 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "3 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 8
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "4 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "4 Off;"
End If

Let MainlineOperationGUI!ThreeByteD.Text = temporarybyte
Let MainlineOperationGUI!FourByteD.Text = ""
Let MainlineOperationGUI!FiveByteD.Text = ""
Let MainlineOperationGUI!SixByteD.Text = ""

End Sub

Private Sub SetFunction5678()

Let temporarybyte = 128 + 32

If MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 1
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "5 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "5 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 2
    Let MainlineOperationGUI.ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "6 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "6 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 4
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "7 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "7 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 8
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "8 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "8 Off;"
End If

Let MainlineOperationGUI!ThreeByteD.Text = temporarybyte
Let MainlineOperationGUI!FourByteD.Text = ""
Let MainlineOperationGUI!FiveByteD.Text = ""
Let MainlineOperationGUI!SixByteD.Text = ""

End Sub

Private Sub SetChangeCV()
    
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
        
        Let TemporaryByteOne = TemporaryByteOne + 128
        Let TemporaryByteOne = TemporaryByteOne + 64
        Let TemporaryByteOne = TemporaryByteOne + 32
        
       If MainlineOperationGUI!ConsistControlCVRead = vbChecked Then
              Let TemporaryByteOne = TemporaryByteOne + 4
        Else
            Let TemporaryByteOne = TemporaryByteOne + 8 + 4
        End If
        
    Let MainlineOperationGUI!ThreeByteD.Text = TemporaryByteOne
    Let MainlineOperationGUI!FourByteD.Text = TemporaryByteTwo
    Let MainlineOperationGUI!FiveByteD.Text = Val(ConsistControlCVValue.Text)
    Let MainlineOperationGUI!SixByteD.Text = ""
    
 Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + _
     "Change CV" + MainlineOperationGUI!ConsistControlCV.Text + " to " + MainlineOperationGUI!ConsistControlCVValue.Text

  
End Sub

Private Sub SetSpeed()

Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Speed "

If MainlineOperationGUI!ConsistControlSpeed128.Value = vbChecked Then
    ' This routine assembles the byte for speed step mode 128
    Let temporary = Val(MainlineOperationGUI!ConsistControlSpeed.Value)
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + Str$(temporary) + " of 128 "
    
    If MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked Then
        temporary = temporary + 128 ' add forward direction
        Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Forward"
    Else
        Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Reverse"
    End If
    
    Let MainlineOperationGUI!ThreeByteD.Text = 63
    Let MainlineOperationGUI!FourByteD.Text = temporary
    Let MainlineOperationGUI!FiveByteD.Text = ""
    Let MainlineOperationGUI!SixByteD.Text = ""
Else
    Let temporary = 64
    If MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked Then
            Let temporary = temporary + 32 ' add forward direction
    End If
    
   If MainlineOperationGUI!ConsistControlSpeed28.Value = vbChecked Then
        'This routine assenmles the byte for speed step mode 28
        Let temp1 = Val(MainlineOperationGUI!ConsistControlSpeed.Value) ' adds the speed
        Let temp2 = temp1 Mod 2
        Let newspeedvalue = Int(temp1 / 2)
        Let temporary = temporary + newspeedvalue
        If temp2 = 1 Then Let temporary = temporary + 16
        Let MainlineOperationGUI!ThreeByteD.Text = temporary
        Let MainlineOperationGUI!FourByteD.Text = ""
        Let MainlineOperationGUI!FiveByteD.Text = ""
        Let MainlineOperationGUI!SixByteD.Text = ""
    Else
        ' This routing assembles the byte for speed step mode 14
        
        Let temporary = temporary + Val(MainlineOperationGUI!ConsistControlSpeed.Value) ' add the speed
        Let MainlineOperationGUI!ThreeByteD.Text = temporary
        Let MainlineOperationGUI!FourByteD.Text = ""
        Let MainlineOperationGUI!FiveByteD.Text = ""
        Let MainlineOperationGUI!SixByteD.Text = ""
    
    End If
End If

End Sub



Public Sub SetSoundDecoderNumber()

If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then

    If MainlineOperationGUI!CheckBoxSoundDecoderShortAddress.Value = unvbChecked Then
        Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!TextBoxMappedDecoderNumber.Text) / 256)
        Let MainlineOperationGUI!TwoByteD.Text = Val(MainlineOperationGUI!TextBoxMappedDecoderNumber.Text) - (Val(MainlineOperationGUI!OneByteD.Text) * 256)
        Let MainlineOperationGUI!OneByteD.Text = Val(MainlineOperationGUI!OneByteD.Text) + 128 + 64
        Let MainlineOperationGUI!ConsistControlComment.Text = "Loco " + MainlineOperationGUI!TextBoxMappedDecoderNumber.Text + "; "
    Else
        Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!TextBoxMappedDecoderNumber.Text))
        Let MainlineOperationGUI!TwoByteD.Text = ""
        Let MainlineOperationGUI!ConsistControlComment.Text = "Consist " + MainlineOperationGUI!TextBoxMappedDecoderNumber.Text + "; "
    End If
End If

End Sub




Private Sub VideoCapture_ErrorMessage(ByVal ErrCode As Long, ByVal ErrString As String)

    Let PreviousNotes = Left$(VideoCaptureNotes.Caption, 200)

    Let VideoCaptureNotes.Caption = "Error Code: " + Str$(ErrCode) + " at " + Str$(Time) + Chr$(13)
    Let VideoCaptureNotes.Caption = VideoCaptureNotes.Caption + "Error Message: " + ErrString + Chr$(13)
    Let VideoCaptureNotes.Caption = VideoCaptureNotes.Caption + Chr$(13) + PreviousNotes

    
End Sub

Private Sub VideoCapture_GotFocus()

Stop
MainlineOperationGUIScreen.PopupMenu menuCaptureDevice

End Sub


Private Sub VideoCapture_LostFocus()
Stop
End Sub


Private Sub VideoCapture_StatusMessage(ByVal StatCode As Long, ByVal StatString As String)

    Let PreviousNotes = Left$(VideoCaptureNotes.Caption, 200)

    Let VideoCaptureNotes.Caption = "Status Code: " + Str$(StatCode) + " at " + Str$(Time) + Chr$(13)
    Let VideoCaptureNotes.Caption = VideoCaptureNotes.Caption + "Status Message: " + StatString + Chr$(13)
    Let VideoCaptureNotes.Caption = VideoCaptureNotes.Caption + Chr$(13) + PreviousNotes

 
End Sub


