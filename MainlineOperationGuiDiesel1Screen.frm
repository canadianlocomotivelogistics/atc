VERSION 4.00
Begin VB.Form MainlineOperationGuiDiesel1Screen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - GUI - Diesel 9-40CW"
   ClientHeight    =   11790
   ClientLeft      =   1125
   ClientTop       =   1425
   ClientWidth     =   15330
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Height          =   12195
   Icon            =   "MainlineOperationGuiDiesel1Screen.frx":0000
   Left            =   1065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12103.15
   ScaleMode       =   0  'User
   ScaleWidth      =   17000
   Tag             =   "0"
   Top             =   1080
   Width           =   15450
   Begin VB.Timer TimerTrackJoint 
      Left            =   15720
      Top             =   1800
   End
   Begin VB.Timer TimerCheckWinsock 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   15720
      Top             =   5100
   End
   Begin VB.Timer TimerAdvertising 
      Interval        =   65535
      Left            =   15720
      Top             =   3960
   End
   Begin VB.Timer TimerStartEngine3 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   15720
      Top             =   1320
   End
   Begin VB.Timer TimerStartEngine2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   15720
      Top             =   840
   End
   Begin VB.Timer TimerStartEngine1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   15720
      Top             =   360
   End
   Begin VB.PictureBox PictureBoxLocomotiveCab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   768
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "d1"
      Top             =   0
      Width           =   15360
      Begin VB.CommandButton ButtonEngineStart 
         Caption         =   "&Start Engine"
         Height          =   255
         Left            =   14040
         TabIndex        =   190
         Top             =   7320
         Width           =   1230
      End
      Begin VB.TextBox TextBoxStatusThrottleDelay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   185
         Text            =   "0"
         Top             =   1560
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxSettingSpeedChangeInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10320
         TabIndex        =   183
         Text            =   "1000"
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxSettingSendCommandInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10320
         TabIndex        =   181
         Text            =   "1000"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextboxSettingAutomaticAirBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10320
         TabIndex        =   179
         Text            =   "1750"
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxSettingIndependentAirBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10320
         TabIndex        =   177
         Text            =   "1750"
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxSettingDeadmannInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10320
         TabIndex        =   175
         Text            =   "50000"
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxStatusDynamicBrakeApplication 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   173
         Text            =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.HScrollBar ScrollTimeAdjustment 
         Height          =   150
         LargeChange     =   10
         Left            =   14040
         Max             =   100
         Min             =   -50
         TabIndex        =   171
         Top             =   7620
         Width           =   1230
      End
      Begin VB.TextBox TextBoxVirtualCabSignalStatusCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   160
         Text            =   "Clear"
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxVirtualCabSignalStatusApproaching 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   158
         Text            =   "Clear"
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxVirtualCabSignalDistance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   155
         Text            =   "5380"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TextBoxStatusSpeedStepNow 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Text            =   "0"
         Top             =   1080
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
         TabIndex        =   10
         Text            =   "0"
         Top             =   840
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
         TabIndex        =   12
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
         TabIndex        =   11
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxStatusTractionEffortTooHigh 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   10320
         TabIndex        =   131
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
         TabIndex        =   130
         Text            =   "1200"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton ButtonHelp 
         Caption         =   "&Help is Off"
         Height          =   255
         Left            =   14040
         TabIndex        =   129
         Top             =   7800
         Width           =   1230
      End
      Begin VB.TextBox TextBoxStatusFuelTooLow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
         Text            =   "60"
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   8
         Left            =   3120
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   101
         Top             =   5760
         Visible         =   0   'False
         Width           =   3990
         Begin VB.TextBox VideoCaptureNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000C000&
            Height          =   2295
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   210
            TabStop         =   0   'False
            Top             =   285
            Width           =   3855
         End
         Begin VB.Label VideoCaptureTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Video Capture Notes"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   1200
            TabIndex        =   154
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 8"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   16
            Left            =   3255
            TabIndex        =   102
            Top             =   2655
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   7
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   99
         Top             =   5760
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
            TabIndex        =   100
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   6
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   97
         Top             =   5760
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
            TabIndex        =   98
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   5
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   95
         Top             =   5760
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
            TabIndex        =   96
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   4
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   93
         Top             =   5760
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
            TabIndex        =   94
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   3
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   91
         Top             =   5760
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
            TabIndex        =   92
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   2
         Left            =   3105
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   89
         Top             =   5760
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
            TabIndex        =   90
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   1
         Left            =   3120
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   87
         Top             =   5760
         Width           =   3990
         Begin VB.PictureBox PictureBoxVirtualCabSignal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   1815
            Left            =   2760
            ScaleHeight     =   1815
            ScaleWidth      =   1215
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
            Begin VB.PictureBox PictureBoxVirtualCabSignalCurrent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   2
               Left            =   720
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   1200
               Width           =   299
            End
            Begin VB.PictureBox PictureBoxVirtualCabSignalApproaching 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   2
               Left            =   240
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   167
               TabStop         =   0   'False
               Top             =   1200
               Width           =   299
            End
            Begin VB.PictureBox PictureBoxVirtualCabSignalCurrent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   1
               Left            =   720
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   166
               TabStop         =   0   'False
               Top             =   840
               Width           =   299
            End
            Begin VB.PictureBox PictureBoxVirtualCabSignalApproaching 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   1
               Left            =   240
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   165
               TabStop         =   0   'False
               Top             =   840
               Width           =   299
            End
            Begin VB.PictureBox PictureBoxVirtualCabSignalCurrent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   0
               Left            =   720
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   480
               Width           =   299
            End
            Begin VB.PictureBox PictureBoxVirtualCabSignalApproaching 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   293
               Index           =   0
               Left            =   240
               ScaleHeight     =   20
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   20
               TabIndex        =   163
               TabStop         =   0   'False
               Top             =   480
               Width           =   299
            End
            Begin VB.Label LabelVirtualCabSignalDistanceFeet 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "0"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   240
               TabIndex        =   170
               Top             =   1560
               Width           =   645
            End
            Begin VB.Label LabelVirtualCabSignalDistance 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "ft"
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   960
               TabIndex        =   169
               Top             =   1560
               Width           =   90
            End
            Begin VB.Label LabelVirtualCabSignalTitle 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "Cab Signals (A)      (C) "
               ForeColor       =   &H0000C000&
               Height          =   375
               Index           =   17
               Left            =   0
               TabIndex        =   162
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.PictureBox PictureBoxSpeedometer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2460
            Left            =   0
            ScaleHeight     =   2460
            ScaleWidth      =   2220
            TabIndex        =   103
            Top             =   0
            Width           =   2220
            Begin VB.TextBox TextBoxDigitalSpeed 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H0000C000&
               Height          =   285
               Left            =   780
               TabIndex        =   104
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               TabIndex        =   109
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
            TabIndex        =   116
            Top             =   2520
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label LabelCombinedPower 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "             Combined Power  Neutral - Notch Zero"
            ForeColor       =   &H0000C000&
            Height          =   690
            Left            =   2160
            TabIndex        =   115
            Top             =   0
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
            TabIndex        =   88
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
         Top             =   8760
         Width           =   375
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   8760
         Width           =   375
      End
      Begin VB.PictureBox PictureBoxResetLeft 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   1500
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   22
         Tag             =   "0"
         Top             =   10200
         Width           =   795
      End
      Begin VB.PictureBox PictureBoxBell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2520
         ScaleHeight     =   540
         ScaleWidth      =   765
         TabIndex        =   21
         Tag             =   "0"
         Top             =   10191
         Width           =   765
      End
      Begin VB.PictureBox PictureBoxThrottle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2038
         Left            =   5520
         ScaleHeight     =   2040
         ScaleWidth      =   660
         TabIndex        =   20
         Tag             =   "0"
         Top             =   9030
         Width           =   664
      End
      Begin VB.PictureBox PictureBoxReverser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1830
         Left            =   4185
         ScaleHeight     =   1830
         ScaleWidth      =   765
         TabIndex        =   19
         Tag             =   "1"
         Top             =   8985
         Width           =   765
      End
      Begin VB.CommandButton ButtonCloseGUI 
         Caption         =   "&Close"
         Height          =   255
         Left            =   14040
         TabIndex        =   18
         Top             =   8100
         Width           =   1230
      End
      Begin VB.PictureBox PictureBoxLight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   13920
         ScaleHeight     =   1485
         ScaleWidth      =   1440
         TabIndex        =   17
         Tag             =   "0"
         Top             =   5715
         Width           =   1440
      End
      Begin VB.PictureBox PictureBoxHorn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2085
         ScaleHeight     =   480
         ScaleWidth      =   795
         TabIndex        =   15
         Tag             =   "0"
         Top             =   10725
         Width           =   795
      End
      Begin VB.TextBox TextBoxStatusSpeedStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TextBoxRadiatorFans1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   11040
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Text            =   "Off"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
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
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   440
         Left            =   2865
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   16
         Tag             =   "0"
         Top             =   9720
         Width           =   675
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   8
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   71
         Top             =   5760
         Visible         =   0   'False
         Width           =   3990
         Begin VB.PictureBox PictureAdvertising 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   0
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   263
            TabIndex        =   207
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Screen 8"
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   72
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   7
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   70
         Top             =   5760
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
            TabIndex        =   73
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   6
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   69
         Top             =   5760
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
            TabIndex        =   74
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   5
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   68
         Top             =   5760
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
            TabIndex        =   75
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   4
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   67
         Top             =   5760
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
            TabIndex        =   76
            Top             =   2660
            Width           =   645
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   3
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   51
         Top             =   5760
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
            TabIndex        =   77
            Top             =   2660
            Width           =   645
         End
         Begin GBAR1.GBar BarTractionEffort 
            Height          =   165
            Left            =   1080
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   960
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   130000
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   786432
            FillColor       =   786432
            FontColor       =   16777215
            Units           =   " of 130,000 lbs"
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
            TabIndex        =   57
            Top             =   960
            Width           =   1005
         End
         Begin GBAR1.GBar BarAmpere 
            Height          =   165
            Left            =   1080
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   720
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   1300
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   786432
            FillColor       =   786432
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
            TabIndex        =   55
            Top             =   720
            Width           =   540
         End
         Begin GBAR1.GBar BarRPM 
            Height          =   165
            Left            =   1080
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   480
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   2000
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49152
            FillColor       =   786432
            FontColor       =   16777215
            Units           =   " of 2000"
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
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   2
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   49
         Top             =   5760
         Visible         =   0   'False
         Width           =   3990
         Begin VB.CommandButton ButtonFillOil 
            Height          =   195
            Left            =   240
            TabIndex        =   120
            Top             =   1200
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillWater 
            Height          =   195
            Left            =   240
            TabIndex        =   119
            Top             =   960
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillSand 
            Height          =   195
            Left            =   240
            TabIndex        =   118
            Top             =   720
            Width           =   135
         End
         Begin VB.CommandButton ButtonFillFuel 
            Height          =   195
            Left            =   240
            TabIndex        =   117
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
            TabIndex        =   78
            Top             =   2660
            Width           =   645
         End
         Begin GBAR1.GBar BarOilTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   80
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49152
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 80 gallons"
         End
         Begin VB.Label LabelOilTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Oil"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   600
            TabIndex        =   65
            Top             =   1200
            Width           =   420
         End
         Begin GBAR1.GBar BarWaterTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   960
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   750
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49152
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 750 gallons"
         End
         Begin VB.Label LabelWaterTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Water"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   480
            TabIndex        =   63
            Top             =   960
            Width           =   555
         End
         Begin GBAR1.GBar BarSandTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   720
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   200
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49152
            FillColor       =   49152
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
            TabIndex        =   61
            Top             =   720
            Width           =   495
         End
         Begin GBAR1.GBar BarFuelTank 
            Height          =   165
            Left            =   1080
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   480
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            BackColor       =   0
            Min             =   1
            Max             =   4500
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   786432
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 4500 gallons"
         End
         Begin VB.Label LabelFuelTank 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Fuel"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   600
            TabIndex        =   59
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
            TabIndex        =   50
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox PictureBoxComputerScreenRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2955
         Index           =   1
         Left            =   9000
         ScaleHeight     =   2925
         ScaleWidth      =   3960
         TabIndex        =   23
         Top             =   5760
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
            TabIndex        =   79
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
            TabIndex        =   40
            Top             =   2355
            Width           =   900
         End
         Begin GBAR1.GBar BarWaterTemperature 
            Height          =   165
            Left            =   1080
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2370
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   2
            Max             =   230
            Value           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49344
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
            TabIndex        =   38
            Top             =   2130
            Width           =   615
         End
         Begin GBAR1.GBar BarWaterPressure 
            Height          =   165
            Left            =   1080
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2130
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
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
            TabIndex        =   36
            Top             =   1875
            Width           =   450
         End
         Begin GBAR1.GBar BarOilTemperature 
            Height          =   165
            Left            =   1080
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1620
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   2
            Max             =   230
            Value           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 230 degrees"
         End
         Begin GBAR1.GBar BarOilPressure 
            Height          =   165
            Left            =   1080
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1395
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   1
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   1050
            Width           =   360
         End
         Begin GBAR1.GBar BarAirIndependentBrake 
            Height          =   165
            Left            =   1080
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   765
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   1
            Max             =   120
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 120 lbs"
         End
         Begin VB.Label LabelIndependent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Independent"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   0
            TabIndex        =   29
            Top             =   765
            Width           =   1020
         End
         Begin GBAR1.GBar BarAirAutomaticBrake 
            Height          =   165
            Left            =   1080
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   540
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   1
            Max             =   120
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 120 lbs"
         End
         Begin VB.Label LabelAutomatic 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Automatic"
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   240
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   15
            Width           =   840
         End
         Begin GBAR1.GBar BarAirMainReservoir 
            Height          =   165
            Left            =   1080
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   300
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   291
            ForeColor       =   49344
            BackColor       =   0
            Min             =   1
            Max             =   120
            Value           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   49344
            FillColor       =   49152
            FontColor       =   16777215
            Units           =   " of 120 lbs"
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
            TabIndex        =   24
            Top             =   300
            Width           =   1065
         End
      End
      Begin TransPicture.TransPictureCtl PictureBoxAutomaticBrake 
         Height          =   1965
         Left            =   7935
         TabIndex        =   266
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9000
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   3466
         TransparentColor=   16777215
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":0442
      End
      Begin TransPicture.TransPictureCtl TransPictureSignal 
         Height          =   2685
         Left            =   14985
         TabIndex        =   265
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1275
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   4736
         TransparentColor=   16777215
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":045E
      End
      Begin TransPicture.TransPictureCtl TransPictureBoxIndependentBrake 
         Height          =   1935
         Left            =   12780
         TabIndex        =   263
         TabStop         =   0   'False
         Tag             =   "10"
         Top             =   8985
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   3413
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureBoxResetRight 
         Height          =   795
         Left            =   14280
         TabIndex        =   262
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1402
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   14
         Left            =   15000
         TabIndex        =   261
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8558
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   13
         Left            =   14760
         TabIndex        =   260
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8574
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   12
         Left            =   14520
         TabIndex        =   259
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8590
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   11
         Left            =   14280
         TabIndex        =   258
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":85AC
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   10
         Left            =   14040
         TabIndex        =   257
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":85C8
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   9
         Left            =   15000
         TabIndex        =   256
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8760
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":85E4
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   8
         Left            =   14760
         TabIndex        =   255
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8760
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8600
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   7
         Left            =   14520
         TabIndex        =   254
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8760
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":861C
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   6
         Left            =   14280
         TabIndex        =   253
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8760
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8638
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   5
         Left            =   14040
         TabIndex        =   252
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8760
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8654
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   4
         Left            =   15000
         TabIndex        =   251
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8400
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":8670
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   3
         Left            =   14760
         TabIndex        =   250
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8400
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":868C
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   2
         Left            =   14520
         TabIndex        =   249
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8400
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":86A8
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   1
         Left            =   14280
         TabIndex        =   248
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8400
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":86C4
      End
      Begin TransPicture.TransPictureCtl TransPictureAuxillarySwitch 
         Height          =   345
         Index           =   0
         Left            =   14040
         TabIndex        =   223
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8400
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Picture         =   "MainlineOperationGuiDiesel1Screen.frx":86E0
      End
      Begin TransPicture.TransPictureCtl TransPictureButtonSide 
         Height          =   525
         Index           =   5
         Left            =   120
         TabIndex        =   221
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8700
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   926
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureButtonSide 
         Height          =   555
         Index           =   4
         Left            =   120
         TabIndex        =   220
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8040
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   979
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureButtonSide 
         Height          =   555
         Index           =   3
         Left            =   120
         TabIndex        =   219
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   7380
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   979
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureButtonSide 
         Height          =   525
         Index           =   1
         Left            =   120
         TabIndex        =   218
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   6060
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   926
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureButtonSide 
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   217
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   6720
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   979
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureSunVisor 
         Height          =   3015
         Left            =   0
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureWipers 
         Height          =   3780
         Left            =   3750
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   0
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   6668
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl PictureBoxRadioPhone 
         Height          =   6270
         Left            =   0
         TabIndex        =   214
         Tag             =   "1"
         Top             =   5280
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   11060
         TransparentColor=   16777215
      End
      Begin TransPicture.TransPictureCtl TransPictureBoxCabWindow 
         Height          =   5085
         Left            =   0
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   0
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   8969
         TransparentColor=   16777215
         Stretch         =   -1  'True
      End
      Begin VIDEOCAPXLib.VideoCapX VideoCapture 
         Height          =   5085
         Left            =   0
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   0
         Width           =   15345
         _Version        =   131072
         _ExtentX        =   27067
         _ExtentY        =   8969
         _StockProps     =   1
         CapFilename     =   ""
         VideoDeviceIndex=   -1
         AudioDeviceIndex=   -1
         VideoCodecQuality=   100
         ServerPort      =   10102
         ServerQuality   =   100
         WMTVersion      =   8
         WMAttributes    =   $"MainlineOperationGuiDiesel1Screen.frx":86FC
         StretchMode     =   1
      End
      Begin VB.Label LabelStatusTimeDelay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSatstusThrottleDelay"
         Height          =   195
         Left            =   1080
         TabIndex        =   186
         Top             =   1560
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label LabelLocomotiveNumberDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   13920
         TabIndex        =   184
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSettingSpeedChangeInterval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7560
         TabIndex        =   182
         Top             =   2400
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSettingSendCommandInterval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7500
         TabIndex        =   180
         Top             =   2160
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSettingAutomaticAirBrake"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7800
         TabIndex        =   178
         Top             =   1920
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSettingIndependentAirBrake"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7605
         TabIndex        =   176
         Top             =   1680
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSettingDeadmannInterval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7755
         TabIndex        =   174
         Top             =   1440
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusDynamicBrakeApplication"
         Height          =   195
         Left            =   240
         TabIndex        =   172
         Top             =   600
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxVirtualCabSignalStatusCurrent"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7440
         TabIndex        =   159
         Top             =   1200
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxVirtualCabSignalStatusApproaching"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7080
         TabIndex        =   157
         Top             =   960
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxVirtualCabSignalDistance"
         Height          =   195
         Left            =   7800
         TabIndex        =   156
         Top             =   720
         Visible         =   0   'False
         Width           =   2370
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSpeedStepNow"
         Height          =   195
         Left            =   360
         TabIndex        =   133
         Top             =   1080
         Visible         =   0   'False
         Width           =   2850
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
         TabIndex        =   134
         Top             =   840
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusAutomaticBrakePressure"
         Height          =   195
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Label LabelRadiatorFans1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxRadiator Fans 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   153
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelRadiatorFans2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxRadiator Fans 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   152
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelRadiatorFans3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxRadiator Fans 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   151
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelDynamicBrakeFan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxDynamic Brake Fan"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   150
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxFuelTank"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   149
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxOilTank"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   148
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxWaterTank"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   147
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label48 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxSandTank"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   146
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelOilTooLowTemperature 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusOilTooLowTemerature"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   145
         Top             =   2160
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label LabelOilTooLowPressure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusOilTooLowPressure"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   144
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label LabelWaterTooLowTemperature 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusWaterTooLowTemperature"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   143
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LabelWaterTooLowPressure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusWaterTooLowPressure"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   142
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label LabelSandTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSandTooLow"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   141
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelWaterTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusWaterTooLow"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   140
         Top             =   3360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelOilTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusOilTooLow"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   139
         Top             =   3600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelFuelTooLow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusFuelTooLow"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   138
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LabelStatusAmpereTooHigh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusAmpereTooHigh"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7560
         TabIndex        =   137
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LabelStatusTractionEffortTooHigh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusTractionEffortTooHigh"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7440
         TabIndex        =   136
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxStatusSpeedStep"
         Height          =   195
         Left            =   1320
         TabIndex        =   132
         Top             =   1320
         Visible         =   0   'False
         Width           =   1890
      End
   End
   Begin FATHMAILOCXLib.SMTP SMTP1 
      Left            =   9240
      Top             =   5640
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin MCI.MMControl SoundAuxilliarySwitch 
      Height          =   315
      Left            =   21180
      TabIndex        =   264
      TabStop         =   0   'False
      Top             =   5940
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   23
      Left            =   21180
      TabIndex        =   247
      TabStop         =   0   'False
      Top             =   12840
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   22
      Left            =   21180
      TabIndex        =   246
      TabStop         =   0   'False
      Top             =   12540
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   21
      Left            =   21180
      TabIndex        =   245
      TabStop         =   0   'False
      Top             =   12240
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   20
      Left            =   21180
      TabIndex        =   244
      TabStop         =   0   'False
      Top             =   11940
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   19
      Left            =   21180
      TabIndex        =   243
      TabStop         =   0   'False
      Top             =   11640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   18
      Left            =   21180
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   11340
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   17
      Left            =   21180
      TabIndex        =   241
      TabStop         =   0   'False
      Top             =   11040
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   16
      Left            =   21180
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   10740
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   14
      Left            =   21180
      TabIndex        =   239
      TabStop         =   0   'False
      Top             =   10440
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   13
      Left            =   21180
      TabIndex        =   238
      TabStop         =   0   'False
      Top             =   10140
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   12
      Left            =   21180
      TabIndex        =   237
      TabStop         =   0   'False
      Top             =   9840
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   11
      Left            =   21180
      TabIndex        =   236
      TabStop         =   0   'False
      Top             =   9540
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   10
      Left            =   21180
      TabIndex        =   235
      TabStop         =   0   'False
      Top             =   9240
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   9
      Left            =   21180
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   8940
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   8
      Left            =   21180
      TabIndex        =   233
      TabStop         =   0   'False
      Top             =   8640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   7
      Left            =   21180
      TabIndex        =   232
      TabStop         =   0   'False
      Top             =   8340
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   6
      Left            =   21180
      TabIndex        =   231
      TabStop         =   0   'False
      Top             =   8040
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   5
      Left            =   21180
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   7740
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   4
      Left            =   21180
      TabIndex        =   229
      TabStop         =   0   'False
      Top             =   7440
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   3
      Left            =   21180
      TabIndex        =   228
      TabStop         =   0   'False
      Top             =   7140
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   300
      Index           =   2
      Left            =   21180
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   6840
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   1
      Left            =   21180
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   6540
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   315
      Index           =   0
      Left            =   21180
      TabIndex        =   225
      TabStop         =   0   'False
      Top             =   6240
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundTrackJoint 
      Height          =   435
      Index           =   15
      Left            =   0
      TabIndex        =   224
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      UpdateInterval  =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TimerTrackJoint - Plays sound of the trak joint at different intervals."
      Height          =   375
      Left            =   16260
      TabIndex        =   222
      Top             =   1860
      Width           =   2775
   End
   Begin FATHMAILOCXLib.SMTP SmtpOcx 
      Left            =   16440
      Top             =   6720
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin FATHMAILOCXLib.Message MessageOcx 
      Left            =   16500
      Top             =   6120
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin FILETRANSXLib.FileTransX FtpControl 
      Height          =   480
      Left            =   15720
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Timeout         =   30
      ServerName      =   "atc.lovethosetrains.com"
      Username        =   "softwareads"
      Password        =   "walnuttree12"
      ProxyName       =   ""
      ProxyUserID     =   ""
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   15780
      Top             =   2400
      _ExtentX        =   767
      _ExtentY        =   556
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TimerCheckWinsock - Used to check the winsock connection, if still present."
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   16260
      TabIndex        =   212
      Top             =   5100
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TimerAdvertise - Use to poll the server and get an advertising message and display it."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   211
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label LabelAdvertise 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TimerAdvertise - Use to poll the server and get an advertising message and display it."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   16320
      TabIndex        =   208
      Top             =   3960
      Width           =   3495
   End
   Begin MCI.MMControl SoundCompressor 
      Height          =   330
      Left            =   21120
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   15720
      Top             =   3480
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   15720
      Top             =   2880
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin MCI.MMControl SoundDeadmann 
      Height          =   330
      Left            =   21120
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   360
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundLeftScreen 
      Height          =   330
      Left            =   21120
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   660
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundRightScreen 
      Height          =   330
      Left            =   21120
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   960
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundResetLeft 
      Height          =   330
      Left            =   21120
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   1260
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundRadioPhone 
      Height          =   330
      Left            =   21120
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   1860
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundSand 
      Height          =   330
      Left            =   21120
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   2160
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundBell 
      Height          =   330
      Left            =   21120
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   2460
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundHorn 
      Height          =   330
      Left            =   21120
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   2760
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundReverser 
      Height          =   330
      Left            =   21120
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   3060
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundThrottle 
      Height          =   330
      Left            =   21120
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   3360
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundResetRight 
      Height          =   330
      Left            =   21120
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   1560
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundBrakeAutomatic 
      Height          =   330
      Left            =   21120
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   3660
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundBrakeIndependent 
      Height          =   330
      Left            =   21120
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   3960
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl SoundLight 
      Height          =   330
      Left            =   21120
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   4260
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.Label LabelEngineStart3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TimerEngineStart3 - To start all other timers, as if the negine was started."
      Height          =   435
      Left            =   16320
      TabIndex        =   189
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label LabelEngineStart2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TimerEngineStart2 - To start the engine sounds."
      Height          =   495
      Left            =   16320
      TabIndex        =   188
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label LabelEngineStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TimerEngineStart1 - Routines for initalizing the screen"
      Height          =   375
      Left            =   16320
      TabIndex        =   187
      Top             =   360
      Width           =   2655
   End
   Begin VB.Menu menuVideoCaptureOptions 
      Caption         =   "Video Capture Options"
      Visible         =   0   'False
      Begin VB.Menu menuAudioCodecDialog 
         Caption         =   "Audio Codec Dialog"
      End
      Begin VB.Menu menuAudioFormatDialog 
         Caption         =   "Audio Format Dialog"
      End
      Begin VB.Menu menuAudioSourceDialog 
         Caption         =   "Audio Source Dialog"
      End
      Begin VB.Menu menuBar 
         Caption         =   "-"
      End
      Begin VB.Menu menuVideoCodecDialog 
         Caption         =   "Video Codec Dialog"
      End
      Begin VB.Menu menuVideoCrossbarDialog 
         Caption         =   "Video Crossbar Dialog"
      End
      Begin VB.Menu menuVideoFormatDialog 
         Caption         =   "Video Format Dialog"
      End
      Begin VB.Menu menuVideoSourceDialog 
         Caption         =   "Video Source Dialog"
      End
   End
End
Attribute VB_Name = "MainlineOperationGuiDiesel1Screen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub ButtonCloseGUI_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, Button Close GUI, Click" & vbCrLf
    End If ' Debug Tag

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Button Close, requested." Then
            GoSub UpdateButtonClose
            Let InternetSettings!textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Button Close, requested."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateButtonClose
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateButtonClose
    End If
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel1 Screen, Button Close GUI, Click" & vbCrLf
    End If ' Debug Tag

Exit Sub

UpdateButtonClose:
    
    Let ButtonCloseGUI.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Video Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.optionHost = True Then
        VideoCapture.StopBroadcast
        Let VideoCapture.Overlay = False
        Let VideoCapture.Preview = False
        Let VideoCapture.Connected = False
        Let VideoCapture.VideoRenderer = False
    ElseIf InternetSettings.optionclient = True Then
        If InternetSettings.NetConnect.Connected = True Then
            VideoCapture.DisplayRemote "", False
        End If
    ElseIf InternetSettings.OptionStandAlone = True Then
        If VideoCapture.GetVideoDeviceCount <> 0 Then
            Let VideoCapture.Preview = False
            Let VideoCapture.Overlay = False
            Let VideoCapture.Connected = False
        End If
    End If
    
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
            If TemporaryScreen = "Mainline Operation GUI Diesel1 Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
Return
    
End Sub

Public Sub ButtonEngineStart_Click()
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, Button Engine Start, Click" & vbCrLf
    End If ' Debug Tag

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Start Engine, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Start Engine, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateStartEngine
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "Stop Engine, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Stop Engine, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateStopEngine
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Start Engine, confirmed." Then
            GoSub UpdateStartEngine
            Let InternetSettings.textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Stop Engine, confirmed." Then
            GoSub UpdateStopEngine
            Let InternetSettings.textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "" Then
            If ButtonEngineStart.Caption = "&Start Engine" Then
                Let InternetSettings!textboxoutboundcommand.Text = "Start Engine, requested."
            Else
                Let InternetSettings!textboxoutboundcommand.Text = "Stop Engine, requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        If ButtonEngineStart.Caption = "&Start Engine" Then
            Let ButtonCloseGUI.Enabled = False
            GoSub UpdateStartEngine
        Else
            GoSub UpdateStopEngine
            Let ButtonCloseGUI.Enabled = True
        End If
    End If
        ' The program must be running with out  any internet connection

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Exit Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel1 Screen, Button Engine Start, Click" & vbCrLf
    End If ' Debug Tag

Exit Sub

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Engine Start Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateStartEngine:
    Let ButtonCloseGUI.Enabled = False
    Let ButtonEngineStart.Caption = "&Stop Engine"
    Let TimerStartEngine1.Enabled = True
    Let TimerStartEngine2.Enabled = True
    Let TimerStartEngine3.Enabled = True

Return

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Engine Stop Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateStopEngine:
    
    Let ButtonEngineStart.Caption = "&Start Engine"
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Beginning of Throttle Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainlineOperationGUI!timerdeadmann.Enabled = False
    Let MainlineOperationGUI!timerdeadmann.Interval = 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Speed to Emergency Stop
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.Winsock.State = sckConnected Then
        If InternetSettings.optionHost.Value = True Then
            Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
        End If
    ElseIf InternetSettings.Winsock.State = sckClosed Then
        If InternetSettings.OptionStandAlone.Value = True Then
            Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
        End If
    End If
                    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Show Speedometer to Zero
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep = 0
    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow = 0
    Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified = 0

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Move Throttle to Zero
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
OneMore:
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
    ElseIf Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
    End If
            
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) <> 0 Then
        Let SoundThrottle.Notify = False
        Let SoundThrottle.Wait = False
        Let SoundThrottle.Shareable = False
        Let SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        Let SoundThrottle.Command = "Open"
        Let SoundThrottle.Notify = True
        Let SoundThrottle.Command = "Play"
    End If
            
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
    
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) >= 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 7 Then
        
            If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
                
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
            ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            End If
                
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
            
            If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
            
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
                    Let MainlineOperationGUI!SoundControlFunction9.Value = vbnChecked
                ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "10" Then
                    Let MainlineOperationGUI!SoundControlFunction10.Value = vbnChecked
                ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "11" Then
                    Let MainlineOperationGUI!SoundControlFunction11.Value = vbnChecked
                ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMminus.Text = "12" Then
                    Let MainlineOperationGUI!SoundControlFunction12.Value = vbnChecked
                End If
            ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            End If
            
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

        End If
                    
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) = 0 Then
        
            If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
                    
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
            ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            End If
            
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
                
        If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < -1 Then
                
            If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
                
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
            ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            End If


            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If
                    
           
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Computer screen with throttle position
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake"
    Else
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle"
    End If
    Let temp = temp + Trim(Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag)) & "(d1).bmp"
    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
    Sleep (500)
    Let LabelCombinedPower.Caption = "Combined Power" + Chr$(13)
    If Val(PictureBoxReverser.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Forward - "
    ElseIf Val(PictureBoxReverser.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Reverse - "
    End If
    If Val(PictureBoxThrottle.Tag) < 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Dynamic Brake" + Chr(13)
    End If
    If Abs(PictureBoxThrottle.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Zero"
    ElseIf Abs(PictureBoxThrottle.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch One"
    ElseIf Abs(PictureBoxThrottle.Tag) = 2 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Two"
    ElseIf Abs(PictureBoxThrottle.Tag) = 3 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Three"
    ElseIf Abs(PictureBoxThrottle.Tag) = 4 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Four"
    ElseIf Abs(PictureBoxThrottle.Tag) = 5 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Five"
    ElseIf Abs(PictureBoxThrottle.Tag) = 6 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Six"
    ElseIf Abs(PictureBoxThrottle.Tag) = 7 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Seven"
    ElseIf Abs(PictureBoxThrottle.Tag) = 8 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Eight"
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Loop if Throttle not in Neutral
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) <> 0 Then GoTo OneMore
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Off all Timers
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainlineOperationGUI.TimerSendCommand.Interval = 0
    Let MainlineOperationGUI.TimerSendCommand.Enabled = False
    Let MainlineOperationGUI.TimerSpeedChange.Interval = 0
    Let MainlineOperationGUI.TimerSpeedChange.Enabled = False
    Let MainlineOperationGUI.timerairmainReservoir.Interval = 0
    Let MainlineOperationGUI.timerairmainReservoir.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarAirMainReservoir.Value = 0
    Let MainlineOperationGUI.TimerAirAutomaticBrake.Interval = 0
    Let MainlineOperationGUI.TimerAirAutomaticBrake.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.barairautomaticbrake.Value = 0
    Let MainlineOperationGUI.TimerAirIndependentBrake.Interval = 0
    Let MainlineOperationGUI.TimerAirIndependentBrake.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarAirIndependentBrake.Value = 0
    Let MainlineOperationGUI.timerthrottledelay.Interval = 0
    Let MainlineOperationGUI.timerthrottledelay.Enabled = False
    Let MainlineOperationGUI.TimerRadiatorFans.Interval = 0
    Let MainlineOperationGUI.TimerRadiatorFans.Enabled = False
    Let MainlineOperationGUI.TimerRPM.Interval = 0
    Let MainlineOperationGUI.TimerRPM.Enabled = False
    Let MainlineOperationGUI!timerdeadmann.Interval = 0
    Let MainlineOperationGUI!timerdeadmann.Enabled = False
    Let MainlineOperationGUI.TimerOilPressure.Interval = 0
    Let MainlineOperationGUI.TimerOilPressure.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarOilPressure.Value = 0
    Let MainlineOperationGUI.TimerOilTemperature.Interval = 0
    Let MainlineOperationGUI.TimerOilTemperature.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarOilTemperature.Value = 0
    Let MainlineOperationGUI.TimerWaterPressure.Interval = 0
    Let MainlineOperationGUI.TimerWaterPressure.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarWaterPressure.Value = 0
    Let MainlineOperationGUI.TimerWaterTemperature.Interval = 0
    Let MainlineOperationGUI.TimerWaterTemperature.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarWaterTemperature.Value = 0
    Let MainlineOperationGUI.timerfuelconsumption.Interval = 0
    Let MainlineOperationGUI.timerfuelconsumption.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarFuelTank.Value = 0
    Let MainlineOperationGUI.TimerOilConsumption.Interval = 0
    Let MainlineOperationGUI.TimerOilConsumption.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarOilTank.Value = 0
    Let MainlineOperationGUI.TimerWaterConsumption.Interval = 0
    Let MainlineOperationGUI.TimerWaterConsumption.Enabled = False
    Let MainlineOperationGuiDiesel1Screen.BarWaterTank.Value = 0
    Let MainlineOperationGuiDiesel1Screen.BarSandTank.Value = 0


' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Off Virtual Cab Signals
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainlineOperationGUI!OptionVirtualcabSignalYes.Value = True Then
        Let MainlineOperationGUI!TimerVirtualCabSignal.Interval = 0
        Let MainlineOperationGUI!TimerVirtualCabSignal.Enabled = False
    End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Now, Turn off the Sound and
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Let MainlineOperationGUI!ConsistControlSpeed.Value = "1"
'
'    If InternetSettings.Winsock.State = sckConnected Then
'        If InternetSettings.OptionHost.value = True Then
'            Call MainlineOperationGUI.SetLocomotiveNumber
'            Call MainlineOperationGUI.SetSpeed
'            Call MainScreen.SendCommandviaTrackQ
'        End If
'    ElseIf InternetSettings.Winsock.State = sckClosed Then
'       If internetsettings.OptionStandAlone.value = true then
'           Call MainlineOperationGUI.SetLocomotiveNumber
'           Call MainlineOperationGUI.SetSpeed
'           Call MainScreen.SendCommandviaTrackQ
'       End if
'    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn the Sound Decoder Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
    
        If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
            
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
        ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
        End If
    
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If

        If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
            
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
        ElseIf MainlineOperationGUI!checkboxqsi.Value = vbChecked Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
        End If

        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If
    End If ' If decoder Equipped
    Let ButtonCloseGUI.Enabled = True
Return

End Sub

Public Sub ButtonFillFuel_Click()
    
    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Fill Fuel, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Fuel, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateFillFuel
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Fill Fuel, confirmed." Then
            GoSub UpdateFillFuel
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Fuel, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
            
        End If
    Else
        GoSub UpdateFillFuel
    End If
    
Exit Sub
    
UpdateFillFuel:
    
    Let SoundRightScreen.Notify = False
    Let SoundRightScreen.Wait = False
    Let SoundRightScreen.Shareable = False
    Let SoundRightScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundRightScreen.Command = "Open"
    Let SoundRightScreen.Notify = True
    Let SoundRightScreen.Command = "Play"
 
    Let TextBoxFuelTank.Text = Val(BarFuelTank.Max) - 100
    Let BarFuelTank.Value = Val(TextBoxFuelTank.Text)

Return

End Sub

Public Sub ButtonFillOil_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Fill Oil, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Oil, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateFillOil
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Fill Oil, confirmed." Then
            GoSub UpdateFillOil
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Oil, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
            
        End If
    Else
        GoSub UpdateFillOil
    End If
    
Exit Sub

UpdateFillOil:


    Let SoundRightScreen.Notify = False
    Let SoundRightScreen.Wait = False
    Let SoundRightScreen.Shareable = False
    Let SoundRightScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundRightScreen.Command = "Open"
    Let SoundRightScreen.Notify = True
    Let SoundRightScreen.Command = "Play"

    Let TextBoxOilTank.Text = Val(BarOilTank.Max) - 10
    Let BarOilTank.Value = Val(TextBoxOilTank.Text)

End Sub

Public Sub ButtonFillSand_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Fill Sand, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Sand, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateFillSand
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Fill Sand, confirmed." Then
            GoSub UpdateFillSand
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Sand, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateFillSand
    End If
    
Exit Sub

UpdateFillSand:

    Let SoundRightScreen.Notify = False
    Let SoundRightScreen.Wait = False
    Let SoundRightScreen.Shareable = False
    Let SoundRightScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundRightScreen.Command = "Open"
    Let SoundRightScreen.Notify = True
    Let SoundRightScreen.Command = "Play"
 
    Let TextBoxSandTank.Text = Val(BarSandTank.Max) - 10
    Let BarSandTank.Value = Val(TextBoxSandTank.Text)

End Sub

Public Sub ButtonFillWater_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Fill Water, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Water, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateFillWater
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "Fill Water, confirmed." Then
            GoSub UpdateFillWater
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Fill Water, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateFillWater
    End If
    
Exit Sub

UpdateFillWater:

    Let SoundRightScreen.Notify = False
    Let SoundRightScreen.Wait = False
    Let SoundRightScreen.Shareable = False
    Let SoundRightScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundRightScreen.Command = "Open"
    Let SoundRightScreen.Notify = True
    Let SoundRightScreen.Command = "Play"

    Let TextBoxWaterTank.Text = Val(BarWaterTank.Max) - 43
    Let BarWaterTank.Value = Val(TextBoxWaterTank.Text)

End Sub

Private Sub ButtonHelp_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!optionclient.Value = True Then
            If ButtonHelp.Caption = "&Help is Off" Then
                Let ButtonHelp.Caption = "&Help is On"
            Else
                Let ButtonHelp.Caption = "&Help is Off"
            End If
        Else
            Let ButtonHelp.Caption = "&Help is Off"
            MsgBox "When using the internet as the host (or server) you cannot use the help button. Help is only available for the remote engineer or without the internet connection.", vbExclamation + vbOKOnly, "Automatic Train Control - Internet Help User Error"
        End If
    ElseIf InternetSettings!Winsock.State = sckClosed Then
        If InternetSettings.OptionStandAlone.Value = True Then
            If ButtonHelp.Caption = "&Help is Off" Then
                Let ButtonHelp.Caption = "&Help is On"
            Else
                Let ButtonHelp.Caption = "&Help is Off"
            End If
        End If
    End If

End Sub



Public Sub ButtonScreenLeft_Click(Index As Integer)

    If InternetSettings!Winsock.State = sckConnected Then
       If InternetSettings!textboxincomingdata.Text = "Left Computer Screen " & Index & ", requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Left Computer Screen " & Index & ", confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateLeftScreen
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "Left Computer Screen " & Index & ", confirmed." Then
            GoSub UpdateLeftScreen
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Left Computer Screen " & Index & ", requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with outany internet connection
        GoSub UpdateLeftScreen
    End If

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateLeftScreen:

    Let SoundLeftScreen.Notify = False
    Let SoundLeftScreen.Wait = False
    Let SoundLeftScreen.Shareable = False
    Let SoundLeftScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundLeftScreen.Command = "Open"
    Let SoundLeftScreen.Notify = True
    Let SoundLeftScreen.Command = "Play"
 
    For X = 1 To 8
        Let PictureBoxComputerScreenLeft(X).Visible = False
    Next X
        
    Let PictureBoxComputerScreenLeft(Index).Visible = True
    
    If Index = 8 Then
        Let VideoCaptureNotes.Text = VideoCaptureNotes.Text & "Current frame rate is " & VideoCapture.GetActualFrameRate() & " fps." & vbCrLf
    End If
    
Return

End Sub

Public Sub ButtonScreenRight_Click(Index As Integer)

    If InternetSettings!Winsock.State = sckConnected Then
       If InternetSettings!textboxincomingdata.Text = "Right Computer Screen " & Index & ", requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Right Computer Screen " & Index & ", confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateRightScreen
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "Right Computer Screen " & Index & ", confirmed." Then
            GoSub UpdateRightScreen
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Right Computer Screen " & Index & ", requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with outany internet connection
        GoSub UpdateRightScreen
    End If

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateRightScreen:

    Let SoundLeftScreen.Notify = False
    Let SoundLeftScreen.Wait = False
    Let SoundLeftScreen.Shareable = False
    Let SoundLeftScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundLeftScreen.Command = "Open"
    Let SoundLeftScreen.Notify = True
    Let SoundLeftScreen.Command = "Play"
 
    For X = 1 To 8
        Let PictureBoxComputerScreenRight(X).Visible = False
    Next X
    PictureBoxComputerScreenRight(Index).Visible = True

Return

End Sub



Private Sub Form_LinkOpen(Cancel As Integer)

End Sub

Private Sub Form_Resize()
    
    If MainlineOperationGuiDiesel1Screen.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineOperationGuiDiesel1Screen.WindowState = vbNormal Then
    
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

Public Sub PictureBoxComputerScreenRight_Click(Index As Integer)
    
    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Right Computer Screen " & Index & ", requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Right Computer Screen " & Index & ", confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateRightScreen
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Right Computer Screen " & Index & ", confirmed." Then
            GoSub UpdateRightScreen
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Right Computer Screen " & Index & ", requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
            
        End If
    Else
        ' The program must be running with outany internet connection
        GoSub UpdateRightScreen
    End If

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateRightScreen:

    Let SoundRightScreen.Notify = False
    Let SoundRightScreen.Wait = False
    Let SoundRightScreen.Shareable = False
    Let SoundRightScreen.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundRightScreen.Command = "Open"
    Let SoundRightScreen.Notify = True
    Let SoundRightScreen.Command = "Play"
    
    For X = 1 To 8
        Let PictureBoxComputerScreenRight(X).Visible = False
    Next X

        Let PictureBoxComputerScreenRight(Index).Visible = True
Return

End Sub





Private Sub Form_Activate()

    DoEvents
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, Form, Activate" & vbCrLf
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
        If TemporaryScreen = "Mainline Operation GUI Diesel1 Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen"
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
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, stack is full, overflow."
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
            For OutsideLoop = 0 To Val(255)
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
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Room Lighting Control - Are the light truly off?
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.optionHost.Value = True Then
        If RoomLightingControl!CheckBoxUseRoomLightingControl.Value = vbChecked Then
            Call RoomLightingControl.ButtonManualControl_Click
        End If
        If RoomLightingControl!CheckBoxUseRoomLightingControl.Value = vbChecked Then
            If RoomLightingControl!ButtonManualControl.Caption = "&Lights are On" Then
                Call RoomLightingControl.ButtonManualControl_Click
            End If
        End If
    End If
' =============================================================================================================================================================================
' Turn Audio/Video On
' =============================================================================================================================================================================
    If InternetSettings.optionHost = True Then
        If InternetSettings.NetConnect.Connected = True Then
            Call VideoModule.VideoOn
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Start Broadcasting
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Video Settings Screen"
            Let Ini.Parameter = "VideoStreamingClientsMaximum"
            VideoCapture.StartBroadcast 20103, Val(Ini.Value)
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Client Mode
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf InternetSettings.optionclient = True Then
        If InternetSettings.NetConnect.Connected = True Then
            If InternetSettings.checkboxclientstreamtypebroadcast.Value = vbChecked Then
                VideoCapture.PlayerOpen "http://atc.server" & CStr(Val(InternetSettings!ComboServerName.ListIndex)) & ".lovethosetrains.com:20103"
                DoEvents
                VideoCapture.PlayerStart
            ElseIf InternetSettings.checkboxclientstreamtypeserver.Value = vbChecked Then
                VideoCapture.DisplayRemote "http://atc.server" & CStr(Val(InternetSettings!ComboServerName.ListIndex)) & ".lovethosetrains.com", True
            End If
        End If
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Activate TimerCheckWinsock
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let TimerCheckWinsock.Enabled = True
    ElseIf InternetSettings.OptionStandAlone = True Then
        If VideoCapture.GetVideoDeviceCount <> 0 Then
            Call VideoModule.VideoOn
         End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display Locomotive Number in GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelLocomotiveNumberDisplay.Caption = Val(MainlineOperationGUI!ComboBoxLocomotiveNumber.Text)
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialize all function in decoders to off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.optionclient.Value <> "True" Then
       If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
           Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction01234
           Call MainScreen.SendCommandviaTrackQ
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction5678
           Call MainScreen.SendCommandviaTrackQ
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction9101112
           Call MainScreen.SendCommandviaTrackQ
       End If
       
       Let MainlineOperationGUI.ConsistControlFunction0.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction1.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction2.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction3.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction7.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction8.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction9.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction10.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction11.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction12.Value = vbUnchecked
    
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction01234
       Call MainScreen.SendCommandviaTrackQ
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction5678
       Call MainScreen.SendCommandviaTrackQ
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction9101112
       Call MainScreen.SendCommandviaTrackQ
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Call BalloonHelpUpdatePart01 ' for RadioPhone, ResetButtonLeft (Alerter), Reverser, Throttle, Horn, Bell and Sander
    Call BalloonHelpUpdatePart02 ' for Automatic Brake, Independent Brake, Light, Reset Right
    Call BalloonHelpUpdatePart03 ' for Buttons F1 to F8 on Right and Left Computer Screen
    Call BalloonHelpUpdatePart04 ' for Right and Left Computer Screens
    Call BalloonHelpUpdatePart05 ' Cab Window
    Call BalloonHelpUpdatePart06 ' Auxillary Switches
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel1 Screen, Form, Activate" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub Form_Deactivate()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, Form, Deactivate" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Diesel1 Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationGuiDiesel1Screen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationGuiDiesel1Screen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationGuiDiesel1Screen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationGuiDiesel1Screen.Height)

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
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Deactivate TImerCheckWinsock
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'If InternetSettings.OptionClient = True Then
        Let TimerCheckWinsock.Enabled = False
    'End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineOperationGuiDiesel1Screen.Hide
    Unload MainlineOperationGuiDiesel1Screen

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel1 Screen, Form, Deactivate" & vbCrLf
    End If ' Debug Tag
  
End Sub


Private Sub Form_Load()


   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, Form, Load" & vbCrLf
    End If ' Debug Tag
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Load Pictures
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryLocomotivePath As String
    Let TemporaryLocomotivePath$ = "\Graphics\Locomotive Diesel1\"
    Let PictureBoxLocomotiveCab.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "CabScreen(d1).bmp")
    Let PictureBoxResetLeft.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ResetLeft" & PictureBoxResetLeft.Tag & "(d1).bmp")
    Let PictureBoxSand.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Sand" & PictureBoxSand.Tag & "(d1).bmp")
    Let PictureBoxBell.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Bell" & PictureBoxBell.Tag & "(d1).bmp")
    Let PictureBoxHorn.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Horn" & PictureBoxHorn.Tag & "(d1).bmp")
    Let PictureBoxReverser.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Reverser" & PictureBoxReverser.Tag & "(d1).bmp")
    Let PictureBoxThrottle.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Throttle" & PictureBoxThrottle.Tag & "(d1).bmp")
    Let PictureBoxAutomaticBrake.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "AutomaticBrake" & PictureBoxAutomaticBrake.Tag & "(d1).bmp")
    Let TransPictureBoxIndependentBrake.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "IndependentBrake" & TransPictureBoxIndependentBrake.Tag & "(d1).bmp")
    Let TransPictureBoxResetRight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ResetRight" & TransPictureBoxResetRight.Tag & "(d1).bmp")
    Let PictureBoxLight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Light" & PictureBoxLight.Tag & "(d1).bmp")
    Let PictureBoxRadioPhone.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "RadioPhone" & PictureBoxRadioPhone.Tag & "(d1).bmp")
    Let PictureBoxSpeedometer.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Speedometer(d1).bmp")
    Let TransPictureBoxCabWindow.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "CabWindow(d1).bmp")
    Let TransPictureSunVisor.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "SunVisor(d1).bmp")
    Let TransPictureWipers.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Wipers(d1).bmp")
    Let TransPictureButtonSide(1).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ButtonSide0(d1).bmp")
    Let TransPictureButtonSide(2).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ButtonSide0(d1).bmp")
    Let TransPictureButtonSide(3).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ButtonSide0(d1).bmp")
    Let TransPictureButtonSide(4).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ButtonSide0(d1).bmp")
    Let TransPictureButtonSide(5).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ButtonSide0(d1).bmp")
    Let TransPictureAuxillarySwitch(0).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(1).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(2).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(3).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(4).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(5).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(6).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(7).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(8).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(9).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(10).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(11).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(12).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(13).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureAuxillarySwitch(14).Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Switch0(d1).bmp")
    Let TransPictureSignal.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Signal(d1).bmp")
    'Let TransPictureCup.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Cup(d1).bmp")

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Select Window State
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Let MainlineOperationGuiDiesel1Screen.Width = (PictureBoxLocomotiveCab.ScaleWidth + 6) * 15
    'Let MainlineOperationGuiDiesel1Screen.Height = (PictureBoxLocomotiveCab.ScaleHeight + 25) * 15
    'If MainlineOperationGuiDiesel1Screen.Width > Screen.Width Then MainlineOperationGuiDiesel1Screen.Width = Screen.Width
    'If MainlineOperationGuiDiesel1Screen.Height > Screen.Height Then MainlineOperationGuiDiesel1Screen.Height = Screen.Height
    Let MainlineOperationGuiDiesel1Screen.Left = (Screen.Width - MainlineOperationGuiDiesel1Screen.Width) / 2   ' Center form horizontally.
    Let MainlineOperationGuiDiesel1Screen.Top = (Screen.Height - MainlineOperationGuiDiesel1Screen.Height) / 2  ' Center form vertiCally.
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel1 Screen, Form, Load" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



Private Sub FtpControl_Timeout()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Ftp Control, connection to server timed out."
    End If
    
End Sub


Private Sub LabelVirtualCabSignalDistance_Change()

Let LabelvirtualcabSignalDistanceFeet.Caption = TextBoxVirtualCabSignalDistance.Text

End Sub

Private Sub menuAudioCodecDialog_Click()

    VideoCapture.ShowAudioCodecDlg
    
End Sub

Private Sub menuAudioFormatDialog_Click()

    VideoCapture.ShowAudioFormatDlg
    
End Sub


Private Sub menuAudioSourceDialog_Click()

    VideoCapture.ShowAudioSourceDlg
    
End Sub


Private Sub menuVideoCodecDialog_Click()

    VideoCapture.ShowVideoCodecDlg
    
End Sub

Private Sub menuVideoCrossbarDialog_Click()

    VideoCapture.ShowVideoCrossbarDlg
    
End Sub


Private Sub menuVideoFormatDialog_Click()

    VideoCapture.ShowVideoFormatDlg
    
End Sub


Private Sub menuVideoSourceDialog_Click()

    VideoCapture.ShowVideoSourceDlg
    
End Sub


Public Sub PictureBoxAutomaticBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ButtonEngineStart.Caption <> "&Start Engine" Then
    
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Automatic Brake, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateAutomaticBrake
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Automatic Brake, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateAutomaticBrake
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Automatic Brake, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateAutomaticBrake
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Automatic Brake, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateAutomaticBrake
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateAutomaticBrake
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Automatic Brake handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        Let BalloonHelpText1 = "The application of the trainline brake is at " & Str$(Int(Val(100 - (TextBoxStatusAutomaticBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Automatic Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxAutomaticBrake)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxAutomaticBrake, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Automatic Brake, unable to setup balloon help for 'PictureBoxAutomaticBrake' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateAutomaticBrake:

    If Button = vbLeftButton Then
        If Val(PictureBoxAutomaticBrake.Tag) < 15 Then
            Let PictureBoxAutomaticBrake.Tag = Val(PictureBoxAutomaticBrake.Tag) + 1
            If Val(PictureBoxAutomaticBrake.Tag) = 1 Then
                Let SoundBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticOpen.wav"
            Else
                Let SoundBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let SoundBrakeAutomatic.Command = "Open"
            Let SoundBrakeAutomatic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum braking, one hundred (100) percent for the automatic brakes (train line brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbRightButton Then
        If Val(PictureBoxAutomaticBrake.Tag) > 0 Then
            Let PictureBoxAutomaticBrake.Tag = Val(PictureBoxAutomaticBrake.Tag) - 1
            If PictureBoxAutomaticBrake.Tag = 0 Then
                Let SoundBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticClose.wav"
            Else
                Let SoundBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let SoundBrakeAutomatic.Command = "Open"
            Let SoundBrakeAutomatic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum braking, zero (0) precent for the automatic brakes (train line brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    
    
    Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake" & Trim$(PictureBoxAutomaticBrake.Tag) & "(d1).bmp"
    Let PictureBoxAutomaticBrake.Picture = LoadPicture(temp)
    
    If Val(PictureBoxThrottle.Tag) <> 0 And Val(PictureBoxReverser.Tag) <> 1 Then
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
        Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
        Let MainlineOperationGUI!timerdeadmann.Enabled = True
    Else
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
    End If

Return

End Sub
Public Sub PictureBoxBell_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Bell, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Bell, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateBell
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Bell, confirmed." Then
            GoSub UpdateBell
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Bell, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateBell
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If PictureBoxBell.Tag = "0" Then
            Let BalloonHelpText1 = "The bell is not active."
        Else 'If PictureBoxBell.Tag = "1" Then
            Let BalloonHelpText1 = "The bell is active."
        End If
        
        Let BalloonHelpText2 = "Bell"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxBell)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxBell, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Bell, unable to setup balloon help for 'PictureBoxBell' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateBell:

    Let SoundBell.Notify = False
    Let SoundBell.Wait = False
    Let SoundBell.Shareable = False
    Let SoundBell.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let SoundBell.Command = "Open"
    Let SoundBell.Notify = True
    Let SoundBell.Command = "Play"

    If PictureBoxBell.Tag = "0" Then
        Let PictureBoxBell.Tag = "1"

        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
        End If
    
'       DoEvents
'       Let MainlineOperationGUI!Wave1.Loop = True
'       Let MainlineOperationGUI!Wave1.filename = App.Path$ & "\Sounds\Graphics\Bell.wav"
'       Let MainlineOperationGUI!Wave1.Action = wAPlay
    
    ElseIf PictureBoxBell.Tag = "1" Then
        Let PictureBoxBell.Tag = "0"
    
        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
        End If

    End If

    Let PictureBoxBell.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\Bell" & PictureBoxBell.Tag & "(d1).bmp")
    
    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If
    End If
Return

End Sub

Public Sub PictureBoxHorn_Click()
        
    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Horn, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Horn, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateHorn
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Horn, confirmed." Then
            GoSub UpdateHorn
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Horn, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with out any internet connection
        GoSub UpdateHorn
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If PictureBoxHorn.Tag = "0" Then
            Let BalloonHelpText1 = "The horn is not active."
        ElseIf PictureBoxHorn.Tag = "1" Then
            Let BalloonHelpText1 = "The horn is active."
        End If
        
        Let BalloonHelpText2 = "Horn"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxHorn)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxHorn, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Horn, unable to setup balloon help for 'PictureBoxHorn' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

Exit Sub
        
' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateHorn:

    If PictureBoxHorn.Tag = "0" Then
        Let PictureBoxHorn.Tag = "1"
        Let temp = App.Path
        Let temp = temp + "\Graphics\Locomotive Diesel1\Horn1(d1).bmp"
        Let PictureBoxHorn.Picture = LoadPicture(temp)
        
        If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
        End If
   
    Else
        Let PictureBoxHorn.Tag = "0"
        Let temp = App.Path
        Let temp = temp + "\Graphics\Locomotive Diesel1\Horn0(d1).bmp"
        Let PictureBoxHorn.Picture = LoadPicture(temp)
        
        
        If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "0" Then
            Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "1" Then
            Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "3" Then
            Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "4" Then
            Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "5" Then
            Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "6" Then
            Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "7" Then
            Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "8" Then
            Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "9" Then
            Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "10" Then
            Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "11" Then
            Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "12" Then
            Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
        End If

    End If
    
    Let SoundHorn.Notify = False
    Let SoundHorn.Wait = False
    Let SoundHorn.Shareable = False
    Let SoundHorn.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let SoundHorn.Command = "Open"
    Let SoundHorn.Notify = True
    Let SoundHorn.Command = "Play"
       
    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If
    End If

Return

End Sub


Public Sub PictureBoxLight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Light, left button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Light, left button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbLeftButton
            GoSub UpdateLight
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Light, left button confirmed." Then
            Let Button = vbLeftButton
            GoSub UpdateLight
            Let InternetSettings.textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Light, right button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Light, right button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbRightButton
            GoSub UpdateLight
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Light, right button confirmed." Then
            Let Button = vbRightButton
            GoSub UpdateLight
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            If Button = vbLeftButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Light, left button requested."
            ElseIf Button = vbRightButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Light, right button requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with outany internet connection
        GoSub UpdateLight
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
        If PictureBoxLight.Tag = "0" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently off."
        ElseIf PictureBoxLight.Tag = "1" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on and dimmed."
        ElseIf PictureBoxLight.Tag = "2" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on."
        End If
        
        Let BalloonHelpText2 = "Headlight"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxLight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxLight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Light, unable to setup balloon help for 'PictureBoxRadioLight' control."
            End If
        End If
        
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault
        
    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateLight:

    If Button = vbRightButton Then
        If Val(PictureBoxLight.Tag) < 2 Then
            Let PictureBoxLight.Tag = Val(PictureBoxLight.Tag) + 1
        End If
    ElseIf Button = vbLeftButton Then
        If Val(PictureBoxLight.Tag) > 0 Then
            Let PictureBoxLight.Tag = Val(PictureBoxLight.Tag) - 1
        End If
    End If
    
    Let SoundLight.Notify = False
    Let SoundLight.Wait = False
    Let SoundLight.Shareable = False
    'Let SoundLight.DeviceType = "WaveAudio"
    Let SoundLight.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let SoundLight.Command = "Open"
    Let SoundLight.Command = "Play"
    
    Let temp = App.Path
    Let temp = temp + "\Graphics\Locomotive Diesel1\Light"
    Let temp = temp + Trim$(PictureBoxLight.Tag)
    Let temp = temp + "(d1).bmp"
    Let PictureBoxLight.Picture = LoadPicture(temp)
    
    If PictureBoxLight.Tag = "0" Then
    
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
        
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
            
    ElseIf PictureBoxLight.Tag = "1" Then
                    
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
        
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
            
    ElseIf PictureBoxLight.Tag = "2" Then
            
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
            
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
    End If

    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If
    End If

Return

End Sub

Public Sub PictureBoxRadioPhone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "RadioPhone, left button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "RadioPhone, left button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbLeftButton
            GoSub UpdateRadioPhone
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "RadioPhone, left button confirmed." Then
            Let Button = vbLeftButton
            GoSub UpdateRadioPhone
            Let InternetSettings.textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "RadioPhone, right button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "RadioPhone, right button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbRightButton
            GoSub UpdateRadioPhone
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "RadioPhone, right button confirmed." Then
            Let Button = vbRightButton
            GoSub UpdateRadioPhone
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            If Button = vbLeftButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "RadioPhone, left button requested."
            ElseIf Button = vbRightButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "RadioPhone, right button requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with outany internet connection
        GoSub UpdateRadioPhone
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
        If Val(PictureBoxRadioPhone.Tag) = 0 Then
            Let BalloonHelpText1 = "The radio phone is off the hook (in use)."
        ElseIf Val(PictureBoxRadioPhone.Tag) = 1 Then
            Let BalloonHelpText1 = "The radio phone is on the hook (not in use)."
        End If
        Let BalloonHelpText2 = "Radio Phone"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxRadioPhone)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxRadioPhone, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Radio Phone, unable to setup balloon help for 'PictureBoxRadioPhone' control."
            End If
        End If
        
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault
        
    End If
    
    Exit Sub
    
' =============================================================================================================================================================================
'
' =============================================================================================================================================================================
    
UpdateRadioPhone:
    
        If Button = vbRightButton Then
            Let MainlineOperationGuiDiesel1Screen!TextBoxVirtualCabSignalStatusApproaching.Text = "Clear"
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!textboxvirtualCabsignalstatuscurrent.Text = "Clear"
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let MainlineOperationGuiDiesel1Screen!PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
        End If
        
        Let SoundRadioPhone.Notify = False
        Let SoundRadioPhone.Wait = False
        Let SoundRadioPhone.Shareable = False
        Let SoundRadioPhone.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let SoundRadioPhone.Command = "Open"
        Let SoundRadioPhone.Notify = True
        Let SoundRadioPhone.Command = "Play"
        
        If Val(PictureBoxRadioPhone.Tag) = 0 Then
            Let PictureBoxRadioPhone.Tag = 1
            Let PictureBoxRadioPhone.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\RadioPhone1(d1).bmp")
        Else
            Let PictureBoxRadioPhone.Tag = 0
            Let PictureBoxRadioPhone.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\RadioPhone0(d1).bmp")
        End If
    
    Return

End Sub

Public Sub PictureBoxResetLeft_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Reset Left, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Reset Left, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateResetLeft
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Reset Left, confirmed." Then
            GoSub UpdateResetLeft
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Reset Left, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateResetLeft
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
'        Dim BalloonHelpText1 As String
'        Dim BalloonHelpText2 As String
'        Dim BalloonHelpSetup As Long
'        Dim BalloonHelpFont As New StdFont
'        Dim BalloonHelpVisibleTime As Long
'        Dim BalloonHelpTimeDelay As Long
'        Dim BalloonHelpShadow As Boolean
'        Dim BalloonHelpCenter As Boolean
'        Dim BalloonHelpShowOnDemand As Boolean
'        Dim BalloonHelpOpacity As Byte
'        Dim BalloonHelpWaveFile As String
'
'        Let Ini.Filename = App.Path$ & "\Atc.ini"
'        Let Ini.Application = "All Screens"
'        Let Ini.Parameter = "BalloonHelpFontName"
'        Let BalloonHelpFont.Name = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontSize"
'        Let BalloonHelpFont.Size = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontBold"
'        Let BalloonHelpFont.Bold = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontItalic"
'        Let BalloonHelpFont.Italic = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontUnderline"
'        Let BalloonHelpFont.Underline = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour1"
'        Let BalloonHelpColour1 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour2"
'        Let BalloonHelpColour2 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour3"
'        Let BalloonHelpColour3 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpVisibleTime"
'        Let BalloonHelpVisibleTime = Ini.Value
'        Let Ini.Parameter = "BalloonHelpDelayTime"
'        Let BalloonHelpDelayTime = Ini.Value
'        Let Ini.Parameter = "BalloonHelpShadow"
'        Let BalloonHelpShadow = Ini.Value
'        Let Ini.Parameter = "BalloonHelpCenter"
'        Let BalloonHelpCenter = Ini.Value
'        Let Ini.Parameter = "BalloonHelpShowOnDemand"
'        Let BalloonHelpShowOnDemand = Ini.Value
'        Let Ini.Parameter = "BalloonHelpWaveFile"
'        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
'        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
'        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
'            BalloonHelpOpacity = 255
'        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
'            Let Ini.Parameter = "BalloonHelpOpacity"
'            Let BalloonHelpOpacity = Ini.Value
'        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
'                Let balloonhelp.Speech = False
'        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
'            Let balloonhelp.Speech = True
'            Let balloonhelp.Voice = 0
'            Let BalloonHelpWaveFile = ""
'        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
'
'        Let BalloonHelpText1 = "The deadmann switch is used to reset the alerter."
'        Let BalloonHelpText2 = "Deadmann Switch"
'
'        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxresetleft)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxrestleft, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, Balloonhelpwavefile)
'        If BalloonHelpSetup = 0 Then
'            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
'                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
'                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
'                Let Ini.Filename = App.Path$ & "\Atc.log"
'                Let Ini.Application = "Log Errors"
'                Let Ini.Parameter = Date$ & " " & Time$
'                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Radio Phone, unable to setup balloon help for 'PictureBoxRadioPhone' control."
'            End If
'        End If
'
'        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault
'
'    End If

    Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateResetLeft:

    Let sounddeadmann.Notify = False
    Let sounddeadmann.Wait = False
    Let sounddeadmann.Shareable = False
    Let sounddeadmann.Command = "Stop"
    'Let sounddeadmann.Notify = False
    'Let sounddeadmann.Command = "Close"
    
    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Clicking the Reset button with sound
    
    Let PictureBoxResetLeft.Tag = "1"
    Let temp = App.Path & "\Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp"
    Let PictureBoxResetLeft.Picture = LoadPicture(temp)
    
    Let SoundResetLeft.Notify = False
    Let SoundResetLeft.Wait = False
    Let SoundResetLeft.Shareable = False
    Let SoundResetLeft.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let SoundResetLeft.Command = "Open"
    Let SoundResetLeft.Notify = True
    Let SoundResetLeft.Command = "Play"
    
    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' End of Routin to Reset Button
 
    Let labeldeadmann.Visible = False
      
    If Val(PictureBoxThrottle.Tag) <> 0 Then
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
        Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
        Let MainlineOperationGUI!timerdeadmann.Enabled = True
    Else
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
    End If
    
    Let PictureBoxResetLeft.Tag = "0"
    Let temp = App.Path & "\Graphics\Locomotive Diesel1\ResetLeft0(d1).bmp"
    Let PictureBoxResetLeft.Picture = LoadPicture(temp)

Return

End Sub





Public Sub PictureBoxReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ButtonEngineStart.Caption <> "&Start Engine" Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Reverser, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Reverser, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateReverser
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Reverser, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateReverser
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Reverser, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Reverser, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateReverser
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Reverser, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateReverser
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Reverser, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Reverser, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateReverser
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Reverser handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If PictureBoxReverser.Tag = 0 Then
            Let BalloonHelpText1 = "This reverser handle is in the forward position."
        ElseIf PictureBoxReverser.Tag = 1 Then
            Let BalloonHelpText1 = "This reverser handle is in the neutral position."
        ElseIf PictureBoxReverser.Tag = 2 Then
            Let BalloonHelpText1 = "This reverser handle is in the reverse position."
        End If
        Let BalloonHelpText2 = "Reverser Handle"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxreverser)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxReverser, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Reverser, unable to setup balloon help for 'PictureReverser' control."
            End If
        End If
    End If
    
    Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault
    
Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateReverser:

    If Val(PictureBoxThrottle.Tag) = 0 Then

        If Val(TextBoxStatusSpeedStepModified.Text) = 0 Then
    
            If Button = vbLeftButton Then
                If Val(PictureBoxReverser.Tag) < 2 Then
                    Let SoundReverser.Command = "Close"
                    Let SoundReverser.Filename = App.Path$ & "\Sounds\Graphics\ControlReverser.wav"
                    Let SoundReverser.Command = "Open"
                    Let SoundReverser.Command = "Play"
                    Let PictureBoxReverser.Tag = Val(PictureBoxReverser.Tag) + 1
                End If
            Else ' If Button = vbRightButton Then
                If Val(PictureBoxReverser.Tag) > 0 Then
                    Let SoundReverser.Command = "Close"
                    Let SoundReverser.Filename = App.Path$ & "\Sounds\Graphics\ControlReverser.wav"
                    Let SoundReverser.Command = "Open"
                    Let SoundReverser.Command = "Play"
                    Let PictureBoxReverser.Tag = Val(PictureBoxReverser.Tag) - 1
                End If
            End If
    
            Let PictureBoxReverser.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\Reverser" & Trim$(PictureBoxReverser.Tag) & "(d1).bmp")
    
            If Val(PictureBoxReverser.Tag) = 2 Then
                Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbChecked
                Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
                Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
            ElseIf Val(PictureBoxReverser.Tag) = 1 Then
                Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
                Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbChecked
                Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
            ElseIf Val(PictureBoxReverser.Tag) = 0 Then
                Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
                Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
                Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked
            End If
        Else
            If MainlineOperationGuiDiesel1Screen.ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You cannot move the reverser handle unless the speed of the locomotive is nill."
                MsgBox Message, vbExclamation, "Engineer Error - Control Interlock"
            End If
        End If
    Else
        If MainlineOperationGuiDiesel1Screen.ButtonHelp.Caption = "&Help is On" Then
            Let Message = "You cannot move the reverser handle unless the throttle is return to notch zero."
            MsgBox Message, vbExclamation, "Engineer Error - Control Interlock"
        End If
    End If
    
    ' ---------------------------------------------------------------------------------------------------------------------------
    ' Update Label with Combined Power
    '
    
    Let LabelCombinedPower.Caption = "Combined Power" & Chr$(13)
    
    If Val(PictureBoxReverser.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Forward - "
    ElseIf Val(PictureBoxReverser.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Reverse - "
    End If
    
    If Val(PictureBoxThrottle.Tag) < 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Dynamic Brake" & Chr(13)
    End If
    
    If Abs(PictureBoxThrottle.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Zero"
    ElseIf Abs(PictureBoxThrottle.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch One"
    ElseIf Abs(PictureBoxThrottle.Tag) = 2 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Two"
    ElseIf Abs(PictureBoxThrottle.Tag) = 3 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Three"
    ElseIf Abs(PictureBoxThrottle.Tag) = 4 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Four"
    ElseIf Abs(PictureBoxThrottle.Tag) = 5 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Five"
    ElseIf Abs(PictureBoxThrottle.Tag) = 6 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Six"
    ElseIf Abs(PictureBoxThrottle.Tag) = 7 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Seven"
    ElseIf Abs(PictureBoxThrottle.Tag) = 8 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Eight"
    End If
    
Return
    
End Sub
Public Sub PictureBoxSand_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Sand, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Sand, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateSand
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Sand, confirmed." Then
            GoSub UpdateSand
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Sand, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateSand
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If PictureBoxSand.Tag = "0" Then
            Let BalloonHelpText1 = "The sand is not active."
        Else 'If PictureBoxSand.Tag = "1" Then
            Let BalloonHelpText1 = "The sand is active."
        End If
        
        Let BalloonHelpText2 = "Sand"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxSand)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSand, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Sand, unable to setup balloon help for 'PictureBoxSand' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateSand:

    If Val(PictureBoxSand.Tag) = 0 Then
        Let PictureBoxSand.Tag = "1"
        Let SoundSand.Notify = False
        Let SoundSand.Wait = False
        Let SoundSand.Shareable = False
        Let SoundSand.Filename = App.Path$ & "\Sounds\Graphics\Sand.wav"
        Let MainlineOperationGUI!TimerSandConsumption.Interval = 2500 * (Val(ScrollTimeAdjustment) / 100 + 1)
    Else
        Let PictureBoxSand.Tag = "0"
        Let SoundSand.Notify = False
        Let SoundSand.Wait = False
        Let SoundSand.Shareable = False
        Let SoundSand.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    End If
    Let SoundSand.Command = "Open"
    Let SoundSand.Notify = True
    Let SoundSand.Command = "Play"
    
    Let Temporary$ = App.Path$
    Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel1\Sand"
    Let Temporary$ = Temporary$ + PictureBoxSand.Tag
    Let Temporary$ = Temporary$ + "(d1).bmp"
    Let PictureBoxSand.Picture = LoadPicture(Temporary$)
Return

End Sub

Public Sub PictureBoxThrottle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ButtonEngineStart.Caption <> "&Start Engine" Then
        If InternetSettings!Winsock.State = sckConnected Then

            If InternetSettings!textboxincomingdata.Text = "Throttle, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Throttle, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateThrottle
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Throttle, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateThrottle
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Throttle, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Throttle, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateThrottle
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Throttle, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateThrottle
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Throttle, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Throttle, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateThrottle
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Throttle handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If PictureBoxThrottle.Tag = -8 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position eight."
        ElseIf PictureBoxThrottle.Tag = -7 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position seven."
        ElseIf PictureBoxThrottle.Tag = -6 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position six."
        ElseIf PictureBoxThrottle.Tag = -5 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position five."
        ElseIf PictureBoxThrottle.Tag = -4 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position four."
        ElseIf PictureBoxThrottle.Tag = -3 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position three."
        ElseIf PictureBoxThrottle.Tag = -2 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position two."
        ElseIf PictureBoxThrottle.Tag = -1 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position one."
        ElseIf PictureBoxThrottle.Tag = 0 Then
            Let BalloonHelpText1 = "The throttle is in the nuetral position."
        ElseIf PictureBoxThrottle.Tag = 1 Then
            Let BalloonHelpText1 = "The throttle is in first notch position."
        ElseIf PictureBoxThrottle.Tag = 2 Then
            Let BalloonHelpText1 = "The throttle is in second notch position."
        ElseIf PictureBoxThrottle.Tag = 3 Then
            Let BalloonHelpText1 = "The throttle is in third notch position."
        ElseIf PictureBoxThrottle.Tag = 4 Then
            Let BalloonHelpText1 = "The throttle is in fourth notch position."
        ElseIf PictureBoxThrottle.Tag = 5 Then
            Let BalloonHelpText1 = "The throttle is in fifth notch position."
        ElseIf PictureBoxThrottle.Tag = 6 Then
            Let BalloonHelpText1 = "The throttle is in sixth notch position."
        ElseIf PictureBoxThrottle.Tag = 7 Then
            Let BalloonHelpText1 = "The throttle is in seventh notch position."
        Else 'If PictureBoxThrottle.Tag = 8 Then
            Let BalloonHelpText1 = "The throttle is in eighth notch position."
        End If
        
        Let BalloonHelpText2 = "Throttle Handle"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxreverser)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxReverser, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Throttle, unable to setup balloon help for 'PictureBoxThrottle' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If
    
Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateThrottle:
    
    If labeldeadmann.Visible = False Then
        If Val(TextBoxStatusThrottleDelay.Text) = 0 Then
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' -------------------------------------------------------------------------------------------------------------------------------------------
            If Button = vbLeftButton Then
                Let OkToNotchUp = True
                If Val(PictureBoxThrottle.Tag) > 7 Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The maximum engine revolution has been reached with the throttle in notch eight."
                        MsgBox Message, vbExclamation, "Engineer Error - Maximum Throttle"
                    End If
                End If
            
                If Val(BarFuelTank.Value) < Val(TextBoxStatusFuelTooLow.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of diesel fuel. You currently have less then " + Str$(Val(TextBoxStatusFuelTooLow.Text))
                        Let Message = Message + " US gallons left."
                        MsgBox Message, vbExclamation, "Engineer Error - Fuel Reserve"
                    End If
                End If
                If Val(BarSandTank.Value) < Val(textboxstatussandtoolow.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of sand. You currently have " + Str$(Val(BarSandTank.Value)) + " pounds of sand left. It is recommended that you have at" + Chr$(13)
                        Let Message = Message + "least " + Str$(Val(textboxstatussandtoolow.Text)) + " pounds of sand."
                        MsgBox Message, vbExclamation, "Engineer Error - Sand Reserve"
                    End If
                End If
                
                If Val(BarWaterTank.Value) < Val(TextBoxStatusWaterTooLow.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of water. You currently have " + Str$(Val(BarWaterTank.Value)) + " US gallons of water left. It is recommended that you have at" + Chr$(13)
                        Let Message = Message + "least " + Str$(Val(TextBoxStatusWaterTooLow.Text)) + " US gallons of water."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Reserve"
                    End If
                End If
                
                If Val(BarOilTank.Value) < Val(TextBoxStatusOilTooLow.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of engine lubricating oil. You currently have " + Str$(Val(BarOilTank.Value)) + " US gallons of lubricating left. It is" + Chr$(13)
                        Let Message = Message + "recommended that you have at least " + Str$(Val(TextBoxStatusOilTooLow.Text)) + " US gallons of oil. "
                        MsgBox Message, vbExclamation, "Engineer Error - Oil Reserve"
                    End If
                End If
                
                If Val(BarOilPressure.Value) < Val(TextBoxStatusOilTooLowPressure.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine oil pressure has not reached a pressure to operate the" + Chr$(13)
                        Let Message = Message + "locomotive safely. The current oil pressure is" + Str$(Val(BarOilPressure.Value)) + " pounds per square" + Chr$(13)
                        Let Message = Message + "inch. We recommend that you wait until the pressure reaches at" + Chr$(13)
                        Let Message = Message + "least" + Str$(Val(TextBoxStatusOilTooLowPressure.Text)) + " pounds per square inch."
                        MsgBox Message, vbExclamation, "Engineer Error - Oil Pressure"
                    End If
                End If
                
                If Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text) * 2 / 3 Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine oil temperature has not reached a temperature to operate the" + Chr$(13)
                        Let Message = Message + "locomotive safely. The current oil temperature is" + Str$(Val(BarOilTemperature.Value)) + " degrees. We recommend" + Chr$(13)
                        Let Message = Message + "that you wait until the temperature reaches at least" + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees."
                        MsgBox Message, vbExclamation, "Engineer Error- Oil Temperature"
                    End If
                End If
                
                If Val(BarWaterPressure.Value) < Val(TextBoxStatusWaterTooLowPressure.Text) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine water pressure has not reached a pressure to operate the " + Chr$(13)
                        Let Message = Message + "locomotive safely. The current water pressure is " + Str$(Val(BarWaterPressure.Value)) + " pounds per square inch. We" + Chr$(13)
                        Let Message = Message + "recommend that you wait until the pressure reaches" + Str$(Val(TextBoxStatusWaterTooLowPressure.Text)) + " pounds per square inch."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Pressure"
                    End If
                End If
            
                If Val(BarWaterTemperature.Value) < Val(TextBoxStatusWaterTooLowTemperature.Text) * 2 / 3 Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine water temperature has not reached a temperature to operate the " + Chr$(13)
                        Let Message = Message + "locomotive safely. The current water temperature is " + Str$(Val(BarWaterTemperature.Value)) + " degrees. We" + Chr$(13)
                        Let Message = Message + "recommend that you wait until the temperature reaches" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Temperature"
                    End If
                End If
            
                If ((Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text)) And _
                    (Val(PictureBoxThrottle.Tag) >= 1)) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Although the engine oil temperature is greater than" + Str$(Val(BarOilTemperature.Value) - 30) + "degrees, it has not reached" + Chr$(13)
                        Let Message = Message + "a temperature to operate above notch one. It is recommended that the oil temerature be greater than" + Chr$(13)
                        Let Message = Message + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                        MsgBox Message, vbExclamation, "Engineer Error- Oil Temperature"
                    End If
                End If
                
                If ((Val(BarWaterTemperature.Value) <= Val(TextBoxStatusWaterTooLowTemperature.Text)) And _
                    (Val(PictureBoxThrottle.Tag) >= 2)) Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Although the engine water temperature is greater than" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text) - 30) + " degrees, it has not reached" + Chr$(13)
                        Let Message = Message + "a temperature to operate above notch two. It is recommended that the water engine temperature be greater than" + Chr$(13)
                        Let Message = Message + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Temperature"
                    End If
               End If
                
                If Val(PictureBoxLight.Tag) = 0 And Val(PictureBoxReverser.Tag) <> 1 Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Just a little reminder that as an engineer, you should put the headlight on before moving the locomotive" + Chr$(13)
                        Let Message = Message + "in either direction. Please remeber Rule 17 the requires to to dim the lights in the yard and " + Chr$(13)
                        Let Message = Message + "and when approaching another locomotive."
                        MsgBox Message, vbExclamation, "Engineer Error - Lights"
                    End If
                End If
                
                If Val(PictureBoxBell.Tag) = 0 And _
                   Val(PictureBoxThrottle.Tag) = 0 And _
                   Val(PictureBoxReverser.Tag) <> 1 And _
                   Val(TextBoxStatusSpeedStepModified.Text) = 0 Then
                    Let OkToNotchUp = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Just a little reminder that as an engineer, you should be activating the bell prior" + Chr$(13)
                        Let Message = Message + " moving in any direction."
                        MsgBox Message, vbExclamation, "Engineer Error - Bell Activation"
                    End If
                End If
                
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Check the current Virtual Block Signal
    '
    
                If MainlineOperationGUI!OptionVirtualcabSignalYes.Value = True Then
                    If textboxvirtualCabsignalstatuscurrent.Text = "Stop" Then
                        Let OkToNotchUp = False
                        If ButtonHelp.Caption = "&Help is On" Then
                            Let Message = "Your throttle has been restricted to notch zero because you are currnetly in a" + Chr$(13)
                            Let Message = Message + "occupied block as advised by your dispatcher."
                            MsgBox Message, vbExclamation, "Engineer Error - Track Occupancy"
                        End If
                    ElseIf textboxvirtualCabsignalstatuscurrent.Text = "Restricted" Then
                        If Val(PictureBoxThrottle.Tag) >= 6 Then
                            Let OkToNotchUp = False
                            If ButtonHelp.Caption = "&Help is On" Then
                                Let Message = "Your throttle has been restricted to notch six because you are currently in a" + Chr(13)
                                Let Message = Message + "restricted block as advised by your dispatcher."
                                MsgBox Message, vbExclamation, "Engineer Error - Track Occupancy"
                            End If
                        End If
                    ElseIf textboxvirtualCabsignalstatuscurrent.Text = "Reduced" Then
                        If Val(PictureBoxThrottle.Tag) >= 3 Then
                            Let OkToNotchUp = False
                            If ButtonHelp.Caption = "&Help is On" Then
                                Let Message = "Your throttle has been restricted to notch three because you are currently in a" + Chr(13)
                                Let Message = Message + "reduced block as advised by your dispatcher."
                                MsgBox Message, vbExclamation, "Engineer Error - Track Occupancy"
                            End If
                        End If
                    End If 'cab signal
                End If ' virtual signal
            
        ' -----------------------------------------------------------------------------------------------------------------------
        '
                    
                If OkToNotchUp = True Then
           
                    Let PictureBoxThrottle.Tag = Val(PictureBoxThrottle.Tag) + 1
                    
                    Let SoundThrottle.Notify = False
                    Let SoundThrottle.Wait = False
                    Let SoundThrottle.Shareable = False
                    Let SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
                    Let SoundThrottle.Command = "Open"
                    Let SoundThrottle.Notify = True
                    Let SoundThrottle.Command = "Play"
                    
                    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                    
                        If Val(PictureBoxThrottle.Tag) > 0 And Val(PictureBoxThrottle.Tag) < 8 Then
                        
                            If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
                                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
                                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
                                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
                                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
                                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
                                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
                                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
                                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
                                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
                                Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
                                Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
                                Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
                                Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
                            End If
                            
                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If

                            If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
                                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
                                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
                                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
                                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
                                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
                                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
                                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
                                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
                                Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
                                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
                                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
                                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
                            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
                                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
                            End If
                            

                            If InternetSettings!Winsock.State = sckConnected Then
                                If InternetSettings.optionHost.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                            
                        ElseIf Val(PictureBoxThrottle.Tag) = 0 Then
                        
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
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                            
                        ElseIf Val(PictureBoxThrottle.Tag) < 0 Then
                        
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
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                        End If ' Throttle Position
                    End If ' If Sound Decoder Equipped
                End If ' If ok to Notch Up
            
                
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Button Right Routine
    '
    ' All these routines are for the right button when it was click'ed.
    '
                
            ElseIf Button = vbRightButton Then
    
                Let OkToNotchDown = True
            
                If Val(PictureBoxThrottle.Tag) < -7 Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        MsgBox "The maximum application has been reached with the dynamic brake in notch eight.", vbOKOnly + vbExclamation, "Engineer Error"
                    End If
                End If
                       
                If Val(BarFuelTank.Value) < Val(TextBoxStatusFuelTooLow.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of diesel fuel. You currently have less then " + Str$(Val(TextBoxStatusFuelTooLow.Text))
                        Let Message = Message + " US gallons left."
                        MsgBox Message, vbExclamation, "Engineer Error - Fuel Reserve"
                    End If
                End If
                
                If Val(BarSandTank.Value) < Val(textboxstatussandtoolow.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of sand. You currently have " + Str$(Val(BarSandTank.Value)) + " pounds of sand left. It is recommended that you have at" + Chr$(13)
                        Let Message = Message + "least " + Str$(Val(textboxstatussandtoolow.Text)) + " pounds of sand."
                        MsgBox Message, vbExclamation, "Engineer Error - Sand Reserve"
                    End If
                End If
                
                If Val(BarWaterTank.Value) < Val(TextBoxStatusWaterTooLow.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of water. You currently have " + Str$(Val(BarWaterTank.Value)) + " US gallons of water left. It is recommended that you have at" + Chr$(13)
                        Let Message = Message + "least " + Str$(Val(TextBoxStatusWaterTooLow.Text)) + " US gallons of water."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Reserve"
                    End If
                End If
                
                If Val(BarOilTank.Value) < Val(TextBoxStatusOilTooLow.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Please contact your dispatcher or maintenence personal and arrange for refilling" + Chr$(13)
                        Let Message = Message + "of engine lubricating oil. You currently have " + Str$(Val(BarOilTank.Value)) + " US gallons of lubricating left. It is" + Chr$(13)
                        Let Message = Message + "recommended that you have at least " + Str$(Val(TextBoxStatusOilTooLow.Text)) + " US gallons of oil. "
                        MsgBox Message, vbExclamation, "Engineer Error - Oil Reserve"
                    End If
                End If
                
                If Val(BarOilPressure.Value) < Val(TextBoxStatusOilTooLowPressure.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine oil pressure has not reached a pressure to operate the" + Chr$(13)
                        Let Message = Message + "locomotive safely. The current oil pressure is" + Str$(Val(BarOilPressure.Value)) + " pounds per square" + Chr$(13)
                        Let Message = Message + "inch. We recommend that you wait until the pressure reaches at" + Chr$(13)
                        Let Message = Message + "least" + Str$(Val(TextBoxStatusOilTooLowPressure.Text)) + " pounds per square inch."
                        MsgBox Message, vbExclamation, "Engineer Error - Oil Pressure"
                    End If
                End If
                
                If Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text) * 2 / 3 Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine oil temperature has not reached a temperature to operate the" + Chr$(13)
                        Let Message = Message + "locomotive safely. The current oil temperature is" + Str$(Val(BarOilTemperature.Value)) + " degrees. We recommend" + Chr$(13)
                        Let Message = Message + "that you wait until the temperature reaches at least" + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees."
                        MsgBox Message, vbExclamation, "Engineer Error- Oil Temperature"
                    End If
                End If
                
                If Val(BarWaterPressure.Value) < Val(TextBoxStatusWaterTooLowPressure.Text) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine water pressure has not reached a pressure to operate the " + Chr$(13)
                        Let Message = Message + "locomotive safely. The current water pressure is " + Str$(Val(BarWaterPressure.Value)) + " pounds per square inch. We" + Chr$(13)
                        Let Message = Message + "recommend that you wait until the pressure reaches" + Str$(Val(TextBoxStatusWaterTooLowPressure.Text)) + " pounds per square inch."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Pressure"
                    End If
                End If
                
                If Val(BarWaterTemperature.Value) < Val(TextBoxStatusWaterTooLowTemperature.Text) * 2 / 3 Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "The engine water temperature has not reached a temperature to operate the " + Chr$(13)
                        Let Message = Message + "locomotive safely. The current water temperature is " + Str$(Val(BarWaterTemperature.Value)) + " degrees. We" + Chr$(13)
                        Let Message = Message + "recommend that you wait until the temperature reaches" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Temperature"
                    End If
                End If
                
                If ((Val(BarOilTemperature.Value) < Val(TextBoxStatusOilTooLowTemperature.Text)) And _
                    (Val(PictureBoxThrottle.Tag) >= 1)) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Although the engine oil temperature is greater than" + Str$(Val(BarOilTemperature.Value) - 30) + "degrees, it has not reached" + Chr$(13)
                        Let Message = Message + "a temperature to operate above notch one. It is recommended that the oil temerature be greater than" + Chr$(13)
                        Let Message = Message + Str$(Val(TextBoxStatusOilTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                        MsgBox Message, vbExclamation, "Engineer Error- Oil Temperature"
                    End If
                End If
                
                If ((Val(BarWaterTemperature.Value) <= Val(TextBoxStatusWaterTooLowTemperature.Text)) And _
                    (Val(PictureBoxThrottle.Tag) >= 2)) Then
                    Let OkToNotchDown = False
                    If ButtonHelp.Caption = "&Help is On" Then
                        Let Message = "Although the engine water temperature is greater than" + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text) - 30) + " degrees, it has not reached" + Chr$(13)
                        Let Message = Message + "a temperature to operate above notch two. It is recommended that the water engine temperature be greater than" + Chr$(13)
                        Let Message = Message + Str$(Val(TextBoxStatusWaterTooLowTemperature.Text)) + " degrees before increasing the speed of the prime mover."
                        MsgBox Message, vbExclamation, "Engineer Error - Water Temperature"
                    End If
                End If
                
                'If Val(pictureboxThrottle.Tag) < 1 Then
                '    If Val(PictureBoxReverser.Tag) <> 1 Then
                '        Let OKToNotchDown = False
                '        If ButtonHelp.Caption = "&Help is On" Then
                '            Let Message = "The dynamic brake can only be engaged when the reverser handle in in nuetral."
                '            MsgBox Message, vbExclamation, "Engineer Error - Reverser Handle and Dynamic Brake",vbokonly+vbexclamation,"Engineer Error"
                '        End If
                '    End If
                'End If
            
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Start Sending Command
    '
    ' Start sending command to the digital command control unit only if the locomotive is allowed to notch down.
    '
            
                If OkToNotchDown = True Then
            
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Playing the Wave File
    '
    
                    Let PictureBoxThrottle.Tag = Val(PictureBoxThrottle.Tag) - 1
                    Let SoundThrottle.Notify = False
                    Let SoundThrottle.Wait = False
                    Let SoundThrottle.Shareable = False
                    Let SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
                    Let SoundThrottle.Command = "Open"
                    Let SoundThrottle.Notify = True
                    Let SoundThrottle.Command = "Play"
                    
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' If Decoder Equipped
    
                
                    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
                        
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Throttle Routines  0>x<7
    '
                        
                        If Val(PictureBoxThrottle.Tag) >= 0 And Val(PictureBoxThrottle.Tag) < 7 Then
                            
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
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction01234
                                        Call MainScreen.SendCommandviaTrackQ
                                        DoEvents
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction5678
                                        Call MainScreen.SendCommandviaTrackQ
                                        DoEvents
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction9101112
                                        Call MainScreen.SendCommandviaTrackQ
                                        DoEvents
                                    End If
                                ElseIf InternetSettings!Winsock.State = sckClosed Then
                                    If InternetSettings.OptionStandAlone.Value = True Then
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction01234
                                        Call MainScreen.SendCommandviaTrackQ
                                        DoEvents
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction5678
                                        Call MainScreen.SendCommandviaTrackQ
                                        DoEvents
                                        Call MainlineOperationGUI.SetSoundDecoderNumber
                                        Call MainlineOperationGUI.SetSoundControlFunction9101112
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
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                            
                        ElseIf Val(PictureBoxThrottle.Tag) = 0 Then
                            
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
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                        ElseIf Val(PictureBoxThrottle.Tag) < 0 Then
                            
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
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            ElseIf InternetSettings!Winsock.State = sckClosed Then
                                If InternetSettings.OptionStandAlone.Value = True Then
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction01234
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction5678
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                    Call MainlineOperationGUI.SetSoundDecoderNumber
                                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                                    Call MainScreen.SendCommandviaTrackQ
                                    DoEvents
                                End If
                            End If
                        End If ' If Throttle
                    End If ' If Decoder Equipped
                End If ' If OKToNotchDown
            End If ' Button
    ' ------------------------------------------------------------------------------------------------------------------------------------------
    ' Updating the Picture
    '
    ' Depending on the value in TextBoxStatusThrottle, beeing that if it is a negative value, then the dynamic brakes are being used and if
    ' a positive numver is being used, the throttle is engaged. Look at the value of the number to determine what picture to load.
    
            Let temp = App.Path
            If Val(PictureBoxThrottle.Tag) < 0 Then
                Let temp = temp + "\Graphics\Locomotive Diesel1\DynamicBrake"
            Else
                Let temp = temp + "\Graphics\Locomotive Diesel1\Throttle"
            End If
            Let temp = temp + Trim$(Abs(PictureBoxThrottle.Tag))
            Let temp = temp + "(d1).bmp"
        
            Let PictureBoxThrottle.Picture = LoadPicture(temp)
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Sending a Speed Command
    '
    ' Sending a spped command to the communication port, depending on the speed steps used. Also, if the reverser handle is in the nuetral
    ' position, then sending the commands to the digital commanf control unit is not preformed.
    '
            If Val(PictureBoxReverser.Tag) <> 1 Then
                ' Only send a speed change if the reverser handle is not in nuetral.
                If Val(PictureBoxThrottle.Tag) >= 0 Then
                    If MainlineOperationGUI!ConsistControlSpeed128 = vbChecked Then
                        Let TextBoxStatusSpeedStep.Text = Int(Val(PictureBoxThrottle.Tag) / 8 * 126)
                    Else
                        If MainlineOperationGUI!ConsistControlSpeed28 = vbChecked Then
                            Let TextBoxStatusSpeedStep.Text = Int(Val(PictureBoxThrottle.Tag) / 8 * 28)
                        Else
                            Let TextBoxStatusSpeedStep.Text = Int(Val(PictureBoxThrottle.Tag) / 8 * 14)
                        End If
                    End If
                End If
            End If
            Let TextBoxStatusThrottleDelay.Text = 1
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Setting the Deadmann Switch
    '
    ' Since the engineer has click'ed on the throttle, we know that he or she is still active with the program. We can then reset the deadmann
    ' timer to a later time.
    
            If Val(PictureBoxThrottle.Tag) <> 0 And Val(PictureBoxReverser.Tag) <> 1 Then
                Let MainlineOperationGUI!timerdeadmann.Enabled = False
                Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
                Let MainlineOperationGUI!timerdeadmann.Enabled = True
            Else ' If PictureBoxThrottle.Tag <> 0 And PictureBoxReverser.Tag <> 1 Then
                Let MainlineOperationGUI!timerdeadmann.Enabled = False
            End If ' If PictureBoxThrottle.Tag <> 0 And PictureBoxReverser.Tag <> 1 Then
            
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    '
    ' Moving the Throttle Handle Quickly
    '
    ' If the engineer moves the handle of the throttle to quickly, then a message box will display, telling the engineer that a recommened
    ' ten seconds is reqired by the locomotive builder between notches.
    
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                MsgBox "You are moving the throttle handle too quickly between otches as recommended by General Electric." + Chr$(13) + " They recommend ten seconds between movements of the throttle and/or dynamic brake.", vbOKOnly + vbExclamation, "Engineer Error"
            End If ' If ButtonHelp.Caption = "&Help is On" Then
        End If ' If Val(TextBoxStatusThrottleDelay.Text) = 0 Then
    
    
    
    ' -------------------------------------------------------------------------------------------------------------------------
    '
    ' Deadmann Switch Setting
    '
    ' In this routine will display a message on the screen to tell the engineer that the locomotive has been sutdown because of inactivity
    ' regarding the deadmann switch. REsetting the deadmann switch is require to activate the engine again.
    '

    Else ' If labeldeadmann.Visible = False Then
        If ButtonHelp.Caption = "&Help is On" Then
            Let Message = "Your locomotive was shut down becuase of the elapsed time of the deadman alarm. Please reset" & Chr$(13) & "the dead man switch before proceeding."
            MsgBox Message, vbExclamation, "Engineer Erorr - Deadmann Switch"
        End If ' If ButtonHelp.Caption = "&Help is On" Then
    End If ' If labeldeadmann.Visible = False Then
    
    ' ---------------------------------------------------------------------------------------------------------------------------
    ' Update Label with Combined Power
    '
    
    Let LabelCombinedPower.Caption = "Combined Power" + Chr$(13)
    If Val(PictureBoxReverser.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Forward - "
    ElseIf Val(PictureBoxReverser.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Reverse - "
    End If
    If Val(PictureBoxThrottle.Tag) < 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Dynamic Brake" & Chr(13)
    End If
    If Abs(PictureBoxThrottle.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Zero"
    ElseIf Abs(PictureBoxThrottle.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch One"
    ElseIf Abs(PictureBoxThrottle.Tag) = 2 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Two"
    ElseIf Abs(PictureBoxThrottle.Tag) = 3 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Three"
    ElseIf Abs(PictureBoxThrottle.Tag) = 4 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Four"
    ElseIf Abs(PictureBoxThrottle.Tag) = 5 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Five"
    ElseIf Abs(PictureBoxThrottle.Tag) = 6 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Six"
    ElseIf Abs(PictureBoxThrottle.Tag) = 7 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Seven"
    ElseIf Abs(PictureBoxThrottle.Tag) = 8 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Eight"
    End If

Return

End Sub

Private Sub SoundBell_Done(NotifyCode As Integer)

    Let SoundBell.Command = "Close"
    
End Sub

Private Sub SoundBrakeAutomatic_Done(NotifyCode As Integer)

    Let SoundBrakeAutomatic.Command = "Close"
    
End Sub

Private Sub SoundBrakeIndependent_Done(NotifyCode As Integer)

    Let SoundBrakeIndependent.Command = "Close"
    
End Sub

Private Sub SoundCompressor_Done(NotifyCode As Integer)
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Close File When Finished
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If NotifyCode = 1 Then
        SoundCompressor.Command = "Close"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Else
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Resevior
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainlineOperationGuiDiesel1Screen!BarAirMainReservoir.Value < 107 Then
            Let SoundCompressor.TimeFormat = mciFormatMilliseconds
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Repeat Middle of Sound
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If SoundCompressor.Position > 7000 Then
                'SoundCompressor.Command = "Stop"
                SoundCompressor.From = 6000
                SoundCompressor.Command = "Play"
            End If
        End If
    End If
    
End Sub

Public Sub sounddeadmann_Done(NotifyCode As Integer)

    If sounddeadmann.Length = sounddeadmann.Position Then

        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Deadmann, requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Deadmann, confirmed."
                Call InternetSettings.ButtonSend_Click
                GoSub UpdateDeadmann
                Let InternetSettings!textboxincomingdata.Text = ""
            ElseIf InternetSettings.textboxincomingdata.Text = "Deadmann, confirmed." Then
                GoSub UpdateDeadmann
                Let InternetSettings.textboxincomingdata.Text = ""
            Else
                Let InternetSettings!textboxoutboundcommand.Text = "Deadmann, requested."
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            GoSub UpdateDeadmann
        End If
    End If

    sounddeadmann.Command = "Close"

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateDeadmann:

    If MainlineOperationGuiDiesel1Screen!PictureBoxResetLeft.Tag <> "1" Then
      
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Change Picture Box to 'Button Depressed'
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp"
        Let MainlineOperationGuiDiesel1Screen!PictureBoxResetLeft.Tag = "1"
        Let MainlineOperationGuiDiesel1Screen!PictureBoxResetLeft.Picture = LoadPicture(temp)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Timer Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
        Let MainlineOperationGUI!timerdeadmann.Interval = 0
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Emergency Stop Locomotive
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetSpeed
                Call MainScreen.SendCommandviaTrackQ
            End If
        ElseIf InternetSettings.Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetSpeed
                Call MainScreen.SendCommandviaTrackQ
            End If
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display GUI Stopped
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStep = 0
        Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepNow = 0
        Let MainlineOperationGuiDiesel1Screen!TextBoxStatusSpeedStepModified = 0

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Move Throttle to Notch Zero
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
OneMore:
           
        If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
        
            If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 And Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 8 Then
            
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
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction01234
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction5678
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction9101112
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If
                ElseIf InternetSettings!Winsock.State = sckClosed Then
                    If InternetSettings.OptionStandAlone.Value = True Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction01234
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction5678
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction9101112
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
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction01234
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction5678
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction9101112
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If
                ElseIf InternetSettings!Winsock.State = sckClosed Then
                    If InternetSettings.OptionStandAlone.Value = True Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction01234
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction5678
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        Call MainlineOperationGUI.SetSoundControlFunction9101112
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If
                End If
            End If
        End If
        
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
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbnChecked
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If
                
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Picture Box Tag Property
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) > 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) - 1
    ElseIf Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag = Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) + 1
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Play Throttle Handle Sound
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let SoundThrottle.Notify = False
    Let SoundThrottle.Wait = False
    Let SoundThrottle.Shareable = False
    Let SoundThrottle.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
    Let SoundThrottle.Command = "Open"
    Let SoundThrottle.Notify = True
    Let SoundThrottle.Command = "Play"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update GUI display in Comupter Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) < 0 Then
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake"
    Else
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle"
    End If
    Let temp = temp & Trim$(Abs(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag)) & "(d1).bmp"
    Let MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Picture = LoadPicture(temp)
    For tt = 1 To 75000
        DoEvents
    Next tt
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Computer Screen in GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelCombinedPower.Caption = "Combined Power" + Chr$(13)
    If Val(PictureBoxReverser.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Forward - "
    ElseIf Val(PictureBoxReverser.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Neutral - "
    Else
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Reverse - "
    End If
    If Val(PictureBoxThrottle.Tag) < 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Dynamic Brake" + Chr(13)
    End If
    If Abs(PictureBoxThrottle.Tag) = 0 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Zero"
    ElseIf Abs(PictureBoxThrottle.Tag) = 1 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch One"
    ElseIf Abs(PictureBoxThrottle.Tag) = 2 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Two"
    ElseIf Abs(PictureBoxThrottle.Tag) = 3 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Three"
    ElseIf Abs(PictureBoxThrottle.Tag) = 4 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Four"
    ElseIf Abs(PictureBoxThrottle.Tag) = 5 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Five"
    ElseIf Abs(PictureBoxThrottle.Tag) = 6 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Six"
    ElseIf Abs(PictureBoxThrottle.Tag) = 7 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Seven"
    ElseIf Abs(PictureBoxThrottle.Tag) = 8 Then
        Let LabelCombinedPower.Caption = LabelCombinedPower.Caption & "Notch Eight"
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check and Loop if Not Notch Zero
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(MainlineOperationGuiDiesel1Screen!PictureBoxThrottle.Tag) <> 0 Then GoTo OneMore
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Dynamic Brake Sound Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
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
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbnChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbnChecked
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        End If
    End If
    
Return
    
End Sub

Private Sub SoundHorn_Done(NotifyCode As Integer)

    Let SoundHorn.Command = "Close"
    
End Sub

Private Sub SoundLeftScreen_Done(NotifyCode As Integer)

    Let SoundLeftScreen.Command = "Close"

End Sub

Private Sub SoundLight_Done(NotifyCode As Integer)

    SoundLight.Command = "Close"
    
    
End Sub

Private Sub SoundRadioPhone_Done(NotifyCode As Integer)

    Let SoundRadioPhone.Command = "Close"
    
End Sub

Private Sub SoundResetLeft_Done(NotifyCode As Integer)

    SoundResetLeft.Command = "Close"
    
End Sub


Private Sub SoundResetRight_Done(NotifyCode As Integer)

    Let SoundResetRight.Command = "Close"

End Sub

Private Sub SoundReverser_Done(NotifyCode As Integer)

    Let SoundReverser.Command = "Close"
    
End Sub

Private Sub SoundRightScreen_Done(NotifyCode As Integer)

    SoundRightScreen.Command = "Close"
    
End Sub

Private Sub SoundSand_Done(NotifyCode As Integer)

    Let SoundSand.Command = "Close"

End Sub

Private Sub SoundThrottle_Done(NotifyCode As Integer)

    Let SoundThrottle.Command = "Close"
    
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
    PictureBoxSpeedometer.DrawWidth = 6
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, , (320 / 57.3), (234 / 57.3), 1
    PictureBoxSpeedometer.DrawWidth = 5
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, &HC0&, (startpoint / 57.3), (EndPoint / 57.3), 1
Else
    PictureBoxSpeedometer.DrawWidth = 6
    PictureBoxSpeedometer.Circle (75 * 15, 73 * 15), 66 * 15, , (320 / 57.3), (234 / 57.3), 1
End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Change the Interval of Timer for Track Joint Sounds
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'    If Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text) = 0 Then
'        Let MainlineOperationGuiDiesel1Screen!TimerTrackJoint.Interval = 0
'    Else
'        Let MainlineOperationGuiDiesel1Screen!TimerTrackJoint.Interval = (126 - Val(MainlineOperationGuiDiesel1Screen!TextBoxDigitalSpeed.Text)) * 12
'    End If
 
End Sub


Private Sub TextBoxStatusDynamicBrakeApplication_Change()
    
    Let MainlineOperationGUI!TimerSendCommand.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingsendcommandInterval.Text) * Val(MainlineOperationGUI!textBoxTonnage.Text) / Val(TextBoxStatusDynamicBrakeApplication.Text / 3 + 1)
    Let MainlineOperationGUI!TimerSendCommand.Enabled = True
    Let MainlineOperationGUI!TimerSpeedChange.Interval = Val(MainlineOperationGuiDiesel1Screen!TextBoxSettingSpeedChangeInterval.Text) * Val(MainlineOperationGUI!textBoxTonnage.Text) / Val(TextBoxStatusDynamicBrakeApplication.Text / 3 + 1)
    Let MainlineOperationGUI!TimerSpeedChange.Enabled = True
    
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
    Let Temporary = temporary1
Else
    Let Temporary = temporary2
End If

If Temporary > 1 Then
    Let Temporary = 1
End If

Let TextBoxStatusSpeedStepModified.Text = Int(Val(TextBoxStatusSpeedStepNow.Text) * Temporary)

End Sub



















Private Sub TextBoxVirtualCabSignalDistance_Change()

Let LabelvirtualcabSignalDistanceFeet.Caption = Int(TextBoxVirtualCabSignalDistance.Text)

End Sub










Private Sub TimerAdvertising_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerAdvertising, Timer" & vbCrLf
    End If 'Darrin
        
    Let TimerAdvertising.Enabled = False

    If InternetSettings.optionclient.Value = True Then
        If InternetSettings.NetConnect.Connected = True Then
            If MainlineOperationGuiDiesel1Screen.Visible = True Then
                Let Ini.Filename = App.Path$ & "\Atc.ini"
                Let Ini.Application = "Mainline Operation GUI Diesel1 Screen"
                Let Ini.Parameter = "AdvertisingCounter"
            
                Dim TemporaryAdvertisingCounter As Long
            
                Let TemporaryAdvertisingCounter = Val(Ini.Value)
                Let TemporaryAdvertisingCounter = TemporaryAdvertisingCounter + 1
                Let Ini.Value = TemporaryAdvertisingCounter
                On Error GoTo TIMEOUT
                FtpControl.Connect
                If FtpControl.FileExists("/ad" & Str$(TemporaryAdvertisingCounter) & ".bmp") = 1 Then
                    FtpControl.Get "/ad" & Str$(TemporaryAdvertisingCounter) & ".bmp", App.Path$ & "\Advertising\ad" & Str$(TemporaryAdvertisingCounter) & ".bmp", 0
                    Let PictureAdvertising.Picture = LoadPicture(App.Path$ & "\Advertising\ad" & Str$(TemporaryAdvertisingCounter) & ".bmp")
                    PictureBoxComputerScreenRight_Click (8)
                    Let Ini.Value = TemporaryAdvertisingCounter
                Else
                    Let Ini.Parameter = "AdvertisingCounter"
                    Let Ini.Value = 0
                End If
                FtpControl.Disconnect
            End If
       End If
    End If
    Let TimerAdvertising.Enabled = True
                
TIMEOUT:
    On Error GoTo 0
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerAdvertising, Timer" & vbCrLf
    End If 'Darrin
    
End Sub


Private Sub TimerCheckWinsock_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerCheckWinsock, Timer" & vbCrLf 'Darrin
    End If 'Darrin

    Let InternetSettings!textboxoutboundcommand.Text = "Checking Winsock Connection."
    Call InternetSettings.ButtonSend_Click
    Let InternetSettings!textboxincomingdata.Text = ""

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerCheckWinsock, Timer" & vbCrLf
    End If 'Darrin

End Sub



Private Sub TimerStartEngine1_Timer()
 
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine1, Timer" & vbCrLf
    End If ' Debug Tag

    Let TimerStartEngine1.Interval = 0
    Let TimerStartEngine1.Enabled = False

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Stop the Locomotive

' If the user selects a locomotive that is currently in use on the layout, and actively moving, then this program will stop the consist.
' This is done by sending a speed command, one, or emergency stop, to the consist.

    Let MainlineOperationGUI!ConsistControlSpeed.Value = "1"
    
    If InternetSettings.Winsock.State = sckConnected Then
        If InternetSettings.optionHost = True Then
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
        End If
    ElseIf InternetSettings.Winsock.State = sckClosed Then
        If InternetSettings.OptionStandAlone.Value = True Then
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Setup the Resources of the Locomotive
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'If InternetSettings.Winsock.State = sckConnected Then
        'If InternetSettings.OptionHost = True Then
            'Call ButtonFillFuel_Click
            'Call ButtonFillOil_Click
            'Call ButtonFillSand_Click
            'Call ButtonFillWater_Click
        'End If
    'ElseIf InternetSettings.Winsock.State = sckClosed Then
        Randomize
        Let TextBoxFuelTank.Text = Int((Val(BarFuelTank.Max) - Val(BarFuelTank.Min) + 1) * Rnd + Val(BarFuelTank.Min))
        Let TextBoxOilTank.Text = Int((Val(BarOilTank.Max) - Val(BarOilTank.Min) + 1) * Rnd + Val(BarOilTank.Min))
        Let TextBoxWaterTank.Text = Int((Val(BarWaterTank.Max) - Val(BarWaterTank.Min) + 1) * Rnd + Val(BarWaterTank.Min))
        Let TextBoxSandTank.Text = Int((Val(BarSandTank.Max) - Val(BarSandTank.Min) + 1) * Rnd + Val(BarSandTank.Min))
    'End If

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Cab Signals
'
' If Cab Signals are used, update then with pictures of the signals.

        If MainlineOperationGUI!OptionVirtualcabSignalYes.Value = True Then
            Let MainlineOperationGUI!TimerVirtualCabSignal.Interval = 1000
            Let MainlineOperationGUI!TimerVirtualCabSignal.Enabled = True
            Let PictureBoxVirtualCabSignal.Visible = True
            Let textboxvirtualCabsignalstatuscurrent = "Clear"
            Let PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
            Let PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let TextBoxVirtualCabSignalStatusApproaching = "Clear"
            Let PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
            Let PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            Let PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
        Else
            Let PictureBoxVirtualCabSignal.Visible = False
        End If
 
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine1, Timer" & vbCrLf
    End If 'Darrin

End Sub


Private Sub TimerStartEngine2_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine2, Timer" & vbCrLf
    End If 'Darrin

    Let TimerStartEngine2.Interval = 0
    Let TimerStartEngine2.Enabled = False

    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' If Sound Equipped
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' If DSX Decoder
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If MainlineOperationGUI!CheckboxSoundtraxxDSX.Value = vbChecked Then
        
            '------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Start Engine Sound
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
            End If
                                                        
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
            
            If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
                Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
                Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
                Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
                Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
                Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
            End If
            
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine2, Timer" & vbCrLf
    End If 'Darrin

End Sub


Private Sub TimerStartEngine3_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine3, Timer" & vbCrLf
    End If 'Darrin

Let TimerStartEngine3.Interval = 0
Let TimerStartEngine3.Enabled = False

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start the timers
'
' Timers are the core of the program. They control many function on a continuos basis.

Let TempTonnage = Val(ScrollTimeAdjustment) / 100 + 1

Let MainlineOperationGUI!TimerSendCommand.Interval = Val(TextBoxSettingsendcommandInterval.Text) * TempTonnage / ((Val(TextBoxStatusDynamicBrakeApplication.Text) / 3) + 1)
Let MainlineOperationGUI!TimerSendCommand.Enabled = True
Let MainlineOperationGUI!TimerSpeedChange.Interval = Val(TextBoxSettingSpeedChangeInterval.Text) * TempTonnage / ((Val(TextBoxStatusDynamicBrakeApplication.Text) / 3) + 1)
Let MainlineOperationGUI!TimerSpeedChange.Enabled = True
Let MainlineOperationGUI!timerairmainReservoir.Interval = 2000 * TempTonnage
Let MainlineOperationGUI!timerairmainReservoir.Enabled = True
Let MainlineOperationGUI!TimerAirAutomaticBrake.Interval = Val(TextboxSettingAutomaticAirBrake.Text) * TempTonnage
Let MainlineOperationGUI!TimerAirAutomaticBrake.Enabled = True
Let MainlineOperationGUI!TimerAirIndependentBrake.Interval = Val(TextboxSettingIndependentAirBrake.Text) * TempTonnage
Let MainlineOperationGUI!TimerAirIndependentBrake.Enabled = True
Let MainlineOperationGUI!timerthrottledelay.Interval = 1000
Let MainlineOperationGUI!timerthrottledelay.Enabled = True
Let MainlineOperationGUI!TimerOilPressure.Interval = 350 * TempTonnage
Let MainlineOperationGUI!TimerOilPressure.Enabled = True
Let MainlineOperationGUI!TimerOilTemperature.Interval = 550 * TempTonnage
Let MainlineOperationGUI!TimerOilTemperature.Enabled = True
Let MainlineOperationGUI!TimerWaterPressure.Interval = 350 * TempTonnage
Let MainlineOperationGUI!TimerWaterPressure.Enabled = True
Let MainlineOperationGUI!TimerWaterTemperature.Interval = 550 * TempTonnage
Let MainlineOperationGUI!TimerWaterTemperature.Enabled = True
Let MainlineOperationGUI!timerfuelconsumption.Interval = 1000 * TempTonnage
Let MainlineOperationGUI!timerfuelconsumption.Enabled = True
Let MainlineOperationGUI!TimerOilConsumption.Interval = 3000 * TempTonnage
Let MainlineOperationGUI!TimerOilConsumption.Enabled = True
Let MainlineOperationGUI!TimerWaterConsumption.Interval = 3000 * TempTonnage
Let MainlineOperationGUI!TimerWaterConsumption.Enabled = True
Let MainlineOperationGUI!TimerRadiatorFans.Interval = 2000 * TempTonnage
Let MainlineOperationGUI!TimerRadiatorFans.Enabled = True
Let MainlineOperationGUI!TimerRPM.Interval = 125 * TempTonnage
Let MainlineOperationGUI!TimerRPM.Enabled = True

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerStartEngine3, Timer" & vbCrLf
    End If 'Darrin

End Sub







Private Sub TimerTrackJoint_Timer()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel1 Screen, TimerTrackJoint, Timer" & vbCrLf
    End If 'Darrin
    
    Dim TemporayValue As String
    Let TemporaryValue = CStr(Int((23 * Rnd) + 1))
    Let Soundtrackjoint(TemporaryValue).Notify = False
    Let Soundtrackjoint(TemporaryValue).Wait = False
    Let Soundtrackjoint(TemporaryValue).Shareable = False
    Let Soundtrackjoint(TemporaryValue).Command = "Close"
    Let Soundtrackjoint(TemporaryValue).Filename = App.Path$ & "\Sounds\Graphics\TrackJoint" & TemporaryValue & ".wav"
    Let Soundtrackjoint(TemporaryValue).Command = "Open"
    Let Soundtrackjoint(TemporaryValue).Command = "Play"

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending,  Mainline Operation GUI Diesel1 Screen, TimerTrackJoint, Timer" & vbCrLf
    End If 'Darrin

End Sub





Public Sub TransPictureAuxillarySwitch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Auxilliary Switches
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '   Index was the only variable passed over internet
    '
    If InternetSettings!Winsock.State = sckConnected Then
    
        If InternetSettings!textboxincomingdata.Text = "Auxilliary Switch " & CStr(Index) & " , left button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Auxilliary Switch " & CStr(Index) & " , left button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbLeftButton
            GoSub UpdateAuxilliarySwitch
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Auxilliary Switch " & CStr(Index) & " , left button confirmed." Then
            Let Button = vbLeftButton
            GoSub UpdateAuxilliarySwitch
            Let InternetSettings.textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Auxilliary Switch " & CStr(Index) & " , right button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Auxilliary Switch " & CStr(Index) & " , right button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbRightButton
            GoSub UpdateAuxilliarySwitch
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Auxilliary Switch " & CStr(Index) & " , right button confirmed." Then
            Let Button = vbRightButton
            GoSub UpdateAuxilliarySwitch
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            If Button = vbLeftButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Auxilliary Switch " & CStr(Index) & " , left button requested."
            ElseIf Button = vbRightButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Auxilliary Switch " & CStr(Index) & " , right button requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateAuxilliarySwitch
    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateAuxilliarySwitch:

    Let sounddeadmann.Notify = False
    Let sounddeadmann.Wait = False
    Let sounddeadmann.Shareable = False
    Let sounddeadmann.Command = "Stop"
    'Let sounddeadmann.Notify = False
    'Let sounddeadmann.Command = "Close"
    
    If TransPictureAuxillarySwitch(Index).Tag = "0" Then
        Let TransPictureAuxillarySwitch(Index).Tag = "1"
        Let TransPictureAuxillarySwitch(Index).Picture = LoadPicture(App.Path & "\Graphics\Locomotive Diesel1\Switch1(d1).bmp")
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Is Decoder Equiped
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Host Mode - Left Button - Turn it On
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    If Button = vbLeftButton Then
                        Call MainlineOperationGUI.SetLocomotiveNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.ConsistControlFunction0.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.ConsistControlFunction1.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.ConsistControlFunction3.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.ConsistControlFunction4.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.ConsistControlFunction5.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.ConsistControlFunction6.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.ConsistControlFunction7.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.ConsistControlFunction8.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.ConsistControlFunction9.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.ConsistControlFunction10.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.ConsistControlFunction11.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.ConsistControlFunction12.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Host Mode - Right Button - Turn it On
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ElseIf Button = vbRightButton Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.SoundControlFunction0.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.SoundControlFunction1.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.SoundControlFunction2.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.SoundControlFunction3.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.SoundControlFunction4.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.SoundControlFunction5.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.SoundControlFunction6.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.SoundControlFunction7.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.SoundControlFunction8.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.SoundControlFunction9.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.SoundControlFunction10.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.SoundControlFunction11.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.SoundControlFunction12.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If ' right or left
                End If ' host mode
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Stand Alone - Left Button Turn it On
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    If Button = vbLeftButton Then
                        Call MainlineOperationGUI.SetLocomotiveNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.ConsistControlFunction0.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.ConsistControlFunction1.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.ConsistControlFunction2.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.ConsistControlFunction3.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.ConsistControlFunction4.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.ConsistControlFunction5.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.ConsistControlFunction6.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.ConsistControlFunction7.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.ConsistControlFunction8.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.ConsistControlFunction9.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.ConsistControlFunction10.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.ConsistControlFunction11.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.ConsistControlFunction12.Value = vbChecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Stand ALone - Right Button - Turn it On
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ElseIf Button = vbRightButton Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.SoundControlFunction0.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.SoundControlFunction1.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.SoundControlFunction2.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.SoundControlFunction3.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.SoundControlFunction4.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.SoundControlFunction5.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.SoundControlFunction6.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.SoundControlFunction7.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.SoundControlFunction8.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.SoundControlFunction9.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.SoundControlFunction10.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.SoundControlFunction11.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.SoundControlFunction12.Value = vbChecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If ' right or left
                End If ' stand alone?
            End If ' if wincok open closed
        End If ' sound decoder equipped
    Else 'if TransPictureAuxillarySwitch(index).tag = "1"
        Let TransPictureAuxillarySwitch(Index).Tag = "0"
        Let TransPictureAuxillarySwitch(Index).Picture = LoadPicture(App.Path & "\Graphics\Locomotive Diesel1\Switch0(d1).bmp")
        If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings.optionHost.Value = True Then
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Host Mode - Left Button - Turn it Off
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    If Button = vbLeftButton Then
                        Call MainlineOperationGUI.SetLocomotiveNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.ConsistControlFunction0.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.ConsistControlFunction1.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.ConsistControlFunction2.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.ConsistControlFunction3.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.ConsistControlFunction7.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.ConsistControlFunction8.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.ConsistControlFunction9.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.ConsistControlFunction10.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.ConsistControlFunction11.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.ConsistControlFunction12.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Stand Alone - Right Button - Turn in Off
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ElseIf Button = vbRightButton Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If ' right or left
                End If ' if host
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Stand Alone Left Button - Turn it Off
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    If Button = vbLeftButton Then
                        Call MainlineOperationGUI.SetLocomotiveNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.ConsistControlFunction0.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.ConsistControlFunction1.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.ConsistControlFunction2.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.ConsistControlFunction3.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction01234
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.ConsistControlFunction7.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.ConsistControlFunction8.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.ConsistControlFunction9.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.ConsistControlFunction10.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.ConsistControlFunction11.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.ConsistControlFunction12.Value = vbUnchecked
                            Call MainlineOperationGUI.SetConsistFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Stand Alone - Right Button - Turn in Off
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
                    ElseIf Button = vbRightButton Then
                        Call MainlineOperationGUI.SetSoundDecoderNumber
                        If Index = 0 Then
                            Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 1 Then
                            Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 2 Then
                            Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 3 Then
                            Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction01234
                        ElseIf Index = 4 Then
                            Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 5 Then
                            Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 6 Then
                            Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 7 Then
                            Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 8 Then
                            Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction5678
                        ElseIf Index = 9 Then
                            Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 10 Then
                            Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 11 Then
                            Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        ElseIf Index = 12 Then
                            Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
                            Call MainlineOperationGUI.SetSoundControlFunction9101112
                        End If
                        Call MainScreen.SendCommandviaTrackQ
                        DoEvents
                    End If ' right or left
                End If ' stand alone mode
            End If ' winsock open or closed
        End If ' sound equipped decoder
    End If ' tage = 1 or tag = 0
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let SoundResetLeft.Notify = False
    Let SoundResetLeft.Wait = False
    Let SoundResetLeft.Shareable = False
    Let SoundResetLeft.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let SoundResetLeft.Command = "Open"
    Let SoundResetLeft.Notify = True
    Let SoundResetLeft.Command = "Play"
    

    Call BalloonHelpUpdatePart06
    
Return

End Sub


Public Sub TransPictureBoxIndependentBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ButtonEngineStart.Caption <> "&Start Engine" Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Independent Brake, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateIndependentBrake
                Let InternetSettings!textboxincomingdata.Text = ""
            
            ElseIf InternetSettings.textboxincomingdata.Text = "Independent Brake, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateIndependentBrake
                Let InternetSettings.textboxincomingdata.Text = ""
            
            ElseIf InternetSettings!textboxincomingdata.Text = "Independent Brake, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateIndependentBrake
                Let InternetSettings!textboxincomingdata.Text = ""
            
            ElseIf InternetSettings.textboxincomingdata.Text = "Independent Brake, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateIndependentBrake
                Let InternetSettings.textboxincomingdata.Text = ""
            
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateIndependentBrake
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Independent Brake handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "The application of the independent brake is at " & Str$(Int(100 - (Val(TextBoxStatusIndependentBrakePressure.Text) * 100))) & " percent"
        Let BalloonHelpText2 = "Automatic Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TranspictureboxIndependentBrake)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBoxIndependentBrake, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Independent Brake, unable to setup balloon help for 'PictureBoxIndependentBrake' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateIndependentBrake:

    If Button = vbLeftButton Then
        If Val(TransPictureBoxIndependentBrake.Tag) < 21 Then
            Let TransPictureBoxIndependentBrake.Tag = Val(TransPictureBoxIndependentBrake.Tag) + 1
            If Val(TransPictureBoxIndependentBrake.Tag) = 1 Then
                Let SoundBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentOpen.wav"
            Else
                Let SoundBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let SoundBrakeIndependent.Notify = False
            Let SoundBrakeIndependent.Wait = False
            Let SoundBrakeIndependent.Shareable = False
            Let SoundBrakeIndependent.Command = "Open"
            Let SoundBrakeIndependent.Notify = True
            Let SoundBrakeIndependent.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum braking, zero (100) percent for the independent brakes (locomotive brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbRightButton Then
        If Val(TransPictureBoxIndependentBrake.Tag) > 0 Then
            Let TransPictureBoxIndependentBrake.Tag = Val(TransPictureBoxIndependentBrake.Tag) - 1
            If Val(TransPictureBoxIndependentBrake.Tag) = 0 Then
                Let SoundBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentClose.wav"
            Else
                Let SoundBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let SoundBrakeIndependent.Notify = False
            Let SoundBrakeIndependent.Wait = False
            Let SoundBrakeIndependent.Shareable = False
            Let SoundBrakeIndependent.Command = "Open"
            Let SoundBrakeIndependent.Notify = True
            Let SoundBrakeIndependent.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum braking, zero (0) percent for the independent brakes (locomotive brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    Let temp = App.Path$
    Let temp = temp + "\Graphics\Locomotive Diesel1\IndependentBrake"
    Let temp = temp + Trim(TransPictureBoxIndependentBrake.Tag)
    Let temp = temp + "(d1).bmp"
    Let TransPictureBoxIndependentBrake.Picture = LoadPicture(temp)
    
    If Val(PictureBoxThrottle.Tag) <> 0 And Val(PictureBoxReverser.Tag) <> 1 Then
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
        Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
        Let MainlineOperationGUI!timerdeadmann.Enabled = True
    Else
        Let MainlineOperationGUI!timerdeadmann.Enabled = False
    End If

Return
End Sub


Public Sub TransPictureBoxResetRight_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Reset Right, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Reset Right, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateResetRight
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Reset Right, confirmed." Then
            GoSub UpdateResetRight
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Reset Right, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateResetRight
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
'        Dim BalloonHelpText1 As String
'        Dim BalloonHelpText2 As String
'        Dim BalloonHelpSetup As Long
'        Dim BalloonHelpFont As New StdFont
'        Dim BalloonHelpVisibleTime As Long
'        Dim BalloonHelpTimeDelay As Long
'        Dim BalloonHelpShadow As Boolean
'        Dim BalloonHelpCenter As Boolean
'        Dim BalloonHelpShowOnDemand As Boolean
'        Dim BalloonHelpOpacity As Byte
'        Dim BalloonHelpWaveFile As String
'
'        Let Ini.Filename = App.Path$ & "\Atc.ini"
'        Let Ini.Application = "All Screens"
'        Let Ini.Parameter = "BalloonHelpFontName"
'        Let BalloonHelpFont.Name = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontSize"
'        Let BalloonHelpFont.Size = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontBold"
'        Let BalloonHelpFont.Bold = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontItalic"
'        Let BalloonHelpFont.Italic = Ini.Value
'        Let Ini.Parameter = "BalloonHelpFontUnderline"
'        Let BalloonHelpFont.Underline = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour1"
'        Let BalloonHelpColour1 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour2"
'        Let BalloonHelpColour2 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpColour3"
'        Let BalloonHelpColour3 = Ini.Value
'        Let Ini.Parameter = "BalloonHelpVisibleTime"
'        Let BalloonHelpVisibleTime = Ini.Value
'        Let Ini.Parameter = "BalloonHelpDelayTime"
'        Let BalloonHelpDelayTime = Ini.Value
'        Let Ini.Parameter = "BalloonHelpShadow"
'        Let BalloonHelpShadow = Ini.Value
'        Let Ini.Parameter = "BalloonHelpCenter"
'        Let BalloonHelpCenter = Ini.Value
'        Let Ini.Parameter = "BalloonHelpShowOnDemand"
'        Let BalloonHelpShowOnDemand = Ini.Value
'        Let Ini.Parameter = "BalloonHelpWaveFile"
'       Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
'       Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
'        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
'            BalloonHelpOpacity = 255
'        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
'            Let Ini.Parameter = "BalloonHelpOpacity"
'            Let BalloonHelpOpacity = Ini.Value
'        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
'                Let balloonhelp.Speech = False
'        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
'            Let balloonhelp.Speech = True
'            Let balloonhelp.Voice = 0
'            Let BalloonHelpWaveFile = ""
'        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
'
'        Let BalloonHelpText1 = "Press this button for emergency application of the digial command control. It will stop all trains on all directions."
'        Let BalloonHelpText2 = "Emergency Stop"
'
'        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxResetRight)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxResetRight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, Balloonhelpwavefile)
'        If BalloonHelpSetup = 0 Then
'            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
'                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
'                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
'                Let Ini.Filename = App.Path$ & "\Atc.log"
'                Let Ini.Application = "Log Errors"
'                Let Ini.Parameter = Date$ & " " & Time$
'                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Reset Right, unable to setup balloon help for 'PictureBoxResetRight' control."
'            End If
'        End If
'
'        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault
'
'    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateResetRight:

    Let TransPictureBoxResetRight.Tag = "1"
    Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\ResetRight1(d1).bmp"
    Let TransPictureBoxResetRight.Picture = LoadPicture(temp)

    Let MainlineOperationGUI!timerdeadmann.Enabled = False
    Let MainlineOperationGUI!timerdeadmann.Interval = 0

    Let MainlineOperationGUI.ConsistControlSpeed.Value = "1"
    '   Set Speed to Emergency Stop
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings.optionHost.Value = True Then
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
            DoEvents
        End If
    ElseIf InternetSettings!Winsock.State = sckClosed Then
        If InternetSettings.OptionStandAlone.Value = True Then
            Call MainlineOperationGUI.SetLocomotiveNumber
            Call MainlineOperationGUI.SetSpeed
            Call MainScreen.SendCommandviaTrackQ
            DoEvents
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Stop GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let TextBoxStatusSpeedStep = 0
    Let TextBoxStatusSpeedStepNow = 0
    Let TextBoxStatusSpeedStepModified = 0
         
OneMore:
           
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
        If Val(PictureBoxThrottle.Tag) > 0 And Val(PictureBoxThrottle.Tag) < 8 Then
    
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If

        ElseIf Val(PictureBoxThrottle.Tag) < 0 Then
    
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbnChecked
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If

    If Val(PictureBoxThrottle.Tag) > 0 Then
        Let PictureBoxThrottle.Tag = Val(PictureBoxThrottle.Tag) - 1
    ElseIf Val(PictureBoxThrottle.Tag) < 0 Then
        Let PictureBoxThrottle.Tag = Val(PictureBoxThrottle.Tag) + 1
    End If
    
    Let SoundResetRight.Notify = False
    Let SoundResetRight.Wait = False
    Let SoundResetRight.Shareable = False
    Let SoundResetRight.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
    Let SoundResetRight.Command = "Open"
    Let SoundResetRight.Command = "Play"
                      
    
    If Val(PictureBoxThrottle.Tag) < 0 Then
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake"
    Else
        Let temp = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle"
    End If
    Let temp = temp + Trim$(Abs(PictureBoxThrottle.Tag))
    Let temp = temp + "(d1).bmp"

    Let PictureBoxThrottle.Picture = LoadPicture(temp)

    For tt = 1 To 150000
        DoEvents
    Next tt

    If Val(PictureBoxThrottle.Tag) <> 0 Then GoTo OneMore
            
    If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
        If Val(PictureBoxThrottle.Tag) = 0 Then
            If MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "0" Then
                Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "1" Then
                Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "2" Then
                Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "3" Then
                Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "4" Then
                Let MainlineOperationGUI!SoundControlFunction4.Value = vbnChecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "5" Then
                Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "6" Then
                Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
            ElseIf MainlineOperationGUI!TextBoxMappedFunctionDynamicBrake.Text = "7" Then
                Let MainlineOperationGUI!SoundControlFunction7.Value = vbnChecked
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
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            ElseIf InternetSettings!Winsock.State = sckClosed Then
                If InternetSettings.OptionStandAlone.Value = True Then
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction01234
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction5678
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                    Call MainlineOperationGUI.SetSoundDecoderNumber
                    Call MainlineOperationGUI.SetSoundControlFunction9101112
                    Call MainScreen.SendCommandviaTrackQ
                    DoEvents
                End If
            End If
        End If
    End If
    
    Let TransPictureBoxResetRight.Tag = "0"
    
    Let Temporary$ = App.Path$ & "\Graphics\Locomotive Diesel1\ResetRight0(d1).bmp"
    Let TransPictureBoxResetRight.Picture = LoadPicture(Temporary$)

Return

End Sub

Private Sub TransPictureBoxResetRight_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("A") Then TransPictureBoxResetRight.Left = TransPictureBoxResetRight.Left - 1
    If KeyAscii = Asc("W") Then TransPictureBoxResetRight.Top = TransPictureBoxResetRight.Top - 1
    If KeyAscii = Asc("S") Then TransPictureBoxResetRight.Left = TransPictureBoxResetRight.Left + 1
    If KeyAscii = Asc("Z") Then TransPictureBoxResetRight.Left = TransPictureBoxResetRight.Top + 1

End Sub


Private Sub TransPictureButtonSide_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If TransPictureButtonSide(Index).Tag = "0" Then
        Let TransPictureButtonSide(Index).Tag = "1"
        Let TransPictureButtonSide(Index).Picture = LoadPicture(App.Path & "\Graphics\Locomotive Diesel1\ButtonSide1(d1).bmp")
        Let MainlineOperationGuiDiesel1Screen.WindowState = 1
        Let TransPictureButtonSide(Index).Tag = "0"
        Let TransPictureButtonSide(Index).Picture = LoadPicture(App.Path & "\Graphics\Locomotive Diesel1\ButtonSide0(d1).bmp")
    Else 'if transpictureButtonSide(index).tag = "1"
        Let TransPictureButtonSide(Index).Tag = "0"
        Let TransPictureButtonSide(Index).Picture = LoadPicture(App.Path & "\Graphics\Locomotive Diesel1\ButtonSide0(d1).bmp")
    End If
End Sub




Private Sub VideoCapture_ConnectionClosed(ByVal RemoteAddress As String, ByVal ErrorCode As Long)

    Call VideoModule.ConnectionClosed(RemoteAddress, ErrorCode)

End Sub

Private Sub VideoCapture_ConnectionRequest(ByVal FromAddress As String, Cancel As Long)

    Call VideoModule.ConnectionRequest(FromAddress, Cancel)
  
End Sub


Private Sub VideoCapture_DeviceLost()
   
    Call VideoModule.DeviceLost
    
End Sub

Private Sub VideoCaptureNotes_Change()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Video Notes
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let VideoCaptureNotes.Text = Right$(VideoCaptureNotes.Text, 32000)
    Let VideoCaptureNotes.SelStart = Len(VideoCaptureNotes.Text)

End Sub


Private Sub BalloonHelpUpdatePart01()
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   RadioPhone
'   ResetButtonLeft (Alerter)
'   Reverser
'   Throttle
'   Horn
'   Bell
'   Sander
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Radio Phone
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If Val(PictureBoxRadioPhone.Tag) = 0 Then
            Let BalloonHelpText1 = "The radio phone is off the hook (in use)."
        Else 'If Val(PictureBoxRadioPhone.Tag) = 1 Then
            Let BalloonHelpText1 = "The radio phone is on the hook (not in use)."
        End If
        Let BalloonHelpText2 = "Radio Phone"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxRadioPhone)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxRadioPhone, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Radio Phone, unable to setup balloon help for 'PictureBoxRadioPhone' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Alerter
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText1 = "The deadmann switch is used to reset the alerter."
        Let BalloonHelpText2 = "Deadmann Switch"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxresetleft)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxResetLeft, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Radio Phone, unable to setup balloon help for 'PictureBoxRadioPhone' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Reverser
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        If PictureBoxReverser.Tag = 0 Then
            Let BalloonHelpText1 = "This reverser handle is in the forward position."
        ElseIf PictureBoxReverser.Tag = 1 Then
            Let BalloonHelpText1 = "This reverser handle is in the neutral position."
        Else 'If PictureBoxReverser.Tag = 2 Then
            Let BalloonHelpText1 = "This reverser handle is in the reverse position."
        End If
        Let BalloonHelpText2 = "Reverser Handle"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxreverser)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxReverser, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Reverser, unable to setup balloon help for 'PictureReverser' control."
            End If
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Throttle
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If PictureBoxThrottle.Tag = -8 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position eight."
        ElseIf PictureBoxThrottle.Tag = -7 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position seven."
        ElseIf PictureBoxThrottle.Tag = -6 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position six."
        ElseIf PictureBoxThrottle.Tag = -5 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position five."
        ElseIf PictureBoxThrottle.Tag = -4 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position four."
        ElseIf PictureBoxThrottle.Tag = -3 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position three."
        ElseIf PictureBoxThrottle.Tag = -2 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position two."
        ElseIf PictureBoxThrottle.Tag = -1 Then
            Let BalloonHelpText1 = "The throttle is in dynamic brake position one."
        ElseIf PictureBoxThrottle.Tag = 0 Then
            Let BalloonHelpText1 = "The throttle is in the nuetral position."
        ElseIf PictureBoxThrottle.Tag = 1 Then
            Let BalloonHelpText1 = "The throttle is in first notch position."
        ElseIf PictureBoxThrottle.Tag = 2 Then
            Let BalloonHelpText1 = "The throttle is in second notch position."
        ElseIf PictureBoxThrottle.Tag = 3 Then
            Let BalloonHelpText1 = "The throttle is in third notch position."
        ElseIf PictureBoxThrottle.Tag = 4 Then
            Let BalloonHelpText1 = "The throttle is in fourth notch position."
        ElseIf PictureBoxThrottle.Tag = 5 Then
            Let BalloonHelpText1 = "The throttle is in fifth notch position."
        ElseIf PictureBoxThrottle.Tag = 6 Then
            Let BalloonHelpText1 = "The throttle is in sixth notch position."
        ElseIf PictureBoxThrottle.Tag = 7 Then
            Let BalloonHelpText1 = "The throttle is in seventh notch position."
        Else 'If PictureBoxThrottle.Tag = 8 Then
            Let BalloonHelpText1 = "The throttle is in eighth notch position."
        End If
        
        Let BalloonHelpText2 = "Throttle Handle"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxThrottle)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxThrottle, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Throttle, unable to setup balloon help for 'PictureBoxThrottle' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Horn
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If PictureBoxHorn.Tag = "0" Then
            Let BalloonHelpText1 = "The horn is not active."
        Else 'If PictureBoxHorn.Tag = "1" Then
            Let BalloonHelpText1 = "The horn is active."
        End If
        
        Let BalloonHelpText2 = "Horn"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxHorn)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxHorn, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Horn, unable to setup balloon help for 'PictureBoxHorn' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Bell
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If PictureBoxBell.Tag = "0" Then
            Let BalloonHelpText1 = "The bell is not active."
        Else 'If PictureBoxBell.Tag = "1" Then
            Let BalloonHelpText1 = "The bell is active."
        End If
        
        Let BalloonHelpText2 = "Bell"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxBell)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxBell, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Bell, unable to setup balloon help for 'PictureBoxBell' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Sand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If PictureBoxSand.Tag = "0" Then
            Let BalloonHelpText1 = "The sand is not active."
        Else 'If PictureBoxSand.Tag = "1" Then
            Let BalloonHelpText1 = "The sand is active."
        End If
        
        Let BalloonHelpText2 = "Sand"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxSand)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSand, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Sand, unable to setup balloon help for 'PictureBoxSand' control."
            End If
        End If

        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

End Sub

Private Sub BalloonHelpUpdatePart02()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   Automatic Brake
'   Independent Brake
'   Headlight
'   ResetRight
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Automatic Brake
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the trainline brake is at " & Str$(Int(Val(100 - (TextBoxStatusAutomaticBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Automatic Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(pictureboxAutomaticBrake)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxAutomaticBrake, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Automatic Brake, unable to setup balloon help for 'PictureBoxAutomaticBrake' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Independent Brake
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the independent brake is at " & Str$(Int(100 - (Val(TextBoxStatusIndependentBrakePressure.Text) * 100))) & " percent"
        Let BalloonHelpText2 = "Independent Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(transpictureboxIndependentBrake)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBoxIndependentBrake, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Transparent Picture Box Independent Brake, unable to setup balloon help for 'TransPictureBoxIndependentBrake' control."
            End If
        End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Headlight
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       
       If PictureBoxLight.Tag = "0" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently off."
        ElseIf PictureBoxLight.Tag = "1" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on and dimmed."
        ElseIf PictureBoxLight.Tag = "2" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on."
        End If
        Let BalloonHelpText2 = "Headlight"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxLight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxLight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Light, unable to setup balloon help for 'PictureBoxRadioLight' control."
            End If
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ResetRight
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText1 = "Press this button for emergency application of the" & vbCrLf & "digial command control. It will stop all trains on all directions."
        Let BalloonHelpText2 = "Emergency Stop"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(transPictureBoxLight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBoxResetRight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Transparent Picture Box Reset Right, unable to setup balloon help for 'TransPictureBoxResetRight' control."
            End If
        End If
        
    End If

End Sub

Private Sub BalloonHelpUpdatePart03()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   FunctionButtonScreenRight (8)
'   FunctionButtonScreenLeft (8)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText2 = "Function Button 'F1'"
        Let BalloonHelpText1 = "This function button displays the digital gauges for air," & vbCrLf & "oil and water; pressure and temperature."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(1), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(1)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F2'"
        Let BalloonHelpText1 = "This function button displays the digital gauges for resources;" & vbCrLf & "fuel, sand, oil and water amounts."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(2), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(2)' control."
            End If
        End If

        Let BalloonHelpText2 = "Function Button 'F3'"
        Let BalloonHelpText1 = "This function button displays the digital gauges for power;" & vbCrLf & "distribution amounts."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(3), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(3)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F4'"
        Let BalloonHelpText1 = "This function button displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(4), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(4)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F5'"
        Let BalloonHelpText1 = "This function button displays computer screen five." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(5), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(5)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F6"
        Let BalloonHelpText1 = "This function button displays computer screen six." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(6), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(6)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F7'"
        Let BalloonHelpText1 = "This function button displays computer screen seven." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(7), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(7)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F8'"
        Let BalloonHelpText1 = "This function button displays computer screen eight." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenRight(8), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenRight(8)' control."
            End If
        End If
                        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText2 = "Function Button 'F1'"
        Let BalloonHelpText1 = "This function button displays the digital gauges for power," & vbCrLf & "and speed."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(1), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(1)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F2'"
        Let BalloonHelpText1 = "This function button displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(2), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(2)' control."
            End If
        End If

        Let BalloonHelpText2 = "Function Button 'F3'"
        Let BalloonHelpText1 = "This function button displays the digital gauges for power;" & vbCrLf & "distribution amounts."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(3), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(3)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F4'"
        Let BalloonHelpText1 = "This function button displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(4), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(4)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F5'"
        Let BalloonHelpText1 = "This function button displays computer screen five." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(5), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(5)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F6"
        Let BalloonHelpText1 = "This function button displays computer screen six." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(6), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(6)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F7'"
        Let BalloonHelpText1 = "This function button displays computer screen seven." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(7), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(7)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Function Button 'F8'"
        Let BalloonHelpText1 = "This function button displays streaming video" & vbCrLf & "information."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPictureButtonscreenleft(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonScreenLeft(8), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart03, unable to setup balloon help for 'ButtonScreenLeft(8)' control."
            End If
        End If
        
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

End Sub

Private Sub BalloonHelpUpdatePart04()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   ComputerScreenRight (8)
'   ComputerScreenLeft (8)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText2 = "Computer Screen for 'F1'"
        Let BalloonHelpText1 = "This computer screen displays the digital gauges for air," & vbCrLf & "oil and water; pressure and temperature."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(1), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(1)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F2'"
        Let BalloonHelpText1 = "This computer screen displays the digital gauges for resources;" & vbCrLf & "fuel, sand, oil and water amounts."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(2), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(2)' control."
            End If
        End If

        Let BalloonHelpText2 = "Computer Screen for 'F3'"
        Let BalloonHelpText1 = "This computer screen displays the digital gauges for power;" & vbCrLf & "distribution amounts."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(3), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(3)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F4'"
        Let BalloonHelpText1 = "This computer screen displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(4), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(4)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F5'"
        Let BalloonHelpText1 = "This computer screen displays computer screen five." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(5), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(5)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F6"
        Let BalloonHelpText1 = "This computer screen displays computer screen six." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(6), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(6)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F7'"
        Let BalloonHelpText1 = "This computer screen displays computer screen seven." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(7), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(7)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F8'"
        Let BalloonHelpText1 = "This computer screen displays computer screen eight." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonPicturePictureBoxComputerScreenRight(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenRight(8), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenRight(8)' control."
            End If
        End If
                        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText2 = "Computer Screen for 'F1'"
        Let BalloonHelpText1 = "This computer screen displays the digital gauges for" & vbCrLf & "power and speed."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(1), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(1)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F2'"
        Let BalloonHelpText1 = "This computer screen displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(2), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(2)' control."
            End If
        End If

        Let BalloonHelpText2 = "Computer Screen for 'F3'"
        Let BalloonHelpText1 = "This computer screen displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(3), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(3)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F4'"
        Let BalloonHelpText1 = "This computer screen displays computer screen four." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(4), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(4)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F5'"
        Let BalloonHelpText1 = "This computer screen displays computer screen five." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(5), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(5)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F6"
        Let BalloonHelpText1 = "This computer screen displays computer screen six." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(6), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(6)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F7'"
        Let BalloonHelpText1 = "This computer screen displays computer screen seven." & vbCrLf & "It is currenly not used for anything."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(7), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(7)' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Computer Screen for 'F8'"
        Let BalloonHelpText1 = "This computer screen displays streaming video" & vbCrLf & "information."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxComputerScreenLeft(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxComputerScreenLeft(8), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart04, unable to setup balloon help for 'PictureBoxComputerScreenLeft(8)' control."
            End If
        End If
        
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

End Sub

Private Sub BalloonHelpUpdatePart05()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   Transparent Picture Box for Cab Window
'   Speedometer
'   Button Start Engine
'   Button Help
'   Button Close GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText2 = "Cab Window"
        Let BalloonHelpText1 = "This is the cab window where you 'see'. The" & vbCrLf & "streaming video is played here."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureBoxCabWindow)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBoxCabWindow, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'TransPictureBoxCabWindow' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Speedometer"
        Let BalloonHelpText1 = "This gauge displays the locomotive" & vbCrLf & "speed. It is not scaled speed but" & vbCrLf & "digital speed steps."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(PictureBoxSpeedometer)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSpeedometer, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'PictureBoxSpeedometer' control."
            End If
        End If
       
        Let BalloonHelpText2 = "Start/Stop Button"
        Let BalloonHelpText1 = "This button is used to start and" & vbCrLf & "stop the prime mover of the locomotive."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonenginestart)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonEngineStart, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'buttonenginestart' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Help Button"
        Let BalloonHelpText1 = "This button is used to help operate" & vbCrLf & "the locomotive with tips for beginners."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(buttonhelp)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonHelp, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonHelp' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Close Button"
        Let BalloonHelpText1 = "This button is used to close the" & vbCrLf & "graphics user screen and return to" & vbCrLf & "the previous screen."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(ButtonCloseGUI)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonCloseGUI, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonCloseGUI' control."
            End If
        End If
       
        Let BalloonHelpText2 = "Fuel Fill Button"
        Let BalloonHelpText1 = "This button is used to imitate" & vbCrLf & "the re-filling of diesel fuel for the locomotive."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(ButtonFillFuel)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonFillFuel, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonFillFuel' control."
            End If
        End If
       
        Let BalloonHelpText2 = "Oil Fill Button"
        Let BalloonHelpText1 = "This button is used to imitate" & vbCrLf & "the re-filling of oil for the locomotive."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(ButtonFillOil)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonFillOil, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonFillOil' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Sand Fill Button"
        Let BalloonHelpText1 = "This button is used to imitate" & vbCrLf & "the re-filling of sand for the locomotive."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(ButtonFillSand)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonFillSand, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonFillSand' control."
            End If
        End If
        
        Let BalloonHelpText2 = "Water Fill Button"
        Let BalloonHelpText1 = "This button is used to imitate" & vbCrLf & "the re-filling of water for the locomotive."

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(ButtonFillWater)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonFillWater, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart05, unable to setup balloon help for 'ButtonFillWater' control."
            End If
        End If
        
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

End Sub

Private Sub BalloonHelpUpdatePart06()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons for
'   Transparent Picture Box for Auxillary Switches
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Computer 'F' Button on Screen Left
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Dim TemporaryValue As Integer
        For TemporaryValue = 0 To 14
            Let BalloonHelpText2 = "Auxillary Switch" & TemporaryValue
            If TransPictureAuxillarySwitch(TemporaryValue).Tag = "0" Then
                Let BalloonHelpText1 = "This auxillary switch is used to turn on/off'" & vbCrLf & "specific functions manually." & vbCrLf & vbCrLf & "This switch is off."
            Else 'if TransPictureAuxillarySwitch(index).tag = "1"
                Let BalloonHelpText1 = "This auxillary switch is used to turn on/off'" & vbCrLf & "specific functions manually." & vbCrLf & vbCrLf & "This switch is on."
            End If
            'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureauxillaryswitch(temporaryvalue))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureAuxillarySwitch(TemporaryValue), BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
            If BalloonHelpSetup = 0 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, BalloonHelpUpdatePart06, unable to setup balloon help for 'TransPictureAuxillarySwitch' control."
                End If
            End If
        Next TemporaryValue
       
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccDefault

    End If

End Sub

















