VERSION 4.00
Begin VB.Form MainlineSpeedTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode  - Speed Table"
   ClientHeight    =   6960
   ClientLeft      =   1320
   ClientTop       =   1725
   ClientWidth     =   11100
   Height          =   7365
   Icon            =   "MainlineSpeedTable.frx":0000
   Left            =   1260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11100
   Top             =   1380
   Width           =   11220
   Begin VB.CommandButton ButtonImport 
      Caption         =   "&Import"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   133
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   8400
      TabIndex        =   132
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "&List"
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonEasyScreen 
      Caption         =   "Easy Screen"
      Height          =   255
      Left            =   1440
      TabIndex        =   87
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   180
      Picture         =   "MainlineSpeedTable.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox SevenByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8040
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   13440
      Width           =   1215
   End
   Begin VB.TextBox SevenByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9360
      TabIndex        =   124
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   13440
      Width           =   1215
   End
   Begin VB.TextBox SevenByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10680
      TabIndex        =   123
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   13440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveNumberGO 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Enabled         =   0   'False
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.Data LocomotiveDatabaseDecoderSpeedTable 
      Caption         =   "Locomotive Number"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDecoders"
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   92
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV69D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   69
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   76
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   69
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   69
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   91
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV91D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   91
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV79D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   79
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   46
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV89D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   89
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton ButtonExport 
      Caption         =   "&Export"
      Height          =   255
      Left            =   5760
      TabIndex        =   89
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   9720
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   7080
      TabIndex        =   88
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV67D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   67
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   82
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   67
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   67
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV68D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   68
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   79
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   68
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   68
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV70D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   70
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   73
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   70
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   70
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV71D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   71
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   70
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   71
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   71
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV72D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   72
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   67
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   72
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   72
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV73D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   73
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   64
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   73
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   73
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV74D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   74
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   61
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   74
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   74
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV75D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   75
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   58
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   75
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   75
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV76D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   76
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   55
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   76
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   76
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV77D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   77
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   52
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   77
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   77
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV78D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   78
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   49
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   78
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   78
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   79
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   79
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV80D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   80
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   43
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   80
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   80
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV81D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   81
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   40
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   81
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   81
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV82D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   82
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   37
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   82
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   82
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV83D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   83
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   34
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   83
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   83
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV84D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   84
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   31
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   84
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   84
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV85D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   85
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   28
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   85
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   85
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV86D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   86
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   86
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   86
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV87D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   87
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   22
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   87
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   87
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV88D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   88
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   88
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   88
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   89
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   89
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV90D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   90
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   90
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   90
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   91
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV92D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   92
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   92
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   93
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   93
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV93D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   93
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVd 
      Alignment       =   2  'Center
      DataField       =   "LocomotiveDecoderCV94D"
      DataSource      =   "LocomotiveDatabaseDecoderSpeedTable"
      Height          =   285
      Index           =   94
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVh 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   94
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDecoderCVb 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   94
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   11160
      Top             =   2040
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   11160
      Top             =   1560
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   11160
      Top             =   960
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label43 
      Caption         =   "Enter the values you would like in the speed table setting. If you need some help, try the easy screen for a visual persective."
      Height          =   375
      Left            =   840
      TabIndex        =   131
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "Decimal"
      Height          =   255
      Left            =   1440
      TabIndex        =   130
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "Hexidecimal"
      Height          =   255
      Left            =   8400
      TabIndex        =   129
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Binary"
      Height          =   255
      Left            =   9720
      TabIndex        =   128
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Error Byte"
      Height          =   195
      Left            =   7200
      TabIndex        =   126
      Top             =   13440
      Width           =   690
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Locomotive Number"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9120
      TabIndex        =   122
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label LocomotiveCommunicationStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   840
      TabIndex        =   121
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "CV67 - Step 1"
      Height          =   195
      Left            =   5760
      TabIndex        =   120
      Top             =   6120
      Width           =   990
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "CV68 - Step 2"
      Height          =   195
      Left            =   5760
      TabIndex        =   119
      Top             =   5760
      Width           =   990
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "CV69 - Step 3"
      Height          =   195
      Left            =   5760
      TabIndex        =   118
      Top             =   5400
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "CV70 - Step 4"
      Height          =   195
      Left            =   5760
      TabIndex        =   117
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "CV71 - Step 5"
      Height          =   195
      Left            =   5760
      TabIndex        =   116
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CV72 - Step 6"
      Height          =   195
      Left            =   5760
      TabIndex        =   115
      Top             =   4320
      Width           =   990
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "CV73 - Step 7"
      Height          =   195
      Left            =   5760
      TabIndex        =   114
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "CV74 - Step 8"
      Height          =   195
      Left            =   5760
      TabIndex        =   113
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "CV75 - Step 9"
      Height          =   195
      Left            =   5760
      TabIndex        =   112
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "CV76 - Step 10"
      Height          =   195
      Left            =   5760
      TabIndex        =   111
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "CV77 - Step 11"
      Height          =   195
      Left            =   5760
      TabIndex        =   110
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "CV78 - Step 12"
      Height          =   195
      Left            =   5760
      TabIndex        =   109
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "CV79 - Step 13"
      Height          =   195
      Left            =   5760
      TabIndex        =   108
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "CV80 - Step 14"
      Height          =   195
      Left            =   5760
      TabIndex        =   107
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "CV81 - Step 15"
      Height          =   195
      Left            =   120
      TabIndex        =   106
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "CV82 - Step 16"
      Height          =   195
      Left            =   120
      TabIndex        =   105
      Top             =   5760
      Width           =   1080
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "CV83 - Step 17"
      Height          =   195
      Left            =   120
      TabIndex        =   104
      Top             =   5400
      Width           =   1080
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "CV84 - Step 18"
      Height          =   195
      Left            =   120
      TabIndex        =   103
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "CV85 - Step 19"
      Height          =   195
      Left            =   120
      TabIndex        =   102
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "CV86 - Step 20"
      Height          =   195
      Left            =   120
      TabIndex        =   101
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "CV87 - Step 21"
      Height          =   195
      Left            =   120
      TabIndex        =   100
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CV88 - Step 22"
      Height          =   195
      Left            =   120
      TabIndex        =   99
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "CV89 - Step 23"
      Height          =   195
      Left            =   120
      TabIndex        =   98
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CV90 - Step 24"
      Height          =   195
      Left            =   120
      TabIndex        =   97
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CV91 - Step 25"
      Height          =   195
      Left            =   120
      TabIndex        =   96
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CV92 - Step 26"
      Height          =   195
      Left            =   120
      TabIndex        =   95
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CV93 - Step 27"
      Height          =   195
      Left            =   120
      TabIndex        =   94
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CV94 - Step 28"
      Height          =   195
      Left            =   120
      TabIndex        =   93
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Decimal"
      Height          =   255
      Left            =   7080
      TabIndex        =   92
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hexidecimal"
      Height          =   255
      Left            =   2760
      TabIndex        =   91
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Binary"
      Height          =   255
      Left            =   4080
      TabIndex        =   90
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "MainlineSpeedTable"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub ButtonClose_Click()

' --------------------' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            If TemporaryScreen = "Mainline Speed Table Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Speed Table Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Speed Table Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Speed Table Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub
Private Sub ButtonEasyScreen_Click()
 
Load MainlineEasyScreenSpeedTable
MainlineEasyScreenSpeedTable.Show vbModeless

End Sub

Private Sub ButtonExport_Click()

Let MousePointer = vbHourglass

Let ButtonUpdate.Enabled = False
Let ButtonClose.Enabled = False
Let ButtonExport.Enabled = False
Let ButtonPreset.Enabled = False

' *******************************************************************************************************************************************
' Exporting - The Configuration Variables
' *******************************************************************************************************************************************
'

For Y = 67 To 94

    Let onebyted.Text = Int(Val(LocomotiveNumberGO.Text) / 256)
    Let twoByteD.Text = Val(LocomotiveNumberGO.Text) - (Val(onebyted.Text) * 256)
    Let onebyted.Text = Val(onebyted.Text) + 128 + 64
    
    Let TemporaryByteOne = 0
    Let TemporaryByteTwo = Y - 1
        
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
        
    Let TemporaryByteOne = TemporaryByteOne + 8 + 4
                
    Let ThreeByteD.Text = TemporaryByteOne
    Let FourByteD.Text = TemporaryByteTwo
    Let FiveByteD.Text = Val(LocomotiveDecoderCVd(Y).Text)
    Let SixByteD.Text = ""
    
' -------------------------------------------------------------------------------------------------------------------------------------------
' Error Byte Calculations
' -------------------------------------------------------------------------------------------------------------------------------------------
'
' The error byte is kepped in the seventh btye (seventh text box on the screen). This is calculated by going through the six bytes and
' counting the number of 'ones', XOR, to determine the error btye.

            Let sevenbyted.Text = "0"

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

                If onebyted.Text <> "" Then
                    If onebyted.Text <> "   " Then
                        If Mid$(OneByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If twoByteD.Text <> "" Then
                    If twoByteD.Text <> "   " Then
                        If Mid$(TwoByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If ThreeByteD.Text <> "" Then
                    If ThreeByteD.Text <> "   " Then
                        If Mid$(ThreeByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If FourByteD.Text <> "" Then
                    If FourByteD.Text <> "   " Then
                        If Mid$(FourByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If FiveByteD.Text <> "" Then
                If FiveByteD.Text <> "   " Then
                    If Mid$(FiveByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If SixByteD.Text <> "" Then
                    If SixByteD.Text <> "   " Then
                        If Mid$(SixByteB.Text, X, 1) = "1" Then
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

                If Int(temp / 2) <> (temp / 2) Then sevenbyted.Text = Val(sevenbyted.Text) + bitvalue

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

            Let LocomotiveCommunicationStatus.Caption = "Status: Updating CV" + Str$(Y) + " Register"

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
            
            If onebyted.Text <> "" Then
                If onebyted.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & OneByteH.Text
                End If
            End If
            If twoByteD.Text <> "" Then
                If twoByteD.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & TwoByteH.Text
                End If
            End If
            If ThreeByteD.Text <> "" Then
                If ThreeByteD.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & ThreeByteH.Text
                End If
            End If
            If FourByteD.Text <> "" Then
                If FourByteD.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & FourByteH.Text
                End If
            End If
            If FiveByteD.Text <> "" Then
                If FiveByteD.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & FiveByteH.Text
                End If
            End If
            If SixByteD.Text <> "" Then
                If SixByteD.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & SixByteH.Text
                End If
            End If
            If sevenbyted.Text <> "" Then
                If sevenbyted.Text <> "   " Then
                    Let CommandControl = CommandControl & " " & SevenByteH.Text
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

            Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + CommandControl + Chr$(10)
            Let MainScreen.TextBoxCommunicationWindowDCC.SelStart = Len(MainScreen.TextBoxCommunicationWindowDCC.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

            MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent parameters as per National Model Railroader Association specification.
    
            Let LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

        If Left$(MainScreen.Labelmode.Text, 6) = "Analog" Then
           While Right$(TextBoxCommunicationWindowDCC.Text, 9) <> "COMMAND: "
               DoEvents
           Wend
        ElseIf Left$(MainScreen.Labelmode.Text, 7) = "Digital" Then
           While Right(TextBoxCommunicationWindowDCC.Text, 1) <> &H80
               DoEvents
           Wend
        End If
            Let LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

        Temporary = DoEvents
' -------------------------------------------------------------------------------------------------------------------------
' Dequeue packet only if packet was queued with "q" command
' -------------------------------------------------------------------------------------------------------------------------
If MainScreen!checkboxdequeuepacket.Value = vbChecked Then

            Let CommandControl = "d"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
                If onebyted.Text <> "" Then
                    If onebyted.Text <> "   " Then
                        Let CommandControl = CommandControl & " " & OneByteH.Text
                    End If
                End If
                If twoByteD.Text <> "" Then
                    If twoByteD.Text <> "   " Then
                        Let CommandControl = CommandControl & " " & TwoByteH.Text
                    End If
                End If
                'If ThreeByteD.Text <> "" Then
                '    If ThreeByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & ThreeByteH.Text
                '    End If
                'End If
                'If FourByteD.Text <> "" Then
                '    If FourByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & FourByteH.Text
                '    End If
                'End If
                'If FiveByteD.Text <> "" Then
                '    If FiveByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & FiveByteH.Text
                '    End If
                'End If
                'If SixByteD.Text <> "" Then
                '    If SixByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & SixByteH.Text
                '    End If
                'End If
                'If SevenByteD.Text <> "" Then
                '    If SevenByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & SevenByteH.Text
                '    End If
                'end If
            
' We finish the string by adding a carriage return to it. The command station will then recognize the command when sent.
            
                Let CommandControl = CommandControl + Chr$(13)
            
' Start Sending the information
'
' The first order of business before sending the command to the communication port is to add the command string to the
' communication window. This communication window is located the the Automatic Train Control Form, and controls which
' controls all the characters going in and out of the communication port.
' The following statement, i believe, set the cursor to the end of the new text being ddisplayed in the communication window.

                Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + CommandControl + Chr$(10)
                Let MainScreen.TextBoxCommunicationWindowDCC.SelStart = Len(MainScreen.TextBoxCommunicationWindowDCC.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

                MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent parameters as per National Model Railroader Association specification.
    
                Let LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

        If Left$(MainScreen.Labelmode.Text, 6) = "Analog" Then
           While Right$(TextBoxCommunicationWindowDCC.Text, 9) <> "COMMAND: "
               DoEvents
           Wend
        ElseIf Left$(MainScreen.Labelmode.Text, 7) = "Digital" Then
           While Right(TextBoxCommunicationWindowDCC.Text, 1) <> &H80
               DoEvents
           Wend
        End If
                Let LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

    End If
            
    
Next Y

Let MousePointer = vbDefault

Let ButtonUpdate.Enabled = True
Let ButtonClose.Enabled = True
Let ButtonExport.Enabled = True
Let ButtonPreset.Enabled = True

End Sub


Private Sub ButtonList_Click()
 
Load LocomotiveCVSpreadsheet
LocomotiveCVSpreadsheet.Show vbModeless

End Sub

Private Sub ButtonPrint_Click()

    MainlineSpeedTable.PrintForm
    
End Sub

Private Sub ButtonUpdate_Click()
    
    LocomotiveDatabaseDecoderSpeedTable.Recordset.Edit
    LocomotiveDatabaseDecoderSpeedTable.Recordset.Update

End Sub






























Private Sub FiveByteD_Change()
    
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
        If TemporaryScreen = "Mainline Speed Table Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Speed Table Screen"
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
            Let Ini.Value = "Mainline Speed Table Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Speed Table Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Speed Table Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Locomotive Number
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let ButtonClose.Enabled = False
    Let MousePointer = vbArrowHourglass
    Let DisplayErrorMessage.Value = vbChecked
    If MainlineDiesel!CheckAddressShort.Value = vbChecked Then
        Let LocomotiveDatabaseDecoderSpeedTable.RecordSource = "LocomotiveDecodersShort"
    Else
        Let LocomotiveDatabaseDecoderSpeedTable.RecordSource = "LocomotiveDecodersLong"
    End If
    LocomotiveDatabaseDecoderSpeedTable.Refresh
    LocomotiveDatabaseDecoderSpeedTable.Recordset.MoveLast
    Let Temporary = Val(MainlineDecoder.LocomotiveNumber.Text)
    LocomotiveDatabaseDecoderSpeedTable.Recordset.AbsolutePosition = Temporary
    Let MousePointer = vbDefault
    Let ButtonClose.Enabled = True

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Speed Table Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineSpeedTable.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineSpeedTable.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineSpeedTable.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineSpeedTable.Height)

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
                Let Ini.Value = "Mainline Speed Table Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Speed Table Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineSpeedTable.Hide
    'unload Mainlinespeedtable

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
    Let Ini.Application = "Mainline Speed Table Screen"
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
        MainlineSpeedTable.Left = (Screen.Width - Width) / 2
        MainlineSpeedTable.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineSpeedTable.Width > Screen.Width Then
            Let MainlineSpeedTable.Left = Screen.Width - MainlineSpeedTable.Width
        Else
            Let MainlineSpeedTable.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineSpeedTable.Height > Screen.Height Then
            Let MainlineSpeedTable.Top = Screen.Height - MainlineSpeedTable.Height
        Else
            Let MainlineSpeedTable.Top = Val(TemporaryValueTop)
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
        Call BalloonHelpPart3
        Call BalloonHelpPart4
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let LocomotiveDatabaseDecoderSpeedTable.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDecoders.mdb"

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

    If MainlineSpeedTable.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineSpeedTable.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Speed Table Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineSpeedTable.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineSpeedTable.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineSpeedTable.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineSpeedTable.Height)
 
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


Private Sub LocomotiveDecoderCVd_Change(Index As Integer)

Let ConfigurationVariableLow = 0
Let ConfigurationVariableHigh = 255

If Val(LocomotiveDecoderCVb(Index).Text) < ConfigurationVariableLow Then
    
    If DisplayErrorMessage.Value = vbChecked Then
        
        Let Message = "The value entered for the Configuration Variable"
        Let Message = Message + Str$(Format(Index, "000"))
        Let Message = Message + " is not valid. You" + Chr$(13)
        Let Message = Message + " enter a number less then"
        Let Message = Message + Str$(ConfigurationVariableLow)
        Let Message = Message + ". Please use a value between"
        Let Message = Message + Str$(ConfigurationVariableLow)
        Let Message = Message + " and"
        Let Message = Message + Str$(ConfigurationVariableHigh)
        Let Message = Message + " only. " + Chr$(13)
        Let Message = Message + Chr$(13)
        Let Message = Message + "Configuration Variable"
        Let Message = Message + Str$(Format(Index, "000"))
        Let Message = Message + " will be reset to"
        Let Message = Message + Str$(ConfigurationVariableLow)
        Let Message = Message + "." + Chr$(13)
        Let Message = Message + Chr$(13)
        Let Message = Message + "Do you wish to continue to receive these warning messages?"
        Let ReturnMessage = MsgBox(Message, vbYesNo + vbInformation, "CV - Entry Error")
        If ReturnMessage = vbNo Then
            Let DisplayErrorMessage.Value = vbUnchecked
        End If
    End If
    'LocomotiveDatabaseDecoder.Recordset.Edit
    'temporary = DoEvents
    LocomotiveDecoderCVd(Index).Text = ConfigurationVariableLow
    'temporary = DoEvents
    'LocomotiveDatabaseDecoder.Recordset.Update
End If
    
Rem If the value of the Primary Address is greater then one hundred and twenty seven, then inform the user that this is
Rem not valid. Reset the value to the maximim and return control to the user.

If Val(LocomotiveDecoderCVd(Index).Text) > ConfigurationVariableHigh Then
    If DisplayErrorMessage.Value = vbChecked Then
        Let Message = "The value entered for the Configuration Variable"
        Let Message = Message + Str$(Format(Index, "000"))
        Let Message = Message + " is not valid. You enter a number" + Chr$(13)
        Let Message = Message + "greater then"
        Let Message = Message + Str$(ConfigurationVariableHigh)
        Let Message = Message + ". Please use a value between"
        Let Message = Message + Str$(ConfigurationVariableLow)
        Let Message = Message + " and"
        Let Message = Message + Str$(ConfigurationVariableHigh)
        Let Message = Message + " only. " + Chr$(13)
        Let Message = Message + Chr$(13)
        Let Message = Message + "Configuration Variable"
        Let Message = Message + Str$(Format(Index, "000"))
        Let Message = Message + " will be reset to"
        Let Message = Message + Str$(ConfigurationVariableHigh)
        Let Message = Message + "." + Chr$(13)
        Let Message = Message + Chr$(13)
        Let Message = Message + "Do you wish to continue to receive these warning messages?"
            
        Let ReturnMessage = MsgBox(Message, vbYesNo + vbInformation, "CV - Entry Error")
            
        If ReturnMessage = vbNo Then
            Let DisplayErrorMessage.Value = vbUnchecked
        End If
    End If
    'LocomotiveDatabaseDecoder.Recordset.Edit
    'temporary = DoEvents
    Let LocomotiveDecoderCVd(Index) = ConfigurationVariableHigh
    'temporary = DoEvents
    'LocomotiveDatabaseDecoder.Recordset.Update
End If

' -------------------------------------------------------------------------------------------------------------------------------------------
' Hexidecimal Converstion
' -------------------------------------------------------------------------------------------------------------------------------------------
    
Let LocomotiveDecoderCVh(Index).Text = Hex(Val(LocomotiveDecoderCVd(Index).Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

If Len(LocomotiveDecoderCVh(Index).Text) = 1 Then
    Let LocomotiveDecoderCVh(Index).Text = "0" + LocomotiveDecoderCVh(Index).Text
End If
    
' -------------------------------------------------------------------------------------------------------------------------------------------
' Binary Conversion
' -------------------------------------------------------------------------------------------------------------------------------------------

Let temp = Val(LocomotiveDecoderCVd(Index).Text)
        
If temp / 128 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = "1"
    Let temp = temp - 128
Else
    Let LocomotiveDecoderCVb(Index).Text = "0"
End If
        
If temp / 64 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 64
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 32 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 32
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 16 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 16
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 8 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 8
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 4 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 4
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 2 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 2
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If
        
If temp / 1 >= 1 Then
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "1"
    Let temp = temp - 1
Else
    Let LocomotiveDecoderCVb(Index).Text = LocomotiveDecoderCVb(Index).Text + "0"
End If

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



Private Sub BalloonHelpPart1()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim t As Boolean
        Dim f As Boolean
        Let t = True
        Let f = False
      
        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step one of twenty-eight, stored in configuration variable sixty-seven," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(67))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(67), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step two of twenty-eight, stored in configuration variable sixty-eight," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(68))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(68), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step three of twenty-eight, stored in configuration variable sixty-nine," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(69))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(69), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step four of twenty-eight, stored in configuration variable seventy," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(70))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(70), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step five of twenty-eight, stored in configuration variable seventy-one," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(71))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(71), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step six of twenty-eight, stored in configuration variable seventy-two," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(72))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(72), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seven of twenty-eight, stored in configuration variable seventy-three," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(73))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(73), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eight of twenty-eight, stored in configuration variable seventy-four," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(74))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(74), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nine of twenty-eight, stored in configuration variable seventy-five," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Nine"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(75))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(75), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step ten of twenty-eight, stored in configuration variable seventy-six," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Ten"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(76))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(76), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eleven of twenty-eight, stored in configuration variable seventy-seven," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Eleven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(77))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(77), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twelve of twenty-eight, stored in configuration variable seventy-eight," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twelve"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(78))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(78), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step thirteen of twenty-eight, stored in configuration variable seventy-nine," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Thirteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(79))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(79), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fourteen of twenty-eight, stored in configuration variable eighty," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Fourteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(80))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(80), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fifteen of twenty-eight, stored in configuration variable eighty-one," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Fifteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(81))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(81), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step sixteen of twenty-eight, stored in configuration variable eighty-two," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Sixteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(82))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(82), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seventeen of twenty-eight, stored in configuration variable eighty-three," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Seventeen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(83))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(83), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eighteen of twenty-eight, stored in configuration variable eighty-four," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Eighteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(84))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(84), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nineteen of twenty-eight, stored in configuration variable eighty-five," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Nineteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(85))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(85), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty of twenty-eight, stored in configuration variable eighty-six," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(86))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(86), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-one of twenty-eight, stored in configuration variable eighty-seven," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(87))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(87), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-two of twenty-eight, stored in configuration variable eighty-eight," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(88))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(88), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-three of twenty-eight, stored in configuration variable eighty-nine," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(89))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(89), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-four of twenty-eight, stored in configuration variable ninety," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(90))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(90), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-five of twenty-eight, stored in configuration variable ninety-one," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(91))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(91), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-six of twenty-eight, stored in configuration variable ninety-two," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(92))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(92), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-seven of twenty-eight, stored in configuration variable ninety-three," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(93))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(93), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-eight of twenty-eight, stored in configuration variable ninety-four," & vbCrLf & "in DECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(94))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(94), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub

Private Sub BalloonHelpPart2()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim t As Boolean
        Dim f As Boolean
        Let t = True
        Let f = False
      
        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step one of twenty-eight, stored in configuration variable sixty-seven," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(67))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(67), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step two of twenty-eight, stored in configuration variable sixty-eight," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(68))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(68), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step three of twenty-eight, stored in configuration variable sixty-nine," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(69))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(69), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step four of twenty-eight, stored in configuration variable seventy," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(70))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(70), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step five of twenty-eight, stored in configuration variable seventy-one," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(71))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(71), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step six of twenty-eight, stored in configuration variable seventy-two," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(72))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(72), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seven of twenty-eight, stored in configuration variable seventy-three," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(73))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(73), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eight of twenty-eight, stored in configuration variable seventy-four," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(74))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(74), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nine of twenty-eight, stored in configuration variable seventy-five," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Nine"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(75))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(75), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step ten of twenty-eight, stored in configuration variable seventy-six," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Ten"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(76))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(76), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eleven of twenty-eight, stored in configuration variable seventy-seven," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Eleven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(77))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(77), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twelve of twenty-eight, stored in configuration variable seventy-eight," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twelve"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(78))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(78), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step thirteen of twenty-eight, stored in configuration variable seventy-nine," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Thirteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(79))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(79), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fourteen of twenty-eight, stored in configuration variable eighty," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Fourteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(80))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(80), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fifteen of twenty-eight, stored in configuration variable eighty-one," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Fifteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(81))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(81), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step sixteen of twenty-eight, stored in configuration variable eighty-two," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Sixteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(82))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(82), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seventeen of twenty-eight, stored in configuration variable eighty-three," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Seventeen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(83))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(83), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eighteen of twenty-eight, stored in configuration variable eighty-four," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Eighteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(84))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(84), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nineteen of twenty-eight, stored in configuration variable eighty-five," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Nineteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(85))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(85), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty of twenty-eight, stored in configuration variable eighty-six," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(86))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(86), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-one of twenty-eight, stored in configuration variable eighty-seven," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(87))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(87), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-two of twenty-eight, stored in configuration variable eighty-eight," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(88))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(88), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-three of twenty-eight, stored in configuration variable eighty-nine," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(89))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(89), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-four of twenty-eight, stored in configuration variable ninety," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(90))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(90), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-five of twenty-eight, stored in configuration variable ninety-one," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(91))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(91), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-six of twenty-eight, stored in configuration variable ninety-two," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(92))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(92), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-seven of twenty-eight, stored in configuration variable ninety-three," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(93))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(93), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-eight of twenty-eight, stored in configuration variable ninety-four," & vbCrLf & "in HEXIDECIMAL format."
        Let TemporaryText2 = "Speed Step Twenty-Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(94))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(94), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub

Private Sub BalloonHelpPart3()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim t As Boolean
        Dim f As Boolean
        Let t = True
        Let f = False
      
        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step one of twenty-eight, stored in configuration variable sixty-seven," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(67))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(67), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step two of twenty-eight, stored in configuration variable sixty-eight," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(68))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(68), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step three of twenty-eight, stored in configuration variable sixty-nine," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(69))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(69), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step four of twenty-eight, stored in configuration variable seventy," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(70))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(70), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step five of twenty-eight, stored in configuration variable seventy-one," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(71))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(71), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step six of twenty-eight, stored in configuration variable seventy-two," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(72))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(72), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seven of twenty-eight, stored in configuration variable seventy-three," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(73))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(73), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eight of twenty-eight, stored in configuration variable seventy-four," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(74))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(74), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nine of twenty-eight, stored in configuration variable seventy-five," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Nine"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(75))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(75), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step ten of twenty-eight, stored in configuration variable seventy-six," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Ten"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(76))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(76), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eleven of twenty-eight, stored in configuration variable seventy-seven," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Eleven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(77))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(77), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twelve of twenty-eight, stored in configuration variable seventy-eight," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twelve"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(78))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(78), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step thirteen of twenty-eight, stored in configuration variable seventy-nine," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Thirteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(79))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(79), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fourteen of twenty-eight, stored in configuration variable eighty," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Fourteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(80))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(80), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step fifteen of twenty-eight, stored in configuration variable eighty-one," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Fifteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(81))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(81), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step sixteen of twenty-eight, stored in configuration variable eighty-two," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Sixteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(82))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(82), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step seventeen of twenty-eight, stored in configuration variable eighty-three," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Seventeen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(83))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(83), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step eighteen of twenty-eight, stored in configuration variable eighty-four," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Eighteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(84))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(84), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step nineteen of twenty-eight, stored in configuration variable eighty-five," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Nineteen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(85))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(85), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty of twenty-eight, stored in configuration variable eighty-six," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(86))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(86), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-one of twenty-eight, stored in configuration variable eighty-seven," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(87))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(87), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-two of twenty-eight, stored in configuration variable eighty-eight," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(88))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(88), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-three of twenty-eight, stored in configuration variable eighty-nine," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(89))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(89), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-four of twenty-eight, stored in configuration variable ninety," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(90))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(90), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-five of twenty-eight, stored in configuration variable ninety-one," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(91))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(91), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-six of twenty-eight, stored in configuration variable ninety-two," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(92))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(92), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-seven of twenty-eight, stored in configuration variable ninety-three," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(93))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(93), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "These text boxes are used to numericaly show the" & vbCrLf & "speed table. This text box will specifically show the value for speed" & vbCrLf & "step twenty-eight of twenty-eight, stored in configuration variable ninety-four," & vbCrLf & "in BINARY format."
        Let TemporaryText2 = "Speed Step Twenty-Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(94))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(94), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub

Private Sub BalloonHelpPart4()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim TemporaryText1 As String
      Dim TemporaryText2 As String
      Dim i As Long
      Dim t As Boolean
      Dim f As Boolean
      Let t = True
      Let f = False
    
      Let TemporaryText1 = "This text box displays the currently selected" & vbCrLf & "locomotive (decoder)."
      Let TemporaryText2 = "Locomotive (decoder) Number"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveNumberGO)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveNumberGO, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
      Let TemporaryText1 = "This button will display the window for listing" & vbCrLf & "all the locomotive (decoder) speed table settings. You" & vbCrLf & "can compare the values of other configuration variable with other" & vbCrLf & "locomotives (decoders)."
      Let TemporaryText2 = "List Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonList)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonList, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
      Let TemporaryText1 = "This button will display an 'Easy Screen' window for setting" & vbCrLf & "all the locomotive (decoder) speed table settings."
      Let TemporaryText2 = "Easy Screen Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonEasyScreen)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonEasyScreen, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
      Let TemporaryText1 = "This button will update the database with the" & vbCrLf & "current values listed in the textboxes. It will not write" & vbCrLf & "these valuse to the loadable spped table in the locomotive (decoder)."
      Let TemporaryText2 = "Update Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonUpdate)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonUpdate, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
      Let TemporaryText1 = "This button will import the values of the configuration" & vbCrLf & "variables from the locomotive (decoder) on the programming" & vbCrLf & "track to this 'Programming Mode- Speed Table' screen."
      Let TemporaryText2 = "Import Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonImport)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonImport, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
      Let TemporaryText1 = "This button will export the values of the configuration" & vbCrLf & "variables to the locomotive (decoder) on the programming" & vbCrLf & "track."
      Let TemporaryText2 = "Exports Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonExport)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonExport, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

      Let TemporaryText1 = "This button prints the current window to your printer."
      Let TemporaryText2 = "Print Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
      Let TemporaryText1 = "This button will close this screen and return you" & vbCrLf & "to the 'Programming Mode - Decoder' screen."
      Let TemporaryText2 = "Close Button"
      'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
      Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")


End Sub
