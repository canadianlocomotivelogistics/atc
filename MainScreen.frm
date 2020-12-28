VERSION 4.00
Begin VB.Form MainScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Main Screen"
   ClientHeight    =   6330
   ClientLeft      =   1725
   ClientTop       =   4305
   ClientWidth     =   9915
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   7020
   Icon            =   "MainScreen.frx":0000
   Left            =   1665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9915
   Top             =   3675
   Width           =   10035
   Begin VB.Timer TimerAutoStart 
      Interval        =   10000
      Left            =   10140
      Top             =   5760
   End
   Begin VB.TextBox TextboxCommunication2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "CMRI device is currently configured for"
      Top             =   5400
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2040
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   10860
      Width           =   3855
   End
   Begin VB.TextBox LabelMode 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox LabelBaudRateSetting3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox LabelBaudRateSetting2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox LabelBaudRateSetting1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox LabelCommunicationPortSetting3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5640
      Width           =   2475
   End
   Begin VB.TextBox LabelCommunicationPortSetting2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5400
      Width           =   2475
   End
   Begin VB.TextBox LabelCommunicationPortSetting1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "(not configured)"
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox LabelNotes 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "Notes:"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox LabelStatus 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "Status:"
      Top             =   120
      Width           =   9675
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      Picture         =   "MainScreen.frx":0442
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   43
      Top             =   6480
      Width           =   540
   End
   Begin VB.CommandButton ButtonExpand 
      Caption         =   ">>"
      Height          =   210
      Left            =   9480
      TabIndex        =   42
      Top             =   6060
      Width           =   375
   End
   Begin VB.TextBox TextboxCommunication3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "Other device is currently configured for"
      Top             =   5640
      Width           =   2970
   End
   Begin VB.Timer TimerTipoftheDay 
      Left            =   10080
      Top             =   3600
   End
   Begin VB.TextBox TextboxCommunication1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "DCC device is currently configured for"
      Top             =   5160
      Width           =   2820
   End
   Begin VB.TextBox SixByteD 
      Alignment       =   2  'Center
      DataField       =   "ThirdByte"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Height          =   315
      Left            =   2040
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxDequeuePackets 
      Caption         =   "Dequeue Packets?"
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   10560
      Width           =   3615
   End
   Begin VB.CheckBox CheckBoxQueueCommand 
      Caption         =   "Queued"
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   10245
      Width           =   975
   End
   Begin VB.TextBox SevenByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox SixByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   9480
      Width           =   1215
   End
   Begin VB.TextBox FiveByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox FourByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox TwoByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox OneByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "00000000"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox SevenByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox SixByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   9480
      Width           =   1215
   End
   Begin VB.TextBox FiveByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox FourByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox TwoByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox OneByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox SevenByteD 
      Alignment       =   2  'Center
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox FiveByteD 
      Alignment       =   2  'Center
      DataField       =   "ThirdByte"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox FourByteD 
      Alignment       =   2  'Center
      DataField       =   "SecondByte"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteD 
      Alignment       =   2  'Center
      DataField       =   "FirstByte"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox TwoByteD 
      Alignment       =   2  'Center
      DataField       =   "AddressByteTwo"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox OneByteD 
      Alignment       =   2  'Center
      DataField       =   "AddressByteOne"
      DataSource      =   "LocomotiveDatabaseConsistControl(0)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1215
   End
   Begin TabDlg.SSTab CommunicationTab 
      Height          =   4095
      Left            =   60
      TabIndex        =   39
      Top             =   480
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Communication Window for DCC"
      TabPicture(0)   =   "MainScreen.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CheckBoxDequeuePacket"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TextBoxCommunicationWindowDCC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CheckBoxCommunicationWindowForDCC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Communication Window for CMRI"
      TabPicture(1)   =   "MainScreen.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CheckBoxCommunicationWindowForCMRI"
      Tab(1).Control(1)=   "TextBoxCommunicationWindowCMRI"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Communication Window for OTHER"
      TabPicture(2)   =   "MainScreen.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CheckBoxCommunicationWindowForOTHER"
      Tab(2).Control(1)=   "TextBoxCommunicationWindowOTHER"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox CheckBoxCommunicationWindowForOTHER 
         Caption         =   "Record Contents of Communication Window for OTHER"
         Height          =   210
         Left            =   -74820
         TabIndex        =   61
         Top             =   3720
         Width           =   4575
      End
      Begin VB.CheckBox CheckBoxCommunicationWindowForCMRI 
         Caption         =   "Record Contents of Communication Window for CMRI"
         Height          =   210
         Left            =   -74820
         TabIndex        =   60
         Top             =   3720
         Width           =   4275
      End
      Begin VB.CheckBox CheckBoxCommunicationWindowForDCC 
         Caption         =   "Record Contents of Communication Window for DCC"
         Height          =   210
         Left            =   180
         TabIndex        =   59
         Top             =   3720
         Width           =   4275
      End
      Begin VB.TextBox TextBoxCommunicationWindowOTHER 
         Height          =   3075
         Left            =   -74820
         TabIndex        =   57
         Top             =   540
         Width           =   9375
      End
      Begin VB.TextBox TextBoxCommunicationWindowCMRI 
         Height          =   3075
         Left            =   -74820
         TabIndex        =   56
         Top             =   540
         Width           =   9375
      End
      Begin VB.TextBox TextBoxCommunicationWindowDCC 
         Height          =   3075
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   540
         Width           =   9375
      End
      Begin VB.CheckBox CheckBoxDequeuePacket 
         Caption         =   "Dequeue all packets?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7620
         TabIndex        =   40
         Top             =   3720
         Width           =   1935
      End
   End
   Begin FATHMAILOCXLib.SMTP SmtpOcx 
      Left            =   10200
      Top             =   5220
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin FATHMAILOCXLib.Message MessageOcx 
      Left            =   10140
      Top             =   4740
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   10080
      Top             =   2400
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin VB.Label LabelDescription 
      Caption         =   $"MainScreen.frx":08D8
      Height          =   615
      Left            =   720
      TabIndex        =   44
      Top             =   6480
      Width           =   9135
   End
   Begin VB.Line Line2 
      X1              =   9840
      X2              =   120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label6 
      Caption         =   "in"
      Height          =   210
      Left            =   7845
      TabIndex        =   41
      Top             =   5160
      Width           =   120
   End
   Begin MBTray.Tray SystemTrayIcon 
      Left            =   10080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "MainScreen.frx":09F4
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   10080
      Top             =   360
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MBTipOfTheDay.TipOfTheDay TipOfTheDay 
      Left            =   10080
      Top             =   3000
      _ExtentX        =   900
      _ExtentY        =   900
      BeginProperty TipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PromptFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "MainScreen.frx":0A10
   End
   Begin IniconLib.Init Ini 
      Left            =   10080
      Top             =   1800
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9840
      Y1              =   4860
      Y2              =   4860
   End
   Begin VB.Label Label3 
      Caption         =   "at"
      Height          =   210
      Left            =   6060
      TabIndex        =   38
      Top             =   5640
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "at"
      Height          =   210
      Left            =   6060
      TabIndex        =   36
      Top             =   5400
      Width           =   135
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   11280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   0   'False
      InputLen        =   1
      NullDiscard     =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   10680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   0   'False
      InputLen        =   1
      NullDiscard     =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Response"
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   10920
      Width           =   1095
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "Binary"
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Hexidecimal"
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Decimal"
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label LabelErrorByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Error Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   9915
      Width           =   1170
   End
   Begin VB.Label LabelSixthByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sixth Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   9555
      Width           =   1185
   End
   Begin VB.Label LabelFifthByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fifth Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   9195
      Width           =   1140
   End
   Begin VB.Label LabelFourthByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fourth Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   8835
      Width           =   1170
   End
   Begin VB.Label LabelThirdByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Third Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   8475
      Width           =   1200
   End
   Begin VB.Label LabelAddressByte2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   8115
      Width           =   1170
   End
   Begin VB.Label LabelAddressByte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   7755
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "at"
      Height          =   210
      Left            =   6060
      TabIndex        =   0
      Top             =   5160
      Width           =   135
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10080
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   0   'False
      NullDiscard     =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
      SThreshold      =   1
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuTransparency 
         Caption         =   "&Transparency is On"
      End
      Begin VB.Menu MenuBackgroundImage 
         Caption         =   "&Background Image is On"
      End
      Begin VB.Menu MenuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu MenuBar3 
         Caption         =   "&Utilities"
         Begin VB.Menu MenuUtilitiesCommandControl 
            Caption         =   "&Command Control"
            Enabled         =   0   'False
         End
         Begin VB.Menu MenuUtilitiesDatabases 
            Caption         =   "&Databases"
            Enabled         =   0   'False
            Begin VB.Menu MenuUtilitiesDatabaseReportingMarks 
               Caption         =   "&Reporting Marks"
               Enabled         =   0   'False
            End
            Begin VB.Menu MenuUtilitiesDatbaseSpeicifcCV 
               Caption         =   "&Specific CV Database"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu menuUtilitiesUpdateSoftware 
            Caption         =   "&Update Software"
         End
      End
      Begin VB.Menu MenuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MenuSound 
      Caption         =   "&Sound"
      WindowList      =   -1  'True
      Begin VB.Menu MenuSoundOptions 
         Caption         =   "Sound Options"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menuCommunication 
      Caption         =   "&Communications"
      Begin VB.Menu menuCommunicationDCC 
         Caption         =   "Communication to DCC is &Closed"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuCommunicationCMRI 
         Caption         =   "Communication to CMRI is &Closed"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuCommunicationOTHER 
         Caption         =   "Communication to OTHER is &Closed"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCommunicationPortSetting 
         Caption         =   "Communication Port &Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuInternetSettings 
         Caption         =   "Internet Settings"
      End
   End
   Begin VB.Menu MenuScaledTime 
      Caption         =   "Scaled &Time"
      Begin VB.Menu MenuScaledTimeSetting 
         Caption         =   "Scaled &Time Setting"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuProgrammingMode 
      Caption         =   "&Programing Mode"
      Begin VB.Menu MenuProgrammingModeDiesel 
         Caption         =   "Programming Mode &Diesel"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuProgrammingModeSteam 
         Caption         =   "Programming Mode &Steam"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuProgrammingModeRollingStock 
         Caption         =   "Programming Mode &Rolling Stock"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuProgrammingModeOther 
         Caption         =   "Programming Mode &Other"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuMainlineModeProgramming 
      Caption         =   "&Mainline Mode"
      Begin VB.Menu MenuMainlineDieselProgramming 
         Caption         =   "Mainline &Diesel Programming"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineSteamProgramming 
         Caption         =   "Mainline &Steam Programming"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineRollingStockProgramming 
         Caption         =   "Mainline &RollingStock Programming"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineOtherProgramming 
         Caption         =   "Mainline &Other"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MenuMainlineConsist 
         Caption         =   "Mainline &Consist Programming"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineScaleSpeedOperation 
         Caption         =   "Mainline Scale Speed Operation"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineOperationATC 
         Caption         =   "Mainline Operation (&ATC)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineOperationGUI 
         Caption         =   "Mainline Operation (&GUI)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuMainlineMacroMaker 
         Caption         =   "Mainline Operation Macro Maker"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuLayoutControl 
      Caption         =   "&Layout Control"
      Begin VB.Menu MenuLayoutDefineBlocks 
         Caption         =   "Define Blocks"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuLayoutCTC 
         Caption         =   "Centralized Trafic Control"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuReports 
      Caption         =   "&Reports"
      Begin VB.Menu MenuReportLocomotiveList 
         Caption         =   "Locomotive List"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuReportConfigurationVariables 
         Caption         =   "Configuration Variables List"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuAssistance 
      Caption         =   "&Assistance"
      Begin VB.Menu MenuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu menuDebugMode 
         Caption         =   "&Debug Mode is Off"
      End
      Begin VB.Menu menuBalloonHelp 
         Caption         =   "&Balloon Help is On"
      End
      Begin VB.Menu menuSpeechHelp 
         Caption         =   "&Speech Help is On"
      End
      Begin VB.Menu MenuTipOfTheDay 
         Caption         =   "&Tip of the Day is On"
      End
      Begin VB.Menu MenuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu menuPrint 
         Caption         =   "&Print Screen"
      End
      Begin VB.Menu MenuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSystemInfo 
         Caption         =   "&System Info"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MainScreen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub ButtonExpand_Click()

    If ButtonExpand.Caption = ">>" Then
        Let ButtonExpand.Caption = "<<"
        Let MainScreen.Height = 12015
    ElseIf ButtonExpand.Caption = "<<" Then
        Let ButtonExpand.Caption = ">>"
        Let MainScreen.Height = 6975
    End If
    
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
    
If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Main Screen, Form, Activate" & vbCrLf
    End If ' Debug Tag
    
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
        If TemporaryScreen = "Main Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Main Screen"
            Let TemporaryCounter = 11
        End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Check Next Item in Stack
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Next TemporaryCounter
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Screen Stack is Full
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryCounter = 10 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Main Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Main Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
    
    Call BalloonHelpUpdatePart01

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Main Screen, Form, Activate" & vbCrLf
    End If 'Darrin
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
   Let TextBoxCommunicationWindowDCC.SelStart = Len(TextBoxCommunicationWindowDCC.Text)
 

End Sub

Private Sub Form_Deactivate()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Main Screen, Form, Deactivate" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainScreen.Height)

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
                Let Ini.Value = "Main Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Main Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    MainScreen.Hide

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Main Screen, Form, Deactivate" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub


Private Sub Form_Load()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Main Screen, Form, Load" & vbCrLf
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
' Initialization of Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Type"
    Dim TemporarySoftwareType As String
    Let TemporarySoftwareType = Ini.Value
    
    If TemporarySoftwareType = "Full" Or TemporarySoftwareType = "Pro" Then
        'Let MenuLogFile.Enabled = True
        'Let MenuTransparency.Enabled = True
        'Let MenuBackgroundImage.Enabled = True
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MenuSound.Enabled = True
        Let MenuSoundOptions.Enabled = True
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let menuCommunicationDCC.Enabled = True
        Let menuCommunicationCMRI.Enabled = True
        Let menuCommunicationOTHER.Enabled = True
        Let MenuCommunicationPortSetting.Enabled = True
        Let menuInternetSettings.Enabled = True
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MenuScaledTime.Enabled = True
        Let MenuScaledTimeSetting.Enabled = False
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MenuProgrammingMode.Enabled = True
        Let MenuMainlineDieselProgramming.Enabled = False
        Let MenuMainlineSteamProgramming.Enabled = False
        Let MenuMainlineRollingStockProgramming.Enabled = False
        Let MenuMainlineOtherProgramming.Enabled = False
        Let menumainlineconsist.Enabled = False
        Let menumainlinescalespeedoperation.Enabled = False
        Let MenuMainlineOperationATC.Enabled = False
        Let MenuMainlineOperationGUI.Enabled = False
        Let MenuMainlineMacroMaker.Enabled = False
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MenuLayoutControl.Enabled = True
        Let MenuLayoutDefineBlocks.Enabled = True
        Let MenuLayoutCTC.Enabled = False
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MenuReports.Enabled = True
        Let MenuReportLocomotiveList.Enabled = True
        Let MenuReportConfigurationVariables.Enabled = True
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Let MenuAssistance.Enabled = True
        'Let MenuHelp.Enabled = True
        'Let menuDebugMode.Enabled = False
        'Let menuBalloonHelp.Enabled = True
        'Let menuspeechHelp.Enabled = True
        'Let menuTipoftheDay.Enabled = True
        'Let menuPrint.Enabled = False
        'Let MenuSystemInfo.Enabled = False
        'Let MenuAbout.Enabled = False
        
    End If
    
    Let Ini.Application = "Communication Settings Screen"
    Let Ini.Parameter = "CommunicationPortForDCC"
    Let LabelCommunicationPortSetting1.Text = Ini.Value
    Let Ini.Parameter = "BaudRateForDCC"
    Let LabelBaudRateSetting1.Text = Ini.Value
    Let Ini.Parameter = "Mode"
    Let Labelmode.Text = Ini.Value

    If Labelmode.Text <> "Standard Mode" Then
        If Labelmode.Text <> "Non-standard Mode" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the mode."
            End If
        End If
    End If
    Let Ini.Parameter = "CommunicationPortForCMRI"
    Let LabelCommunicationPortSetting2.Text = Ini.Value

    Let Ini.Parameter = "BaudRateForCMRI"
    Let LabelBaudRateSetting2.Text = Ini.Value
    Let Ini.Parameter = "CommunicationPortForOTHER"
    Let LabelCommunicationPortSetting3.Text = Ini.Value
    Let Ini.Parameter = "BaudRateForOTHER"
    Let LabelBaudRateSetting3.Text = Ini.Value
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "CheckBoxCommunicationWindowForDCC"
    Let checkboxcommunicationwindowfordcc.Value = Ini.Value
    Let Ini.Parameter = "CheckBoxCommunicationWindowForCMRI"
    Let CheckBoxCommunicationWindowForCMRI.Value = Ini.Value
    Let Ini.Parameter = "CheckBoxCommunicationWindowForOTHER"
    Let CheckBoxCommunicationWindowForOTHER.Value = Ini.Value

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        MainScreen.Left = (Screen.Width - Width) / 2
        MainScreen.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainScreen.Width > Screen.Width Then
            Let MainScreen.Left = Screen.Width - MainScreen.Width
        Else
            Let MainScreen.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainScreen.Height > Screen.Height Then
            Let MainScreen.Top = Screen.Height - MainScreen.Height
        Else
            Let MainScreen.Top = Val(TemporaryValueTop)
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Clearing Screen Stack
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Screen Stack"
    Dim TemporaryCounter As Integer
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Loop for Clearing Sceen Stack
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryCounter = 0 To 9
        Let Ini.Parameter = CStr(TemporaryCounter)
        Let Ini.Value = "Unused"
    Next TemporaryCounter
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the log file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let Ini.Application = "Main Screen"
    'Let Ini.Parameter = "LogFile"
    'Dim TemporaryLogFile As String
    'Let TemporaryLogFile = Ini.Value
    'If TemporaryLogFile = "On" Then
    '    Let MenuLogFile.Caption = "&Log File is On"
    'ElseIf TemporaryLogFile = "Off" Then
    '    Let MenuLogFile.Caption = "&Log File is Off"
    'Else
    '    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
    '        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
    '        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
    '        Let Ini.Filename = App.Path$ & "\Atc.log"
    '        Let Ini.Application = "Log Errors"
    '        Let Ini.Parameter = Date$ & " " & Time$
    '        Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'Log File' file."
    '    End If
    'End If

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the log file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "BackgroundImage"
    Dim TemporaryBackgroundImage As String
    Let TemporaryBackgroundImage = Ini.Value
    
    If TemporaryBackgroundImage = "On" Then
        Let MenuBackgroundImage.Caption = "&Background Image is On"
        Let MenuTransparency.Enabled = True
    ElseIf TemporaryBackgroundImage = "Off" Then
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "Transparency"
        Let Ini.Value = "Off"
        Let MenuTransparency.Enabled = False
        Let MenuBackgroundImage.Caption = "&Background Image is Off"
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'Background Image' file."
        End If
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the log file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "Transparency"
    Dim TemporaryTransparency As String
    Let TemporaryTransparency = Ini.Value
    If TemporaryTransparency = "On" Then
        Let MenuTransparency.Caption = "&Transparency is On"
        Let AlphaBlend.Enabled = True
    ElseIf TemporaryTransparency = "Off" Then
        Let MenuTransparency.Caption = "&Transparency is Off"
        Let AlphaBlend.Enabled = False
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'Transparency' file."
        End If
    End If
   
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the balloon help file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "DebugMode"
    Dim TemporaryDebugMode As String
    Let TemporaryDebugMode = Ini.Value
    If TemporaryDebugMode = "True" Then
        Let menuDebugMode.Caption = "&Debug Mode is On"
        DebugMode.Show vbModeless
    ElseIf TemporaryDebugMode = "False" Then
        Let menuDebugMode.Caption = "&Debug Mode is Off"
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'DebugMode' file."
        End If
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the balloon help file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "BalloonHelp"
    Dim TemporaryBalloon As String
    Let TemporaryBalloon = Ini.Value
    If TemporaryBalloon = "True" Then
        Let menuBalloonHelp.Caption = "&Balloon Help is On"
    ElseIf TemporaryBalloon = "False" Then
        Let menuBalloonHelp.Caption = "&Balloon Help is Off"
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'BalloonHelp' file."
        End If
    End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the Speech help file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "SpeechHelp"
    Dim TemporarySpeech As String
    Let TemporarySpeech = Ini.Value
    If TemporarySpeech = "True" Then
        Let menuspeechHelp.Caption = "&Speech Help is On"
    ElseIf TemporarySpeech = "False" Then
        Let menuspeechHelp.Caption = "&Speech Help is Off"
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'SpeechHelp' file."
        End If
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the tip of the day in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "TipOfTheDay"
    Dim TemporaryTipoftheDay As String
    Let TemporaryTipoftheDay = Ini.Value
    If TemporaryTipoftheDay = "True" Then
        Let TimerTipoftheDay.Interval = 65535
        Let menuTipoftheDay.Caption = "&Tip of the Day is On"
    ElseIf TemporaryTipoftheDay = "False" Then
        Let menuTipoftheDay.Caption = "&Tip of the Day is Off"
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'Tip of the Day' file."
        End If
    End If
   
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the dequeu packet in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "DequeuePacket"
    Dim TemporaryDequeuePacket As String
    Let TemporaryDequeuePacket = Ini.Value
    If TemporaryDequeuePacket = "0" Then
        Let checkboxdequeuepacket.Value = 0
    ElseIf TemporaryDequeuePacket = "1" Then
        Let checkboxdequeuepacket.Value = 1
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, an incorrect value is stored in ATC.INI for the 'DequeuePacket' file."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let TipOfTheDay.SourceFile = App.Path$ & "\Atc.tip"

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding System Icon Tray
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let SystemTrayIcon.ToolTip = About!LabelTitle.Caption
    SystemTrayIcon.Create
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send SMS
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MessageOcx.Text = " Software has been started by " & Password!TextBoxUsersEmailAddress.Text
    Let MessageOcx.Sender = "canadianlocomotivelogistics@gmail.com"
    Let MessageOcx.Recipients = "5196305494@pcs.rogers.com"
    Let MessageOcx.Subject = "Server Status"
    Let MessageOcx.ReplyTo = Password.TextBoxUsersEmailAddress.Text
    Let SmtpOcx.TimeoutMS = 6000
    Let SmtpOcx.UserName = "canadianlocomotivelogistics@gmail.com"
    Let SmtpOcx.Password = "walnuttree12"
    Let SmtpOcx.LoginMethod = AuthLoginMethod
    Let SmtpOcx.ServerAddr = "smtp.gmail.com"
    Let SmtpOcx.ServerPort = 587
    Let TemporaryValue = SmtpOcx.Send(MessageOcx.GetRaw)
    If TemporaryValue <> 0 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Main Screen Screen, Form Load, unable to send text message showing server status."
        End If
    End If

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Main Screen, Form, Load" & vbCrLf
    End If ' Debug Tag

' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub




Private Sub Form_Resize()

    If MainScreen.WindowState = vbMinimized Then
    
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
        
    ElseIf MainScreen.WindowState = vbNormal Then
    
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

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving the screen size
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainScreen.Height)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

Private Sub MenuAbout_Click()
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Load and display the new Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load About
    About.Show vbModeless
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub MenuBackgroundImage_Click()

    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Change Status of Background Image
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    
    If MenuBackgroundImage.Caption = "&Background Image is Off" Then
        Let MenuBackgroundImage.Caption = "&Background Image is On"
        Let Ini.Parameter = "BackgroundImage"
        Let Ini.Value = "On"
        'Let BackGround!ImageBoxBackGround.Width = BackGround.Width / 15
        'Let BackGround!ImageBoxBackGround.Height = BackGround.Height / 15
        'Let BackGround!ImageBoxBackGround.Visible = True
        Let BackGround.Visible = True
        Let MenuTransparency.Enabled = True
        MainScreen.Show vbModeless
    ElseIf MenuBackgroundImage.Caption = "&Background Image is On" Then
        Let MenuBackgroundImage.Caption = "&Background Image is Off"
        Let Ini.Parameter = "BackgroundImage"
        Let Ini.Value = "Off"
        'Let BackGround!ImageBoxBackGround.Width = BackGround.Width / 15
        'Let BackGround!ImageBoxBackGround.Height = BackGround.Height / 15
        'Let BackGround!ImageBoxBackGround.Visible = True
        Let BackGround.Visible = False
        If MenuTransparency.Caption = "&Transparency is On" Then
            Call MenuTransparency_Click
        End If
        Let MenuTransparency.Enabled = False
    End If

End Sub

Private Sub menuBalloonHelp_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
     
    If menuBalloonHelp.Caption = "&Balloon Help is Off" Then
        Let menuBalloonHelp.Caption = "&Balloon Help is On"
        Let Ini.Parameter = "BalloonHelp"
        Let Ini.Value = "True"
    ElseIf menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Let menuBalloonHelp.Caption = "&Balloon Help is Off"
        Let Ini.Parameter = "BalloonHelp"
        Let Ini.Value = "False"
    End If
    
    Call BalloonHelpUpdatePart01
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub menuCommunicationCMRI_Click()

If menuCommunicationCMRI.Caption = "Communication to CMRI is &Closed" Then
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Opening First Port for DCC
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If LabelCommunicationPortSetting2.Text <> "Not Used" Then
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'If Port Open
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MSComm2.PortOpen = False Then
            If LabelCommunicationPortSetting2.Text = "Communication Port One (com1:)" Then
                Let MainScreen.MSComm2.CommPort = "1"
            ElseIf LabelCommunicationPortSetting2.Text = "Communication Port Two (com2:)" Then
                Let MainScreen.MSComm2.CommPort = "2"
            ElseIf LabelCommunicationPortSetting2.Text = "Communication Port Three (com3:)" Then
                Let MainScreen.MSComm2.CommPort = "3"
            ElseIf LabelCommunicationPortSetting2.Text = "Communication Port Four (com4:)" Then
                Let MainScreen.MSComm2.CommPort = "4"
            End If
            If LabelBaudRateSetting2.Text = "9600 bits per second" Then
                Let MainScreen.MSComm2.Settings = "9600,n,8,1"
            ElseIf LabelBaudRateSetting2.Text = "19200 bits per scond" Then
                Let MainScreen.MSComm2.Settings = "19200,n,8,1"
            End If
            Let LabelStatus.Text = "Status: Opening CMRI Communication Port"
            Let MSComm2.PortOpen = True
        End If
        Let MousePointer = vbHourglass
        
        Let LabelStatus.Text = "Status: Waiting for CMRI Confirmation of Port Open"
        While Right(TextBoxCommunicationWindowDCC.Text, 1) <> Chr$(&H80)
            DoEvents
        Wend
        Let LabelStatus.Text = "Status: Clear"
        Let MousePointer = vbDefault
                
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Chnaging the Valid Option in the Menu
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
           
        Let menuCommunicationCMRI.Caption = "Communicaton to CMRI is &Open"
        Let MenuCommunicationPortSetting.Enabled = False
        'Let MenuScaledTimeSetting.Enabled = True
        'Let MenuProgrammingModeDiesel.Enabled = True
        'Let MenuProgrammingModeSteam.Enabled = True
        'Let MenuProgrammingModeRollingStock.Enabled = True
        'Let MenuProgrammingModeOther.Enabled = True
        'Let MenuMainlineDieselProgramming.Enabled = True
        'Let MenuMainlineSteamProgramming.Enabled = True
        'Let MenuMainlineRollingStockProgramming.Enabled = True
        'Let MenuMainlineOtherProgramming.Enabled = True
        'Let menumainlineconsist.Enabled = True
        'Let menumainlinescalespeedoperation.Enabled = True
        'Let MenuMainlineOperationatc.Enabled = True
        'Let MenuMainlineOperationGUI.Enabled = True
        'Let MenuMainlineMacroMaker.Enabled = True
        Let MenuLayoutDefineBlocks.Enabled = True
        Let MenuLayoutCTC.Enabled = True
        Let TextBoxCommunicationWindowCMRI.Enabled = True

    End If
Else
    
    ' -------------------------------------------------------------------------------------------------------------------------
    ' Port Open - Communication Port
    ' -------------------------------------------------------------------------------------------------------------------------
    
    Let LabelStatus.Text = "Status: Closing Communication Port"
    Let MSComm2.PortOpen = False
    
    Let menuCommunicationCMRI.Caption = "Communication to CMRI is &Closed"
    If menuCommunicationDCC.Caption = "Communication to DCC is &Open" Or _
       menuCommunicationCMRI.Caption = "Communication to CMRI is &Open" Or _
       menuCommunicationOTHER.Caption = "Communication to OTHER is &Open" Then
        Let MenuCommunicationPortSetting.Enabled = False
    Else
        Let MenuCommunicationPortSetting.Enabled = True
    End If
    'Let MenuScaledTimeSetting.Enabled = False
    'Let MenuProgrammingModeDiesel.Enabled = False
    'Let MenuProgrammingModeSteam.Enabled = False
    'Let MenuProgrammingModeRollingStock.Enabled = False
    'Let MenuProgrammingModeOther.Enabled = False
    'Let MenuMainlineDieselProgramming.Enabled = False
    'Let MenuMainlineSteamProgramming.Enabled = False
    'Let MenuMainlineRollingStockProgramming.Enabled = False
    'Let MenuMainlineOtherProgramming.Enabled = False
    'Let menumainlineconsist.Enabled = False
    'Let menumainlinescalespeedoperation.Enabled = False
    'Let MenuMainlineOperationatc.Enabled = False
    'Let MenuMainlineOperationGUI.Enabled = False
    'Let MenuMainlineMacroMaker.Enabled = False
    Let MenuLayoutDefineBlocks.Enabled = False
    Let MenuLayoutCTC.Enabled = False
    Let TextBoxCommunicationWindowCMRI.Enabled = False
    
    Let LabelStatus.Text = "Status: Clear"

End If

' -------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub menuCommunicationDCC_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Main Screen, Menu Communication DCC, Click" & vbCrLf
    End If ' Debug Mode

    Dim TemporaryInput As String
    Dim TemporaryTimeOutCounter As Long
    
    Let MousePointer = vbHourglass
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Opening Communication Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If menuCommunicationDCC.Caption = "Communication to DCC is &Closed" Then

        If LabelCommunicationPortSetting1.Text <> "Not Used" Then
            If MSComm1.PortOpen = False Then
                If LabelCommunicationPortSetting1.Text = "Communication Port One (com1:)" Then
                    Let MainScreen.MSComm1.CommPort = "1"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Two (com2:)" Then
                    Let MainScreen.MSComm1.CommPort = "2"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Three (com3:)" Then
                    Let MainScreen.MSComm1.CommPort = "3"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Four (com4:)" Then
                    Let MainScreen.MSComm1.CommPort = "4"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Five (com5:)" Then
                    Let MainScreen.MSComm1.CommPort = "5"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Six (com6:)" Then
                    Let MainScreen.MSComm1.CommPort = "6"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Seven (com7:)" Then
                    Let MainScreen.MSComm1.CommPort = "7"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Eight (com8:)" Then
                    Let MainScreen.MSComm1.CommPort = "8"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Nine (com9:)" Then
                    Let MainScreen.MSComm1.CommPort = "9"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Ten (com10:)" Then
                    Let MainScreen.MSComm1.CommPort = "10"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Eleven (com11:)" Then
                    Let MainScreen.MSComm1.CommPort = "11"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Twelve (com12:)" Then
                    Let MainScreen.MSComm1.CommPort = "12"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Thirteen (com13:)" Then
                    Let MainScreen.MSComm1.CommPort = "13"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Fourteen (com14:)" Then
                    Let MainScreen.MSComm1.CommPort = "14"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Fifteen (com15:)" Then
                    Let MainScreen.MSComm1.CommPort = "15"
                ElseIf LabelCommunicationPortSetting1.Text = "Communication Port Sixteen (com16:)" Then
                    Let MainScreen.MSComm1.CommPort = "16"
                ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Menu Communication DCC, ATC.ini file is set for unknown communication port."
                End If
            
                If LabelBaudRateSetting1.Text = "9600 bits per second" Then
                    Let MainScreen.MSComm1.Settings = "9600,n,8,1"
                ElseIf LabelBaudRateSetting1.Text = "19200 bits per second" Then
                    Let MainScreen.MSComm1.Settings = "19200,n,8,1"
                ElseIf LabelBaudRateSetting1.Text = "57600 bits per second" Then
                    Let MainScreen.MSComm1.Settings = "57600,n,8,1"
                ElseIf LabelBaudRateSetting1.Text = "14400 bits per second" Then
                    Let MainScreen.MSComm1.Settings = "1400,n,8,1"
                ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Menu Communication DCC, ATC.ini file is set for unknown baud rate."
                End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Port Open - Communication Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Let LabelStatus.Text = "Status: Opening DCC Communication Port"
                Let MSComm1.PortOpen = True
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command to DCC system
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                MSComm1.InBufferCount = 0
                MSComm1.Output = Chr$(&HAA)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                If MainScreen.Labelmode.Text = "Standard Mode" Then
                    Let TemporaryTimeOutCounter = 0
                    Do While MainScreen.MSComm1.InBufferCount < 3
                        Let TemporaryTimeOutCounter = TemporaryTimeOutCounter + 1
                        If TemporaryTimeOutCounter > 750000 Then
                            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                Let Ini.Filename = App.Path$ & "\Atc.log"
                                Let Ini.Application = "Log Errors"
                                Let Ini.Parameter = Date$ & " " & Time$
                                Let Ini.Value = "Main Screen, Menu Communication DCC, time out occured by digital command control unit after sending command."
                            End If
                            Let MSComm1.PortOpen = False
                            Let MousePointer = vbDefault
                            Let LabelStatus.Text = "Status: Digital command control unit, timed out."
                            Exit Sub
                        End If
                        Let LabelStatus.Text = "Status: Wait for response from digitial command control unit."
                        DoEvents
                    Loop
                    Let TemporaryInput = MainScreen.MSComm1.Input
                    If Len(TemporaryInput) <> 3 Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Menu Communication DCC, wrong number of bytes returned from digital command control unit."
                        End If
                    End If
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & "ROM version " & Asc(Left$(TemporaryInput, 1)) & "." & Asc(Mid$(TemporaryInput, 2, 1)) & "." & Asc(Right$(TemporaryInput, 1)) & Chr$(13) & Chr$(10)
                End If
            End If
        End If
        Let LabelStatus.Text = "Status: Communication Port to DCC Open"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Changing the Valid Option in the Menu
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let menuCommunicationDCC.Caption = "Communication to DCC is &Open"
        Let MenuCommunicationPortSetting.Enabled = False
        Let menuInternetSettings.Enabled = False
        Let MenuScaledTimeSetting.Enabled = True
        Let MenuProgrammingModeDiesel.Enabled = True
        Let MenuProgrammingModeSteam.Enabled = True
        Let MenuProgrammingModeRollingStock.Enabled = True
        Let MenuProgrammingModeOther.Enabled = True
        Let MenuMainlineDieselProgramming.Enabled = True
        Let MenuMainlineSteamProgramming.Enabled = True
        Let MenuMainlineRollingStockProgramming.Enabled = True
        Let MenuMainlineOtherProgramming.Enabled = True
        Let menumainlineconsist.Enabled = True
        Let menumainlinescalespeedoperation.Enabled = True
        Let MenuMainlineOperationATC.Enabled = True
        Let MenuMainlineOperationGUI.Enabled = True
        Let MenuMainlineMacroMaker.Enabled = True
        Let MenuUtilitiescommandcontrol.Enabled = True
        'Let MenuLayoutDefineBlocks.Enabled = True
        'Let MenuLayoutCTC.Enabled = True
        Let TextBoxCommunicationWindowDCC.Enabled = True
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Close Communication Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf menuCommunicationDCC.Caption = "Communication to DCC is &Open" Then
        Let LabelStatus.Text = "Status: Closing Communication Port"
        If LabelCommunicationPortSetting1.Text <> "Not Used" Then
            If MSComm1.PortOpen = True Then
                Let MSComm1.PortOpen = False
            End If
        End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Menu Options
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let menuCommunicationDCC.Caption = "Communication to DCC is &Closed"
        If menuCommunicationDCC.Caption = "Communication to DCC is &Open" Or _
        menuCommunicationCMRI.Caption = "Communication to CMRI is &Open" Or _
        menuCommunicationOTHER.Caption = "Communication to OTHER is &Open" Then
            Let MenuCommunicationPortSetting.Enabled = False
        Else
            Let MenuCommunicationPortSetting.Enabled = True
        End If
        Let menuInternetSettings.Enabled = True
        Let MenuScaledTimeSetting.Enabled = False
        Let MenuProgrammingModeDiesel.Enabled = False
        Let MenuProgrammingModeSteam.Enabled = False
        Let MenuProgrammingModeRollingStock.Enabled = False
        Let MenuProgrammingModeOther.Enabled = False
        Let MenuMainlineDieselProgramming.Enabled = False
        Let MenuMainlineSteamProgramming.Enabled = False
        Let MenuMainlineRollingStockProgramming.Enabled = False
        Let MenuMainlineOtherProgramming.Enabled = False
        Let menumainlineconsist.Enabled = False
        Let menumainlinescalespeedoperation.Enabled = False
        Let MenuMainlineOperationATC.Enabled = False
        Let MenuMainlineOperationGUI.Enabled = False
        Let MenuMainlineMacroMaker.Enabled = False
        Let MenuUtilitiescommandcontrol.Enabled = False
        'Let MenuLayoutDefineBlocks.Enabled = False
        'Let MenuLayoutCTC.Enabled = False
        Let TextBoxCommunicationWindowDCC.Enabled = False
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Status
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let LabelStatus.Text = "Status: Communication Port to DCC is Closed"
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Main Screen, Menu Communication DCC, programm error."
    End If
    
    Let MousePointer = vbDefault

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Main Screen, Menu Communication DCC, Click" & vbCrLf
    End If ' Debug Mode

End Sub



Private Sub MenuCommunicationPortSetting_Click()
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Loading the Communication Setting Window
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
    Load CommunicationSetting
    CommunicationSetting.Show vbModeless

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub menuDebugMode_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
     
    If menuDebugMode.Caption = "&Debug Mode is Off" Then
        Let menuDebugMode.Caption = "&Debug Mode is On"
        Let Ini.Parameter = "DebugMode"
        Let Ini.Value = "True"
        DebugMode.Show vbModeless
    ElseIf menuDebugMode.Caption = "&Debug Mode is On" Then
        Let menuDebugMode.Caption = "&Debug Mode is Off"
        Let Ini.Parameter = "DebugMode"
        Let Ini.Value = "False"
        DebugMode.Hide
    End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
'

End Sub

Private Sub MenuExit_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainScreen.Height)
    Let Ini.Parameter = "DequeuePacket"
    Let Ini.Value = checkboxdequeuepacket.Value
    Let Ini.Parameter = "CheckBoxCommunicationWindowForDCC"
    Let Ini.Value = checkboxcommunicationwindowfordcc.Value
    Let Ini.Parameter = "CheckBoxCommunicationWindowForCMRI"
    Let Ini.Value = CheckBoxCommunicationWindowForCMRI.Value
    Let Ini.Parameter = "CheckBoxCommunicationWindowForOTHER"
    Let Ini.Value = CheckBoxCommunicationWindowForOTHER.Value

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
                Let Ini.Value = "Main Screen, Meun Exit, Click, an incorrect value is stored in ATC.INI for the 'Transparency' file."
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
            Let Ini.Value = "Main Screen Screen, Menu Exit, Click, an incorrect value is stored in ATC.INI for the 'Background Image' file."
        End If
    End If

    MainScreen.Hide
      
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Terminate Program
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End

End Sub

Private Sub MenuHelp_Click()

    Let temp = "hh.exe " + App.Path$ & "\Help\atc.chm"
    Shell temp, vbNormalFocus

End Sub

Private Sub menuInternetSettings_Click()
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Loading the Internet Settings Window
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       
    Load InternetSettings
    InternetSettings.Show vbModeless

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
End Sub

Private Sub MenuLayoutCTC_Click()

    Load LayoutCTC
    LayoutCTC.Show vbModeless

End Sub

Private Sub MenuLayoutDefineBlocks_Click()

    Load DefineBlocks
    DefineBlocks.Show vbModeless

End Sub













Private Sub MenuLogFile_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    If MainScreen!MenuLogFile.Caption = "&Log File is Off" Then
        Let MenuLogFile.Caption = "&Log File is On"
        Let Ini.Parameter = "LogFile"
        Let Ini.Value = "On"
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let MenuLogFile.Caption = "&Log File is Off"
        Let Ini.Parameter = "LogFile"
        Let Ini.Value = "Off"
    End If

End Sub

Private Sub MenuMainlineConsist_Click()

    Load MainlineConsist
    MainlineConsist.Show vbModeless

End Sub

Private Sub MenuMainlineDieselProgramming_Click()

    Load MainlineDiesel
    MainlineDiesel.Show vbModeless

End Sub



Private Sub MenuMainlineMacroMaker_Click()
  
    Load MainlineMacroMaker
    MainlineMacroMaker.Show vbModeless

End Sub

Private Sub MenuMainlineOperationATC_Click()

    Load MainlineOperationATC
    MainlineOperationATC.Show vbModeless

End Sub

Private Sub MenuMainlineOperationGUI_Click()
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Main Screen, Menu Mainline Operation GUI, Click" & vbCrLf
    End If ' Debug Tag
    
    Load MainlineOperationGUI
    MainlineOperationGUI.Show vbModeless
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Main Screen, Menu Mainline Operation GUI, Click" & vbCrLf
    End If ' Debug Tag

End Sub


Private Sub MenuMainlineOtherProgramming_Click()
    
    

    Load MainlineOther
    MainlineOther.Show vbModeless

End Sub

Private Sub MenuMainlineRollingStockProgramming_Click()
    
    

    Load MainlineRollingStock
    MainlineRollingStock.Show vbModeless

End Sub


Private Sub MenuMainlineScaleSpeedOperation_Click()
    
    

    Load MainlineScaleSpeedOperation
    MainlineScaleSpeedOperation.Show vbModeless

End Sub

Private Sub MenuMainlineSteamProgramming_Click()
    
    
    
    Load MainlineSteam
    MainlineSteam.Show vbModeless

End Sub





Private Sub menuPrint_Click()

    MainScreen.PrintForm
    
End Sub

Private Sub MenuProgrammingModeDiesel_Click()
    
    
    
    Load ProgrammingDiesel
    ProgrammingDiesel.Show vbModeless

End Sub

Private Sub MenuProgrammingModeOther_Click()
    
    

    Load ProgrammingOther
    ProgrammingOther.Show vbModeless

End Sub

Private Sub MenuProgrammingModeRollingStock_Click()
    
    

    Load ProgrammingRollingStock
    ProgrammingRollingStock.Show vbModeless

End Sub

Private Sub MenuProgrammingModeSteam_Click()
    
    

    Load ProgrammingSteam
    ProgrammingSteam.Show vbModeless

End Sub

Private Sub MenuReportConfigurationVariables_Click()
    
    
    
    Load LocomotiveCVSpreadsheet
    LocomotiveCVSpreadsheet.Show vbModeless

End Sub

Private Sub MenuReportLocomotiveList_Click()
    
    
    
    Load LocomotiveSpreadsheet
    LocomotiveSpreadsheet.Show vbModeless

End Sub

Private Sub MenuScaledTimeSetting_Click()
    
    Load ClockScreen
    ClockScreen.Show vbModeless

End Sub


Private Sub MenuSoundOptions_Click()
    
    

    Load SoundScreen
    SoundScreen.Show vbModeless

End Sub

Private Sub menuSpeechHelp_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
     
    If menuspeechHelp.Caption = "&Speech Help is Off" Then
        Let menuspeechHelp.Caption = "&Speech Help is On"
        Let Ini.Parameter = "SpeechHelp"
        Let Ini.Value = "True"
        Let balloonhelp.Speech = vbChecked
    ElseIf menuspeechHelp.Caption = "&Speech Help is On" Then
        Let menuspeechHelp.Caption = "&Speech Help is Off"
        Let Ini.Parameter = "SpeechHelp"
        Let Ini.Value = "False"
        Let balloonhelp.Speech = vbUnchecked
    End If
    
    Call BalloonHelpUpdatePart01
    
End Sub

Private Sub MenuSystemInfo_Click()

    Load SystemInformation
    SystemInformation.Show vbModeless

End Sub


Private Sub menuTipoftheDay_Click()

    If menuTipoftheDay.Caption = "&Tip of the Day is Off" Then
        Let menuTipoftheDay.Caption = "&Tip of the Day is On"
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "TipOfTheDay"
        'Let TipOfTheDay.ShowAtStartUp = True
        TipOfTheDay.Show
        Let TimerTipoftheDay.Interval = 65535
    ElseIf menuTipoftheDay.Caption = "&Tip of the Day is On" Then
        Let menuTipoftheDay.Caption = "&Tip of the Day is Off"
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "TipOfTheDay"
        Let Ini.Value = "False"
        Let TimerTipoftheDay.Interval = 0
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Sub

Private Sub MenuTransparency_Click()
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    If MenuTransparency.Caption = "&Transparency is Off" Then
        Let MenuTransparency.Caption = "&Transparency is On"
        Let Ini.Parameter = "Transparency"
        Let Ini.Value = "On"
        Let Ini.Parameter = "Opacity"
        Let AlphaBlend.Opacity = Val(Ini.Value)
        Let AlphaBlend.Enabled = True
    ElseIf MenuTransparency.Caption = "&Transparency is On" Then
        Let MenuTransparency.Caption = "&Transparency is Off"
        Let Ini.Parameter = "Transparency"
        Let Ini.Value = "Off"
        Let AlphaBlend.Enabled = False
    End If
End Sub


Private Sub MenuUtilitiescommandcontrol_Click()

    Load UtilitiesForCommandControl
    UtilitiesForCommandControl.Show vbModeless
    
End Sub

Private Sub menuUtilitiesUpdateSoftware_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainScreen.Height)
    Let Ini.Parameter = "DequeuePacket"
    Let Ini.Value = checkboxdequeuepacket.Value
    
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
                Let Ini.Value = "Main Screen, Meun Exit, Click, an incorrect value is stored in ATC.INI for the 'Transparency' file."
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
            Let Ini.Value = "Main Screen Screen, Menu Exit, Click, an incorrect value is stored in ATC.INI for the 'Background Image' file."
        End If
    End If

    MainScreen.Hide
    On Error Resume Next
    FileCopy "AtcUpdater.new", "AtcUpdater.exe"
    Shell ("AtcUpdater.exe"), vbNormalFocus
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Terminate Program
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End

End Sub

Private Sub MSComm1_OnComm()
    
    If MSComm1.CommEvent = comEvSend Then
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Receive Character from Communication Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf MSComm1.CommEvent = comEvReceive Then
        'If Labelmode.text = "Standard Mode" Then
        '    While Right$(TextBoxCommunicationWindowDCC.Text, 1) <> Chr$(&H80)
        '        While MSComm1.InBufferCount <> 0
        '            Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & MSComm1.Input
        '            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
        '            ' Too Many Charaters in TextBox
        '            ' -----------------------------------------------------------------------------------------------------------------------------------------------------
        '            If Len(TextBoxCommunicationWindowDCC.Text) >= 10000 Then
        '                Let TextBoxCommunicationWindowDCC.Text = Right$(TextBoxCommunicationWindowDCC.Text, 4000)
        '            End If
        '            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
        '            ' Move Cursor to End of Text in TextBox
        '            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------
        '            Let TextBoxCommunicationWindowDCC.SelStart = Len(TextBoxCommunicationWindowDCC.Text)
        '        Wend
        '    DoEvents
        '    Wend
        'End If
'   ElseIf MSComm1.CommEvent = comEvCTS Then
'   ElseIf MSComm1.CommEvent = comEvDSR Then
'   ElseIf MSComm1.CommEvent = comEvCD Then
'   ElseIf MSComm1.CommEvent = comEvRing Then
'   ElseIf MSComm1.CommEvent = eveof Then
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


Private Sub TextBoxCommunicationWindowCMRI_Change()

    Dim tempExitEarly As Boolean
' -------------------------------------------------------------------------------------------------------------------------------------------------------
' Displaying Last 30,000 Characters
' -------------------------------------------------------------------------------------------------------------------------------------------------------
    If Len(TextBoxCommunicationWindowCMRI.Text) > 30000 Then
        Let TextBoxCommunicationWindowCMRI.Text = Right$(TextBoxCommunicationWindowCMRI.Text, 30000)
        Let tempExitEarly = True
    End If
    Let TextBoxCommunicationWindowDCC.SelStart = Len(TextBoxCommunicationWindowDCC.Text)
' -------------------------------------------------------------------------------------------------------------------------------------------------------
' Record last line to File
' -------------------------------------------------------------------------------------------------------------------------------------------------------
    If tempExitEarly = True Then Exit Sub: Rem stop duplication of event caused by truncating the text

    If CheckBoxCommunicationWindowForCMRI.Value = vbChecked Then
        If Right$(TextBoxCommunicationWindowCMRI.Text, 2) = vbCrLf Then
            On Error GoTo Adam
            Open App.Path & "\AtcCMRI.txt" For Append As #1
            For TemporaryValue = Len(TextBoxCommunicationWindowCMRI.Text) - 3 To 0 Step -1
                If Mid$(TextBoxCommunicationWindowCMRI.Text, TemporaryValue, 2) = vbCrLf Then
                    Print #1, Mid$(TextBoxCommunicationWindowCMRI.Text, TemporaryValue + 2, Len(TextBoxCommunicationWindowCMRI.Text) - 2);
                    Let TemporaryValue = -2
                End If
            Next TemporaryValue
            If TemporaryValue = -1 Then
                Print #1, TextBoxCommunicationWindowCMRI.Text;
            End If
Adam:
            Close #1
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub TextBoxCommunicationWindowDCC_Change()

    Dim tempExitEarly As Boolean
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Displaying Last 30,000 Characters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Len(TextBoxCommunicationWindowDCC.Text) > 30000 Then
        Let TextBoxCommunicationWindowDCC.Text = Right$(TextBoxCommunicationWindowDCC.Text, 30000)
        Let tempExitEarly = True
    End If
    Let TextBoxCommunicationWindowDCC.SelStart = Len(TextBoxCommunicationWindowDCC.Text)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Record last line to File
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If tempExitEarly = True Then Exit Sub: Rem stop duplication of event caused by truncating the text

    If checkboxcommunicationwindowfordcc.Value = vbChecked Then
        If Right$(TextBoxCommunicationWindowDCC.Text, 2) = vbCrLf Then
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Wright to Local File & Network File
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            On Error Resume Next
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "NetworkPath"
            
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If Ini.Value <> "Not Used" Then
                Stop
                Open Ini.Value & "\AtcDCC.txt" For Append As #1
                If Err <> 0 Then
                    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        Let Ini.Filename = App.Path$ & "\Atc.log"
                        Let Ini.Application = "Log Errors"
                        Let Ini.Parameter = Date$ & " " & Time$
                        Let Ini.Value = "Main Screen, TextBoxCommunicationWindowDCC, Change, invalid network path or filename (" & Err & "," & Error$ & ")."
                    End If
                End If
            End If
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            Open App.Path$ & "\AtcDCC.txt" For Append As #2
            If Err <> 0 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, TextBoxCommunicationWindowDCC, Change, unable to open file (" & Err & "," & Error$ & ")."
                End If
            End If
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            For TemporaryValue = Len(TextBoxCommunicationWindowDCC.Text) - 3 To 0 Step -1
                If Mid$(TextBoxCommunicationWindowDCC.Text, TemporaryValue, 2) = vbCrLf Then
                    Let Ini.Filename = App.Path$ & "\Atc.ini"
                    Let Ini.Application = "Main Screen"
                    Let Ini.Parameter = "NetworkPath"
                    If Ini.Value <> "Not Used" Then
                        Print #1, Mid$(TextBoxCommunicationWindowDCC.Text, TemporaryValue + 2, Len(TextBoxCommunicationWindowDCC.Text) - 2);
                        If Err <> 0 Then
                            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                Let Ini.Filename = App.Path$ & "\Atc.log"
                                Let Ini.Application = "Log Errors"
                                Let Ini.Parameter = Date$ & " " & Time$
                                Let Ini.Value = "Main Screen, TextBoxCommunicationWindowDCC, Change, unable to append file (" & Err & "," & Error$ & ")."
                            End If
                        End If
                    End If
                    
                    Print #2, Mid$(TextBoxCommunicationWindowDCC.Text, TemporaryValue + 2, Len(TextBoxCommunicationWindowDCC.Text) - 2);
                    If Err <> 0 Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                           Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, TextBoxCommunicationWindowDCC, Change, invalid network path in ATC.INI file (" & Err & "," & Error$ & ")."
                        End If
                    End If
                    Let TemporaryValue = -2
                End If
            Next TemporaryValue
            If TemporaryValue = -1 Then
                Print #1, TextBoxCommunicationWindowDCC.Text;
            End If
            Close #2
            Close #1
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub TextBoxCommunicationWindowOTHER_Change()

    Dim tempExitEarly As Boolean
' -------------------------------------------------------------------------------------------------------------------------------------------------------
' Displaying Last 30,000 Characters
' -------------------------------------------------------------------------------------------------------------------------------------------------------
    If Len(TextBoxCommunicationWindowOTHER.Text) > 30000 Then
        Let TextBoxCommunicationWindowOTHER.Text = Right$(TextBoxCommunicationWindowOTHER.Text, 30000)
        Let tempExitEarly = True
    End If
    Let TextBoxCommunicationWindowOTHER.SelStart = Len(TextBoxCommunicationWindowOTHER.Text)
' -------------------------------------------------------------------------------------------------------------------------------------------------------
' Record last line to File
' -------------------------------------------------------------------------------------------------------------------------------------------------------
    If tempExitEarly = True Then Exit Sub: Rem stop duplication of event caused by truncating the text

    If CheckBoxCommunicationWindowForOTHER.Value = vbChecked Then
        If Right$(TextBoxCommunicationWindowOTHER.Text, 2) = vbCrLf Then
            On Error GoTo Adam
            Open App.Path & "\AtcOTHER.txt" For Append As #1
            For TemporaryValue = Len(TextBoxCommunicationWindowOTHER.Text) - 3 To 0 Step -1
                If Mid$(TextBoxCommunicationWindowOTHER.Text, TemporaryValue, 2) = vbCrLf Then
                    Print #1, Mid$(TextBoxCommunicationWindowOTHER.Text, TemporaryValue + 2, Len(TextBoxCommunicationWindowOTHER.Text) - 2);
                    Let TemporaryValue = -2
                End If
            Next TemporaryValue
            If TemporaryValue = -1 Then
                Print #1, TextBoxCommunicationWindowOTHER.Text;
            End If
Adam:
            Close #1
            On Error GoTo 0
        End If
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






Private Sub TimerAutoStart_Timer()
    
    'Let Ini.Filename = App.Path$ & "\Atc.ini"
    'Let Ini.Application = "Main Screen"
    'If MainScreen!MenuLogFile.Caption = "&Log File is Off" Then
        'Let MenuLogFile.Caption = "&Log File is On"
        'Let Ini.Parameter = "LogFile"
        'Let Ini.Value = "On"
    'ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        'Let MenuLogFile.Caption = "&Log File is Off"
        'Let Ini.Parameter = "LogFile"
        'Let Ini.Value = "Off"
    'End If
       
    Dim TemporaryTime As Integer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  Turn Timer Off Perminently
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let TimerAutoStart.Interval = 0
    Let TimerAutoStart.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get AutoHost Preference
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "AutoStart"
    If Ini.Value = "True" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Turn on Communication Port
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call menuCommunicationDCC_Click
        For TemporaryTime = -32768 To 32766
            DoEvents
        Next TemporaryTime
        'Sleep (1000)
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Open GUI Setup Screen
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Load MainlineOperationGUI
        MainlineOperationGUI.Show vbModeless
        For TemporaryTime = -32768 To 32766
            DoEvents
        Next TemporaryTime
        'Sleep (1000)
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Open Internet Settings Screen
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Load InternetSettings
        InternetSettings.Show vbModeless
        For TemporaryTime = -32768 To 32766
            DoEvents
        Next TemporaryTime
        'Sleep (1000)
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Select Host mode
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let InternetSettings.optionHost.Value = True
        For TemporaryTime = -32768 To 32766
            DoEvents
        Next TemporaryTime
        'Sleep (1000)
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Turn Auto Listen On
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call InternetSettings.ButtonAutoListen_Click
        For TemporaryTime = -32768 To 32766
            DoEvents
        Next TemporaryTime
        'Sleep (1000)
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Start Listening for Connection
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call InternetSettings.ButtonListen_Click
    End If
End Sub

Private Sub TimerTipoftheDay_Timer()

    If TipOfTheDay.ShowAtStartUp Then
        TipOfTheDay.Show
    Else
        Let menuTipoftheDay.Caption = "&Tip of the Day is Off"
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "TipOfTheDay"
        Let Ini.Value = "False"
        Let TimerTipoftheDay.Interval = 0
    End If
End Sub



Private Sub BalloonHelpUpdatePart01()

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
        Let MainScreen.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "If there are any changes in the status of the communication port or" & vbCrLf & "data passing through the port, it will be reflected here."
        Let BalloonHelpText2 = "Status"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelStatus)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelStatus, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This tab allow you to select which communication window" & vbCrLf & "is visible. You can monitor the activity of each serial port."
        Let BalloonHelpText2 = "Communication Window Tab"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CommunicationTab)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CommunicationTab, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'Communication Tab' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This textbox is the communication window which shows" & vbCrLf & "the activity of the port to the DCC device."
        Let BalloonHelpText2 = "Communication Window for DCC device"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowDCC)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowDCC, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunicationWindowDCC' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This textbox is the communication window which shows" & vbCrLf & "the activity of the port to the CMRI device."
        Let BalloonHelpText2 = "Communication Window for CMRI device"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowCMRI)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowCMRI, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunicatiopnWindowCMRI' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This textbox is the communication window which shows" & vbCrLf & "the activity of the port to the OTHER device."
        Let BalloonHelpText2 = "Communication Window for OTHER device"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowOTHER)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowOTHER, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunicationWindowOTHER' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Click on this option to have the program dequeue" & vbCrLf & "all packets sent to the digital command control after" & vbCrLf & "queuing (sending) them."
        Let BalloonHelpText2 = "Dequeue Packets"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(checkboxdequeuepacket)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxdequeuepacket, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'checkboxdequeuepacket' object."
            End If
        End If
        
        Let BalloonHelpText1 = "If there are any notes the user needs to know about, the program will display them here."
        Let BalloonHelpText2 = "Notes"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelNotes)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelNotes, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelNotes' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "settings for your DCC device. To change this" & vbCrLf & "select 'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxCommunication1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxCommunication1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunication1' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "settings for your CMRI device. To change this" & vbCrLf & "select 'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxCommunication2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxCommunication2, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunication2' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "settings for your OTHER device. To change this" & vbCrLf & "select 'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxCommunication3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxCommunication3, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'TextboxCommunication3' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "for your DCC device. To change this select" & vbCrLf & "'Communications' from the main menu, then 'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelCommunicationPortSetting1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelCommunicationPortSetting1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelCommunicationPortSetting1' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "for your CMRI device. To change this select" & vbCrLf & "'Communications' from the main menu, then 'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelCommunicationPortSetting2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelCommunicationPortSetting2, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelCommunicationPortSettings2' object."
            End If
        End If
        Let BalloonHelpText1 = "This label displays the current cummunication port" & vbCrLf & "for your OTHER device. To change this select" & vbCrLf & "'Communications' from the main menu, then 'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelCommunicationPortSetting3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelCommunicationPortSetting3, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelCommunicationPortSetting3' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current baud rate for" & vbCrLf & "your DCC device. To change this select" & vbCrLf & "'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelBaudRateSetting1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelBaudRateSetting1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelBaudRateSetting1' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current baud rate for" & vbCrLf & "your CMRI device. To change this select" & vbCrLf & "'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelBaudRateSetting2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelBaudRateSetting2, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelBaudRateSetting2' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This label displays the current baud rate for" & vbCrLf & "your OTHER device. To change this select" & vbCrLf & "'Communications' from the main menu, then" & vbCrLf & "'Settings'."
        Let BalloonHelpText2 = "Communication Port Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelBaudRateSetting3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelBaudRateSetting3, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'LabelBaudRateSetting3' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This command button will expand or contract the Main" & vbCrLf & "Screen to show addtion objects. The additional objects" & vbCrLf & "are used by the software to send commands to the digital" & vbCrLf & "command control unit."
        Let BalloonHelpText2 = "Expand Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonExpand)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonExpand, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'ButtonExpand' object."
            End If
        End If
        
        MainScreen.MousePointer = ccDefault
        
    Else 'If menuBalloonHelp.Caption = "&Balloon Help is Off" Then
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, Balloon Help Update, unable to setup balloon help for 'ButtonExpand' object."
            End If
        End If
    End If

End Sub






Public Sub SendCommandviaTrackQ()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainscreen, General, Send Command Via TrackQ" & vbCrLf
    End If ' Debug Tag
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Speed and Direction Commands only
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryByteCount As Integer
    Dim TemporaryValue As Integer
    Dim TemporaryCommand As String
    Dim TemporaryCommandH As String
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Checksum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let sevenbyted.Text = "0"
    For X = 1 To 8
        Let TemporaryValue = 0
        If onebyted.Text <> "" Then
            If onebyted.Text <> "   " Then
                If Mid$(OneByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If twoByteD.Text <> "" Then
            If twoByteD.Text <> "   " Then
                If Mid$(TwoByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If ThreeByteD.Text <> "" Then
            If ThreeByteD.Text <> "   " Then
                If Mid$(ThreeByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If FourByteD.Text <> "" Then
            If FourByteD.Text <> "   " Then
                If Mid$(FourByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If FiveByteD.Text <> "" Then
            If FiveByteD.Text <> "   " Then
                If Mid$(FiveByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If SixByteD.Text <> "" Then
            If SixByteD.Text <> "   " Then
                If Mid$(SixByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If X = 1 Then bitvalue = 128
        If X = 2 Then bitvalue = 64
        If X = 3 Then bitvalue = 32
        If X = 4 Then bitvalue = 16
        If X = 5 Then bitvalue = 8
        If X = 6 Then bitvalue = 4
        If X = 7 Then bitvalue = 2
        If X = 8 Then bitvalue = 1
        If Int(TemporaryValue / 2) <> (TemporaryValue / 2) Then sevenbyted.Text = Val(sevenbyted.Text) + bitvalue
    Next X
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Assemble the TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelStatus.Text = "Status: Sending Command"
    Let TemporaryCommand = ""
    Let TemporaryCommandH = ""
    Let TemporaryByteCount = 0
    If onebyted.Text <> "" Then
        If onebyted.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(onebyted.Text))
            Let TemporaryCommandH = TemporaryCommandH & OneByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If twoByteD.Text <> "" Then
        If twoByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(twoByteD.Text))
            Let TemporaryCommandH = TemporaryCommandH & TwoByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
       End If
    End If
    If ThreeByteD.Text <> "" Then
        If ThreeByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(ThreeByteD.Text))
            Let TemporaryCommandH = TemporaryCommandH & ThreeByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If FourByteD.Text <> "" Then
        If FourByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(FourByteD.Text))
            Let TemporaryCommandH = TemporaryCommandH & FourByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If FiveByteD.Text <> "" Then
        If FiveByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(FiveByteD.Text))
            Let TemporaryCommandH = TemporaryCommandH & FiveByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
       End If
    End If
    If SixByteD.Text <> "" Then
        If SixByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(SixByteD.Text))
            Let TemporaryCommandH = TemporaryCommandH & SixByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If sevenbyted.Text <> "" Then
        If sevenbyted.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(sevenbyted.Text))
            Let TemporaryCommandH = TemporaryCommandH & SevenByteH.Text
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Finish Assembly of TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainScreen.MSComm1.InBufferCount = 0
    If TemporaryByteCount = 3 Then
        If LabelCommunicationPortSetting1.Text <> "Not Used" Then
            If CommunicationSetting.optiondccsystemnce.Value = True Then
                Let MainScreen.MSComm1.Output = Chr$(&HA3) & TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0xA3" & TemporaryCommandH
            ElseIf CommunicationSetting.optiondccsystemDigitrax.Value = True Then
                Let MainScreen.MSComm1.Output = TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0x" & TemporaryCommandH
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, SendCommand, invalid value for Option DCC System."
            End If
        End If
    ElseIf TemporaryByteCount = 4 Then
        If LabelCommunicationPortSetting1.Text <> "Not Used" Then
            If CommunicationSetting.optiondccsystemnce.Value = True Then
                Let MainScreen.MSComm1.Output = Chr$(&HA4) & TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0xA4" & TemporaryCommandH
            ElseIf CommunicationSetting.optiondccsystemDigitrax.Value = True Then
                Let MainScreen.MSComm1.Output = TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0x" & TemporaryCommandH
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, SendCommand, invalid value for Option DCC System."
            End If
        End If
    ElseIf TemporaryByteCount = 5 Then
        If LabelCommunicationPortSetting1.Text <> "Not Used" Then
            If CommunicationSetting.optiondccsystemnce.Value = True Then
                Let MainScreen.MSComm1.Output = Chr$(&HA5) & TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0xA5" & TemporaryCommandH
            ElseIf CommunicationSetting.optiondccsystemDigitrax.Value = True Then
                Let MainScreen.MSComm1.Output = TemporaryCommand
                Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " 0x" & TemporaryCommandH
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, SendCommand, invalid value for Option DCC System."
            End If
        End If
        
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Main Screen, SendCommand, invalid value for byte counter."
    End If
        
    TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " at " & Date$ & " " & Time$ & vbCrLf
    
    Let LabelStatus.Text = "Status: Command Sent"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If LabelCommunicationPortSetting1.Text <> "Not Used" Then
        If MainScreen.Labelmode.Text = "Standard Mode" Then
            Do While MainScreen.MSComm1.InBufferCount < 1
                Let LabelStatus.Text = "Status: Wait for response from digitial command control unit."
                DoEvents
            Loop
            Let TemporaryInput = MainScreen.MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, wrong number of bytes returned from digital command control unit."
                End If
            End If
            If TemporaryInput = "1" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, locomotive address out of range when queing packet."
                End If
            ElseIf TemporaryInput = "2" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, cab address out of range when queing packet."
                End If
            ElseIf TemporaryInput = "3" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, data out of range when queing packet."
                End If
            ElseIf TemporaryInput = "4" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, byte count out range when queing packet."
                End If
            ElseIf TemporaryInput <> "!" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Main Screen, Send Command Via Track Q, unknown return value by command control unit when queing packet."
                End If
            End If
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Dequeuing TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.checkboxdequeuepacket.Value = vbChecked Then
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Checksum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let sevenbyted.Text = "0"
        For X = 1 To 8
            If onebyted.Text <> "" Then
                If onebyted.Text <> "   " Then
                    If Mid$(OneByteB.Text, X, 1) = "1" Then
                        Let TemporaryValue = TemporaryValue + 1
                    End If
                End If
            End If
            If twoByteD.Text <> "" Then
                If twoByteD.Text <> "   " Then
                    If Mid$(TwoByteB.Text, X, 1) = "1" Then
                        Let TemporaryValue = TemporaryValue + 1
                    End If
                End If
            End If
'            If ThreeByteD.Text <> "" Then
'                If ThreeByteD.Text <> "   " Then
'                    If Mid$(ThreeByteB.Text, X, 1) = "1" Then
'                        Let TemporaryValue = TemporaryValue + 1
'                    End If
'                End If
'            End If
'            If FourByteD.Text <> "" Then
'                If FourByteD.Text <> "   " Then
'                    If Mid$(FourByteB.Text, X, 1) = "1" Then
'                        Let TemporaryValue = TemporaryValue + 1
'                    End If
'                End If
'            End If
'            If FiveByteD.Text <> "" Then
'                If FiveByteD.Text <> "   " Then
'                    If Mid$(FiveByteB.Text, X, 1) = "1" Then
'                        Let TemporaryValue = TemporaryValue + 1
'                    End If
'                End If
'            End If
'            If SixByteD.Text <> "" Then
'                If SixByteD.Text <> "   " Then
'                    If Mid$(SixByteB.Text, X, 1) = "1" Then
'                        Let TemporaryValue = TemporaryValue + 1
'                    End If
'                End If
'            End If
            If X = 1 Then bitvalue = 128
            If X = 2 Then bitvalue = 64
            If X = 3 Then bitvalue = 32
            If X = 4 Then bitvalue = 16
            If X = 5 Then bitvalue = 8
            If X = 6 Then bitvalue = 4
            If X = 7 Then bitvalue = 2
            If X = 8 Then bitvalue = 1
            If Int(TemporaryValue / 2) <> (TemporaryValue / 2) Then sevenbyted.Text = Val(sevenbyted.Text) + bitvalue
        Next X
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Assemble the TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let LabelStatus.Text = "Status: Sending Command"
        Let TemporaryCommand = ""
'        Let TemporaryByteCount = 0
            If onebyted.Text <> "" Then
                If onebyted.Text <> "   " Then
                    Let TemporaryCommand = TemporaryCommand & Chr$(Val(onebyted.Text))
'                   Let TemporaryByteCount = TemporaryByteCount + 1
                End If
            End If
            If twoByteD.Text <> "" Then
                If twoByteD.Text <> "   " Then
                    Let TemporaryCommand = TemporaryCommand & Chr$(Val(twoByteD.Text))
'                    Let TemporaryByteCount = TemporaryByteCount + 1
               End If
            End If
'            If ThreeByteD.Text <> "" Then
'                If ThreeByteD.Text <> "   " Then
'                    Let TemporaryCommand = TemporaryCommand & Chr$(val(ThreeByteD.Text))
'                    Let TemporaryByteCount = TemporaryByteCount + 1
'                End If
'            End If
'            If FourByteD.Text <> "" Then
'                If FourByteD.Text <> "   " Then
'                    Let TemporaryCommand = TemporaryCommand & Chr$(val(FourByteD.Text))
'                    Let TemporaryByteCount = TemporaryByteCount + 1
'                End If
'            End If
'            If FiveByteD.Text <> "" Then
'                If FiveByteD.Text <> "   " Then
'                    Let TemporaryCommand = TemporaryCommand & Chr$(val(FiveByteD.Text))
'                    Let TemporaryByteCount = TemporaryByteCount + 1
'               End If
'            End If
'            If SixByteD.Text <> "" Then
'                If SixByteD.Text <> "   " Then
'                    Let TemporaryCommand = TemporaryCommand & Chr$(val(SixByteD.Text))
'                    Let TemporaryByteCount = TemporaryByteCount + 1
'                End If
'            End If
'            If sevenbyted.Text <> "" Then
'                If sevenbyted.Text <> "   " Then
'                    Let TemporaryCommand = TemporaryCommand & Chr$(val(SevenByteH.Text))
'                    'Let TemporaryByteCount = TemporaryByteCount + 1
'                End If
'            End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Finish Assembly of TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainScreen.MSComm1.InBufferCount = 0
        If LabelCommunicationPortSetting1.Text <> "No Used" Then
            If CommunicationSetting.optiondccsystemnce.Value = True Then
                Let MainScreen.MSComm1.Output = Chr$(&H88) & TemporaryCommand
            End If
            Let LabelStatus.Text = "Status: Command Sent"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If LabelCommunicationPortSetting1.Text <> "Not Used" Then
                If MainScreen.Labelmode.Text = "Standard Mode" Then
                    Do While MainScreen.MSComm1.InBufferCount < 1
                        Let LabelStatus.Text = "Status: Wait for response from digitial command control unit."
                        DoEvents
                    Loop
                    Let TemporaryInput = MainScreen.MSComm1.Input
                    If Len(TemporaryInput) <> 1 Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, wrong number of bytes returned from digital command control unit when dequeing packet."
                        End If
                    End If
                    If TemporaryInput = "1" Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, locomotive address out of range when dequeing packet."
                        End If
                    ElseIf TemporaryInput = "2" Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, cab address out of range when dequeing packet."
                        End If
                    ElseIf TemporaryInput = "3" Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, data out of range when dequeing packet."
                        End If
                    ElseIf TemporaryInput = "4" Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, byte count out range when dequeing packet."
                        End If
                    ElseIf TemporaryInput <> "!" Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Main Screen, Send Command Via Track Q, unknown return value by command control unit when dequeing packet."
                        End If
                    End If
                End If
            End If
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainscreen, General, Send Command Via TrackQ" & vbCrLf
    End If ' Debug Mode

End Sub

Public Sub SendCommandviaTempQ()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainscreen, General, Send Command Via TempQ" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Speed and Direction Commands only
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryByteCount As Integer
    Dim TemporaryValue As Integer
    Dim TemporaryCommand As String
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Checksum
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let sevenbyted.Text = "0"
    For X = 1 To 8
        Let TemporaryValue = 0
        If onebyted.Text <> "" Then
            If onebyted.Text <> "   " Then
                If Mid$(OneByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If twoByteD.Text <> "" Then
            If twoByteD.Text <> "   " Then
                If Mid$(TwoByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If ThreeByteD.Text <> "" Then
            If ThreeByteD.Text <> "   " Then
                If Mid$(ThreeByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If FourByteD.Text <> "" Then
            If FourByteD.Text <> "   " Then
                If Mid$(FourByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If FiveByteD.Text <> "" Then
            If FiveByteD.Text <> "   " Then
                If Mid$(FiveByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If SixByteD.Text <> "" Then
            If SixByteD.Text <> "   " Then
                If Mid$(SixByteB.Text, X, 1) = "1" Then
                    Let TemporaryValue = TemporaryValue + 1
                End If
            End If
        End If
        If X = 1 Then bitvalue = 128
        If X = 2 Then bitvalue = 64
        If X = 3 Then bitvalue = 32
        If X = 4 Then bitvalue = 16
        If X = 5 Then bitvalue = 8
        If X = 6 Then bitvalue = 4
        If X = 7 Then bitvalue = 2
        If X = 8 Then bitvalue = 1
        If Int(TemporaryValue / 2) <> (TemporaryValue / 2) Then sevenbyted.Text = Val(sevenbyted.Text) + bitvalue
    Next X
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Assemble the TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelStatus.Text = "Status: Sending Command"
    Let TemporaryCommand = ""
    Let TemporaryByteCount = 0
    If onebyted.Text <> "" Then
        If onebyted.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(onebyted.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If twoByteD.Text <> "" Then
        If twoByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(twoByteD.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
       End If
    End If
    If ThreeByteD.Text <> "" Then
        If ThreeByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(ThreeByteD.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If FourByteD.Text <> "" Then
        If FourByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(FourByteD.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If FiveByteD.Text <> "" Then
        If FiveByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(FiveByteD.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
       End If
    End If
    If SixByteD.Text <> "" Then
        If SixByteD.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(SixByteD.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
    If sevenbyted.Text <> "" Then
        If sevenbyted.Text <> "   " Then
            Let TemporaryCommand = TemporaryCommand & Chr$(Val(SevenByteH.Text))
            Let TemporaryByteCount = TemporaryByteCount + 1
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Finish Assembly of TemporaryCommand
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainScreen!MSComm1.InBufferCount = 0
    If TemporaryByteCount = 3 Then
        Let MainScreen!MSComm1.Output = Chr$(&H93) & Chr$(5) & TemporaryCommand
    ElseIf TemporaryByteCount = 4 Then
        Let MainScreen!MSComm1.Output = Chr$(&H94) & Chr$(5) & TemporaryCommand
    ElseIf TemporaryByteCount = 5 Then
        Let MainScreen!MSComm1.Output = Chr$(&H95) & Chr$(5) & TemporaryCommand
    ElseIf TemporaryByteCount = 6 Then
        Let MainScreen!MSComm1.Output = Chr$(&H96) & Chr$(5) & TemporaryCommand
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Main Screen, Send Command Via Temporary Q, invalid value for byte counter."
    End If
    Let LabelStatus.Text = "Status: Command Sent"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.Labelmode.Text = "Standard Mode" Then
        Do While MainScreen.MSComm1.InBufferCount < 1
            Let LabelStatus.Text = "Status: Wait for response from digitial command control unit."
            DoEvents
        Loop
        Let TemporaryInput = MainScreen.MSComm1.Input
        If Len(TemporaryInput) <> 1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, wrong number of bytes returned from digital command control unit."
            End If
        End If
        If TemporaryInput = "1" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, locomotive address out of range when queing packet."
            End If
        ElseIf TemporaryInput = "2" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, cab address out of range when queing packet."
            End If
        ElseIf TemporaryInput = "3" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, data out of range when queing packet."
            End If
        ElseIf TemporaryInput = "4" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, byte count out range when queing packet."
            End If
        ElseIf TemporaryInput <> "!" Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen, Send Command Via Track Q, unknown return value by command control unit when queing packet."
            End If
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainscreen, General, Send Command Via TempQ" & vbCrLf
    End If ' Debug Tag

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


