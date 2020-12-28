VERSION 4.00
Begin VB.Form MainlineOperationATC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - ATC"
   ClientHeight    =   6720
   ClientLeft      =   945
   ClientTop       =   1905
   ClientWidth     =   9945
   Height          =   7125
   Icon            =   "MainlineOperationATC.frx":0000
   Left            =   885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9945
   Top             =   1560
   Width           =   10065
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   8640
      TabIndex        =   235
      Top             =   6080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   180
      Picture         =   "MainlineOperationATC.frx":0442
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   234
      Top             =   135
      Width           =   540
   End
   Begin VB.HScrollBar TimerIntervalSetting 
      Height          =   255
      LargeChange     =   100
      Left            =   7560
      TabIndex        =   231
      Top             =   2520
      Value           =   1000
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7320
      Top             =   120
   End
   Begin VB.CheckBox CommunicationPortBusy 
      Caption         =   "Communication Port Busy"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   224
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CheckBox CommunicationPortBusy 
      Caption         =   "Communication Port Busy"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   223
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CheckBox CommunicationPortBusy 
      Caption         =   "Communication Port Busy"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   222
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CheckBox CommunicationPortBusy 
      Caption         =   "Communication Port Busy"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   221
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CheckBox CommunicationPortBusy 
      Caption         =   "Communication Port Busy"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   220
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   8640
      TabIndex        =   44
      Top             =   6360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ATC #0"
      TabPicture(0)   =   "MainlineOperationATC.frx":0884
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label12(0)"
      Tab(0).Control(1)=   "Label14(0)"
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(3)=   "Label10(0)"
      Tab(0).Control(4)=   "Label9(0)"
      Tab(0).Control(5)=   "Label8(0)"
      Tab(0).Control(6)=   "Label7(0)"
      Tab(0).Control(7)=   "Label6(0)"
      Tab(0).Control(8)=   "Label5(0)"
      Tab(0).Control(9)=   "Label4(0)"
      Tab(0).Control(10)=   "Label3(0)"
      Tab(0).Control(11)=   "Label2(0)"
      Tab(0).Control(12)=   "Label1(0)"
      Tab(0).Control(13)=   "LocomotiveCommunicationStatus(0)"
      Tab(0).Control(14)=   "Text1(0)"
      Tab(0).Control(15)=   "LocomotiveDatabaseConsistControl(0)"
      Tab(0).Control(16)=   "Frame1(0)"
      Tab(0).Control(17)=   "TimeDelay(0)"
      Tab(0).Control(18)=   "LocomotiveDatabaseConsistControlLoop(0)"
      Tab(0).Control(19)=   "ButtonStart(0)"
      Tab(0).Control(20)=   "SevenByteD(0)"
      Tab(0).Control(21)=   "SixByteD(0)"
      Tab(0).Control(22)=   "SevenByteB(0)"
      Tab(0).Control(23)=   "SevenByteH(0)"
      Tab(0).Control(24)=   "SixByteB(0)"
      Tab(0).Control(25)=   "SixByteH(0)"
      Tab(0).Control(26)=   "FiveByteB(0)"
      Tab(0).Control(27)=   "FiveByteH(0)"
      Tab(0).Control(28)=   "FiveByteD(0)"
      Tab(0).Control(29)=   "FourByteB(0)"
      Tab(0).Control(30)=   "FourByteH(0)"
      Tab(0).Control(31)=   "FourByteD(0)"
      Tab(0).Control(32)=   "ThreeByteB(0)"
      Tab(0).Control(33)=   "ThreeByteH(0)"
      Tab(0).Control(34)=   "ThreeByteD(0)"
      Tab(0).Control(35)=   "TwoByteB(0)"
      Tab(0).Control(36)=   "TwoByteH(0)"
      Tab(0).Control(37)=   "TwoByteD(0)"
      Tab(0).Control(38)=   "OneByteB(0)"
      Tab(0).Control(39)=   "OneByteH(0)"
      Tab(0).Control(40)=   "OneByteD(0)"
      Tab(0).Control(41)=   "LocomotiveConsistDefault(0)"
      Tab(0).Control(42)=   "ButtonStop(0)"
      Tab(0).Control(43)=   "TimeLeft(0)"
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "ATC #1"
      TabPicture(1)   =   "MainlineOperationATC.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(3)=   "Label2(1)"
      Tab(1).Control(4)=   "Label12(1)"
      Tab(1).Control(5)=   "Label14(1)"
      Tab(1).Control(6)=   "Label10(1)"
      Tab(1).Control(7)=   "Label9(1)"
      Tab(1).Control(8)=   "Label8(1)"
      Tab(1).Control(9)=   "Label7(1)"
      Tab(1).Control(10)=   "Label6(1)"
      Tab(1).Control(11)=   "Label5(1)"
      Tab(1).Control(12)=   "LocomotiveCommunicationStatus(1)"
      Tab(1).Control(13)=   "OneByteD(1)"
      Tab(1).Control(14)=   "TwoByteD(1)"
      Tab(1).Control(15)=   "ThreeByteD(1)"
      Tab(1).Control(16)=   "FourByteD(1)"
      Tab(1).Control(17)=   "FiveByteD(1)"
      Tab(1).Control(18)=   "SixByteD(1)"
      Tab(1).Control(19)=   "SevenByteD(1)"
      Tab(1).Control(20)=   "TimeDelay(1)"
      Tab(1).Control(21)=   "Text1(1)"
      Tab(1).Control(22)=   "OneByteH(1)"
      Tab(1).Control(23)=   "TwoByteH(1)"
      Tab(1).Control(24)=   "ThreeByteH(1)"
      Tab(1).Control(25)=   "FourByteH(1)"
      Tab(1).Control(26)=   "FiveByteH(1)"
      Tab(1).Control(27)=   "SixByteH(1)"
      Tab(1).Control(28)=   "SevenByteH(1)"
      Tab(1).Control(29)=   "OneByteB(1)"
      Tab(1).Control(30)=   "TwoByteB(1)"
      Tab(1).Control(31)=   "ThreeByteB(1)"
      Tab(1).Control(32)=   "FourByteB(1)"
      Tab(1).Control(33)=   "FiveByteB(1)"
      Tab(1).Control(34)=   "SixByteB(1)"
      Tab(1).Control(35)=   "SevenByteB(1)"
      Tab(1).Control(36)=   "LocomotiveDatabaseConsistControlLoop(1)"
      Tab(1).Control(37)=   "Frame1(1)"
      Tab(1).Control(38)=   "ButtonStart(1)"
      Tab(1).Control(39)=   "ButtonStop(1)"
      Tab(1).Control(40)=   "LocomotiveDatabaseConsistControl(1)"
      Tab(1).Control(41)=   "LocomotiveConsistDefault(1)"
      Tab(1).Control(42)=   "TimeLeft(1)"
      Tab(1).ControlCount=   43
      TabCaption(2)   =   "ATC #2"
      TabPicture(2)   =   "MainlineOperationATC.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3(2)"
      Tab(2).Control(1)=   "Label4(2)"
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(3)=   "Label1(2)"
      Tab(2).Control(4)=   "Label12(2)"
      Tab(2).Control(5)=   "Label14(2)"
      Tab(2).Control(6)=   "Label10(2)"
      Tab(2).Control(7)=   "Label9(2)"
      Tab(2).Control(8)=   "Label8(2)"
      Tab(2).Control(9)=   "Label7(2)"
      Tab(2).Control(10)=   "Label6(2)"
      Tab(2).Control(11)=   "Label5(2)"
      Tab(2).Control(12)=   "LocomotiveCommunicationStatus(2)"
      Tab(2).Control(13)=   "OneByteD(2)"
      Tab(2).Control(14)=   "TwoByteD(2)"
      Tab(2).Control(15)=   "ThreeByteD(2)"
      Tab(2).Control(16)=   "FourByteD(2)"
      Tab(2).Control(17)=   "FiveByteD(2)"
      Tab(2).Control(18)=   "SixByteD(2)"
      Tab(2).Control(19)=   "SevenByteD(2)"
      Tab(2).Control(20)=   "TimeDelay(2)"
      Tab(2).Control(21)=   "Text1(2)"
      Tab(2).Control(22)=   "OneByteH(2)"
      Tab(2).Control(23)=   "TwoByteH(2)"
      Tab(2).Control(24)=   "ThreeByteH(2)"
      Tab(2).Control(25)=   "FourByteH(2)"
      Tab(2).Control(26)=   "FiveByteH(2)"
      Tab(2).Control(27)=   "SixByteH(2)"
      Tab(2).Control(28)=   "SevenByteH(2)"
      Tab(2).Control(29)=   "OneByteB(2)"
      Tab(2).Control(30)=   "TwoByteB(2)"
      Tab(2).Control(31)=   "ThreeByteB(2)"
      Tab(2).Control(32)=   "FourByteB(2)"
      Tab(2).Control(33)=   "FiveByteB(2)"
      Tab(2).Control(34)=   "SixByteB(2)"
      Tab(2).Control(35)=   "SevenByteB(2)"
      Tab(2).Control(36)=   "LocomotiveDatabaseConsistControlLoop(2)"
      Tab(2).Control(37)=   "Frame1(2)"
      Tab(2).Control(38)=   "ButtonStart(2)"
      Tab(2).Control(39)=   "ButtonStop(2)"
      Tab(2).Control(40)=   "LocomotiveDatabaseConsistControl(2)"
      Tab(2).Control(41)=   "LocomotiveConsistDefault(2)"
      Tab(2).Control(42)=   "TimeLeft(2)"
      Tab(2).ControlCount=   43
      TabCaption(3)   =   "ATC #3"
      TabPicture(3)   =   "MainlineOperationATC.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3(3)"
      Tab(3).Control(1)=   "Label4(3)"
      Tab(3).Control(2)=   "Label2(3)"
      Tab(3).Control(3)=   "Label12(3)"
      Tab(3).Control(4)=   "Label14(3)"
      Tab(3).Control(5)=   "Label10(3)"
      Tab(3).Control(6)=   "Label9(3)"
      Tab(3).Control(7)=   "Label8(3)"
      Tab(3).Control(8)=   "Label7(3)"
      Tab(3).Control(9)=   "Label6(3)"
      Tab(3).Control(10)=   "Label5(3)"
      Tab(3).Control(11)=   "Label1(3)"
      Tab(3).Control(12)=   "LocomotiveCommunicationStatus(3)"
      Tab(3).Control(13)=   "OneByteD(3)"
      Tab(3).Control(14)=   "TwoByteD(3)"
      Tab(3).Control(15)=   "ThreeByteD(3)"
      Tab(3).Control(16)=   "FourByteD(3)"
      Tab(3).Control(17)=   "FiveByteD(3)"
      Tab(3).Control(18)=   "SixByteD(3)"
      Tab(3).Control(19)=   "SevenByteD(3)"
      Tab(3).Control(20)=   "TimeDelay(3)"
      Tab(3).Control(21)=   "Text1(3)"
      Tab(3).Control(22)=   "OneByteH(3)"
      Tab(3).Control(23)=   "TwoByteH(3)"
      Tab(3).Control(24)=   "ThreeByteH(3)"
      Tab(3).Control(25)=   "FourByteH(3)"
      Tab(3).Control(26)=   "FiveByteH(3)"
      Tab(3).Control(27)=   "SixByteH(3)"
      Tab(3).Control(28)=   "SevenByteH(3)"
      Tab(3).Control(29)=   "OneByteB(3)"
      Tab(3).Control(30)=   "TwoByteB(3)"
      Tab(3).Control(31)=   "ThreeByteB(3)"
      Tab(3).Control(32)=   "FourByteB(3)"
      Tab(3).Control(33)=   "FiveByteB(3)"
      Tab(3).Control(34)=   "SixByteB(3)"
      Tab(3).Control(35)=   "SevenByteB(3)"
      Tab(3).Control(36)=   "LocomotiveDatabaseConsistControlLoop(3)"
      Tab(3).Control(37)=   "Frame1(3)"
      Tab(3).Control(38)=   "ButtonStart(3)"
      Tab(3).Control(39)=   "ButtonStop(3)"
      Tab(3).Control(40)=   "LocomotiveDatabaseConsistControl(3)"
      Tab(3).Control(41)=   "LocomotiveConsistDefault(3)"
      Tab(3).Control(42)=   "TimeLeft(3)"
      Tab(3).ControlCount=   43
      TabCaption(4)   =   "ATC #4"
      TabPicture(4)   =   "MainlineOperationATC.frx":08F4
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label3(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label4(4)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label2(4)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label12(4)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label14(4)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label10(4)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label9(4)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label8(4)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label7(4)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label6(4)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label5(4)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label1(4)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "LocomotiveCommunicationStatus(4)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "OneByteD(4)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "TwoByteD(4)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "ThreeByteD(4)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "FourByteD(4)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "FiveByteD(4)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "SixByteD(4)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "SevenByteD(4)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "TimeDelay(4)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Text1(4)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "OneByteH(4)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "TwoByteH(4)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "ThreeByteH(4)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "FourByteH(4)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "FiveByteH(4)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "SixByteH(4)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "SevenByteH(4)"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "OneByteB(4)"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "TwoByteB(4)"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "ThreeByteB(4)"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).Control(32)=   "FourByteB(4)"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "FiveByteB(4)"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "SixByteB(4)"
      Tab(4).Control(34).Enabled=   0   'False
      Tab(4).Control(35)=   "SevenByteB(4)"
      Tab(4).Control(35).Enabled=   0   'False
      Tab(4).Control(36)=   "LocomotiveDatabaseConsistControlLoop(4)"
      Tab(4).Control(36).Enabled=   0   'False
      Tab(4).Control(37)=   "Frame1(4)"
      Tab(4).Control(37).Enabled=   0   'False
      Tab(4).Control(38)=   "ButtonStart(4)"
      Tab(4).Control(38).Enabled=   0   'False
      Tab(4).Control(39)=   "ButtonStop(4)"
      Tab(4).Control(39).Enabled=   0   'False
      Tab(4).Control(40)=   "LocomotiveDatabaseConsistControl(4)"
      Tab(4).Control(40).Enabled=   0   'False
      Tab(4).Control(41)=   "LocomotiveConsistDefault(4)"
      Tab(4).Control(41).Enabled=   0   'False
      Tab(4).Control(42)=   "TimeLeft(4)"
      Tab(4).Control(42).Enabled=   0   'False
      Tab(4).ControlCount=   43
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   230
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   229
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   228
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   227
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   226
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -69240
         TabIndex        =   219
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox LocomotiveConsistDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default to Consist Address"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   218
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox LocomotiveConsistDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default to Consist Address"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   255
         Index           =   3
         Left            =   -71760
         TabIndex        =   217
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox LocomotiveConsistDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default to Consist Address"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   255
         Index           =   2
         Left            =   -71760
         TabIndex        =   216
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox LocomotiveConsistDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default to Consist Address"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   255
         Index           =   1
         Left            =   -71760
         TabIndex        =   215
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox LocomotiveConsistDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default to Consist Address"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   255
         Index           =   0
         Left            =   -71760
         TabIndex        =   214
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   4
         Left            =   5760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomoriveDatabaseConsist"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   3
         Left            =   -69240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomoriveDatabaseConsist"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   2
         Left            =   -69240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomoriveDatabaseConsist"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -69240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomoriveDatabaseConsist"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   208
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   -69240
         TabIndex        =   207
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   -69240
         TabIndex        =   206
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -69240
         TabIndex        =   205
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   204
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   -69240
         TabIndex        =   203
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   -69240
         TabIndex        =   202
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -69240
         TabIndex        =   201
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   4
         Left            =   240
         TabIndex        =   198
         Top             =   4920
         Width           =   6855
         Begin VB.CommandButton ButtonControlFile 
            Caption         =   "Control File"
            Height          =   255
            Index           =   4
            Left            =   5520
            TabIndex        =   200
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox LocomotiveDatabaseConsistControlFileName 
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   199
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   3
         Left            =   -74760
         TabIndex        =   195
         Top             =   4920
         Width           =   6855
         Begin VB.CommandButton ButtonControlFile 
            Caption         =   "Control File"
            Height          =   255
            Index           =   3
            Left            =   5520
            TabIndex        =   197
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox LocomotiveDatabaseConsistControlFileName 
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   2
         Left            =   -74760
         TabIndex        =   192
         Top             =   4920
         Width           =   6855
         Begin VB.CommandButton ButtonControlFile 
            Caption         =   "Control File"
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   194
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox LocomotiveDatabaseConsistControlFileName 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   193
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   1
         Left            =   -74760
         TabIndex        =   189
         Top             =   4920
         Width           =   6855
         Begin VB.CommandButton ButtonControlFile 
            Caption         =   "Control File"
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   191
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox LocomotiveDatabaseConsistControlFileName 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   188
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   187
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   186
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   185
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   184
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   183
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   182
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   181
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   180
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   179
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   178
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   177
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   176
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   175
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   174
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   173
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   172
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   171
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   170
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   169
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   168
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   167
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   166
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   165
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   164
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   163
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   162
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   161
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   160
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -70680
         TabIndex        =   159
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -70680
         TabIndex        =   158
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70680
         TabIndex        =   157
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   156
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   155
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   154
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   153
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   152
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   151
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   150
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   149
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   148
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   147
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   146
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   145
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   144
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   143
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   142
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   141
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   140
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   139
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   138
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   137
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   136
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   135
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   134
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   133
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   132
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   131
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   130
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   129
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "ConsistControlComment"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   94
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         DataField       =   "ConsistControlComment"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         TabIndex        =   93
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         DataField       =   "ConsistControlComment"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         TabIndex        =   92
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         DataField       =   "ConsistControlComment"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         TabIndex        =   91
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox TimeDelay 
         Alignment       =   2  'Center
         DataField       =   "Time Delay"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   90
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeDelay 
         Alignment       =   2  'Center
         DataField       =   "Time Delay"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   89
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeDelay 
         Alignment       =   2  'Center
         DataField       =   "Time Delay"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   88
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeDelay 
         Alignment       =   2  'Center
         DataField       =   "Time Delay"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   87
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox SevenByteD 
         Alignment       =   2  'Center
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   86
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteD 
         Alignment       =   2  'Center
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   85
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteD 
         Alignment       =   2  'Center
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   84
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteD 
         Alignment       =   2  'Center
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   83
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SixByteD 
         Alignment       =   2  'Center
         DataField       =   "FourthByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   82
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteD 
         Alignment       =   2  'Center
         DataField       =   "FourthByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   81
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteD 
         Alignment       =   2  'Center
         DataField       =   "FourthByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   80
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteD 
         Alignment       =   2  'Center
         DataField       =   "FourthByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   79
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox FiveByteD 
         Alignment       =   2  'Center
         DataField       =   "ThirdByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   78
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteD 
         Alignment       =   2  'Center
         DataField       =   "ThirdByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   77
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteD 
         Alignment       =   2  'Center
         DataField       =   "ThirdByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   76
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteD 
         Alignment       =   2  'Center
         DataField       =   "ThirdByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   75
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FourByteD 
         Alignment       =   2  'Center
         DataField       =   "SecondByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   74
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteD 
         Alignment       =   2  'Center
         DataField       =   "SecondByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   73
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteD 
         Alignment       =   2  'Center
         DataField       =   "SecondByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   72
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteD 
         Alignment       =   2  'Center
         DataField       =   "SecondByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   71
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteD 
         Alignment       =   2  'Center
         DataField       =   "FirstByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   70
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteD 
         Alignment       =   2  'Center
         DataField       =   "FirstByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   69
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteD 
         Alignment       =   2  'Center
         DataField       =   "FirstByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   68
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteD 
         Alignment       =   2  'Center
         DataField       =   "FirstByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   67
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TwoByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteTwo"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   66
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteTwo"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   65
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteTwo"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   64
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox OneByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteOne"
         DataSource      =   "LocomotiveDatabaseConsistControl(4)"
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   63
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteOne"
         DataSource      =   "LocomotiveDatabaseConsistControl(3)"
         Height          =   285
         Index           =   3
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   62
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteOne"
         DataSource      =   "LocomotiveDatabaseConsistControl(2)"
         Height          =   285
         Index           =   2
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   61
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TwoByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteTwo"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   60
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox OneByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteOne"
         DataSource      =   "LocomotiveDatabaseConsistControl(1)"
         Height          =   285
         Index           =   1
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   59
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteOne"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   28
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   27
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TwoByteD 
         Alignment       =   2  'Center
         DataField       =   "AddressByteTwo"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   25
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   24
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteD 
         Alignment       =   2  'Center
         DataField       =   "FirstByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   22
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   21
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox FourByteD 
         Alignment       =   2  'Center
         DataField       =   "SecondByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   19
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   18
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FiveByteD 
         Alignment       =   2  'Center
         DataField       =   "ThirdByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   16
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   15
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   14
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   13
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   12
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70680
         TabIndex        =   11
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SixByteD 
         Alignment       =   2  'Center
         DataField       =   "FourthByte"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   10
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SevenByteD 
         Alignment       =   2  'Center
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton ButtonStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -69240
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   7
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox TimeDelay 
         Alignment       =   2  'Center
         DataField       =   "Time Delay"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   0
         Left            =   -74760
         TabIndex        =   3
         Top             =   4920
         Width           =   6855
         Begin VB.TextBox LocomotiveDatabaseConsistControlFileName 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   5175
         End
         Begin VB.CommandButton ButtonControlFile 
            Caption         =   "&Control File"
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -69240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LocomoriveDatabaseConsist"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "ConsistControlComment"
         DataSource      =   "LocomotiveDatabaseConsistControl(0)"
         Height          =   285
         Index           =   0
         Left            =   -73320
         TabIndex        =   2
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   213
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   3
         Left            =   -73800
         TabIndex        =   212
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   211
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   1
         Left            =   -73800
         TabIndex        =   210
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   0
         Left            =   -73800
         TabIndex        =   209
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   128
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   127
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   126
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   125
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   124
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   123
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Third Byte"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   122
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Third Byte"
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   121
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Third Byte"
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   120
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Third Byte"
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   119
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fourth Byte"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   118
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fourth Byte"
         Height          =   195
         Index           =   3
         Left            =   -74280
         TabIndex        =   117
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fourth Byte"
         Height          =   195
         Index           =   2
         Left            =   -74280
         TabIndex        =   116
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fourth Byte"
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   115
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fifth Byte"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   114
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fifth Byte"
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   113
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fifth Byte"
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   112
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fifth Byte"
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   111
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sixth Byte"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   110
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sixth Byte"
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   109
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sixth Byte"
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   108
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sixth Byte"
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   107
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Error Byte"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   106
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Error Byte"
         Height          =   195
         Index           =   3
         Left            =   -74160
         TabIndex        =   105
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Error Byte"
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   104
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Error Byte"
         Height          =   195
         Index           =   1
         Left            =   -74160
         TabIndex        =   103
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time Delay"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   102
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time Delay"
         Height          =   195
         Index           =   3
         Left            =   -74280
         TabIndex        =   101
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time Delay"
         Height          =   195
         Index           =   2
         Left            =   -74280
         TabIndex        =   100
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time Delay"
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   99
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   98
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   3
         Left            =   -74280
         TabIndex        =   97
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   96
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   95
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   58
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   57
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   3
         Left            =   -70680
         TabIndex        =   56
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   2
         Left            =   -70680
         TabIndex        =   55
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   1
         Left            =   -70680
         TabIndex        =   54
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   53
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   3
         Left            =   -72000
         TabIndex        =   52
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   2
         Left            =   -72000
         TabIndex        =   51
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   50
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   3
         Left            =   -73320
         TabIndex        =   49
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   2
         Left            =   -73320
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   47
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   1
         Left            =   -72000
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   45
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status: Clear"
         Height          =   195
         Index           =   5
         Left            =   -74640
         TabIndex        =   43
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   42
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   0
         Left            =   -70680
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   0
         Left            =   -73320
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   0
         Left            =   -72000
         TabIndex        =   39
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address Byte"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   38
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Third Byte"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   37
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fourth Byte"
         Height          =   195
         Index           =   0
         Left            =   -74280
         TabIndex        =   36
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fifth Byte"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   35
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sixth Byte"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   34
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Error Byte"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   33
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   15
         Left            =   -73800
         TabIndex        =   32
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time Delay"
         Height          =   195
         Index           =   0
         Left            =   -74280
         TabIndex        =   31
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   0
         Left            =   -74280
         TabIndex        =   30
         Top             =   4200
         Width           =   855
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7800
      Top             =   4440
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7800
      Top             =   3960
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7800
      Top             =   3360
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Time delay control of"
      Height          =   195
      Left            =   7560
      TabIndex        =   233
      Top             =   2280
      Width           =   1470
   End
   Begin VB.Label TimerIntervalSettingT 
      AutoSize        =   -1  'True
      Caption         =   "1000"
      Height          =   195
      Left            =   9120
      TabIndex        =   232
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label15 
      Caption         =   "Is the Communication Port Busy?"
      Height          =   255
      Left            =   7560
      TabIndex        =   225
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label13 
      Caption         =   "Letting the computer take control of your locomotive, just select the Locomotive Control File and start the journey."
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   135
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog LocomotiveDatabaseConsistControlSelect 
      Left            =   6720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Locomotive Control File"
      Filter          =   "Control Files|Con*.mdb"
   End
End
Attribute VB_Name = "MainlineOperationATC"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub ButtonClose_Click()

Let Timer1.Interval = 0
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
            If TemporaryScreen = "Mainline Operation ATC Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation ATC Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Operation ATC Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Operation ATC Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub ButtonControlFile_Click(Index As Integer)

Timer1.Interval = 0

' Common Dialog Box to get a filename.
'
' Displays the CommonDialog control's Open dialog box.
' The object placeholder represents an object expression that evaluates to an object in the Applies To list.

    Let LocomotiveDatabaseConsistControlSelect.Filter = "Consist Databases Files|Consist*.mdb"
    Let LocomotiveDatabaseConsistControlSelect.InitDir = App.Path$ & "\Databases\"
    LocomotiveDatabaseConsistControlSelect.ShowOpen
    
' Once the user has selected a file to use, place the file name in the text box for visible clarification and place
' the name in the database control property.  Please note that the 'record set' is hard coded. This prevents any
' non-consist control database files form being accidentally used. Not error routines have been made.

    Let LocomotiveDatabaseConsistControlFileName(Index).Text = LocomotiveDatabaseConsistControlSelect.Filename
    Let LocomotiveDatabaseConsistControl(Index).DatabaseName = LocomotiveDatabaseConsistControlSelect.Filename
    Let LocomotiveDatabaseConsistControl(Index).RecordSource = "LocomoriveDatabaseConsist"
        LocomotiveDatabaseConsistControl(Index).Refresh
        
' Ok, we're out of here.
 
    If LocomotiveDatabaseConsistControlFileName(Index).Text = "" Then
        Let ButtonStart(Index).Enabled = False
        Let ButtonStop(Index).Enabled = False
        Let ButtonControlFile(Index).Enabled = True
    Else
        Let ButtonStart(Index).Enabled = True
        Let ButtonStop(Index).Enabled = False
        Let ButtonControlFile(Index).Enabled = True
    End If
   
    Let Timer1.Interval = Val(TimerIntervalSetting.Value)
      
End Sub






Private Sub ButtonPrint_Click()

    MainlineOperationATC.PrintForm
    
End Sub

Private Sub ButtonStart_Click(Index As Integer)
 
Let ButtonStart(Index).Enabled = False
Let ButtonStop(Index).Enabled = True
Let ButtonControlFile(Index).Enabled = False

End Sub





Private Sub ButtonStop_Click(Index As Integer)
 
Let ButtonStart(Index).Enabled = False
Let ButtonStop(Index).Enabled = False
Let ButtonControlFile(Index).Enabled = True

End Sub







Private Sub FiveByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let FiveByteH(Index).Text = Hex(Val(FiveByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(FiveByteH(Index).Text) = 1 Then
        Let FiveByteH(Index).Text = "0" + FiveByteH(Index).Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(FiveByteD(Index).Text)
        
        If temp / 128 >= 1 Then
                Let FiveByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                FiveByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let FiveByteB(Index).Text = FiveByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                FiveByteB(Index).Text = FiveByteB(Index).Text + "0"
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
        If TemporaryScreen = "Mainline Operation ATC Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Operation ATC Screen"
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
            Let Ini.Value = "Mainline Operation ATC Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Operation ATC Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation ATC Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Operation ATC Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationATC.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationATC.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationATC.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationATC.Height)

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
                Let Ini.Value = "Mainline Operation ATC Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation ATC Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineOperationATC.Hide
    'unload MainlineoperationATC

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
    Let Ini.Application = "Mainline Operation ATC Screen"
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
        MainlineOperationATC.Left = (Screen.Width - Width) / 2
        MainlineOperationATC.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineOperationATC.Width > Screen.Width Then
            Let MainlineOperationATC.Left = Screen.Width - MainlineOperationATC.Width
        Else
            Let MainlineOperationATC.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineOperationATC.Height > Screen.Height Then
            Let MainlineOperationATC.Top = Screen.Height - MainlineOperationATC.Height
        Else
            Let MainlineOperationATC.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This text box is where all information from your" & vbCrLf & "serial port is displayed. Commands given by the" & vbCrLf & "program are displayed here. You can also type your" & vbCrLf & "own commands, providing the port is not busy."
        Let TemporaryText2 = "Communication Window"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowDCC)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowDCC, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button closes the current window and" & vbCrLf & "returns to to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let LocomotiveDatabaseConsistControl(0).DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
    Let LocomotiveDatabaseConsistControl(1).DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
    Let LocomotiveDatabaseConsistControl(2).DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
    Let LocomotiveDatabaseConsistControl(3).DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
    Let LocomotiveDatabaseConsistControl(4).DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
    Let LocomotiveDatabaseConsistControlSelect.InitDir = App.Path$ & "\Databases"
 
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

    If MainlineOperationATC.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineOperationATC.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Operation ATC Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationATC.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationATC.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationATC.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationATC.Height)
 
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

Private Sub LocomotiveConsistDefault_Click(Index As Integer)

If LocomotiveConsistDefault(Index).Value = vbChecked Then
    Let onebyted(Index).BackColor = &HC0C0C0
    Let onebyted(Index).Enabled = False
    Let onebyted(Index).DataField = ""
    Let onebyted(Index).Text = MainlineConsist.LocomotiveConsistNumber.Text
    Let twoByteD(Index).BackColor = &HC0C0C0
    Let twoByteD(Index).Enabled = False
    Let twoByteD(Index).DataField = ""
    Let twoByteD(Index).Text = ""
Else
    Let onebyted(Index).BackColor = &H80000005
    Let onebyted(Index).Enabled = True
    Let onebyted(Index).DataField = "AddressByteOne"
    Let twoByteD(Index).BackColor = &H80000005
    Let twoByteD(Index).Enabled = True
    Let twoByteD(Index).DataField = "AddressByteTwo"
End If

End Sub

Private Sub OneByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let OneByteH(Index).Text = Hex(Val(onebyted(Index).Text))
    
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(OneByteH(Index).Text) = 1 Then
        Let OneByteH(Index).Text = "0" + OneByteH(Index).Text
    End If
    
Rem Binary Conversion
    
        Let temp = Val(onebyted(Index).Text)
        
        If temp / 128 >= 1 Then
                Let OneByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                Let OneByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let OneByteB(Index).Text = OneByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                Let OneByteB(Index).Text = OneByteB(Index).Text + "0"
        End If

End Sub


Private Sub FourByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let FourByteH(Index).Text = Hex(Val(FourByteD(Index).Text))
    
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(FourByteH(Index).Text) = 1 Then
        Let FourByteH(Index).Text = "0" + FourByteH(Index).Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(FourByteD(Index).Text)
        
        If temp / 128 >= 1 Then
                Let FourByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                Let FourByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let FourByteB(Index).Text = FourByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                Let FourByteB(Index).Text = FourByteB(Index).Text + "0"
        End If

End Sub




Private Sub Timer1_Timer()

'------------------------------------------------------------------------------------------------------------
' Loop routine
'
' THis is almost a continuous loop. Since Visual Basic cannot do more that one event at a time, I have to
' process all five possible consist in automatic run time. This is done by cycling through each one of the
' consist and then processing the next one (only if the start button is activated.

    For Y = 0 To 4 Step 1

'------------------------------------------------------------------------------------------------------------
' Conditional Statement
'
' Check to see if the start button is diabled (the user them clicked on it) and the stop button is enabled
' (the user has opened a consist control file and is in operation)

        If ButtonStart(Y).Enabled = False And _
            ButtonStop(Y).Enabled = True Then
        
'------------------------------------------------------------------------------------------------------------
' Move to the first Record
'
            If LocomotiveDatabaseConsistControl(Y).Recordset.EOF = True Then
                If LocomotiveDatabaseConsistControlLoop(Y).Value = vbChecked Then
                    LocomotiveDatabaseConsistControl(Y).Recordset.MoveFirst
                Else
                    Let ButtonStart(Y).Enabled = False
                    Let ButtonStop(Y).Enabled = False
                    Let ButtonControlFile(Y).Enabled = True
                    Let LocomotiveCommunicationStatus(Y).Caption = "Status: End of File"
                End If
            End If




' -----------------------------------------------------------------------------------------------------------
' Check if delayed time
'


            If Val(TimeDelay(Y).Text) <> Val(TimeLeft(Y).Text) Then
                    Let TimeLeft(Y).Text = Val(TimeLeft(Y).Text) + 1
                    Let LocomotiveCommunicationStatus(Y).Caption = "Status: Time Delay of " + TimeLeft(Y).Text + " seconds of " + TimeDelay(Y).Text
                Else
                    LocomotiveDatabaseConsistControl(Y).Recordset.MoveNext
                    Let TimeLeft(Y).Text = "1"

            If LocomotiveDatabaseConsistControl(Y).Recordset.EOF = False Then

            Let sevenbyted(Y).Text = ""

                For X = 1 To 8
                
                    Let temp = 0


                        If onebyted(Y).Text <> "" Then
                            If onebyted(Y).Text <> "   " Then
                                If Mid$(OneByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If twoByteD(Y).Text <> "" Then
                            If twoByteD(Y).Text <> "   " Then
                                If Mid$(TwoByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If ThreeByteD(Y).Text <> "" Then
                            If ThreeByteD(Y).Text <> "   " Then
                                If Mid$(ThreeByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If FourByteD(Y).Text <> "" Then
                            If FourByteD(Y).Text <> "   " Then
                                If Mid$(FourByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If FiveByteD(Y).Text <> "" Then
                            If FiveByteD(Y).Text <> "   " Then
                                If Mid$(FiveByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If SixByteD(Y).Text <> "" Then
                            If SixByteD(Y).Text <> "   " Then
                                If Mid$(SixByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
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
    
                If Int(temp / 2) <> (temp / 2) Then
                    Let sevenbyted(Y).Text = Val(sevenbyted(Y).Text) + bitvalue
                End If
                        
            Next X
                    
' -----------------------------------------------------------------------------------------------------------
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.

            Let LocomotiveCommunicationStatus(Y).Caption = "Status: Sending Command"

            Let CommandControl = "q"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
            If onebyted(Y).Text <> "" Then
                If onebyted(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & OneByteH(Y).Text
                End If
            End If
            If twoByteD(Y).Text <> "" Then
                If twoByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & TwoByteH(Y).Text
                End If
            End If
            If ThreeByteD(Y).Text <> "" Then
                If ThreeByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & ThreeByteH(Y).Text
                End If
            End If
            If FourByteD(Y).Text <> "" Then
                If FourByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & FourByteH(Y).Text
                End If
            End If
            If FiveByteD(Y).Text <> "" Then
                If FiveByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & FiveByteH(Y).Text
                End If
            End If
            If SixByteD(Y).Text <> "" Then
                If SixByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & SixByteH(Y).Text
                End If
            End If
            If sevenbyted(Y).Text <> "" Then
                If sevenbyted(Y).Text <> "   " Then
                    Let CommandControl = CommandControl & " " & SevenByteH(Y).Text
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

            Let CommunicationPortBusy(Y).Value = vbChecked

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

            MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent parameters as per National Model Railroader Association specification.
    
            Let LocomotiveCommunicationStatus(Y).Caption = "Status: Command Sent"

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
            Let LocomotiveCommunicationStatus(Y).Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

            Let CommunicationPortBusy(Y).Value = vbUnchecked

' -------------------------------------------------------------------------------------------------------------------------------------------
' Dequing Packet
'
            If MainScreen!checkboxdequeuepacket.Value = vbChecked Then

                Let LocomotiveCommunicationStatus(Y).Caption = "Status: Sending Command"

' As well, I initially set the 'commandcontrol' string to the North Coast Engineering command for sending a command to the
' decoder. The following format is used in sending a packet:
'       's cxx yy yy..'
'   where 's' repersent the command to send a packet
'   where 'c' represent the nottation of number of times to repeat this packet.
'   where 'xx' is the number of times to send this packet in hexidecimal. I've hardcoded this to four.
'   where 'yy' is the data to be sent to the command station, and repeated as often as necessary.
' The last hexidecimal should be the error byte.

                Let CommandControl = "d"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
                If onebyted(Y).Text <> "" Then
                   If onebyted(Y).Text <> "   " Then
                        Let CommandControl = CommandControl & " " & OneByteH(Y).Text
                    End If
                End If
                If twoByteD(Y).Text <> "" Then
                    If twoByteD(Y).Text <> "   " Then
                        Let CommandControl = CommandControl & " " & TwoByteH(Y).Text
                    End If
                End If
                'If ThreeByteD(y).Text <> "" Then
                '    If ThreeByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & ThreeByteH(y).Text
                '    End If
                'End If
                'If FourByteD(y).Text <> "" Then
                '    If FourByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & FourByteH(y).Text
                '    End If
                'End If
                'If FiveByteD(y).Text <> "" Then
                '    If FiveByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & FiveByteH(y).Text
                '    End If
                'End If
                'If SixByteD(y).Text <> "" Then
                '    If SixByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & SixByteH(y).Text
                '    End If
                'End If
                'If SevenByteD(y).Text <> "" Then
                '    If SevenByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl & " " & SevenByteH(y).Text
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

                Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + CommandControl + Chr$(10)
                Let MainScreen.TextBoxCommunicationWindowDCC.SelStart = Len(MainScreen.TextBoxCommunicationWindowDCC.Text)

                Let CommunicationPortBusy(Y).Value = vbChecked

                MainScreen.MSComm1.Output = CommandControl
    
                Let LocomotiveCommunicationStatus(Y).Caption = "Status: Command Sent"

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
                Let LocomotiveCommunicationStatus(Y).Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

                Let CommunicationPortBusy(Y).Value = vbUnchecked

' -----------------------------------------------------------------------------------------------------------
' Conditional Statement Ending
'
' If packet is to be dequed then

                End If

                'If Val(OneByteD(Y).Text) = 0 Then
                '    If Val(TwoByteD(Y).Text) = 0 Then
                '        If Val(ThreeByteD(Y).Text) = 0 Then
                '            If Val(FourByteD(Y).Text) = 0 Then
                '                If Val(FiveByteD(Y).Text) = 0 Then
                '                    If Val(SixByteD(Y).Text) = 0 Then
                '                        If Val(SevenByteD(Y).Text) = 0 Then
                                            If Val(TimeDelay(Y).Text) = 0 Then
                                                If LocomotiveDatabaseConsistControlLoop(Y).Value = vbChecked Then
                                                    LocomotiveDatabaseConsistControl(Y).Recordset.MoveFirst
                                                Else
                                                Let ButtonStart(Y).Enabled = False
                                                Let ButtonStop(Y).Enabled = False
                                                Let ButtonControlFile(Y).Enabled = True
                                                Let LocomotiveCommunicationStatus(Y).Caption = "Status: End of File"
                                                End If
                                            End If
                '                        End If
                '                    End If
                '                End If
                '            End If
                '        End If
                '    End If
                'End If
        
' Let's set the communication visual status back to clear, we are now finished with the consist control.

'                Let LocomotiveCommunicationStatus(Y).Caption = "Status: Clear"

' -----------------------------------------------------------------------------------------------------------
' Conditional Statement - Timer Delay?
'
                End If



' -----------------------------------------------------------------------------------------------------------
' End of Conditional Statement - While Not EOF
'
            End If

' -----------------------------------------------------------------------------------------------------------
' End of Conditional Statement - If Start Button
'
        End If
    
' -----------------------------------------------------------------------------------------------------------
' End of Loop for Y
'
    Next Y
    
End Sub

Private Sub TimerIntervalSetting_Change()

    Let Timer1.Interval = Val(TimerIntervalSetting.Value)
    Let TimerIntervalSettingT.Caption = Val(TimerIntervalSetting.Value)
    
End Sub

Private Sub TwoByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let TwoByteH(Index).Text = Hex(Val(twoByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(TwoByteH(Index).Text) = 1 Then
        Let TwoByteH(Index).Text = "0" + TwoByteH(Index).Text
    End If

' This routine prforms the Binary Conversion
    
        Let temp = Val(twoByteD(Index).Text)
        
        If temp / 128 >= 1 Then
                Let TwoByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                Let TwoByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                Let TwoByteB(Index).Text = TwoByteB(Index).Text + "0"
        End If

End Sub
Private Sub SevenByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let SevenByteH(Index).Text = Hex(Val(sevenbyted(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SevenByteH(Index).Text) = 1 Then
        Let SevenByteH(Index).Text = "0" + SevenByteH(Index).Text
    End If

Rem Binary Conversion
    
        Let temp = Val(sevenbyted(Index).Text)
        
        If temp / 128 >= 1 Then
                Let SevenByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                Let SevenByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                Let SevenByteB(Index).Text = SevenByteB(Index).Text + "0"
        End If

End Sub
Private Sub SixByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let SixByteH(Index).Text = Hex(Val(SixByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SixByteH(Index).Text) = 1 Then
        Let SixByteH(Index).Text = "0" + SixByteH(Index).Text
    End If

Rem Binary Conversion
    
        Let temp = Val(SixByteD(Index).Text)
        
        If temp / 128 >= 1 Then
                Let SixByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                Let SixByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let SixByteB(Index).Text = SixByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                Let SixByteB(Index).Text = SixByteB(Index).Text + "0"
        End If

End Sub
Private Sub ThreeByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let ThreeByteH(Index).Text = Hex(Val(ThreeByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(ThreeByteH(Index).Text) = 1 Then
        Let ThreeByteH(Index).Text = "0" + ThreeByteH(Index).Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(ThreeByteD(Index).Text)
        
        If temp / 128 >= 1 Then
                Let ThreeByteB(Index).Text = "1"
                Let temp = temp - 128
            Else
                ThreeByteB(Index).Text = "0"
        End If
        
        If temp / 64 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 64
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 32 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 32
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 16 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 16
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 8 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 8
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 4 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 4
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 2 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 2
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If
        
        If temp / 1 >= 1 Then
                Let ThreeByteB(Index).Text = ThreeByteB(Index).Text + "1"
                Let temp = temp - 1
            Else
                ThreeByteB(Index).Text = ThreeByteB(Index).Text + "0"
        End If

End Sub




