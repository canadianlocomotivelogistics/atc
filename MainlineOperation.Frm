VERSION 4.00
Begin VB.Form MainlineOperation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Operation"
   ClientHeight    =   6720
   ClientLeft      =   285
   ClientTop       =   1500
   ClientWidth     =   9945
   Height          =   7125
   Icon            =   "MainlineOperation.frx":0000
   Left            =   225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Top             =   1155
   Width           =   10065
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   180
      Picture         =   "MainlineOperation.frx":0442
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
      TabPicture(0)   =   "MainlineOperation.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "LocomotiveCommunicationStatus(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "LocomotiveDatabaseConsistControl(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TimeDelay(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "LocomotiveDatabaseConsistControlLoop(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ButtonStart(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "SevenByteD(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "SixByteD(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "SevenByteB(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "SevenByteH(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "SixByteB(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "SixByteH(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "FiveByteB(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "FiveByteH(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "FiveByteD(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "FourByteB(0)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "FourByteH(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "FourByteD(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "ThreeByteB(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "ThreeByteH(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ThreeByteD(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TwoByteB(0)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TwoByteH(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TwoByteD(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "OneByteB(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "OneByteH(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "OneByteD(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "LocomotiveConsistDefault(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "ButtonStop(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TimeLeft(0)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "ATC #1"
      TabPicture(1)   =   "MainlineOperation.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TimeLeft(1)"
      Tab(1).Control(1)=   "LocomotiveConsistDefault(1)"
      Tab(1).Control(2)=   "LocomotiveDatabaseConsistControl(1)"
      Tab(1).Control(3)=   "ButtonStop(1)"
      Tab(1).Control(4)=   "ButtonStart(1)"
      Tab(1).Control(5)=   "Frame1(1)"
      Tab(1).Control(6)=   "LocomotiveDatabaseConsistControlLoop(1)"
      Tab(1).Control(7)=   "SevenByteB(1)"
      Tab(1).Control(8)=   "SixByteB(1)"
      Tab(1).Control(9)=   "FiveByteB(1)"
      Tab(1).Control(10)=   "FourByteB(1)"
      Tab(1).Control(11)=   "ThreeByteB(1)"
      Tab(1).Control(12)=   "TwoByteB(1)"
      Tab(1).Control(13)=   "OneByteB(1)"
      Tab(1).Control(14)=   "SevenByteH(1)"
      Tab(1).Control(15)=   "SixByteH(1)"
      Tab(1).Control(16)=   "FiveByteH(1)"
      Tab(1).Control(17)=   "FourByteH(1)"
      Tab(1).Control(18)=   "ThreeByteH(1)"
      Tab(1).Control(19)=   "TwoByteH(1)"
      Tab(1).Control(20)=   "OneByteH(1)"
      Tab(1).Control(21)=   "Text1(1)"
      Tab(1).Control(22)=   "TimeDelay(1)"
      Tab(1).Control(23)=   "SevenByteD(1)"
      Tab(1).Control(24)=   "SixByteD(1)"
      Tab(1).Control(25)=   "FiveByteD(1)"
      Tab(1).Control(26)=   "FourByteD(1)"
      Tab(1).Control(27)=   "ThreeByteD(1)"
      Tab(1).Control(28)=   "TwoByteD(1)"
      Tab(1).Control(29)=   "OneByteD(1)"
      Tab(1).Control(30)=   "LocomotiveCommunicationStatus(1)"
      Tab(1).Control(31)=   "Label5(1)"
      Tab(1).Control(32)=   "Label6(1)"
      Tab(1).Control(33)=   "Label7(1)"
      Tab(1).Control(34)=   "Label8(1)"
      Tab(1).Control(35)=   "Label9(1)"
      Tab(1).Control(36)=   "Label10(1)"
      Tab(1).Control(37)=   "Label14(1)"
      Tab(1).Control(38)=   "Label12(1)"
      Tab(1).Control(39)=   "Label2(1)"
      Tab(1).Control(40)=   "Label1(1)"
      Tab(1).Control(41)=   "Label4(1)"
      Tab(1).Control(42)=   "Label3(1)"
      Tab(1).ControlCount=   43
      TabCaption(2)   =   "ATC #2"
      TabPicture(2)   =   "MainlineOperation.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TimeLeft(2)"
      Tab(2).Control(1)=   "LocomotiveConsistDefault(2)"
      Tab(2).Control(2)=   "LocomotiveDatabaseConsistControl(2)"
      Tab(2).Control(3)=   "ButtonStop(2)"
      Tab(2).Control(4)=   "ButtonStart(2)"
      Tab(2).Control(5)=   "Frame1(2)"
      Tab(2).Control(6)=   "LocomotiveDatabaseConsistControlLoop(2)"
      Tab(2).Control(7)=   "SevenByteB(2)"
      Tab(2).Control(8)=   "SixByteB(2)"
      Tab(2).Control(9)=   "FiveByteB(2)"
      Tab(2).Control(10)=   "FourByteB(2)"
      Tab(2).Control(11)=   "ThreeByteB(2)"
      Tab(2).Control(12)=   "TwoByteB(2)"
      Tab(2).Control(13)=   "OneByteB(2)"
      Tab(2).Control(14)=   "SevenByteH(2)"
      Tab(2).Control(15)=   "SixByteH(2)"
      Tab(2).Control(16)=   "FiveByteH(2)"
      Tab(2).Control(17)=   "FourByteH(2)"
      Tab(2).Control(18)=   "ThreeByteH(2)"
      Tab(2).Control(19)=   "TwoByteH(2)"
      Tab(2).Control(20)=   "OneByteH(2)"
      Tab(2).Control(21)=   "Text1(2)"
      Tab(2).Control(22)=   "TimeDelay(2)"
      Tab(2).Control(23)=   "SevenByteD(2)"
      Tab(2).Control(24)=   "SixByteD(2)"
      Tab(2).Control(25)=   "FiveByteD(2)"
      Tab(2).Control(26)=   "FourByteD(2)"
      Tab(2).Control(27)=   "ThreeByteD(2)"
      Tab(2).Control(28)=   "TwoByteD(2)"
      Tab(2).Control(29)=   "OneByteD(2)"
      Tab(2).Control(30)=   "LocomotiveCommunicationStatus(2)"
      Tab(2).Control(31)=   "Label5(2)"
      Tab(2).Control(32)=   "Label6(2)"
      Tab(2).Control(33)=   "Label7(2)"
      Tab(2).Control(34)=   "Label8(2)"
      Tab(2).Control(35)=   "Label9(2)"
      Tab(2).Control(36)=   "Label10(2)"
      Tab(2).Control(37)=   "Label14(2)"
      Tab(2).Control(38)=   "Label12(2)"
      Tab(2).Control(39)=   "Label1(2)"
      Tab(2).Control(40)=   "Label2(2)"
      Tab(2).Control(41)=   "Label4(2)"
      Tab(2).Control(42)=   "Label3(2)"
      Tab(2).ControlCount=   43
      TabCaption(3)   =   "ATC #3"
      TabPicture(3)   =   "MainlineOperation.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TimeLeft(3)"
      Tab(3).Control(1)=   "LocomotiveConsistDefault(3)"
      Tab(3).Control(2)=   "LocomotiveDatabaseConsistControl(3)"
      Tab(3).Control(3)=   "ButtonStop(3)"
      Tab(3).Control(4)=   "ButtonStart(3)"
      Tab(3).Control(5)=   "Frame1(3)"
      Tab(3).Control(6)=   "LocomotiveDatabaseConsistControlLoop(3)"
      Tab(3).Control(7)=   "SevenByteB(3)"
      Tab(3).Control(8)=   "SixByteB(3)"
      Tab(3).Control(9)=   "FiveByteB(3)"
      Tab(3).Control(10)=   "FourByteB(3)"
      Tab(3).Control(11)=   "ThreeByteB(3)"
      Tab(3).Control(12)=   "TwoByteB(3)"
      Tab(3).Control(13)=   "OneByteB(3)"
      Tab(3).Control(14)=   "SevenByteH(3)"
      Tab(3).Control(15)=   "SixByteH(3)"
      Tab(3).Control(16)=   "FiveByteH(3)"
      Tab(3).Control(17)=   "FourByteH(3)"
      Tab(3).Control(18)=   "ThreeByteH(3)"
      Tab(3).Control(19)=   "TwoByteH(3)"
      Tab(3).Control(20)=   "OneByteH(3)"
      Tab(3).Control(21)=   "Text1(3)"
      Tab(3).Control(22)=   "TimeDelay(3)"
      Tab(3).Control(23)=   "SevenByteD(3)"
      Tab(3).Control(24)=   "SixByteD(3)"
      Tab(3).Control(25)=   "FiveByteD(3)"
      Tab(3).Control(26)=   "FourByteD(3)"
      Tab(3).Control(27)=   "ThreeByteD(3)"
      Tab(3).Control(28)=   "TwoByteD(3)"
      Tab(3).Control(29)=   "OneByteD(3)"
      Tab(3).Control(30)=   "LocomotiveCommunicationStatus(3)"
      Tab(3).Control(31)=   "Label1(3)"
      Tab(3).Control(32)=   "Label5(3)"
      Tab(3).Control(33)=   "Label6(3)"
      Tab(3).Control(34)=   "Label7(3)"
      Tab(3).Control(35)=   "Label8(3)"
      Tab(3).Control(36)=   "Label9(3)"
      Tab(3).Control(37)=   "Label10(3)"
      Tab(3).Control(38)=   "Label14(3)"
      Tab(3).Control(39)=   "Label12(3)"
      Tab(3).Control(40)=   "Label2(3)"
      Tab(3).Control(41)=   "Label4(3)"
      Tab(3).Control(42)=   "Label3(3)"
      Tab(3).ControlCount=   43
      TabCaption(4)   =   "ATC #4"
      TabPicture(4)   =   "MainlineOperation.frx":08F4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TimeLeft(4)"
      Tab(4).Control(1)=   "LocomotiveConsistDefault(4)"
      Tab(4).Control(2)=   "LocomotiveDatabaseConsistControl(4)"
      Tab(4).Control(3)=   "ButtonStop(4)"
      Tab(4).Control(4)=   "ButtonStart(4)"
      Tab(4).Control(5)=   "Frame1(4)"
      Tab(4).Control(6)=   "LocomotiveDatabaseConsistControlLoop(4)"
      Tab(4).Control(7)=   "SevenByteB(4)"
      Tab(4).Control(8)=   "SixByteB(4)"
      Tab(4).Control(9)=   "FiveByteB(4)"
      Tab(4).Control(10)=   "FourByteB(4)"
      Tab(4).Control(11)=   "ThreeByteB(4)"
      Tab(4).Control(12)=   "TwoByteB(4)"
      Tab(4).Control(13)=   "OneByteB(4)"
      Tab(4).Control(14)=   "SevenByteH(4)"
      Tab(4).Control(15)=   "SixByteH(4)"
      Tab(4).Control(16)=   "FiveByteH(4)"
      Tab(4).Control(17)=   "FourByteH(4)"
      Tab(4).Control(18)=   "ThreeByteH(4)"
      Tab(4).Control(19)=   "TwoByteH(4)"
      Tab(4).Control(20)=   "OneByteH(4)"
      Tab(4).Control(21)=   "Text1(4)"
      Tab(4).Control(22)=   "TimeDelay(4)"
      Tab(4).Control(23)=   "SevenByteD(4)"
      Tab(4).Control(24)=   "SixByteD(4)"
      Tab(4).Control(25)=   "FiveByteD(4)"
      Tab(4).Control(26)=   "FourByteD(4)"
      Tab(4).Control(27)=   "ThreeByteD(4)"
      Tab(4).Control(28)=   "TwoByteD(4)"
      Tab(4).Control(29)=   "OneByteD(4)"
      Tab(4).Control(30)=   "LocomotiveCommunicationStatus(4)"
      Tab(4).Control(31)=   "Label1(4)"
      Tab(4).Control(32)=   "Label5(4)"
      Tab(4).Control(33)=   "Label6(4)"
      Tab(4).Control(34)=   "Label7(4)"
      Tab(4).Control(35)=   "Label8(4)"
      Tab(4).Control(36)=   "Label9(4)"
      Tab(4).Control(37)=   "Label10(4)"
      Tab(4).Control(38)=   "Label14(4)"
      Tab(4).Control(39)=   "Label12(4)"
      Tab(4).Control(40)=   "Label2(4)"
      Tab(4).Control(41)=   "Label4(4)"
      Tab(4).Control(42)=   "Label3(4)"
      Tab(4).ControlCount=   43
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   230
         Text            =   "00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TimeLeft 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
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
         Left            =   5760
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
         Left            =   -71760
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
         Left            =   3240
         TabIndex        =   214
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Data LocomotiveDatabaseConsistControl 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   4
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
         Caption         =   "Data1"
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
         Caption         =   "Data1"
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
         Caption         =   "Data1"
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
         Left            =   -69240
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
         Left            =   -69240
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
         Left            =   -74760
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
         Left            =   -74640
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   184
         Text            =   "00000000"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   180
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   176
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   172
         Text            =   "00000000"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   168
         Text            =   "00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   164
         Text            =   "00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -70680
         TabIndex        =   160
         Text            =   "00000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   156
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   152
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   148
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   144
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   140
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   136
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72000
         TabIndex        =   132
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   -73320
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   28
         Text            =   "00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox OneByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   25
         Text            =   "00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TwoByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   22
         Text            =   "00"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox ThreeByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   19
         Text            =   "00"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox FourByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   16
         Text            =   "00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox FiveByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   15
         Text            =   "00000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox SixByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Text            =   "00"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SixByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Text            =   "00000000"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox SevenByteH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         Text            =   "00"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox SevenByteB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4320
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
         Left            =   1680
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
         Left            =   1680
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
         Left            =   5760
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox LocomotiveDatabaseConsistControlLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuos Loop "
         Height          =   255
         Index           =   0
         Left            =   360
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Index           =   0
         Left            =   240
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
         Left            =   5760
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
         Left            =   1680
         TabIndex        =   2
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label LocomotiveCommunicationStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Index           =   4
         Left            =   -73800
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
         Left            =   1200
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
         Left            =   -74400
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
         Left            =   -74400
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
         Left            =   -74160
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
         Left            =   -74280
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
         Left            =   -74160
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
         Left            =   -74160
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
         Left            =   -74160
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
         Left            =   -74280
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
         Left            =   -74280
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
         Left            =   -70680
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
         Left            =   -72000
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
         Left            =   -73320
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
         Left            =   600
         TabIndex        =   42
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Index           =   0
         Left            =   3000
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
         Left            =   600
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
         Left            =   840
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
         Left            =   720
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
         Left            =   840
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
         Left            =   840
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
         Left            =   840
         TabIndex        =   33
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   15
         Left            =   1200
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
         Left            =   720
         TabIndex        =   31
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Top             =   4200
         Width           =   855
      End
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
Attribute VB_Name = "MainlineOperation"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub ButtonClose_Click()

Let Timer1.Interval = 0

' =========================================================================================================================
' Close the Programming Diesel Window
'
   
' My Programming Notes
'
' Now that the commubication port is opened (hopefully to the correct communicaton port connected to the North Coast
' Enginerring Power HOuse Pro) the additional options on the menu bar can be enabled or disabled.
    
    Let MainScreen.MenuCommunicationPortOpen.Enabled = False
    Let MainScreen.MenuCommunicationPortClosed.Enabled = True
    Let MainScreen.MenuCommunicationPortSetting.Enabled = False
    Let MainScreen.MenuScaledTimeSetting.Enabled = True
    Let MainScreen.MenuProgrammingModeDiesel.Enabled = True
    Let MainScreen.MenuProgrammingModeSteam.Enabled = True
    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = True
    Let MainScreen.MenuProgrammingModeOther.Enabled = True
    Let MainScreen.MenuMainlineDieselProgramming.Enabled = True
    Let MainScreen.MenuMainlineSteamProgramming.Enabled = True
    Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = True
    Let MainScreen.MenuMainlineOtherProgramming.Enabled = True
    Let MainScreen.menumainlineconsist.Enabled = True
    Let MainScreen.menumainlinescalespeedoperation.Enabled = True
    Let MainScreen.MenuMainlineOperation.Enabled = True
    Let MainScreen.MenuMainlineOperationGUI.Enabled = True
    Let MainScreen.MenuMainlineMacroMaker = True
    Let MainScreen.MenuLayoutDefineBlocks.Enabled = True
    Let MainScreen.MenuLayoutCTC.Enabled = True

' =========================================================================================================================
' Hide Method
'
'Hides an MDIForm or Form object but doesn't unload it.
'
' The syntax is as follows,
'
' object.Hide
'
' The object placeholder represents an object expression that evaluates to an object in the Applies To list.  If object
' is omitted, the form with the focus is assumed to be object. When a form is hidden, it's removed from the screen and
' its Visible property is set to False.  A hidden form's controls aren't accessible to the user, but they are available
' to the running Visual Basic application, to other processes that may be communicating with the application
' through DDE, and to Timer control events. When a form is hidden, the user can't interact with the application
' until all code in the event procedure that caused the form to be hidden has finished executing. If the form isn't
' loaded when the Hide method is invoked, the Hide method loads the form but doesn't display it.

    MainlineOperation.Hide
    
' =========================================================================================================================
' Uload Statement
'
' Unloads a form or control from memory.
'
' The syntax is as follows
'
'   Unload object
'
' The object placeholder is the name of a Form object or control array element to unload. Unloading a form or control may
' be necessary or expedient in some cases where the memory used is needed for something else or when you need to reset
' properties to their original values. Before a form is unloaded, the Query_Unload event procedure occurs, followed by
' the Form_Unload event procedure.  Setting the cancel argument to True in either of these events keeps the form from
' being unloaded.  For MDIForm objects, the MDIForm object's Query_Unload event procedure occurs, followed by
' the Query_Unload event procedure and Form_Unload event procedure for each MDI child form, and finally the MDIForm
' object's Form_Unload event procedure. When a form is unloaded, all controls placed on the form at run time are no
' longer accessible.  Controls placed on the form at design time remain intact; however, any run-time changes to
' those controls and their properties are lost when the form is reloaded.  All changes to form properties are
' also lost. When a form is unloaded, only the displayed component is unloaded.  The code associated with the form
' module remains in memory. Only control array elements added to a form at run time can be unloaded with the Unload
' statement.  The properties of unloaded controls are reinitialized when the controls are reloaded.

    Unload MainlineOperation

' =========================================================================================================================
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


Private Sub ButtonControlFile_Click(Index As Integer)

Timer1.Interval = 0

' Common Dialog Box to get a filename.
'
' Displays the CommonDialog control's Open dialog box.
' The object placeholder represents an object expression that evaluates to an object in the Applies To list.

    Let LocomotiveDatabaseConsistControlSelect.Filter = "Consist Databases Files|Consist*.mdb"
    Let LocomotiveDatabaseConsistControlSelect.InitDir = App.Path + "\Databases\"
    LocomotiveDatabaseConsistControlSelect.ShowOpen
    
' Once the user has selected a file to use, place the file name in the text box for visible clarification and place
' the name in the database control property.  Please note that the 'record set' is hard coded. This prevents any
' non-consist control database files form being accidentally used. Not error routines have been made.

    Let LocomotiveDatabaseConsistControlFileName(Index).Text = LocomotiveDatabaseConsistControlSelect.filename
    Let LocomotiveDatabaseConsistControl(Index).DatabaseName = LocomotiveDatabaseConsistControlSelect.filename
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

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form verti'cally.
  
Let LocomotiveDatabaseConsistControl(0).DatabaseName = App.Path + "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseConsistControl(1).DatabaseName = App.Path + "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseConsistControl(2).DatabaseName = App.Path + "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseConsistControl(3).DatabaseName = App.Path + "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseConsistControl(4).DatabaseName = App.Path + "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseConsistControlSelect.InitDir = App.Path + "\Databases"
  
  
' My Programming Notes
'
' Now that the clock window is opened the additional options on the menu bar can be enabled or disabled.
    
    Let MainScreen.MenuCommunicationPortOpen.Enabled = False
    Let MainScreen.MenuCommunicationPortClosed.Enabled = False
    Let MainScreen.MenuCommunicationPortSetting.Enabled = False
    Let MainScreen.MenuScaledTimeSetting.Enabled = False
    Let MainScreen.MenuProgrammingModeDiesel.Enabled = False
    Let MainScreen.MenuProgrammingModeSteam.Enabled = False
    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = False
    Let MainScreen.MenuProgrammingModeOther.Enabled = False
    Let MainScreen.MenuMainlineDieselProgramming.Enabled = False
    Let MainScreen.MenuMainlineSteamProgramming.Enabled = False
    Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = False
    Let MainScreen.MenuMainlineOtherProgramming.Enabled = False
    Let MainScreen.menumainlineconsist.Enabled = False
    Let MainScreen.menumainlinescalespeedoperation.Enabled = False
    Let MainScreen.MenuMainlineOperation.Enabled = True
    Let MainScreen.MenuMainlineOperationGUI.Enabled = True ' Temporary NMRA
    Let MainScreen.MenuMainlineMacroMaker.Enabled = False
    Let MainScreen.MenuLayoutDefineBlocks.Enabled = False
    Let MainScreen.MenuLayoutCTC.Enabled = False

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
'
' Every time a new window is opened in Autoamtic Train Control we check the screen size and compare it to the window screen size. If the window cannot be displayed in the current screen size a
' message box is displayed. This allows time for the user to change the screen attributes to correct size.

Do
    If Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY Then
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    End If
Loop While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY
    
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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


Private Sub LocomotiveConsistDefault_Click(Index As Integer)

If LocomotiveConsistDefault(Index).Value = vbChecked Then
    Let OneByteD(Index).BackColor = &HC0C0C0
    Let OneByteD(Index).Enabled = False
    Let OneByteD(Index).DataField = ""
    Let OneByteD(Index).Text = MainlineConsist.LocomotiveConsistNumber.Text
    Let TwoByteD(Index).BackColor = &HC0C0C0
    Let TwoByteD(Index).Enabled = False
    Let TwoByteD(Index).DataField = ""
    Let TwoByteD(Index).Text = ""
Else
    Let OneByteD(Index).BackColor = &H80000005
    Let OneByteD(Index).Enabled = True
    Let OneByteD(Index).DataField = "AddressByteOne"
    Let TwoByteD(Index).BackColor = &H80000005
    Let TwoByteD(Index).Enabled = True
    Let TwoByteD(Index).DataField = "AddressByteTwo"
End If

End Sub

Private Sub OneByteD_Change(Index As Integer)
    
Rem Hexidecimal Converstion
        
    Let OneByteH(Index).Text = Hex(Val(OneByteD(Index).Text))
    
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(OneByteH(Index).Text) = 1 Then
        Let OneByteH(Index).Text = "0" + OneByteH(Index).Text
    End If
    
Rem Binary Conversion
    
        Let temp = Val(OneByteD(Index).Text)
        
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

            Let SevenByteD(Y).Text = ""

                For X = 1 To 8
                
                    Let temp = 0


                        If OneByteD(Y).Text <> "" Then
                            If OneByteD(Y).Text <> "   " Then
                                If Mid$(OneByteB(Y).Text, X, 1) = "1" Then
                                    Let temp = temp + 1
                                End If
                            End If
                        End If
                        If TwoByteD(Y).Text <> "" Then
                            If TwoByteD(Y).Text <> "   " Then
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
                    Let SevenByteD(Y).Text = Val(SevenByteD(Y).Text) + bitvalue
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
            
            If OneByteD(Y).Text <> "" Then
                If OneByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + OneByteH(Y).Text
                End If
            End If
            If TwoByteD(Y).Text <> "" Then
                If TwoByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + TwoByteH(Y).Text
                End If
            End If
            If ThreeByteD(Y).Text <> "" Then
                If ThreeByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + ThreeByteH(Y).Text
                End If
            End If
            If FourByteD(Y).Text <> "" Then
                If FourByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + FourByteH(Y).Text
                End If
            End If
            If FiveByteD(Y).Text <> "" Then
                If FiveByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + FiveByteH(Y).Text
                End If
            End If
            If SixByteD(Y).Text <> "" Then
                If SixByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + SixByteH(Y).Text
                End If
            End If
            If SevenByteD(Y).Text <> "" Then
                If SevenByteD(Y).Text <> "   " Then
                    Let CommandControl = CommandControl + " " + SevenByteH(Y).Text
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

            Let CommunicationPortBusy(Y).Value = vbChecked

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

            MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent paramters as per National Model Railroader Association specification.
    
            Let LocomotiveCommunicationStatus(Y).Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

            While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
                Let temp = DoEvents
            Wend
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
            
                If OneByteD(Y).Text <> "" Then
                   If OneByteD(Y).Text <> "   " Then
                        Let CommandControl = CommandControl + " " + OneByteH(Y).Text
                    End If
                End If
                If TwoByteD(Y).Text <> "" Then
                    If TwoByteD(Y).Text <> "   " Then
                        Let CommandControl = CommandControl + " " + TwoByteH(Y).Text
                    End If
                End If
                'If ThreeByteD(y).Text <> "" Then
                '    If ThreeByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + ThreeByteH(y).Text
                '    End If
                'End If
                'If FourByteD(y).Text <> "" Then
                '    If FourByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + FourByteH(y).Text
                '    End If
                'End If
                'If FiveByteD(y).Text <> "" Then
                '    If FiveByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + FiveByteH(y).Text
                '    End If
                'End If
                'If SixByteD(y).Text <> "" Then
                '    If SixByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + SixByteH(y).Text
                '    End If
                'End If
                'If SevenByteD(y).Text <> "" Then
                '    If SevenByteD(y).Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + SevenByteH(y).Text
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

                Let CommunicationPortBusy(Y).Value = vbChecked

                MainScreen.MSComm1.Output = CommandControl
    
                Let LocomotiveCommunicationStatus(Y).Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

                While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
                    Let temp = DoEvents
                Wend
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
        
    Let TwoByteH(Index).Text = Hex(Val(TwoByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(TwoByteH(Index).Text) = 1 Then
        Let TwoByteH(Index).Text = "0" + TwoByteH(Index).Text
    End If

' This routine prforms the Binary Conversion
    
        Let temp = Val(TwoByteD(Index).Text)
        
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
        
    Let SevenByteH(Index).Text = Hex(Val(SevenByteD(Index).Text))
        
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SevenByteH(Index).Text) = 1 Then
        Let SevenByteH(Index).Text = "0" + SevenByteH(Index).Text
    End If

Rem Binary Conversion
    
        Let temp = Val(SevenByteD(Index).Text)
        
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




