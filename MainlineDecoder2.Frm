VERSION 4.00
Begin VB.Form MainlineDecoder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Programming for Decoders"
   ClientHeight    =   10785
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Height          =   11190
   Left            =   15
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Top             =   -30
   Width           =   8880
   Begin VB.Data DecoderDatabase 
      Caption         =   "DecoderDatabase"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCvs.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DecoderList"
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "&List"
      Height          =   255
      Left            =   120
      TabIndex        =   221
      Top             =   10440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      Picture         =   "MainlineDecoder.frx":0000
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   209
      Top             =   120
      Width           =   540
   End
   Begin VB.TextBox LocomotiveNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseDecoder"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   206
      Top             =   360
      Width           =   1335
   End
   Begin VB.Data LocomotiveDatabaseDecoder 
      Caption         =   "LocomotiveDatabaseDecoder"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\LocomotiveDatabaseDecoders.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDecoders"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CheckBox DisplayErrorMessage 
      Caption         =   "Display Error Messages"
      Height          =   195
      Left            =   10200
      TabIndex        =   122
      Top             =   3960
      Value           =   2  'Grayed
      Width           =   2295
   End
   Begin VB.TextBox OneByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   114
      Text            =   "00000000"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox OneByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   113
      Text            =   "00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox OneByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   112
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TwoByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   111
      Text            =   "00000000"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox TwoByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   110
      Text            =   "00"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox TwoByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   109
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   108
      Text            =   "00000000"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   107
      Text            =   "00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   106
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox FourByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   105
      Text            =   "00000000"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox FourByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   104
      Text            =   "00"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox FourByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   103
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox FiveByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   102
      Text            =   "00000000"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox FiveByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   101
      Text            =   "00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox FiveByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   100
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox SixByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   99
      Text            =   "00000000"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox SixByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   98
      Text            =   "00"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox SixByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   97
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox SevenByteB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   12840
      TabIndex        =   96
      Text            =   "00000000"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox SevenByteH 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      TabIndex        =   95
      Text            =   "00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox SevenByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10200
      TabIndex        =   94
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10200
      Top             =   4320
   End
   Begin VB.CommandButton ButtonExport 
      Caption         =   "&Export"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   10440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Identification"
      TabPicture(0)   =   "MainlineDecoder.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line9"
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(4)=   "Line1"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(13)=   "Label43"
      Tab(0).Control(14)=   "Line12"
      Tab(0).Control(15)=   "Label42"
      Tab(0).Control(16)=   "Label45"
      Tab(0).Control(17)=   "Line13"
      Tab(0).Control(18)=   "LocomotiveDecoderCVb(19)"
      Tab(0).Control(19)=   "LocomotiveDecoderCVh(19)"
      Tab(0).Control(20)=   "LocomotiveDecoderCVd(19)"
      Tab(0).Control(21)=   "LocomotiveDecoderCVc(18)"
      Tab(0).Control(22)=   "LocomotiveDecoderCVb(18)"
      Tab(0).Control(23)=   "LocomotiveDecoderCVh(18)"
      Tab(0).Control(24)=   "LocomotiveDecoderCVd(18)"
      Tab(0).Control(25)=   "LocomotiveDecoderCVb(17)"
      Tab(0).Control(26)=   "LocomotiveDecoderCVh(17)"
      Tab(0).Control(27)=   "LocomotiveDecoderCVd(17)"
      Tab(0).Control(28)=   "LocomotiveDecoderCVb(8)"
      Tab(0).Control(29)=   "LocomotiveDecoderCVh(8)"
      Tab(0).Control(30)=   "LocomotiveDecoderCVd(8)"
      Tab(0).Control(31)=   "LocomotiveDecoderCVb(7)"
      Tab(0).Control(32)=   "LocomotiveDecoderCVh(7)"
      Tab(0).Control(33)=   "LocomotiveDecoderCVd(7)"
      Tab(0).Control(34)=   "LocomotiveDecoderCVb(1)"
      Tab(0).Control(35)=   "LocomotiveDecoderCVh(1)"
      Tab(0).Control(36)=   "LocomotiveDecoderCVd(1)"
      Tab(0).Control(37)=   "LocomotiveDecoderCVd(105)"
      Tab(0).Control(38)=   "LocomotiveDecoderCVd(106)"
      Tab(0).Control(39)=   "LocomotiveDecoderCVh(105)"
      Tab(0).Control(40)=   "LocomotiveDecoderCVh(106)"
      Tab(0).Control(41)=   "LocomotiveDecoderCVb(105)"
      Tab(0).Control(42)=   "LocomotiveDecoderCVb(106)"
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "Voltage Configuration"
      TabPicture(1)   =   "MainlineDecoder.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label44"
      Tab(1).Control(1)=   "Line7"
      Tab(1).Control(2)=   "Label41"
      Tab(1).Control(3)=   "Line6"
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(5)=   "Label16"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(9)=   "Label12"
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(11)=   "LocomotiveDecoderCVd(65)"
      Tab(1).Control(12)=   "LocomotiveDecoderCVh(65)"
      Tab(1).Control(13)=   "LocomotiveDecoderCVb(65)"
      Tab(1).Control(14)=   "LocomotiveDecoderCVb(5)"
      Tab(1).Control(15)=   "LocomotiveDecoderCVh(5)"
      Tab(1).Control(16)=   "LocomotiveDecoderCVd(5)"
      Tab(1).Control(17)=   "LocomotiveDecoderCVb(6)"
      Tab(1).Control(18)=   "LocomotiveDecoderCVh(6)"
      Tab(1).Control(19)=   "LocomotiveDecoderCVd(6)"
      Tab(1).Control(20)=   "LocomotiveDecoderCVd(2)"
      Tab(1).Control(21)=   "LocomotiveDecoderCVh(2)"
      Tab(1).Control(22)=   "LocomotiveDecoderCVb(2)"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Speed Regulation"
      TabPicture(2)   =   "MainlineDecoder.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label40"
      Tab(2).Control(1)=   "Label39"
      Tab(2).Control(2)=   "Line4"
      Tab(2).Control(3)=   "Label25"
      Tab(2).Control(4)=   "Label24"
      Tab(2).Control(5)=   "Line3"
      Tab(2).Control(6)=   "Label23"
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(8)=   "Label21"
      Tab(2).Control(9)=   "Label20"
      Tab(2).Control(10)=   "Label19"
      Tab(2).Control(11)=   "Label18"
      Tab(2).Control(12)=   "Label77"
      Tab(2).Control(13)=   "LocomotiveDecoderCVd(95)"
      Tab(2).Control(14)=   "LocomotiveDecoderCVh(95)"
      Tab(2).Control(15)=   "LocomotiveDecoderCVb(95)"
      Tab(2).Control(16)=   "LocomotiveDecoderCVd(66)"
      Tab(2).Control(17)=   "LocomotiveDecoderCVh(66)"
      Tab(2).Control(18)=   "LocomotiveDecoderCVb(66)"
      Tab(2).Control(19)=   "ButtonSpeedTable"
      Tab(2).Control(20)=   "LocomotiveDecoderCVc(24)"
      Tab(2).Control(21)=   "LocomotiveDecoderCVc(23)"
      Tab(2).Control(22)=   "LocomotiveDecoderCVd(24)"
      Tab(2).Control(23)=   "LocomotiveDecoderCVh(24)"
      Tab(2).Control(24)=   "LocomotiveDecoderCVb(24)"
      Tab(2).Control(25)=   "LocomotiveDecoderCVd(23)"
      Tab(2).Control(26)=   "LocomotiveDecoderCVh(23)"
      Tab(2).Control(27)=   "LocomotiveDecoderCVb(23)"
      Tab(2).Control(28)=   "LocomotiveDecoderCVd(4)"
      Tab(2).Control(29)=   "LocomotiveDecoderCVh(4)"
      Tab(2).Control(30)=   "LocomotiveDecoderCVb(4)"
      Tab(2).Control(31)=   "LocomotiveDecoderCVd(3)"
      Tab(2).Control(32)=   "LocomotiveDecoderCVh(3)"
      Tab(2).Control(33)=   "LocomotiveDecoderCVb(3)"
      Tab(2).ControlCount=   34
      TabCaption(3)   =   "General Configuration"
      TabPicture(3)   =   "MainlineDecoder.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label37"
      Tab(3).Control(1)=   "Line5"
      Tab(3).Control(2)=   "Label28"
      Tab(3).Control(3)=   "Label27"
      Tab(3).Control(4)=   "Label26"
      Tab(3).Control(5)=   "Label29"
      Tab(3).Control(6)=   "LocomotiveDecoderCVd(29)"
      Tab(3).Control(7)=   "LocomotiveDecoderCVh(29)"
      Tab(3).Control(8)=   "LocomotiveDecoderCVb(29)"
      Tab(3).Control(9)=   "Command2"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Other Settings"
      TabPicture(4)   =   "MainlineDecoder.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label47"
      Tab(4).Control(1)=   "Label55"
      Tab(4).Control(2)=   "Label56"
      Tab(4).Control(3)=   "Label57"
      Tab(4).Control(4)=   "Label58"
      Tab(4).Control(5)=   "Label72"
      Tab(4).Control(6)=   "Label74"
      Tab(4).Control(7)=   "Label76"
      Tab(4).Control(8)=   "LocomotiveDecoderCVd(9)"
      Tab(4).Control(9)=   "LocomotiveDecoderCVh(9)"
      Tab(4).Control(10)=   "LocomotiveDecoderCVb(9)"
      Tab(4).Control(11)=   "LocomotiveDecoderCVd(10)"
      Tab(4).Control(12)=   "LocomotiveDecoderCVh(10)"
      Tab(4).Control(13)=   "LocomotiveDecoderCVb(10)"
      Tab(4).Control(14)=   "LocomotiveDecoderCVd(11)"
      Tab(4).Control(15)=   "LocomotiveDecoderCVh(11)"
      Tab(4).Control(16)=   "LocomotiveDecoderCVb(11)"
      Tab(4).Control(17)=   "LocomotiveDecoderCVd(12)"
      Tab(4).Control(18)=   "LocomotiveDecoderCVh(12)"
      Tab(4).Control(19)=   "LocomotiveDecoderCVb(12)"
      Tab(4).Control(20)=   "LocomotiveDecoderCVd(13)"
      Tab(4).Control(21)=   "LocomotiveDecoderCVh(13)"
      Tab(4).Control(22)=   "LocomotiveDecoderCVb(13)"
      Tab(4).ControlCount=   23
      TabCaption(5)   =   "Functions"
      TabPicture(5)   =   "MainlineDecoder.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label59"
      Tab(5).Control(1)=   "Label60"
      Tab(5).Control(2)=   "Label61"
      Tab(5).Control(3)=   "Label62"
      Tab(5).Control(4)=   "Label63"
      Tab(5).Control(5)=   "Label64"
      Tab(5).Control(6)=   "Label65"
      Tab(5).Control(7)=   "Label66"
      Tab(5).Control(8)=   "Label67"
      Tab(5).Control(9)=   "Line11"
      Tab(5).Control(10)=   "Label68"
      Tab(5).Control(11)=   "Label69"
      Tab(5).Control(12)=   "Label70"
      Tab(5).Control(13)=   "Label71"
      Tab(5).Control(14)=   "Label73"
      Tab(5).Control(15)=   "Label75"
      Tab(5).Control(16)=   "Line8"
      Tab(5).Control(17)=   "Label80"
      Tab(5).Control(18)=   "Label81"
      Tab(5).Control(19)=   "LocomotiveDecoderCVd(33)"
      Tab(5).Control(20)=   "LocomotiveDecoderCVd(34)"
      Tab(5).Control(21)=   "LocomotiveDecoderCVd(35)"
      Tab(5).Control(22)=   "LocomotiveDecoderCVd(36)"
      Tab(5).Control(23)=   "LocomotiveDecoderCVd(37)"
      Tab(5).Control(24)=   "LocomotiveDecoderCVd(38)"
      Tab(5).Control(25)=   "LocomotiveDecoderCVd(39)"
      Tab(5).Control(26)=   "LocomotiveDecoderCVd(40)"
      Tab(5).Control(27)=   "LocomotiveDecoderCVd(41)"
      Tab(5).Control(28)=   "LocomotiveDecoderCVd(42)"
      Tab(5).Control(29)=   "LocomotiveDecoderCVd(21)"
      Tab(5).Control(30)=   "LocomotiveDecoderCVd(22)"
      Tab(5).Control(31)=   "LocomotiveDecoderCVh(21)"
      Tab(5).Control(32)=   "LocomotiveDecoderCVh(22)"
      Tab(5).Control(33)=   "LocomotiveDecoderCVh(33)"
      Tab(5).Control(34)=   "LocomotiveDecoderCVh(34)"
      Tab(5).Control(35)=   "LocomotiveDecoderCVh(35)"
      Tab(5).Control(36)=   "LocomotiveDecoderCVh(36)"
      Tab(5).Control(37)=   "LocomotiveDecoderCVh(37)"
      Tab(5).Control(38)=   "LocomotiveDecoderCVh(38)"
      Tab(5).Control(39)=   "LocomotiveDecoderCVh(39)"
      Tab(5).Control(40)=   "LocomotiveDecoderCVh(40)"
      Tab(5).Control(41)=   "LocomotiveDecoderCVh(41)"
      Tab(5).Control(42)=   "LocomotiveDecoderCVh(42)"
      Tab(5).Control(43)=   "LocomotiveDecoderCVb(21)"
      Tab(5).Control(44)=   "LocomotiveDecoderCVb(22)"
      Tab(5).Control(45)=   "LocomotiveDecoderCVb(33)"
      Tab(5).Control(46)=   "LocomotiveDecoderCVb(34)"
      Tab(5).Control(47)=   "LocomotiveDecoderCVb(35)"
      Tab(5).Control(48)=   "LocomotiveDecoderCVb(36)"
      Tab(5).Control(49)=   "LocomotiveDecoderCVb(37)"
      Tab(5).Control(50)=   "LocomotiveDecoderCVb(38)"
      Tab(5).Control(51)=   "LocomotiveDecoderCVb(39)"
      Tab(5).Control(52)=   "LocomotiveDecoderCVb(40)"
      Tab(5).Control(53)=   "LocomotiveDecoderCVb(41)"
      Tab(5).Control(54)=   "LocomotiveDecoderCVb(42)"
      Tab(5).Control(55)=   "Command1"
      Tab(5).Control(56)=   "ButtonEasyScreenFunctions"
      Tab(5).ControlCount=   57
      TabCaption(6)   =   "Decoder Specific"
      TabPicture(6)   =   "MainlineDecoder.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label78"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label85"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label86"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label30"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "SSTab2"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "TextBoxManufacturerName"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "ComboBoxDecoderName"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "ButtonEasyScreen"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).ControlCount=   8
      Begin VB.CommandButton ButtonEasyScreen 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   7080
         TabIndex        =   222
         Top             =   9120
         Width           =   1335
      End
      Begin VB.ComboBox ComboBoxDecoderName 
         Height          =   315
         ItemData        =   "MainlineDecoder.frx":0506
         Left            =   4440
         List            =   "MainlineDecoder.frx":0508
         TabIndex        =   220
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   218
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton ButtonEasyScreenFunctions 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   217
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   216
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TextBoxManufacturerName 
         Height          =   285
         Left            =   1560
         TabIndex        =   212
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   106
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   204
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   105
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   203
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   106
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   202
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   105
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   201
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV106D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   106
         Left            =   -70920
         TabIndex        =   200
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV105D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   105
         Left            =   -70920
         TabIndex        =   199
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   -68040
         TabIndex        =   195
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   -68040
         TabIndex        =   194
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   -68040
         TabIndex        =   193
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   39
         Left            =   -68040
         TabIndex        =   192
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   38
         Left            =   -68040
         TabIndex        =   191
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   37
         Left            =   -68040
         TabIndex        =   190
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   36
         Left            =   -68040
         TabIndex        =   189
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   -68040
         TabIndex        =   188
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   -68040
         TabIndex        =   187
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   -68040
         TabIndex        =   186
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   -68040
         TabIndex        =   185
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   -68040
         TabIndex        =   184
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   -69480
         TabIndex        =   183
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   -69480
         TabIndex        =   182
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   -69480
         TabIndex        =   181
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   39
         Left            =   -69480
         TabIndex        =   180
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   38
         Left            =   -69480
         TabIndex        =   179
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   37
         Left            =   -69480
         TabIndex        =   178
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   36
         Left            =   -69480
         TabIndex        =   177
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   -69480
         TabIndex        =   176
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   -69480
         TabIndex        =   175
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   -69480
         TabIndex        =   174
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   -69480
         TabIndex        =   173
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   -69480
         TabIndex        =   167
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV22D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   22
         Left            =   -70920
         TabIndex        =   163
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV21D"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   21
         Left            =   -70920
         TabIndex        =   162
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV42D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   42
         Left            =   -70920
         TabIndex        =   153
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV41D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   41
         Left            =   -70920
         TabIndex        =   152
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV40D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   40
         Left            =   -70920
         TabIndex        =   151
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV39D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   39
         Left            =   -70920
         TabIndex        =   150
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV38D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   38
         Left            =   -70920
         TabIndex        =   149
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV37D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   37
         Left            =   -70920
         TabIndex        =   148
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV36D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   36
         Left            =   -70920
         TabIndex        =   147
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV35D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   35
         Left            =   -70920
         TabIndex        =   146
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV34D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   34
         Left            =   -70920
         TabIndex        =   145
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV33D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   33
         Left            =   -70920
         TabIndex        =   143
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV13D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   13
         Left            =   -70920
         TabIndex        =   139
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV12D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   12
         Left            =   -70920
         TabIndex        =   135
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV11D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   11
         Left            =   -70920
         TabIndex        =   131
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV10D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   10
         Left            =   -70920
         TabIndex        =   127
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV9D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   9
         Left            =   -70920
         TabIndex        =   123
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "LocomotiveDecoderCV1D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "LocomotiveDecoderCV7D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "LocomotiveDecoderCV8D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "LocomotiveDecoderCV17D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "LocomotiveDecoderCV18D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV19D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   19
         Left            =   -70920
         TabIndex        =   43
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "00000000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV2D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   2
         Left            =   -70920
         TabIndex        =   38
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV6D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   6
         Left            =   -70920
         TabIndex        =   37
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "00000000"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV5D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   5
         Left            =   -70920
         TabIndex        =   34
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "00000000"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "00000000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV3D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   3
         Left            =   -70920
         TabIndex        =   29
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "00000000"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV4D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   4
         Left            =   -70920
         TabIndex        =   26
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "00000000"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV23D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   23
         Left            =   -70920
         TabIndex        =   23
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "00000000"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV24D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   24
         Left            =   -70920
         TabIndex        =   20
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   23
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   24
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton ButtonSpeedTable 
         Caption         =   "Speed Table"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -68040
         TabIndex        =   17
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV29D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   29
         Left            =   -70920
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   66
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "00000000"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   66
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV66D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   66
         Left            =   -70920
         TabIndex        =   11
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   95
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "00000000"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   95
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV95D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   95
         Left            =   -70920
         TabIndex        =   8
         Text            =   "0"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   65
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "00000000"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   65
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV65D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   65
         Left            =   -70920
         TabIndex        =   5
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6855
         Left            =   120
         TabIndex        =   224
         Top             =   1920
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12091
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "CV 49 - 64"
         TabPicture(0)   =   "MainlineDecoder.frx":050A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LabelCVdescription(49)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LabelCVdescription(50)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LabelCVdescription(51)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LabelCVdescription(52)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LabelCVdescription(53)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LabelCVdescription(54)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LabelCVdescription(55)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LabelCVdescription(56)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "LabelCVdescription(57)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LabelCVdescription(58)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "LabelCVdescription(59)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "LabelCVdescription(60)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "LabelCVdescription(61)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "LabelCVdescription(62)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "LabelCVdescription(63)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "LabelCVdescription(64)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label31(49)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label31(50)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label31(51)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label31(52)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label31(53)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label31(54)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label31(55)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label31(56)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label31(57)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Label31(58)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Label31(59)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Label31(60)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Label31(61)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Label31(62)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Label31(63)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Label31(64)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "LocomotiveDecoderCVb(56)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "LocomotiveDecoderCVb(55)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "LocomotiveDecoderCVb(54)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "LocomotiveDecoderCVb(53)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "LocomotiveDecoderCVb(52)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "LocomotiveDecoderCVb(51)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "LocomotiveDecoderCVb(50)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "Option1(49)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "LocomotiveDecoderCVh(49)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "LocomotiveDecoderCVd(49)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "LocomotiveDecoderCVb(57)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "LocomotiveDecoderCVb(58)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "LocomotiveDecoderCVb(59)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "LocomotiveDecoderCVb(60)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "LocomotiveDecoderCVb(61)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "LocomotiveDecoderCVb(62)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "LocomotiveDecoderCVb(63)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "LocomotiveDecoderCVb(64)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "LocomotiveDecoderCVh(50)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "LocomotiveDecoderCVh(51)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "LocomotiveDecoderCVh(52)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "LocomotiveDecoderCVh(53)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "LocomotiveDecoderCVh(54)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "LocomotiveDecoderCVh(55)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "LocomotiveDecoderCVh(56)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "LocomotiveDecoderCVh(57)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "LocomotiveDecoderCVh(58)"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "LocomotiveDecoderCVh(59)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "LocomotiveDecoderCVh(60)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "LocomotiveDecoderCVh(61)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "LocomotiveDecoderCVh(62)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "LocomotiveDecoderCVh(63)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "LocomotiveDecoderCVh(64)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "LocomotiveDecoderCVd(50)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "LocomotiveDecoderCVd(51)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "LocomotiveDecoderCVd(52)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "LocomotiveDecoderCVd(53)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "LocomotiveDecoderCVd(54)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "LocomotiveDecoderCVd(55)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "LocomotiveDecoderCVd(56)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "LocomotiveDecoderCVd(57)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "LocomotiveDecoderCVd(58)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "LocomotiveDecoderCVd(59)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "LocomotiveDecoderCVd(60)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "LocomotiveDecoderCVd(61)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "LocomotiveDecoderCVd(62)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "LocomotiveDecoderCVd(63)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).Control(79)=   "LocomotiveDecoderCVd(64)"
         Tab(0).Control(79).Enabled=   0   'False
         Tab(0).Control(80)=   "Option1(50)"
         Tab(0).Control(80).Enabled=   0   'False
         Tab(0).Control(81)=   "Option1(51)"
         Tab(0).Control(81).Enabled=   0   'False
         Tab(0).Control(82)=   "Option1(52)"
         Tab(0).Control(82).Enabled=   0   'False
         Tab(0).Control(83)=   "Option1(53)"
         Tab(0).Control(83).Enabled=   0   'False
         Tab(0).Control(84)=   "Option1(54)"
         Tab(0).Control(84).Enabled=   0   'False
         Tab(0).Control(85)=   "Option1(55)"
         Tab(0).Control(85).Enabled=   0   'False
         Tab(0).Control(86)=   "Option1(56)"
         Tab(0).Control(86).Enabled=   0   'False
         Tab(0).Control(87)=   "Option1(57)"
         Tab(0).Control(87).Enabled=   0   'False
         Tab(0).Control(88)=   "Option1(58)"
         Tab(0).Control(88).Enabled=   0   'False
         Tab(0).Control(89)=   "Option1(59)"
         Tab(0).Control(89).Enabled=   0   'False
         Tab(0).Control(90)=   "Option1(60)"
         Tab(0).Control(90).Enabled=   0   'False
         Tab(0).Control(91)=   "Option1(61)"
         Tab(0).Control(91).Enabled=   0   'False
         Tab(0).Control(92)=   "Option1(62)"
         Tab(0).Control(92).Enabled=   0   'False
         Tab(0).Control(93)=   "Option1(63)"
         Tab(0).Control(93).Enabled=   0   'False
         Tab(0).Control(94)=   "Option1(64)"
         Tab(0).Control(94).Enabled=   0   'False
         Tab(0).Control(95)=   "LocomotiveDecoderCVb(49)"
         Tab(0).Control(95).Enabled=   0   'False
         Tab(0).Control(96)=   "SpecificCVs(50)"
         Tab(0).Control(96).Enabled=   0   'False
         Tab(0).Control(97)=   "SpecificCVs(51)"
         Tab(0).Control(97).Enabled=   0   'False
         Tab(0).Control(98)=   "SpecificCVs(52)"
         Tab(0).Control(98).Enabled=   0   'False
         Tab(0).Control(99)=   "SpecificCVs(53)"
         Tab(0).Control(99).Enabled=   0   'False
         Tab(0).Control(100)=   "SpecificCVs(54)"
         Tab(0).Control(100).Enabled=   0   'False
         Tab(0).Control(101)=   "SpecificCVs(55)"
         Tab(0).Control(101).Enabled=   0   'False
         Tab(0).Control(102)=   "SpecificCVs(56)"
         Tab(0).Control(102).Enabled=   0   'False
         Tab(0).Control(103)=   "SpecificCVs(57)"
         Tab(0).Control(103).Enabled=   0   'False
         Tab(0).Control(104)=   "SpecificCVs(58)"
         Tab(0).Control(104).Enabled=   0   'False
         Tab(0).Control(105)=   "SpecificCVs(59)"
         Tab(0).Control(105).Enabled=   0   'False
         Tab(0).Control(106)=   "SpecificCVs(60)"
         Tab(0).Control(106).Enabled=   0   'False
         Tab(0).Control(107)=   "SpecificCVs(61)"
         Tab(0).Control(107).Enabled=   0   'False
         Tab(0).Control(108)=   "SpecificCVs(62)"
         Tab(0).Control(108).Enabled=   0   'False
         Tab(0).Control(109)=   "SpecificCVs(63)"
         Tab(0).Control(109).Enabled=   0   'False
         Tab(0).Control(110)=   "SpecificCVs(64)"
         Tab(0).Control(110).Enabled=   0   'False
         Tab(0).Control(111)=   "SpecificCVs(49)"
         Tab(0).Control(111).Enabled=   0   'False
         Tab(0).ControlCount=   112
         TabCaption(1)   =   "CV112 - 128"
         TabPicture(1)   =   "MainlineDecoder.frx":0526
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SpecificCVs(128)"
         Tab(1).Control(1)=   "SpecificCVs(127)"
         Tab(1).Control(2)=   "SpecificCVs(126)"
         Tab(1).Control(3)=   "SpecificCVs(125)"
         Tab(1).Control(4)=   "SpecificCVs(124)"
         Tab(1).Control(5)=   "SpecificCVs(123)"
         Tab(1).Control(6)=   "SpecificCVs(122)"
         Tab(1).Control(7)=   "SpecificCVs(121)"
         Tab(1).Control(8)=   "SpecificCVs(120)"
         Tab(1).Control(9)=   "SpecificCVs(119)"
         Tab(1).Control(10)=   "SpecificCVs(118)"
         Tab(1).Control(11)=   "SpecificCVs(117)"
         Tab(1).Control(12)=   "SpecificCVs(116)"
         Tab(1).Control(13)=   "SpecificCVs(115)"
         Tab(1).Control(14)=   "SpecificCVs(114)"
         Tab(1).Control(15)=   "SpecificCVs(113)"
         Tab(1).Control(16)=   "SpecificCVs(112)"
         Tab(1).Control(17)=   "Option1(128)"
         Tab(1).Control(18)=   "Option1(127)"
         Tab(1).Control(19)=   "Option1(126)"
         Tab(1).Control(20)=   "Option1(125)"
         Tab(1).Control(21)=   "Option1(124)"
         Tab(1).Control(22)=   "Option1(123)"
         Tab(1).Control(23)=   "Option1(122)"
         Tab(1).Control(24)=   "Option1(121)"
         Tab(1).Control(25)=   "Option1(120)"
         Tab(1).Control(26)=   "Option1(119)"
         Tab(1).Control(27)=   "Option1(118)"
         Tab(1).Control(28)=   "Option1(117)"
         Tab(1).Control(29)=   "Option1(116)"
         Tab(1).Control(30)=   "Option1(115)"
         Tab(1).Control(31)=   "Option1(114)"
         Tab(1).Control(32)=   "Option1(113)"
         Tab(1).Control(33)=   "LocomotiveDecoderCVd(128)"
         Tab(1).Control(34)=   "LocomotiveDecoderCVd(127)"
         Tab(1).Control(35)=   "LocomotiveDecoderCVd(126)"
         Tab(1).Control(36)=   "LocomotiveDecoderCVd(125)"
         Tab(1).Control(37)=   "LocomotiveDecoderCVd(124)"
         Tab(1).Control(38)=   "LocomotiveDecoderCVd(123)"
         Tab(1).Control(39)=   "LocomotiveDecoderCVd(122)"
         Tab(1).Control(40)=   "LocomotiveDecoderCVd(121)"
         Tab(1).Control(41)=   "LocomotiveDecoderCVd(120)"
         Tab(1).Control(42)=   "LocomotiveDecoderCVd(119)"
         Tab(1).Control(43)=   "LocomotiveDecoderCVd(118)"
         Tab(1).Control(44)=   "LocomotiveDecoderCVd(117)"
         Tab(1).Control(45)=   "LocomotiveDecoderCVd(116)"
         Tab(1).Control(46)=   "LocomotiveDecoderCVd(115)"
         Tab(1).Control(47)=   "LocomotiveDecoderCVd(114)"
         Tab(1).Control(48)=   "LocomotiveDecoderCVd(113)"
         Tab(1).Control(49)=   "LocomotiveDecoderCVh(128)"
         Tab(1).Control(50)=   "LocomotiveDecoderCVh(127)"
         Tab(1).Control(51)=   "LocomotiveDecoderCVh(126)"
         Tab(1).Control(52)=   "LocomotiveDecoderCVh(125)"
         Tab(1).Control(53)=   "LocomotiveDecoderCVh(124)"
         Tab(1).Control(54)=   "LocomotiveDecoderCVh(123)"
         Tab(1).Control(55)=   "LocomotiveDecoderCVh(122)"
         Tab(1).Control(56)=   "LocomotiveDecoderCVh(121)"
         Tab(1).Control(57)=   "LocomotiveDecoderCVh(120)"
         Tab(1).Control(58)=   "LocomotiveDecoderCVh(119)"
         Tab(1).Control(59)=   "LocomotiveDecoderCVh(118)"
         Tab(1).Control(60)=   "LocomotiveDecoderCVh(117)"
         Tab(1).Control(61)=   "LocomotiveDecoderCVh(116)"
         Tab(1).Control(62)=   "LocomotiveDecoderCVh(115)"
         Tab(1).Control(63)=   "LocomotiveDecoderCVh(114)"
         Tab(1).Control(64)=   "LocomotiveDecoderCVh(113)"
         Tab(1).Control(65)=   "LocomotiveDecoderCVb(128)"
         Tab(1).Control(66)=   "LocomotiveDecoderCVb(127)"
         Tab(1).Control(67)=   "LocomotiveDecoderCVb(126)"
         Tab(1).Control(68)=   "LocomotiveDecoderCVb(125)"
         Tab(1).Control(69)=   "LocomotiveDecoderCVb(124)"
         Tab(1).Control(70)=   "LocomotiveDecoderCVb(123)"
         Tab(1).Control(71)=   "LocomotiveDecoderCVb(122)"
         Tab(1).Control(72)=   "LocomotiveDecoderCVb(121)"
         Tab(1).Control(73)=   "LocomotiveDecoderCVb(120)"
         Tab(1).Control(74)=   "LocomotiveDecoderCVb(119)"
         Tab(1).Control(75)=   "LocomotiveDecoderCVb(118)"
         Tab(1).Control(76)=   "LocomotiveDecoderCVb(117)"
         Tab(1).Control(77)=   "LocomotiveDecoderCVb(116)"
         Tab(1).Control(78)=   "LocomotiveDecoderCVb(115)"
         Tab(1).Control(79)=   "LocomotiveDecoderCVb(114)"
         Tab(1).Control(80)=   "LocomotiveDecoderCVb(113)"
         Tab(1).Control(81)=   "LocomotiveDecoderCVb(112)"
         Tab(1).Control(82)=   "LocomotiveDecoderCVh(112)"
         Tab(1).Control(83)=   "LocomotiveDecoderCVd(112)"
         Tab(1).Control(84)=   "Option1(112)"
         Tab(1).Control(85)=   "LabelCVdescription(128)"
         Tab(1).Control(86)=   "LabelCVdescription(127)"
         Tab(1).Control(87)=   "LabelCVdescription(126)"
         Tab(1).Control(88)=   "LabelCVdescription(125)"
         Tab(1).Control(89)=   "LabelCVdescription(124)"
         Tab(1).Control(90)=   "LabelCVdescription(123)"
         Tab(1).Control(91)=   "LabelCVdescription(122)"
         Tab(1).Control(92)=   "LabelCVdescription(121)"
         Tab(1).Control(93)=   "LabelCVdescription(120)"
         Tab(1).Control(94)=   "LabelCVdescription(119)"
         Tab(1).Control(95)=   "LabelCVdescription(118)"
         Tab(1).Control(96)=   "LabelCVdescription(117)"
         Tab(1).Control(97)=   "LabelCVdescription(116)"
         Tab(1).Control(98)=   "LabelCVdescription(115)"
         Tab(1).Control(99)=   "LabelCVdescription(114)"
         Tab(1).Control(100)=   "LabelCVdescription(113)"
         Tab(1).Control(101)=   "LabelCVdescription(112)"
         Tab(1).Control(102)=   "Label31(128)"
         Tab(1).Control(103)=   "Label31(127)"
         Tab(1).Control(104)=   "Label31(126)"
         Tab(1).Control(105)=   "Label31(125)"
         Tab(1).Control(106)=   "Label31(124)"
         Tab(1).Control(107)=   "Label31(123)"
         Tab(1).Control(108)=   "Label31(122)"
         Tab(1).Control(109)=   "Label31(121)"
         Tab(1).Control(110)=   "Label31(120)"
         Tab(1).Control(111)=   "Label31(119)"
         Tab(1).Control(112)=   "Label31(118)"
         Tab(1).Control(113)=   "Label31(117)"
         Tab(1).Control(114)=   "Label31(116)"
         Tab(1).Control(115)=   "Label31(115)"
         Tab(1).Control(116)=   "Label31(114)"
         Tab(1).Control(117)=   "Label31(113)"
         Tab(1).Control(118)=   "Label31(112)"
         Tab(1).ControlCount=   119
         TabCaption(2)   =   "Additional CVs"
         TabPicture(2)   =   "MainlineDecoder.frx":0542
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   128
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV128"
            Top             =   6360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   127
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV127"
            Top             =   6000
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   126
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV126"
            Top             =   5640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   125
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV125"
            Top             =   5280
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   124
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV124"
            Top             =   4920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   123
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV123"
            Top             =   4560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   122
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV122"
            Top             =   4200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   121
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV121"
            Top             =   3840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   120
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV120"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   119
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV119"
            Top             =   3120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   118
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV118"
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   117
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV117"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   116
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV116"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   115
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV115"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   114
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV114"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   113
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV113"
            Top             =   960
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   112
            Left            =   -72720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV112"
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   49
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV49"
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   64
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV64"
            Top             =   6000
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   63
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV63"
            Top             =   5640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   62
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV62"
            Top             =   5280
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   61
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV61"
            Top             =   4920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   60
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV60"
            Top             =   4560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   59
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV59"
            Top             =   4200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   58
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV58"
            Top             =   3840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   57
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV57"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   56
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV56"
            Top             =   3120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   55
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV55"
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   54
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV54"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   53
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV53"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   52
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV52"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   51
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV51"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
            Exclusive       =   0   'False
            Height          =   300
            Index           =   50
            Left            =   2280
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "CV50"
            Top             =   960
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   128
            Left            =   -74880
            TabIndex        =   356
            Top             =   6360
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   127
            Left            =   -74880
            TabIndex        =   355
            Top             =   6000
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   126
            Left            =   -74880
            TabIndex        =   354
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   125
            Left            =   -74880
            TabIndex        =   353
            Top             =   5280
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   124
            Left            =   -74880
            TabIndex        =   352
            Top             =   4920
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   123
            Left            =   -74880
            TabIndex        =   351
            Top             =   4560
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   122
            Left            =   -74880
            TabIndex        =   350
            Top             =   4200
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   121
            Left            =   -74880
            TabIndex        =   349
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   120
            Left            =   -74880
            TabIndex        =   348
            Top             =   3480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   119
            Left            =   -74880
            TabIndex        =   347
            Top             =   3120
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   118
            Left            =   -74880
            TabIndex        =   346
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   117
            Left            =   -74880
            TabIndex        =   345
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   116
            Left            =   -74880
            TabIndex        =   344
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   115
            Left            =   -74880
            TabIndex        =   343
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   114
            Left            =   -74880
            TabIndex        =   342
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   113
            Left            =   -74880
            TabIndex        =   341
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV128D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   128
            Left            =   -71040
            TabIndex        =   340
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV127D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   127
            Left            =   -71040
            TabIndex        =   339
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV126D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   126
            Left            =   -71040
            TabIndex        =   338
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV125D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   125
            Left            =   -71040
            TabIndex        =   337
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV124D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   124
            Left            =   -71040
            TabIndex        =   336
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV123D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   123
            Left            =   -71040
            TabIndex        =   335
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV122D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   122
            Left            =   -71040
            TabIndex        =   334
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV121D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   121
            Left            =   -71040
            TabIndex        =   333
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV120D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   120
            Left            =   -71040
            TabIndex        =   332
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV119D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   119
            Left            =   -71040
            TabIndex        =   331
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV118D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   118
            Left            =   -71040
            TabIndex        =   330
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV117D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   117
            Left            =   -71040
            TabIndex        =   329
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV116D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   116
            Left            =   -71040
            TabIndex        =   328
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV115D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   115
            Left            =   -71040
            TabIndex        =   327
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV114D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   114
            Left            =   -71040
            TabIndex        =   326
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV113D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   113
            Left            =   -71040
            TabIndex        =   325
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   128
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   324
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   127
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   323
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   126
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   322
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   125
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   321
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   124
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   320
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   123
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   319
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   122
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   318
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   121
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   317
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   120
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   316
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   119
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   315
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   118
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   314
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   117
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   313
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   116
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   312
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   115
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   311
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   114
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   310
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   113
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   309
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   128
            Left            =   -68160
            TabIndex        =   308
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   127
            Left            =   -68160
            TabIndex        =   307
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   126
            Left            =   -68160
            TabIndex        =   306
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   125
            Left            =   -68160
            TabIndex        =   305
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   124
            Left            =   -68160
            TabIndex        =   304
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   123
            Left            =   -68160
            TabIndex        =   303
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   122
            Left            =   -68160
            TabIndex        =   302
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   121
            Left            =   -68160
            TabIndex        =   301
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   120
            Left            =   -68160
            TabIndex        =   300
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   119
            Left            =   -68160
            TabIndex        =   299
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   118
            Left            =   -68160
            TabIndex        =   298
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   117
            Left            =   -68160
            TabIndex        =   297
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   116
            Left            =   -68160
            TabIndex        =   296
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   115
            Left            =   -68160
            TabIndex        =   295
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   114
            Left            =   -68160
            TabIndex        =   294
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   113
            Left            =   -68160
            TabIndex        =   293
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   112
            Left            =   -68160
            TabIndex        =   292
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   6840
            TabIndex        =   291
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   112
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   290
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV112D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   112
            Left            =   -71040
            TabIndex        =   289
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   112
            Left            =   -74880
            TabIndex        =   288
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   287
            Top             =   6000
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   286
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   285
            Top             =   5280
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   284
            Top             =   4920
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   283
            Top             =   4560
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   282
            Top             =   4200
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   281
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   280
            Top             =   3480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   279
            Top             =   3120
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   278
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   277
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   276
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   275
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   120
            TabIndex        =   274
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   273
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV64D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   64
            Left            =   3960
            TabIndex        =   272
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV63D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   3960
            TabIndex        =   271
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV62D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   62
            Left            =   3960
            TabIndex        =   270
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV61D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   61
            Left            =   3960
            TabIndex        =   269
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV60D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   60
            Left            =   3960
            TabIndex        =   268
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV59D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   3960
            TabIndex        =   267
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV58D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   58
            Left            =   3960
            TabIndex        =   266
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV57D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   57
            Left            =   3960
            TabIndex        =   265
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV56D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   3960
            TabIndex        =   264
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV55D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   55
            Left            =   3960
            TabIndex        =   263
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV54D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   3960
            TabIndex        =   262
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV53D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   3960
            TabIndex        =   261
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV52D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   3960
            TabIndex        =   260
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV51D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   51
            Left            =   3960
            TabIndex        =   259
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV50D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   3960
            TabIndex        =   258
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   64
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   257
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   256
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   62
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   255
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   61
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   254
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   60
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   252
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   58
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   251
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   57
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   250
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   249
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   55
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   248
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   247
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   246
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   245
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   51
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   244
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   243
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   64
            Left            =   6840
            TabIndex        =   242
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   63
            Left            =   6840
            TabIndex        =   241
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   62
            Left            =   6840
            TabIndex        =   240
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   61
            Left            =   6840
            TabIndex        =   239
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   60
            Left            =   6840
            TabIndex        =   238
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   6840
            TabIndex        =   237
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   58
            Left            =   6840
            TabIndex        =   236
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   57
            Left            =   6840
            TabIndex        =   235
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVd 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "LocomotiveDecoderCV49D"
            DataSource      =   "LocomotiveDatabaseDecoder"
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   3960
            TabIndex        =   234
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   233
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   232
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   6840
            TabIndex        =   231
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   51
            Left            =   6840
            TabIndex        =   230
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   6840
            TabIndex        =   229
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   6840
            TabIndex        =   228
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   6840
            TabIndex        =   227
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   55
            Left            =   6840
            TabIndex        =   226
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   6840
            TabIndex        =   225
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(128)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   128
            Left            =   -74040
            TabIndex        =   422
            Top             =   6360
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(127)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   127
            Left            =   -74040
            TabIndex        =   421
            Top             =   6000
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(126)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   126
            Left            =   -74040
            TabIndex        =   420
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(125)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   125
            Left            =   -74040
            TabIndex        =   419
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(124)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   124
            Left            =   -74040
            TabIndex        =   418
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(123)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   123
            Left            =   -74040
            TabIndex        =   417
            Top             =   4560
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(122)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   122
            Left            =   -74040
            TabIndex        =   416
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(121)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   121
            Left            =   -74040
            TabIndex        =   415
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(120)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   120
            Left            =   -74040
            TabIndex        =   414
            Top             =   3480
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(119)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   119
            Left            =   -74040
            TabIndex        =   413
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(118)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   118
            Left            =   -74040
            TabIndex        =   412
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(117)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   117
            Left            =   -74040
            TabIndex        =   411
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(116)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   116
            Left            =   -74040
            TabIndex        =   410
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(115)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   115
            Left            =   -74040
            TabIndex        =   409
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(114)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   114
            Left            =   -74040
            TabIndex        =   408
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(113)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   113
            Left            =   -74040
            TabIndex        =   407
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(112)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   112
            Left            =   -74040
            TabIndex        =   406
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label31 
            Caption         =   "CV128 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   128
            Left            =   -74640
            TabIndex        =   405
            Top             =   6360
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV127 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   127
            Left            =   -74640
            TabIndex        =   404
            Top             =   6000
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV126 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   126
            Left            =   -74640
            TabIndex        =   403
            Top             =   5640
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV125 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   125
            Left            =   -74640
            TabIndex        =   402
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV124 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   124
            Left            =   -74640
            TabIndex        =   401
            Top             =   4920
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV123 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   123
            Left            =   -74640
            TabIndex        =   400
            Top             =   4560
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV122 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   122
            Left            =   -74640
            TabIndex        =   399
            Top             =   4200
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV121 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   121
            Left            =   -74640
            TabIndex        =   398
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV120 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   120
            Left            =   -74640
            TabIndex        =   397
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV119 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   119
            Left            =   -74640
            TabIndex        =   396
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV118 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   118
            Left            =   -74640
            TabIndex        =   395
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV117 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   117
            Left            =   -74640
            TabIndex        =   394
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV116 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   116
            Left            =   -74640
            TabIndex        =   393
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV115 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   115
            Left            =   -74640
            TabIndex        =   392
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV114 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   114
            Left            =   -74640
            TabIndex        =   391
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV113 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   113
            Left            =   -74640
            TabIndex        =   390
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV112 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   112
            Left            =   -74640
            TabIndex        =   389
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV64 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   360
            TabIndex        =   388
            Top             =   6000
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV63 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   360
            TabIndex        =   387
            Top             =   5640
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV62 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   360
            TabIndex        =   386
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV61 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   360
            TabIndex        =   385
            Top             =   4920
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV60 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   360
            TabIndex        =   384
            Top             =   4560
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV59 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   360
            TabIndex        =   383
            Top             =   4200
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV58 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   360
            TabIndex        =   382
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV57 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   360
            TabIndex        =   381
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV56 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   360
            TabIndex        =   380
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV55 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   360
            TabIndex        =   379
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV54 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   360
            TabIndex        =   378
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV53 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   360
            TabIndex        =   377
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV52 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   360
            TabIndex        =   376
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV51 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   360
            TabIndex        =   375
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV50 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   360
            TabIndex        =   374
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "CV49 -"
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   360
            TabIndex        =   373
            Top             =   600
            Width           =   495
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(64)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   960
            TabIndex        =   372
            Top             =   6000
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(63)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   960
            TabIndex        =   371
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(62)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   960
            TabIndex        =   370
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(61)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   960
            TabIndex        =   369
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(60)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   960
            TabIndex        =   368
            Top             =   4560
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(59)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   960
            TabIndex        =   367
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(58)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   960
            TabIndex        =   366
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(57)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   960
            TabIndex        =   365
            Top             =   3480
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(56)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   960
            TabIndex        =   364
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(55)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   960
            TabIndex        =   363
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(54)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   960
            TabIndex        =   362
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(53)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   960
            TabIndex        =   361
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(52)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   960
            TabIndex        =   360
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(51)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   960
            TabIndex        =   359
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(50)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   960
            TabIndex        =   358
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label LabelCVdescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(49)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   960
            TabIndex        =   357
            Top             =   600
            Width           =   2895
         End
      End
      Begin VB.Label Label30 
         Caption         =   $"MainlineDecoder.frx":055E
         Height          =   495
         Left            =   240
         TabIndex        =   223
         Top             =   8880
         Width           =   6495
      End
      Begin VB.Label Label29 
         Caption         =   "For additional help with setting the configuration variables for CV29, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   219
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label Label81 
         Caption         =   "For additional help with setting the configuration variables for functions, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   215
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label Label80 
         Caption         =   "For additional help with setting the configuration variables for functions, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   214
         Top             =   6480
         Width           =   6495
      End
      Begin VB.Line Line8 
         X1              =   -66720
         X2              =   -74760
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label86 
         Caption         =   "Decoder Type"
         Height          =   255
         Left            =   4440
         TabIndex        =   213
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label85 
         Caption         =   "Manufacturer"
         Height          =   255
         Left            =   1560
         TabIndex        =   211
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label78 
         Caption         =   $"MainlineDecoder.frx":0605
         Height          =   375
         Left            =   360
         TabIndex        =   210
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "Speed Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   208
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   205
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line13 
         X1              =   -66720
         X2              =   -74760
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "CV106 - User Identifier 2"
         Height          =   195
         Left            =   -74760
         TabIndex        =   198
         Top             =   4320
         Width           =   1725
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "CV105 - User Identifier 1"
         Height          =   195
         Left            =   -74760
         TabIndex        =   197
         Top             =   3960
         Width           =   1725
      End
      Begin VB.Line Line12 
         X1              =   -66720
         X2              =   -74760
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Identification Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   196
         Top             =   4680
         Width           =   1410
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   172
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   171
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   170
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   169
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   168
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "CV21 - Active for Consist Functions 1-8"
         Height          =   195
         Left            =   -74760
         TabIndex        =   166
         Top             =   1320
         Width           =   2760
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "CV22 - Active for Consist Function Light"
         Height          =   195
         Left            =   -74760
         TabIndex        =   165
         Top             =   1680
         Width           =   2805
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "CV42 - Output for Function Eight"
         Height          =   195
         Left            =   -74760
         TabIndex        =   164
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Line Line11 
         X1              =   -74760
         X2              =   -66720
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "CV41 - Output for Function Seven"
         Height          =   195
         Left            =   -74760
         TabIndex        =   161
         Top             =   5640
         Width           =   2400
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "CV40- Output for Function Six"
         Height          =   195
         Left            =   -74760
         TabIndex        =   160
         Top             =   5280
         Width           =   2100
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "CV39 - Output for Function Five"
         Height          =   195
         Left            =   -74760
         TabIndex        =   159
         Top             =   4920
         Width           =   2235
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "CV38 - Output for Function Four"
         Height          =   195
         Left            =   -74760
         TabIndex        =   158
         Top             =   4560
         Width           =   2250
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "CV37 - Output Function Three"
         Height          =   195
         Left            =   -74760
         TabIndex        =   157
         Top             =   4200
         Width           =   2130
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "CV36 - Output Function Two"
         Height          =   195
         Left            =   -74760
         TabIndex        =   156
         Top             =   3840
         Width           =   2025
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "CV35 - Output for Function One"
         Height          =   195
         Left            =   -74760
         TabIndex        =   155
         Top             =   3480
         Width           =   2235
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "CV34 - Output for Rear Light"
         Height          =   195
         Left            =   -74760
         TabIndex        =   154
         Top             =   3120
         Width           =   2010
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "CV33 - Output for Funtion Front Light"
         Height          =   195
         Left            =   -74760
         TabIndex        =   144
         Top             =   2760
         Width           =   2595
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "CV13 - Analog Mode Function Status"
         Height          =   195
         Left            =   -74760
         TabIndex        =   142
         Top             =   2640
         Width           =   2625
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "CV12 - Power Source Conversion"
         Height          =   195
         Left            =   -74760
         TabIndex        =   138
         Top             =   2280
         Width           =   2370
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "CV11 - Packet Time Out Value"
         Height          =   195
         Left            =   -74760
         TabIndex        =   132
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "CV10 - EMF  FeedBack Cut out"
         Height          =   195
         Left            =   -74760
         TabIndex        =   130
         Top             =   1560
         Width           =   2235
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "CV9 - Total PWM Period"
         Height          =   195
         Left            =   -74760
         TabIndex        =   124
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CV1 - Short Address"
         Height          =   195
         Left            =   -74760
         TabIndex        =   92
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   91
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   90
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   89
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CV7 - Manufacturer Version Number"
         Height          =   195
         Left            =   -74760
         TabIndex        =   88
         Top             =   3000
         Width           =   2550
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CV8 - Manufacturer ID Number"
         Height          =   195
         Left            =   -74760
         TabIndex        =   87
         Top             =   3360
         Width           =   2190
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CV17 - Extended Address"
         Height          =   195
         Left            =   -74760
         TabIndex        =   86
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CV18 - Extended Address"
         Height          =   195
         Left            =   -74760
         TabIndex        =   85
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   -72360
         TabIndex        =   84
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "CV19 - Consist Address"
         Height          =   195
         Left            =   -74760
         TabIndex        =   83
         Top             =   2520
         Width           =   1650
      End
      Begin VB.Line Line2 
         X1              =   -66720
         X2              =   -74760
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   82
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "CV2 - Starting Voltage"
         Height          =   195
         Left            =   -74760
         TabIndex        =   81
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   80
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   79
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   -72360
         TabIndex        =   78
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "CV6 - Mid Voltage"
         Height          =   195
         Left            =   -74760
         TabIndex        =   77
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "CV5 - High Voltage"
         Height          =   195
         Left            =   -74760
         TabIndex        =   76
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   75
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   74
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   73
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   -72360
         TabIndex        =   72
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "CV3 - Acceleration Rate"
         Height          =   195
         Left            =   -74760
         TabIndex        =   71
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "CV4 - Deceleration Rate"
         Height          =   195
         Left            =   -74760
         TabIndex        =   70
         Top             =   2040
         Width           =   1725
      End
      Begin VB.Line Line3 
         X1              =   -66720
         X2              =   -74760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "CV23 - Acceleration Adjustment"
         Height          =   195
         Left            =   -74760
         TabIndex        =   69
         Top             =   1560
         Width           =   2235
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "CV24 - Deceleration Adjustment"
         Height          =   195
         Left            =   -74760
         TabIndex        =   68
         Top             =   2400
         Width           =   2250
      End
      Begin VB.Line Line4 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   67
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   66
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "CV29  - Configuration"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Line Line5 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   64
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line6 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "CV66 - Forward Speed Trim"
         Height          =   195
         Left            =   -74760
         TabIndex        =   63
         Top             =   2880
         Width           =   1950
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "CV95 - Reverse Trim"
         Height          =   195
         Left            =   -74760
         TabIndex        =   62
         Top             =   3240
         Width           =   1470
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "CV65 - Kick Start"
         Height          =   195
         Left            =   -74760
         TabIndex        =   61
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Line Line7 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Voltage Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   60
         Top             =   2880
         Width           =   1050
      End
      Begin VB.Line Line9 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.Label Label79 
      AutoSize        =   -1  'True
      Caption         =   "Locomotive Number"
      Height          =   195
      Left            =   1200
      TabIndex        =   207
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label54 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   9120
      TabIndex        =   121
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   9120
      TabIndex        =   120
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label Label52 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Third Byte"
      Height          =   195
      Left            =   9240
      TabIndex        =   119
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fourth Byte"
      Height          =   195
      Left            =   9240
      TabIndex        =   118
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fifth Byte"
      Height          =   195
      Left            =   9360
      TabIndex        =   117
      Top             =   2760
      Width           =   660
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sixth Byte"
      Height          =   195
      Left            =   9360
      TabIndex        =   116
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Error Byte"
      Height          =   195
      Left            =   9360
      TabIndex        =   115
      Top             =   3480
      Width           =   690
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      Caption         =   "Update all your general configuration variables on the mainline"
      Height          =   195
      Left            =   3240
      TabIndex        =   93
      Top             =   120
      Width           =   4380
   End
   Begin VB.Label LocomotiveCommunicationStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "MainlineDecoder"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim cv(1 To 128) As String


Private Sub ButtonClose_Click()

    MainlineDecoder.Hide
    Unload MainlineDecoder

End Sub


Private Sub ButtonEasyScreen_Click()

temp = 0
For t = 49 To 64
    If Option1(t).Value = True Then
        Let temp = 1
    End If
Next t

For t = 112 To 128
    If Option1(t).Value = True Then
        Let temp = 1
    End If
Next t

If temp = 1 Then
    Load MainlineEasyScreenSpecificCvs
    MainlineEasyScreenSpecificCvs.Show 1
End If

End Sub

Private Sub ButtonEasyScreenFunctions_Click()

' =========================================================================================================================
' Loading the Programming Screen for 'Easy screen' Functions
'

' =========================================================================================================================
' Load Statement
'
' Loads a form or control into memory.
'
' Syntax is as the folowing:
'
'   Load object
'
' The object placeholder is the name of a Form object, MDIForm object, or control array element to load. You don 't need
' to use the Load statement with forms unless you want to load a form without displaying it.  Any reference to a form
' (except in a Set or If...TypeOf statement) automatically loads it if it's not already loaded.  For example, the Show
' method loads a form before displaying it.  Once the form is loaded, its properties and controls can be altered by the
' application, whether or not the form is actually visible.  Under some circumstances, you may want to load all your
' forms during initialization and display them later as they're needed. When more than one instance of a form is loaded,
' each form has its own set of properties and variables that can be changed independently from all other instances of the
' form. When Visual Basic loads a Form object, it sets form properties to their initial values and then performs the
' Load event procedure.  When an application starts, Visual Basic automatically loads and displays the application's
' startup form. If you load a Form with the MDIChild property set to True before loading the MDIForm, the MDIForm is
' automatically loaded before the child form.  Loading an MDI child form causes it to be visible immediately after
' the Form_Load event procedure ends because MDI child forms can't be hidden. The standard dialog boxes produced by
' Visual Basic functions such as MsgBox and InputBox do not need to be loaded, shown, or unloaded, but can simply be
' invoked directly.

    Load MainlineEasyScreenFunctions

' =========================================================================================================================
' Show Method
'
'Displays an MDIForm or Form object.  Doesn't support named arguments. For Visual Basic in-process OLE server (DLL)
' considerations and restrictions that apply to this topic, see Make OLE DLL File Command.
'
' Syntax is as the following:
'
'   object.Show Style
'
' The Show method syntax has these parts where 'object' is optional. An object expression that evaluates to an object in
' the Applies To list. If object is omitted, the form associated with the active form module is assumed to be object. The
' 'style' is optional. Integer that determines if the form is modal or modeless.  If style is 0, the form is modeless;
' if style is 1, the form is modal. If the specified form isn't loaded when the Show method is invoked, Visual Basic
' automatically loads it. When Show displays a modeless form, subsequent code is executed as it's
' encountered.  When Show displays a modal form, no subsequent code is executed until the form is hidden or
' unloaded. When Show displays a modal form, no input (keyboard or mouse click) can occur except to objects on the
' modal form.  The program must hide or unload a modal form (usually in response to some user action) before input
' to another form can occur.  An MDIForm can't be modal. Although other forms in your application are disabled when
' a modal form is displayed, other applications aren't. The startup form of an application is automatically shown
' after its Load event is invoked.

    MainlineEasyScreenFunctions.Show 0

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

Private Sub ButtonExport_Click()

End Sub

Private Sub ButtonList_Click()

Load LocomotiveCVSpreadsheet
LocomotiveCVSpreadsheet.Show 0

End Sub

Private Sub ButtonSpeedTable_Click()

    Let MousePointer = vbHourglass
    Load MainlineSpeedTable
    MainlineSpeedTable.Show 0
    Let MousePointer = vbDefault

End Sub








Private Sub ButtonUpdate_Click()
    
    LocomotiveDatabaseDecoder.Recordset.Edit
    LocomotiveDatabaseDecoder.Recordset.Update

End Sub




Private Sub ComboBoxDecoderName_Click()

Let MousePointer = vbArrowHourglass

DecoderDatabase.Recordset.MoveFirst

Do While Not (DecoderDatabase.Recordset.Fields("Manufacturer") = TextBoxManufacturerName.Text And DecoderDatabase.Recordset.Fields("Decoder") = ComboBoxDecoderName.Text)
DecoderDatabase.Recordset.MoveNext
Loop

For tempDatabase = 49 To 64
    SpecificCVs(tempDatabase).Recordset.MoveFirst
    While SpecificCVs(tempDatabase).Recordset.Fields("RecordNumber") <> DecoderDatabase.Recordset.Fields("RecordNumber")
        SpecificCVs(tempDatabase).Recordset.MoveNext
    Wend
    DoEvents
    If SpecificCVs(tempDatabase).Recordset.Fields("Used") = True Then
        Let cv(tempDatabase) = "Yes"
        Let LabelCVdescription(tempDatabase).Enabled = True
        Let Label31(tempDatabase).Enabled = True
        Let LabelCVdescription(tempDatabase).Enabled = True
        Let LocomotiveDecoderCVd(tempDatabase).Enabled = True
        If SpecificCVs(tempDatabase).Recordset.Fields("BitUsed") = True Then
            Let Option1(tempDatabase).Enabled = True
        Else
            Let Option1(tempDatabase).Enabled = False
        End If
    Else
        Let cv(tempDatabase) = "No"
        Let LabelCVdescription(tempDatabase).Enabled = False
        Let Label31(tempDatabase).Enabled = False
        Let LabelCVdescription(tempDatabase).Enabled = False
        Let LocomotiveDecoderCVd(tempDatabase).Enabled = False
        Let LabelCVdescription(tempDatabase).Caption = "CV not used by this decoder."
        Let Option1(tempDatabase).Enabled = False
    End If
Next tempDatabase

For tempDatabase = 112 To 128
    SpecificCVs(tempDatabase).Recordset.MoveFirst
    While SpecificCVs(tempDatabase).Recordset.Fields("RecordNumber") <> DecoderDatabase.Recordset.Fields("RecordNumber")
        SpecificCVs(tempDatabase).Recordset.MoveNext
    Wend
    DoEvents
    If SpecificCVs(tempDatabase).Recordset.Fields("Used") = True Then
        Let cv(tempDatabase) = "Yes"
        Let LabelCVdescription(tempDatabase).Enabled = True
        Let Label31(tempDatabase).Enabled = True
        Let LabelCVdescription(tempDatabase).Enabled = True
        Let LocomotiveDecoderCVd(tempDatabase).Enabled = True
        If SpecificCVs(tempDatabase).Recordset.Fields("BitUsed") = True Then
            Let Option1(tempDatabase).Enabled = True
        Else
            Let Option1(tempDatabase).Enabled = False
        End If
    Else
        Let cv(tempDatabase) = "No"
        Let LabelCVdescription(tempDatabase).Enabled = False
        Let Label31(tempDatabase).Enabled = False
        Let LabelCVdescription(tempDatabase).Enabled = False
        Let LocomotiveDecoderCVd(tempDatabase).Enabled = False
        Let LabelCVdescription(tempDatabase).Caption = "CV not used by this decoder."
        Let Option1(tempDatabase).Enabled = False
    End If
Next tempDatabase

Let MousePointer = Default

End Sub


Private Sub Command1_Click()

' =========================================================================================================================
' Loading the Programming Screen for 'Easy screen' Functions
'

' =========================================================================================================================
' Load Statement
'
' Loads a form or control into memory.
'
' Syntax is as the folowing:
'
'   Load object
'
' The object placeholder is the name of a Form object, MDIForm object, or control array element to load. You don 't need
' to use the Load statement with forms unless you want to load a form without displaying it.  Any reference to a form
' (except in a Set or If...TypeOf statement) automatically loads it if it's not already loaded.  For example, the Show
' method loads a form before displaying it.  Once the form is loaded, its properties and controls can be altered by the
' application, whether or not the form is actually visible.  Under some circumstances, you may want to load all your
' forms during initialization and display them later as they're needed. When more than one instance of a form is loaded,
' each form has its own set of properties and variables that can be changed independently from all other instances of the
' form. When Visual Basic loads a Form object, it sets form properties to their initial values and then performs the
' Load event procedure.  When an application starts, Visual Basic automatically loads and displays the application's
' startup form. If you load a Form with the MDIChild property set to True before loading the MDIForm, the MDIForm is
' automatically loaded before the child form.  Loading an MDI child form causes it to be visible immediately after
' the Form_Load event procedure ends because MDI child forms can't be hidden. The standard dialog boxes produced by
' Visual Basic functions such as MsgBox and InputBox do not need to be loaded, shown, or unloaded, but can simply be
' invoked directly.

    Load MainlineEasyScreenConsistFunctions

' =========================================================================================================================
' Show Method
'
'Displays an MDIForm or Form object.  Doesn't support named arguments. For Visual Basic in-process OLE server (DLL)
' considerations and restrictions that apply to this topic, see Make OLE DLL File Command.
'
' Syntax is as the following:
'
'   object.Show Style
'
' The Show method syntax has these parts where 'object' is optional. An object expression that evaluates to an object in
' the Applies To list. If object is omitted, the form associated with the active form module is assumed to be object. The
' 'style' is optional. Integer that determines if the form is modal or modeless.  If style is 0, the form is modeless;
' if style is 1, the form is modal. If the specified form isn't loaded when the Show method is invoked, Visual Basic
' automatically loads it. When Show displays a modeless form, subsequent code is executed as it's
' encountered.  When Show displays a modal form, no subsequent code is executed until the form is hidden or
' unloaded. When Show displays a modal form, no input (keyboard or mouse click) can occur except to objects on the
' modal form.  The program must hide or unload a modal form (usually in response to some user action) before input
' to another form can occur.  An MDIForm can't be modal. Although other forms in your application are disabled when
' a modal form is displayed, other applications aren't. The startup form of an application is automatically shown
' after its Load event is invoked.

    MainlineEasyScreenConsistFunctions.Show 0

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

Private Sub Command2_Click()

' =========================================================================================================================
' Loading the Programming Screen for 'Easy screen' Functions
'

' =========================================================================================================================
' Load Statement
'
' Loads a form or control into memory.
'
' Syntax is as the folowing:
'
'   Load object
'
' The object placeholder is the name of a Form object, MDIForm object, or control array element to load. You don 't need
' to use the Load statement with forms unless you want to load a form without displaying it.  Any reference to a form
' (except in a Set or If...TypeOf statement) automatically loads it if it's not already loaded.  For example, the Show
' method loads a form before displaying it.  Once the form is loaded, its properties and controls can be altered by the
' application, whether or not the form is actually visible.  Under some circumstances, you may want to load all your
' forms during initialization and display them later as they're needed. When more than one instance of a form is loaded,
' each form has its own set of properties and variables that can be changed independently from all other instances of the
' form. When Visual Basic loads a Form object, it sets form properties to their initial values and then performs the
' Load event procedure.  When an application starts, Visual Basic automatically loads and displays the application's
' startup form. If you load a Form with the MDIChild property set to True before loading the MDIForm, the MDIForm is
' automatically loaded before the child form.  Loading an MDI child form causes it to be visible immediately after
' the Form_Load event procedure ends because MDI child forms can't be hidden. The standard dialog boxes produced by
' Visual Basic functions such as MsgBox and InputBox do not need to be loaded, shown, or unloaded, but can simply be
' invoked directly.

    Load MainlineEasyScreenConfiguration

' =========================================================================================================================
' Show Method
'
'Displays an MDIForm or Form object.  Doesn't support named arguments. For Visual Basic in-process OLE server (DLL)
' considerations and restrictions that apply to this topic, see Make OLE DLL File Command.
'
' Syntax is as the following:
'
'   object.Show Style
'
' The Show method syntax has these parts where 'object' is optional. An object expression that evaluates to an object in
' the Applies To list. If object is omitted, the form associated with the active form module is assumed to be object. The
' 'style' is optional. Integer that determines if the form is modal or modeless.  If style is 0, the form is modeless;
' if style is 1, the form is modal. If the specified form isn't loaded when the Show method is invoked, Visual Basic
' automatically loads it. When Show displays a modeless form, subsequent code is executed as it's
' encountered.  When Show displays a modal form, no subsequent code is executed until the form is hidden or
' unloaded. When Show displays a modal form, no input (keyboard or mouse click) can occur except to objects on the
' modal form.  The program must hide or unload a modal form (usually in response to some user action) before input
' to another form can occur.  An MDIForm can't be modal. Although other forms in your application are disabled when
' a modal form is displayed, other applications aren't. The startup form of an application is automatically shown
' after its Load event is invoked.

    MainlineEasyScreenConfiguration.Show 0

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










Private Sub DecoderDatabase_Reposition()

Let DecoderDatabase.Caption = CStr(DecoderDatabase.Recordset.AbsolutePosition)
 
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

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.

Let LocomotiveDatabaseDecoder.DatabaseName = App.Path + "\Databases\LocomotiveDatabaseDecoders.mdb"
'Let SpecificCVsDatabase.DatabaseName = App.Path + "\Databases\SpecificDecoderCvs.mdb"

' Initializes the configuration variable to be read. This allows me to change the CV read acounting for the different
' decoder and manufactureres.

    Let cv(1) = "Yes"
    Let cv(2) = "Yes"
    Let cv(3) = "Yes"
    Let cv(4) = "Yes"
    Let cv(5) = "Yes"
    Let cv(6) = "Yes"
    Let cv(7) = "Yes"
    Let cv(8) = "Yes"
    Let cv(9) = "Yes"
    Let cv(10) = "Yes"
    Let cv(11) = "Yes"
    Let cv(12) = "Yes"
    Let cv(13) = "Yes"
    Let cv(17) = "Yes"
    Let cv(18) = "Yes"
    Let cv(19) = "Yes"
    Let cv(21) = "Yes"
    Let cv(22) = "Yes"
    Let cv(23) = "Yes"
    Let cv(24) = "Yes"
    Let cv(29) = "Yes"
    Let cv(33) = "Yes"
    Let cv(34) = "Yes"
    Let cv(35) = "Yes"
    Let cv(36) = "Yes"
    Let cv(37) = "Yes"
    Let cv(38) = "Yes"
    Let cv(39) = "Yes"
    Let cv(40) = "Yes"
    Let cv(41) = "Yes"
    Let cv(42) = "Yes"
    Let cv(65) = "Yes"
    Let cv(66) = "Yes"
    Let cv(95) = "Yes"
    Let cv(105) = "Yes"
    Let cv(106) = "Yes"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

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
        Else
            Let FourByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 64
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 32
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 16
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 8
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 4
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 2
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let FourByteB.Text = FourByteB.Text + "1"
            Let temp = temp - 1
        Else
            Let FourByteB.Text = FourByteB.Text + "0"
        End If

End Sub































Private Sub LocomotiveDatabaseDecoder_Reposition()

Let LocomotiveDatabaseDecoder.Caption = "Loco " & CStr(LocomotiveDatabaseDecoder.Recordset.AbsolutePosition)

End Sub

Private Sub LocomotiveDecoderCVd_Change(Index As Integer)

Let ConfigurationVariableLow = 0
Let ConfigurationVariableHigh = 255

If Index = 1 Then
    Let ConfigurationVariableLow = 1
    Let ConfigurationVariableHigh = 127
End If

If Val(LocomotiveDecoderCVb(Index).Text) < ConfigurationVariableLow Then
    
    If DisplayErrorMessage.Value = vbChecked Then
        
        Let message = "The value entered for the Configuration Variable"
        Let message = message + Str$(Format(Index, "000"))
        Let message = message + " is not valid. You" + Chr$(13)
        Let message = message + " enter a number less then"
        Let message = message + Str$(ConfigurationVariableLow)
        Let message = message + ". Please use a value between"
        Let message = message + Str$(ConfigurationVariableLow)
        Let message = message + " and"
        Let message = message + Str$(ConfigurationVariableHigh)
        Let message = message + " only. " + Chr$(13)
        Let message = message + Chr$(13)
        Let message = message + "Configuration Variable"
        Let message = message + Str$(Format(Index, "000"))
        Let message = message + " will be reset to"
        Let message = message + Str$(ConfigurationVariableLow)
        Let message = message + "." + Chr$(13)
        Let message = message + Chr$(13)
        Let message = message + "Do you wish to continue to receive these warning messages?"
        Let ReturnMessage = MsgBox(message, vbYesNo + vbInformation, "CV - Entry Error")
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
        Let message = "The value entered for the Configuration Variable"
        Let message = message + Str$(Format(Index, "000"))
        Let message = message + " is not valid. You enter a number" + Chr$(13)
        Let message = message + "greater then"
        Let message = message + Str$(ConfigurationVariableHigh)
        Let message = message + ". Please use a value between"
        Let message = message + Str$(ConfigurationVariableLow)
        Let message = message + " and"
        Let message = message + Str$(ConfigurationVariableHigh)
        Let message = message + " only. " + Chr$(13)
        Let message = message + Chr$(13)
        Let message = message + "Configuration Variable"
        Let message = message + Str$(Format(Index, "000"))
        Let message = message + " will be reset to"
        Let message = message + Str$(ConfigurationVariableHigh)
        Let message = message + "." + Chr$(13)
        Let message = message + Chr$(13)
        Let message = message + "Do you wish to continue to receive these warning messages?"
            
        Let ReturnMessage = MsgBox(message, vbYesNo + vbInformation, "CV - Entry Error")
            
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




' ===========================================================================================================================================
' Special Routines
' ===========================================================================================================================================
'
' For Configuration Variables - Starting Voltage -Mid Voltage - High Voltage - Kick Voltage
'
If Index = 2 Or _
   Index = 5 Or _
   Index = 6 Or _
   Index = 65 Then

    Let s_voltage = (Int(Val(LocomotiveDecoderCVd(2).Text) * 12 / 255 * 100) / 100)
    Let m_voltage = (Int(Val(LocomotiveDecoderCVd(6).Text) * 12 / 255 * 100) / 100)
    Let h_voltage = (Int(Val(LocomotiveDecoderCVd(5).Text) * 12 / 255 * 100) / 100)
    Let k_voltage = (Int(Val(LocomotiveDecoderCVd(65).Text) * 12 / 255 * 100) / 100)

    Let Label44.Caption = "Voltage Nots:" + Chr$(13)
    Let Label44.Caption = Label44.Caption + "CV2, on a twelve volt system sets starting voltage to be approximately"
    Let Label44.Caption = Label44.Caption + Str$(s_voltage) + " volt(s)." + Chr$(13)
    Let Label44.Caption = Label44.Caption + "CV6, on a twelve volt system sets midrange voltage to be approximately"
    Let Label44.Caption = Label44.Caption + Str$(m_voltage) + " volt(s)." + Chr$(13)
    Let Label44.Caption = Label44.Caption + "CV5, on a twelve volt system sets high voltage to be approximately"
    Let Label44.Caption = Label44.Caption + Str$(h_voltage) + " volt(s)." + Chr$(13)
    Let Label44.Caption = Label44.Caption + "CV65, on a twelve volt system sets kick voltage to be approximately"
    Let Label44.Caption = Label44.Caption + Str$(k_voltage) + " volt(s)." + Chr$(13)
    
End If

' ===========================================================================================================================================
' Special Routines
' ===========================================================================================================================================
'
' For Configuration Variables - Manufacturer Identification and Decoder Identification

If Index = 7 Or _
   Index = 8 Then
   
    Let Label43.Caption = "Identification Notes:" + Chr$(13)
    
    If Val(LocomotiveDecoderCVd(8).Text) = 137 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is A-Train Electronics" + Chr$(13)
        Let TextBoxManufacturerName.Text = "A-Train Electronics"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 173 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Arnold - Rivarossi" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Arnold - Rivarossi"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 127 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Atlas Model Railroad Products" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Atlas Model Railroad Products"
        If Val(LocomotiveDecoderCVd(7).Text) = 45 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is AH120FXLP"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 135 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is CVP Products" + Chr$(13)
        Let TextBoxManufacturerName.Text = "CVP Products"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 129 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Digitrax" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Digitrax"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 151 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Electronic Solutions" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Electronic Solutions"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 155 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Bebr, Fleischmann GmdH & Co." + Chr$(13)
        Let TextBoxManufacturerName.Text = "Bebr, Fleischmann GmdH & Co."
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 133 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Intelligent Command Control" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Intelligent Command Control"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 157 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Kuehn Ing." + Chr$(13)
        Let TextBoxManufacturerName.Text = "Kuehn Ing."
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 99 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Lenz" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Lenz"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 159 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is LGB (Ernst Paul Lehmann Patentwerk)" + Chr$(13)
        Let TextBoxManufacturerName.Text = "LGB (Ernst Paul Lehmann Patentwerk)"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 143 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Model Rectifier Corporation" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Model Rectifier Corporation"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 11 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is North Coast Engineering" + Chr$(13)
        Let TextBoxManufacturerName.Text = "North Coast Engineering"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 14 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is PSI - Dynatrol" + Chr$(13)
        Let TextBoxManufacturerName.Text = "PSI - Dynatrol"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 13 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Public Domain and Do-It-Yourself Decoders" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Public Domain"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 15 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Ranfixx Technologies (Wangrow)" + Chr$(13)
        Let TextBoxManufacturerName.Text = "anfixx Technologies (Wangrow)"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 139 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is RealRail Effects" + Chr$(13)
        Let TextBoxManufacturerName.Text = "RealRail Effects"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 149 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Rock Junction Controls" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Rock Junction Controls"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 161 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Roco Modelspielwarren" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Roco Modelspielwarren"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
    If Val(LocomotiveDecoderCVd(8).Text) = 141 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Throttle-Up (Soundtraxx)" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Throttle-Up (Soundtraxx)"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 153 Then
        Let Label43.Caption = Label53.Caption + "CV8, Manufacturer is Train Control Systems" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Train Control Systems"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 131 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Trix ModelHeisenbaln" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Trix ModelHeisenbaln"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 85 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Uhelmbrock" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Uhelmbrock"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 47 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Umelec Ing. Buero" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Umelec Ing. Buero"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 163 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is WP RailShops" + Chr$(13)
        Let TextBoxManufacturerName.Text = "WP RailShops"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 145 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is Zimo Elektronik" + Chr$(13)
        Let TextBoxManufacturerName.Text = "Zimo Elektronik"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
        
    If Val(LocomotiveDecoderCVd(8).Text) = 132 Then
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is " + Chr$(13)
        Let TextBoxManufacturerName.Text = "ZTC"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
End If


' ===========================================================================================================================================
' Special Routines
' ===========================================================================================================================================
'
' For Configuration Variables - Starting Voltage -Mid Voltage - High Voltage - Kick Voltage
'

If Index = 17 Or _
   Index = 18 Then

    Let LocomotiveDecoderCVc(18).Text = (Val(LocomotiveDecoderCVd(17).Text) - 128 - 64) * 256
    Let LocomotiveDecoderCVc(18).Text = Val(LocomotiveDecoderCVc(18).Text) + Val(LocomotiveDecoderCVd(18).Text)
        
End If


' ===========================================================================================================================================
' Special Routines
' ===========================================================================================================================================
'
' For Configuration Variables - Starting Voltage -Mid Voltage - High Voltage - Kick Voltage
'

If Index = 3 Or _
   Index = 23 Then

    If Val(LocomotiveDecoderCVd(23).Text) >= 128 Then
        Let LocomotiveDecoderCVc(23).Text = Val(LocomotiveDecoderCVd(3).Text) - Val(LocomotiveDecoderCVd(23).Text) + 128
    End If
    If Val(LocomotiveDecoderCVd(23).Text) <= 127 Then
        Let LocomotiveDecoderCVc(23).Text = Val(LocomotiveDecoderCVd(3).Text) + Val(LocomotiveDecoderCVd(23).Text)
    End If
   
Let Label77.Caption = "Speed Notes:" + Chr$(13) + _
    "CV3, using 128 speed steps, sets acceleration to" + Str$((Val(LocomotiveDecoderCVd(3).Text) * 0.896 / 128)) + " seconds for each speed step." + Chr$(13) + _
    "CV4, using 128 speed steps, sets acceleration to" + Str$((Val(LocomotiveDecoderCVd(4).Text) * 0.896 / 128)) + " seconds for each speed step." + Chr$(13) + _
    "CV23, using 128 speed steps, adjusts acceleration by" + Str$((Val(LocomotiveDecoderCVd(23).Text) * 0.896 / 128)) + " to" + Str$((Val(LocomotiveDecoderCVc(23).Text) * 0.896 / 128)) + "seconds for each speed step." + Chr$(13) + _
    "CV24, using 128 speed steps, adjusts acceleration by" + Str$((Val(LocomotiveDecoderCVd(24).Text) * 0.896 / 128)) + " to" + Str$((Val(LocomotiveDecoderCVc(24).Text) * 0.896 / 128)) + "seconds for each speed step." + Chr$(13)

End If

If Index = 4 Or _
   Index = 24 Then

    If Val(LocomotiveDecoderCVd(24).Text) >= 128 Then
        Let LocomotiveDecoderCVc(24).Text = Val(LocomotiveDecoderCVd(4).Text) - Val(LocomotiveDecoderCVd(24).Text) + 128
    End If
    If Val(LocomotiveDecoderCVd(24).Text) <= 127 Then
        Let LocomotiveDecoderCVc(24).Text = Val(LocomotiveDecoderCVd(4).Text) + Val(LocomotiveDecoderCVd(24).Text)
    End If
   
Let Label77.Caption = "Speed Notes:" + Chr$(13) + _
    "CV3, using 128 speed steps, sets acceleration to" + Str$((Val(LocomotiveDecoderCVd(3).Text) * 0.896 / 128)) + " seconds for each speed step." + Chr$(13) + _
    "CV4, using 128 speed steps, sets acceleration to" + Str$((Val(LocomotiveDecoderCVd(4).Text) * 0.896 / 128)) + " seconds for each speed step." + Chr$(13) + _
    "CV23, using 128 speed steps, adjusts acceleration by" + Str$((Val(LocomotiveDecoderCVd(23).Text) * 0.896 / 128)) + " to" + Str$((Val(LocomotiveDecoderCVc(23).Text) * 0.896 / 128)) + "seconds for each speed step." + Chr$(13) + _
    "CV24, using 128 speed steps, adjusts acceleration by" + Str$((Val(LocomotiveDecoderCVd(24).Text) * 0.896 / 128)) + " to" + Str$((Val(LocomotiveDecoderCVc(24).Text) * 0.896 / 128)) + "seconds for each speed step." + Chr$(13)
    
End If

If Index = 29 Then
    If (Val(LocomotiveDecoderCVd(29)) And 16) Then
        ButtonSpeedTable.Enabled = True
    Else
        ButtonSpeedTable.Enabled = False
    End If
End If

End Sub

Private Sub OneByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let OneByteH.Text = Hex(Val(OneByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(OneByteH.Text) = 1 Then
        Let OneByteH.Text = "0" + OneByteH.Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(OneByteD.Text)
        
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
        
    Let SevenByteH.Text = Hex(Val(SevenByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SevenByteH.Text) = 1 Then
        Let SevenByteH.Text = "0" + SevenByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(SevenByteD.Text)
        
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





Private Sub SpecificCVs_Reposition(Index As Integer)

'Let SpecificCVs(Index).Caption = CStr(SpecificCVs(Index).Recordset.AbsolutePosition)

End Sub

Private Sub TextBoxManufacturerName_Change()

Let MousePointer = vbArrowHourglass

DecoderDatabase.Recordset.MoveFirst

ComboBoxDecoderName.Clear

While Not DecoderDatabase.Recordset.EOF

    If DecoderDatabase.Recordset.Fields("Manufacturer") = TextBoxManufacturerName.Text Then
        Let temp = DecoderDatabase.Recordset.Fields("decoder")
        ComboBoxDecoderName.AddItem temp
    End If
    DecoderDatabase.Recordset.MoveNext
Wend

Let MousePointer = Default

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

Let MousePointer = vbArrowHourglass

Let ButtonClose.Enabled = False

Let Timer1.Interval = 0

Let DisplayErrorMessage.Value = vbChecked

    LocomotiveDatabaseDecoder.Recordset.MoveLast
    
    ' Which screen called the decoder screen? was is mainline diesel? was is mainline decoder?
    
    If Val(MainlineConsist.LocomotiveNumberConsist.Text) <> 0 Then
        Let temp = Val(MainlineConsist.LocomotiveNumberConsist.Text)
    End If
    If Val(MainlineDiesel.LocomotiveNumberGO.Text) <> 0 Then
        Let temp = Val(MainlineDiesel.LocomotiveNumberGO.Text)
    End If
    
    LocomotiveDatabaseDecoder.Recordset.AbsolutePosition = temp
    
Let MousePointer = vbDefault

Let ButtonClose.Enabled = True

End Sub


Private Sub TwoByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let TwoByteH.Text = Hex(Val(TwoByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(TwoByteH.Text) = 1 Then
        Let TwoByteH.Text = "0" + TwoByteH.Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(TwoByteD.Text)
        
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



Private Sub SetErrorByte()
   
' ===========================================================================================================================================
' Error Byte Calculations
' ===========================================================================================================================================
'
' The error byte is kepped in the seventh btye (seventh text box on the screen). This is calculated by going through the six bytes and
' counting the number of 'ones', XOR, to determine the error btye.

            Let SevenByteD.Text = "0"

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

                If OneByteD.Text <> "" Then
                    If OneByteD.Text <> "   " Then
                        If Mid$(OneByteB.Text, X, 1) = "1" Then
                            Let temp = temp + 1
                        End If
                    End If
                End If
                If TwoByteD.Text <> "" Then
                    If TwoByteD.Text <> "   " Then
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

                If Int(temp / 2) <> (temp / 2) Then SevenByteD.Text = Val(SevenByteD.Text) + bitvalue

' My Notes:
'
' This is where we need to return to the top of the 'for next' loop. Again the loop is preformed eight times, once for
' each bit inthe varible.
            
            Next X

End Sub

Private Sub SendCommand()
   
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
            
            If OneByteD.Text <> "" Then
                If OneByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + OneByteH.Text
                End If
            End If
            If TwoByteD.Text <> "" Then
                If TwoByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + TwoByteH.Text
                End If
            End If
            If ThreeByteD.Text <> "" Then
                If ThreeByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + ThreeByteH.Text
                End If
            End If
            If FourByteD.Text <> "" Then
                If FourByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + FourByteH.Text
                End If
            End If
            If FiveByteD.Text <> "" Then
                If FiveByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + FiveByteH.Text
                End If
            End If
            If SixByteD.Text <> "" Then
                If SixByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + SixByteH.Text
                End If
            End If
            If SevenByteD.Text <> "" Then
                If SevenByteD.Text <> "   " Then
                    Let CommandControl = CommandControl + " " + SevenByteH.Text
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
    
            Let LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

            While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
                Let temp = DoEvents
            Wend
            Let LocomotiveCommunicationStatus.Caption = "Status: Clear"
    

' =========================================================================================================================
' Dequeue packet only if packet was queued with "q" command
' =========================================================================================================================

If MainScreen!checkboxdequeuepacket.Value = vbChecked Then

                Let CommandControl = "d"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
                If OneByteD.Text <> "" Then
                    If OneByteD.Text <> "   " Then
                        Let CommandControl = CommandControl + " " + OneByteH.Text
                    End If
                End If
                If TwoByteD.Text <> "" Then
                    If TwoByteD.Text <> "   " Then
                        Let CommandControl = CommandControl + " " + TwoByteH.Text
                    End If
                End If
                'If ThreeByteD.Text <> "" Then
                '    If ThreeByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + ThreeByteH.Text
                '    End If
                'End If
                'If FourByteD.Text <> "" Then
                '    If FourByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + FourByteH.Text
                '    End If
                'End If
                'If FiveByteD.Text <> "" Then
                '    If FiveByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + FiveByteH.Text
                '    End If
                'End If
                'If SixByteD.Text <> "" Then
                '    If SixByteD.Text <> "   " Then
                '        Let CommandControl = CommandControl + " " + SixByteH.Text
                '    End If
                'End If
                If SevenByteD.Text <> "" Then
                    If SevenByteD.Text <> "   " Then
                        Let CommandControl = CommandControl + " " + SevenByteH.Text
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
    
                Let LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

                While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
                    Let temp = DoEvents
                Wend
                Let LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

End If
       
End Sub
