VERSION 4.00
Begin VB.Form MainlineDecoder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Decoders"
   ClientHeight    =   10770
   ClientLeft      =   2715
   ClientTop       =   2355
   ClientWidth     =   8745
   Height          =   11175
   Icon            =   "MainlineDecoder.frx":0000
   Left            =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   8745
   Top             =   2010
   Width           =   8865
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   6120
      TabIndex        =   519
      Top             =   10440
      Width           =   1215
   End
   Begin VB.Data DecoderDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DecoderList"
      Top             =   480
      Width           =   1140
   End
   Begin VB.Data LocomotiveDatabaseDecoder 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDecoders"
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "&List"
      Height          =   255
      Left            =   120
      TabIndex        =   255
      Top             =   10440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "MainlineDecoder.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox LocomotiveNumber 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseDecoder"
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton ButtonExport 
      Caption         =   "&Export"
      Height          =   255
      Left            =   1440
      TabIndex        =   256
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   7440
      TabIndex        =   0
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   258
      Top             =   10440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   7
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
      TabPicture(0)   =   "MainlineDecoder.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelCVDescription(106)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelCVDescription(105)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label43"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelCVDescription(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LabelCVDescription(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LabelCVDescription(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LabelCVDescription(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LabelCVDescription(18)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "LabelCVDescription(19)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line9"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "LabelCVPrefix(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LabelCVPrefix(17)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "LabelCVPrefix(18)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "LabelCVPrefix(19)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LabelCVPrefix(7)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "LabelCVPrefix(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "LabelCVPrefix(105)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "LabelCVPrefix(106)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CheckBoxCVChanged(106)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CheckBoxCVChanged(105)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "CheckBoxCVChanged(8)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "CheckBoxCVChanged(7)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "CheckBoxCVChanged(19)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "CheckBoxCVChanged(18)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CheckBoxCVChanged(17)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "CheckBoxCVChanged(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "LocomotiveDecoderCVb(106)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "LocomotiveDecoderCVb(105)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "LocomotiveDecoderCVh(106)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "LocomotiveDecoderCVh(105)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "LocomotiveDecoderCVd(106)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "LocomotiveDecoderCVd(105)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "LocomotiveDecoderCVd(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "LocomotiveDecoderCVh(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "LocomotiveDecoderCVb(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "LocomotiveDecoderCVd(7)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "LocomotiveDecoderCVh(7)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "LocomotiveDecoderCVb(7)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "LocomotiveDecoderCVd(8)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "LocomotiveDecoderCVh(8)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "LocomotiveDecoderCVb(8)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "LocomotiveDecoderCVd(17)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "LocomotiveDecoderCVh(17)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "LocomotiveDecoderCVb(17)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "LocomotiveDecoderCVd(18)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "LocomotiveDecoderCVh(18)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "LocomotiveDecoderCVb(18)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "LocomotiveDecoderCVc(18)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "LocomotiveDecoderCVd(19)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "LocomotiveDecoderCVh(19)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "LocomotiveDecoderCVb(19)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "SpecificCVs(1)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "SpecificCVs(17)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "SpecificCVs(18)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "SpecificCVs(19)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "SpecificCVs(7)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "SpecificCVs(8)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "SpecificCVs(105)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "SpecificCVs(106)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).ControlCount=   67
      TabCaption(1)   =   "Voltage Configuration"
      TabPicture(1)   =   "MainlineDecoder.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "LabelCVDescription(2)"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(5)=   "LabelCVDescription(6)"
      Tab(1).Control(6)=   "LabelCVDescription(5)"
      Tab(1).Control(7)=   "Line6"
      Tab(1).Control(8)=   "LabelCVDescription(65)"
      Tab(1).Control(9)=   "Line7"
      Tab(1).Control(10)=   "Label44"
      Tab(1).Control(11)=   "LabelCVPrefix(2)"
      Tab(1).Control(12)=   "LabelCVPrefix(6)"
      Tab(1).Control(13)=   "LabelCVPrefix(5)"
      Tab(1).Control(14)=   "LabelCVPrefix(65)"
      Tab(1).Control(15)=   "CheckBoxCVChanged(65)"
      Tab(1).Control(16)=   "CheckBoxCVChanged(5)"
      Tab(1).Control(17)=   "CheckBoxCVChanged(6)"
      Tab(1).Control(18)=   "CheckBoxCVChanged(2)"
      Tab(1).Control(19)=   "LocomotiveDecoderCVb(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "LocomotiveDecoderCVh(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "LocomotiveDecoderCVd(2)"
      Tab(1).Control(22)=   "LocomotiveDecoderCVd(6)"
      Tab(1).Control(23)=   "LocomotiveDecoderCVh(6)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "LocomotiveDecoderCVb(6)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "LocomotiveDecoderCVd(5)"
      Tab(1).Control(26)=   "LocomotiveDecoderCVh(5)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "LocomotiveDecoderCVb(5)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "LocomotiveDecoderCVb(65)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "LocomotiveDecoderCVh(65)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "LocomotiveDecoderCVd(65)"
      Tab(1).Control(31)=   "SpecificCVs(2)"
      Tab(1).Control(32)=   "SpecificCVs(6)"
      Tab(1).Control(33)=   "SpecificCVs(5)"
      Tab(1).Control(34)=   "SpecificCVs(65)"
      Tab(1).ControlCount=   35
      TabCaption(2)   =   "Speed Regulation"
      TabPicture(2)   =   "MainlineDecoder.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label77"
      Tab(2).Control(1)=   "Label18"
      Tab(2).Control(2)=   "Label19"
      Tab(2).Control(3)=   "Label20"
      Tab(2).Control(4)=   "Label21"
      Tab(2).Control(5)=   "LabelCVDescription(4)"
      Tab(2).Control(6)=   "Line3"
      Tab(2).Control(7)=   "LabelCVDescription(23)"
      Tab(2).Control(8)=   "LabelCVDescription(24)"
      Tab(2).Control(9)=   "Line4"
      Tab(2).Control(10)=   "LabelCVDescription(66)"
      Tab(2).Control(11)=   "LabelCVDescription(95)"
      Tab(2).Control(12)=   "Line10"
      Tab(2).Control(13)=   "LabelCVPrefix(3)"
      Tab(2).Control(14)=   "LabelCVPrefix(23)"
      Tab(2).Control(15)=   "LabelCVPrefix(4)"
      Tab(2).Control(16)=   "LabelCVPrefix(24)"
      Tab(2).Control(17)=   "LabelCVPrefix(66)"
      Tab(2).Control(18)=   "LabelCVPrefix(95)"
      Tab(2).Control(19)=   "LabelCVDescription(3)"
      Tab(2).Control(20)=   "Line15"
      Tab(2).Control(21)=   "Label5"
      Tab(2).Control(22)=   "CheckBoxCVChanged(95)"
      Tab(2).Control(23)=   "CheckBoxCVChanged(66)"
      Tab(2).Control(24)=   "CheckBoxCVChanged(24)"
      Tab(2).Control(25)=   "CheckBoxCVChanged(4)"
      Tab(2).Control(26)=   "CheckBoxCVChanged(23)"
      Tab(2).Control(27)=   "CheckBoxCVChanged(3)"
      Tab(2).Control(28)=   "LocomotiveDecoderCVb(3)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "LocomotiveDecoderCVh(3)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "LocomotiveDecoderCVd(3)"
      Tab(2).Control(31)=   "LocomotiveDecoderCVb(4)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "LocomotiveDecoderCVh(4)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "LocomotiveDecoderCVd(4)"
      Tab(2).Control(34)=   "LocomotiveDecoderCVb(23)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "LocomotiveDecoderCVh(23)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "LocomotiveDecoderCVd(23)"
      Tab(2).Control(37)=   "LocomotiveDecoderCVb(24)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "LocomotiveDecoderCVh(24)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "LocomotiveDecoderCVd(24)"
      Tab(2).Control(40)=   "LocomotiveDecoderCVc(23)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "LocomotiveDecoderCVc(24)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "ButtonSpeedTable"
      Tab(2).Control(43)=   "LocomotiveDecoderCVb(66)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "LocomotiveDecoderCVh(66)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "LocomotiveDecoderCVd(66)"
      Tab(2).Control(46)=   "LocomotiveDecoderCVb(95)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "LocomotiveDecoderCVh(95)"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "LocomotiveDecoderCVd(95)"
      Tab(2).Control(49)=   "SpecificCVs(3)"
      Tab(2).Control(50)=   "SpecificCVs(23)"
      Tab(2).Control(51)=   "SpecificCVs(4)"
      Tab(2).Control(52)=   "SpecificCVs(24)"
      Tab(2).Control(53)=   "SpecificCVs(66)"
      Tab(2).Control(54)=   "SpecificCVs(95)"
      Tab(2).ControlCount=   55
      TabCaption(3)   =   "General Configuration"
      TabPicture(3)   =   "MainlineDecoder.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label29"
      Tab(3).Control(1)=   "Label26"
      Tab(3).Control(2)=   "Label27"
      Tab(3).Control(3)=   "LabelCVDescription(29)"
      Tab(3).Control(4)=   "Line5"
      Tab(3).Control(5)=   "Label37"
      Tab(3).Control(6)=   "Label32"
      Tab(3).Control(7)=   "LabelCVPrefix(29)"
      Tab(3).Control(8)=   "Command2"
      Tab(3).Control(9)=   "CheckBoxCVChanged(29)"
      Tab(3).Control(10)=   "LocomotiveDecoderCVb(29)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "LocomotiveDecoderCVh(29)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "LocomotiveDecoderCVd(29)"
      Tab(3).Control(13)=   "SpecificCVs(29)"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Other Settings"
      TabPicture(4)   =   "MainlineDecoder.frx":08F4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label76"
      Tab(4).Control(1)=   "Label74"
      Tab(4).Control(2)=   "Label72"
      Tab(4).Control(3)=   "LabelCVDescription(13)"
      Tab(4).Control(4)=   "LabelCVDescription(12)"
      Tab(4).Control(5)=   "LabelCVDescription(11)"
      Tab(4).Control(6)=   "LabelCVDescription(10)"
      Tab(4).Control(7)=   "LabelCVDescription(9)"
      Tab(4).Control(8)=   "Label33"
      Tab(4).Control(9)=   "LabelCVPrefix(9)"
      Tab(4).Control(10)=   "LabelCVPrefix(10)"
      Tab(4).Control(11)=   "LabelCVPrefix(11)"
      Tab(4).Control(12)=   "LabelCVPrefix(12)"
      Tab(4).Control(13)=   "LabelCVPrefix(13)"
      Tab(4).Control(14)=   "LabelCVPrefix(15)"
      Tab(4).Control(15)=   "LabelCVDescription(15)"
      Tab(4).Control(16)=   "LabelCVDescription(16)"
      Tab(4).Control(17)=   "LabelCVPrefix(16)"
      Tab(4).Control(18)=   "Line16"
      Tab(4).Control(19)=   "LabelCVPrefix(14)"
      Tab(4).Control(20)=   "LabelCVDescription(14)"
      Tab(4).Control(21)=   "CheckBoxCVChanged(13)"
      Tab(4).Control(22)=   "CheckBoxCVChanged(12)"
      Tab(4).Control(23)=   "CheckBoxCVChanged(11)"
      Tab(4).Control(24)=   "CheckBoxCVChanged(10)"
      Tab(4).Control(25)=   "CheckBoxCVChanged(9)"
      Tab(4).Control(26)=   "LocomotiveDecoderCVb(13)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "LocomotiveDecoderCVh(13)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "LocomotiveDecoderCVd(13)"
      Tab(4).Control(29)=   "LocomotiveDecoderCVb(12)"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "LocomotiveDecoderCVh(12)"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "LocomotiveDecoderCVd(12)"
      Tab(4).Control(32)=   "LocomotiveDecoderCVb(11)"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "LocomotiveDecoderCVh(11)"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "LocomotiveDecoderCVd(11)"
      Tab(4).Control(35)=   "LocomotiveDecoderCVb(10)"
      Tab(4).Control(35).Enabled=   0   'False
      Tab(4).Control(36)=   "LocomotiveDecoderCVh(10)"
      Tab(4).Control(36).Enabled=   0   'False
      Tab(4).Control(37)=   "LocomotiveDecoderCVd(10)"
      Tab(4).Control(38)=   "LocomotiveDecoderCVb(9)"
      Tab(4).Control(38).Enabled=   0   'False
      Tab(4).Control(39)=   "LocomotiveDecoderCVh(9)"
      Tab(4).Control(39).Enabled=   0   'False
      Tab(4).Control(40)=   "LocomotiveDecoderCVd(9)"
      Tab(4).Control(41)=   "SpecificCVs(9)"
      Tab(4).Control(42)=   "SpecificCVs(10)"
      Tab(4).Control(43)=   "SpecificCVs(11)"
      Tab(4).Control(44)=   "SpecificCVs(12)"
      Tab(4).Control(45)=   "SpecificCVs(13)"
      Tab(4).Control(46)=   "SpecificCVs(15)"
      Tab(4).Control(47)=   "LocomotiveDecoderCVd(15)"
      Tab(4).Control(48)=   "LocomotiveDecoderCVh(15)"
      Tab(4).Control(48).Enabled=   0   'False
      Tab(4).Control(49)=   "LocomotiveDecoderCVb(15)"
      Tab(4).Control(49).Enabled=   0   'False
      Tab(4).Control(50)=   "LocomotiveDecoderCVb(16)"
      Tab(4).Control(50).Enabled=   0   'False
      Tab(4).Control(51)=   "LocomotiveDecoderCVh(16)"
      Tab(4).Control(51).Enabled=   0   'False
      Tab(4).Control(52)=   "LocomotiveDecoderCVd(16)"
      Tab(4).Control(53)=   "SpecificCVs(16)"
      Tab(4).Control(54)=   "SpecificCVs(14)"
      Tab(4).Control(55)=   "LocomotiveDecoderCVd(14)"
      Tab(4).Control(56)=   "LocomotiveDecoderCVh(14)"
      Tab(4).Control(56).Enabled=   0   'False
      Tab(4).Control(57)=   "LocomotiveDecoderCVb(14)"
      Tab(4).Control(57).Enabled=   0   'False
      Tab(4).ControlCount=   58
      TabCaption(5)   =   "Functions"
      TabPicture(5)   =   "MainlineDecoder.frx":0910
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label7"
      Tab(5).Control(1)=   "Label1"
      Tab(5).Control(2)=   "Line8"
      Tab(5).Control(3)=   "Label75"
      Tab(5).Control(4)=   "Label73"
      Tab(5).Control(5)=   "Label71"
      Tab(5).Control(6)=   "LabelCVDescription(21)"
      Tab(5).Control(7)=   "LabelCVDescription(22)"
      Tab(5).Control(8)=   "LabelCVDescription(42)"
      Tab(5).Control(9)=   "Line11"
      Tab(5).Control(10)=   "LabelCVDescription(41)"
      Tab(5).Control(11)=   "LabelCVDescription(40)"
      Tab(5).Control(12)=   "LabelCVDescription(39)"
      Tab(5).Control(13)=   "LabelCVDescription(38)"
      Tab(5).Control(14)=   "LabelCVDescription(37)"
      Tab(5).Control(15)=   "LabelCVDescription(36)"
      Tab(5).Control(16)=   "LabelCVDescription(35)"
      Tab(5).Control(17)=   "LabelCVDescription(34)"
      Tab(5).Control(18)=   "LabelCVDescription(33)"
      Tab(5).Control(19)=   "Label34"
      Tab(5).Control(20)=   "LabelCVPrefix(21)"
      Tab(5).Control(21)=   "LabelCVPrefix(22)"
      Tab(5).Control(22)=   "LabelCVPrefix(33)"
      Tab(5).Control(23)=   "LabelCVPrefix(34)"
      Tab(5).Control(24)=   "LabelCVPrefix(35)"
      Tab(5).Control(25)=   "LabelCVPrefix(36)"
      Tab(5).Control(26)=   "LabelCVPrefix(37)"
      Tab(5).Control(27)=   "LabelCVPrefix(38)"
      Tab(5).Control(28)=   "LabelCVPrefix(39)"
      Tab(5).Control(29)=   "LabelCVPrefix(40)"
      Tab(5).Control(30)=   "LabelCVPrefix(41)"
      Tab(5).Control(31)=   "LabelCVPrefix(42)"
      Tab(5).Control(32)=   "Command1"
      Tab(5).Control(33)=   "ButtonEasyScreenFunctions"
      Tab(5).Control(34)=   "CheckBoxCVChanged(42)"
      Tab(5).Control(35)=   "CheckBoxCVChanged(41)"
      Tab(5).Control(36)=   "CheckBoxCVChanged(40)"
      Tab(5).Control(37)=   "CheckBoxCVChanged(39)"
      Tab(5).Control(38)=   "CheckBoxCVChanged(38)"
      Tab(5).Control(39)=   "CheckBoxCVChanged(37)"
      Tab(5).Control(40)=   "CheckBoxCVChanged(36)"
      Tab(5).Control(41)=   "CheckBoxCVChanged(35)"
      Tab(5).Control(42)=   "CheckBoxCVChanged(34)"
      Tab(5).Control(43)=   "CheckBoxCVChanged(33)"
      Tab(5).Control(44)=   "CheckBoxCVChanged(22)"
      Tab(5).Control(45)=   "CheckBoxCVChanged(21)"
      Tab(5).Control(46)=   "LocomotiveDecoderCVb(42)"
      Tab(5).Control(46).Enabled=   0   'False
      Tab(5).Control(47)=   "LocomotiveDecoderCVb(41)"
      Tab(5).Control(47).Enabled=   0   'False
      Tab(5).Control(48)=   "LocomotiveDecoderCVb(40)"
      Tab(5).Control(48).Enabled=   0   'False
      Tab(5).Control(49)=   "LocomotiveDecoderCVb(39)"
      Tab(5).Control(49).Enabled=   0   'False
      Tab(5).Control(50)=   "LocomotiveDecoderCVb(38)"
      Tab(5).Control(50).Enabled=   0   'False
      Tab(5).Control(51)=   "LocomotiveDecoderCVb(37)"
      Tab(5).Control(51).Enabled=   0   'False
      Tab(5).Control(52)=   "LocomotiveDecoderCVb(36)"
      Tab(5).Control(52).Enabled=   0   'False
      Tab(5).Control(53)=   "LocomotiveDecoderCVb(35)"
      Tab(5).Control(53).Enabled=   0   'False
      Tab(5).Control(54)=   "LocomotiveDecoderCVb(34)"
      Tab(5).Control(54).Enabled=   0   'False
      Tab(5).Control(55)=   "LocomotiveDecoderCVb(33)"
      Tab(5).Control(55).Enabled=   0   'False
      Tab(5).Control(56)=   "LocomotiveDecoderCVb(22)"
      Tab(5).Control(56).Enabled=   0   'False
      Tab(5).Control(57)=   "LocomotiveDecoderCVb(21)"
      Tab(5).Control(57).Enabled=   0   'False
      Tab(5).Control(58)=   "LocomotiveDecoderCVh(42)"
      Tab(5).Control(58).Enabled=   0   'False
      Tab(5).Control(59)=   "LocomotiveDecoderCVh(41)"
      Tab(5).Control(59).Enabled=   0   'False
      Tab(5).Control(60)=   "LocomotiveDecoderCVh(40)"
      Tab(5).Control(60).Enabled=   0   'False
      Tab(5).Control(61)=   "LocomotiveDecoderCVh(39)"
      Tab(5).Control(61).Enabled=   0   'False
      Tab(5).Control(62)=   "LocomotiveDecoderCVh(38)"
      Tab(5).Control(62).Enabled=   0   'False
      Tab(5).Control(63)=   "LocomotiveDecoderCVh(37)"
      Tab(5).Control(63).Enabled=   0   'False
      Tab(5).Control(64)=   "LocomotiveDecoderCVh(36)"
      Tab(5).Control(64).Enabled=   0   'False
      Tab(5).Control(65)=   "LocomotiveDecoderCVh(35)"
      Tab(5).Control(65).Enabled=   0   'False
      Tab(5).Control(66)=   "LocomotiveDecoderCVh(34)"
      Tab(5).Control(66).Enabled=   0   'False
      Tab(5).Control(67)=   "LocomotiveDecoderCVh(33)"
      Tab(5).Control(67).Enabled=   0   'False
      Tab(5).Control(68)=   "LocomotiveDecoderCVh(22)"
      Tab(5).Control(68).Enabled=   0   'False
      Tab(5).Control(69)=   "LocomotiveDecoderCVh(21)"
      Tab(5).Control(69).Enabled=   0   'False
      Tab(5).Control(70)=   "LocomotiveDecoderCVd(22)"
      Tab(5).Control(71)=   "LocomotiveDecoderCVd(21)"
      Tab(5).Control(72)=   "LocomotiveDecoderCVd(42)"
      Tab(5).Control(73)=   "LocomotiveDecoderCVd(41)"
      Tab(5).Control(74)=   "LocomotiveDecoderCVd(40)"
      Tab(5).Control(75)=   "LocomotiveDecoderCVd(39)"
      Tab(5).Control(76)=   "LocomotiveDecoderCVd(38)"
      Tab(5).Control(77)=   "LocomotiveDecoderCVd(37)"
      Tab(5).Control(78)=   "LocomotiveDecoderCVd(36)"
      Tab(5).Control(79)=   "LocomotiveDecoderCVd(35)"
      Tab(5).Control(80)=   "LocomotiveDecoderCVd(34)"
      Tab(5).Control(81)=   "LocomotiveDecoderCVd(33)"
      Tab(5).Control(82)=   "SpecificCVs(21)"
      Tab(5).Control(83)=   "SpecificCVs(22)"
      Tab(5).Control(84)=   "SpecificCVs(33)"
      Tab(5).Control(85)=   "SpecificCVs(34)"
      Tab(5).Control(86)=   "SpecificCVs(35)"
      Tab(5).Control(87)=   "SpecificCVs(36)"
      Tab(5).Control(88)=   "SpecificCVs(37)"
      Tab(5).Control(89)=   "SpecificCVs(38)"
      Tab(5).Control(90)=   "SpecificCVs(39)"
      Tab(5).Control(91)=   "SpecificCVs(40)"
      Tab(5).Control(92)=   "SpecificCVs(41)"
      Tab(5).Control(93)=   "SpecificCVs(42)"
      Tab(5).ControlCount=   94
      TabCaption(6)   =   "Decoder Specific"
      TabPicture(6)   =   "MainlineDecoder.frx":092C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ComboBoxDecoderName"
      Tab(6).Control(1)=   "TextBoxManufacturerName"
      Tab(6).Control(2)=   "ButtonEasyScreen"
      Tab(6).Control(3)=   "SSTab2"
      Tab(6).Control(4)=   "Label78"
      Tab(6).Control(5)=   "Label85"
      Tab(6).Control(6)=   "Label86"
      Tab(6).Control(7)=   "Label30"
      Tab(6).Control(8)=   "LabelDecoderSpecificCV(51)"
      Tab(6).ControlCount=   9
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   14
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   518
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   14
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   517
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV13D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   14
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   516
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   14
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV14"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   16
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV16"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV9D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   16
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   511
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   16
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   510
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   16
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   509
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   15
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   508
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   15
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   507
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV9D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   15
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   506
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   15
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV15"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   42
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV42"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   41
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV41"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   40
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV40"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   39
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV39"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   38
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV38"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   37
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV37"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   36
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV36"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   35
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV35"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   34
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV34"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   33
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV33"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   22
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV22"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   21
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV21"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   13
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV13"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   12
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV12"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   11
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV11"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   10
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV10"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   9
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV9"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   29
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV29"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   95
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV95"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   66
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV66"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   24
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV24"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   4
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV4"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   23
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV23"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   3
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV3"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   65
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV65"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   5
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV5"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   6
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV6"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   2
         Left            =   -72360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV2"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   106
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV106"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   105
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV105"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   8
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV8"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   7
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV7"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   19
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV18"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   18
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV18"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   17
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV17"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data SpecificCVs 
         Connect         =   "Access"
         DatabaseName    =   ""
         Exclusive       =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CV1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV65D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   65
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   38
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   65
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   65
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV95D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   95
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   58
         Text            =   "0"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   95
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   95
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV66D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   66
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   55
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   66
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   66
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV29D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   62
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   29
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   29
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton ButtonSpeedTable 
         Caption         =   "Speed Table"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -68040
         TabIndex        =   61
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   24
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   23
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV24D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   24
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   52
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   24
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   24
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV23D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   23
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   45
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   23
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   23
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV4D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   4
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   48
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV3D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   3
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   41
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV5D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   5
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV6D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   6
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   32
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV2D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   2
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   19
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   19
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV19D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   19
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   14
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   18
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   18
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   18
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "LocomotiveDecoderCV18D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   18
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   17
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   17
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "LocomotiveDecoderCV17D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   17
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "LocomotiveDecoderCV8D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   8
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "LocomotiveDecoderCV7D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   7
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV1D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   1
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   4
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
         MaxLength       =   3
         TabIndex        =   66
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV10D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   10
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   69
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   10
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   10
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV11D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   11
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   72
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   11
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   11
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV12D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   12
         Left            =   -70920
         TabIndex        =   75
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   12
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   12
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV13D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   13
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   78
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   13
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   13
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV33D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   33
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   88
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV34D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   34
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   91
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV35D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   35
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   94
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV36D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   36
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   97
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV37D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   37
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   100
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV38D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   38
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   103
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV39D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   39
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   106
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV40D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   40
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   109
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV41D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   41
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   112
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV42D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   42
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   115
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveDecoderCV21D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   21
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   81
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
         MaxLength       =   3
         TabIndex        =   84
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   21
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   22
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   33
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   34
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   35
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   36
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   37
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   38
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   39
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   40
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   41
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   42
         Left            =   -69480
         Locked          =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   21
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   22
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   33
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   34
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   35
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   36
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   37
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   38
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   39
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   40
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   41
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   42
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "LocomotiveDecoderCV105D"
         DataSource      =   "LocomotiveDatabaseDecoder"
         Height          =   285
         Index           =   105
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   23
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
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   26
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   105
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVh 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   106
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   105
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveDecoderCVb 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   106
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1335
      End
      Begin VB.ComboBox ComboBoxDecoderName 
         Height          =   315
         ItemData        =   "MainlineDecoder.frx":0948
         Left            =   -70440
         List            =   "MainlineDecoder.frx":094A
         TabIndex        =   120
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TextBoxManufacturerName 
         Height          =   285
         Left            =   -73320
         TabIndex        =   119
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   1
         Left            =   8640
         TabIndex        =   296
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   17
         Left            =   8640
         TabIndex        =   295
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   18
         Left            =   8640
         TabIndex        =   294
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   19
         Left            =   8640
         TabIndex        =   293
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   7
         Left            =   8640
         TabIndex        =   292
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   8
         Left            =   8640
         TabIndex        =   291
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   105
         Left            =   8640
         TabIndex        =   290
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   106
         Left            =   8640
         TabIndex        =   289
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   2
         Left            =   -66360
         TabIndex        =   288
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   6
         Left            =   -66360
         TabIndex        =   287
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   5
         Left            =   -66360
         TabIndex        =   286
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   65
         Left            =   -66360
         TabIndex        =   285
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   3
         Left            =   -66360
         TabIndex        =   284
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   23
         Left            =   -66360
         TabIndex        =   283
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   4
         Left            =   -66360
         TabIndex        =   282
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   24
         Left            =   -66360
         TabIndex        =   281
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   66
         Left            =   -66360
         TabIndex        =   280
         Top             =   2880
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   95
         Left            =   -66360
         TabIndex        =   279
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   29
         Left            =   -66360
         TabIndex        =   278
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   9
         Left            =   -66360
         TabIndex        =   277
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   10
         Left            =   -66360
         TabIndex        =   276
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   11
         Left            =   -66360
         TabIndex        =   275
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   12
         Left            =   -66360
         TabIndex        =   274
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   13
         Left            =   -66360
         TabIndex        =   273
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   21
         Left            =   -66360
         TabIndex        =   272
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   22
         Left            =   -66360
         TabIndex        =   271
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   33
         Left            =   -66360
         TabIndex        =   270
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   34
         Left            =   -66360
         TabIndex        =   269
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   35
         Left            =   -66360
         TabIndex        =   268
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   36
         Left            =   -66360
         TabIndex        =   267
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   37
         Left            =   -66360
         TabIndex        =   266
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   38
         Left            =   -66360
         TabIndex        =   265
         Top             =   4560
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   39
         Left            =   -66360
         TabIndex        =   264
         Top             =   4920
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   40
         Left            =   -66360
         TabIndex        =   263
         Top             =   5280
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   41
         Left            =   -66360
         TabIndex        =   262
         Top             =   5640
         Width           =   2655
      End
      Begin VB.CheckBox CheckBoxCVChanged 
         Caption         =   "Configuration Variable Changed?"
         Height          =   255
         Index           =   42
         Left            =   -66360
         TabIndex        =   261
         Top             =   6000
         Width           =   2655
      End
      Begin VB.CommandButton ButtonEasyScreenFunctions 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   118
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   87
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Easy Screen"
         Height          =   255
         Left            =   -68040
         TabIndex        =   65
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton ButtonEasyScreen 
         Caption         =   "Easy Screen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -67920
         TabIndex        =   254
         Top             =   9120
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   121
         Top             =   1920
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12091
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "CV 49 - 64"
         TabPicture(0)   =   "MainlineDecoder.frx":094C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LabelCVPrefix(64)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LabelCVPrefix(63)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LabelCVPrefix(62)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LabelCVPrefix(61)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LabelCVPrefix(60)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LabelCVPrefix(59)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LabelCVPrefix(58)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LabelCVPrefix(57)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "LabelCVPrefix(56)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LabelCVPrefix(55)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "LabelCVPrefix(54)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "LabelCVPrefix(53)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "LabelCVPrefix(52)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "LabelCVPrefix(51)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "LabelCVPrefix(50)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "LabelCVPrefix(49)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "LabelCVDescription(64)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "LabelCVDescription(63)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "LabelCVDescription(62)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "LabelCVDescription(61)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "LabelCVDescription(60)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "LabelCVDescription(59)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "LabelCVDescription(58)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "LabelCVDescription(57)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "LabelCVDescription(56)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "LabelCVDescription(55)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "LabelCVDescription(54)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "LabelCVDescription(53)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "LabelCVDescription(52)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "LabelCVDescription(51)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "LabelCVDescription(50)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "LabelCVDescription(49)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "SpecificCVs(49)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "SpecificCVs(64)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "SpecificCVs(63)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "SpecificCVs(62)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "SpecificCVs(61)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "SpecificCVs(60)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "SpecificCVs(59)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "SpecificCVs(58)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "SpecificCVs(57)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "SpecificCVs(56)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "SpecificCVs(55)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "SpecificCVs(54)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "SpecificCVs(53)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "SpecificCVs(52)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "SpecificCVs(51)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "SpecificCVs(50)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "LocomotiveDecoderCVb(49)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "Option1(64)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "Option1(63)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "Option1(62)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "Option1(61)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "Option1(60)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "Option1(59)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "Option1(58)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "Option1(57)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "Option1(56)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "Option1(55)"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "Option1(54)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "Option1(53)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "Option1(52)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "Option1(51)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "Option1(50)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "LocomotiveDecoderCVd(64)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "LocomotiveDecoderCVd(63)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "LocomotiveDecoderCVd(62)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "LocomotiveDecoderCVd(61)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "LocomotiveDecoderCVd(60)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "LocomotiveDecoderCVd(59)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "LocomotiveDecoderCVd(58)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "LocomotiveDecoderCVd(57)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "LocomotiveDecoderCVd(56)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "LocomotiveDecoderCVd(55)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "LocomotiveDecoderCVd(54)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "LocomotiveDecoderCVd(53)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "LocomotiveDecoderCVd(52)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "LocomotiveDecoderCVd(51)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "LocomotiveDecoderCVd(50)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).Control(79)=   "LocomotiveDecoderCVh(64)"
         Tab(0).Control(79).Enabled=   0   'False
         Tab(0).Control(80)=   "LocomotiveDecoderCVh(63)"
         Tab(0).Control(80).Enabled=   0   'False
         Tab(0).Control(81)=   "LocomotiveDecoderCVh(62)"
         Tab(0).Control(81).Enabled=   0   'False
         Tab(0).Control(82)=   "LocomotiveDecoderCVh(61)"
         Tab(0).Control(82).Enabled=   0   'False
         Tab(0).Control(83)=   "LocomotiveDecoderCVh(60)"
         Tab(0).Control(83).Enabled=   0   'False
         Tab(0).Control(84)=   "LocomotiveDecoderCVh(59)"
         Tab(0).Control(84).Enabled=   0   'False
         Tab(0).Control(85)=   "LocomotiveDecoderCVh(58)"
         Tab(0).Control(85).Enabled=   0   'False
         Tab(0).Control(86)=   "LocomotiveDecoderCVh(57)"
         Tab(0).Control(86).Enabled=   0   'False
         Tab(0).Control(87)=   "LocomotiveDecoderCVh(56)"
         Tab(0).Control(87).Enabled=   0   'False
         Tab(0).Control(88)=   "LocomotiveDecoderCVh(55)"
         Tab(0).Control(88).Enabled=   0   'False
         Tab(0).Control(89)=   "LocomotiveDecoderCVh(54)"
         Tab(0).Control(89).Enabled=   0   'False
         Tab(0).Control(90)=   "LocomotiveDecoderCVh(53)"
         Tab(0).Control(90).Enabled=   0   'False
         Tab(0).Control(91)=   "LocomotiveDecoderCVh(52)"
         Tab(0).Control(91).Enabled=   0   'False
         Tab(0).Control(92)=   "LocomotiveDecoderCVh(51)"
         Tab(0).Control(92).Enabled=   0   'False
         Tab(0).Control(93)=   "LocomotiveDecoderCVh(50)"
         Tab(0).Control(93).Enabled=   0   'False
         Tab(0).Control(94)=   "LocomotiveDecoderCVb(64)"
         Tab(0).Control(94).Enabled=   0   'False
         Tab(0).Control(95)=   "LocomotiveDecoderCVb(63)"
         Tab(0).Control(95).Enabled=   0   'False
         Tab(0).Control(96)=   "LocomotiveDecoderCVb(62)"
         Tab(0).Control(96).Enabled=   0   'False
         Tab(0).Control(97)=   "LocomotiveDecoderCVb(61)"
         Tab(0).Control(97).Enabled=   0   'False
         Tab(0).Control(98)=   "LocomotiveDecoderCVb(60)"
         Tab(0).Control(98).Enabled=   0   'False
         Tab(0).Control(99)=   "LocomotiveDecoderCVb(59)"
         Tab(0).Control(99).Enabled=   0   'False
         Tab(0).Control(100)=   "LocomotiveDecoderCVb(58)"
         Tab(0).Control(100).Enabled=   0   'False
         Tab(0).Control(101)=   "LocomotiveDecoderCVb(57)"
         Tab(0).Control(101).Enabled=   0   'False
         Tab(0).Control(102)=   "LocomotiveDecoderCVd(49)"
         Tab(0).Control(102).Enabled=   0   'False
         Tab(0).Control(103)=   "LocomotiveDecoderCVh(49)"
         Tab(0).Control(103).Enabled=   0   'False
         Tab(0).Control(104)=   "Option1(49)"
         Tab(0).Control(104).Enabled=   0   'False
         Tab(0).Control(105)=   "LocomotiveDecoderCVb(50)"
         Tab(0).Control(105).Enabled=   0   'False
         Tab(0).Control(106)=   "LocomotiveDecoderCVb(51)"
         Tab(0).Control(106).Enabled=   0   'False
         Tab(0).Control(107)=   "LocomotiveDecoderCVb(52)"
         Tab(0).Control(107).Enabled=   0   'False
         Tab(0).Control(108)=   "LocomotiveDecoderCVb(53)"
         Tab(0).Control(108).Enabled=   0   'False
         Tab(0).Control(109)=   "LocomotiveDecoderCVb(54)"
         Tab(0).Control(109).Enabled=   0   'False
         Tab(0).Control(110)=   "LocomotiveDecoderCVb(55)"
         Tab(0).Control(110).Enabled=   0   'False
         Tab(0).Control(111)=   "LocomotiveDecoderCVb(56)"
         Tab(0).Control(111).Enabled=   0   'False
         Tab(0).Control(112)=   "CheckBoxCVChanged(49)"
         Tab(0).Control(112).Enabled=   0   'False
         Tab(0).Control(113)=   "CheckBoxCVChanged(50)"
         Tab(0).Control(113).Enabled=   0   'False
         Tab(0).Control(114)=   "CheckBoxCVChanged(51)"
         Tab(0).Control(114).Enabled=   0   'False
         Tab(0).Control(115)=   "CheckBoxCVChanged(52)"
         Tab(0).Control(115).Enabled=   0   'False
         Tab(0).Control(116)=   "CheckBoxCVChanged(53)"
         Tab(0).Control(116).Enabled=   0   'False
         Tab(0).Control(117)=   "CheckBoxCVChanged(54)"
         Tab(0).Control(117).Enabled=   0   'False
         Tab(0).Control(118)=   "CheckBoxCVChanged(55)"
         Tab(0).Control(118).Enabled=   0   'False
         Tab(0).Control(119)=   "CheckBoxCVChanged(56)"
         Tab(0).Control(119).Enabled=   0   'False
         Tab(0).Control(120)=   "CheckBoxCVChanged(57)"
         Tab(0).Control(120).Enabled=   0   'False
         Tab(0).Control(121)=   "CheckBoxCVChanged(58)"
         Tab(0).Control(121).Enabled=   0   'False
         Tab(0).Control(122)=   "CheckBoxCVChanged(59)"
         Tab(0).Control(122).Enabled=   0   'False
         Tab(0).Control(123)=   "CheckBoxCVChanged(60)"
         Tab(0).Control(123).Enabled=   0   'False
         Tab(0).Control(124)=   "CheckBoxCVChanged(61)"
         Tab(0).Control(124).Enabled=   0   'False
         Tab(0).Control(125)=   "CheckBoxCVChanged(62)"
         Tab(0).Control(125).Enabled=   0   'False
         Tab(0).Control(126)=   "CheckBoxCVChanged(63)"
         Tab(0).Control(126).Enabled=   0   'False
         Tab(0).Control(127)=   "CheckBoxCVChanged(64)"
         Tab(0).Control(127).Enabled=   0   'False
         Tab(0).ControlCount=   128
         TabCaption(1)   =   "CV112 - 128"
         TabPicture(1)   =   "MainlineDecoder.frx":0968
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CheckBoxCVChanged(128)"
         Tab(1).Control(1)=   "CheckBoxCVChanged(127)"
         Tab(1).Control(2)=   "CheckBoxCVChanged(126)"
         Tab(1).Control(3)=   "CheckBoxCVChanged(125)"
         Tab(1).Control(4)=   "CheckBoxCVChanged(124)"
         Tab(1).Control(5)=   "CheckBoxCVChanged(123)"
         Tab(1).Control(6)=   "CheckBoxCVChanged(122)"
         Tab(1).Control(7)=   "CheckBoxCVChanged(121)"
         Tab(1).Control(8)=   "CheckBoxCVChanged(120)"
         Tab(1).Control(9)=   "CheckBoxCVChanged(119)"
         Tab(1).Control(10)=   "CheckBoxCVChanged(118)"
         Tab(1).Control(11)=   "CheckBoxCVChanged(117)"
         Tab(1).Control(12)=   "CheckBoxCVChanged(116)"
         Tab(1).Control(13)=   "CheckBoxCVChanged(115)"
         Tab(1).Control(14)=   "CheckBoxCVChanged(114)"
         Tab(1).Control(15)=   "CheckBoxCVChanged(113)"
         Tab(1).Control(16)=   "CheckBoxCVChanged(112)"
         Tab(1).Control(17)=   "Option1(112)"
         Tab(1).Control(18)=   "LocomotiveDecoderCVd(112)"
         Tab(1).Control(19)=   "LocomotiveDecoderCVh(112)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "LocomotiveDecoderCVb(112)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "LocomotiveDecoderCVb(113)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "LocomotiveDecoderCVb(114)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "LocomotiveDecoderCVb(115)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "LocomotiveDecoderCVb(116)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "LocomotiveDecoderCVb(117)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "LocomotiveDecoderCVb(118)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "LocomotiveDecoderCVb(119)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "LocomotiveDecoderCVb(120)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "LocomotiveDecoderCVb(121)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "LocomotiveDecoderCVb(122)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "LocomotiveDecoderCVb(123)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "LocomotiveDecoderCVb(124)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "LocomotiveDecoderCVb(125)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "LocomotiveDecoderCVb(126)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "LocomotiveDecoderCVb(127)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "LocomotiveDecoderCVb(128)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "LocomotiveDecoderCVh(113)"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "LocomotiveDecoderCVh(114)"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "LocomotiveDecoderCVh(115)"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "LocomotiveDecoderCVh(116)"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "LocomotiveDecoderCVh(117)"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "LocomotiveDecoderCVh(118)"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "LocomotiveDecoderCVh(119)"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "LocomotiveDecoderCVh(120)"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "LocomotiveDecoderCVh(121)"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "LocomotiveDecoderCVh(122)"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "LocomotiveDecoderCVh(123)"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "LocomotiveDecoderCVh(124)"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "LocomotiveDecoderCVh(125)"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "LocomotiveDecoderCVh(126)"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "LocomotiveDecoderCVh(127)"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "LocomotiveDecoderCVh(128)"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "LocomotiveDecoderCVd(113)"
         Tab(1).Control(54)=   "LocomotiveDecoderCVd(114)"
         Tab(1).Control(55)=   "LocomotiveDecoderCVd(115)"
         Tab(1).Control(56)=   "LocomotiveDecoderCVd(116)"
         Tab(1).Control(57)=   "LocomotiveDecoderCVd(117)"
         Tab(1).Control(58)=   "LocomotiveDecoderCVd(118)"
         Tab(1).Control(59)=   "LocomotiveDecoderCVd(119)"
         Tab(1).Control(60)=   "LocomotiveDecoderCVd(120)"
         Tab(1).Control(61)=   "LocomotiveDecoderCVd(121)"
         Tab(1).Control(62)=   "LocomotiveDecoderCVd(122)"
         Tab(1).Control(63)=   "LocomotiveDecoderCVd(123)"
         Tab(1).Control(64)=   "LocomotiveDecoderCVd(124)"
         Tab(1).Control(65)=   "LocomotiveDecoderCVd(125)"
         Tab(1).Control(66)=   "LocomotiveDecoderCVd(126)"
         Tab(1).Control(67)=   "LocomotiveDecoderCVd(127)"
         Tab(1).Control(68)=   "LocomotiveDecoderCVd(128)"
         Tab(1).Control(69)=   "Option1(113)"
         Tab(1).Control(70)=   "Option1(114)"
         Tab(1).Control(71)=   "Option1(115)"
         Tab(1).Control(72)=   "Option1(116)"
         Tab(1).Control(73)=   "Option1(117)"
         Tab(1).Control(74)=   "Option1(118)"
         Tab(1).Control(75)=   "Option1(119)"
         Tab(1).Control(76)=   "Option1(120)"
         Tab(1).Control(77)=   "Option1(121)"
         Tab(1).Control(78)=   "Option1(122)"
         Tab(1).Control(79)=   "Option1(123)"
         Tab(1).Control(80)=   "Option1(124)"
         Tab(1).Control(81)=   "Option1(125)"
         Tab(1).Control(82)=   "Option1(126)"
         Tab(1).Control(83)=   "Option1(127)"
         Tab(1).Control(84)=   "Option1(128)"
         Tab(1).Control(85)=   "SpecificCVs(112)"
         Tab(1).Control(86)=   "SpecificCVs(113)"
         Tab(1).Control(87)=   "SpecificCVs(114)"
         Tab(1).Control(88)=   "SpecificCVs(115)"
         Tab(1).Control(89)=   "SpecificCVs(116)"
         Tab(1).Control(90)=   "SpecificCVs(117)"
         Tab(1).Control(91)=   "SpecificCVs(118)"
         Tab(1).Control(92)=   "SpecificCVs(119)"
         Tab(1).Control(93)=   "SpecificCVs(120)"
         Tab(1).Control(94)=   "SpecificCVs(121)"
         Tab(1).Control(95)=   "SpecificCVs(122)"
         Tab(1).Control(96)=   "SpecificCVs(123)"
         Tab(1).Control(97)=   "SpecificCVs(124)"
         Tab(1).Control(98)=   "SpecificCVs(125)"
         Tab(1).Control(99)=   "SpecificCVs(126)"
         Tab(1).Control(100)=   "SpecificCVs(127)"
         Tab(1).Control(101)=   "SpecificCVs(128)"
         Tab(1).Control(102)=   "LabelCVPrefix(112)"
         Tab(1).Control(103)=   "LabelCVPrefix(113)"
         Tab(1).Control(104)=   "LabelCVPrefix(114)"
         Tab(1).Control(105)=   "LabelCVPrefix(115)"
         Tab(1).Control(106)=   "LabelCVPrefix(116)"
         Tab(1).Control(107)=   "LabelCVPrefix(117)"
         Tab(1).Control(108)=   "LabelCVPrefix(118)"
         Tab(1).Control(109)=   "LabelCVPrefix(119)"
         Tab(1).Control(110)=   "LabelCVPrefix(120)"
         Tab(1).Control(111)=   "LabelCVPrefix(121)"
         Tab(1).Control(112)=   "LabelCVPrefix(122)"
         Tab(1).Control(113)=   "LabelCVPrefix(123)"
         Tab(1).Control(114)=   "LabelCVPrefix(124)"
         Tab(1).Control(115)=   "LabelCVPrefix(125)"
         Tab(1).Control(116)=   "LabelCVPrefix(126)"
         Tab(1).Control(117)=   "LabelCVPrefix(127)"
         Tab(1).Control(118)=   "LabelCVPrefix(128)"
         Tab(1).Control(119)=   "LabelCVDescription(112)"
         Tab(1).Control(120)=   "LabelCVDescription(113)"
         Tab(1).Control(121)=   "LabelCVDescription(114)"
         Tab(1).Control(122)=   "LabelCVDescription(115)"
         Tab(1).Control(123)=   "LabelCVDescription(116)"
         Tab(1).Control(124)=   "LabelCVDescription(117)"
         Tab(1).Control(125)=   "LabelCVDescription(118)"
         Tab(1).Control(126)=   "LabelCVDescription(119)"
         Tab(1).Control(127)=   "LabelCVDescription(120)"
         Tab(1).Control(128)=   "LabelCVDescription(121)"
         Tab(1).Control(129)=   "LabelCVDescription(122)"
         Tab(1).Control(130)=   "LabelCVDescription(123)"
         Tab(1).Control(131)=   "LabelCVDescription(124)"
         Tab(1).Control(132)=   "LabelCVDescription(125)"
         Tab(1).Control(133)=   "LabelCVDescription(126)"
         Tab(1).Control(134)=   "LabelCVDescription(127)"
         Tab(1).Control(135)=   "LabelCVDescription(128)"
         Tab(1).ControlCount=   136
         TabCaption(2)   =   "Additional CVs"
         TabPicture(2)   =   "MainlineDecoder.frx":0984
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   128
            Left            =   -66240
            TabIndex        =   462
            Top             =   6360
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   127
            Left            =   -66240
            TabIndex        =   461
            Top             =   6000
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   126
            Left            =   -66240
            TabIndex        =   460
            Top             =   5640
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   125
            Left            =   -66240
            TabIndex        =   459
            Top             =   5280
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   124
            Left            =   -66240
            TabIndex        =   458
            Top             =   4920
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   123
            Left            =   -66240
            TabIndex        =   457
            Top             =   4560
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   122
            Left            =   -66240
            TabIndex        =   456
            Top             =   4200
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   121
            Left            =   -66240
            TabIndex        =   455
            Top             =   3840
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   120
            Left            =   -66240
            TabIndex        =   454
            Top             =   3480
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   119
            Left            =   -66240
            TabIndex        =   453
            Top             =   3120
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   118
            Left            =   -66240
            TabIndex        =   452
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   117
            Left            =   -66240
            TabIndex        =   451
            Top             =   2400
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   116
            Left            =   -66240
            TabIndex        =   450
            Top             =   2040
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   115
            Left            =   -66240
            TabIndex        =   449
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   114
            Left            =   -66240
            TabIndex        =   448
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   113
            Left            =   -66240
            TabIndex        =   447
            Top             =   960
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   112
            Left            =   -66240
            TabIndex        =   446
            Top             =   600
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   64
            Left            =   8760
            TabIndex        =   445
            Top             =   6000
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   63
            Left            =   8760
            TabIndex        =   444
            Top             =   5640
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   62
            Left            =   8760
            TabIndex        =   443
            Top             =   5280
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   61
            Left            =   8760
            TabIndex        =   442
            Top             =   4920
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   60
            Left            =   8760
            TabIndex        =   441
            Top             =   4560
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   59
            Left            =   8760
            TabIndex        =   440
            Top             =   4200
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   58
            Left            =   8760
            TabIndex        =   439
            Top             =   3840
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   57
            Left            =   8760
            TabIndex        =   438
            Top             =   3480
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   56
            Left            =   8760
            TabIndex        =   437
            Top             =   3120
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   55
            Left            =   8760
            TabIndex        =   436
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   54
            Left            =   8760
            TabIndex        =   435
            Top             =   2400
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   53
            Left            =   8760
            TabIndex        =   434
            Top             =   2040
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   52
            Left            =   8760
            TabIndex        =   433
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   51
            Left            =   8760
            TabIndex        =   432
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   50
            Left            =   8760
            TabIndex        =   431
            Top             =   960
            Width           =   2655
         End
         Begin VB.CheckBox CheckBoxCVChanged 
            Caption         =   "Configuration Variable Changed?"
            Height          =   255
            Index           =   49
            Left            =   8760
            TabIndex        =   430
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   56
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   55
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   54
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   53
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   52
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   51
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   50
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   122
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   49
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   600
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
            MaxLength       =   3
            TabIndex        =   123
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   57
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   58
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   59
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   60
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   61
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   62
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   63
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   64
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   50
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   51
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   52
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   53
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   54
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   55
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   56
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   57
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   58
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   59
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   164
            TabStop         =   0   'False
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   60
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   61
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   62
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   63
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   64
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   6000
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
            MaxLength       =   3
            TabIndex        =   127
            Top             =   960
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
            TabIndex        =   131
            Top             =   1320
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
            MaxLength       =   3
            TabIndex        =   135
            Top             =   1680
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
            TabIndex        =   139
            Top             =   2040
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
            MaxLength       =   3
            TabIndex        =   143
            Top             =   2400
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
            TabIndex        =   147
            Top             =   2760
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
            MaxLength       =   3
            TabIndex        =   151
            Top             =   3120
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
            MaxLength       =   3
            TabIndex        =   155
            Top             =   3480
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
            MaxLength       =   3
            TabIndex        =   159
            Top             =   3840
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
            TabIndex        =   163
            Top             =   4200
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
            TabIndex        =   167
            Top             =   4560
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
            TabIndex        =   171
            Top             =   4920
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
            MaxLength       =   3
            TabIndex        =   175
            Top             =   5280
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
            TabIndex        =   179
            Top             =   5640
            Width           =   1335
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
            MaxLength       =   3
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   126
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   120
            TabIndex        =   130
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   134
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   138
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   142
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   146
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   150
            Top             =   3120
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   154
            Top             =   3480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   158
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   162
            Top             =   4200
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   166
            Top             =   4560
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   120
            TabIndex        =   170
            Top             =   4920
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   174
            Top             =   5280
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   178
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   182
            Top             =   6000
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   112
            Left            =   -74880
            TabIndex        =   186
            Top             =   600
            Width           =   255
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
            MaxLength       =   3
            TabIndex        =   187
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   112
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   49
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   112
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   113
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   114
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   115
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   201
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   116
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   205
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   117
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   209
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   118
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   213
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   119
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   217
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   120
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   221
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   121
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   225
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   122
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   123
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   233
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   124
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   237
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   125
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   241
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   126
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   244
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   127
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   249
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVb 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   128
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   253
            TabStop         =   0   'False
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   113
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   114
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   115
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   200
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   116
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   204
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   117
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   208
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   118
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   119
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   216
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   120
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   121
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   224
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   122
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   228
            TabStop         =   0   'False
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   123
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   232
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   124
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   236
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   125
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   240
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   126
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   245
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   127
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   248
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox LocomotiveDecoderCVh 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   128
            Left            =   -69600
            Locked          =   -1  'True
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   6360
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
            MaxLength       =   3
            TabIndex        =   191
            Top             =   960
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
            TabIndex        =   195
            Top             =   1320
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
            TabIndex        =   199
            Top             =   1680
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
            MaxLength       =   3
            TabIndex        =   203
            Top             =   2040
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
            TabIndex        =   207
            Top             =   2400
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
            MaxLength       =   3
            TabIndex        =   211
            Top             =   2760
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
            TabIndex        =   215
            Top             =   3120
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
            MaxLength       =   3
            TabIndex        =   219
            Top             =   3480
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
            MaxLength       =   3
            TabIndex        =   223
            Top             =   3840
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
            MaxLength       =   3
            TabIndex        =   227
            Top             =   4200
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
            MaxLength       =   3
            TabIndex        =   231
            Top             =   4560
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
            MaxLength       =   3
            TabIndex        =   235
            Top             =   4920
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
            MaxLength       =   3
            TabIndex        =   239
            Top             =   5280
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
            TabIndex        =   243
            Top             =   5640
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
            MaxLength       =   3
            TabIndex        =   247
            Top             =   6000
            Width           =   1335
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
            MaxLength       =   3
            TabIndex        =   251
            Top             =   6360
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   113
            Left            =   -74880
            TabIndex        =   190
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   114
            Left            =   -74880
            TabIndex        =   194
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   115
            Left            =   -74880
            TabIndex        =   198
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   116
            Left            =   -74880
            TabIndex        =   202
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   117
            Left            =   -74880
            TabIndex        =   206
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   118
            Left            =   -74880
            TabIndex        =   210
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   119
            Left            =   -74880
            TabIndex        =   214
            Top             =   3120
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   120
            Left            =   -74880
            TabIndex        =   218
            Top             =   3480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   121
            Left            =   -74880
            TabIndex        =   222
            Top             =   3840
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   122
            Left            =   -74880
            TabIndex        =   226
            Top             =   4200
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   123
            Left            =   -74880
            TabIndex        =   230
            Top             =   4560
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   124
            Left            =   -74880
            TabIndex        =   234
            Top             =   4920
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   125
            Left            =   -74880
            TabIndex        =   238
            Top             =   5280
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   126
            Left            =   -74880
            TabIndex        =   242
            Top             =   5640
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   127
            Left            =   -74880
            TabIndex        =   246
            Top             =   6000
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Enabled         =   0   'False
            Height          =   255
            Index           =   128
            Left            =   -74880
            TabIndex        =   250
            Top             =   6360
            Width           =   255
         End
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   ""
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
         Begin VB.Data SpecificCVs 
            Connect         =   "Access"
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
            DatabaseName    =   ""
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
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(49)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   840
            TabIndex        =   429
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(50)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   840
            TabIndex        =   428
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(51)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   840
            TabIndex        =   427
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(52)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   840
            TabIndex        =   426
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(53)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   840
            TabIndex        =   425
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(54)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   840
            TabIndex        =   424
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(55)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   840
            TabIndex        =   423
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(56)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   840
            TabIndex        =   422
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(57)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   840
            TabIndex        =   421
            Top             =   3480
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(58)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   840
            TabIndex        =   420
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(59)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   840
            TabIndex        =   419
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(60)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   840
            TabIndex        =   418
            Top             =   4560
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(61)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   840
            TabIndex        =   417
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(62)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   840
            TabIndex        =   416
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(63)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   840
            TabIndex        =   415
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(64)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   840
            TabIndex        =   414
            Top             =   6000
            Width           =   2895
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV49,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   49
            Left            =   360
            TabIndex        =   413
            Top             =   600
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV50,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   50
            Left            =   360
            TabIndex        =   412
            Top             =   960
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV51,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   360
            TabIndex        =   411
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV52,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   52
            Left            =   360
            TabIndex        =   410
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV53,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   360
            TabIndex        =   409
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV54,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   54
            Left            =   360
            TabIndex        =   408
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV55,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   55
            Left            =   360
            TabIndex        =   407
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV56,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   56
            Left            =   360
            TabIndex        =   406
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV57,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   57
            Left            =   360
            TabIndex        =   405
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV58,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   58
            Left            =   360
            TabIndex        =   404
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV59,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   59
            Left            =   360
            TabIndex        =   403
            Top             =   4200
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV60,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   60
            Left            =   360
            TabIndex        =   402
            Top             =   4560
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV61,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   61
            Left            =   360
            TabIndex        =   401
            Top             =   4920
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV62,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   62
            Left            =   360
            TabIndex        =   400
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV63,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   63
            Left            =   360
            TabIndex        =   399
            Top             =   5640
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV64,"
            Enabled         =   0   'False
            Height          =   255
            Index           =   64
            Left            =   360
            TabIndex        =   398
            Top             =   6000
            Width           =   495
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV112,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   112
            Left            =   -74640
            TabIndex        =   397
            Top             =   600
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV113,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   113
            Left            =   -74640
            TabIndex        =   396
            Top             =   960
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV114,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   114
            Left            =   -74640
            TabIndex        =   395
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV115,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   115
            Left            =   -74640
            TabIndex        =   394
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV116,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   116
            Left            =   -74640
            TabIndex        =   393
            Top             =   2040
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV117,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   117
            Left            =   -74640
            TabIndex        =   392
            Top             =   2400
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV118,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   118
            Left            =   -74640
            TabIndex        =   391
            Top             =   2760
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV119,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   119
            Left            =   -74640
            TabIndex        =   390
            Top             =   3120
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV120,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   120
            Left            =   -74640
            TabIndex        =   389
            Top             =   3480
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV121,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   121
            Left            =   -74640
            TabIndex        =   388
            Top             =   3840
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV122,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   122
            Left            =   -74640
            TabIndex        =   387
            Top             =   4200
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV123,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   123
            Left            =   -74640
            TabIndex        =   386
            Top             =   4560
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV124,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   124
            Left            =   -74640
            TabIndex        =   385
            Top             =   4920
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV125,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   125
            Left            =   -74640
            TabIndex        =   384
            Top             =   5280
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV126,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   126
            Left            =   -74640
            TabIndex        =   383
            Top             =   5640
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV127,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   127
            Left            =   -74640
            TabIndex        =   382
            Top             =   6000
            Width           =   525
         End
         Begin VB.Label LabelCVPrefix 
            Caption         =   "CV128,"
            Enabled         =   0   'False
            Height          =   195
            Index           =   128
            Left            =   -74640
            TabIndex        =   381
            Top             =   6360
            Width           =   525
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(112)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   112
            Left            =   -74040
            TabIndex        =   380
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(113)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   113
            Left            =   -74040
            TabIndex        =   379
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(114)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   114
            Left            =   -74040
            TabIndex        =   378
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(115)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   115
            Left            =   -74040
            TabIndex        =   377
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(116)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   116
            Left            =   -74040
            TabIndex        =   376
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(117)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   117
            Left            =   -74040
            TabIndex        =   375
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(118)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   118
            Left            =   -74040
            TabIndex        =   374
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(119)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   119
            Left            =   -74040
            TabIndex        =   373
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(120)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   120
            Left            =   -74040
            TabIndex        =   372
            Top             =   3480
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(121)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   121
            Left            =   -74040
            TabIndex        =   371
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(122)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   122
            Left            =   -74040
            TabIndex        =   370
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(123)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   123
            Left            =   -74040
            TabIndex        =   369
            Top             =   4560
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(124)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   124
            Left            =   -74040
            TabIndex        =   368
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(125)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   125
            Left            =   -74040
            TabIndex        =   367
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(126)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   126
            Left            =   -74040
            TabIndex        =   366
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(127)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   127
            Left            =   -74040
            TabIndex        =   365
            Top             =   6000
            Width           =   2895
         End
         Begin VB.Label LabelCVDescription 
            Caption         =   "No Data"
            DataField       =   "Description"
            DataSource      =   "SpecificCVs(128)"
            Enabled         =   0   'False
            Height          =   255
            Index           =   128
            Left            =   -74040
            TabIndex        =   364
            Top             =   6360
            Width           =   2895
         End
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(14)"
         Height          =   195
         Index           =   14
         Left            =   -74280
         TabIndex        =   515
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV14,"
         Height          =   195
         Index           =   14
         Left            =   -74760
         TabIndex        =   514
         Top             =   3000
         Width           =   435
      End
      Begin VB.Line Line16 
         X1              =   -66720
         X2              =   -74760
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV16,"
         Height          =   195
         Index           =   16
         Left            =   -74760
         TabIndex        =   513
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(16)"
         Height          =   195
         Index           =   16
         Left            =   -74280
         TabIndex        =   512
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(15)"
         Height          =   195
         Index           =   15
         Left            =   -74280
         TabIndex        =   505
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV15,"
         Height          =   195
         Index           =   15
         Left            =   -74760
         TabIndex        =   504
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   $"MainlineDecoder.frx":09A0
         Height          =   615
         Left            =   -74760
         TabIndex        =   503
         Top             =   3720
         Width           =   7215
      End
      Begin VB.Line Line15 
         X1              =   -66720
         X2              =   -74760
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV42,"
         Height          =   195
         Index           =   42
         Left            =   -74760
         TabIndex        =   502
         Top             =   6000
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV41,"
         Height          =   195
         Index           =   41
         Left            =   -74760
         TabIndex        =   501
         Top             =   5640
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV40,"
         Height          =   195
         Index           =   40
         Left            =   -74760
         TabIndex        =   500
         Top             =   5280
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV39,"
         Height          =   195
         Index           =   39
         Left            =   -74760
         TabIndex        =   499
         Top             =   4920
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV38,"
         Height          =   195
         Index           =   38
         Left            =   -74760
         TabIndex        =   498
         Top             =   4560
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV37,"
         Height          =   195
         Index           =   37
         Left            =   -74760
         TabIndex        =   497
         Top             =   4200
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV36,"
         Height          =   195
         Index           =   36
         Left            =   -74760
         TabIndex        =   496
         Top             =   3840
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         AutoSize        =   -1  'True
         Caption         =   "CV35,"
         Height          =   195
         Index           =   35
         Left            =   -74760
         TabIndex        =   495
         Top             =   3480
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV34,"
         Height          =   195
         Index           =   34
         Left            =   -74760
         TabIndex        =   494
         Top             =   3120
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV33,"
         Height          =   195
         Index           =   33
         Left            =   -74760
         TabIndex        =   493
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV22,"
         Height          =   195
         Index           =   22
         Left            =   -74760
         TabIndex        =   492
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV21,"
         Height          =   195
         Index           =   21
         Left            =   -74760
         TabIndex        =   491
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV13,"
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   490
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV12,"
         Height          =   195
         Index           =   12
         Left            =   -74760
         TabIndex        =   489
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV11,"
         Height          =   195
         Index           =   11
         Left            =   -74760
         TabIndex        =   488
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         AutoSize        =   -1  'True
         Caption         =   "CV10,"
         Height          =   195
         Index           =   10
         Left            =   -74760
         TabIndex        =   487
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV9,"
         Height          =   195
         Index           =   9
         Left            =   -74760
         TabIndex        =   486
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV29,"
         Height          =   195
         Index           =   29
         Left            =   -74760
         TabIndex        =   485
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(3)"
         Height          =   195
         Index           =   3
         Left            =   -74280
         TabIndex        =   484
         Top             =   1200
         Width           =   3270
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV95,"
         Height          =   195
         Index           =   95
         Left            =   -74760
         TabIndex        =   483
         Top             =   3240
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV66,"
         Height          =   195
         Index           =   66
         Left            =   -74760
         TabIndex        =   482
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV24,"
         Height          =   195
         Index           =   24
         Left            =   -74760
         TabIndex        =   481
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV4,"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   480
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV23,"
         Height          =   195
         Index           =   23
         Left            =   -74760
         TabIndex        =   479
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV3,"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   478
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV65,"
         Height          =   195
         Index           =   65
         Left            =   -74760
         TabIndex        =   477
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV5,"
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   476
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV6,"
         Height          =   195
         Index           =   6
         Left            =   -74760
         TabIndex        =   475
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV2,"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   474
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV106,"
         Height          =   195
         Index           =   106
         Left            =   240
         TabIndex        =   473
         Top             =   4320
         Width           =   525
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV105,"
         Height          =   195
         Index           =   105
         Left            =   240
         TabIndex        =   472
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV8,"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   471
         Top             =   3360
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV7,"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   470
         Top             =   3000
         Width           =   345
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV19,"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   469
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV18,"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   468
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV17,"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   467
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label LabelCVPrefix 
         Caption         =   "CV1,"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   466
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label34 
         Caption         =   "Function Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   465
         Top             =   7080
         Width           =   1125
      End
      Begin VB.Label Label33 
         Caption         =   "Other Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   464
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label32 
         Caption         =   "Configuration Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   463
         Top             =   2280
         Width           =   1440
      End
      Begin VB.Line Line10 
         X1              =   -74760
         X2              =   -66720
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line9 
         X1              =   8280
         X2              =   240
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Voltage Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   363
         Top             =   2880
         Width           =   1050
      End
      Begin VB.Line Line7 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(65)"
         Height          =   195
         Index           =   65
         Left            =   -74280
         TabIndex        =   362
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(95)"
         Height          =   195
         Index           =   95
         Left            =   -74280
         TabIndex        =   361
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(66)"
         Height          =   195
         Index           =   66
         Left            =   -74280
         TabIndex        =   360
         Top             =   2880
         Width           =   3270
      End
      Begin VB.Line Line6 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   359
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line5 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(29)"
         Height          =   255
         Index           =   29
         Left            =   -74280
         TabIndex        =   358
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   357
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   356
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line4 
         X1              =   -66720
         X2              =   -74760
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(24)"
         Height          =   195
         Index           =   24
         Left            =   -74280
         TabIndex        =   355
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(23)"
         Height          =   195
         Index           =   23
         Left            =   -74280
         TabIndex        =   354
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Line Line3 
         X1              =   -66720
         X2              =   -74760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(4)"
         Height          =   195
         Index           =   4
         Left            =   -74280
         TabIndex        =   353
         Top             =   2040
         Width           =   3285
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   -72360
         TabIndex        =   352
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   351
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   350
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   349
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(5)"
         Height          =   195
         Index           =   5
         Left            =   -74280
         TabIndex        =   348
         Top             =   1920
         Width           =   3270
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(6)"
         Height          =   195
         Index           =   6
         Left            =   -74280
         TabIndex        =   347
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   -72360
         TabIndex        =   346
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   345
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   344
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(2)"
         Height          =   195
         Index           =   2
         Left            =   -74355
         TabIndex        =   343
         Top             =   1200
         Width           =   3360
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   342
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   8280
         X2              =   240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(19)"
         Height          =   195
         Index           =   19
         Left            =   720
         TabIndex        =   341
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Combined"
         Height          =   255
         Left            =   2640
         TabIndex        =   340
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   8280
         X2              =   240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(18)"
         Height          =   195
         Index           =   18
         Left            =   720
         TabIndex        =   339
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(17)"
         Height          =   195
         Index           =   17
         Left            =   720
         TabIndex        =   338
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(8)"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   337
         Top             =   3360
         Width           =   3270
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(7)"
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   336
         Top             =   3000
         Width           =   3270
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   6960
         TabIndex        =   335
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   5520
         TabIndex        =   334
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   4080
         TabIndex        =   333
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(1)"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   332
         Top             =   1200
         Width           =   3345
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(9)"
         Height          =   195
         Index           =   9
         Left            =   -74280
         TabIndex        =   331
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(10)"
         Height          =   195
         Index           =   10
         Left            =   -74280
         TabIndex        =   330
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(11)"
         Height          =   195
         Index           =   11
         Left            =   -74280
         TabIndex        =   329
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(12)"
         Height          =   195
         Index           =   12
         Left            =   -74280
         TabIndex        =   328
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(13)"
         Height          =   195
         Index           =   13
         Left            =   -74280
         TabIndex        =   327
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(33)"
         Height          =   195
         Index           =   33
         Left            =   -74280
         TabIndex        =   326
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(34)"
         Height          =   195
         Index           =   34
         Left            =   -74280
         TabIndex        =   325
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(35)"
         Height          =   195
         Index           =   35
         Left            =   -74280
         TabIndex        =   324
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(36)"
         Height          =   195
         Index           =   36
         Left            =   -74280
         TabIndex        =   323
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(37)"
         Height          =   195
         Index           =   37
         Left            =   -74280
         TabIndex        =   322
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(38)"
         Height          =   195
         Index           =   38
         Left            =   -74280
         TabIndex        =   321
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(39)"
         Height          =   195
         Index           =   39
         Left            =   -74280
         TabIndex        =   320
         Top             =   4920
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(40)"
         Height          =   195
         Index           =   40
         Left            =   -74280
         TabIndex        =   319
         Top             =   5280
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(41)"
         Height          =   195
         Index           =   41
         Left            =   -74280
         TabIndex        =   318
         Top             =   5640
         Width           =   3255
      End
      Begin VB.Line Line11 
         X1              =   -74760
         X2              =   -66720
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(42)"
         Height          =   195
         Index           =   42
         Left            =   -74280
         TabIndex        =   317
         Top             =   6000
         Width           =   3255
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(22)"
         Height          =   195
         Index           =   22
         Left            =   -74280
         TabIndex        =   316
         Top             =   1680
         Width           =   3285
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(21)"
         Height          =   195
         Index           =   21
         Left            =   -74280
         TabIndex        =   315
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   314
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   313
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         Caption         =   "Hexidecimal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   312
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   311
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "Binary"
         Height          =   255
         Left            =   -68040
         TabIndex        =   310
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Identification Notes:"
         Height          =   195
         Left            =   240
         TabIndex        =   309
         Top             =   4680
         Width           =   1410
      End
      Begin VB.Line Line12 
         X1              =   8280
         X2              =   240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(105)"
         Height          =   195
         Index           =   105
         Left            =   840
         TabIndex        =   308
         Top             =   3960
         Width           =   3165
      End
      Begin VB.Label LabelCVDescription 
         DataField       =   "Description"
         DataSource      =   "SpecificCVs(106)"
         Height          =   195
         Index           =   106
         Left            =   840
         TabIndex        =   307
         Top             =   4320
         Width           =   3165
      End
      Begin VB.Line Line13 
         X1              =   8280
         X2              =   240
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Caption         =   "Decimal"
         Height          =   255
         Left            =   -70920
         TabIndex        =   306
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "Speed Notes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   305
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label78 
         Caption         =   $"MainlineDecoder.frx":0A75
         Height          =   375
         Left            =   -74640
         TabIndex        =   304
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label85 
         Caption         =   "Manufacturer"
         Height          =   255
         Left            =   -73320
         TabIndex        =   303
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label86 
         Caption         =   "Decoder Type"
         Height          =   255
         Left            =   -70440
         TabIndex        =   302
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Line Line8 
         X1              =   -66720
         X2              =   -74760
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label1 
         Caption         =   "For additional help with setting the configuration variables for functions, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   301
         Top             =   6480
         Width           =   6495
      End
      Begin VB.Label Label7 
         Caption         =   "For additional help with setting the configuration variables for functions, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   300
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label Label29 
         Caption         =   "For additional help with setting the configuration variables for CV29, click on the 'easy screen' button."
         Height          =   375
         Left            =   -74760
         TabIndex        =   299
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label Label30 
         Caption         =   $"MainlineDecoder.frx":0B53
         Height          =   495
         Left            =   -74760
         TabIndex        =   298
         Top             =   8880
         Width           =   6495
      End
      Begin VB.Label LabelDecoderSpecificCV 
         DataField       =   "CV51Description"
         DataSource      =   "DecoderSpecificCVs"
         Enabled         =   0   'False
         Height          =   195
         Index           =   51
         Left            =   -74400
         TabIndex        =   297
         Top             =   2760
         Width           =   3285
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   8880
      Top             =   2820
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   8880
      Top             =   2280
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   8880
      Top             =   1680
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label LocomotiveCommunicationStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   3480
      TabIndex        =   260
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label84 
      Caption         =   "Locomotive Number"
      Height          =   255
      Left            =   1080
      TabIndex        =   259
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label46 
      Caption         =   "Update all your general configuration variables on the programming track."
      Height          =   195
      Left            =   3480
      TabIndex        =   257
      Top             =   120
      Width           =   5160
   End
End
Attribute VB_Name = "MainlineDecoder"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim cv(1 To 128) As String

Private Sub ButtonClose_Click()

    Let MainlineDiesel.TextBoxLastTimeDecoderViewed.Text = Date$ & " at " & Time$
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
            If TemporaryScreen = "Mainline Decoder Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Decoder Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                CommunicationSetting.Show vbmodelss
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
                Let Ini.Value = "Mainline Decoder Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Decoder Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub ButtonEasyScreen_Click()
    
    Load MainlineEasyScreenSpecificCvs
    MainlineEasyScreenSpecificCvs.Show vbModeless

End Sub

Private Sub ButtonEasyScreenFunctions_Click()

    Load MainlineEasyScreenFunctions
    MainlineEasyScreenFunctions.Show vbModeless


End Sub

Private Sub ButtonExport_Click()

Let MainlineDiesel.TextBoxLastTimeDecoderWritten.Text = Date$ & " at " & Time$

Let MousePointer = vbHourglass
Let ButtonClose.Enabled = False
Let ButtonUpdate.Enabled = False
Let ButtonExport.Enabled = False

For Y = 1 To 128

If cv(Y) = "Yes" Then
    If Y <> 29 Then GoTo 10
        Let onebyted.Text = Int(Val(LocomotiveNumber.Text) / 256)
        Let twoByteD.Text = Val(LocomotiveNumber) - (Val(onebyted.Text) * 256)
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
    
        Call SetErrorByte
        Call MainScreen.SendCommandviaTrackQ
        
10
    
    ' Now that the Locomotive Communication Window has be updated...
    
        End If
    For tt = 1 To 500
        DoEvents
    Next tt

Next Y

Let ButtonClose.Enabled = True
Let ButtonUpdate.Enabled = True
Let ButtonExport.Enabled = True

Let MousePointer = vbDefault

End Sub


Private Sub ButtonList_Click()
 
Load LocomotiveCVSpreadsheet
LocomotiveCVSpreadsheet.Show vbModeless

End Sub

Private Sub ButtonPrint_Click()

    mainlinedeocder.PrintForm
    
End Sub

Private Sub ButtonSpeedTable_Click()

    Load MainlineSpeedTable
    MainlineSpeedTable.Show vbModeless

End Sub

Private Sub ButtonUpdate_Click()

' -------------------------------------------------------------------------------------------------------------------------
' Enabled Property
'
'Returns or sets a value that determines whether a form or control can respond to user-generated events.
'
' The syntax is a follows:
'
'   object.Enabled [= boolean]
'
'   The Enabled property syntax has these parts: where 'object' is an object expression that evaluates to an object in the
' Applies To list.  If object is omitted, the form associated with the active form module is assumed to be object; and
' 'boolean' is a Boolean expression that specifies whether object can respond to user-generated events. The Enabled
' property allows forms and controls to be enabled or disabled at run time.  For example, you can disable objects that
' don't apply to the current state of the application.  You can also disable a control used purely for display purposes,
' such as a text box that provides read-only information. Disabling a Timer control by setting Enabled to False cancels
' the countdown set up by the control's Interval property. For a Menu control, Enabled is normally read/write at run time.  But Enabled is read-only for menu items that are exposed or supplied by Visual Basic to add-ins, such as the Add-In Manager command on the Add-Ins menu.
'
' My Programming Notes
'
' We need to control the buttons that are available to the user while a subroutine is going.

    Let ButtonExport.Enabled = False
'   Let ButtonImport.Enabled = False
    Let ButtonClose.Enabled = False
    Let ButtonUpdate.Enabled = False
    Let ButtonSpeedTable.Enabled = False
    
' -------------------------------------------------------------------------------------------------------------------------
' Updating the Record
'
    LocomotiveDatabaseDecoder.Recordset.Edit
    LocomotiveDatabaseDecoder.Recordset.Update
       
' -------------------------------------------------------------------------------------------------------------------------
' Enabled Property
'
'Returns or sets a value that determines whether a form or control can respond to user-generated events.
'
' The syntax is a follows:
'
'   object.Enabled [= boolean]
'
'   The Enabled property syntax has these parts: where 'object' is an object expression that evaluates to an object in the
' Applies To list.  If object is omitted, the form associated with the active form module is assumed to be object; and
' 'boolean' is a Boolean expression that specifies whether object can respond to user-generated events. The Enabled
' property allows forms and controls to be enabled or disabled at run time.  For example, you can disable objects that
' don't apply to the current state of the application.  You can also disable a control used purely for display purposes,
' such as a text box that provides read-only information. Disabling a Timer control by setting Enabled to False cancels
' the countdown set up by the control's Interval property. For a Menu control, Enabled is normally read/write at run time.  But Enabled is read-only for menu items that are exposed or supplied by Visual Basic to add-ins, such as the Add-In Manager command on the Add-Ins menu.
'
' My Programming Notes
'
' We need to control the buttons that are available to the user while a subroutine is going.

    Let ButtonExport.Enabled = True
'    Let ButtonImport.Enabled = True
    Let ButtonClose.Enabled = True
    Let ButtonUpdate.Enabled = False
    Let ButtonSpeedTable.Enabled = True


' -------------------------------------------------------------------------------------------------------------------------
' End Subroutine
'
' Ends a procedure or block.
'
' The syntax is as follows:
'
'   End Sub
'
'   The End statement syntax has these forms: where required to end a Sub statement. For Visual Basic in-process OLE
' server (DLL) considerations and restrictions that apply to this topic, see Make OLE DLL File Command. When executed,
' the End statement resets all module-level variables and all static local variables in all modules.  If you need to
' preserve the value of these variables, use the Stop Statement instead.  You can then resume execution while preserving
' the value of those variables.

End Sub








Private Sub ComboBoxDecoderName_Click()

Let MousePointer = vbArrowHourglass

DecoderDatabase.Recordset.MoveFirst

Do While Not (DecoderDatabase.Recordset.Fields("Manufacturer") = TextBoxManufacturerName.Text And DecoderDatabase.Recordset.Fields("Decoder") = ComboBoxDecoderName.Text)
    DecoderDatabase.Recordset.MoveNext
Loop

Update:

If TempDatabase = 105 Then
    Let TempDatabase = 106
ElseIf TempDatabase = 95 Then
    Let TempDatabase = 105
ElseIf TempDatabase = 66 Then
    Let TempDatabase = 95
ElseIf TempDatabase = 65 Then
    Let TempDatabase = 66
ElseIf TempDatabase = 42 Then
    Let TempDatabase = 65
ElseIf TempDatabase = 41 Then
    Let TempDatabase = 42
ElseIf TempDatabase = 40 Then
    Let TempDatabase = 41
ElseIf TempDatabase = 39 Then
    Let TempDatabase = 40
ElseIf TempDatabase = 38 Then
    Let TempDatabase = 39
ElseIf TempDatabase = 37 Then
    Let TempDatabase = 38
ElseIf TempDatabase = 36 Then
    Let TempDatabase = 37
ElseIf TempDatabase = 35 Then
    Let TempDatabase = 36
ElseIf TempDatabase = 35 Then
    Let TempDatabase = 35
ElseIf TempDatabase = 34 Then
    Let TempDatabase = 35
ElseIf TempDatabase = 33 Then
    Let TempDatabase = 34
ElseIf TempDatabase = 29 Then
    Let TempDatabase = 33
ElseIf TempDatabase = 24 Then
    Let TempDatabase = 29
ElseIf TempDatabase = 23 Then
    Let TempDatabase = 24
ElseIf TempDatabase = 22 Then
    Let TempDatabase = 23
ElseIf TempDatabase = 21 Then
    Let TempDatabase = 22
ElseIf TempDatabase = 19 Then
    Let TempDatabase = 21
ElseIf TempDatabase = 18 Then
    Let TempDatabase = 19
ElseIf TempDatabase = 17 Then
    Let TempDatabase = 18
ElseIf TempDatabase = 16 Then
    Let TempDatabase = 17
ElseIf TempDatabase = 15 Then
    Let TempDatabase = 16
ElseIf TempDatabase = 14 Then
    Let TempDatabase = 15
ElseIf TempDatabase = 13 Then
    Let TempDatabase = 14
ElseIf TempDatabase = 12 Then
    Let TempDatabase = 13
ElseIf TempDatabase = 11 Then
    Let TempDatabase = 12
ElseIf TempDatabase = 10 Then
    Let TempDatabase = 11
ElseIf TempDatabase = 9 Then
    Let TempDatabase = 10
ElseIf TempDatabase = 8 Then
    Let TempDatabase = 9
ElseIf TempDatabase = 7 Then
    Let TempDatabase = 8
ElseIf TempDatabase = 6 Then
    Let TempDatabase = 7
ElseIf TempDatabase = 5 Then
    Let TempDatabase = 6
ElseIf TempDatabase = 4 Then
    Let TempDatabase = 5
ElseIf TempDatabase = 3 Then
    Let TempDatabase = 4
ElseIf TempDatabase = 2 Then
    Let TempDatabase = 3
ElseIf TempDatabase = 1 Then
    Let TempDatabase = 2
ElseIf TempDatabase = 0 Then
    Let TempDatabase = 1
End If

    SpecificCVs(TempDatabase).Recordset.MoveFirst
    While SpecificCVs(TempDatabase).Recordset.Fields("RecordNumber") <> DecoderDatabase.Recordset.Fields("RecordNumber")
        SpecificCVs(TempDatabase).Recordset.MoveNext
    Wend
    DoEvents
    If SpecificCVs(TempDatabase).Recordset.Fields("Used") = True Then
        Let LabelCVPrefix(TempDatabase).Enabled = True
        Let LabelCVdescription(TempDatabase).Enabled = True
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = True
        Let LocomotiveDecoderCVh(TempDatabase).Enabled = True
        Let LocomotiveDecoderCVb(TempDatabase).Enabled = True
     Else
        Let LabelCVPrefix(TempDatabase).Enabled = False
        Let LabelCVdescription(TempDatabase).Enabled = False
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = False
        Let LocomotiveDecoderCVh(TempDatabase).Enabled = False
        Let LocomotiveDecoderCVb(TempDatabase).Enabled = False
    End If

If TempDatabase <> 106 Then GoTo Update

For TempDatabase = 49 To 64
    SpecificCVs(TempDatabase).Recordset.MoveFirst
    While SpecificCVs(TempDatabase).Recordset.Fields("RecordNumber") <> DecoderDatabase.Recordset.Fields("RecordNumber")
        SpecificCVs(TempDatabase).Recordset.MoveNext
    Wend
    DoEvents
    If SpecificCVs(TempDatabase).Recordset.Fields("Used") = True Then
        Let cv(TempDatabase) = "Yes"
        Let LabelCVPrefix(TempDatabase).Enabled = True
        Let LabelCVdescription(TempDatabase).Enabled = True
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = True
        If SpecificCVs(TempDatabase).Recordset.Fields("BitUsed") = True Then
            Let Option1(TempDatabase).Enabled = True
        Else
            Let Option1(TempDatabase).Enabled = False
        End If
    Else
        Let cv(TempDatabase) = "No"
        Let LabelCVPrefix(TempDatabase).Enabled = False
        Let LabelCVdescription(TempDatabase).Enabled = False
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = False
        Let Option1(TempDatabase).Enabled = False
    End If
Next TempDatabase

For TempDatabase = 112 To 128
    SpecificCVs(TempDatabase).Recordset.MoveFirst
    While SpecificCVs(TempDatabase).Recordset.Fields("RecordNumber") <> DecoderDatabase.Recordset.Fields("RecordNumber")
        SpecificCVs(TempDatabase).Recordset.MoveNext
    Wend
    DoEvents
    If SpecificCVs(TempDatabase).Recordset.Fields("Used") = True Then
        Let cv(TempDatabase) = "Yes"
        Let LabelCVPrefix(TempDatabase).Enabled = True
        Let LabelCVdescription(TempDatabase).Enabled = True
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = True
        If SpecificCVs(TempDatabase).Recordset.Fields("BitUsed") = True Then
            Let Option1(TempDatabase).Enabled = True
        Else
            Let Option1(TempDatabase).Enabled = False
        End If
    Else
        Let cv(TempDatabase) = "No"
        Let LabelCVPrefix(TempDatabase).Enabled = False
        Let LabelCVdescription(TempDatabase).Enabled = False
        Let LocomotiveDecoderCVd(TempDatabase).Enabled = False
        Let Option1(TempDatabase).Enabled = False
    End If
Next TempDatabase

ButtonUpdate.Enabled = False

Let MousePointer = Default

End Sub


Private Sub Command1_Click()

    Load MainlineEasyScreenConsistFunctions
    MainlineEasyScreenConsistFunctions.Show vbModeless

End Sub

Private Sub Command2_Click()

    Load MainlineEasyScreenConfiguration
    MainlineEasyScreenConfiguration.Show vbModeless

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
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Add to Screen Stack
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryScreen As String
    Dim TemporaryCounter As Integer
    Dim TemporaryRecord As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Loop for Checking Sceen Stack
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Screen Stack"
    For TemporaryCounter = 0 To 9
        Let Ini.Parameter = CStr(TemporaryCounter)
        Let TemporaryScreen = Ini.Value
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Already Present in INI
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If TemporaryScreen = "Mainline Decoder Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Decoder Screen"
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
            Let Ini.Value = "Programming Decoder Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Programming Decoder Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Programming Decoder Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

   
   
   
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set properties of the objects in the window
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MousePointer = vbHourglass
    Let ButtonClose.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the database source
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainlineDiesel!CheckAddressShort.Value = vbChecked Then
        Let LocomotiveDatabaseDecoder.RecordSource = "LocomotiveDecodersShort"
    Else
        Let LocomotiveDatabaseDecoder.RecordSource = "LocomotiveDecodersLong"
    End If
    Let Temporary = "LocomotiveNumber like " + Str$(Val(MainlineDiesel!LocomotiveNumberGO.Text))
    LocomotiveDatabaseDecoder.Recordset.FindFirst Temporary
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Return properties of objects in windows.
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let ButtonClose.Enabled = True
    Let MousePointer = vbDefault
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Decoder Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineDecoder.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineDecoder.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineDecoder.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineDecoder.Height)

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
                Let Ini.Value = "Mainline Decoder Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Decoder Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineDecoder.Hide
    'unload Mainlinedecoder

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
    Let Ini.Application = "Mainline Decoder Screen"
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
        MainlineDecoder.Left = (Screen.Width - Width) / 2
        MainlineDecoder.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineDecoder.Width > Screen.Width Then
            Let MainlineDecoder.Left = Screen.Width - MainlineDecoder.Width
        Else
            Let MainlineDecoder.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineDecoder.Height > Screen.Height Then
            Let MainlineDecoder.Top = Screen.Height - MainlineDecoder.Height
        Else
            Let MainlineDecoder.Top = Val(TemporaryValueTop)
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
        Call InitializeBalloonHelpPart1
        Call InitializeBalloonHelpPart2
        Call InitializeBalloonHelpPart3
        Call InitializeBalloonHelpPart4
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let DecoderDatabase.DatabaseName = App.Path$ & "\Databases\SpecificDecoderCvs.mdb"
    Let LocomotiveDatabaseDecoder.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDecoders.mdb"
    Let SpecificCVs(1).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(2).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(3).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(4).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(5).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(6).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(7).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(8).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(9).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(10).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(11).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(12).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(13).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(14).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(15).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(16).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(17).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(18).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(19).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(21).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(22).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(23).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(24).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(29).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(33).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(34).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(35).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(36).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(37).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(38).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(39).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(40).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(41).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(42).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(49).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(50).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(51).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(52).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(53).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(54).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(55).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(56).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(57).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(58).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(59).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(60).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(61).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(62).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(63).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(64).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(65).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(66).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(95).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    
    Let SpecificCVs(105).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(106).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    
    Let SpecificCVs(112).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(113).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(114).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(115).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(116).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(117).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(118).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(119).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(120).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(121).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(122).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(123).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(124).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(125).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(126).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(127).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    Let SpecificCVs(128).DatabaseName = App.Path$ & "\Databases\SpecificDecoderCVs.mdb"
    
    
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


Private Sub Form_Resize()

    If MainlineDecoder.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineDecoder.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Decoder Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineDecoder.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineDecoder.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineDecoder.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineDecoder.Height)
 
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

























Private Sub LocomotiveDecoderCVc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If Index = 18 Then
    Let temp17 = Int(Val(LocomotiveDecoderCVc(18).Text) / 256)
    Let temp18 = Val(LocomotiveDecoderCVc(18).Text) - (256 * temp17)
    Let temp17 = temp17 + 128 + 64

    Let LocomotiveDecoderCVd(17).Text = temp17
    Let LocomotiveDecoderCVd(18).Text = temp18
End If
 
End Sub


Private Sub LocomotiveDecoderCVd_Change(Index As Integer)

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




' -------------------------------------------------------------------------------------------------------------------------------------------
' Special Routines
' -------------------------------------------------------------------------------------------------------------------------------------------
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

' -------------------------------------------------------------------------------------------------------------------------------------------
' Special Routines
' -------------------------------------------------------------------------------------------------------------------------------------------
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
        Let TextBoxManufacturerName.Text = "Manufacturer is Arnold - Rivarossi"
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
        Let TextBoxManufacturerName.Text = "Ranfixx Technologies (Wangrow)"
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
        Let TextBoxManufacturerName.Text = "ModelHeisenbaln"
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
        Let Label43.Caption = Label43.Caption + "CV8, Manufacturer is ZTC" + Chr$(13)
        Let TextBoxManufacturerName.Text = "ZTC"
        If Val(LocomotiveDecoderCVd(7).Text) = 1 Then
            Let Label43.Caption = Label43.Caption + "CV7, Decoder version is unknown"
        End If
    End If
End If


' -------------------------------------------------------------------------------------------------------------------------------------------
' Special Routines
' -------------------------------------------------------------------------------------------------------------------------------------------

If Index = 17 Or _
   Index = 18 Then

    Let LocomotiveDecoderCVc(18).Text = (Val(LocomotiveDecoderCVd(17).Text) - 128 - 64) * 256
    Let LocomotiveDecoderCVc(18).Text = Val(LocomotiveDecoderCVc(18).Text) + Val(LocomotiveDecoderCVd(18).Text)
        
End If


' -------------------------------------------------------------------------------------------------------------------------------------------
' Special Routines
' -------------------------------------------------------------------------------------------------------------------------------------------

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

Private Sub LocomotiveDecoderCVd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Noting which CV changed
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let CheckBoxCVChanged(Index).Value = vbChecked
    Let ButtonUpdate.Enabled = True
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check for Valid Entry
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If Val(LocomotiveDecoderCVd(Index).Text) < SpecificCVs(Index).Recordset.Fields("ValueLow").Value Then
        If DisplayErrorMessage.Value = vbChecked Then
            Let Message = "The value," & Str(LocomotiveDecoderCVd(Index).Text) & ", entered for the Configuration Varible" & Str$(Format(Index, "000")) & " is not valid. You" + Chr$(13)
            Let Message = Message + " enter a number less then" & Str$(SpecificCVs(Index).Recordset.Fields("ValueLow").Value) & ". Please use a value between" & Str$(ConfigurationVariableLow) & " and" & Str$(SpecificCVs(Index).Recordset.Fields("ValueHigh").Value) & " only. " & Chr$(13)
            Let Message = Message + "Configuration Variable" & Str$(Format(Index, "000")) & " will be reset to" & Str$(ConfigurationVariableLow) & "." + Chr$(13)
            Let Message = Message + Chr$(13) & "Do you wish to continue to receive these warning messages?"
            Let ReturnMessage = MsgBox(Message, vbYesNo + vbInformation, "CV - Entry Error")
            If ReturnMessage = vbNo Then
                Let DisplayErrorMessage.Value = vbUnchecked
            End If
        End If
        LocomotiveDecoderCVd(Index).Text = SpecificCVs(Index).Recordset.Fields("ValueLow").Value
    End If
    
    If Val(LocomotiveDecoderCVd(Index).Text) > SpecificCVs(Index).Recordset.Fields("ValueHigh").Value Then
        If DisplayErrorMessage.Value = vbChecked Then
            Let Message = "The value," & Str(LocomotiveDecoderCVd(Index).Text) & ", entered for the Configuration Variable" & Str$(Format(Index, "000")) & " is not valid. You" + Chr$(13)
            Let Message = Message + " enter a number greater then" & Str$(SpecificCVs(Index).Recordset.Fields("ValueHigh").Value) & ". Please use a value between" & Str$(ConfigurationVariableLow) & " and" & Str$(SpecificCVs(Index).Recordset.Fields("ValueHigh").Value) & " only. " & Chr$(13)
            Let Message = Message + "Configuration Variable" & Str$(Format(Index, "000")) & " will be reset to" & Str$(ConfigurationVariableLow) & "." + Chr$(13)
            Let Message = Message + Chr$(13) & "Do you wish to continue to receive these warning messages?"
            Let ReturnMessage = MsgBox(Message, vbYesNo + vbInformation, "CV - Entry Error")
            If ReturnMessage = vbNo Then
                Let DisplayErrorMessage.Value = vbUnchecked
            End If
        End If
        Let LocomotiveDecoderCVd(Index) = SpecificCVs(Index).Recordset.Fields("ValueHigh").Value
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

Private Sub Option1_Click(Index As Integer)
     
    Let ButtonEasyScreen.Enabled = True
 
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




Private Sub TimerInitialize_Timer()

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



Public Sub SetErrorByte()
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

End Sub

Public Sub SendCommand()
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

End If
      
End Sub




Private Sub InitializeBalloonHelpPart1()

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

        Let TemporaryText1 = "This text box displays the locomotive decoder number" & vbCrLf & "or the locomotive number."
        Let TemporaryText2 = "Locomotive (decoder) Number"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveNumber)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveNumber, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Tab
        Let TemporaryText1 = "This tab selection groups the configuraton variables" & vbCrLf & "into seven different areas, 'click' on the tab" & vbCrLf & "to see the configuration variables in that group."
        Let TemporaryText2 = "Tab Groups for Configuration Variables"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SSTab1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SSTab1, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'List
        Let TemporaryText1 = "This button displays a new window to list all the" & vbCrLf & "locomotives in your database" & vbCrLf & "based on thier activation or decoder equip."
        Let TemporaryText2 = "List Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonList, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Export
        Let TemporaryText1 = "This button will cause all the configuration variables" & vbCrLf & "to be exported (sent out) to the command station (to the locomotive)."
        Let TemporaryText2 = "Export Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonExport)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonExport, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Update
        Let TemporaryText1 = "This button will cause an update to the database" & vbCrLf & "with any changes you have made."
        Let TemporaryText2 = "Update Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonUpdate)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonUpdate, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Print
        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Close
        Let TemporaryText1 = "This button will close the Mainline Decoder Window" & vbCrLf & "and return you to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'SpeedTable
        Let TemporaryText1 = "This button will open the Mainline Speed Table" & vbCrLf & "Window. It is used to change the loaocomotive's speed table values" & vbCrLf & "(conguration variables 66-94)."
        Let TemporaryText2 = "Speed Table Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonSpeedTable)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonSpeedTable, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Command1
        Let TemporaryText1 = "This button will open an Easy Screen for setting" & vbCrLf & "consist functions (front and rear headlight)."
        Let TemporaryText2 = "Consist Functions"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(Command1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(Command1, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'Command2
        Let TemporaryText1 = "This button will open an Easy Screen for setting" & vbCrLf & "general configurations for the decoder."
        Let TemporaryText2 = "General Configuration"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(Command2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(Command2, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'buttoneastscreenfunctions
        Let TemporaryText1 = "This button will open an Easy Screen for setting" & vbCrLf & " for consist functions (one through eight)."
        Let TemporaryText2 = "Consist Functions"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonEasyScreenFunctions)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonEasyScreenFunctions, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'SpeedTable
        Let TemporaryText1 = "This button will open an Easy Screen for setting" & vbCrLf & "manufacturer specific configuration variables."
        Let TemporaryText2 = "Manufacturer Specific CVs"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonEasyScreen)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonEasyScreen, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'textbocManufacturername
        Let TemporaryText1 = "This textbox will display the manufacturer" & vbCrLf & "name of the decoder."
        Let TemporaryText2 = "Manufacturer Name of Decoder"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxManufacturerName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxManufacturerName, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'ComboBoxdecodername
        Let TemporaryText1 = "This combination box will display the possible decoders models. Selecting the correct decoder will enable specific manufacturer's configuration variable."
        Let TemporaryText2 = "Decoder Model"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBoxDecoderName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxDecoderName, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub

Private Sub InitializeBalloonHelpPart2()

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
'1
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one, the short address of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 1"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one, the short address of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 1"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one, the short address of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 1"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'2
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "two, the starting motor voltage of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 2"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "two, the starting motor voltage of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 2"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "two, the starting motor voltage of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 2"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'3
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 3"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 3"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 3"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'4
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 4"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 4"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "three, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 4"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'5
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "five, the maximum motor voltage of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 5"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "five, the maximum motor voltage of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 5"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "five, the maximum motor voltage of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 5"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'6
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "six, the middle motor voltage of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 6"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "six, the middle motor voltage of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 6"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "six, the middle motor voltage of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 6"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'7
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seven, the manuafacturer version number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 7"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seven, the manuafacturer version number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 7"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seven, the manuafacturer version number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 7"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'8
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eight, the manufacturer identification number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 8"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eight, the manufacturer identification number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 8"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eight, the manufacturer identification number in DECIMAL notation." & vbCrLf & "You cannot change this value."
        Let TemporaryText2 = "Configuration Variable 8"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'9
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nine, the pulse width modulation period of the decoder" & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 9"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(9))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(9), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nine, the pulse width modulation period of the decoder" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 9"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(9))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(9), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nine, the pulse width modulation period of the decoder" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 9"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(9))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(9), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'10
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ten, the feedback cut of of the decoder" & vbCrLf & " in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 10"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(10))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(10), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ten, the feedback cut of of the decoder" & vbCrLf & " in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 10"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(10))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(10), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ten, the feedback cut of of the decoder" & vbCrLf & " in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 10"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(10))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(10), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'11
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eleven, the packet time out value of the decoder" & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 11"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(11))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(11), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eleven, the packet time out value of the decoder" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 11"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(11))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(11), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eleven, the packet time out value of the decoder" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 11"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(11))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(11), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'12
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twelve, the power source conversion of the decoder" & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 12"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(12))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(12), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twelve, the power source conversion of the decoder" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 12"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(12))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(12), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twelve, the power source conversion of the decoder" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 12"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(12))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(12), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'13
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirteen, the analog mode function status of the decoder " & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 13"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(13))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(13), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirteen, the analog mode function status of the decoder " & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 13"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(13))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(13), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirteen, the analog mode function status of the decoder " & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 13"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(13))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(13), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'14
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourteen, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 14"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(14))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(14), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourteen, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 14"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(14))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(14), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourteen, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 14"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(14))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(14), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'15
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fifteen, the decoder locking feature, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 15"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(15))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(15), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fifteen, the decoder locking feature, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 15"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(15))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(15), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fifteen, the decoder locking feature, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 15"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(15))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(15), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'16
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixteen, the decoder locking feature, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 16"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(16))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(16), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixteen, the decoder locking feature, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 16"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(16))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(16), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixteen, the decoder locking feature, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 16"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(16))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(16), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'17
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seventeen, part of the extended address of the decoder, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 17"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(17))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(17), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seventeen, part of the extended address of the decoder, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 17"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(17))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(17), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seventeen, part of the extended address of the decoder, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 17"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(17))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(17), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'18
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eighteen, part of the extented address of the decoder, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 18"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(18))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(18), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eighteen, part of the extented address of the decoder, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 18"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(18))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(18), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "eighteen, part of the extented address of the decoder, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 18"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(18))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(18), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "seventeen and eighteen, the extented address of the decoder, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 18"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVc(18))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVc(18), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'19
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nineteen, the consist address of the decoder, in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 19"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(19))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(19), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nineteen, the consist address of the decoder, in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 19"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(19))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(19), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "nineteen, the consist address of the decoder, in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 19"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(19))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(19), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'21
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-one, the functions active when the decoder is consisted" & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 21"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(21))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(21), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-one, the functions active when the decoder is consisted" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 21"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(21))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(21), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-one, the functions active when the decoder is consisted" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 21"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(21))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(21), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'22
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-two, the headlight functions active when the decoder is" & vbCrLf & "consisted in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 22"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(22))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(22), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-two, the headlight functions active when the decoder is" & vbCrLf & "consisted in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 22"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(22))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(22), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-two, the headlight functions active when the decoder is" & vbCrLf & "consisted in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 22"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(22))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(22), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'23
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 23"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(23))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(23), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 23"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(23))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(23), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-three, the motor voltage acceleration rate of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 23"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(23))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(23), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'24
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-four, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 24"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(24))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(24), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-four, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 24"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(24))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(24), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-four, the motor voltage deacceleration rate of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 24"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(24))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(24), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'29
        'Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-nine, the general configuration of the decoder," & vbCrLf & "in DECIMAL notation."
        'Let TemporaryText2 = "Configuration Variable 29"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(29))
        'Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(29), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        'Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-nine, the general configuratin of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        'Let TemporaryText2 = "Configuration Variable 29"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(29))
        'Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(29), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        'Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-nine, the general configuration of the decoder," & vbCrLf & "in BINARY notation."
        'Let TemporaryText2 = "Configuration Variable 29"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(29))
        'Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(29), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")


End Sub

Private Sub InitializeBalloonHelpPart3()

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
'33
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-three, the output for the front light function of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 33"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(33))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(33), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-three, the output for the front function of the decoder" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 33"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(33))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(33), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-three, the output for the front light function of the decoder" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 33"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(33))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(33), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'34
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-four, the output for the rear light functon of the decoder" & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 34"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(34))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(34), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-four, the output for the rear light function of the decoder" & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 34"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(34))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(34), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "twenty-nine, the output for the rear light function of the decoder" & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 34"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(34))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(34), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'35
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-five, the outut for function one of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 35"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(35))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(35), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-five, the output for function one of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 35"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(35))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(35), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-five, the output for function one of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 35"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(35))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(35), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'36
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-six, the outut for function two of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 36"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(36))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(36), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-six, the output for function two of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 36"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(36))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(36), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-six, the output for function two of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 36"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(36))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(36), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'37
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-seven, the outut for function three of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 37"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(37))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(37), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-seven, the output for function three of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 37"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(37))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(37), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-seven, the output for function three of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 37"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(37))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(37), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'38
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-eight, the outut for function four of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 38"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(38))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(38), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-eight, the output for function four of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 38"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(38))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(38), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-eight, the output for function four of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 38"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(38))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(38), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'39
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-nine, the outut for function five of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 39"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(39))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(39), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-nine, the output for function five of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 39"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(39))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(39), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "thirty-nine, the output for function five of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 39"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(39))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(39), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'40
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty, the outut for function six of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 40"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(40))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(40), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty, the output for function six of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 40"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(40))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(40), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty, the output for function six of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 40"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(40))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(40), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'41
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-one, the outut for function seven of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 41"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(41))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(41), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-one, the output for function seven of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 41"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(41))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(41), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-one, the output for function seven of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 41"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(41))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(41), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'42
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-two, the outut for function eight of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 42"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(42))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(42), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-two, the output for function eight of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 42"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(42))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(42), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "fourty-two, the output for function eight of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 42"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(42))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(42), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'65
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-five, the kick start voltage of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 65"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(65))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(65), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-five, the kick start voltage of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 65"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(65))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(65), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-five, the kick start voltage of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 65"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(65))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(65), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'66
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-six, the forward speed trim of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 66"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(66))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(66), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-six, the forward speed trim of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 66"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(66))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(66), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "sixty-six, the forward speed trim of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 66"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(66))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(66), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'95
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ninety-five, the reverse speed trim of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 95"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(95))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(95), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ninety-five, the forward speed trim of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 95"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(95))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(95), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "ninety-five, the forward speed trim of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 95"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(95))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(95), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'105
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and five, the user identifier one of the decoder." & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 105"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(105))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(105), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and five, the user identifier one of the decoder." & vbCrLf & "in HEXIECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 105"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(105))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(105), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and five, the user identifier one of the decoder." & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 105"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(105))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(105), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'106
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and six, the user identifier two of the decoder," & vbCrLf & "in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 106"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(106))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(106), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and six, the user identifier two of the decoder," & vbCrLf & "in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable 106"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(106))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(106), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for configuraton variable" & vbCrLf & "one hundred and six, the user identifier two of the decoder," & vbCrLf & "in BINARY notation."
        Let TemporaryText2 = "Configuration Variable 106"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(106))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(106), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End Sub

Private Sub InitializeBalloonHelpPart4()

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

    For Temporary = 49 To 64
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in BINARY notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    Next Temporary
    For Temporary = 112 To 128
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in DECIMAL notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVd(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVd(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in HEXIDECIMAL notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVh(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVh(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box displays the value for the" & vbCrLf & "specific manufacturer configuraton variable" & vbCrLf & " in BINARY notation."
        Let TemporaryText2 = "Configuration Variable" + Str$(Temporary)
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderCVb(Temporary))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderCVb(Temporary), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    Next Temporary
End Sub
