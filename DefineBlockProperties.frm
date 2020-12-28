VERSION 4.00
Begin VB.Form DefineBlockProperties 
   Caption         =   "Automatic Train Control - Define Block Properties"
   ClientHeight    =   9420
   ClientLeft      =   1785
   ClientTop       =   1455
   ClientWidth     =   12660
   Height          =   9825
   Icon            =   "DefineBlockProperties.frx":0000
   Left            =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12660
   Top             =   1110
   Width           =   12780
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   10080
      TabIndex        =   166
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   11400
      TabIndex        =   19
      Top             =   9120
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxTrackIcon 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "DefineBlockProperties.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox TextBoxLength 
      DataField       =   "Length"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      TabIndex        =   11
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox TextBoxTop 
      BackColor       =   &H8000000F&
      DataField       =   "PictureBoxTop"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox TextBoxLeft 
      BackColor       =   &H8000000F&
      DataField       =   "PictureBoxLeft"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox TextBoxFileName 
      BackColor       =   &H8000000F&
      DataField       =   "PictureBoxFileName"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox TextBoxObjectName 
      BackColor       =   &H8000000F&
      DataField       =   "PictureBoxName"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox TextBoxRecordCounter 
      BackColor       =   &H8000000F&
      DataField       =   "PictureBoxName"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Data DatabaseBlockProperties 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\TrackPlanDatabase.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5355
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TrackPlan"
      Top             =   600
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4935
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8705
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Connecting Blocks"
      TabPicture(0)   =   "DefineBlockProperties.frx":0884
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Track Detection Settings"
      TabPicture(1)   =   "DefineBlockProperties.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Signal Lights"
      TabPicture(2)   =   "DefineBlockProperties.frx":08BC
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label9"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label10"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label12"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label13"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label14"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label15"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label16"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Bulb6(7)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Bulb6(6)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Bulb6(5)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Bulb6(4)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Bulb6(3)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Bulb6(2)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Bulb6(1)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Bulb5(8)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Bulb5(7)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Bulb5(6)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Bulb5(5)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Bulb5(4)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Bulb5(3)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Bulb5(2)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Bulb5(1)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Bulb4(8)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Bulb4(7)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Bulb4(6)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Bulb4(5)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Bulb4(4)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Bulb4(3)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Bulb4(2)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Bulb4(1)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Bulb3(8)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Bulb3(7)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Bulb3(6)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Bulb3(5)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Bulb3(4)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Bulb3(3)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Bulb3(2)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Bulb3(1)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Bulb2(8)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Bulb2(7)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Bulb2(6)"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Bulb2(5)"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Bulb2(4)"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Bulb2(3)"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Bulb2(2)"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Bulb2(1)"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Bulb1(8)"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Bulb1(7)"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "TextBoxNodeNumber(1)"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "TextBoxNodeNumber(2)"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "TextBoxNodeNumber(3)"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "TextBoxNodeNumber(4)"
      Tab(2).Control(58).Enabled=   0   'False
      Tab(2).Control(59)=   "TextBoxNodeNumber(5)"
      Tab(2).Control(59).Enabled=   0   'False
      Tab(2).Control(60)=   "TextBoxNodeNumber(6)"
      Tab(2).Control(60).Enabled=   0   'False
      Tab(2).Control(61)=   "TextBoxNodeNumber(7)"
      Tab(2).Control(61).Enabled=   0   'False
      Tab(2).Control(62)=   "TextBoxNodeNumber(8)"
      Tab(2).Control(62).Enabled=   0   'False
      Tab(2).Control(63)=   "TextBoxNodeNumber(9)"
      Tab(2).Control(63).Enabled=   0   'False
      Tab(2).Control(64)=   "TextBoxCardNumber(1)"
      Tab(2).Control(64).Enabled=   0   'False
      Tab(2).Control(65)=   "TextBoxCardNumber(2)"
      Tab(2).Control(65).Enabled=   0   'False
      Tab(2).Control(66)=   "TextBoxCardNumber(3)"
      Tab(2).Control(66).Enabled=   0   'False
      Tab(2).Control(67)=   "TextBoxCardNumber(4)"
      Tab(2).Control(67).Enabled=   0   'False
      Tab(2).Control(68)=   "TextBoxCardNumber(5)"
      Tab(2).Control(68).Enabled=   0   'False
      Tab(2).Control(69)=   "TextBoxCardNumber(6)"
      Tab(2).Control(69).Enabled=   0   'False
      Tab(2).Control(70)=   "TextBoxCardNumber(7)"
      Tab(2).Control(70).Enabled=   0   'False
      Tab(2).Control(71)=   "TextBoxCardNumber(8)"
      Tab(2).Control(71).Enabled=   0   'False
      Tab(2).Control(72)=   "TextBoxCardNumber(9)"
      Tab(2).Control(72).Enabled=   0   'False
      Tab(2).Control(73)=   "Bulb1(1)"
      Tab(2).Control(73).Enabled=   0   'False
      Tab(2).Control(74)=   "Bulb1(2)"
      Tab(2).Control(74).Enabled=   0   'False
      Tab(2).Control(75)=   "Bulb1(3)"
      Tab(2).Control(75).Enabled=   0   'False
      Tab(2).Control(76)=   "Bulb1(4)"
      Tab(2).Control(76).Enabled=   0   'False
      Tab(2).Control(77)=   "Bulb1(5)"
      Tab(2).Control(77).Enabled=   0   'False
      Tab(2).Control(78)=   "Bulb1(6)"
      Tab(2).Control(78).Enabled=   0   'False
      Tab(2).Control(79)=   "Bulb6(8)"
      Tab(2).Control(79).Enabled=   0   'False
      Tab(2).Control(80)=   "Bulb7(1)"
      Tab(2).Control(80).Enabled=   0   'False
      Tab(2).Control(81)=   "Bulb7(2)"
      Tab(2).Control(81).Enabled=   0   'False
      Tab(2).Control(82)=   "Bulb7(3)"
      Tab(2).Control(82).Enabled=   0   'False
      Tab(2).Control(83)=   "Bulb7(4)"
      Tab(2).Control(83).Enabled=   0   'False
      Tab(2).Control(84)=   "Bulb7(5)"
      Tab(2).Control(84).Enabled=   0   'False
      Tab(2).Control(85)=   "Bulb7(6)"
      Tab(2).Control(85).Enabled=   0   'False
      Tab(2).Control(86)=   "Bulb7(7)"
      Tab(2).Control(86).Enabled=   0   'False
      Tab(2).Control(87)=   "Bulb7(8)"
      Tab(2).Control(87).Enabled=   0   'False
      Tab(2).Control(88)=   "Bulb8(1)"
      Tab(2).Control(88).Enabled=   0   'False
      Tab(2).Control(89)=   "Bulb8(2)"
      Tab(2).Control(89).Enabled=   0   'False
      Tab(2).Control(90)=   "Bulb8(3)"
      Tab(2).Control(90).Enabled=   0   'False
      Tab(2).Control(91)=   "Bulb8(4)"
      Tab(2).Control(91).Enabled=   0   'False
      Tab(2).Control(92)=   "Bulb8(5)"
      Tab(2).Control(92).Enabled=   0   'False
      Tab(2).Control(93)=   "Bulb8(6)"
      Tab(2).Control(93).Enabled=   0   'False
      Tab(2).Control(94)=   "Bulb8(7)"
      Tab(2).Control(94).Enabled=   0   'False
      Tab(2).Control(95)=   "Bulb8(8)"
      Tab(2).Control(95).Enabled=   0   'False
      Tab(2).Control(96)=   "Bulb9(1)"
      Tab(2).Control(96).Enabled=   0   'False
      Tab(2).Control(97)=   "Bulb9(3)"
      Tab(2).Control(97).Enabled=   0   'False
      Tab(2).Control(98)=   "Bulb9(4)"
      Tab(2).Control(98).Enabled=   0   'False
      Tab(2).Control(99)=   "Bulb9(5)"
      Tab(2).Control(99).Enabled=   0   'False
      Tab(2).Control(100)=   "Bulb9(6)"
      Tab(2).Control(100).Enabled=   0   'False
      Tab(2).Control(101)=   "Bulb9(7)"
      Tab(2).Control(101).Enabled=   0   'False
      Tab(2).Control(102)=   "Bulb9(8)"
      Tab(2).Control(102).Enabled=   0   'False
      Tab(2).Control(103)=   "CheckBoxMultipleBits(5)"
      Tab(2).Control(103).Enabled=   0   'False
      Tab(2).Control(104)=   "CheckBoxMultipleBits(7)"
      Tab(2).Control(104).Enabled=   0   'False
      Tab(2).Control(105)=   "CheckBoxMultipleBits(8)"
      Tab(2).Control(105).Enabled=   0   'False
      Tab(2).Control(106)=   "CheckBoxMultipleBits(9)"
      Tab(2).Control(106).Enabled=   0   'False
      Tab(2).Control(107)=   "CheckBoxMultipleBits(2)"
      Tab(2).Control(107).Enabled=   0   'False
      Tab(2).Control(108)=   "Bulb9(2)"
      Tab(2).Control(108).Enabled=   0   'False
      Tab(2).Control(109)=   "CheckBoxMultipleBits(6)"
      Tab(2).Control(109).Enabled=   0   'False
      Tab(2).Control(110)=   "CheckBoxMultipleBits(1)"
      Tab(2).Control(110).Enabled=   0   'False
      Tab(2).Control(111)=   "CheckBoxMultipleBits(3)"
      Tab(2).Control(111).Enabled=   0   'False
      Tab(2).Control(112)=   "CheckBoxMultipleBits(4)"
      Tab(2).Control(112).Enabled=   0   'False
      Tab(2).Control(113)=   "ButtonTest(1)"
      Tab(2).Control(113).Enabled=   0   'False
      Tab(2).Control(114)=   "ButtonTest(2)"
      Tab(2).Control(114).Enabled=   0   'False
      Tab(2).Control(115)=   "ButtonTest(3)"
      Tab(2).Control(115).Enabled=   0   'False
      Tab(2).Control(116)=   "ButtonTest(4)"
      Tab(2).Control(116).Enabled=   0   'False
      Tab(2).Control(117)=   "ButtonTest(5)"
      Tab(2).Control(117).Enabled=   0   'False
      Tab(2).Control(118)=   "ButtonTest(6)"
      Tab(2).Control(118).Enabled=   0   'False
      Tab(2).Control(119)=   "ButtonTest(7)"
      Tab(2).Control(119).Enabled=   0   'False
      Tab(2).Control(120)=   "ButtonTest(8)"
      Tab(2).Control(120).Enabled=   0   'False
      Tab(2).Control(121)=   "ButtonTest(9)"
      Tab(2).Control(121).Enabled=   0   'False
      Tab(2).ControlCount=   122
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Ligh On"
         Height          =   255
         Index           =   9
         Left            =   11040
         TabIndex        =   165
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Ligh On"
         Height          =   255
         Index           =   8
         Left            =   11040
         TabIndex        =   164
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Ligh On"
         Height          =   255
         Index           =   7
         Left            =   11040
         TabIndex        =   163
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Light On"
         Height          =   255
         Index           =   6
         Left            =   11040
         TabIndex        =   162
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Light On"
         Height          =   255
         Index           =   5
         Left            =   11040
         TabIndex        =   161
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Light On"
         Height          =   255
         Index           =   4
         Left            =   11040
         TabIndex        =   160
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Light On"
         Height          =   255
         Index           =   3
         Left            =   11040
         TabIndex        =   159
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Ligh On"
         Height          =   255
         Index           =   2
         Left            =   11040
         TabIndex        =   158
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ButtonTest 
         Caption         =   "&Turn Light On"
         Height          =   255
         Index           =   1
         Left            =   11040
         TabIndex        =   157
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T2L1M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   9600
         TabIndex        =   156
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T1L3M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   9600
         TabIndex        =   155
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   154
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T2L3M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   9600
         TabIndex        =   153
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   152
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         Caption         =   " "
         DataField       =   "T1L2M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   151
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T3L3M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   9600
         TabIndex        =   139
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T3L2M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   9600
         TabIndex        =   138
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T3L1M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   9600
         TabIndex        =   137
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox CheckBoxMultipleBits 
         DataField       =   "T2L1M"
         DataSource      =   "DatabaseBlockProperties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   9600
         TabIndex        =   136
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   135
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   134
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   133
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   132
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   131
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   130
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb9 
         DataField       =   "T3L3B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   116
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   115
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   114
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   113
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   112
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   111
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   110
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   109
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb8 
         DataField       =   "T3L2B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   108
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   107
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   106
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   105
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   104
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   103
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   102
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   101
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb7 
         DataField       =   "T3L1B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   100
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   99
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   98
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   97
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   96
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   95
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   94
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   93
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L3C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   92
         Text            =   "n/a"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L2C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   8
         Left            =   3960
         TabIndex        =   91
         Text            =   "n/a"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L1C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   7
         Left            =   3960
         TabIndex        =   90
         Text            =   "n/a"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L3C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   6
         Left            =   3960
         TabIndex        =   89
         Text            =   "n/a"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L2C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   88
         Text            =   "n/a"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L1C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   87
         Text            =   "n/a"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L3C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   86
         Text            =   "n/a"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L2N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   85
         Text            =   "n/a"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextBoxCardNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L1C"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   84
         Text            =   "n/a"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L3N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   83
         Text            =   "n/a"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L2N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   82
         Text            =   "n/a"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T3L1N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   81
         Text            =   "n/a"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L3N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   80
         Text            =   "n/a"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L2N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   79
         Text            =   "n/a"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T2L1N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   78
         Text            =   "n/a"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L3N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   77
         Text            =   "n/a"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L2N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   76
         Text            =   "n/a"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TextBoxNodeNumber 
         Alignment       =   1  'Right Justify
         DataField       =   "T1L1N"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   75
         Text            =   "n/a"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Automatic"
         Height          =   255
         Left            =   -71160
         TabIndex        =   74
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox TextBoxDetectionCmriNode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -72720
         TabIndex        =   73
         Text            =   "n/a"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TextBoxDetectionCmriBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -70440
         TabIndex        =   72
         Text            =   "n/a"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox CheckboxDetection 
         Caption         =   "Primary block detection using"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   71
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox CheckboxDetection 
         Caption         =   "Secondary block detection using"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   70
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox TextBoxDetectionCmriNode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -72720
         TabIndex        =   69
         Text            =   "n/a"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextBoxDetectionCmriBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -70440
         TabIndex        =   68
         Text            =   "n/a"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextBoxWestBoundBlock 
         DataField       =   "WestBound1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   5
         Left            =   -71040
         TabIndex        =   67
         Text            =   "n/a"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextBoxWestBoundBlock 
         DataField       =   "WestBound2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   4
         Left            =   -71040
         TabIndex        =   66
         Text            =   "n/a"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TextBoxWestBoundBlock 
         DataField       =   "WestBound3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   0
         Left            =   -71040
         TabIndex        =   65
         Text            =   "n/a"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TextBoxEastBoundBlock 
         DataField       =   "EastBound1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   5
         Left            =   -71040
         TabIndex        =   64
         Text            =   "n/a"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TextBoxEastBoundBlock 
         DataField       =   "EastBound2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   4
         Left            =   -71040
         TabIndex        =   63
         Text            =   "n/a"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox TextBoxEastBoundBlock 
         DataField       =   "EastBound3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   285
         Index           =   2
         Left            =   -71040
         TabIndex        =   62
         Text            =   "n/a"
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   61
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb1 
         DataField       =   "T1L1B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   60
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   59
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   58
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   57
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   56
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   55
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   54
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   53
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb2 
         DataField       =   "T1L2B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   52
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   51
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   50
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   49
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   48
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   47
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   46
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   45
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb3 
         DataField       =   "T1L3B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   44
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   43
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   42
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   41
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   40
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   39
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   38
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   37
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb4 
         DataField       =   "T2L1B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   36
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   35
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   34
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B3"
         DataSource      =   "DatabaseBlockProperties "
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   33
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   32
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   31
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   30
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2B2L7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   29
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb5 
         DataField       =   "T2L2B8"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   28
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B1"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   27
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B2"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   26
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B3"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   25
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B4"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   24
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B5"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   23
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B6"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   22
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Bulb6 
         DataField       =   "T2L3B7"
         DataSource      =   "DatabaseBlockProperties"
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   21
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Illuminating Items"
         Height          =   255
         Left            =   240
         TabIndex        =   150
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Light Green (botton target)"
         Height          =   255
         Left            =   240
         TabIndex        =   149
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "Light Yellow (bottom target)"
         Height          =   255
         Left            =   240
         TabIndex        =   148
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Light Red (bottom target)"
         Height          =   255
         Left            =   240
         TabIndex        =   147
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Light Green (middle target)"
         Height          =   255
         Left            =   240
         TabIndex        =   146
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Light Yellow (middle target)"
         Height          =   255
         Left            =   240
         TabIndex        =   145
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Light Red (middle target)"
         Height          =   255
         Left            =   240
         TabIndex        =   144
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Light Green (top target)"
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Light Yellow (top target)"
         Height          =   255
         Left            =   240
         TabIndex        =   142
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Light Red (top target)"
         Height          =   255
         Left            =   240
         TabIndex        =   141
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Oscillating Bits"
         Height          =   255
         Left            =   9360
         TabIndex        =   140
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Card Number"
         Height          =   255
         Left            =   3960
         TabIndex        =   129
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Node Number"
         Height          =   255
         Left            =   2280
         TabIndex        =   128
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LabelDetectionCmriNode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Node Number"
         Height          =   315
         Index           =   3
         Left            =   -73800
         TabIndex        =   127
         Top             =   960
         Width           =   990
      End
      Begin VB.Label LabelDetectionCmriBit 
         Alignment       =   1  'Right Justify
         Caption         =   "Bit for Detection"
         Height          =   255
         Index           =   3
         Left            =   -71760
         TabIndex        =   126
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LabelDetectionCmriNode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Node Number"
         Height          =   315
         Index           =   2
         Left            =   -73800
         TabIndex        =   125
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label LabelDetectionCmriBit 
         Alignment       =   1  'Right Justify
         Caption         =   "Bit for Detection"
         Height          =   255
         Index           =   2
         Left            =   -71760
         TabIndex        =   124
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label LabelWestBoundBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "West  bound link to next block (primary)"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   123
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label LabelWestBoundBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "West bound link to next block (secondary)"
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   122
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label LabelWestBoundBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "West bound link to next block (secondary)"
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   121
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label LabelEastBoundBlocks 
         Alignment       =   1  'Right Justify
         Caption         =   "East bound link to next block (primary)"
         Height          =   255
         Index           =   5
         Left            =   -74280
         TabIndex        =   120
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label LabelEastBoundBlocks 
         Alignment       =   1  'Right Justify
         Caption         =   "East bound link to next block (secondary)"
         Height          =   255
         Index           =   4
         Left            =   -74280
         TabIndex        =   119
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label LabelEastBoundBlocks 
         Alignment       =   1  'Right Justify
         Caption         =   "East bound link to next block (primary)"
         Height          =   255
         Index           =   3
         Left            =   -74280
         TabIndex        =   118
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Bit Number     1     2      3      4      5      6      7      8"
         Height          =   255
         Left            =   5280
         TabIndex        =   117
         Top             =   600
         Width           =   3735
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7440
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Label LabelTrackIconProperties 
      Caption         =   "Track Icon Properties"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   2895
   End
   Begin IniconLib.Init Ini 
      Left            =   7440
      Top             =   240
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7440
      Top             =   840
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   $"DefineBlockProperties.frx":08D8
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   960
      TabIndex        =   14
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label LabelName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name of Block"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Length of Block"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LabelTop 
      Alignment       =   1  'Right Justify
      Caption         =   "Top Position on Track Map"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label LabelLeft 
      Alignment       =   1  'Right Justify
      Caption         =   "Left Position on Track Map"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LabelFileName 
      Alignment       =   1  'Right Justify
      Caption         =   "File Name"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label LabelObjectName 
      Alignment       =   1  'Right Justify
      Caption         =   "Object Name"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label LabelRecordCounter 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Counter"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "DefineBlockProperties"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Bulb1_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb1(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb1(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(1).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb1(temporaryVariable).Value = vbUnchecked Then
                Let Bulb1(temporaryVariable).Enabled = False
            ElseIf Bulb1(temporaryVariable).Value = vbChecked Then
                Let Bulb1(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(1).Value = vbUnchecked
        If Bulb1(1).Value = vbChecked Then
            Let Bulb1(2).Enabled = True
            Let Bulb1(3).Enabled = False
            Let Bulb1(4).Enabled = False
            Let Bulb1(5).Enabled = False
            Let Bulb1(6).Enabled = False
            Let Bulb1(7).Enabled = False
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(2).Value = vbChecked Then
            Let Bulb1(1).Enabled = True
            Let Bulb1(3).Enabled = True
            Let Bulb1(4).Enabled = False
            Let Bulb1(5).Enabled = False
            Let Bulb1(6).Enabled = False
            Let Bulb1(7).Enabled = False
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(3).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = True
            Let Bulb1(4).Enabled = True
            Let Bulb1(5).Enabled = False
            Let Bulb1(6).Enabled = False
            Let Bulb1(7).Enabled = False
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(4).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = False
            Let Bulb1(3).Enabled = True
            Let Bulb1(5).Enabled = True
            Let Bulb1(6).Enabled = False
            Let Bulb1(7).Enabled = False
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(5).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = False
            Let Bulb1(3).Enabled = False
            Let Bulb1(4).Enabled = True
            Let Bulb1(6).Enabled = True
            Let Bulb1(7).Enabled = False
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(6).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = False
            Let Bulb1(3).Enabled = False
            Let Bulb1(4).Enabled = False
            Let Bulb1(5).Enabled = True
            Let Bulb1(7).Enabled = True
            Let Bulb1(8).Enabled = False
        ElseIf Bulb1(7).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = False
            Let Bulb1(3).Enabled = False
            Let Bulb1(4).Enabled = False
            Let Bulb1(5).Enabled = False
            Let Bulb1(6).Enabled = True
            Let Bulb1(8).Enabled = True
        ElseIf Bulb1(8).Value = vbChecked Then
            Let Bulb1(1).Enabled = False
            Let Bulb1(2).Enabled = False
            Let Bulb1(3).Enabled = False
            Let Bulb1(4).Enabled = False
            Let Bulb1(5).Enabled = False
            Let Bulb1(6).Enabled = False
            Let Bulb1(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(1).Value = vbUnchecked
        Let Bulb1(1).Enabled = True
        Let Bulb1(2).Enabled = True
        Let Bulb1(3).Enabled = True
        Let Bulb1(4).Enabled = True
        Let Bulb1(5).Enabled = True
        Let Bulb1(6).Enabled = True
        Let Bulb1(7).Enabled = True
        Let Bulb1(8).Enabled = True
    End If
End Sub


Private Sub Bulb2_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb2(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb2(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(2).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb2(temporaryVariable).Value = vbUnchecked Then
                Let Bulb2(temporaryVariable).Enabled = False
            ElseIf Bulb2(temporaryVariable).Value = vbChecked Then
                Let Bulb2(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(2).Value = vbUnchecked
        If Bulb2(1).Value = vbChecked Then
            Let Bulb2(2).Enabled = True
            Let Bulb2(3).Enabled = False
            Let Bulb2(4).Enabled = False
            Let Bulb2(5).Enabled = False
            Let Bulb2(6).Enabled = False
            Let Bulb2(7).Enabled = False
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(2).Value = vbChecked Then
            Let Bulb2(1).Enabled = True
            Let Bulb2(3).Enabled = True
            Let Bulb2(4).Enabled = False
            Let Bulb2(5).Enabled = False
            Let Bulb2(6).Enabled = False
            Let Bulb2(7).Enabled = False
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(3).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = True
            Let Bulb2(4).Enabled = True
            Let Bulb2(5).Enabled = False
            Let Bulb2(6).Enabled = False
            Let Bulb2(7).Enabled = False
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(4).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = False
            Let Bulb2(3).Enabled = True
            Let Bulb2(5).Enabled = True
            Let Bulb2(6).Enabled = False
            Let Bulb2(7).Enabled = False
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(5).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = False
            Let Bulb2(3).Enabled = False
            Let Bulb2(4).Enabled = True
            Let Bulb2(6).Enabled = True
            Let Bulb2(7).Enabled = False
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(6).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = False
            Let Bulb2(3).Enabled = False
            Let Bulb2(4).Enabled = False
            Let Bulb2(5).Enabled = True
            Let Bulb2(7).Enabled = True
            Let Bulb2(8).Enabled = False
        ElseIf Bulb2(7).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = False
            Let Bulb2(3).Enabled = False
            Let Bulb2(4).Enabled = False
            Let Bulb2(5).Enabled = False
            Let Bulb2(6).Enabled = True
            Let Bulb2(8).Enabled = True
        ElseIf Bulb2(8).Value = vbChecked Then
            Let Bulb2(1).Enabled = False
            Let Bulb2(2).Enabled = False
            Let Bulb2(3).Enabled = False
            Let Bulb2(4).Enabled = False
            Let Bulb2(5).Enabled = False
            Let Bulb2(6).Enabled = False
            Let Bulb2(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(2).Value = vbUnchecked
        Let Bulb2(1).Enabled = True
        Let Bulb2(2).Enabled = True
        Let Bulb2(3).Enabled = True
        Let Bulb2(4).Enabled = True
        Let Bulb2(5).Enabled = True
        Let Bulb2(6).Enabled = True
        Let Bulb2(7).Enabled = True
        Let Bulb2(8).Enabled = True
    End If
End Sub


Private Sub Bulb3_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb3(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb3(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(3).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb3(temporaryVariable).Value = vbUnchecked Then
                Let Bulb3(temporaryVariable).Enabled = False
            ElseIf Bulb3(temporaryVariable).Value = vbChecked Then
                Let Bulb3(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(3).Value = vbUnchecked
        If Bulb3(1).Value = vbChecked Then
            Let Bulb3(2).Enabled = True
            Let Bulb3(3).Enabled = False
            Let Bulb3(4).Enabled = False
            Let Bulb3(5).Enabled = False
            Let Bulb3(6).Enabled = False
            Let Bulb3(7).Enabled = False
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(2).Value = vbChecked Then
            Let Bulb3(1).Enabled = True
            Let Bulb3(3).Enabled = True
            Let Bulb3(4).Enabled = False
            Let Bulb3(5).Enabled = False
            Let Bulb3(6).Enabled = False
            Let Bulb3(7).Enabled = False
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(3).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = True
            Let Bulb3(4).Enabled = True
            Let Bulb3(5).Enabled = False
            Let Bulb3(6).Enabled = False
            Let Bulb3(7).Enabled = False
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(4).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = False
            Let Bulb3(3).Enabled = True
            Let Bulb3(5).Enabled = True
            Let Bulb3(6).Enabled = False
            Let Bulb3(7).Enabled = False
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(5).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = False
            Let Bulb3(3).Enabled = False
            Let Bulb3(4).Enabled = True
            Let Bulb3(6).Enabled = True
            Let Bulb3(7).Enabled = False
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(6).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = False
            Let Bulb3(3).Enabled = False
            Let Bulb3(4).Enabled = False
            Let Bulb3(5).Enabled = True
            Let Bulb3(7).Enabled = True
            Let Bulb3(8).Enabled = False
        ElseIf Bulb3(7).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = False
            Let Bulb3(3).Enabled = False
            Let Bulb3(4).Enabled = False
            Let Bulb3(5).Enabled = False
            Let Bulb3(6).Enabled = True
            Let Bulb3(8).Enabled = True
        ElseIf Bulb3(8).Value = vbChecked Then
            Let Bulb3(1).Enabled = False
            Let Bulb3(2).Enabled = False
            Let Bulb3(3).Enabled = False
            Let Bulb3(4).Enabled = False
            Let Bulb3(5).Enabled = False
            Let Bulb3(6).Enabled = False
            Let Bulb3(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(3).Value = vbUnchecked
        Let Bulb3(1).Enabled = True
        Let Bulb3(2).Enabled = True
        Let Bulb3(3).Enabled = True
        Let Bulb3(4).Enabled = True
        Let Bulb3(5).Enabled = True
        Let Bulb3(6).Enabled = True
        Let Bulb3(7).Enabled = True
        Let Bulb3(8).Enabled = True
    End If
   
End Sub


Private Sub Bulb4_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb4(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb4(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(4).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb4(temporaryVariable).Value = vbUnchecked Then
                Let Bulb4(temporaryVariable).Enabled = False
            ElseIf Bulb4(temporaryVariable).Value = vbChecked Then
                Let Bulb4(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(4).Value = vbUnchecked
        If Bulb4(1).Value = vbChecked Then
            Let Bulb4(2).Enabled = True
            Let Bulb4(3).Enabled = False
            Let Bulb4(4).Enabled = False
            Let Bulb4(5).Enabled = False
            Let Bulb4(6).Enabled = False
            Let Bulb4(7).Enabled = False
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(2).Value = vbChecked Then
            Let Bulb4(1).Enabled = True
            Let Bulb4(3).Enabled = True
            Let Bulb4(4).Enabled = False
            Let Bulb4(5).Enabled = False
            Let Bulb4(6).Enabled = False
            Let Bulb4(7).Enabled = False
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(3).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = True
            Let Bulb4(4).Enabled = True
            Let Bulb4(5).Enabled = False
            Let Bulb4(6).Enabled = False
            Let Bulb4(7).Enabled = False
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(4).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = False
            Let Bulb4(3).Enabled = True
            Let Bulb4(5).Enabled = True
            Let Bulb4(6).Enabled = False
            Let Bulb4(7).Enabled = False
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(5).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = False
            Let Bulb4(3).Enabled = False
            Let Bulb4(4).Enabled = True
            Let Bulb4(6).Enabled = True
            Let Bulb4(7).Enabled = False
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(6).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = False
            Let Bulb4(3).Enabled = False
            Let Bulb4(4).Enabled = False
            Let Bulb4(5).Enabled = True
            Let Bulb4(7).Enabled = True
            Let Bulb4(8).Enabled = False
        ElseIf Bulb4(7).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = False
            Let Bulb4(3).Enabled = False
            Let Bulb4(4).Enabled = False
            Let Bulb4(5).Enabled = False
            Let Bulb4(6).Enabled = True
            Let Bulb4(8).Enabled = True
        ElseIf Bulb4(8).Value = vbChecked Then
            Let Bulb4(1).Enabled = False
            Let Bulb4(2).Enabled = False
            Let Bulb4(3).Enabled = False
            Let Bulb4(4).Enabled = False
            Let Bulb4(5).Enabled = False
            Let Bulb4(6).Enabled = False
            Let Bulb4(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(4).Value = vbUnchecked
        Let Bulb4(1).Enabled = True
        Let Bulb4(2).Enabled = True
        Let Bulb4(3).Enabled = True
        Let Bulb4(4).Enabled = True
        Let Bulb4(5).Enabled = True
        Let Bulb4(6).Enabled = True
        Let Bulb4(7).Enabled = True
        Let Bulb4(8).Enabled = True
    End If
    
End Sub


Private Sub Bulb5_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb5(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb5(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(5).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb5(temporaryVariable).Value = vbUnchecked Then
                Let Bulb5(temporaryVariable).Enabled = False
            ElseIf Bulb5(temporaryVariable).Value = vbChecked Then
                Let Bulb5(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(5).Value = vbUnchecked
        If Bulb5(1).Value = vbChecked Then
            Let Bulb5(2).Enabled = True
            Let Bulb5(3).Enabled = False
            Let Bulb5(4).Enabled = False
            Let Bulb5(5).Enabled = False
            Let Bulb5(6).Enabled = False
            Let Bulb5(7).Enabled = False
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(2).Value = vbChecked Then
            Let Bulb5(1).Enabled = True
            Let Bulb5(3).Enabled = True
            Let Bulb5(4).Enabled = False
            Let Bulb5(5).Enabled = False
            Let Bulb5(6).Enabled = False
            Let Bulb5(7).Enabled = False
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(3).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = True
            Let Bulb5(4).Enabled = True
            Let Bulb5(5).Enabled = False
            Let Bulb5(6).Enabled = False
            Let Bulb5(7).Enabled = False
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(4).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = False
            Let Bulb5(3).Enabled = True
            Let Bulb5(5).Enabled = True
            Let Bulb5(6).Enabled = False
            Let Bulb5(7).Enabled = False
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(5).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = False
            Let Bulb5(3).Enabled = False
            Let Bulb5(4).Enabled = True
            Let Bulb5(6).Enabled = True
            Let Bulb5(7).Enabled = False
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(6).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = False
            Let Bulb5(3).Enabled = False
            Let Bulb5(4).Enabled = False
            Let Bulb5(5).Enabled = True
            Let Bulb5(7).Enabled = True
            Let Bulb5(8).Enabled = False
        ElseIf Bulb5(7).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = False
            Let Bulb5(3).Enabled = False
            Let Bulb5(4).Enabled = False
            Let Bulb5(5).Enabled = False
            Let Bulb5(6).Enabled = True
            Let Bulb5(8).Enabled = True
        ElseIf Bulb5(8).Value = vbChecked Then
            Let Bulb5(1).Enabled = False
            Let Bulb5(2).Enabled = False
            Let Bulb5(3).Enabled = False
            Let Bulb5(4).Enabled = False
            Let Bulb5(5).Enabled = False
            Let Bulb5(6).Enabled = False
            Let Bulb5(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(5).Value = vbUnchecked
        Let Bulb5(1).Enabled = True
        Let Bulb5(2).Enabled = True
        Let Bulb5(3).Enabled = True
        Let Bulb5(4).Enabled = True
        Let Bulb5(5).Enabled = True
        Let Bulb5(6).Enabled = True
        Let Bulb5(7).Enabled = True
        Let Bulb5(8).Enabled = True
    End If
    
End Sub


Private Sub Bulb6_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb6(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb6(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(6).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb6(temporaryVariable).Value = vbUnchecked Then
                Let Bulb6(temporaryVariable).Enabled = False
            ElseIf Bulb6(temporaryVariable).Value = vbChecked Then
                Let Bulb6(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(6).Value = vbUnchecked
        If Bulb6(1).Value = vbChecked Then
            Let Bulb6(2).Enabled = True
            Let Bulb6(3).Enabled = False
            Let Bulb6(4).Enabled = False
            Let Bulb6(5).Enabled = False
            Let Bulb6(6).Enabled = False
            Let Bulb6(7).Enabled = False
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(2).Value = vbChecked Then
            Let Bulb6(1).Enabled = True
            Let Bulb6(3).Enabled = True
            Let Bulb6(4).Enabled = False
            Let Bulb6(5).Enabled = False
            Let Bulb6(6).Enabled = False
            Let Bulb6(7).Enabled = False
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(3).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = True
            Let Bulb6(4).Enabled = True
            Let Bulb6(5).Enabled = False
            Let Bulb6(6).Enabled = False
            Let Bulb6(7).Enabled = False
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(4).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = False
            Let Bulb6(3).Enabled = True
            Let Bulb6(5).Enabled = True
            Let Bulb6(6).Enabled = False
            Let Bulb6(7).Enabled = False
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(5).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = False
            Let Bulb6(3).Enabled = False
            Let Bulb6(4).Enabled = True
            Let Bulb6(6).Enabled = True
            Let Bulb6(7).Enabled = False
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(6).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = False
            Let Bulb6(3).Enabled = False
            Let Bulb6(4).Enabled = False
            Let Bulb6(5).Enabled = True
            Let Bulb6(7).Enabled = True
            Let Bulb6(8).Enabled = False
        ElseIf Bulb6(7).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = False
            Let Bulb6(3).Enabled = False
            Let Bulb6(4).Enabled = False
            Let Bulb6(5).Enabled = False
            Let Bulb6(6).Enabled = True
            Let Bulb6(8).Enabled = True
        ElseIf Bulb6(8).Value = vbChecked Then
            Let Bulb6(1).Enabled = False
            Let Bulb6(2).Enabled = False
            Let Bulb6(3).Enabled = False
            Let Bulb6(4).Enabled = False
            Let Bulb6(5).Enabled = False
            Let Bulb6(6).Enabled = False
            Let Bulb6(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(6).Value = vbUnchecked
        Let Bulb6(1).Enabled = True
        Let Bulb6(2).Enabled = True
        Let Bulb6(3).Enabled = True
        Let Bulb6(4).Enabled = True
        Let Bulb6(5).Enabled = True
        Let Bulb6(6).Enabled = True
        Let Bulb6(7).Enabled = True
        Let Bulb6(8).Enabled = True
    End If
    
End Sub


Private Sub Bulb7_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb7(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb7(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(7).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb7(temporaryVariable).Value = vbUnchecked Then
                Let Bulb7(temporaryVariable).Enabled = False
            ElseIf Bulb7(temporaryVariable).Value = vbChecked Then
                Let Bulb7(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(7).Value = vbUnchecked
        If Bulb7(1).Value = vbChecked Then
            Let Bulb7(2).Enabled = True
            Let Bulb7(3).Enabled = False
            Let Bulb7(4).Enabled = False
            Let Bulb7(5).Enabled = False
            Let Bulb7(6).Enabled = False
            Let Bulb7(7).Enabled = False
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(2).Value = vbChecked Then
            Let Bulb7(1).Enabled = True
            Let Bulb7(3).Enabled = True
            Let Bulb7(4).Enabled = False
            Let Bulb7(5).Enabled = False
            Let Bulb7(6).Enabled = False
            Let Bulb7(7).Enabled = False
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(3).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = True
            Let Bulb7(4).Enabled = True
            Let Bulb7(5).Enabled = False
            Let Bulb7(6).Enabled = False
            Let Bulb7(7).Enabled = False
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(4).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = False
            Let Bulb7(3).Enabled = True
            Let Bulb7(5).Enabled = True
            Let Bulb7(6).Enabled = False
            Let Bulb7(7).Enabled = False
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(5).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = False
            Let Bulb7(3).Enabled = False
            Let Bulb7(4).Enabled = True
            Let Bulb7(6).Enabled = True
            Let Bulb7(7).Enabled = False
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(6).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = False
            Let Bulb7(3).Enabled = False
            Let Bulb7(4).Enabled = False
            Let Bulb7(5).Enabled = True
            Let Bulb7(7).Enabled = True
            Let Bulb7(8).Enabled = False
        ElseIf Bulb7(7).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = False
            Let Bulb7(3).Enabled = False
            Let Bulb7(4).Enabled = False
            Let Bulb7(5).Enabled = False
            Let Bulb7(6).Enabled = True
            Let Bulb7(8).Enabled = True
        ElseIf Bulb7(8).Value = vbChecked Then
            Let Bulb7(1).Enabled = False
            Let Bulb7(2).Enabled = False
            Let Bulb7(3).Enabled = False
            Let Bulb7(4).Enabled = False
            Let Bulb7(5).Enabled = False
            Let Bulb7(6).Enabled = False
            Let Bulb7(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(7).Value = vbUnchecked
        Let Bulb7(1).Enabled = True
        Let Bulb7(2).Enabled = True
        Let Bulb7(3).Enabled = True
        Let Bulb7(4).Enabled = True
        Let Bulb7(5).Enabled = True
        Let Bulb7(6).Enabled = True
        Let Bulb7(7).Enabled = True
        Let Bulb7(8).Enabled = True
    End If
    
End Sub

Private Sub Bulb8_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb8(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb8(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(8).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb8(temporaryVariable).Value = vbUnchecked Then
                Let Bulb8(temporaryVariable).Enabled = False
            ElseIf Bulb8(temporaryVariable).Value = vbChecked Then
                Let Bulb8(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(8).Value = vbUnchecked
        If Bulb8(1).Value = vbChecked Then
            Let Bulb8(2).Enabled = True
            Let Bulb8(3).Enabled = False
            Let Bulb8(4).Enabled = False
            Let Bulb8(5).Enabled = False
            Let Bulb8(6).Enabled = False
            Let Bulb8(7).Enabled = False
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(2).Value = vbChecked Then
            Let Bulb8(1).Enabled = True
            Let Bulb8(3).Enabled = True
            Let Bulb8(4).Enabled = False
            Let Bulb8(5).Enabled = False
            Let Bulb8(6).Enabled = False
            Let Bulb8(7).Enabled = False
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(3).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = True
            Let Bulb8(4).Enabled = True
            Let Bulb8(5).Enabled = False
            Let Bulb8(6).Enabled = False
            Let Bulb8(7).Enabled = False
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(4).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = False
            Let Bulb8(3).Enabled = True
            Let Bulb8(5).Enabled = True
            Let Bulb8(6).Enabled = False
            Let Bulb8(7).Enabled = False
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(5).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = False
            Let Bulb8(3).Enabled = False
            Let Bulb8(4).Enabled = True
            Let Bulb8(6).Enabled = True
            Let Bulb8(7).Enabled = False
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(6).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = False
            Let Bulb8(3).Enabled = False
            Let Bulb8(4).Enabled = False
            Let Bulb8(5).Enabled = True
            Let Bulb8(7).Enabled = True
            Let Bulb8(8).Enabled = False
        ElseIf Bulb8(7).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = False
            Let Bulb8(3).Enabled = False
            Let Bulb8(4).Enabled = False
            Let Bulb8(5).Enabled = False
            Let Bulb8(6).Enabled = True
            Let Bulb8(8).Enabled = True
        ElseIf Bulb8(8).Value = vbChecked Then
            Let Bulb8(1).Enabled = False
            Let Bulb8(2).Enabled = False
            Let Bulb8(3).Enabled = False
            Let Bulb8(4).Enabled = False
            Let Bulb8(5).Enabled = False
            Let Bulb8(6).Enabled = False
            Let Bulb8(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(8).Value = vbUnchecked
        Let Bulb8(1).Enabled = True
        Let Bulb8(2).Enabled = True
        Let Bulb8(3).Enabled = True
        Let Bulb8(4).Enabled = True
        Let Bulb8(5).Enabled = True
        Let Bulb8(6).Enabled = True
        Let Bulb8(7).Enabled = True
        Let Bulb8(8).Enabled = True
    End If

End Sub


Private Sub Bulb9_Click(Index As Integer)

    Dim temporaryMultiple As Integer
    Dim temporaryInteger As Integer
    Dim temporaryVariable As Integer
    
    Let temporaryMultiple = 0
    
    For temporaryInteger = 1 To 8
        If Bulb9(temporaryInteger) = vbChecked Then
            Let temporaryMultiple = temporaryMultiple + 1
        End If
    Next temporaryInteger
    
    If temporaryMultiple > 2 Then
        MsgBox "Error in assingning bits to specific outputs", vbExclamation + vbOKOnly, "Data Entry Error"
        For temporaryInteger = 1 To 8
            Let Bulb9(temporaryInteger).Value = vbUnchecked
        Next temporaryInteger
    ElseIf temporaryMultiple = 2 Then
        Let checkboxmultiplebits(9).Value = vbChecked
        For temporaryVariable = 1 To 8
            If Bulb9(temporaryVariable).Value = vbUnchecked Then
                Let Bulb9(temporaryVariable).Enabled = False
            ElseIf Bulb9(temporaryVariable).Value = vbChecked Then
                Let Bulb9(temporaryVariable).Enabled = True
            End If
        Next temporaryVariable
    ElseIf temporaryMultiple = 1 Then
        Let checkboxmultiplebits(9).Value = vbUnchecked
        If Bulb9(1).Value = vbChecked Then
            Let Bulb9(2).Enabled = True
            Let Bulb9(3).Enabled = False
            Let Bulb9(4).Enabled = False
            Let Bulb9(5).Enabled = False
            Let Bulb9(6).Enabled = False
            Let Bulb9(7).Enabled = False
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(2).Value = vbChecked Then
            Let Bulb9(1).Enabled = True
            Let Bulb9(3).Enabled = True
            Let Bulb9(4).Enabled = False
            Let Bulb9(5).Enabled = False
            Let Bulb9(6).Enabled = False
            Let Bulb9(7).Enabled = False
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(3).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = True
            Let Bulb9(4).Enabled = True
            Let Bulb9(5).Enabled = False
            Let Bulb9(6).Enabled = False
            Let Bulb9(7).Enabled = False
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(4).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = False
            Let Bulb9(3).Enabled = True
            Let Bulb9(5).Enabled = True
            Let Bulb9(6).Enabled = False
            Let Bulb9(7).Enabled = False
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(5).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = False
            Let Bulb9(3).Enabled = False
            Let Bulb9(4).Enabled = True
            Let Bulb9(6).Enabled = True
            Let Bulb9(7).Enabled = False
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(6).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = False
            Let Bulb9(3).Enabled = False
            Let Bulb9(4).Enabled = False
            Let Bulb9(5).Enabled = True
            Let Bulb9(7).Enabled = True
            Let Bulb9(8).Enabled = False
        ElseIf Bulb9(7).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = False
            Let Bulb9(3).Enabled = False
            Let Bulb9(4).Enabled = False
            Let Bulb9(5).Enabled = False
            Let Bulb9(6).Enabled = True
            Let Bulb9(8).Enabled = True
        ElseIf Bulb9(8).Value = vbChecked Then
            Let Bulb9(1).Enabled = False
            Let Bulb9(2).Enabled = False
            Let Bulb9(3).Enabled = False
            Let Bulb9(4).Enabled = False
            Let Bulb9(5).Enabled = False
            Let Bulb9(6).Enabled = False
            Let Bulb9(7).Enabled = True
        End If
    ElseIf temporaryMultiple = 0 Then
        Let checkboxmultiplebits(9).Value = vbUnchecked
        Let Bulb9(1).Enabled = True
        Let Bulb9(2).Enabled = True
        Let Bulb9(3).Enabled = True
        Let Bulb9(4).Enabled = True
        Let Bulb9(5).Enabled = True
        Let Bulb9(6).Enabled = True
        Let Bulb9(7).Enabled = True
        Let Bulb9(8).Enabled = True
    End If
    
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
            If TemporaryScreen = "Define Block Properties Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Block Properties Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
             '   FunScreen.Show vbModeless
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
                Let Ini.Value = "Define Block Properties Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Define Block Properties Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub



Private Sub ButtonPrint_Click()

    DefineBlockProperties.PrintForm

End Sub

Private Sub ButtonTest_Click(Index As Integer)

    If ButtonTest(Index).Caption = "&Turn Light On" Then
        Let ButtonTest(Index).Caption = "&Turn Light Off"
    Else
        Let ButtonTest(Index).Caption = "&Turn Light On"
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
        If TemporaryScreen = "Define Block Properties Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Define Block Properties Screen"
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
            Let Ini.Value = "Define Block Properties Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Define Block Properties Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Define Block Properties Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Define Block Properties Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(DefineBlockProperties.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(DefineBlockProperties.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(DefineBlockProperties.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(DefineBlockProperties.Height)

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
                Let Ini.Value = "Define Block Properties Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
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
            Let Ini.Value = "Define Block Properties Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    DefineBlockProperties.Hide
    'unload defineblockproperties

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
    Let Ini.Application = "Define Block Properties"
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
        DefineBlockProperties.Left = (Screen.Width - Width) / 2
        DefineBlockProperties.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + DefineBlockProperties.Width > Screen.Width Then
            Let DefineBlockProperties.Left = Screen.Width - DefineBlockProperties.Width
        Else
            Let DefineBlockProperties.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + DefineBlockProperties.Height > Screen.Height Then
            Let DefineBlockProperties.Top = Screen.Height - DefineBlockProperties.Height
        Else
            Let DefineBlockProperties.Top = Val(TemporaryValueTop)
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
        
        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
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
    DatabaseBlockProperties.DatabaseName = App.Path$ + "\Databases\TrackPlanDatabase.mdb"
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub




Private Sub Form_Resize()

    If DefineBlockProperties.WindowState = vbMinimized Then
    
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
        
    ElseIf DefineBlockProperties.WindowState = vbNormal Then
    
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

Private Sub TextBoxFileName_Change()

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Update the Picture in the Track Icon
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let PictureBoxTrackIcon.Picture = LoadPicture(App.Path$ & "\Graphics\" & TextBoxFileName.Text)

End Sub


Private Sub TextBoxNodeNumber_Change(Index As Integer)

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Check Value
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TextBoxNodeNumber(Index).Text) > 127 Then
        MsgBox "The maximum number of nodes you can have is one hundred and" & vbCrLf & "twenty-seven. Please enter a new number.", vbExclamation + vbOKOnly, "Error Entering Node Number for Signal"
        Let TextBoxNodeNumber(Index).Text = 127
    ElseIf Val(TextBoxNodeNumber(Index).Text) < 0 Then
        MsgBox "The minimum node address you can have is zero. Please enter" & vbCrLf & "a new number."
        Let TextBoxNodeNumber(Index).Text = 0
    End If
    
End Sub


