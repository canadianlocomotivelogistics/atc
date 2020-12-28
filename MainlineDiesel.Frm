VERSION 4.00
Begin VB.Form MainlineDiesel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Diesel Database"
   ClientHeight    =   10860
   ClientLeft      =   3465
   ClientTop       =   1680
   ClientWidth     =   6705
   Height          =   11265
   Icon            =   "MainlineDiesel.frx":0000
   Left            =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   6705
   Top             =   1335
   Width           =   6825
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   5400
      TabIndex        =   59
      Top             =   8820
      Width           =   1215
   End
   Begin VB.Timer TimerInitialize 
      Interval        =   1
      Left            =   7200
      Top             =   1800
   End
   Begin VB.TextBox TextBoxLastTimeLocomotiveViewed 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "LastTimeLocomotiveViewed"
      DataSource      =   "LocomotiveNumber"
      Height          =   225
      Left            =   3960
      TabIndex        =   53
      Text            =   "UnKnown"
      Top             =   9840
      Width           =   2655
   End
   Begin VB.TextBox TextBoxLastTimeDecoderViewed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "LastTimeDeocderViewed"
      DataSource      =   "LocomotiveNumber"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4200
      TabIndex        =   52
      Text            =   "Unknown"
      Top             =   10080
      Width           =   2415
   End
   Begin VB.TextBox TextBoxLastTimeDecoderRead 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "LastTimeDecoderRead"
      DataSource      =   "LocomotiveNumber"
      Height          =   225
      Left            =   3960
      TabIndex        =   51
      Text            =   "Unknown"
      Top             =   10320
      Width           =   2535
   End
   Begin VB.TextBox TextBoxLastTimeDecoderWritten 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "LastTimeDecoderWritten"
      DataSource      =   "LocomotiveNumber"
      Height          =   225
      Left            =   4320
      TabIndex        =   50
      Text            =   "Unknown"
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Data ReportingMarksDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ReportingMarks"
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "List"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   8220
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   32
      Top             =   8520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   60
      Picture         =   "MainlineDiesel.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   210
      Width           =   480
   End
   Begin VB.CheckBox LocomotiveDatabaseActiveSearch 
      Caption         =   "Locomotive Active List Search"
      DataField       =   "LocomotiveActiveList"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data LocomotiveDatabaseSearch 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBaseLongV030002"
      Top             =   1320
      Width           =   1140
   End
   Begin VB.TextBox LocomotiveDatabaseNumberSearch 
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   285
      Left            =   7200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data LocomotiveNumber 
      Caption         =   "Locomotive Number"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDatabaseLongV030002"
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox LocomotiveImageFilename 
      DataField       =   "LocomotiveImage"
      DataSource      =   "LocomotiveNumber"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton ButtonChangePicture 
      Caption         =   "Change Picture"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton ButtonDecoderProg 
      Caption         =   "&Decoder Prog"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame LocomotiveType 
      Caption         =   "Locomotive Type"
      Height          =   5070
      Left            =   120
      TabIndex        =   34
      Top             =   3780
      Width           =   3015
      Begin VB.ComboBox LocomotiveModelManufacturer 
         DataField       =   "LocomotiveModelManufacturer"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0884
         Left            =   175
         List            =   "MainlineDiesel.frx":089A
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "Locomotive Model Manufacturer"
         Top             =   1320
         Width           =   2670
      End
      Begin VB.TextBox LocomotiveFacts 
         DataField       =   "LocomotiveFacts"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   3075
         Left            =   175
         MaxLength       =   65535
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "MainlineDiesel.frx":0924
         Top             =   1800
         Width           =   2670
      End
      Begin VB.ComboBox LocomotiveModel 
         DataField       =   "LocomotiveModel"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0937
         Left            =   175
         List            =   "MainlineDiesel.frx":0939
         TabIndex        =   9
         Text            =   "Locomotive Model"
         Top             =   840
         Width           =   2670
      End
      Begin VB.ComboBox LocomotiveManufacturer 
         DataField       =   "LocomotiveManufacturer"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":093B
         Left            =   175
         List            =   "MainlineDiesel.frx":0957
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Locomotive Manufacturer"
         Top             =   360
         Width           =   2670
      End
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Frame LocoID 
      Caption         =   "Locomotive Identification"
      Height          =   2130
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
      Begin VB.CheckBox CheckAddressShort 
         Caption         =   "Sort Address"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox LocomotiveNumberGO 
         Height          =   315
         Left            =   175
         TabIndex        =   2
         Text            =   "0000"
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton ButtonGo 
         Caption         =   "Go"
         Height          =   285
         Left            =   2085
         TabIndex        =   7
         Top             =   1725
         Width           =   765
      End
      Begin VB.CheckBox LocomotiveDecoderEquiped 
         Caption         =   "Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   255
         Left            =   175
         TabIndex        =   6
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CheckBox LocomotiveActiveList 
         Caption         =   "Locomotive Active in List"
         DataField       =   "LocomotiveActiveList"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   255
         Left            =   175
         TabIndex        =   5
         Top             =   1395
         Width           =   2175
      End
      Begin VB.ComboBox LocomotivePrototype 
         DataField       =   "LocomotivePrototype"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0A06
         Left            =   175
         List            =   "MainlineDiesel.frx":0A08
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Proto Type Owner"
         Top             =   930
         Width           =   2670
      End
      Begin VB.Line Line2 
         X1              =   2820
         X2              =   180
         Y1              =   780
         Y2              =   780
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   3360
      TabIndex        =   16
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Prototype"
      TabPicture(0)   =   "MainlineDiesel.frx":0A0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LocomotivePrototypeHorsePower"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LocomotivePrototypeWeight"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LocomotivePrototypeAdhesionFactor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LocomotivePrototypeTractionEffort"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LocomotivePrototypeDrawBarPull"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ButtonPrototypeInfo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ComboBoxCouplerManufacturer"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Model"
      TabPicture(1)   =   "MainlineDiesel.frx":0A26
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(5)=   "LocomotiveModelHorsePower"
      Tab(1).Control(6)=   "LocomotiveModelWeight"
      Tab(1).Control(7)=   "LocomotiveModelAdhesionFactor"
      Tab(1).Control(8)=   "LocomotiveModelTractionEffort"
      Tab(1).Control(9)=   "LocomotiveModelDrawBarPull"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Couplers"
      TabPicture(2)   =   "MainlineDiesel.frx":0A42
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LabelCoupler"
      Tab(2).Control(1)=   "LabelCouplerFront(0)"
      Tab(2).Control(2)=   "LabelCouplerRear(1)"
      Tab(2).Control(3)=   "LineCoupler"
      Tab(2).Control(4)=   "Image1"
      Tab(2).Control(5)=   "ComboBoxCouplerFront"
      Tab(2).Control(6)=   "ComboBoxCouplerRear"
      Tab(2).ControlCount=   7
      Begin VB.ComboBox ComboBoxCouplerRear 
         DataField       =   "LocomotiveCouplerTypeRear"
         DataSource      =   "LocomotiveNumber"
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0A5E
         Left            =   -74820
         List            =   "MainlineDiesel.frx":0A9B
         TabIndex        =   29
         Text            =   "Choose a Rear Coupler Type"
         Top             =   2400
         Width           =   2880
      End
      Begin VB.ComboBox ComboBoxCouplerFront 
         DataField       =   "LocomotiveCouplerTypeFront"
         DataSource      =   "LocomotiveNumber"
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0D17
         Left            =   -74820
         List            =   "MainlineDiesel.frx":0D54
         TabIndex        =   28
         Text            =   "Choose a Front Coupler Type"
         Top             =   1680
         Width           =   2880
      End
      Begin VB.ComboBox ComboBoxCouplerManufacturer 
         Height          =   315
         ItemData        =   "MainlineDiesel.frx":0FD0
         Left            =   15175
         List            =   "MainlineDiesel.frx":0FDD
         TabIndex        =   47
         Text            =   "Choose and manufacturer"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton ButtonPrototypeInfo 
         Caption         =   "Prototype Info"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox LocomotiveModelDrawBarPull 
         DataField       =   "LocomotiveModelDrawBarPull"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74825
         TabIndex        =   27
         Text            =   "Draw Bar Pull"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelTractionEffort 
         DataField       =   "LocomotiveModelTractionEffort"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74825
         TabIndex        =   26
         Text            =   "Traction Effort"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelAdhesionFactor 
         DataField       =   "LocomotiveModelAdhesionFactor"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74825
         TabIndex        =   25
         Text            =   "Adhesion Factor"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelWeight 
         DataField       =   "LocomotiveModelWeight"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74825
         TabIndex        =   24
         Text            =   "Weight"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelHorsePower 
         DataField       =   "LocomotiveModelHorsePower"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74825
         TabIndex        =   23
         Text            =   "Horse Power"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeDrawBarPull 
         DataField       =   "LocomotivePrototypeDrawBarPull"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   175
         TabIndex        =   21
         Text            =   "Draw Bar Pull"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeTractionEffort 
         DataField       =   "LocomotivePrototypeTractionEffort"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   175
         TabIndex        =   20
         Text            =   "Traction Effort"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeAdhesionFactor 
         DataField       =   "LocomotivePrototypeAdhesionFactor"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   175
         TabIndex        =   19
         Text            =   "Adhesion Factor"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeWeight 
         DataField       =   "LocomotivePrototypeWeight"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   175
         TabIndex        =   18
         Text            =   "Weight"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeHorsePower 
         DataField       =   "LocomotivePrototypeHorsePower"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   175
         TabIndex        =   17
         Text            =   "Horse Power"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   -74760
         Picture         =   "MainlineDiesel.frx":1004
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Line LineCoupler 
         X1              =   -71995
         X2              =   -74820
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label LabelCouplerRear 
         Caption         =   "The rear coupler is defined as"
         Height          =   255
         Index           =   1
         Left            =   -74820
         TabIndex        =   49
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label LabelCouplerFront 
         Caption         =   "The front copuler is definded as"
         Height          =   255
         Index           =   0
         Left            =   -74820
         TabIndex        =   48
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label LabelCoupler 
         BackColor       =   &H00C0C0C0&
         Caption         =   "This is a general listing of the couplers installed on your model locomotive."
         Height          =   855
         Left            =   -73800
         TabIndex        =   46
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Draw Bar (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   44
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Traction (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Adhesion %"
         Height          =   195
         Left            =   -73320
         TabIndex        =   42
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Weight (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   41
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Horse Power"
         Height          =   255
         Left            =   -73320
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Draw Bar (tons)"
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Traction (tons)"
         Height          =   255
         Left            =   1680
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Adhesion %"
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Weight (tons)"
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horse Power"
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   720
         Width           =   915
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7200
      Top             =   4440
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7200
      Top             =   3840
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label LabelLastTimeLocomotiveViewed 
      Caption         =   "The last time this locomotive's data was viewed was"
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   9840
      Width           =   3675
   End
   Begin VB.Label LabelLastTimeDecoderViewed 
      Caption         =   "The last time this locomotive's decoder was viewed was"
      Height          =   195
      Left            =   120
      TabIndex        =   57
      Top             =   10080
      Width           =   4035
   End
   Begin VB.Label LabelLastTimeDecoderRead 
      Caption         =   "The last time this locomotive's decoder was read was"
      Height          =   195
      Left            =   120
      TabIndex        =   56
      Top             =   10320
      Width           =   3795
   End
   Begin VB.Label LabelLastTimeDecoderWritten 
      Caption         =   "The last time this locomotive's decoder was written to was"
      Height          =   195
      Left            =   120
      TabIndex        =   55
      Top             =   10560
      Width           =   4080
   End
   Begin VB.Label LabelAdditionalNotes 
      Caption         =   "Additional Notes:"
      Height          =   195
      Left            =   120
      TabIndex        =   54
      Top             =   9600
      Width           =   1200
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Label Label11 
      Caption         =   $"MainlineDiesel.frx":52CE
      Height          =   615
      Left            =   840
      TabIndex        =   45
      Top             =   120
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog PictureGet 
      Left            =   3360
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open an Image"
      Filter          =   "Picture Files (*.gif)|*.gif|Picture FIles (*.jpg)|*.jpg"
   End
   Begin VB.Image LocomotiveImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "MainlineDiesel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButtonChangePicture_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Loading the Picture File Name Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    PictureGet.ShowOpen
    Let LocomotiveImageFilename.Text = PictureGet.Filename
    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub ButtonClose_Click()

Let textboxlasttimelocomotiveviewed.Text = Date$ & " at " & Time$
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
            If TemporaryScreen = "Mainline Diesel Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Diesel Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Diesel Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Diesel Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub ButtonDecoderProg_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Loading the Programming Decoder Window
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Load MainlineDecoder
    MainlineDecoder.Show vbModeless

End Sub

Private Sub ButtonGo_Click()
      
    Let textboxlasttimelocomotiveviewed.Text = Date$ & " at " & Time$
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set all properties to false
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LocomotivePrototype.Enabled = False
    Let LocomotiveActiveList.Enabled = False
    Let LocomotiveDecoderEquiped.Enabled = False
    Let LocomotiveManufacturer.Enabled = False
    Let LocomotiveModel.Enabled = False
    Let LocomotiveModelManufacturer.Enabled = False
    Let LocomotiveFacts.Enabled = False
    Let LocomotiveImage.Enabled = False
    Let LocomotiveImageFilename.Enabled = False
    Let SSTab1.Enabled = False
    Let LocomotivePrototypeHorsePower.Enabled = False
    Let LocomotivePrototypeWeight.Enabled = False
    Let LocomotivePrototypeAdhesionFactor.Enabled = False
    Let LocomotivePrototypeTractionEffort.Enabled = False
    Let LocomotivePrototypeDrawBarPull.Enabled = False
    Let LocomotiveModelHorsePower.Enabled = False
    Let LocomotiveModelWeight.Enabled = False
    Let LocomotiveModelAdhesionFactor.Enabled = False
    Let LocomotiveModelTractionEffort.Enabled = False
    Let LocomotiveModelDrawBarPull.Enabled = False
    Let ButtonChangePicture.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check for valid entries before searching
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(LocomotiveNumberGO.Text) < 1 Or Val(LocomotiveNumberGO.Text) > 127 And CheckAddressShort.Value = vbChecked Then
        MsgBox "The value entered for the locomotive number is not valid. Please" & Chr$(13) & "use the values between one (1) and one hundred and twenty-seven" & Chr$(13) & "(127) for a short address as you have selected.", vbOKOnly + vbExclamation, "Locomotive Entry Error"
    ElseIf Val(LocomotiveNumberGO.Text) < 1 Or Val(LocomotiveNumberGO.Text) > 9999 And CheckAddressShort.Value = vbUnchecked Then
        MsgBox "The value entered for the locomotive number is not valid. Please" & Chr$(13) & "use values between one (1) and nine thousand nine hundred and ninety-nine" & Chr$(13) & "(9999) only.", vbOKOnly + vbExclamation, "Locomotive Entry Error"
    Else
        Let MousePointer = vbHourglass
        Let Temporary = "LocomotiveNumber like " + Str$(Val(LocomotiveNumberGO.Text))
        LocomotiveNumber.Recordset.FindFirst Temporary
        Let MousePointer = vbDefault
        If Not LocomotiveNumber.Recordset.NoMatch Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Record found, set enabled properties to true
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let LocomotivePrototype.Enabled = True
            Let LocomotiveActiveList.Enabled = True
            Let LocomotiveDecoderEquiped.Enabled = True
            Let LocomotiveManufacturer.Enabled = True
            Let LocomotiveModel.Enabled = True
            Let LocomotiveModelManufacturer.Enabled = True
            Let LocomotiveFacts.Enabled = True
            Let LocomotiveImage.Enabled = True
            Let LocomotiveImageFilename.Enabled = True
            Let ButtonChangePicture.Enabled = True
            Let SSTab1.Enabled = True
            Let LocomotivePrototypeHorsePower.Enabled = True
            Let LocomotivePrototypeWeight.Enabled = True
            Let LocomotivePrototypeAdhesionFactor.Enabled = True
            Let LocomotivePrototypeTractionEffort.Enabled = True
            Let LocomotivePrototypeDrawBarPull.Enabled = True
            Let LocomotiveModelHorsePower.Enabled = True
            Let LocomotiveModelWeight.Enabled = True
            Let LocomotiveModelAdhesionFactor.Enabled = True
            Let LocomotiveModelTractionEffort.Enabled = True
            Let LocomotiveModelDrawBarPull.Enabled = True
            Let ButtonDecoderProg.Visible = True
        Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' No record found, set enabled properties to false
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            MsgBox "The locomotive number was not found in this database.", vbOKOnly + vbExclamation, "Locomotive Entry Error"
            Let LocomotivePrototype.Enabled = False
            Let LocomotiveActiveList.Enabled = False
            Let LocomotiveDecoderEquiped.Enabled = False
            Let LocomotiveManufacturer.Enabled = False
            Let LocomotiveModel.Enabled = False
            Let LocomotiveModelManufacturer.Enabled = False
            Let LocomotiveFacts.Enabled = False
            Let LocomotiveImage.Enabled = False
            Let LocomotiveImageFilename.Enabled = False
            Let ButtonChangePicture.Enabled = False
            Let SSTab1.Enabled = False
            Let LocomotivePrototypeHorsePower.Enabled = False
            Let LocomotivePrototypeWeight.Enabled = False
            Let LocomotivePrototypeAdhesionFactor.Enabled = False
            Let LocomotivePrototypeTractionEffort.Enabled = False
            Let LocomotivePrototypeDrawBarPull.Enabled = False
            Let LocomotiveModelHorsePower.Enabled = False
            Let LocomotiveModelWeight.Enabled = False
            Let LocomotiveModelAdhesionFactor.Enabled = False
            Let LocomotiveModelTractionEffort.Enabled = False
            Let LocomotiveModelDrawBarPull.Enabled = False
            Let LocomotiveDecoderProgram.Visible = False
            If MainScreen!MenuLogFile.Caption = "&Log File is Off" Then
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Error
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbExclamation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "ProgrammingDiesel, ButonGo, Click, locomotive number was not found in the database."
            End If
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub ButtonList_Click()

Load LocomotiveSpreadsheet
LocomotiveSpreadsheet.Show vbModeless

End Sub

Private Sub ButtonPrint_Click()

    MainlineDiesel.PrintForm
    
End Sub

Private Sub ButtonPrototypeInfo_Click()
MainlineDiesel.Hide
Load MainlinePrototypeInfo
MainlinePrototypeInfo.Show vbModeless

End Sub



Private Sub ButtonUpdate_Click()
        
    LocomotiveNumber.Recordset.Edit
    LocomotiveNumber.Recordset.Update

    Let ButtonUpdate.Enabled = False

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub









Private Sub CheckAddressShort_Click()
    
    Let ButtonClose.Enabled = False
    Let MainlineDiesel.MousePointer = vbArrowHourglass
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set all properties to false
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LocomotivePrototype.Enabled = False
    Let LocomotiveActiveList.Enabled = False
    Let LocomotiveDecoderEquiped.Enabled = False
    Let LocomotiveManufacturer.Enabled = False
    Let LocomotiveModel.Enabled = False
    Let LocomotiveModelManufacturer.Enabled = False
    Let LocomotiveFacts.Enabled = False
    Let LocomotiveImage.Enabled = False
    Let LocomotiveImageFilename.Enabled = False
    Let SSTab1.Enabled = False
    Let LocomotivePrototypeHorsePower.Enabled = False
    Let LocomotivePrototypeWeight.Enabled = False
    Let LocomotivePrototypeAdhesionFactor.Enabled = False
    Let LocomotivePrototypeTractionEffort.Enabled = False
    Let LocomotivePrototypeDrawBarPull.Enabled = False
    Let LocomotiveModelHorsePower.Enabled = False
    Let LocomotiveModelWeight.Enabled = False
    Let LocomotiveModelAdhesionFactor.Enabled = False
    Let LocomotiveModelTractionEffort.Enabled = False
    Let LocomotiveModelDrawBarPull.Enabled = False
    Let ButtonChangePicture.Enabled = False
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Move to first (blank) record for binding objects
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveNumber.Recordset.MoveFirst
    Let textboxlasttimelocomotiveviewed.Text = Date$ & " at " & Time$
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the recordsource property
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If CheckAddressShort.Value = vbChecked Then
        Let LocomotiveNumber.RecordSource = "LocomotiveDatabaseShortV030002"
        Let LocomotiveDatabaseSearch.RecordSource = "LocomotiveDatabaseShortV030002"
    Else
        Let LocomotiveNumber.RecordSource = "LocomotiveDatabaseLongV030002"
        Let LocomotiveDatabaseSearch.RecordSource = "LocomotiveDatabaseLongV030002"
    End If
    LocomotiveNumber.Refresh
    LocomotiveDatabaseSearch.Refresh
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the locomotive selection from database
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveNumberGO.Clear
    LocomotiveDatabaseSearch.Recordset.MoveFirst
    Do While Not LocomotiveDatabaseSearch.Recordset.EOF
        LocomotiveDatabaseSearch.Recordset.MoveNext
        If LocomotiveDatabaseActiveSearch.Value = vbChecked Then
             LocomotiveNumberGO.AddItem Format(Val(LocomotiveDatabaseNumberSearch.Text), "0000")
        End If
    Loop
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the prototype owners list
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotivePrototype.Clear
    ReportingMarksDatabase.Recordset.MoveFirst
    While Not ReportingMarksDatabase.Recordset.EOF
        Let TemporaryReportingMark = ReportingMarksDatabase.Recordset.Fields("ReportingMarks") + " (" + ReportingMarksDatabase.Recordset.Fields("RegisteredOwner") + ")"
        LocomotivePrototype.AddItem TemporaryReportingMark
        ReportingMarksDatabase.Recordset.MoveNext
        DoEvents
    Wend
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Return properties to
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let ButtonClose.Enabled = True
    Let MainlineDiesel.MousePointer = vbDefault

End Sub

Private Sub ComboBoxCouplerFront_Change()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub ComboBoxCouplerRear_Change()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub








Private Sub Form_Activate()

    DoEvents
    
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
        If TemporaryScreen = "Programming Diesel Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Programming Diesel Screen"
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
            Let Ini.Value = "Programming Diesel Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Programming Diesel Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Programming Diesel Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    Call CheckAddressShort_Click
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Diesel Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineDiesel.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineDiesel.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineDiesel.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineDiesel.Height)

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
                Let Ini.Value = "Mainline Diesel Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Diesel Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineDiesel.Hide
    'unload MainlineDiesel

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
    Let Ini.Application = "Mainline Diesel Screen"
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
        MainlineDiesel.Left = (Screen.Width - Width) / 2
        MainlineDiesel.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineDiesel.Width > Screen.Width Then
            Let MainlineDiesel.Left = Screen.Width - MainlineDiesel.Width
        Else
            Let MainlineDiesel.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineDiesel.Height > Screen.Height Then
            Let MainlineDiesel.Top = Screen.Height - MainlineDiesel.Height
        Else
            Let MainlineDiesel.Top = Val(TemporaryValueTop)
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
    Call BalloonHelpUpdatePart01
    Call BalloonHelpUpdatePart02
    Call BalloonHelpUpdatePart03
    Call BalloonHelpUpdatePart04

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let LocomotiveNumber.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
    Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
    Let ReportingMarksDatabase.DatabaseName = App.Path$ & "\Databases\ReportingMarksDatabase.mdb"
    Let PictureGet.InitDir = App.Path$ & "\Pictures"

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

    If MainlineDiesel.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineDiesel.WindowState = vbNormal Then
    
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

Private Sub LocomotiveActiveList_GotFocus()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub LocomotiveDecoderEquiped_GotFocus()

    Let ButtonUpdate.Enabled = True
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub




Private Sub LocomotiveDecoderEquiped_Click()
    
    If LocomotiveDecoderEquiped.Value = vbUnchecked Then
        Let ButtonDecoderProg.Enabled = False
    Else 'If LocomotiveDecoderEquiped.Value = vbChecked Then
        Let ButtonDecoderProg.Enabled = True
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveFacts_Click()

Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveImageFilename_Change()

    On Error Resume Next
    Let LocomotiveImage.Picture = LoadPicture(LocomotiveImageFilename.Text)
    If Err.Number = 53 Then
        LocomotiveImage.Picture = LoadPicture()
        MsgBox "Your picture listed on file was not found." + Chr$(13) + "Please update this record.", vbExclamation, "Locomotive Picture not Found"
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub LocomotiveImageFilename_KeyPress(KeyAscii As Integer)

Let ButtonUpdate.Enabled = True

End Sub


Private Sub LocomotiveManufacturer_Change()

If LocomotiveManufacturer.Text = "American Locomotive Company" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "C-636"
    LocomotiveModel.AddItem "C-630"
    LocomotiveModel.AddItem "C-628"
    LocomotiveModel.AddItem "C-430"
    LocomotiveModel.AddItem "C-425"
    LocomotiveModel.AddItem "C-424"
    LocomotiveModel.AddItem "C-420"
    LocomotiveModel.AddItem "RS36"
    LocomotiveModel.AddItem "RS32"
    LocomotiveModel.AddItem "RS27"
    LocomotiveModel.AddItem "RSD15"
    LocomotiveModel.AddItem "RSD12"
    LocomotiveModel.AddItem "RD11"
    LocomotiveModel.AddItem "RS3"
    LocomotiveModel.AddItem "RS2"
    LocomotiveModel.AddItem "RS1"
    LocomotiveModel.AddItem "S-1"
    LocomotiveModel.AddItem "S-2"
    LocomotiveModel.AddItem "S-3"
    LocomotiveModel.AddItem "S-4"
    LocomotiveModel.AddItem "S-6"
    LocomotiveModel.AddItem "T-6"
    LocomotiveModel.AddItem "PA-3"
    LocomotiveModel.AddItem "PA-2"
    LocomotiveModel.AddItem "PA-1"
    LocomotiveModel.AddItem "FA-2"
    LocomotiveModel.AddItem "FA-1"
End If

If LocomotiveManufacturer.Text = "Baldwin" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "VO 1000"
    LocomotiveModel.AddItem "DRS-6-6-15 or AS-616"
    LocomotiveModel.AddItem "DRS-6-4-15 or AS-416"
    LocomotiveModel.AddItem "DRS-4-4-15 or AS-16"
    LocomotiveModel.AddItem "DS-4-4-10"
    LocomotiveModel.AddItem "S-12"
End If

If LocomotiveManufacturer.Text = "Bombardier" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "HR616"
    End If
    
If LocomotiveManufacturer.Text = "FairBanks-Morse" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "H-12-44"
End If

If LocomotiveManufacturer.Text = "General Electric" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "AC6000"
    LocomotiveModel.AddItem "AC4400CW"
    LocomotiveModel.AddItem "9-44CW"
    LocomotiveModel.AddItem "9-44C"
    LocomotiveModel.AddItem "9-44BW"
    LocomotiveModel.AddItem "9-44B"
    LocomotiveModel.AddItem "9-40CW"
    LocomotiveModel.AddItem "9-40C"
    LocomotiveModel.AddItem "9-40BW"
    LocomotiveModel.AddItem "9-40B"
    LocomotiveModel.AddItem "8-41CW"
    LocomotiveModel.AddItem "8-41C"
    LocomotiveModel.AddItem "8-40CW"
    LocomotiveModel.AddItem "8-40CM"
    LocomotiveModel.AddItem "8-40C"
    LocomotiveModel.AddItem "8-40BW"
    LocomotiveModel.AddItem "8-40B"
    LocomotiveModel.AddItem "8-39CE"
    LocomotiveModel.AddItem "8-39C"
    LocomotiveModel.AddItem "8-39B"
    LocomotiveModel.AddItem "8-32C"
    LocomotiveModel.AddItem "8-32BWH or P32-8BWH"
    LocomotiveModel.AddItem "8-32B"
    LocomotiveModel.AddItem "7-36B"
    LocomotiveModel.AddItem "7-36C"
    LocomotiveModel.AddItem "7-33C"
    LocomotiveModel.AddItem "7-30B-A1"
    LocomotiveModel.AddItem "7-30C-A"
    LocomotiveModel.AddItem "7-30B-A"
    LocomotiveModel.AddItem "7-30C"
    LocomotiveModel.AddItem "7-30B"
    LocomotiveModel.AddItem "7-23BQ"
    LocomotiveModel.AddItem "7-23B"
    LocomotiveModel.AddItem "U36C"
    LocomotiveModel.AddItem "U36B"
    LocomotiveModel.AddItem "U33C"
    LocomotiveModel.AddItem "U33B"
    LocomotiveModel.AddItem "SF30C"
    LocomotiveModel.AddItem "U30C"
    LocomotiveModel.AddItem "U30B"
    LocomotiveModel.AddItem "U28B"
    LocomotiveModel.AddItem "U25C"
    LocomotiveModel.AddItem "U25B"
    LocomotiveModel.AddItem "U23C"
    LocomotiveModel.AddItem "U23B"
    LocomotiveModel.AddItem "U18B"
End If

If LocomotiveManufacturer.Text = "General Motors - Electromotive Division" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "SD90/43MAC"
    LocomotiveModel.AddItem "SD80MAC"
    LocomotiveModel.AddItem "SD80MC"
    LocomotiveModel.AddItem "SD75M"
    LocomotiveModel.AddItem "SD70MAC"
    LocomotiveModel.AddItem "SD70M"
    LocomotiveModel.AddItem "SD70I"
    LocomotiveModel.AddItem "SD70"
    LocomotiveModel.AddItem "F69PH-AC"
    LocomotiveModel.AddItem "SD60MAC"
    LocomotiveModel.AddItem "SD60M"
    LocomotiveModel.AddItem "SD60F"
    LocomotiveModel.AddItem "SD60I"
    LocomotiveModel.AddItem "SD60"
    LocomotiveModel.AddItem "SD50M"
    LocomotiveModel.AddItem "SD50F"
    LocomotiveModel.AddItem "SD50I"
    LocomotiveModel.AddItem "SD50"
    LocomotiveModel.AddItem "GP60M"
    LocomotiveModel.AddItem "GP60"
    LocomotiveModel.AddItem "AMD-103"
    LocomotiveModel.AddItem "F59PHI"
    LocomotiveModel.AddItem "F59PH"
    LocomotiveModel.AddItem "GP59"
    LocomotiveModel.AddItem "GP50"
    LocomotiveModel.AddItem "DD40AX"
    LocomotiveModel.AddItem "SD45-2T"
    LocomotiveModel.AddItem "SD45-2"
    LocomotiveModel.AddItem "SDP45"
    LocomotiveModel.AddItem "FP45"
    LocomotiveModel.AddItem "F45"
    LocomotiveModel.AddItem "SD45"
    LocomotiveModel.AddItem "SD40-2W"
    LocomotiveModel.AddItem "SD40-2T"
    LocomotiveModel.AddItem "SD40-2F"
    LocomotiveModel.AddItem "SD40-2"
    LocomotiveModel.AddItem "SDP40F"
    LocomotiveModel.AddItem "SDP40"
    LocomotiveModel.AddItem "F40PHM-2"
    LocomotiveModel.AddItem "F40PH-2C"
    LocomotiveModel.AddItem "F40PH"
    LocomotiveModel.AddItem "F40C"
    LocomotiveModel.AddItem "SD40"
    LocomotiveModel.AddItem "GP40-2"
    LocomotiveModel.AddItem "GP40X"
    LocomotiveModel.AddItem "GP40P"
    LocomotiveModel.AddItem "GP40W"
    LocomotiveModel.AddItem "GP40"
    LocomotiveModel.AddItem "GP39-2"
    LocomotiveModel.AddItem "SDL39"
    LocomotiveModel.AddItem "SD39"
    LocomotiveModel.AddItem "GP39"
    LocomotiveModel.AddItem "SD38-2"
    LocomotiveModel.AddItem "GP39-2W"
    LocomotiveModel.AddItem "GP38-2"
    LocomotiveModel.AddItem "SD38"
    LocomotiveModel.AddItem "GP38"
    LocomotiveModel.AddItem "SDP35"
    LocomotiveModel.AddItem "SD35"
    LocomotiveModel.AddItem "GP35"
    LocomotiveModel.AddItem "GP30"
    LocomotiveModel.AddItem "SD24"
    LocomotiveModel.AddItem "GP20"
    LocomotiveModel.AddItem "GP15-1"
    LocomotiveModel.AddItem "GP15T"
    LocomotiveModel.AddItem "GP15"
    LocomotiveModel.AddItem "SD18"
    LocomotiveModel.AddItem "GP18"
    LocomotiveModel.AddItem "SD9"
    LocomotiveModel.AddItem "GP9"
    LocomotiveModel.AddItem "SD7"
    LocomotiveModel.AddItem "GP7"
    LocomotiveModel.AddItem "CF7"
    LocomotiveModel.AddItem "GMD1"
    LocomotiveModel.AddItem "RS1325"
    LocomotiveModel.AddItem "NM5"
    LocomotiveModel.AddItem "NW2"
    LocomotiveModel.AddItem "TR5"
    LocomotiveModel.AddItem "TR4"
    LocomotiveModel.AddItem "MP15T"
    LocomotiveModel.AddItem "MP15AC"
    LocomotiveModel.AddItem "MP15(DC)"
    LocomotiveModel.AddItem "SW1504"
    LocomotiveModel.AddItem "SW1500"
    LocomotiveModel.AddItem "SW1200"
    LocomotiveModel.AddItem "SW1001"
    LocomotiveModel.AddItem "SW1000"
    LocomotiveModel.AddItem "SW900"
    LocomotiveModel.AddItem "SW600"
    LocomotiveModel.AddItem "SW9"
    LocomotiveModel.AddItem "SW8"
    LocomotiveModel.AddItem "SW7"
    LocomotiveModel.AddItem "SW1"
    LocomotiveModel.AddItem "FL9"
    LocomotiveModel.AddItem "FP9"
    LocomotiveModel.AddItem "F9"
    LocomotiveModel.AddItem "FP7"
    LocomotiveModel.AddItem "F7"
    LocomotiveModel.AddItem "F3"
    LocomotiveModel.AddItem "E9A"
    LocomotiveModel.AddItem "E8A"
    LocomotiveModel.AddItem "BL2"

End If

If LocomotiveManufacturer.Text = "Montreal Locomotive Works" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "S-13"
        LocomotiveModel.AddItem "M636"
        LocomotiveModel.AddItem "M630"
        LocomotiveModel.AddItem "M420"
        LocomotiveModel.AddItem "M420R"
End If

If LocomotiveManufacturer.Text = "Morrison Knudsen" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "MK5000C"
        LocomotiveModel.AddItem "MK1200G"
        LocomotiveModel.AddItem "MK-F40PHL-2"
        LocomotiveModel.AddItem "MKGP40FH-2"
End If

Let ButtonUpdate.Enabled = True

End Sub

Private Sub LocomotiveManufacturer_Click()

If LocomotiveManufacturer.Text = "American Locomotive Company" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "C-636"
    LocomotiveModel.AddItem "C-630"
    LocomotiveModel.AddItem "C-628"
    LocomotiveModel.AddItem "C-430"
    LocomotiveModel.AddItem "C-425"
    LocomotiveModel.AddItem "C-424"
    LocomotiveModel.AddItem "C-420"
    LocomotiveModel.AddItem "RS36"
    LocomotiveModel.AddItem "RS32"
    LocomotiveModel.AddItem "RS27"
    LocomotiveModel.AddItem "RSD15"
    LocomotiveModel.AddItem "RSD12"
    LocomotiveModel.AddItem "RD11"
    LocomotiveModel.AddItem "RS3"
    LocomotiveModel.AddItem "RS2"
    LocomotiveModel.AddItem "RS1"
    LocomotiveModel.AddItem "S-1"
    LocomotiveModel.AddItem "S-2"
    LocomotiveModel.AddItem "S-3"
    LocomotiveModel.AddItem "S-4"
    LocomotiveModel.AddItem "S-6"
    LocomotiveModel.AddItem "T-6"
    LocomotiveModel.AddItem "PA-3"
    LocomotiveModel.AddItem "PA-2"
    LocomotiveModel.AddItem "PA-1"
    LocomotiveModel.AddItem "FA-2"
    LocomotiveModel.AddItem "FA-1"
End If

If LocomotiveManufacturer.Text = "Baldwin" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "VO 1000"
    LocomotiveModel.AddItem "DRS-6-6-15 or AS-616"
    LocomotiveModel.AddItem "DRS-6-4-15 or AS-416"
    LocomotiveModel.AddItem "DRS-4-4-15 or AS-16"
    LocomotiveModel.AddItem "DS-4-4-10"
    LocomotiveModel.AddItem "S-12"
End If

If LocomotiveManufacturer.Text = "Bombardier" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "HR616"
    End If
    
If LocomotiveManufacturer.Text = "FairBanks-Morse" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "H-12-44"
End If

If LocomotiveManufacturer.Text = "General Electric" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "AC6000"
    LocomotiveModel.AddItem "AC4400CW"
    LocomotiveModel.AddItem "9-44CW"
    LocomotiveModel.AddItem "9-44C"
    LocomotiveModel.AddItem "9-44BW"
    LocomotiveModel.AddItem "9-44B"
    LocomotiveModel.AddItem "9-40CW"
    LocomotiveModel.AddItem "9-40C"
    LocomotiveModel.AddItem "9-40BW"
    LocomotiveModel.AddItem "9-40B"
    LocomotiveModel.AddItem "8-41CW"
    LocomotiveModel.AddItem "8-41C"
    LocomotiveModel.AddItem "8-40CW"
    LocomotiveModel.AddItem "8-40CM"
    LocomotiveModel.AddItem "8-40C"
    LocomotiveModel.AddItem "8-40BW"
    LocomotiveModel.AddItem "8-40B"
    LocomotiveModel.AddItem "8-39CE"
    LocomotiveModel.AddItem "8-39C"
    LocomotiveModel.AddItem "8-39B"
    LocomotiveModel.AddItem "8-32C"
    LocomotiveModel.AddItem "8-32BWH or P32-8BWH"
    LocomotiveModel.AddItem "8-32B"
    LocomotiveModel.AddItem "7-36B"
    LocomotiveModel.AddItem "7-36C"
    LocomotiveModel.AddItem "7-33C"
    LocomotiveModel.AddItem "7-30B-A1"
    LocomotiveModel.AddItem "7-30C-A"
    LocomotiveModel.AddItem "7-30B-A"
    LocomotiveModel.AddItem "7-30C"
    LocomotiveModel.AddItem "7-30B"
    LocomotiveModel.AddItem "7-23BQ"
    LocomotiveModel.AddItem "7-23B"
    LocomotiveModel.AddItem "U36C"
    LocomotiveModel.AddItem "U36B"
    LocomotiveModel.AddItem "U33C"
    LocomotiveModel.AddItem "U33B"
    LocomotiveModel.AddItem "SF30C"
    LocomotiveModel.AddItem "U30C"
    LocomotiveModel.AddItem "U30B"
    LocomotiveModel.AddItem "U28B"
    LocomotiveModel.AddItem "U25C"
    LocomotiveModel.AddItem "U25B"
    LocomotiveModel.AddItem "U23C"
    LocomotiveModel.AddItem "U23B"
    LocomotiveModel.AddItem "U18B"
End If

If LocomotiveManufacturer.Text = "General Motors - Electromotive Division" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "SD90/43MAC"
    LocomotiveModel.AddItem "SD80MAC"
    LocomotiveModel.AddItem "SD80MC"
    LocomotiveModel.AddItem "SD75M"
    LocomotiveModel.AddItem "SD70MAC"
    LocomotiveModel.AddItem "SD70M"
    LocomotiveModel.AddItem "SD70I"
    LocomotiveModel.AddItem "SD70"
    LocomotiveModel.AddItem "F69PH-AC"
    LocomotiveModel.AddItem "SD60MAC"
    LocomotiveModel.AddItem "SD60M"
    LocomotiveModel.AddItem "SD60F"
    LocomotiveModel.AddItem "SD60I"
    LocomotiveModel.AddItem "SD60"
    LocomotiveModel.AddItem "SD50M"
    LocomotiveModel.AddItem "SD50F"
    LocomotiveModel.AddItem "SD50I"
    LocomotiveModel.AddItem "SD50"
    LocomotiveModel.AddItem "GP60M"
    LocomotiveModel.AddItem "GP60"
    LocomotiveModel.AddItem "AMD-103"
    LocomotiveModel.AddItem "F59PHI"
    LocomotiveModel.AddItem "F59PH"
    LocomotiveModel.AddItem "GP59"
    LocomotiveModel.AddItem "GP50"
    LocomotiveModel.AddItem "DD40AX"
    LocomotiveModel.AddItem "SD45-2T"
    LocomotiveModel.AddItem "SD45-2"
    LocomotiveModel.AddItem "SDP45"
    LocomotiveModel.AddItem "FP45"
    LocomotiveModel.AddItem "F45"
    LocomotiveModel.AddItem "SD45"
    LocomotiveModel.AddItem "SD40-2W"
    LocomotiveModel.AddItem "SD40-2T"
    LocomotiveModel.AddItem "SD40-2F"
    LocomotiveModel.AddItem "SD40-2"
    LocomotiveModel.AddItem "SDP40F"
    LocomotiveModel.AddItem "SDP40"
    LocomotiveModel.AddItem "F40PHM-2"
    LocomotiveModel.AddItem "F40PH-2C"
    LocomotiveModel.AddItem "F40PH"
    LocomotiveModel.AddItem "F40C"
    LocomotiveModel.AddItem "SD40"
    LocomotiveModel.AddItem "GP40-2"
    LocomotiveModel.AddItem "GP40X"
    LocomotiveModel.AddItem "GP40P"
    LocomotiveModel.AddItem "GP40W"
    LocomotiveModel.AddItem "GP40"
    LocomotiveModel.AddItem "GP39-2"
    LocomotiveModel.AddItem "SDL39"
    LocomotiveModel.AddItem "SD39"
    LocomotiveModel.AddItem "GP39"
    LocomotiveModel.AddItem "SD38-2"
    LocomotiveModel.AddItem "GP39-2W"
    LocomotiveModel.AddItem "GP38-2"
    LocomotiveModel.AddItem "SD38"
    LocomotiveModel.AddItem "GP38"
    LocomotiveModel.AddItem "SDP35"
    LocomotiveModel.AddItem "SD35"
    LocomotiveModel.AddItem "GP35"
    LocomotiveModel.AddItem "GP30"
    LocomotiveModel.AddItem "SD24"
    LocomotiveModel.AddItem "GP20"
    LocomotiveModel.AddItem "GP15-1"
    LocomotiveModel.AddItem "GP15T"
    LocomotiveModel.AddItem "GP15"
    LocomotiveModel.AddItem "SD18"
    LocomotiveModel.AddItem "GP18"
    LocomotiveModel.AddItem "SD9"
    LocomotiveModel.AddItem "GP9"
    LocomotiveModel.AddItem "SD7"
    LocomotiveModel.AddItem "GP7"
    LocomotiveModel.AddItem "CF7"
    LocomotiveModel.AddItem "GMD1"
    LocomotiveModel.AddItem "RS1325"
    LocomotiveModel.AddItem "NM5"
    LocomotiveModel.AddItem "NW2"
    LocomotiveModel.AddItem "TR5"
    LocomotiveModel.AddItem "TR4"
    LocomotiveModel.AddItem "MP15T"
    LocomotiveModel.AddItem "MP15AC"
    LocomotiveModel.AddItem "MP15(DC)"
    LocomotiveModel.AddItem "SW1504"
    LocomotiveModel.AddItem "SW1500"
    LocomotiveModel.AddItem "SW1200"
    LocomotiveModel.AddItem "SW1001"
    LocomotiveModel.AddItem "SW1000"
    LocomotiveModel.AddItem "SW900"
    LocomotiveModel.AddItem "SW600"
    LocomotiveModel.AddItem "SW9"
    LocomotiveModel.AddItem "SW8"
    LocomotiveModel.AddItem "SW7"
    LocomotiveModel.AddItem "SW1"
    LocomotiveModel.AddItem "FL9"
    LocomotiveModel.AddItem "FP9"
    LocomotiveModel.AddItem "F9"
    LocomotiveModel.AddItem "FP7"
    LocomotiveModel.AddItem "F7"
    LocomotiveModel.AddItem "F3"
    LocomotiveModel.AddItem "E9A"
    LocomotiveModel.AddItem "E8A"
    LocomotiveModel.AddItem "BL2"

End If

If LocomotiveManufacturer.Text = "Montreal Locomotive Works" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "S-13"
        LocomotiveModel.AddItem "M636"
        LocomotiveModel.AddItem "M630"
        LocomotiveModel.AddItem "M420"
        LocomotiveModel.AddItem "M420R"
End If

If LocomotiveManufacturer.Text = "Morrison Knudsen" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "MK5000C"
        LocomotiveModel.AddItem "MK1200G"
        LocomotiveModel.AddItem "MK-F40PHL-2"
        LocomotiveModel.AddItem "MKGP40FH-2"
End If

Let ButtonUpdate.Enabled = True

End Sub


Private Sub LocomotiveModel_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelAdhesionFactor_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelDrawBarPull_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelHorsePower_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelManufacturer_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelTractionEffort_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveModelWeight_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotiveNumberGO_Click()

Let textboxlasttimelocomotiveviewed.Text = Date$ & " at " & Time$

' My Programming Notes
'
' I just wanted some way to show the user of the program that the search in the data base has been started; and to
' please wait unitl the results are displayed.

       Let MousePointer = vbHourglass
    
' -------------------------------------------------------------------------------------------------------------------------
' Searching for
    
       Let Temporary = "LocomotiveNumber like" + Str$(Val(LocomotiveNumberGO.Text))
    
' -------------------------------------------------------------------------------------------------------------------------
' Findfirst Method

    LocomotiveNumber.Recordset.FindFirst Temporary
    Let MousePointer = vbDefault
    If Not LocomotiveNumber.Recordset.NoMatch Then
            Let LocomotivePrototype.Enabled = True
            Let LocomotiveActiveList.Enabled = True
            Let LocomotiveDecoderEquiped.Enabled = True
            Let LocomotiveManufacturer.Enabled = True
            Let LocomotiveModel.Enabled = True
            Let LocomotiveModelManufacturer.Enabled = True
            Let LocomotiveFacts.Enabled = True
            Let LocomotiveImage.Enabled = True
            Let LocomotiveImageFilename.Enabled = True
            Let SSTab1.Enabled = True
            Let LocomotivePrototypeHorsePower.Enabled = True
            Let LocomotivePrototypeWeight.Enabled = True
            Let LocomotivePrototypeAdhesionFactor.Enabled = True
            Let LocomotivePrototypeTractionEffort.Enabled = True
            Let LocomotivePrototypeDrawBarPull.Enabled = True
            Let LocomotiveModelHorsePower.Enabled = True
            Let LocomotiveModelWeight.Enabled = True
            Let LocomotiveModelAdhesionFactor.Enabled = True
            Let LocomotiveModelTractionEffort.Enabled = True
            Let LocomotiveModelDrawBarPull.Enabled = True
            Let ButtonChangePicture.Enabled = True
    Else
            MsgBox "No such locomotive could be found in the database.", vbExclamation, "Locomotive Entry Error"
            Let LocomotivePrototype.Enabled = False
            Let LocomotiveActiveList.Enabled = False
            Let LocomotiveDecoderEquiped.Enabled = False
            Let LocomotiveManufacturer.Enabled = False
            Let LocomotiveModel.Enabled = False
            Let LocomotiveModelManufacturer.Enabled = False
            Let LocomotiveFacts.Enabled = False
            Let LocomotiveImage.Enabled = False
            Let LocomotiveImageFilename.Enabled = False
            Let SSTab1.Enabled = False
            Let LocomotivePrototypeHorsePower.Enabled = False
            Let LocomotivePrototypeWeight.Enabled = False
            Let LocomotivePrototypeAdhesionFactor.Enabled = False
            Let LocomotivePrototypeTractionEffort.Enabled = False
            Let LocomotivePrototypeDrawBarPull.Enabled = False
            Let LocomotiveModelHorsePower.Enabled = False
            Let LocomotiveModelWeight.Enabled = False
            Let LocomotiveModelAdhesionFactor.Enabled = False
            Let LocomotiveModelTractionEffort.Enabled = False
            Let LocomotiveModelDrawBarPull.Enabled = False
            Let ButtonChangePicture.Enabled = False
    End If

End Sub


Private Sub LocomotiveNumberGO_KeyDown(KeyCode As Integer, Shift As Integer)

LocomotiveNumber.Recordset.MoveFirst
        
        Let LocomotivePrototype.Enabled = False
        Let LocomotiveActiveList.Enabled = False
        Let LocomotiveDecoderEquiped.Enabled = False
        Let LocomotiveManufacturer.Enabled = False
        Let LocomotiveModel.Enabled = False
        Let LocomotiveModelManufacturer.Enabled = False
        Let LocomotiveFacts.Enabled = False
        Let LocomotiveImage.Enabled = False
        Let LocomotiveImageFilename.Enabled = False
        Let SSTab1.Enabled = False
        Let LocomotivePrototypeHorsePower.Enabled = False
        Let LocomotivePrototypeWeight.Enabled = False
        Let LocomotivePrototypeAdhesionFactor.Enabled = False
        Let LocomotivePrototypeTractionEffort.Enabled = False
        Let LocomotivePrototypeDrawBarPull.Enabled = False
        Let LocomotiveModelHorsePower.Enabled = False
        Let LocomotiveModelWeight.Enabled = False
        Let LocomotiveModelAdhesionFactor.Enabled = False
        Let LocomotiveModelTractionEffort.Enabled = False
        Let LocomotiveModelDrawBarPull.Enabled = False
        Let ButtonChangePicture.Enabled = False
End Sub


Private Sub LocomotivePrototype_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotivePrototypeAdhesionFactor_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotivePrototypeDrawBarPull_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotivePrototypeHorsePower_Click()

Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotivePrototypeTractionEffort_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub LocomotivePrototypeWeight_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub Picture1_Click()

    Let ButtonUpdate.Enabled = True

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub



Private Sub TemporaryTextBox_Change()

Let MainlineDiesel.LocomotiveNumberGO.Text = Format(LocomotiveNumber.Recordset("LocomotiveNumber"), "0000")
Let MainlineDiesel.LocomotiveNumber.Caption = Format(LocomotiveNumber.Recordset("LocomotiveNumber"), "0000")

End Sub
















Private Sub TimerInitialize_Timer()

    Let TimerInitialize.Enabled = False
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Locomotive List
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let MainlineDiesel.MousePointer = vbArrowHourglass
    Let ButtonClose.Enabled = False
    LocomotiveNumberGO.Clear
   
    LocomotiveDatabaseSearch.Refresh
    LocomotiveDatabaseSearch.Recordset.MoveFirst

    Do
    If LocomotiveDatabaseActiveSearch.Value = vbChecked Then
         LocomotiveNumberGO.AddItem Format(Val(LocomotiveDatabaseNumberSearch.Text), "0000")
    End If
    LocomotiveDatabaseSearch.Recordset.MoveNext
    Loop While Not LocomotiveDatabaseSearch.Recordset.EOF

    Let ButtonClose.Enabled = True
    Let MainlineDiesel.MousePointer = vbDefault
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Prototype Owners List
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

LocomotivePrototype.Clear
While Not ReportingMarksDatabase.Recordset.EOF
    Let temp = ReportingMarksDatabase.Recordset.Fields("ReportingMarks") + " (" + ReportingMarksDatabase.Recordset.Fields("RegisteredOwner") + ")"
    LocomotivePrototype.AddItem temp
    ReportingMarksDatabase.Recordset.MoveNext
Wend

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
        Else 'If menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainScreen.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "This combinaton box is used to display all the locomotive" & vbCrLf & "in your database. You can select a number or type in a" & vbCrLf & "new number."
        Let BalloonHelpText2 = "Locomotive (decoder) Number"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelStatus)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveNumberGO, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If
        

        Let BalloonHelpText1 = "This check box is used to change the database to short" & vbCrLf & "adresses (two digits), otherwise long addresses" & vbCrLf & "(four digits) will be used."
        Let BalloonHelpText2 = "Short Address"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckAddressShort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckAddressShort, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This combinaton box is used to display the prototype" & vbCrLf & "owner of the locomotive."
        Let BalloonHelpText2 = "Prototype Owner"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototype)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototype, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This check box flags the locomotive (number) as being" & vbCrLf & "active or used. Making the locomotive (number) inactive" & vbCrLf & "removes it from the database."
        Let BalloonHelpText2 = "Locomotive Active?"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveActiveList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveActiveList, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This check box flags the locomotive (number) as being" & vbCrLf & "equiped with a digital decoder (DCC). Making the locomotive " & vbCrLf & "(number) not equiped is for DC units."
        Let BalloonHelpText2 = "Locomotive Decoder Equiped?"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDecoderEquiped)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDecoderEquiped, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button makes the program search for the locomotive" & vbCrLf & "(number) in the database for editing."
        Let BalloonHelpText2 = "Go Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonGo)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonGo, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This combinaton box allows for the selection of the" & vbCrLf & "locomotive manufacturer."
        Let BalloonHelpText2 = "Locomotive Manufacturer"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveManufacturer)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveManufacturer, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This combinaton box allows for the selection of the" & vbCrLf & "locomotive model."
        Let BalloonHelpText2 = "Locomotive Model"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModel)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModel, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This combinaton box allows for the selection of the" & vbCrLf & "model's manufacturer."
        Let BalloonHelpText2 = "Locomotive Model Manufacturer"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelManufacturer)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelManufacturer, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This text box allows you to place additional information" & vbCrLf & "about the prototype locomotive or the model into your database."
        Let BalloonHelpText2 = "Locomotive Facts"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveFacts)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveFacts, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
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

Private Sub BalloonHelpUpdatePart02()

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
        Else 'If menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainScreen.MousePointer = ccHourglass
 
        Let BalloonHelpText1 = "This picture box allows to to show a picture of your" & vbCrLf & "model or the prototype locomotive."
        Let BalloonHelpText2 = "Locomotive Picture"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveImage)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveImage, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the current filename of the" & vbCrLf & "picture above."
        Let BalloonHelpText2 = "Locomotive Picture Filename"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveImageFilename)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveImageFilename, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button allows you to change the file name for" & vbCrLf & "this locomotive (number) in your database."
        Let BalloonHelpText2 = "Change the Picture"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonChangePicture)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonChangePicture, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This tab toggles grouped inforamtion about the" & vbCrLf & "locomotive prototype or model."
        Let BalloonHelpText2 = "Tabbed Inforamtion"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SSTab1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SSTab1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the horse power rating for the" & vbCrLf & "prototype locomotive."
        Let BalloonHelpText2 = "Horsepower"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototypeHorsePower)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototypeHorsePower, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the gross weight for the" & vbCrLf & "prototype locomotive."
        Let BalloonHelpText2 = "Weight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototypeWeight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototypeWeight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the adhesion factor for the" & vbCrLf & "prototype locomotive."
        Let BalloonHelpText2 = "Adhesion Factor"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototypeAdhesionFactor)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototypeAdhesionFactor, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the traction effort for the" & vbCrLf & "prototype locomotive."
        Let BalloonHelpText2 = "Traction Effort"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototypeTractionEffort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototypeTractionEffort, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the draw bar pull for the" & vbCrLf & "prototype locomotive."
        Let BalloonHelpText2 = "Draw Bar Pull"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotivePrototypeDrawBarPull)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotivePrototypeDrawBarPull, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
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

Private Sub BalloonHelpUpdatePart03()

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
        Else 'If menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainScreen.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "This textbox displays the horse power rating for the" & vbCrLf & "model locomotive."
        Let BalloonHelpText2 = "Horsepower"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelHorsePower)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelHorsePower, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the gross weight for the" & vbCrLf & "model locomotive."
        Let BalloonHelpText2 = "Weight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelWeight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelWeight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the adhesion factor for the" & vbCrLf & "model locomotive."
        Let BalloonHelpText2 = "Adhesion Factor"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelAdhesionFactor)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelAdhesionFactor, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the traction effort for the" & vbCrLf & "model locomotive."
        Let BalloonHelpText2 = "Traction Effort"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelTractionEffort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelTractionEffort, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the draw bar pull for the" & vbCrLf & "model locomotive."
        Let BalloonHelpText2 = "Draw Bar Pull"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveModelDrawBarPull)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveModelDrawBarPull, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If


        Let BalloonHelpText1 = "This combination box displays the couple used on" & vbCrLf & "the front of the model locomotive."
        Let BalloonHelpText2 = "Front Coupler"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBoxCouplerFront)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxCouplerFront, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This combination box displays the coupler used on" & vbCrLf & "the rear of the model locomotive."
        Let BalloonHelpText2 = "Rear Coupler"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBoxCouplerRear)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxCouplerRear, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
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

Private Sub BalloonHelpUpdatePart04()

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
        Else 'If menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainScreen.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "This button loads a new window for obtaining" & vbCrLf & "additional information about the selected prototype" & vbCrLf & "locomotive."
        Let BalloonHelpText2 = "Prototype Information Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrototypeInfo)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrototypeInfo, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button displays a new window for veiwing" & vbCrLf & "the decoder settings of  the selected locomotive (number)."
        Let BalloonHelpText2 = "Decoder Programming Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonDecoderProg)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonDecoderProg, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button displays a new window for obtaining" & vbCrLf & " a list of locomotives in your database."
        Let BalloonHelpText2 = "List Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonList, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button updates the database with any" & vbCrLf & "changes made in this window (Programming Diesel)."
        Let BalloonHelpText2 = "Update Buton"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonUpdate)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonUpdate, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button prints the current window to your printer."
        Let BalloonHelpText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "This button closes this window (Programming Diesel)."
        Let BalloonHelpText2 = "Close Buton"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Diesel Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
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
