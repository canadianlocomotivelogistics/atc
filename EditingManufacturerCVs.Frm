VERSION 4.00
Begin VB.Form EditingManufacturerCVs 
   Caption         =   "Automatic Train Control - Editor for Manufacturer Specific CVs"
   ClientHeight    =   9135
   ClientLeft      =   4275
   ClientTop       =   3645
   ClientWidth     =   12960
   Height          =   9540
   Icon            =   "EditingManufacturerCVs.frx":0000
   Left            =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   12960
   Top             =   3300
   Width           =   13080
   Begin VB.TextBox TextboxNotes 
      DataField       =   "Notes"
      DataSource      =   "SpecificCVsDatabase"
      Height          =   735
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   59
      Text            =   "EditingManufacturerCVs.frx":0442
      Top             =   7320
      Width           =   12135
   End
   Begin VB.Data DecoderDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DecoderList"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   11640
      TabIndex        =   46
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Data SpecificCVsDatabase 
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\SpecificDecoderCVs.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CV1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit8Used 
      DataField       =   "Bit8Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit7Used 
      DataField       =   "Bit7Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   6480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      DataField       =   "Decoder"
      DataSource      =   "DecoderDatabase"
      Height          =   285
      Left            =   8520
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Manufacturer"
      DataSource      =   "DecoderDatabase"
      Height          =   285
      Left            =   3000
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ComboBox ComboBoxDecoderName 
      Height          =   315
      ItemData        =   "EditingManufacturerCVs.frx":0448
      Left            =   8520
      List            =   "EditingManufacturerCVs.frx":044A
      TabIndex        =   40
      Text            =   "No decoder selected"
      Top             =   720
      Width           =   3495
   End
   Begin VB.ComboBox ComboBoxManufacturerName 
      Height          =   315
      ItemData        =   "EditingManufacturerCVs.frx":044C
      Left            =   3000
      List            =   "EditingManufacturerCVs.frx":049E
      TabIndex        =   36
      Text            =   "No manufacturer selected"
      Top             =   720
      Width           =   3495
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit6Used 
      DataField       =   "Bit6Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit5Used 
      DataField       =   "Bit5Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit4Used 
      DataField       =   "Bit4Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit3Used 
      DataField       =   "Bit3Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox TextBoxSpecificCVbit8Off 
      DataField       =   "Bit8Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   31
      Text            =   "No Data"
      Top             =   6840
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit7Off 
      DataField       =   "Bit7Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   30
      Text            =   "No Data"
      Top             =   6480
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit6Off 
      DataField       =   "Bit6Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   29
      Text            =   "No Data"
      Top             =   6120
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit5Off 
      DataField       =   "Bit5Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   28
      Text            =   "No Data"
      Top             =   5760
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit4Off 
      DataField       =   "Bit4Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   27
      Text            =   "No Data"
      Top             =   5400
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit3Off 
      DataField       =   "Bit3Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   26
      Text            =   "No Data"
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit8On 
      DataField       =   "Bit8On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   25
      Text            =   "No Data"
      Top             =   6840
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit7On 
      DataField       =   "Bit7On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   24
      Text            =   "No Data"
      Top             =   6480
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit6On 
      DataField       =   "Bit6On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   23
      Text            =   "No Data"
      Top             =   6120
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit5On 
      DataField       =   "Bit5On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   22
      Text            =   "No Data"
      Top             =   5760
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit4On 
      DataField       =   "Bit4On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Text            =   "No Data"
      Top             =   5400
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "EditingManufacturerCVs.frx":06CF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox TextBoxSpecificCVbit3On 
      DataField       =   "Bit3On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Text            =   "No Data"
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit2Off 
      DataField       =   "Bit2Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   17
      Text            =   "No Data"
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit2On 
      DataField       =   "Bit2On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Text            =   "No Data"
      Top             =   4680
      Width           =   4695
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit2Used 
      DataField       =   "Bit2Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.TextBox TextBoxSpecificCVbit1Off 
      DataField       =   "Bit1Off"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   14
      Text            =   "No Data"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox TextBoxSpecificCVbit1On 
      DataField       =   "Bit1On"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "No Data"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.CheckBox CheckBoxSpecificCVbit1Used 
      DataField       =   "Bit1Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton ButtonAddNew 
      Caption         =   "&Add New"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CheckBox CheckBoxSpecificCVbitUsed 
      DataField       =   "BitUsed"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox TextBoxSpecificCVdescription 
      DataField       =   "Description"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Text            =   "No Data"
      Top             =   2760
      Width           =   5895
   End
   Begin VB.CheckBox CheckBoxSpecificCVused 
      DataField       =   "Used"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox TextBoxValueDefault 
      DataField       =   "ValueDefault"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   9720
      TabIndex        =   6
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox TextBoxValueMaximum 
      DataField       =   "ValueHigh"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Text            =   "255"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox TextBoxValueMinimum 
      DataField       =   "ValueLow"
      DataSource      =   "SpecificCVsDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Notes:"
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   7320
      Width           =   465
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   12840
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label16 
      Caption         =   "Is bit eight used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   57
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label15 
      Caption         =   "Is bit seven used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   56
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Is bit six used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   55
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label13 
      Caption         =   "Is bit five used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   54
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label12 
      Caption         =   "Is bit four used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   53
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "Is bit three used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   52
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Is bit two used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   51
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Is bit one used by the manufacturer?"
      Height          =   255
      Left            =   445
      TabIndex        =   50
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Is this configuration variable used by the manufacturer?"
      Height          =   255
      Left            =   480
      TabIndex        =   49
      Top             =   2400
      Width           =   3975
   End
   Begin IniconLib.Init Ini 
      Left            =   1200
      Top             =   8520
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
      Left            =   600
      Top             =   8520
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   120
      Top             =   8520
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Label Label6 
      Caption         =   "Is this configuration variable bit manipulative?"
      Height          =   255
      Left            =   445
      TabIndex        =   48
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "What is the general function of the configuration of the manufacturer specific CV?"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   12840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "with a decoder name of "
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "with a decoder name of "
      Height          =   315
      Left            =   6600
      TabIndex        =   39
      Top             =   720
      Width           =   1710
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "In the database a manufacturer was"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Search for Manufacturer"
      Height          =   255
      Left            =   1080
      TabIndex        =   37
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   12840
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   12840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label3 
      Caption         =   $"EditingManufacturerCVs.frx":0B11
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   120
      Width           =   12015
   End
   Begin VB.Label LabelValueDefault 
      Caption         =   "; with a default value for CV is"
      Height          =   195
      Index           =   49
      Left            =   7560
      TabIndex        =   5
      Top             =   3480
      Width           =   2100
   End
   Begin VB.Label LabelValueMaximum 
      Caption         =   "; the maximum value for CV is"
      Height          =   195
      Index           =   49
      Left            =   4560
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label LabelValueMinimum 
      Caption         =   "For this configuration the minimum value for CV is"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   128
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV1"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 1."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV2"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 2."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV3"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 3."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV4"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 4."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV5"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 5."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV6"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 6."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV7"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 7."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV8"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 8."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV9"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 9."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV10"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 10."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV11"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 11."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV12"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 12."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV13"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 13."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV14"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 14."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV15"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 15."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV16"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 16"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV17"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 17."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV18"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 18."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV19"
            Object.Tag             =   ""
            ToolTipText     =   "Select this tab to edit configuration variable 19."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV20"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab21 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV21"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab22 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV22"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab23 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV23"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab24 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV24"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab25 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV25"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab26 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV26"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab27 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV27"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab28 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV28"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab29 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV29"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab30 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV30"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab31 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV31"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab32 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV32"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab33 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV33"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab34 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV34"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab35 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV35"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab36 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV36"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab37 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV37"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab38 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV38"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab39 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV39"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab40 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV40"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab41 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV41"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab42 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV42"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab43 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV43"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab44 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV44"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab45 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV45"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab46 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV46"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab47 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV47"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab48 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV48"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab49 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV49"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab50 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV50"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab51 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV51"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab52 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV52"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab53 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV53"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab54 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV54"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab55 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV55"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab56 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV56"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab57 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV57"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab58 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV58"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab59 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV59"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab60 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV60"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab61 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV61"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab62 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV62"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab63 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV63"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab64 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV64"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab65 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV65"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab66 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV66"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab67 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV67"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab68 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV68"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab69 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV69"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab70 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV70"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab71 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV71"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab72 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV72"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab73 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV73"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab74 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV74"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab75 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV75"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab76 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV76"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab77 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV77"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab78 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV78"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab79 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV79"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab80 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV80"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab81 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV81"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab82 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV82"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab83 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV83"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab84 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV84"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab85 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV85"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab86 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV86"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab87 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV87"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab88 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV88"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab89 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV89"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab90 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV90"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab91 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV91"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab92 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV92"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab93 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV93"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab94 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV94"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab95 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV95"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab96 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV96"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab97 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV97"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab98 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV98"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab99 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV99"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab100 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV100"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab101 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV101"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab102 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV102"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab103 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV103"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab104 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV104"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab105 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV105"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab106 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV106"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab107 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV107"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab108 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV108"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab109 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV109"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab110 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV110"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab111 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV111"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab112 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV112"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab113 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV113"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab114 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV114"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab115 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV115"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab116 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV116"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab117 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV117"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab118 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV118"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab119 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV119"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab120 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV120"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab121 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV121"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab122 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV122"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab123 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV123"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab124 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV124"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab125 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV125"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab126 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV126"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab127 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV127"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab128 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CV128"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "EditingManufacturerCVs"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub SSTab1_DblClick()

End Sub


Private Sub CheckSpecificCVused_Click()


End Sub


Private Sub ButtonAddNew_Click()

For temp = 1 To 128
    Let temporary$ = "CV" & CStr(temp)
    SpecificCVsDatabase.RecordSource = temporary$
    SpecificCVsDatabase.Refresh
    SpecificCVsDatabase.Recordset.AddNew
    SpecificCVsDatabase.Recordset.Update
Next temp

DecoderDatabase.Recordset.AddNew
DecoderDatabase.Recordset.Update

DecoderDatabase.Recordset.MoveLast
Call TabStrip_Click


End Sub

Private Sub ButtonClear_Click()

End Sub

Private Sub ButtonClose_Click()

    End

End Sub

Private Sub ButtonUpdate_Click()
    
    Let CheckBoxSpecificCVused.Enabled = False
    Call CheckBoxSpecificCVused_Click
    Let ButtonClose.Enabled = True
    Let ButtonUpdate.Enabled = False
    SpecificCVsDatabase.Recordset.Update
    Let TabStrip.Enabled = False
    
End Sub

Private Sub CheckBoxSpecificCVbit1Used_Click()
    
    If CheckBoxSpecificCVbit1Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit1Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit1On.Enabled = True
            Let TextBoxSpecificCVbit1Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit1On.Enabled = False
            Let TextBoxSpecificCVbit1Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit1On.Enabled = False
        Let TextBoxSpecificCVbit1Off.Enabled = False
    End If

End Sub

Private Sub CheckBoxSpecificCVbit1Used_GotFocus()
            
        Let ButtonUpdate.Enabled = True
        Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVbit2Used_Click()
    
    If CheckBoxSpecificCVbit2Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit2Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit2On.Enabled = True
            Let TextBoxSpecificCVbit2Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit2On.Enabled = False
            Let TextBoxSpecificCVbit2Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit2On.Enabled = False
        Let TextBoxSpecificCVbit2Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit2Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVbit3Used_Click()
    
    If CheckBoxSpecificCVbit3Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit3Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit3On.Enabled = True
            Let TextBoxSpecificCVbit3Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit3On.Enabled = False
            Let TextBoxSpecificCVbit3Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit3On.Enabled = False
        Let TextBoxSpecificCVbit3Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit3Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVbit4Used_Click()
    
    If CheckBoxSpecificCVbit4Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit4Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit4On.Enabled = True
            Let TextBoxSpecificCVbit4Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit4On.Enabled = False
            Let TextBoxSpecificCVbit4Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit4On.Enabled = False
        Let TextBoxSpecificCVbit4Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit4Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVbit5Used_Click()
    
    If CheckBoxSpecificCVbit5Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit5Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit5On.Enabled = True
            Let TextBoxSpecificCVbit5Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit5On.Enabled = False
            Let TextBoxSpecificCVbit5Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit5On.Enabled = False
        Let TextBoxSpecificCVbit5Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit5Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVbit6Used_Click()
    
    If CheckBoxSpecificCVbit6Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit6Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit6On.Enabled = True
            Let TextBoxSpecificCVbit6Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit6On.Enabled = False
            Let TextBoxSpecificCVbit6Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit6On.Enabled = False
        Let TextBoxSpecificCVbit6Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit6Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False


End Sub


Private Sub CheckBoxSpecificCVbit7Used_Click()
    
    If CheckBoxSpecificCVbit7Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit7Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit7On.Enabled = True
            Let TextBoxSpecificCVbit7Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit7On.Enabled = False
            Let TextBoxSpecificCVbit7Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit7On.Enabled = False
        Let TextBoxSpecificCVbit7Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit7Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False


End Sub


Private Sub CheckBoxSpecificCVbit8Used_Click()
    
    If CheckBoxSpecificCVbit8Used.Enabled = True Then
        
        If CheckBoxSpecificCVbit8Used.Value = vbChecked Then
            Let TextBoxSpecificCVbit8On.Enabled = True
            Let TextBoxSpecificCVbit8Off.Enabled = True
        Else
            Let TextBoxSpecificCVbit8On.Enabled = False
            Let TextBoxSpecificCVbit8Off.Enabled = False
        End If
    Else
        Let TextBoxSpecificCVbit8On.Enabled = False
        Let TextBoxSpecificCVbit8Off.Enabled = False
    End If

End Sub


Private Sub CheckBoxSpecificCVbit8Used_GotFocus()
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False


End Sub


Private Sub CheckBoxSpecificCVbitUsed_Click()

    If CheckBoxSpecificCVbitUsed.Enabled = True Then
        
        If CheckBoxSpecificCVbitUsed.Value = vbChecked Then
            Let CheckBoxSpecificCVbit1Used.Enabled = True
            Let CheckBoxSpecificCVbit2Used.Enabled = True
            Let CheckBoxSpecificCVbit3Used.Enabled = True
            Let CheckBoxSpecificCVbit4Used.Enabled = True
            Let CheckBoxSpecificCVbit5Used.Enabled = True
            Let CheckBoxSpecificCVbit6Used.Enabled = True
            Let CheckBoxSpecificCVbit7Used.Enabled = True
            Let CheckBoxSpecificCVbit8Used.Enabled = True
        Else
            Let CheckBoxSpecificCVbit1Used.Enabled = False
            Let CheckBoxSpecificCVbit2Used.Enabled = False
            Let CheckBoxSpecificCVbit3Used.Enabled = False
            Let CheckBoxSpecificCVbit4Used.Enabled = False
            Let CheckBoxSpecificCVbit5Used.Enabled = False
            Let CheckBoxSpecificCVbit6Used.Enabled = False
            Let CheckBoxSpecificCVbit7Used.Enabled = False
            Let CheckBoxSpecificCVbit8Used.Enabled = False
        End If
    Else
        Let CheckBoxSpecificCVbit1Used.Enabled = False
        Let CheckBoxSpecificCVbit2Used.Enabled = False
        Let CheckBoxSpecificCVbit3Used.Enabled = False
        Let CheckBoxSpecificCVbit4Used.Enabled = False
        Let CheckBoxSpecificCVbit5Used.Enabled = False
        Let CheckBoxSpecificCVbit6Used.Enabled = False
        Let CheckBoxSpecificCVbit7Used.Enabled = False
        Let CheckBoxSpecificCVbit8Used.Enabled = False
    End If
        Call CheckBoxSpecificCVbit1Used_Click
        Call CheckBoxSpecificCVbit2Used_Click
        Call CheckBoxSpecificCVbit3Used_Click
        Call CheckBoxSpecificCVbit4Used_Click
        Call CheckBoxSpecificCVbit5Used_Click
        Call CheckBoxSpecificCVbit6Used_Click
        Call CheckBoxSpecificCVbit7Used_Click
        Call CheckBoxSpecificCVbit8Used_Click

End Sub

Private Sub CheckBoxSpecificCVbitUsed_GotFocus()
            
        Let ButtonUpdate.Enabled = True
        Let ButtonClose.Enabled = False

End Sub


Private Sub CheckBoxSpecificCVused_Click()

    If CheckBoxSpecificCVused.Enabled = True Then
        If CheckBoxSpecificCVused.Value = vbChecked Then
            Let TextBoxSpecificCVdescription.Enabled = True
            Let TextBoxValueMinimum.Enabled = True
            Let TextBoxValueMaximum.Enabled = True
            Let TextBoxValueDefault.Enabled = True
            Let CheckBoxSpecificCVbitUsed.Enabled = True
        Else
            Let TextBoxSpecificCVdescription.Enabled = False
            Let TextBoxValueMinimum.Enabled = False
            Let TextBoxValueMaximum.Enabled = False
            Let TextBoxValueDefault.Enabled = False
            Let CheckBoxSpecificCVbitUsed.Enabled = False
        End If
    Else
        Let CheckBoxSpecificCVbitUsed.Enabled = False
    End If
    
    Call CheckBoxSpecificCVbitUsed_Click

End Sub

Private Sub CheckBoxSpecificCVused_GotFocus()
            
        Let ButtonUpdate.Enabled = True
        Let ButtonClose.Enabled = False

End Sub


Private Sub ComboBoxDecoderName_Click()

Let MousePointer = vbArrowHourglass

DecoderDatabase.Recordset.MoveFirst
SpecificCVsDatabase.Recordset.MoveFirst

Do While Not (DecoderDatabase.Recordset.Fields("Manufacturer") = ComboBoxManufacturerName.Text And DecoderDatabase.Recordset.Fields("Decoder") = ComboBoxDecoderName.Text)
    DecoderDatabase.Recordset.MoveNext
    SpecificCVsDatabase.Recordset.MoveNext
Loop

Call TabStrip_Click
Let TabStrip.Enabled = True

Let MousePointer = Default

End Sub


Private Sub ComboBoxManufacturerName_Click()

Let MousePointer = vbArrowHourglass

DecoderDatabase.Recordset.MoveFirst

ComboBoxDecoderName.Clear

While Not DecoderDatabase.Recordset.EOF

    If DecoderDatabase.Recordset.Fields("Manufacturer") = ComboBoxManufacturerName.Text Then
        Let temp = DecoderDatabase.Recordset.Fields("decoder")
        ComboBoxDecoderName.AddItem temp
    End If
    DecoderDatabase.Recordset.MoveNext
Wend

Let MousePointer = Default

End Sub


Private Sub DecoderDatabase_Reposition()

    Let DecoderDatabase.Caption = DecoderDatabase.Recordset.AbsolutePosition
End Sub

Private Sub Form_Activate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transparency Screen Delay
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "Transparency"
    Dim TemporaryTransparency As String
    Let TemporaryTransparency = Ini.Value
    If TemporaryTransparency = "On" Then
        Let AlphaBlend.Enabled = True
        Let Ini.Parameter = "Opacity"
        Dim TemporaryOpacity As String
        Let TemporaryOpacity = Ini.Value
        Dim TemporaryScreenDelay As String
        Let temporaryscreenselay = Ini.Value
        Dim OutsideLoop As Integer
        Dim InsideLoop As Integer
        For OutsideLoop = 0 To Val(TemporaryOpacity)
            Let AlphaBlend.Opacity = OutsideLoop
            For InsideLoop = 0 To Val(TemporaryScreenDelay)
                DoEvents
            Next InsideLoop
        Next OutsideLoop
    ElseIf TemporaryTransparency = "Off" Then
       Let AlphaBlend.Enabled = False
    Else
        If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
            Stop
        End If
    End If
    
End Sub

Private Sub Form_Load()

' =============================================================================================================================================================================
' Checking the Screen Resolution
' =============================================================================================================================================================================
'
'   Every time a new window is opened in Autoamtic Train Control we check the screen size and compare it to the window screen size.
' If the window cannot be displayed in the current screen size a message box is displayed. This allows time for the user to change
' the screen attributes to correct size.

Do
    If Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY Then
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    End If
Loop While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY

' =============================================================================================================================================================================
' Initialization of Screen
' =============================================================================================================================================================================

    Let Ini.filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Editor Manufacturer Specific CVs"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value

' =============================================================================================================================================================================
' Positioning the Screen
' =============================================================================================================================================================================
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Firest, check to see if any values are present in the ini file, if so, use then, otherwise center the screen.
    
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        EditingManufacturerCVs.Left = (Screen.Width - Width) / 2   ' Center form horizontally.
        EditingManufacturerCVs.Top = (Screen.Height - Height) / 2  ' Center form vertiCally.
    Else
    
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Check to see if the window is off the screen, width wise.
        
        If Val(TemporaryValueLeft) + EditingManufacturerCVs.Width > Screen.Width Then
            Let EditingManufacturerCVs.Left = Screen.Width - EditingManufacturerCVs.Width
        Else
            Let EditingManufacturerCVs.Left = Val(TemporaryValueLeft)
        End If
        
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Check to see if the window is off the screen, height wise.
        
        If Val(TemporaryValueTop) + EditingManufacturerCVs.Height > Screen.Height Then
            Let EditingManufacturerCVs.Top = Screen.Height - EditingManufacturerCVs.Height
        Else
            Let EditingManufacturerCVs.Top = Val(TemporaryValueTop)
        End If
    End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
    '    Let AlphaBlend.Enabled = True
    'ElseIf MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
    '    Let AlphaBlend.Enabled = False
    'End If
' =============================================================================================================================================================================
' Adding Balloons
' =============================================================================================================================================================================
    'If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
        'Dim TemporaryText1 As String
        'Dim TemporaryText2 As String
        'Dim i As Integer
        'Dim t As Boolean
        'Dim f As Boolean
        'Let t = True
        'Let f = False

        'Let TemporaryText1 = "This text box is where all information from your" + vbCrLf + "serial port is displayed. Commands given by the" + vbCrLf + "program are displayed here. You can also type your" + vbCrLf + "own commands, providing the port is not busy."
        'Let TemporaryText2 = "Communication Window"
        'i = BalloonHelp.DestroyToolTip(TextBoxCommunicationWindowDCC.hWnd)
        'i = BalloonHelp.AddToolTip(TextBoxCommunicationWindowDCC.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)
    'End If
' =============================================================================================================================================================================
' Defining Databases and files
' =============================================================================================================================================================================

    DecoderDatabase.DatabaseName = App.Path & "\Databases\SpecificDecoderCVs.mdb"
    SpecificCVsDatabase.DatabaseName = App.Path & "\Databases\SpecificDecoderCVs.mdb"
    
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub SpecificCVsDatabase_Reposition()

    Let SpecificCVsDatabase.Caption = SpecificCVsDatabase.Recordset.AbsolutePosition

End Sub

Private Sub TabStrip_Click()
    
    Let SpecificCVsDatabase.RecordSource = TabStrip.SelectedItem
    SpecificCVsDatabase.Refresh
    
    SpecificCVsDatabase.Recordset.MoveFirst
    While SpecificCVsDatabase.Recordset.AbsolutePosition <> DecoderDatabase.Recordset.AbsolutePosition And Not SpecificCVsDatabase.Recordset.EOF
        SpecificCVsDatabase.Recordset.MoveNext
    Wend
    
    CheckBoxSpecificCVused.Enabled = True
    Call CheckBoxSpecificCVused_Click
    SpecificCVsDatabase.Recordset.Edit
    
    ButtonClose.Enabled = False
    ButtonUpdate.Enabled = True

End Sub


Private Sub TextBoxSpecificCVbit1Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit1On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit2Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit2On_LinkError(LinkErr As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit3Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit3On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit4Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit4On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit5Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit5On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit6Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit6On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit7Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit7On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit8Off_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVbit8On_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxSpecificCVdescription_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxValueDefault_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxValueMaximum_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


Private Sub TextBoxValueMinimum_KeyPress(KeyAscii As Integer)
        
    Let ButtonUpdate.Enabled = True
    Let ButtonClose.Enabled = False

End Sub


