VERSION 4.00
Begin VB.Form DefineBlocksProperties 
   Caption         =   "Form1"
   ClientHeight    =   10440
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   10845
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.TextBox TextboxWestboundBlocks 
      Height          =   285
      Left            =   5160
      TabIndex        =   17
      Text            =   "n/a"
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "DefineBlocksProperties.frx":0000
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox TextBoxLength 
      DataField       =   "Length"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox TextBoxTop 
      DataField       =   "PictureBoxTop"
      DataSource      =   "DatabaseBlockProperties"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox TextBoxLeft 
      DataField       =   "PictureBoxLeft"
      DataSource      =   "DatabaseBlockProperties"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox TextBoxFileName 
      DataField       =   "PictureBoxFileName"
      DataSource      =   "DatabaseBlockProperties"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox TextBoxObjectName 
      DataField       =   "PictureBoxName"
      DataSource      =   "DatabaseBlockProperties"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox TextBoxRecordCounter 
      DataField       =   "RecordCounter"
      DataSource      =   "DatabaseBlockProperties"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
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
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TrackPlan"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   $"DefineBlocksProperties.frx":0442
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
      DataField       =   "Name"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelLength 
      Alignment       =   1  'Right Justify
      Caption         =   "Length of Block"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LabelTop 
      Alignment       =   1  'Right Justify
      Caption         =   "Top Position on Track Map"
      DataField       =   "PictureBoxTop"
      DataSource      =   "DatabaseBlockProperties"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label LabelLeft 
      Alignment       =   1  'Right Justify
      Caption         =   "Left Position on Track Map"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label LabelFileName 
      Alignment       =   1  'Right Justify
      Caption         =   "File Name"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label LabelObjectName 
      Alignment       =   1  'Right Justify
      Caption         =   "Object Name"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label LabelRecordCounter 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Counter"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "DefineBlocksProperties"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
