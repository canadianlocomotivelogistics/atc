VERSION 4.00
Begin VB.Form PrototypeInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Prototype Information"
   ClientHeight    =   9090
   ClientLeft      =   9765
   ClientTop       =   3075
   ClientWidth     =   6690
   Height          =   9495
   Icon            =   "PrototypeInfo.frx":0000
   Left            =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   2730
   Width           =   6810
   Begin VB.CommandButton ButtonRecordAdd 
      Caption         =   "&Add New"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton ButtonRecordDelete 
      Caption         =   "&Delete"
      Height          =   255
      Left            =   5400
      TabIndex        =   33
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Data PrototypeInfoDatabase 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\LocomotiveDatabasePrototypeInfo.mdb"
      Exclusive       =   -1  'True
      Height          =   270
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PrototypeInfo"
      Top             =   7680
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "PrototypeInfo.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox PrototypeUnitsBuilt 
      DataField       =   "PrototypeUnitsBuilt"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Text            =   "Units Built"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox PrototypeDateManufactured 
      DataField       =   "PrototypeDateManufactured"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Text            =   "Date Manufactured"
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox PrototypeCylinders 
      DataField       =   "PrototypeCylinders"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "Cylinders"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox PrototypeLength 
      DataField       =   "PrototypeLength"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "Length"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Picture"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox PrototypeImageFilename 
      DataField       =   "PrototypeImageFilename"
      DataSource      =   "PrototypeInfoDatabase"
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox PrototypeDrawBarPull 
      DataField       =   "PrototypeDrawBarPull"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Text            =   "Draw Bar Pull"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox PrototypeTractionEffort 
      DataField       =   "PrototypeTractionEffort"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Text            =   "Traction Effort"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox PrototypeAdhesionFactor 
      DataField       =   "PrototypeAdhesionFactor"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Text            =   "Adhesion Factor"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox PrototypeWeight 
      DataField       =   "PrototypeWeight"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "Weight"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox PrototypeHorsePower 
      DataField       =   "PrototypeHorsePower"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "Horse Power"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox PrototypeFacts 
      DataField       =   "PrototypeFacts"
      DataSource      =   "PrototypeInfoDatabase"
      Height          =   5055
      Left            =   120
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "PrototypeInfo.frx":0884
      Top             =   2400
      Width           =   3015
   End
   Begin VB.ComboBox PrototypeModel 
      DataField       =   "PrototypeModel"
      DataSource      =   "PrototypeInfoDatabase"
      Height          =   315
      ItemData        =   "PrototypeInfo.frx":0896
      Left            =   120
      List            =   "PrototypeInfo.frx":0898
      TabIndex        =   3
      Text            =   "Prototype Model"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.ComboBox PrototypeManufacturer 
      DataField       =   "PrototypeManufacturer"
      DataSource      =   "PrototypeInfoDatabase"
      Height          =   315
      ItemData        =   "PrototypeInfo.frx":089A
      Left            =   120
      List            =   "PrototypeInfo.frx":08B6
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Prototype Manufacturer"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.PictureBox AlphaBlend 
      Height          =   480
      Left            =   7080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   30
      Top             =   1440
      Width           =   1200
   End
   Begin VB.PictureBox Ini 
      Height          =   480
      Left            =   7080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   31
      Top             =   840
      Width           =   1200
   End
   Begin VB.PictureBox BalloonHelp 
      Height          =   480
      Left            =   7080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   32
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label11 
      Caption         =   $"PrototypeInfo.frx":0965
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   8160
      Width           =   6495
   End
   Begin VB.Line Line4 
      X1              =   6600
      X2              =   120
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label LabelDetails 
      Caption         =   "Locomotive Details"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   3360
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label LabelFacts 
      Caption         =   "Locomotive Facts"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label10 
      Caption         =   "Units Built (number)"
      Height          =   195
      Left            =   4920
      TabIndex        =   26
      Top             =   7080
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "Date (year-year)"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Cylinders (number)"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Draw Bar Pull (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Traction Effort (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Adhesion (percentage)"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Weight (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Horse Power (hp)"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Length (feet, Inches)"
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "To search for a specific model:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label LabelPrototypeInfo 
      Caption         =   "Here to can edit the prototype information and place this information into your locomotive database."
      Height          =   495
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog PictureGet 
      Left            =   3360
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open an Image"
      Filter          =   "Picture Files (*.bmp)|*.bmp|Picture Files (*.gif)|*.gif|Picture Files (*.jpg)|*.jpg"
   End
   Begin VB.Image PrototypeImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "PrototypeInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub buttonAdoptInfo_Click()

End Sub

Private Sub ButtonClose_Click()

End

End Sub

Private Sub ButtonPrint_Click()

    MainlinePrototypeInfo.PrintForm
    
End Sub

Private Sub ButtonRecordAdd_Click()

    If ButtonRecordAdd.Caption = "&Add New" Then
        Let ButtonRecordAdd.Caption = "&Update"
        Let ButtonRecordDelete.Enabled = False
        PrototypeInfoDatabase.Recordset.AddNew
    ElseIf ButtonRecordAdd.Caption = "&Update" Then
        Let ButtonRecordAdd.Caption = "&Add New"
        Let ButtonRecordDelete.Enabled = True
        PrototypeInfoDatabase.Recordset.Update
        PrototypeInfoDatabase.Recordset.MoveLast
    End If
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub ButtonRecordMoveBack_Click()

    If Val(PrototypeInfoDatabase.Recordset.AbsolutePosition) > 1 Then
        PrototypeInfoDatabase.Recordset.AbsolutePosition = PrototypeInfoDatabase.Recordset.AbsolutePosition - 1
    End If
End Sub


Private Sub ButtonRecordMoveForward_Click()
    
    If Not PrototypeInfoDatabase.Recordset.EOF Then
        PrototypeInfoDatabase.Recordset.AbsolutePosition = PrototypeInfoDatabase.Recordset.AbsolutePosition + 1
    End If
End Sub


Private Sub ButtonRecordDelete_Click()

    PrototypeInfoDatabase.Recordset.Delete
    PrototypeInfoDatabase.Recordset.Refresh
    
End Sub

Private Sub Command3_Click()

    PictureGet.ShowOpen
    Let PrototypeImageFilename.Text = PictureGet.filename

End Sub

Private Sub Form_Load()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
     '   Dim TemporaryText1 As String
      '  Dim TemporaryText2 As String
       ' Dim i As Integer
        'Dim t As Boolean
        'Dim f As Boolean
        'Let t = True
        'Let f = False
'
 '       Let TemporaryText1 = "This button prints the current window to your printer."
  '      Let TemporaryText2 = "Print Button"
   '     i = BalloonHelp.DestroyToolTip(ButtonPrint.hWnd)
    '    i = BalloonHelp.AddToolTip(ButtonPrint.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

        'Let TemporaryText1 = "This text box is where all information from your" + vbCrLf + "serial port is displayed. Commands given by the" + vbCrLf + "program are displayed here. You can also type your" + vbCrLf + "own commands, providing the port is not busy."
        'Let TemporaryText2 = "Communication Window"
        'i = BalloonHelp.DestroyToolTip(TextBoxCommunicationWindowDCC.hWnd)
        'i = BalloonHelp.AddToolTip(TextBoxCommunicationWindowDCC.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

    'End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let PrototypeInfoDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabasePrototypeInfo.mdb"
    PrototypeInfoDatabase.Refresh

'    If MainlineDiesel.LocomotiveManufacturer.Text <> "Locomotive Manufacturer" Then
 '       Let PrototypeManufacturer.Text = MainlineDiesel!LocomotiveManufacturer.Text
 '   End If
  '  If MainlineDiesel.LocomotiveModel.Text <> "Locomotive Model" Then
   '     Let PrototypeModel.Text = MainlineDiesel!LocomotiveModel.Text
    'End If

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

Private Sub PrototypeImageFilename_Change()

    On Error Resume Next
    Let PROTOTYPEIMAGE.Picture = LoadPicture(PrototypeImageFilename.Text)
    If Err.Number = 53 Then
        PROTOTYPEIMAGE.Picture = LoadPicture()
    End If

End Sub



Private Sub PrototypeInfoDatabase_Reposition()
Let PrototypeInfoDatabase.Caption = PrototypeInfoDatabase.Recordset.AbsolutePosition
End Sub

Private Sub PrototypeManufacturer_Change()

If PrototypeManufacturer.Text = "American Locomotive Company" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "C-636"
    PrototypeModel.AddItem "C-630"
    PrototypeModel.AddItem "C-628"
    PrototypeModel.AddItem "C-430"
    PrototypeModel.AddItem "C-425"
    PrototypeModel.AddItem "C-424"
    PrototypeModel.AddItem "C-420"
    PrototypeModel.AddItem "RS36"
    PrototypeModel.AddItem "RS32"
    PrototypeModel.AddItem "RS27"
    PrototypeModel.AddItem "RSD15"
    PrototypeModel.AddItem "RSD12"
    PrototypeModel.AddItem "RD11"
    PrototypeModel.AddItem "RS3"
    PrototypeModel.AddItem "RS2"
    PrototypeModel.AddItem "RS1"
    PrototypeModel.AddItem "S-1"
    PrototypeModel.AddItem "S-2"
    PrototypeModel.AddItem "S-3"
    PrototypeModel.AddItem "S-4"
    PrototypeModel.AddItem "S-6"
    PrototypeModel.AddItem "T-6"
    PrototypeModel.AddItem "PA-3"
    PrototypeModel.AddItem "PA-2"
    PrototypeModel.AddItem "PA-1"
    PrototypeModel.AddItem "FA-2"
    PrototypeModel.AddItem "FA-1"
End If

If PrototypeManufacturer.Text = "Baldwin" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "VO 1000"
    PrototypeModel.AddItem "DRS-6-6-15 or AS-616"
    PrototypeModel.AddItem "DRS-6-4-15 or AS-416"
    PrototypeModel.AddItem "DRS-4-4-15 or AS-16"
    PrototypeModel.AddItem "DS-4-4-10"
    PrototypeModel.AddItem "S-12"
End If

If PrototypeManufacturer.Text = "Bombardier" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "HR616"
    End If
    
If PrototypeManufacturer.Text = "FairBanks-Morse" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "H-12-44"
End If

If PrototypeManufacturer.Text = "General Electric" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "AC6000"
    PrototypeModel.AddItem "AC4400CW"
    PrototypeModel.AddItem "9-44CW"
    PrototypeModel.AddItem "9-44C"
    PrototypeModel.AddItem "9-44BW"
    PrototypeModel.AddItem "9-44B"
    PrototypeModel.AddItem "9-40CW"
    PrototypeModel.AddItem "9-40C"
    PrototypeModel.AddItem "9-40BW"
    PrototypeModel.AddItem "9-40B"
    PrototypeModel.AddItem "8-41CW"
    PrototypeModel.AddItem "8-41C"
    PrototypeModel.AddItem "8-40CW"
    PrototypeModel.AddItem "8-40CM"
    PrototypeModel.AddItem "8-40C"
    PrototypeModel.AddItem "8-40BW"
    PrototypeModel.AddItem "8-40B"
    PrototypeModel.AddItem "8-39CE"
    PrototypeModel.AddItem "8-39C"
    PrototypeModel.AddItem "8-39B"
    PrototypeModel.AddItem "8-32C"
    PrototypeModel.AddItem "8-32BWH or P32-8BWH"
    PrototypeModel.AddItem "8-32B"
    PrototypeModel.AddItem "7-36B"
    PrototypeModel.AddItem "7-36C"
    PrototypeModel.AddItem "7-33C"
    PrototypeModel.AddItem "7-30B-A1"
    PrototypeModel.AddItem "7-30C-A"
    PrototypeModel.AddItem "7-30B-A"
    PrototypeModel.AddItem "7-30C"
    PrototypeModel.AddItem "7-30B"
    PrototypeModel.AddItem "7-23BQ"
    PrototypeModel.AddItem "7-23B"
    PrototypeModel.AddItem "U36C"
    PrototypeModel.AddItem "U36B"
    PrototypeModel.AddItem "U33C"
    PrototypeModel.AddItem "U33B"
    PrototypeModel.AddItem "SF30C"
    PrototypeModel.AddItem "U30C"
    PrototypeModel.AddItem "U30B"
    PrototypeModel.AddItem "U28B"
    PrototypeModel.AddItem "U25C"
    PrototypeModel.AddItem "U25B"
    PrototypeModel.AddItem "U23C"
    PrototypeModel.AddItem "U23B"
    PrototypeModel.AddItem "U18B"
End If

If PrototypeManufacturer.Text = "General Motors - Electromotive Division" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "SD90/43MAC"
    PrototypeModel.AddItem "SD80MAC"
    PrototypeModel.AddItem "SD80MC"
    PrototypeModel.AddItem "SD75M"
    PrototypeModel.AddItem "SD70MAC"
    PrototypeModel.AddItem "SD70M"
    PrototypeModel.AddItem "SD70I"
    PrototypeModel.AddItem "SD70"
    PrototypeModel.AddItem "F69PH-AC"
    PrototypeModel.AddItem "SD60MAC"
    PrototypeModel.AddItem "SD60M"
    PrototypeModel.AddItem "SD60F"
    PrototypeModel.AddItem "SD60I"
    PrototypeModel.AddItem "SD60"
    PrototypeModel.AddItem "SD50M"
    PrototypeModel.AddItem "SD50F"
    PrototypeModel.AddItem "SD50I"
    PrototypeModel.AddItem "SD50"
    PrototypeModel.AddItem "GP60M"
    PrototypeModel.AddItem "GP60"
    PrototypeModel.AddItem "AMD-103"
    PrototypeModel.AddItem "F59PHI"
    PrototypeModel.AddItem "F59PH"
    PrototypeModel.AddItem "GP59"
    PrototypeModel.AddItem "GP50"
    PrototypeModel.AddItem "DD40AX"
    PrototypeModel.AddItem "SD45-2T"
    PrototypeModel.AddItem "SD45-2"
    PrototypeModel.AddItem "SDP45"
    PrototypeModel.AddItem "FP45"
    PrototypeModel.AddItem "F45"
    PrototypeModel.AddItem "SD45"
    PrototypeModel.AddItem "SD40-2W"
    PrototypeModel.AddItem "SD40-2T"
    PrototypeModel.AddItem "SD40-2F"
    PrototypeModel.AddItem "SD40-2"
    PrototypeModel.AddItem "SDP40F"
    PrototypeModel.AddItem "SDP40"
    PrototypeModel.AddItem "F40PHM-2"
    PrototypeModel.AddItem "F40PH-2C"
    PrototypeModel.AddItem "F40PH"
    PrototypeModel.AddItem "F40C"
    PrototypeModel.AddItem "SD40"
    PrototypeModel.AddItem "GP40-2"
    PrototypeModel.AddItem "GP40X"
    PrototypeModel.AddItem "GP40P"
    PrototypeModel.AddItem "GP40W"
    PrototypeModel.AddItem "GP40"
    PrototypeModel.AddItem "GP39-2"
    PrototypeModel.AddItem "SDL39"
    PrototypeModel.AddItem "SD39"
    PrototypeModel.AddItem "GP39"
    PrototypeModel.AddItem "SD38-2"
    PrototypeModel.AddItem "GP39-2W"
    PrototypeModel.AddItem "GP38-2"
    PrototypeModel.AddItem "SD38"
    PrototypeModel.AddItem "GP38"
    PrototypeModel.AddItem "SDP35"
    PrototypeModel.AddItem "SD35"
    PrototypeModel.AddItem "GP35"
    PrototypeModel.AddItem "GP30"
    PrototypeModel.AddItem "SD24"
    PrototypeModel.AddItem "GP20"
    PrototypeModel.AddItem "GP15-1"
    PrototypeModel.AddItem "GP15T"
    PrototypeModel.AddItem "GP15"
    PrototypeModel.AddItem "SD18"
    PrototypeModel.AddItem "GP18"
    PrototypeModel.AddItem "SD9"
    PrototypeModel.AddItem "GP9"
    PrototypeModel.AddItem "SD7"
    PrototypeModel.AddItem "GP7"
    PrototypeModel.AddItem "CF7"
    PrototypeModel.AddItem "GMD1"
    PrototypeModel.AddItem "RS1325"
    PrototypeModel.AddItem "NM5"
    PrototypeModel.AddItem "NW2"
    PrototypeModel.AddItem "TR5"
    PrototypeModel.AddItem "TR4"
    PrototypeModel.AddItem "MP15T"
    PrototypeModel.AddItem "MP15AC"
    PrototypeModel.AddItem "MP15(DC)"
    PrototypeModel.AddItem "SW1504"
    PrototypeModel.AddItem "SW1500"
    PrototypeModel.AddItem "SW1200"
    PrototypeModel.AddItem "SW1001"
    PrototypeModel.AddItem "SW1000"
    PrototypeModel.AddItem "SW900"
    PrototypeModel.AddItem "SW600"
    PrototypeModel.AddItem "SW9"
    PrototypeModel.AddItem "SW8"
    PrototypeModel.AddItem "SW7"
    PrototypeModel.AddItem "SW1"
    PrototypeModel.AddItem "FL9"
    PrototypeModel.AddItem "FP9"
    PrototypeModel.AddItem "F9"
    PrototypeModel.AddItem "FP7"
    PrototypeModel.AddItem "F7"
    PrototypeModel.AddItem "F3"
    PrototypeModel.AddItem "E9A"
    PrototypeModel.AddItem "E8A"
    PrototypeModel.AddItem "BL2"

End If

If PrototypeManufacturer.Text = "Montreal Locomotive Works" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "S-13"
        PrototypeModel.AddItem "M636"
        PrototypeModel.AddItem "M630"
        PrototypeModel.AddItem "M420"
        PrototypeModel.AddItem "M420R"
End If

If PrototypeManufacturer.Text = "Morrison Knudsen" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "MK5000C"
        PrototypeModel.AddItem "MK1200G"
        PrototypeModel.AddItem "MK-F40PHL-2"
        PrototypeModel.AddItem "MKGP40FH-2"
End If

Let PrototypeManufacturer.SelStart = 0
Let PrototypeManufacturer.SelLength = 0

End Sub

Private Sub PrototypeManufacturer_Click()

If PrototypeManufacturer.Text = "American Locomotive Company" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "C-636"
    PrototypeModel.AddItem "C-630"
    PrototypeModel.AddItem "C-628"
    PrototypeModel.AddItem "C-430"
    PrototypeModel.AddItem "C-425"
    PrototypeModel.AddItem "C-424"
    PrototypeModel.AddItem "C-420"
    PrototypeModel.AddItem "RS36"
    PrototypeModel.AddItem "RS32"
    PrototypeModel.AddItem "RS27"
    PrototypeModel.AddItem "RSD15"
    PrototypeModel.AddItem "RSD12"
    PrototypeModel.AddItem "RD11"
    PrototypeModel.AddItem "RS3"
    PrototypeModel.AddItem "RS2"
    PrototypeModel.AddItem "RS1"
    PrototypeModel.AddItem "S-1"
    PrototypeModel.AddItem "S-2"
    PrototypeModel.AddItem "S-3"
    PrototypeModel.AddItem "S-4"
    PrototypeModel.AddItem "S-6"
    PrototypeModel.AddItem "T-6"
    PrototypeModel.AddItem "PA-3"
    PrototypeModel.AddItem "PA-2"
    PrototypeModel.AddItem "PA-1"
    PrototypeModel.AddItem "FA-2"
    PrototypeModel.AddItem "FA-1"
End If

If PrototypeManufacturer.Text = "Baldwin" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "VO 1000"
    PrototypeModel.AddItem "DRS-6-6-15 or AS-616"
    PrototypeModel.AddItem "DRS-6-4-15 or AS-416"
    PrototypeModel.AddItem "DRS-4-4-15 or AS-16"
    PrototypeModel.AddItem "DS-4-4-10"
    PrototypeModel.AddItem "S-12"
End If

If PrototypeManufacturer.Text = "Bombardier" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "HR616"
    End If
    
If PrototypeManufacturer.Text = "FairBanks-Morse" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "H-12-44"
End If

If PrototypeManufacturer.Text = "General Electric" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "AC6000"
    PrototypeModel.AddItem "AC4400CW"
    PrototypeModel.AddItem "9-44CW"
    PrototypeModel.AddItem "9-44C"
    PrototypeModel.AddItem "9-44BW"
    PrototypeModel.AddItem "9-44B"
    PrototypeModel.AddItem "9-40CW"
    PrototypeModel.AddItem "9-40C"
    PrototypeModel.AddItem "9-40BW"
    PrototypeModel.AddItem "9-40B"
    PrototypeModel.AddItem "8-41CW"
    PrototypeModel.AddItem "8-41C"
    PrototypeModel.AddItem "8-40CW"
    PrototypeModel.AddItem "8-40CM"
    PrototypeModel.AddItem "8-40C"
    PrototypeModel.AddItem "8-40BW"
    PrototypeModel.AddItem "8-40B"
    PrototypeModel.AddItem "8-39CE"
    PrototypeModel.AddItem "8-39C"
    PrototypeModel.AddItem "8-39B"
    PrototypeModel.AddItem "8-32C"
    PrototypeModel.AddItem "8-32BWH or P32-8BWH"
    PrototypeModel.AddItem "8-32B"
    PrototypeModel.AddItem "7-36B"
    PrototypeModel.AddItem "7-36C"
    PrototypeModel.AddItem "7-33C"
    PrototypeModel.AddItem "7-30B-A1"
    PrototypeModel.AddItem "7-30C-A"
    PrototypeModel.AddItem "7-30B-A"
    PrototypeModel.AddItem "7-30C"
    PrototypeModel.AddItem "7-30B"
    PrototypeModel.AddItem "7-23BQ"
    PrototypeModel.AddItem "7-23B"
    PrototypeModel.AddItem "U36C"
    PrototypeModel.AddItem "U36B"
    PrototypeModel.AddItem "U33C"
    PrototypeModel.AddItem "U33B"
    PrototypeModel.AddItem "SF30C"
    PrototypeModel.AddItem "U30C"
    PrototypeModel.AddItem "U30B"
    PrototypeModel.AddItem "U28B"
    PrototypeModel.AddItem "U25C"
    PrototypeModel.AddItem "U25B"
    PrototypeModel.AddItem "U23C"
    PrototypeModel.AddItem "U23B"
    PrototypeModel.AddItem "U18B"
End If

If PrototypeManufacturer.Text = "General Motors - Electromotive Division" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "SD90/43MAC"
    PrototypeModel.AddItem "SD80MAC"
    PrototypeModel.AddItem "SD80MC"
    PrototypeModel.AddItem "SD75M"
    PrototypeModel.AddItem "SD70MAC"
    PrototypeModel.AddItem "SD70M"
    PrototypeModel.AddItem "SD70I"
    PrototypeModel.AddItem "SD70"
    PrototypeModel.AddItem "F69PH-AC"
    PrototypeModel.AddItem "SD60MAC"
    PrototypeModel.AddItem "SD60M"
    PrototypeModel.AddItem "SD60F"
    PrototypeModel.AddItem "SD60I"
    PrototypeModel.AddItem "SD60"
    PrototypeModel.AddItem "SD50M"
    PrototypeModel.AddItem "SD50F"
    PrototypeModel.AddItem "SD50I"
    PrototypeModel.AddItem "SD50"
    PrototypeModel.AddItem "GP60M"
    PrototypeModel.AddItem "GP60"
    PrototypeModel.AddItem "AMD-103"
    PrototypeModel.AddItem "F59PHI"
    PrototypeModel.AddItem "F59PH"
    PrototypeModel.AddItem "GP59"
    PrototypeModel.AddItem "GP50"
    PrototypeModel.AddItem "DD40AX"
    PrototypeModel.AddItem "SD45-2T"
    PrototypeModel.AddItem "SD45-2"
    PrototypeModel.AddItem "SDP45"
    PrototypeModel.AddItem "FP45"
    PrototypeModel.AddItem "F45"
    PrototypeModel.AddItem "SD45"
    PrototypeModel.AddItem "SD40-2W"
    PrototypeModel.AddItem "SD40-2T"
    PrototypeModel.AddItem "SD40-2F"
    PrototypeModel.AddItem "SD40-2"
    PrototypeModel.AddItem "SDP40F"
    PrototypeModel.AddItem "SDP40"
    PrototypeModel.AddItem "F40PHM-2"
    PrototypeModel.AddItem "F40PH-2C"
    PrototypeModel.AddItem "F40PH"
    PrototypeModel.AddItem "F40C"
    PrototypeModel.AddItem "SD40"
    PrototypeModel.AddItem "GP40-2"
    PrototypeModel.AddItem "GP40X"
    PrototypeModel.AddItem "GP40P"
    PrototypeModel.AddItem "GP40W"
    PrototypeModel.AddItem "GP40"
    PrototypeModel.AddItem "GP39-2"
    PrototypeModel.AddItem "SDL39"
    PrototypeModel.AddItem "SD39"
    PrototypeModel.AddItem "GP39"
    PrototypeModel.AddItem "SD38-2"
    PrototypeModel.AddItem "GP39-2W"
    PrototypeModel.AddItem "GP38-2"
    PrototypeModel.AddItem "SD38"
    PrototypeModel.AddItem "GP38"
    PrototypeModel.AddItem "SDP35"
    PrototypeModel.AddItem "SD35"
    PrototypeModel.AddItem "GP35"
    PrototypeModel.AddItem "GP30"
    PrototypeModel.AddItem "SD24"
    PrototypeModel.AddItem "GP20"
    PrototypeModel.AddItem "GP15-1"
    PrototypeModel.AddItem "GP15T"
    PrototypeModel.AddItem "GP15"
    PrototypeModel.AddItem "SD18"
    PrototypeModel.AddItem "GP18"
    PrototypeModel.AddItem "SD9"
    PrototypeModel.AddItem "GP9"
    PrototypeModel.AddItem "SD7"
    PrototypeModel.AddItem "GP7"
    PrototypeModel.AddItem "CF7"
    PrototypeModel.AddItem "GMD1"
    PrototypeModel.AddItem "RS1325"
    PrototypeModel.AddItem "NM5"
    PrototypeModel.AddItem "NW2"
    PrototypeModel.AddItem "TR5"
    PrototypeModel.AddItem "TR4"
    PrototypeModel.AddItem "MP15T"
    PrototypeModel.AddItem "MP15AC"
    PrototypeModel.AddItem "MP15(DC)"
    PrototypeModel.AddItem "SW1504"
    PrototypeModel.AddItem "SW1500"
    PrototypeModel.AddItem "SW1200"
    PrototypeModel.AddItem "SW1001"
    PrototypeModel.AddItem "SW1000"
    PrototypeModel.AddItem "SW900"
    PrototypeModel.AddItem "SW600"
    PrototypeModel.AddItem "SW9"
    PrototypeModel.AddItem "SW8"
    PrototypeModel.AddItem "SW7"
    PrototypeModel.AddItem "SW1"
    PrototypeModel.AddItem "FL9"
    PrototypeModel.AddItem "FP9"
    PrototypeModel.AddItem "F9"
    PrototypeModel.AddItem "FP7"
    PrototypeModel.AddItem "F7"
    PrototypeModel.AddItem "F3"
    PrototypeModel.AddItem "E9A"
    PrototypeModel.AddItem "E8A"
    PrototypeModel.AddItem "BL2"

End If

If PrototypeManufacturer.Text = "Montreal Locomotive Works" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "S-13"
        PrototypeModel.AddItem "M636"
        PrototypeModel.AddItem "M630"
        PrototypeModel.AddItem "M420"
        PrototypeModel.AddItem "M420R"
End If

If PrototypeManufacturer.Text = "Morrison Knudsen" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "MK5000C"
        PrototypeModel.AddItem "MK1200G"
        PrototypeModel.AddItem "MK-F40PHL-2"
        PrototypeModel.AddItem "MKGP40FH-2"
End If

Let PrototypeManufacturer.SelStart = 0
Let PrototypeManufacturer.SelLength = 0

End Sub


