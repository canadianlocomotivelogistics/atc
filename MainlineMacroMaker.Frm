VERSION 4.00
Begin VB.Form MainlineMacroMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline  Mode - Macro Maker"
   ClientHeight    =   7800
   ClientLeft      =   510
   ClientTop       =   5520
   ClientWidth     =   10905
   Height          =   8205
   Icon            =   "MainlineMacroMaker.frx":0000
   Left            =   450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10905
   Top             =   5175
   Width           =   11025
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   8280
      TabIndex        =   79
      Top             =   7320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   150
      Picture         =   "MainlineMacroMaker.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   78
      Top             =   135
      Width           =   480
   End
   Begin VB.TextBox ConsistControlComment 
      DataField       =   "ConsistControlComment"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      TabIndex        =   74
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8880
      Top             =   7320
   End
   Begin VB.Data LocomotiveDatabaseSearch 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBase"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data LocomotiveDatabaseConsistControl 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomoriveDatabaseConsist"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "Changing Configuration Variables"
      Height          =   1575
      Left            =   5640
      TabIndex        =   62
      Top             =   5640
      Width           =   5175
      Begin VB.CheckBox ConsistControlCVWright 
         Caption         =   "Wright Configuration Variable"
         Height          =   255
         Left            =   1200
         TabIndex        =   71
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox ConsistControlCVRead 
         Caption         =   "Read or"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox ConsistControlCVValue 
         Height          =   285
         Left            =   2400
         TabIndex        =   69
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox ConsistControlCV 
         Height          =   285
         Left            =   240
         TabIndex        =   68
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   63
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   73
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Configuration Variable"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sending out the Command"
      Height          =   1215
      Left            =   240
      TabIndex        =   58
      Top             =   5280
      Width           =   4935
      Begin VB.CommandButton Command8 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlSendCS 
         Caption         =   "Command Station"
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox CommandStationSendCF 
         Caption         =   "Send Commands to File and/or to"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Function Controls - Group Two"
      Height          =   1095
      Left            =   5640
      TabIndex        =   52
      Top             =   4440
      Width           =   5175
      Begin VB.CheckBox ConsistControlFunction8 
         Caption         =   "Function 8"
         Height          =   255
         Left            =   3840
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction7 
         Caption         =   "Function 7"
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction6 
         Caption         =   "Function 6"
         Height          =   255
         Left            =   1440
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction5 
         Caption         =   "Function 5"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Locomotive Number or Consist"
      Height          =   1215
      Left            =   5640
      TabIndex        =   48
      Top             =   240
      Width           =   5175
      Begin VB.CheckBox ShortAddress 
         Caption         =   "Short Address"
         Height          =   255
         Left            =   2040
         TabIndex        =   76
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox LocomotiveDatabaseDecoderSearch 
         Caption         =   "Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "LocomotiveDatabaseSearch"
         Height          =   255
         Left            =   1680
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox LocomotiveDatabaseNumberSearch 
         DataField       =   "LocomotiveNumber"
         DataSource      =   "LocomotiveDatabaseSearch"
         Height          =   285
         Left            =   360
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox ConsistControlMacroLocomotiveNumber 
         Height          =   315
         Left            =   360
         TabIndex        =   49
         Text            =   "4805"
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function Controls - Group One"
      Height          =   1095
      Left            =   5640
      TabIndex        =   41
      Top             =   3240
      Width           =   5175
      Begin VB.CheckBox ConsistControlFunction4 
         Caption         =   "Function 4"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction3 
         Caption         =   "Function 3"
         Height          =   255
         Left            =   3840
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction2 
         Caption         =   "Function 2"
         Height          =   255
         Left            =   2640
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction1 
         Caption         =   "Function 1"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlFunction0 
         Caption         =   "Function 0"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed & Direction Controls"
      Height          =   1575
      Left            =   5640
      TabIndex        =   34
      Top             =   1560
      Width           =   5175
      Begin VB.CommandButton Command5 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   47
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlDirectionR 
         Caption         =   "Reversed Direction"
         Height          =   255
         Left            =   1320
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox ConsistControlDirectionF 
         Caption         =   "Forward or "
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ConsistControlSpeed128 
         Caption         =   "128 Step Sped Mode"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox ConsistControlSpeed28 
         Caption         =   "28 Step or"
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ConsistControlSpeed14 
         Caption         =   "14 Step,"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.HScrollBar ConsistControlSpeed 
         Height          =   255
         Left            =   240
         Max             =   15
         TabIndex        =   35
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label ConsistControlSpeedValue 
         Caption         =   "0"
         Height          =   255
         Left            =   4440
         TabIndex        =   67
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox TimeDelay 
      Alignment       =   2  'Center
      DataField       =   "Time Delay"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox SevenByteD 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox SevenByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   23
      Text            =   "00"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox SevenByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   22
      Text            =   "00000000"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox SixByteD 
      Alignment       =   2  'Center
      DataField       =   "FourthByte"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox SixByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   20
      Text            =   "00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox SixByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   19
      Text            =   "00000000"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox FiveByteD 
      Alignment       =   2  'Center
      DataField       =   "ThirdByte"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox FiveByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Text            =   "00"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox FiveByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Text            =   "00000000"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox FourByteD 
      Alignment       =   2  'Center
      DataField       =   "SecondByte"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox FourByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Text            =   "00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox FourByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   13
      Text            =   "00000000"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteD 
      Alignment       =   2  'Center
      DataField       =   "FirstByte"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Text            =   "00"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox ThreeByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Text            =   "00000000"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TwoByteD 
      Alignment       =   2  'Center
      DataField       =   "AddressByteTwo"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox TwoByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox TwoByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Text            =   "00000000"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox OneByteD 
      Alignment       =   2  'Center
      DataField       =   "AddressByteOne"
      DataSource      =   "LocomotiveDatabaseConsistControl"
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox OneByteH 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "00"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox OneByteB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "00000000"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marco File Name"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "&Open"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox ConsistControlFilename 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   9600
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7620
      Top             =   7380
      _ExtentX        =   873
      _ExtentY        =   556
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   11160
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   11160
      Top             =   720
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label2 
      Caption         =   $"MainlineMacroMaker.frx":0884
      Height          =   915
      Left            =   810
      TabIndex        =   77
      Top             =   120
      Width           =   4365
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments"
      Height          =   255
      Left            =   240
      TabIndex        =   75
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label LocomotiveCommunicationStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   360
      TabIndex        =   66
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5400
      Y1              =   840
      Y2              =   7440
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Time Delay"
      Height          =   195
      Left            =   360
      TabIndex        =   33
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Error Byte"
      Height          =   195
      Left            =   480
      TabIndex        =   32
      Top             =   3840
      Width           =   690
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sixth Byte"
      Height          =   195
      Left            =   480
      TabIndex        =   31
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fifth Byte"
      Height          =   195
      Left            =   480
      TabIndex        =   30
      Top             =   3120
      Width           =   660
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fourth Byte"
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Third Byte"
      Height          =   195
      Left            =   360
      TabIndex        =   28
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address Byte"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   1680
      Width           =   930
   End
   Begin MSComDlg.CommonDialog ConsistControlDialog 
      Left            =   8280
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open a new Consist Control"
      Filter          =   "Control Files|*.mdb"
      InitDir         =   "c:\Automatic Train Control\ConsistControlDatabases\"
   End
End
Attribute VB_Name = "MainlineMacroMaker"
Attribute VB_Creatable = False
Attribute VB_Exposed = False







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
            If TemporaryScreen = "Mainline Macro Maker Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Macro Maker Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Macro Maker Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Macro Maker Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonPrint_Click()

    MainlineMacroMaker.PrintForm
    
End Sub

Private Sub Command3_Click()
 
If Command3.Caption = "&Open" Then

' -------------------------------------------------------------------------------------------------------------------------
' Initial Directory
'
    ConsistControlDialog.InitDir = App.Path$ & "\Databases\"

    ConsistControlDialog.ShowOpen
    ConsistControlFilename.Text = ConsistControlDialog.Filename
    LocomotiveDatabaseConsistControl.DatabaseName = ConsistControlDialog.Filename
    LocomotiveDatabaseConsistControl.Refresh
    
    Let Command3.Caption = "&Stop"
    LocomotiveDatabaseConsistControl.Recordset.MoveFirst
Else
    Let Command3.Caption = "&Open"
End If

End Sub

Private Sub Command4_Click()

Let ConsistControlComment.Text = ConsistControlComment.Text + "Function "

Let TemporaryByte = 128

If ConsistControlFunction0.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 16
    Let ConsistControlComment.Text = ConsistControlComment.Text + "0 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "0 Off;"
End If

If ConsistControlFunction1.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 1
    Let ConsistControlComment.Text = ConsistControlComment.Text + "1 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "1 Off;"
End If

If ConsistControlFunction2.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 2
    Let ConsistControlComment.Text = ConsistControlComment.Text + "2 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "2 Off;"
End If

If ConsistControlFunction3.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 4
    Let ConsistControlComment.Text = ConsistControlComment.Text + "3 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "3 Off;"
End If

If ConsistControlFunction4.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 8
    Let ConsistControlComment.Text = ConsistControlComment.Text + "4 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "4 Off;"
End If

Let ThreeByteD.Text = TemporaryByte
Let FourByteD.Text = ""
Let FiveByteD.Text = ""
Let SixByteD.Text = ""


End Sub


Private Sub Command5_Click()

Let ConsistControlComment.Text = ConsistControlComment.Text + "Speed "

If ConsistControlSpeed128.Value = vbChecked Then
    ' This routine assembles the byte for speed step mode 128
    Let Temporary = Val(ConsistControlSpeed.Value)
    Let ConsistControlComment.Text = ConsistControlComment.Text + Str$(Temporary) + " of 128 "
    
    If ConsistControlDirectionF.Value = vbChecked Then
        Temporary = Temporary + 128 ' add forward direction
        Let ConsistControlComment.Text = ConsistControlComment.Text + "Forward"
    Else
        Let ConsistControlComment.Text = ConsistControlComment.Text + "Reverse"
    End If
    
    Let ThreeByteD.Text = 63
    Let FourByteD.Text = Temporary
    Let FiveByteD.Text = ""
    Let SixByteD.Text = ""
Else
    Let Temporary = 64
    If ConsistControlDirectionF.Value = vbChecked Then
        Let Temporary = Temporary + 32 ' add forward direction
    End If
    
    If ConsistControlSpeed28.Value = vbChecked Then
        'This routine assenmles the byte for speed step mode 28
        Let temp1 = Val(ConsistControlSpeed.Value) ' adds the speed
        Let temp2 = temp1 Mod 2
        Let newspeedvalue = Int(temp1 / 2)
        Let Temporary = Temporary + newspeedvalue
        If temp2 = 1 Then Let Temporary = Temporary + 16
        Let ThreeByteD.Text = Temporary
        Let FourByteD.Text = ""
        Let FiveByteD.Text = ""
        Let SixByteD.Text = ""
    Else
        ' This routing assembles the byte for speed step mode 14
        
        Let Temporary = Temporary + Val(ConsistControlSpeed.Value) ' add the speed
        Let ThreeByteD.Text = Temporary
        Let FourByteD.Text = ""
        Let FiveByteD.Text = ""
        Let SixByteD.Text = ""
    
    End If
End If



End Sub

Private Sub Command6_Click()

If ShortAdDress.Value = unvbChecked Then
    Let onebyted.Text = Int(Val(consistcontrolmacrolocomotivenumber.Text) / 256)
    Let twoByteD.Text = Val(consistcontrolmacrolocomotivenumber.Text) - (Val(onebyted.Text) * 256)
    Let onebyted.Text = Val(onebyted.Text) + 128 + 64
    Let ConsistControlComment.Text = "Loco " + consistcontrolmacrolocomotivenumber.Text + "; "
End If

If ShortAdDress.Value = vbChecked Then
    Let onebyted.Text = Int(Val(consistcontrolmacrolocomotivenumber.Text))
    Let twoByteD.Text = ""
    Let ConsistControlComment.Text = "Consist " + consistcontrolmacrolocomotivenumber.Text + "; "
End If

End Sub


Private Sub Command7_Click()

Let TemporaryByte = 128 + 32

If ConsistControlFunction5.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 1
    Let ConsistControlComment.Text = ConsistControlComment.Text + "5 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "5 Off;"
End If

If ConsistControlFunction6.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 2
    Let ConsistControlComment.Text = ConsistControlComment.Text + "6 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "6 Off;"
End If

If ConsistControlFunction7.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 4
    Let ConsistControlComment.Text = ConsistControlComment.Text + "7 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "7 Off;"
End If

If ConsistControlFunction8.Value = vbChecked Then
    Let TemporaryByte = TemporaryByte + 8
    Let ConsistControlComment.Text = ConsistControlComment.Text + "8 On; "
Else
    Let ConsistControlComment.Text = ConsistControlComment.Text + "8 Off;"
End If

Let ThreeByteD.Text = TemporaryByte
Let FourByteD.Text = ""
Let FiveByteD.Text = ""
Let SixByteD.Text = ""

End Sub

Private Sub Command8_Click()

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
  
If ConsistControlSendCS.Value = vbChecked Then
    
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.
'
' Before setting the communication port, I used this let statement to set the visual status on the screen. Most of the
' Screen contain this lable to help notify the user what is happening with the program.
'
' Let Statements
'
' Two Visual Basic statements are used in combination with the assignment operator (=).
' The Let statement, although usually implicit, is used for assigning values.
' The Set statement, which must always be explicit, is used for assigning object references.
' If you use Let instead of Set when assigning an object reference, you will generally end up assigning the value of the object's default property.
' Attempting to use the resulting variable as an object reference will usually result in an error, such as  Error 424 Object required.

Let LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

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
    
    
If MainScreen!checkboxdequeuepacket.Value = vbChecked Then
    
' Now that the Locomotive Communication Window has be updated...
    
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

Let LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

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
        'If SevenByteD.Text <> "" Then
        '    If SevenByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl & " " & SevenByteH.Text
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
End If

' -------------------------------------------------------------------------------------------------------------------------
' Automatic Addtion to Comments line
'




End Sub

Private Sub Command9_Click()
    
        Let TemporaryByteOne = 0
        Let TemporaryByteTwo = Val(ConsistControlCV.Text) - 1
        
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
        
       If ConsistControlCVRead = vbChecked Then
              Let TemporaryByteOne = TemporaryByteOne + 4
        Else
            Let TemporaryByteOne = TemporaryByteOne + 8 + 4
        End If
        
    Let ThreeByteD.Text = TemporaryByteOne
    Let FourByteD.Text = TemporaryByteTwo
    Let FiveByteD.Text = Val(ConsistControlCVValue.Text)
    Let SixByteD.Text = ""
    
 Let ConsistControlComment.Text = ConsistControlComment.Text + _
     "Change CV" + ConsistControlCV.Text + " to " + ConsistControlCVValue.Text

    
End Sub

Private Sub ConsistControlCVRead_Click()

If ConsistControlCVRead.Value = vbChecked Then
    Let ConsistControlCVWright.Value = vbUnchecked
Else
    Let ConsistControlCVWright.Value = vbChecked
End If

End Sub

Private Sub ConsistControlCVWright_Click()

If ConsistControlCVWright.Value = vbChecked Then
    Let ConsistControlCVRead.Value = vbUnchecked
Else
    Let ConsistControlCVRead.Value = vbChecked
End If

End Sub



Private Sub ConsistControlDirectionF_Click()

If ConsistControlDirectionF.Value = vbcheck Then
    Let ConsistControlDirectionF.Value = vbUnchecked
    Let ConsistControlDirectionR.Value = vbChecked
Else
    Let ConsistControlDirectionF.Value = vbChecked
    Let ConsistControlDirectionR.Value = vbUnchecked
End If

End Sub

Private Sub ConsistControlDirectionR_Click()

If ConsistControlDirectionR.Value = vbcheck Then
    Let ConsistControlDirectionR.Value = vbUnchecked
    Let ConsistControlDirectionF.Value = vbChecked
Else
    Let ConsistControlDirectionR.Value = vbChecked
    Let ConsistControlDirectionF.Value = vbUnchecked
End If
    

End Sub

Private Sub ConsistControlFilename_Change()


If ConsistControlFilename.Text = "" Then
    Let Command4.Enabled = False
    Let Command5.Enabled = False
    Let Command6.Enabled = False
    Let Command7.Enabled = False
    Let Command8.Enabled = False
    Let Command9.Enabled = False
    Let LocomotiveDatabaseConsistControl.Enabled = False
Else
    Let Command4.Enabled = True
    Let Command5.Enabled = True
    Let Command6.Enabled = True
    Let Command7.Enabled = True
    Let Command8.Enabled = True
    Let Command9.Enabled = True
    Let LocomotiveDatabaseConsistControl.Enabled = True

    
Let LocomotiveDatabaseConsistControl.DatabaseName = ConsistControlFilename.Text

End If

End Sub

Private Sub ConsistControlSpeed_Change()

Let ConsistControlSpeedValue.Caption = ConsistControlSpeed.Value

End Sub

Private Sub ConsistControlSpeed128_Click()

If ConsistControlSpeed128.Value = vbChecked Then
    Let ConsistControlSpeed28.Value = vbUnchecked
    Let ConsistControlSpeed14.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 127
Let ConsistControlSpeed.Value = 0

End Sub

Private Sub ConsistControlSpeed14_Click()

If ConsistControlSpeed14.Value = vbChecked Then
    Let ConsistControlSpeed28.Value = vbUnchecked
    Let ConsistControlSpeed128.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 15
Let ConsistControlSpeed.Value = 0



End Sub

Private Sub ConsistControlSpeed28_Click()

If ConsistControlSpeed28.Value = vbChecked Then
    Let ConsistControlSpeed14.Value = vbUnchecked
    Let ConsistControlSpeed128.Value = vbUnchecked
End If

Let ConsistControlSpeed.Min = 0
Let ConsistControlSpeed.Max = 29
Let ConsistControlSpeed.Value = 0

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
        If TemporaryScreen = "Mainline Macro Maker Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Macro Maker Screen"
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
            Let Ini.Value = "Mainline Macro Maker Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Macro Maker Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Macro Maker Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Macro Maker Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineMacroMaker.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineMacroMaker.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineMacroMaker.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineMacroMaker.Height)

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
                Let Ini.Value = "Mainline Macro Maker Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Macro Maker Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineMacroMaker.Hide
    'unload Mainlinemacromaker

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
    Let Ini.Application = "Mainline Macro Maker Screen"
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
        MainlineMacroMaker.Left = (Screen.Width - Width) / 2
        MainlineMacroMaker.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineMacroMaker.Width > Screen.Width Then
            Let MainlineMacroMaker.Left = Screen.Width - MainlineMacroMaker.Width
        Else
            Let MainlineMacroMaker.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineMacroMaker.Height > Screen.Height Then
            Let MainlineMacroMaker.Top = Screen.Height - MainlineMacroMaker.Height
        Else
            Let MainlineMacroMaker.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button closes the current window and returns" & vbCrLf & "you to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This text box is where all information from your" & vbCrLf & "serial port is displayed. Commands given by the" & vbCrLf & "program are displayed here. You can also type your" & vbCrLf & "own commands, providing the port is not busy."
        Let TemporaryText2 = "Communication Window"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowDCC)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowDCC, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Let LocomotiveDatabaseConsistControl.DatabaseName = App.Path$ & "\Databases\Consist Control Database for Consist 0.mdb"
Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"

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

    If MainlineMacroMaker.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineMacroMaker.WindowState = vbNormal Then
    
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

Private Sub Form_Terminate()
    If CloseMode = vbFormControlMenu Then
        MsgBox "Please use the Close button. Do not close this window buy eXiting."
        Cancel = True
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
    Let Ini.Application = "Mainline Marco Maker Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineMacroMaker.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineMacroMaker.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineMacroMaker.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineMacroMaker.Height)
 
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

    Let ButtonClose.Enabled = False

    Let Timer1.Interval = 0

' Returns or sets a value indicating the type of mouse pointer displayed when the mouse is over a particular part of an object at run time.

    consistcontrolmacrolocomotivenumber.Clear

' Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

    LocomotiveDatabaseSearch.Recordset.MoveFirst

' Use the RecordCount property to find out how many records in a Recordset or TableDef object have been accessed.
' RecordCount doesn't indicate how many records are contained in a dynaset- or snapshot-type Recordset until all records have been accessed.
' Once the last record has been accessed, the RecordCount property indicates the total number of undeleted records in the Recordset or TableDef.
' To force the last record to be accessed, use the MoveLast or FindLast method on the Recordset.
' You can also use an SQL Count function to determine the approximate number of records your query will return.

    Do While Not LocomotiveDatabaseSearch.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    LocomotiveDatabaseSearch.Recordset.MoveNext
    If LocomotiveDatabaseDecoderSearch.Value = vbChecked Then
         consistcontrolmacrolocomotivenumber.AddItem Format(Val(LocomotiveDatabaseNumberSearch.Text), "0000")
     End If

' DoEvents passes control to the operating system.
' Control is not returned until the operating system has finished processing the events in its queue and  all keys in the SendKeys queue have been sent.
' If parts of your code take up a lot of processor time, use DoEvents periodiCally to relinquish control to the operating system so that events, such as keyboard input and mouse clicks, can be processed without significant delay.
  
      Let temp = DoEvents
     
    ' Just in cause the user of the program closes the window while the database is still open searching
    ' This will terminate the search and allow the window to be close withou gerating and error.
         
Loop

Let ButtonClose.Enabled = True

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


