VERSION 4.00
Begin VB.Form MainlineOperationGuiDiesel3Screen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Operation GUI for SD40-2 Diesel"
   ClientHeight    =   8775
   ClientLeft      =   1245
   ClientTop       =   1050
   ClientWidth     =   12060
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Height          =   9180
   Icon            =   "MainlineOperationGuiDiesel3Screen.frx":0000
   Left            =   1185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   914.292
   ScaleMode       =   0  'User
   ScaleWidth      =   1454.176
   Top             =   705
   Width           =   12180
   Begin VB.PictureBox PictureSpeedometer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      ForeColor       =   &H00000000&
      Height          =   1410
      Left            =   8310
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   24
      Top             =   315
      Width           =   1350
      Begin VB.Label LabelScaleSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   450
         TabIndex        =   25
         Top             =   855
         Width           =   465
      End
      Begin VB.Line LineSpeedometer 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         DrawMode        =   1  'Blackness
         X1              =   46
         X2              =   15
         Y1              =   47
         Y2              =   64
      End
   End
   Begin VB.Timer TimerCheckWinsock 
      Left            =   15495
      Top             =   10755
   End
   Begin VB.Timer TimerEngineRunning 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15480
      Top             =   270
   End
   Begin VB.CommandButton ButtonHelp 
      Caption         =   "&Help is Off"
      Height          =   255
      Left            =   14040
      TabIndex        =   7
      Top             =   10860
      Width           =   1215
   End
   Begin VB.CommandButton ButtonCloseGUI 
      Caption         =   "&Close"
      Height          =   255
      Left            =   14040
      TabIndex        =   5
      Top             =   11160
      Width           =   1230
   End
   Begin VB.CommandButton ButtonEngineStart 
      Caption         =   "&Start Engine"
      Height          =   255
      Left            =   14040
      TabIndex        =   4
      Top             =   10560
      Width           =   1215
   End
   Begin FATHMAILOCXLib.SMTP SmtpOcx 
      Left            =   16710
      Top             =   705
      _Version        =   65536
      _ExtentX        =   688
      _ExtentY        =   820
      _StockProps     =   0
   End
   Begin FATHMAILOCXLib.Message MessageOcx 
      Left            =   16695
      Top             =   240
      _Version        =   65536
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   0
   End
   Begin TransPicture.TransPictureCtl TransPictureThrottleRight 
      Height          =   240
      Left            =   7875
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8235
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   423
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0442
   End
   Begin TransPicture.TransPictureCtl TransPictureSandRight 
      Height          =   345
      Left            =   4725
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7425
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   609
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":045E
   End
   Begin TransPicture.TransPictureCtl TransPictureHeadlightRight 
      Height          =   720
      Left            =   9315
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   9330
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":047A
   End
   Begin VB.Label LabelLocomotiveNumberDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14475
      TabIndex        =   22
      Top             =   8539
      Width           =   675
   End
   Begin MCI.MMControl MMControlHeadlight 
      Height          =   435
      Left            =   16020
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureHeadlight 
      Height          =   795
      Left            =   4485
      TabIndex        =   20
      Tag             =   "0"
      Top             =   10035
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1402
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0496
   End
   Begin TransPicture.TransPictureCtl TransPictureThrottle 
      Height          =   705
      Left            =   7230
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1244
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":04B2
   End
   Begin MCI.MMControl MMControlThrottle 
      Height          =   435
      Left            =   16020
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl MMControlBrakeDynamic 
      Height          =   435
      Left            =   16020
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin MCI.MMControl MMControlReverser 
      Height          =   435
      Left            =   16020
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4020
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureBrakeDynamic 
      Height          =   1155
      Left            =   5700
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7530
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   2037
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":04CE
   End
   Begin TransPicture.TransPictureCtl TransPictureReverser 
      Height          =   780
      Left            =   7410
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   9405
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1376
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":04EA
   End
   Begin MCI.MMControl MMControlSand 
      Height          =   435
      Left            =   16020
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureSand 
      Height          =   540
      Left            =   3450
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8310
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0506
   End
   Begin MCI.MMControl MMControlBell 
      Height          =   435
      Left            =   16020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureBell 
      Height          =   345
      Left            =   3570
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   10950
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0522
   End
   Begin MCI.MMControl MMControlBrakeIndependent 
      Height          =   435
      Left            =   16020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureBrakeIndependent 
      Height          =   840
      Left            =   1680
      TabIndex        =   8
      Tag             =   "0"
      Top             =   10665
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1482
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":053E
   End
   Begin MCI.MMControl MMControlBrakeAutomatic 
      Height          =   435
      Left            =   16020
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureBrakeAutomatic 
      Height          =   2550
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4498
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":055A
   End
   Begin MCI.MMControl MMControlHorn 
      Height          =   435
      Left            =   16020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureHorn 
      Height          =   390
      Left            =   3345
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   5430
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0576
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   16080
      Top             =   720
      _ExtentX        =   661
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   16020
      Top             =   240
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   16020
      Top             =   1260
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin TransPicture.TransPictureCtl TransPictureCab 
      Height          =   11520
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   20320
      TransparentColor=   16777215
      Picture         =   "MainlineOperationGuiDiesel3Screen.frx":0592
   End
   Begin VIDEOCAPXLib.VideoCapX VideoCapture 
      Height          =   7343
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15359
      _Version        =   131072
      _ExtentX        =   27093
      _ExtentY        =   12965
      _StockProps     =   1
      CapFilename     =   ""
      ColorFormat     =   0
      AudioInputIndex =   0
      StretchMode     =   1
   End
End
Attribute VB_Name = "MainlineOperationGuiDiesel3Screen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




















Public Sub ButtonEngineStart_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Start Engine, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Start Engine, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateEngineStart
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings!textboxincomingdata.Text = "Stop Engine, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Stop Engine, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateEngineStop
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Start Engine, confirmed." Then
            GoSub UpdateEngineStart
            Let InternetSettings.textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Stop Engine, confirmed." Then
            GoSub UpdateEngineStop
            Let InternetSettings.textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "" Then
            If ButtonEngineStart.Caption = "&Start Engine" Then
                Let InternetSettings!textboxoutboundcommand.Text = "Start Engine, requested."
            Else
                Let InternetSettings!textboxoutboundcommand.Text = "Stop Engine, requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        If ButtonEngineStart.Caption = "&Start Engine" Then
            Let ButtonCloseGUI.Enabled = False
            GoSub UpdateEngineStart
        Else
            GoSub UpdateEngineStop
            Let ButtonCloseGUI.Enabled = True
        End If
    End If
        ' The program must be running with out  any internet connection

Exit Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateEngineStart:
    Let ButtonEngineStart.Caption = "&Stop Engine"
    Let ButtonCloseGUI.Enabled = False
    Let TimerEngineRunning.Enabled = True
Return

UpdateEngineStop:
    Let ButtonEngineStart.Caption = "&Start Engine"
    Let ButtonCloseGUI.Enabled = True
    Let TimerEngineRunning.Enabled = False
Return

End Sub

Private Sub ButtonHelp_Click()

If ButtonHelp.Caption = "&Help is Off" Then
    Let ButtonHelp.Caption = "&Help is On"
Else
    Let ButtonHelp.Caption = "&Help is Off"
End If

End Sub



Public Sub ButtonCloseGUI_Click()

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Button Close, requested." Then
            GoSub UpdateButtonClose
            Let InternetSettings!textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Button Close, requested."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateButtonClose
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateButtonClose
    End If

Exit Sub

UpdateButtonClose:
    
    Let ButtonCloseGUI.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Video Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.optionHost = True Then
        VideoCapture.StopBroadcast
        Let VideoCapture.Overlay = False
        Let VideoCapture.Preview = False
        Let VideoCapture.Connected = False
        Let VideoCapture.VideoRenderer = False
    ElseIf InternetSettings.optionclient = True Then
        If InternetSettings.NetConnect.Connected = True Then
            VideoCapture.DisplayRemote "", False
        End If
    ElseIf InternetSettings.OptionStandAlone = True Then
        If VideoCapture.GetVideoDeviceCount <> 0 Then
            Let VideoCapture.Preview = False
            Let VideoCapture.Overlay = False
            Let VideoCapture.Connected = False
        End If
    End If
    
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
            If TemporaryScreen = "Mainline Operation GUI Diesel3 Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel3 Screen" Then
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
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
Return
   
End Sub


Private Sub Form_Activate()

    DoEvents
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Mainline Operation GUI Diesel3 Screen, Form, Activate" & vbCrLf
    End If ' Debug Tag

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
        If TemporaryScreen = "Mainline Operation GUI Diesel3 Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen"
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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Form Activate, stack is full, overflow."
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
            For OutsideLoop = 0 To Val(255)
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
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Room Lighitng Control - Are the light truly off?
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If RoomLightingControl!CheckBoxUseRoomLightingControl.Value = vbChecked Then
        Call RoomLightingControl.ButtonManualControl_Click
    End If
    If RoomLightingControl!CheckBoxUseRoomLightingControl.Value = vbChecked Then
        If RoomLightingControl!ButtonManualControl.Caption = "&Lights are On" Then
            Call RoomLightingControl.ButtonManualControl_Click
        End If
    End If
' =============================================================================================================================================================================
' Turn Audio/Video On
' =============================================================================================================================================================================
    If InternetSettings.optionHost = True Then
        If InternetSettings.NetConnect.Connected = True Then
            Call VideoModule.VideoOn
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Start Broadcasting
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Video Settings Screen"
            Let Ini.Parameter = "VideoStreamingClientsMaximum"
            VideoCapture.StartBroadcast 20103, Val(Ini.Value)
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Client Mode
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf InternetSettings.optionclient = True Then
        If InternetSettings.NetConnect.Connected = True Then
            If InternetSettings.checkboxclientstreamtypebroadcast.Value = vbChecked Then
                VideoCapture.PlayerOpen "http://atc.server" & CStr(Val(InternetSettings!ComboServerName.ListIndex)) & ".lovethosetrains.com:20103"
                DoEvents
                VideoCapture.PlayerStart
            ElseIf InternetSettings.checkboxclientstreamtypeserver.Value = vbChecked Then
                VideoCapture.DisplayRemote "atc.server" & CStr(Val(InternetSettings!ComboServerName.ListIndex)) & ".lovethosetrains.com", True
            End If
        End If
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Activate TimerCheckWinsock
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let TimerCheckWinsock.Enabled = True
    ElseIf InternetSettings.OptionStandAlone = True Then
        If VideoCapture.GetVideoDeviceCount <> 0 Then
            Call VideoModule.VideoOn
         End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display Locomotive Number in GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelLocomotiveNumberDisplay.Caption = Val(MainlineOperationGUI!ComboBoxLocomotiveNumber.Text)
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialize all function in decoders to off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.optionclient.Value <> "True" Then
       If MainlineOperationGUI.CheckBoxSoundDecoderEquipped = vbChecked Then
           Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
           Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction01234
           Call MainScreen.SendCommandviaTrackQ
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction5678
           Call MainScreen.SendCommandviaTrackQ
           Call MainlineOperationGUI.SetSoundDecoderNumber
           Call MainlineOperationGUI.SetSoundControlFunction9101112
           Call MainScreen.SendCommandviaTrackQ
       End If
       
       Let MainlineOperationGUI.ConsistControlFunction0.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction1.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction2.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction3.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction4.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction5.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction6.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction7.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction8.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction9.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction10.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction11.Value = vbUnchecked
       Let MainlineOperationGUI.ConsistControlFunction12.Value = vbUnchecked
    
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction01234
       Call MainScreen.SendCommandviaTrackQ
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction5678
       Call MainScreen.SendCommandviaTrackQ
       Call MainlineOperationGUI.SetLocomotiveNumber
       Call MainlineOperationGUI.SetConsistFunction9101112
       Call MainScreen.SendCommandviaTrackQ
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Mainline Operation GUI Diesel3 Screen, Form, Activate" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Operation GUI Diesel3 Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationGuiDiesel3Screen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationGuiDiesel3Screen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationGuiDiesel3Screen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationGuiDiesel3Screen.Height)

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
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Form Deactivate, variable error in ATC.INI file for 'Tranparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineOperationGuiDiesel3Screen.Hide
    Unload MainlineOperationGuiDiesel3Screen

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub


Private Sub Form_Load()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Load Pictures
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryLocomotivePath As String
    Let TemporaryLocomotivePath$ = "\Graphics\Locomotive Diesel3\"

    TransPictureCab.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "CabScreen(d3).bmp")
    TransPictureHorn.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Horn0(d3).bmp")
    TransPictureBrakeAutomatic.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "BrakeAutomatic0(d3).bmp")
    TransPictureBrakeIndependent.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "BrakeIndependent0(d3).bmp")
    TransPictureSand.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Sand0(d3).bmp")
    TransPictureSandRight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "SandRight0(d3).bmp")
    TransPictureBell.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Bell0(d3).bmp")
    TransPictureThrottle.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Throttle0(d3).bmp")
    TransPictureThrottleRight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "ThrottleRight0(d3).bmp")
    TransPictureBrakeDynamic.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "BrakeDynamic0(d3).bmp")
    TransPictureReverser.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Reverser1(d3).bmp")
    TransPictureHeadlight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Headlight0(d3).bmp")
    TransPictureHeadlightRight.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "HeadlightRight0(d3).bmp")
    PictureSpeedometer.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ & "Speedometer(d3).bmp")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub

























Private Sub Form_Resize()

    If MainlineOperationGuiDiesel3Screen.WindowState = vbMinimized Then

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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Resize, variable not set correctly for 'BackGround Image' in ATC.INI file."
        End If

    ElseIf MainlineOperationGuiDiesel3Screen.WindowState = vbNormal Then

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
            Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Resize, variable not set correctly for 'BackGround Image' in ATC.INI file."
        End If

    End If

End Sub

Private Sub PictureBoxAutomaticBrake_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyAscii = Asc("A") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) + 1

End Sub

Private Sub PictureBoxAutomaticBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxAutomaticBrake.Tag) > 0 Then
        Let PictureBoxAutomaticBrake.Tag = Trim(Str(Val(PictureBoxAutomaticBrake.Tag) - 1))
        If Val(PictureBoxAutomaticBrake.Tag) = 0 Then
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentClose.wav"
        Else
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxAutomaticBrake.Tag) < 17 Then
        Let PictureBoxAutomaticBrake.Tag = Trim(Str$(Val(PictureBoxAutomaticBrake.Tag) + 1))
        If Val(PictureBoxAutomaticBrake.Tag) = 1 Then
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentOpen.wav"
        Else
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\AutomaticBrake"
Let Temporary$ = Temporary$ + PictureBoxAutomaticBrake.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxAutomaticBrake.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxBell_Click()

Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
Let MainlineOperationGUI!Wave1.Action = wAPlay

If Val(PictureBoxBell.Tag) = 0 Then
    Let PictureBoxBell.Tag = "1"
Else
    Let PictureBoxBell.Tag = "0"
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\Bell"
Let Temporary$ = Temporary$ + PictureBoxBell.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxBell.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxBell_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxBell.Left = Val(PictureBoxBell.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxBell.Left = Val(PictureBoxBell.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxBell.Top = Val(PictureBoxBell.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxBell.Top = Val(PictureBoxBell.Top) + 1

End Sub


Private Sub PictureBoxDynamicBrake_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxDynamicBrake.Left = Val(PictureBoxDynamicBrake.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxDynamicBrake.Left = Val(PictureBoxDynamicBrake.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxDynamicBrake.Top = Val(PictureBoxDynamicBrake.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxDynamicBrake.Top = Val(PictureBoxDynamicBrake.Top) + 1

End Sub

Private Sub PictureBoxDynamicBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxDynamicBrake.Tag) > 0 Then
        Let PictureBoxDynamicBrake.Tag = Trim(Str$(Val(PictureBoxDynamicBrake.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ContrlThrottle.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the dynamic brake."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
    
ElseIf Button = vbRightButton Then
    If Val(PictureBoxDynamicBrake.Tag) < 9 Then
        Let PictureBoxDynamicBrake.Tag = Trim(Str(Val(PictureBoxDynamicBrake.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the dynamic brake"
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\DynamicBrake"
Let Temporary$ = Temporary$ + PictureBoxDynamicBrake.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxDynamicBrake.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxHeadlight_KeyPress(KeyAscii As Integer)

End Sub










Private Sub LabelLocomotiveCurrentSpeed_Click()

End Sub


Private Sub TimerCheckWinsock_Timer()

    Let InternetSettings!textboxoutboundcommand.Text = "Checking Winsock Connection."
    Call InternetSettings.ButtonSend_Click
    Let InternetSettings!textboxincomingdata.Text = ""

End Sub


Private Sub TimerEngineRunnig_Timer()

End Sub

Private Sub TimerEngineRunning_Timer()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Spped Step
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let TempThrottle = 15.875 * Val(TransPictureThrottle.Tag)
    Let TempBrakeAutomatic = 4 * Val(TransPictureBrakeAutomatic.Tag)
    Let TempBrakeIndependent = 4 * Val(TransPictureBrakeIndependent.Tag)
    Let TempBrakeDynamic = 8 * Val(TransPictureBrakeDynamic.Tag)

    Let TempSpeed = Int(TempThrottle - TempBrakeAutomatic - TempBrakeIndependent - TempBrakeDynamic)
    If TempSpeed < 0 Then Let TempSpeed = 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Apply Speed
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TempSpeed > MainlineOperationGUI.ConsistControlSpeed.Value Then
        Let MainlineOperationGUI.ConsistControlSpeed.Value = MainlineOperationGUI.ConsistControlSpeed.Value + 1
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                If MainlineOperationGUI.ConsistControlSpeed.Value <> 1 Then
                    Call MainlineOperationGUI.SetLocomotiveNumber
                    Call MainlineOperationGUI.SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                End If
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings!OptionStandAlone.Value = True Then
                If MainlineOperationGUI.ConsistControlSpeed.Value <> 1 Then
                    Call MainlineOperationGUI.SetLocomotiveNumber
                    Call MainlineOperationGUI.SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                End If
            End If
        End If
    ElseIf TempSpeed < MainlineOperationGUI.ConsistControlSpeed.Value Then
        Let MainlineOperationGUI.ConsistControlSpeed.Value = MainlineOperationGUI.ConsistControlSpeed.Value - 1
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                If MainlineOperationGUI.ConsistControlSpeed.Value <> 1 Then
                    Call MainlineOperationGUI.SetLocomotiveNumber
                    Call MainlineOperationGUI.SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                End If
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings!OptionStandAlone.Value = True Then
                If MainlineOperationGUI.ConsistControlSpeed.Value <> 1 Then
                    Call MainlineOperationGUI.SetLocomotiveNumber
                    Call MainlineOperationGUI.SetSpeed
                    Call MainScreen.SendCommandviaTrackQ
                End If
            End If
        End If
    End If
    
    Let LineSpeedometer.X2 = LineSpeedometer.X1 - Cos((MainlineOperationGUI.ConsistControlSpeed.Value * 2 - 30) / 180 * 3.14159265) * 40
    Let LineSpeedometer.Y2 = LineSpeedometer.Y1 - Sin((MainlineOperationGUI.ConsistControlSpeed.Value * 2 - 30) / 180 * 3.14159265) * 40
    Let LabelScaleSpeed.Caption = CStr(MainlineOperationGUI.ConsistControlSpeed.Value)
    
End Sub
Public Sub TransPictureBell_Click()

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Bell, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Bell, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateBell
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Bell, confirmed." Then
            GoSub UpdateBell
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Bell, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateBell
    End If

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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If TransPictureBell.Tag = "0" Then
            Let BalloonHelpText1 = "The bell is not active."
        Else 'If TransPictureBell.Tag = "1" Then
            Let BalloonHelpText1 = "The bell is active."
        End If
        
        Let BalloonHelpText2 = "Bell"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureBell)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBell, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Bell, unable to setup balloon help for 'TransPictureBell' control."
            End If
        End If

    End If
    
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateBell:

    Let MMControlBell.Command = "Close"
    Let MMControlBell.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let MMControlBell.Command = "Open"
    Let MMControlBell.Command = "Play"

    If TransPictureBell.Tag = "0" Then
        Let TransPictureBell.Tag = "1"

        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
        End If
        
    ElseIf TransPictureBell.Tag = "1" Then
        Let TransPictureBell.Tag = "0"
    
        If MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionBell.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbUnchecked
        End If

    End If

    Let TransPictureBell.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\Bell" & TransPictureBell.Tag & "(d3).bmp")
    
    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
            End If
        End If
    End If
Return

End Sub

Public Sub TransPictureBrakeAutomatic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
    
    If ButtonEngineStart.Caption = "&Stop Engine" Then
    
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Automatic Brake, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateBrakeAutomatic
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Automatic Brake, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateBrakeAutomatic
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Automatic Brake, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateBrakeAutomatic
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Automatic Brake, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateBrakeAutomatic
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Automatic Brake, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateBrakeAutomatic
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Automatic Brake handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the trainline brake is at " & Str$(Int(Val(100 - (TextBoxStatusAutomaticBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Automatic Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(transpicturebrakeautomatic)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBrakeAutomatic, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Automatic Brake, unable to setup balloon help for 'transpicturebrakeautomatic' control."
            End If
        End If

    End If

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateBrakeAutomatic:

    If Button = vbRightButton Then
        If Val(TransPictureBrakeAutomatic.Tag) < 15 Then
            Let TransPictureBrakeAutomatic.Tag = Val(TransPictureBrakeAutomatic.Tag) + 1
            Let MMControlBrakeAutomatic.Command = "Close"
            If Val(TransPictureBrakeAutomatic.Tag) = 1 Then
                Let MMControlBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticOpen.wav"
            Else
                Let MMControlBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let MMControlBrakeAutomatic.Command = "Open"
            Let MMControlBrakeAutomatic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum braking, one hundred (100) percent for the automatic brakes (train line brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbLeftButton Then
        If Val(TransPictureBrakeAutomatic.Tag) > 0 Then
            Let TransPictureBrakeAutomatic.Tag = Val(TransPictureBrakeAutomatic.Tag) - 1
            Let MMControlBrakeAutomatic.Command = "Close"
            If TransPictureBrakeAutomatic.Tag = 0 Then
                Let MMControlBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticClose.wav"
            Else
                Let MMControlBrakeAutomatic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let MMControlBrakeAutomatic.Command = "Open"
            Let MMControlBrakeAutomatic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum braking, zero (0) precent for the automatic brakes (train line brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    Let temp = App.Path$ & "\Graphics\Locomotive Diesel3\BrakeAutomatic" & Trim$(TransPictureBrakeAutomatic.Tag) & "(d3).bmp"
    Let TransPictureBrakeAutomatic.Picture = LoadPicture(temp)
    
'    If Val(TransPictureThrottle.Tag) <> 0 And Val(TransPictureReverser.Tag) <> 1 Then
'        Let MainlineOperationGUI!timerdeadmann.Enabled = False
'        Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
'        Let MainlineOperationGUI!timerdeadmann.Enabled = True
'    Else
'        Let MainlineOperationGUI!timerdeadmann.Enabled = False
'    End If

Return
End Sub


Public Sub TransPictureBrakeDynamic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
    
    If ButtonEngineStart.Caption = "&Stop Engine" Then
    
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Dynamic Brake, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Dynamic Brake, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateBrakeDynamic
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Dynamic Brake, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateBrakeDynamic
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Dynamic Brake, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Dynamic Brake, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateBrakeDynamic
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Dynamic Brake, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateBrakeDynamic
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Dynamic Brake, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Dynamic Brake, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateBrakeDynamic
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Automatic Brake handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the trainline brake is at " & Str$(Int(Val(100 - (TextBoxStatusAutomaticBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Automatic Brakes (Trainline)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureBrakeDynamic)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBrakeDynamic, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Dynamic Brake, unable to setup balloon help for 'TransPictureBrakeDynamic' control."
            End If
        End If

    End If

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateBrakeDynamic:

    If Button = vbRightButton Then
        If Val(TransPictureBrakeDynamic.Tag) < 9 Then
            Let TransPictureBrakeDynamic.Tag = Val(TransPictureBrakeDynamic.Tag) + 1
            Let MMControlBrakeDynamic.Command = "Close"
            Let MMControlBrakeDynamic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            Let MMControlBrakeDynamic.Command = "Open"
            Let MMControlBrakeDynamic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum braking, one hundred (100) percent for the dynamic brakes."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbLeftButton Then
        If Val(TransPictureBrakeDynamic.Tag) > 0 Then
            Let TransPictureBrakeDynamic.Tag = Val(TransPictureBrakeDynamic.Tag) - 1
            Let MMControlBrakeDynamic.Command = "Close"
            Let MMControlBrakeDynamic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            Let MMControlBrakeDynamic.Command = "Open"
            Let MMControlBrakeDynamic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum braking, zero (0) precent for the dynamic brakes."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    Let TransPictureBrakeDynamic.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\BrakeDynamic" & Trim$(TransPictureBrakeDynamic.Tag) & "(d3).bmp")
    

Return

End Sub


Public Sub TransPictureBrakeIndependent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
    
    If ButtonEngineStart.Caption = "&Stop Engine" Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Independent Brake, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateBrakeIndependent
                Let InternetSettings!textboxincomingdata.Text = ""
            
            ElseIf InternetSettings.textboxincomingdata.Text = "Independent Brake, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateBrakeIndependent
                Let InternetSettings.textboxincomingdata.Text = ""
            
            ElseIf InternetSettings!textboxincomingdata.Text = "Independent Brake, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateBrakeIndependent
                Let InternetSettings!textboxincomingdata.Text = ""
            
            ElseIf InternetSettings.textboxincomingdata.Text = "Independent Brake, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateBrakeIndependent
                Let InternetSettings.textboxincomingdata.Text = ""
            
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Independent Brake, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateBrakeIndependent
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the Independent Brake handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
    
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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the locomotive brake is at " & Str$(Int(Val(100 - (TextBoxStatusIndependentBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Independent Brakes (Locomotive)"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureBrakeIndependent)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureBrakeIndependent, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Independent Brake, unable to setup balloon help for 'TransPictureBrakeIndependent' control."
            End If
        End If

    End If

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateBrakeIndependent:

    If Button = vbRightButton Then
        If Val(TransPictureBrakeIndependent.Tag) < 15 Then
            Let TransPictureBrakeIndependent.Tag = Val(TransPictureBrakeIndependent.Tag) + 1
            Let MMControlBrakeIndependent.Command = "Close"
            If Val(TransPictureBrakeIndependent.Tag) = 1 Then
                Let MMControlBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentOpen.wav"
            Else
                Let MMControlBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let MMControlBrakeIndependent.Command = "Open"
            Let MMControlBrakeIndependent.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum braking, one hundred (100) percent for the independent brakes (locomotive brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbLeftButton Then
        If Val(TransPictureBrakeIndependent.Tag) > 0 Then
            Let TransPictureBrakeIndependent.Tag = Val(TransPictureBrakeIndependent.Tag) - 1
            Let MMControlBrakeIndependent.Command = "Close"
            If TransPictureBrakeIndependent.Tag = 0 Then
                Let MMControlBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentClose.wav"
            Else
                Let MMControlBrakeIndependent.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            End If
            Let MMControlBrakeIndependent.Command = "Open"
            Let MMControlBrakeIndependent.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum braking, zero (0) precent for the independent brakes (locomotive brakes)."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    Let temp = App.Path$ & "\Graphics\Locomotive Diesel3\BrakeIndependent" & Trim$(TransPictureBrakeIndependent.Tag) & "(d3).bmp"
    Let TransPictureBrakeIndependent.Picture = LoadPicture(temp)
    
'    If Val(TransPictureThrottle.Tag) <> 0 And Val(TransPictureReverser.Tag) <> 1 Then
'        Let MainlineOperationGUI!timerdeadmann.Enabled = False
'        Let MainlineOperationGUI!timerdeadmann.Interval = MainlineOperationGuiDiesel1Screen!TextBoxSettingDeadmannInterval.Text
'        Let MainlineOperationGUI!timerdeadmann.Enabled = True
'    Else
'        Let MainlineOperationGUI!timerdeadmann.Enabled = False
'    End If

Return
End Sub



Private Sub TransPictureHeadlightLeft_Click()

End Sub

Public Sub TransPictureHeadlight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
        
    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Headlight, left button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Headlight, left button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbLeftButton
            GoSub UpdateHeadlight
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Headlight, left button confirmed." Then
            Let Button = vbLeftButton
            GoSub UpdateHeadlight
            Let InternetSettings.textboxincomingdata.Text = ""
            
        ElseIf InternetSettings!textboxincomingdata.Text = "Headlight, right button requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Headlight, right button confirmed."
            Call InternetSettings.ButtonSend_Click
            Let Button = vbRightButton
            GoSub UpdateHeadlight
            Let InternetSettings!textboxincomingdata.Text = ""
            
        ElseIf InternetSettings.textboxincomingdata.Text = "Headlight, right button confirmed." Then
            Let Button = vbRightButton
            GoSub UpdateHeadlight
            Let InternetSettings.textboxincomingdata.Text = ""
            
        Else
            If Button = vbLeftButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Headlight, left button requested."
            ElseIf Button = vbRightButton Then
                Let InternetSettings!textboxoutboundcommand.Text = "Headlight, right button requested."
            End If
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
         GoSub UpdateHeadlight
    End If

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If TransPictureHeadlight.Tag = "0" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently off."
        ElseIf TransPictureHeadlight.Tag = "1" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on and dimmed."
        ElseIf TransPictureHeadlight.Tag = "2" Then
            Let BalloonHelpText1 = "The headlight of the lead locomotive is currently on."
        End If
        
        Let BalloonHelpText2 = "Headlight"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureHeadlight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureHeadlight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Headlight, unable to setup balloon help for 'PictureBoxHeadlight' control."
            End If
        End If
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureHeadlight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureHeadlightRight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Picture Box Headlight Right, unable to setup balloon help for 'PictureBoxHeadlightRight' control."
            End If
        End If
    End If
    
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateHeadlight:

    If Button = vbRightButton Then
        If Val(TransPictureHeadlight.Tag) < 2 Then
            Let TransPictureHeadlight.Tag = Val(TransPictureHeadlight.Tag) + 1
        End If
    ElseIf Button = vbLeftButton Then
        If Val(TransPictureHeadlight.Tag) > 0 Then
            Let TransPictureHeadlight.Tag = Val(TransPictureHeadlight.Tag) - 1
        End If
    End If
    
    Let MMControlHeadlight.Command = "Close"
    Let MMControlHeadlight.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
    Let MMControlHeadlight.Command = "Open"
    Let MMControlHeadlight.Command = "Play"
    
    Let TransPictureHeadlight.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\HeadLight" & Trim$(TransPictureHeadlight.Tag) & "(d3).bmp")
    Let TransPictureHeadlightRight.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\HeadlightRight" & Trim$(TransPictureHeadlight.Tag) & "(d3).bmp")
    
    If TransPictureHeadlight.Tag = "0" Then
    
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
        
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
            
    ElseIf TransPictureHeadlight.Tag = "1" Then
                    
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
        
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
            
    ElseIf TransPictureHeadlight.Tag = "2" Then
            
        If MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderLight.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
        End If
            
        If MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "0" Then
            Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "1" Then
            Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "2" Then
            Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "3" Then
            Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "4" Then
            Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "5" Then
            Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "6" Then
            Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "7" Then
            Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "8" Then
            Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "9" Then
            Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "10" Then
            Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "11" Then
            Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI!TextBoxLocomotiveDecoderDim.Text = "12" Then
            Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
        End If
    End If

    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction01234
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction5678
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction9101112
                Call MainScreen.SendCommandviaTrackQ
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction01234
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction5678
                Call MainScreen.SendCommandviaTrackQ
                Call MainlineOperationGUI.SetLocomotiveNumber
                Call MainlineOperationGUI.SetConsistFunction9101112
                Call MainScreen.SendCommandviaTrackQ
            End If
        End If
    End If

Return

End Sub

Private Sub TransPictureHeadlightRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call TransPictureHeadlight_MouseDown(Button, Shift, X, Y)
    
End Sub


Public Sub TransPictureHorn_Click()
        
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Horn, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Horn, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateHorn
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Horn, confirmed." Then
            GoSub UpdateHorn
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Horn, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        ' The program must be running with out any internet connection
        GoSub UpdateHorn
    End If

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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        If TransPictureHorn.Tag = "0" Then
            Let BalloonHelpText1 = "The horn is not active."
        ElseIf TransPictureHorn.Tag = "1" Then
            Let BalloonHelpText1 = "The horn is active."
        End If
        
        Let BalloonHelpText2 = "Horn"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureHorn)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureHorn, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Horn, unable to setup balloon help for 'TransPictureHorn' control."
            End If
        End If

        
    End If
    
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub
        
' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateHorn:

    If TransPictureHorn.Tag = "0" Then
        Let TransPictureHorn.Tag = "1"
        Let temp = App.Path
        Let temp = temp + "\Graphics\Locomotive Diesel3\Horn1(d3).bmp"
        Let TransPictureHorn.Picture = LoadPicture(temp)
        
        If MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "0" Then
            Let MainlineOperationGUI!SoundControlFunction0.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "1" Then
            Let MainlineOperationGUI!SoundControlFunction1.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI!SoundControlFunction2.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "3" Then
            Let MainlineOperationGUI!SoundControlFunction3.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "4" Then
            Let MainlineOperationGUI!SoundControlFunction4.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "5" Then
            Let MainlineOperationGUI!SoundControlFunction5.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "6" Then
            Let MainlineOperationGUI!SoundControlFunction6.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "7" Then
            Let MainlineOperationGUI!SoundControlFunction7.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "8" Then
            Let MainlineOperationGUI!SoundControlFunction8.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "9" Then
            Let MainlineOperationGUI!SoundControlFunction9.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "10" Then
            Let MainlineOperationGUI!SoundControlFunction10.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "11" Then
            Let MainlineOperationGUI!SoundControlFunction11.Value = vbChecked
        ElseIf MainlineOperationGUI!TextBoxMappedFunctionHorn.Text = "12" Then
            Let MainlineOperationGUI!SoundControlFunction12.Value = vbChecked
        End If
   
    Else
        Let TransPictureHorn.Tag = "0"
        Let temp = App.Path
        Let temp = temp + "\Graphics\Locomotive Diesel3\Horn0(d3).bmp"
        Let TransPictureHorn.Picture = LoadPicture(temp)
        
        
        If MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "0" Then
            Let MainlineOperationGUI.SoundControlFunction0.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "1" Then
            Let MainlineOperationGUI.SoundControlFunction1.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "2" Then
            Let MainlineOperationGUI.SoundControlFunction2.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "3" Then
            Let MainlineOperationGUI.SoundControlFunction3.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "4" Then
            Let MainlineOperationGUI.SoundControlFunction4.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "5" Then
            Let MainlineOperationGUI.SoundControlFunction5.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "6" Then
            Let MainlineOperationGUI.SoundControlFunction6.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "7" Then
            Let MainlineOperationGUI.SoundControlFunction7.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "8" Then
            Let MainlineOperationGUI.SoundControlFunction8.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "9" Then
            Let MainlineOperationGUI.SoundControlFunction9.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "10" Then
            Let MainlineOperationGUI.SoundControlFunction10.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "11" Then
            Let MainlineOperationGUI.SoundControlFunction11.Value = vbUnchecked
        ElseIf MainlineOperationGUI.TextBoxMappedFunctionHorn.Text = "12" Then
            Let MainlineOperationGUI.SoundControlFunction12.Value = vbUnchecked
        End If

    End If
    
    Let MMControlHorn.Command = "Close"
    Let MMControlHorn.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    Let MMControlHorn.Command = "Open"
    Let MMControlHorn.Command = "Play"
       
    If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings.optionHost.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        ElseIf InternetSettings!Winsock.State = sckClosed Then
            If InternetSettings.OptionStandAlone.Value = True Then
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction01234
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction5678
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
                Call MainlineOperationGUI.SetSoundDecoderNumber
                Call MainlineOperationGUI.SetSoundControlFunction9101112
                Call MainScreen.SendCommandviaTrackQ
                DoEvents
            End If
        End If
    End If

Return

End Sub

Private Sub PictureBoxLocomotiveBrake_KeyPress(KeyAscii As Integer)

End Sub

Private Sub PictureBoxLocomotiveBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxLocomotiveBrake.Tag) > 0 Then
        Let PictureBoxLocomotiveBrake.Tag = Trim(Str$(Val(PictureBoxLocomotiveBrake.Tag) - 1))
        If Val(PictureBoxLocomotiveBrake.Tag) = 0 Then
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentClose.wav"
        Else
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxLocomotiveBrake.Tag) < 16 Then
        Let PictureBoxLocomotiveBrake.Tag = Trim(Str(Val(PictureBoxLocomotiveBrake.Tag) + 1))
        If Val(PictureBoxLocomotiveBrake.Tag) = 1 Then
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentOpen.wav"
        Else
            Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        End If
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\LocomotiveBrake"
Let Temporary$ = Temporary$ + PictureBoxLocomotiveBrake.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxLocomotiveBrake.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxReverser_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) + 1

End Sub

Private Sub PictureBoxReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxReverser.Tag) > -1 Then
        Let PictureBoxReverser.Tag = Trim(Str$(Val(PictureBoxReverser.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the reverser."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxReverser.Tag) < 1 Then
        Let PictureBoxReverser.Tag = Trim(Str$(Val(PictureBoxReverser.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the reverser."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\Reverser"
Let Temporary$ = Temporary$ + PictureBoxReverser.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxReverser.Picture = LoadPicture(Temporary$)

End Sub




Private Sub PictureBoxSand_Click()

Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
Let MainlineOperationGUI!Wave1.Action = wAPlay

If Val(PictureBoxSand.Tag) = 0 Then
    Let PictureBoxSand.Tag = "1"
Else
    Let PictureBoxSand.Tag = "0"
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Diesel3\Sand"
Let Temporary$ = Temporary$ + PictureBoxSand.Tag
Let Temporary$ = Temporary$ + "(d3).bmp"

Let PictureBoxSand.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxSand_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) + 1

End Sub

Private Sub PictureBoxThrottle_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxThrottle.Left = Val(PictureBoxThrottle.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxThrottle.Left = Val(PictureBoxThrottle.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxThrottle.Top = Val(PictureBoxThrottle.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxThrottle.Top = Val(PictureBoxThrottle.Top) + 1

End Sub






Private Sub TimerStartEngine_Timer()

Let TimerStartEngine.Enabled = False

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Stop the Locomotive
'
' If the user selects a locomotive that is currently in use on the layout, and activly moving, then this program will stop the consist.
' This is done by sending a speed command, one, or emergency stop, to the consist.

Let MainlineOperationGUI!ConsistControlSpeed.Value = "1"
Call MainlineOperationGUI.SetLocomotiveNumber
Call MainlineOperationGUI.SetSpeed
Call MainScreen.SendCommandviaTrackQ

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Note The Consist Number
'
' Just move the locomotive or consist number to the graphics user interface display


'Let LabelLocomotiveNumberDisplay.Caption = Val(MainlineOperationGUI!ComboBoxLocomotive Number.Text)


' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Give Resources
'
' Set the resources of the locomotive you are using. Sometime in the future this will be in a database, so that it is remembered.

'Randomize
'Let TextBoxFuelTank.Text = Int((Val(BarFuelTank.Max) - Val(BarFuelTank.Min) + 1) * Rnd + Val(BarFuelTank.Min))
'Let TextBoxOilTank.Text = Int((Val(BarOilTank.Max) - Val(BarOilTank.Min) + 1) * Rnd + Val(BarOilTank.Min))
'Let TextBoxWaterTank.Text = Int((Val(BarWaterTank.Max) - Val(BarWaterTank.Min) + 1) * Rnd + Val(BarWaterTank.Min))
'Let TextBoxSandTank.Text = Int((Val(BarSandTank.Max) - Val(BarSandTank.Min) + 1) * Rnd + Val(BarSandTank.Min))


' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Cab Signals
'
' If Cab Signals are used, update then with pictures of the signals.
'
'If MainlineOperationGUI!OptionVirtualcabSignalYes.Value = True Then
'    Let MainlineOperationGUI!TimerVirtualCabSignal.Interval = 1000
'    Let MainlineOperationGUI!TimerVirtualCabSignal.Enabled = True
'
'    Let PictureBoxVirtualCabSignal.Visible = True
'
'    Let textboxvirtualCabsignalstatuscurrent = "Clear"
'    Let PictureBoxVirtualCabSignalCurrent(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightGreen(d3).bmp")
'    Let PictureBoxVirtualCabSignalCurrent(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightRed(d3).bmp")
'    Let PictureBoxVirtualCabSignalCurrent(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightRed(d3).bmp")'
'
'    Let TextBoxVirtualCabSignalStatusApproaching = "Clear"
'    Let PictureBoxVirtualCabSignalApproaching(0).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightGreen(d3).bmp")
'    Let PictureBoxVirtualCabSignalApproaching(1).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightRed(d3).bmp")
'    Let PictureBoxVirtualCabSignalApproaching(2).Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SignalLightRed(d3).bmp")
'    Stop
'Else
'    Let PictureBoxVirtualCabSignal.Visible = False
'End If
'
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Command Button for Streming Video
'
' If there is a video device, then allow the command buttons to be visible.
    
    If VideoCapture.NumCapDevs > 0 Then
        Let ButtonVideoSettings.Enabled = True
        Let ButtonVideo.Enabled = True
    Else
        Let ButtonVideoSettings.Enabled = False
        Let ButtonVideo.Enabled = False
    End If

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send commands to Sound Decoder
'
' If there user has opted for the use of the sound decoder, send commands to the sound decoder to start the engine sounds.
' Using a Soundtraxx decoder this is done by marking the function RPMPlus with a check and then unchecked.

If MainlineOperationGUI!CheckBoxSoundDecoderEquipped.Value = vbChecked Then

    If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
        Let MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
        Let MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
        Let MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
        Let MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
        Let MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
        Let MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
        Let MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
        Let MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
        Let MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
        Let MainlineOperationGUI!ConsistControlFunction9.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
        Let MainlineOperationGUI!ConsistControlFunction10.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
        Let MainlineOperationGUI!ConsistControlFunction11.Value = vbChecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
        Let MainlineOperationGUI!ConsistControlFunction12.Value = vbChecked
    End If
                                                
    Call MainlineOperationGUI.SetSoundDecoderNumber
    Call MainlineOperationGUI.SetSoundControlFunction01234
    Call MainScreen.SendCommandviaTrackQ
    DoEvents
    Call MainlineOperationGUI.SetSoundControlFunction5678
    Call MainScreen.SendCommandviaTrackQ
    DoEvents
    Call MainlineOperationGUI.SetSoundControlFunction9101112
    Call MainScreen.SendCommandviaTrackQ
    DoEvents
            
    If MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "0" Then
        Let MainlineOperationGUI!ConsistControlFunction0.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "1" Then
        Let MainlineOperationGUI!ConsistControlFunction1.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "2" Then
        Let MainlineOperationGUI!ConsistControlFunction2.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "3" Then
        Let MainlineOperationGUI!ConsistControlFunction3.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "4" Then
        Let MainlineOperationGUI!ConsistControlFunction4.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "5" Then
        Let MainlineOperationGUI!ConsistControlFunction5.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "6" Then
        Let MainlineOperationGUI!ConsistControlFunction6.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "7" Then
        Let MainlineOperationGUI!ConsistControlFunction7.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "8" Then
        Let MainlineOperationGUI!ConsistControlFunction8.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "9" Then
        Let MainlineOperationGUI!ConsistControlFunction9.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "10" Then
        Let MainlineOperationGUI!ConsistControlFunction10.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "11" Then
        Let MainlineOperationGUI!ConsistControlFunction11.Value = vbUnchecked
    ElseIf MainlineOperationGUI!TextBoxMappedFunctionRPMPlus.Text = "12" Then
        Let MainlineOperationGUI!ConsistControlFunction12.Value = vbUnchecked
    End If
                                                
    Call MainlineOperationGUI.SetSoundDecoderNumber
    Call MainlineOperationGUI.SetSoundControlFunction01234
    Call MainScreen.SendCommandviaTrackQ
    DoEvents
    Call MainlineOperationGUI.SetSoundControlFunction5678
    Call MainScreen.SendCommandviaTrackQ
    DoEvents
    Call MainlineOperationGUI.SetSoundControlFunction9101112
    Call MainScreen.SendCommandviaTrackQ
    DoEvents

End If

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start the timers
'
' Timers are the core of the program. They control many function on a continuos basis.

End Sub


Private Sub TimerStartEngine2_Timer()
Let TempTonnage = Val(ScrollTimeAdjustment) / 100 + 1

'Let MainlineOperationGUI!TimerSendCommand.Interval = Val(TextBoxSettingsendcommandInterval.Text) * TempTonnage / ((Val(TextBoxStatusDynamicBrakeApplication.Text) / 3) + 1)
'Let MainlineOperationGUI!TimerSendCommand.Enabled = True
'Let MainlineOperationGUI!TimerSpeedChange.Interval = Val(TextBoxSettingSpeedChangeInterval.Text) * TempTonnage / ((Val(TextBoxStatusDynamicBrakeApplication.Text) / 3) + 1)
'Let MainlineOperationGUI!TimerSpeedChange.Enabled = True
'Let MainlineOperationGUI!timerairmainReservoir.Interval = 2000 * TempTonnage
'Let MainlineOperationGUI!timerairmainReservoir.Enabled = True
'Let MainlineOperationGUI!TimerAirAutomaticBrake.Interval = Val(TextboxSettingAutomaticAirBrake.Text) * TempTonnage
'Let MainlineOperationGUI!TimerAirAutomaticBrake.Enabled = True
'Let MainlineOperationGUI!TimerAirIndependentBrake.Interval = Val(TextboxSettingIndependentAirBrake.Text) * TempTonnage
'Let MainlineOperationGUI!TimerAirIndependentBrake.Enabled = True
'Let MainlineOperationGUI!timerthrottledelay.Interval = 1000
'Let MainlineOperationGUI!timerthrottledelay.Enabled = True
'Let MainlineOperationGUI!TimerOilPressure.Interval = 350 * TempTonnage
'Let MainlineOperationGUI!TimerOilPressure.Enabled = True
'Let MainlineOperationGUI!TimerOilTemperature.Interval = 550 * TempTonnage
'Let MainlineOperationGUI!TimerOilTemperature.Enabled = True
'Let MainlineOperationGUI!TimerWaterPressure.Interval = 350 * TempTonnage
'Let MainlineOperationGUI!TimerWaterPressure.Enabled = True
'Let MainlineOperationGUI!TimerWaterTemperature.Interval = 550 * TempTonnage
'Let MainlineOperationGUI!TimerWaterTemperature.Enabled = True
'Let MainlineOperationGUI!timerfuelconsumption.Interval = 1000 * TempTonnage
'Let MainlineOperationGUI!timerfuelconsumption.Enabled = True
'Let MainlineOperationGUI!TimerOilConsumption.Interval = 3000 * TempTonnage
'Let MainlineOperationGUI!TimerOilConsumption.Enabled = True
'Let MainlineOperationGUI!TimerWaterConsumption.Interval = 3000 * TempTonnage
'Let MainlineOperationGUI!TimerWaterConsumption.Enabled = True
'Let MainlineOperationGUI!TimerRadiatorFans.Interval = 2000 * TempTonnage
'Let MainlineOperationGUI!TimerRadiatorFans.Enabled = True
'Let MainlineOperationGUI!TimerRPM.Interval = 125 * TempTonnage
'Let MainlineOperationGUI!TimerRPM.Enabled = True

End Sub








Private Sub TransPictureHorn_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 40 Then
        Let TransPictureHorn.Top = TransPictureHorn.Top + 1
    ElseIf KeyCode = 38 Then
        Let TransPictureHorn.Top = TransPictureHorn.Top - 1
    ElseIf KeyCode = 39 Then
        Let TransPictureHorn.Left = TransPictureHorn.Left + 1
    ElseIf KeyCode = 37 Then
        Let TransPictureHorn.Left = TransPictureHorn.Left - 1
    End If
    
    Debug.Print TransPictureHorn.Top
    Debug.Print TransPictureHorn.Left

End Sub

Public Sub TransPictureReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Val(LabelScaleSpeed.Caption) = 0 Then

        Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
        
        If ButtonEngineStart.Caption <> "&Start Engine" Then
            If InternetSettings!Winsock.State = sckConnected Then
                If InternetSettings!textboxincomingdata.Text = "Reverser, left button requested." Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Reverser, left button confirmed."
                    Call InternetSettings.ButtonSend_Click
                    Let Button = vbLeftButton
                    GoSub UpdateReverser
                    Let InternetSettings!textboxincomingdata.Text = ""
                    
                ElseIf InternetSettings.textboxincomingdata.Text = "Reverser, left button confirmed." Then
                    Let Button = vbLeftButton
                    GoSub UpdateReverser
                    Let InternetSettings.textboxincomingdata.Text = ""
                    
                ElseIf InternetSettings!textboxincomingdata.Text = "Reverser, right button requested." Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Reverser, right button confirmed."
                    Call InternetSettings.ButtonSend_Click
                    Let Button = vbRightButton
                    GoSub UpdateReverser
                    Let InternetSettings!textboxincomingdata.Text = ""
                    
                ElseIf InternetSettings.textboxincomingdata.Text = "Reverser, right button confirmed." Then
                    Let Button = vbRightButton
                    GoSub UpdateReverser
                    Let InternetSettings.textboxincomingdata.Text = ""
                    
                Else
                    If Button = vbLeftButton Then
                        Let InternetSettings!textboxoutboundcommand.Text = "Reverser, left button requested."
                    ElseIf Button = vbRightButton Then
                        Let InternetSettings!textboxoutboundcommand.Text = "Reverser, right button requested."
                    End If
                    Call InternetSettings.ButtonSend_Click
                    Let InternetSettings!textboxincomingdata.Text = ""
                End If
            Else
                GoSub UpdateReverser
            End If
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                MsgBox "Before making any movements on the Reverser handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
            End If
        End If

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
            Dim BalloonHelpWaveFile As String '
    
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
            Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
                Let balloonhelp.Speech = True
                Let balloonhelp.Voice = 0
                Let BalloonHelpWaveFile = ""
            End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            If TransPictureReverser.Tag = 0 Then
                Let BalloonHelpText1 = "This reverser handle is in the forward position."
            ElseIf TransPictureReverser.Tag = 1 Then
                Let BalloonHelpText1 = "This reverser handle is in the neutral position."
            ElseIf TransPictureReverser.Tag = 2 Then
                Let BalloonHelpText1 = "This reverser handle is in the reverse position."
            End If
            Let BalloonHelpText2 = "Reverser Handle"
            
            'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureReverser)
            Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureReverser, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
            If BalloonHelpSetup = 0 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Reverser, unable to setup balloon help for 'PictureReverser' control."
                End If
            End If
        End If
        
        Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault
        
    Else
        If MainlineOperationGuiDiesel1Screen.ButtonHelp.Caption = "&Help is On" Then
            Let Message = "You cannot move the reverser handle unless the speed of the locomotive is zero."
            MsgBox Message, vbExclamation, "Engineer Error - Control Interlock"
        End If
    End If

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateReverser:


    If Button = vbRightButton Then
        If Val(TransPictureReverser.Tag) < 2 Then
            Let MMControlReverser.Command = "Close"
            Let MMControlReverser.Filename = App.Path$ & "\Sounds\Graphics\ControlReverser.wav"
            Let MMControlReverser.Command = "Open"
            Let MMControlReverser.Command = "Play"
            Let TransPictureReverser.Tag = Val(TransPictureReverser.Tag) + 1
        End If
    Else 'If Button = vbleftButton Then
        If Val(TransPictureReverser.Tag) > 0 Then
            Let MMControlReverser.Command = "Close"
            Let MMControlReverser.Filename = App.Path$ & "\Sounds\Graphics\ControlReverser.wav"
            Let MMControlReverser.Command = "Open"
            Let MMControlReverser.Command = "Play"
            Let TransPictureReverser.Tag = Val(TransPictureReverser.Tag) - 1
        End If
    End If

    Let TransPictureReverser.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\Reverser" & Trim$(TransPictureReverser.Tag) & "(d3).bmp")

    If Val(TransPictureReverser.Tag) = 2 Then
        Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
        Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
        Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked
    ElseIf Val(TransPictureReverser.Tag) = 1 Then
        Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbUnchecked
        Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbChecked
        Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
    ElseIf Val(TransPictureReverser.Tag) = 0 Then
        Let MainlineOperationGUI!ConsistControlDirectionR.Value = vbChecked
        Let MainlineOperationGUI!ConsistControlDirectionN.Value = vbUnchecked
        Let MainlineOperationGUI!ConsistControlDirectionF.Value = vbUnchecked
    End If
    
Return

End Sub


Public Sub TransPictureSand_Click()
    
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass

    If InternetSettings!Winsock.State = sckConnected Then
        If InternetSettings!textboxincomingdata.Text = "Sand, requested." Then
            Let InternetSettings!textboxoutboundcommand.Text = "Sand, confirmed."
            Call InternetSettings.ButtonSend_Click
            GoSub UpdateSand
            Let InternetSettings!textboxincomingdata.Text = ""
        ElseIf InternetSettings.textboxincomingdata.Text = "Sand, confirmed." Then
            GoSub UpdateSand
            Let InternetSettings.textboxincomingdata.Text = ""
        Else
            Let InternetSettings!textboxoutboundcommand.Text = "Sand, requested."
            Call InternetSettings.ButtonSend_Click
            Let InternetSettings!textboxincomingdata.Text = ""
        End If
    Else
        GoSub UpdateSand
    End If

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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let MainlineOperationGuiDiesel1Screen.MousePointer = ccHourglass

        If TransPictureSand.Tag = "0" Then
            Let BalloonHelpText1 = "The sand is not active."
        Else 'If TransPictureSand.Tag = "1" Then
            Let BalloonHelpText1 = "The sand is active."
        End If
        
        Let BalloonHelpText2 = "Sand"

        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureSand)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureSand, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Trans Picture Sand, unable to setup balloon help for 'TransPictureSand' control."
            End If
        End If

    End If
    
    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' =============================================================================================================================================================================
'
' =============================================================================================================================================================================

UpdateSand:

    If Val(TransPictureSand.Tag) = 0 Then
        Let TransPictureSand.Tag = "1"
        Let MMControlSand.Command = "Close"
        Let MMControlSand.Filename = App.Path$ & "\Sounds\Graphics\Sand.wav"
    Else
        Let TransPictureSand.Tag = "0"
        Let MMControlSand.Filename = App.Path$ & "\Sounds\Graphics\control.wav"
    End If
    Let MMControlSand.Command = "Open"
    Let MMControlSand.Command = "Play"
    
    Let TransPictureSand.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\Sand" & TransPictureSand.Tag & "(d3).bmp")
    Let TransPictureSandRight.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\SandRight" & TransPictureSand.Tag & "(d3).bmp")
    
Return

End Sub


Public Sub TransPictureThrottle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccHourglass
    
    If ButtonEngineStart.Caption = "&Stop Engine" Then
    
        If InternetSettings!Winsock.State = sckConnected Then
            If InternetSettings!textboxincomingdata.Text = "Throttle, left button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Throttle, left button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbLeftButton
                GoSub UpdateThrottle
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Throttle, left button confirmed." Then
                Let Button = vbLeftButton
                GoSub UpdateThrottle
                Let InternetSettings.textboxincomingdata.Text = ""
                
            ElseIf InternetSettings!textboxincomingdata.Text = "Throttle, right button requested." Then
                Let InternetSettings!textboxoutboundcommand.Text = "Throttle, right button confirmed."
                Call InternetSettings.ButtonSend_Click
                Let Button = vbRightButton
                GoSub UpdateThrottle
                Let InternetSettings!textboxincomingdata.Text = ""
                
            ElseIf InternetSettings.textboxincomingdata.Text = "Throttle, right button confirmed." Then
                Let Button = vbRightButton
                GoSub UpdateThrottle
                Let InternetSettings.textboxincomingdata.Text = ""
                
            Else
                If Button = vbLeftButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Throttle, left button requested."
                ElseIf Button = vbRightButton Then
                    Let InternetSettings!textboxoutboundcommand.Text = "Throttle, right button requested."
                End If
                Call InternetSettings.ButtonSend_Click
                Let InternetSettings!textboxincomingdata.Text = ""
            End If
        Else
            ' The program must be running with outany internet connection
            GoSub UpdateThrottle
        End If
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            MsgBox "Before making any movements on the throttle handle, you should start the locomotive engine. Then please allow for time, as the engine needs to warm up.", vbExclamation + vbOKOnly, "Automatic Train Control - Engine not Started"
        End If
    End If
    
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
        Dim BalloonHelpWaveFile As String '

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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let BalloonHelpText1 = "The application of the throttle is at " & Str$(Int(Val(100 - (TextBoxStatusAutomaticBrakePressure.Text) * 100))) & " percent."
        Let BalloonHelpText2 = "Throttle"
        
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureThrottle)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureThrottle, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Throttle, unable to setup balloon help for 'TransPictureThrottle' control."
            End If
        End If
        'Let BalloonHelpSetup = BalloonHelp.DestroyToolTip(TransPictureThrottleRight)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TransPictureThrottleRight, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel3 Screen, Picture Box Throttle Right, unable to setup balloon help for 'TransPictureThrottle' control."
            End If
        End If

    End If

    Let MainlineOperationGuiDiesel3Screen.MousePointer = ccDefault

Exit Sub

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

UpdateThrottle:

    If Button = vbLeftButton Then
        If Val(TransPictureThrottle.Tag) < 8 Then
            Let TransPictureThrottle.Tag = Val(TransPictureThrottle.Tag) + 1
            Let MMControlBrakeDynamic.Command = "Close"
            Let MMControlBrakeDynamic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            Let MMControlBrakeDynamic.Command = "Open"
            Let MMControlBrakeDynamic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the maximum throttle, one hundred (100) percent."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    ElseIf Button = vbRightButton Then
        If Val(TransPictureThrottle.Tag) > 0 Then
            Let TransPictureThrottle.Tag = Val(TransPictureThrottle.Tag) - 1
            Let MMControlBrakeDynamic.Command = "Close"
            Let MMControlBrakeDynamic.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
            Let MMControlBrakeDynamic.Command = "Open"
            Let MMControlBrakeDynamic.Command = "Play"
        Else
            If ButtonHelp.Caption = "&Help is On" Then
                Let Message = "You have reached the minimum throttle, zero (0) precent."
                MsgBox Message, vbExclamation, "ATC - Engineer Error"
            End If
        End If
    End If
    
    Let TransPictureThrottle.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\Throttle" & Trim$(TransPictureThrottle.Tag) & "(d3).bmp")
    Let TransPictureThrottleRight.Picture = LoadPicture(App.Path$ & "\Graphics\Locomotive Diesel3\ThrottleRight" & Trim$(TransPictureThrottle.Tag) & "(d3).bmp")
    
Return

End Sub


Private Sub VideoCapture_ConnectionClosed(ByVal RemoteAddress As String, ByVal ErrorCode As Long)

    Call VideoModule.ConnectionClosed(RemoteAddress, ErrorCode)
    
End Sub


Private Sub VideoCapture_ConnectionRequest(ByVal FromAddress As String, Cancel As Long)

    Call VideoModule.ConnectionRequest(FromAddress, Cancel)
    
End Sub


Private Sub VideoCapture_DeviceLost()

    Call VideoModule.DeviceLost
    
End Sub
    


