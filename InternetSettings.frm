VERSION 4.00
Begin VB.Form InternetSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Internet Settings"
   ClientHeight    =   9660
   ClientLeft      =   3435
   ClientTop       =   2250
   ClientWidth     =   6675
   Height          =   10065
   Icon            =   "InternetSettings.frx":0000
   Left            =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   6675
   Top             =   1905
   Width           =   6795
   Begin VB.ComboBox ComboServerName 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "InternetSettings.frx":0442
      Left            =   60
      List            =   "InternetSettings.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   3060
      Width           =   6555
   End
   Begin VB.CheckBox CheckboxAdditionalAudioStream 
      Caption         =   "An additional live audio stream with your video/audio from the locomotive is available."
      Height          =   255
      Left            =   60
      TabIndex        =   49
      Top             =   7620
      Value           =   1  'Checked
      Width           =   6555
   End
   Begin VB.CheckBox CheckboxClientStreamTypeBroadcast 
      Caption         =   "Broadcast Mode streaming mode is used."
      Height          =   315
      Left            =   3060
      TabIndex        =   48
      Top             =   6120
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin VB.CheckBox CheckboxClientStreamTypeServer 
      Caption         =   "Server Mode or"
      Height          =   315
      Left            =   1620
      TabIndex        =   47
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Timer TimerCheckWinsock 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   7920
      Top             =   6600
   End
   Begin VB.CommandButton ButtonRoomLightingControl 
      Caption         =   "&Light Settings"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   36
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton ButtonAudioSettings 
      Caption         =   "&Audio Settings"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   33
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonVideoSettings 
      Caption         =   "&Video Settings"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton OptionStandAlone 
      Caption         =   "Stand alone (operates a niether a host or client)"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   2160
      Value           =   -1  'True
      Width           =   3975
   End
   Begin VB.Timer TimerScreenCapture 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   7920
      Top             =   5460
   End
   Begin VB.CheckBox CheckBoxNetConnectionViaProxy 
      Caption         =   "Proxy server is used to connect to the internet."
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   4800
      Width           =   5175
   End
   Begin VB.CheckBox CheckBoxNetConnectionViaModem 
      Caption         =   "A modem is used to connect to the internet."
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CheckBox CheckBoxNetConnectionViaLan 
      Caption         =   "Local Area network used for this internet connection."
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   4320
      Width           =   5175
   End
   Begin VB.OptionButton OptionClient 
      Caption         =   "Client (remote user of the software)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1900
      Width           =   3495
   End
   Begin VB.OptionButton OptionHost 
      Caption         =   "Host (connected to a DCC layout)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton ButtonAutoListen 
      Caption         =   "&Auto Listen Off"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   3900
      Width           =   1215
   End
   Begin VB.TextBox TextboxIncomingData 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton ButtonSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   13200
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TextboxOutBoundCommand 
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Text            =   "Enter a command here."
      Top             =   3120
      Width           =   4815
   End
   Begin VB.TextBox TextboxOutBoundData 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   975
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "InternetSettings.frx":0446
      Top             =   1680
      Width           =   6495
   End
   Begin VB.TextBox TextboxInBoundData 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   975
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "InternetSettings.frx":0463
      Top             =   360
      Width           =   6495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "InternetSettings.frx":0480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton ButtonDisconnect 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   4500
      Width           =   1215
   End
   Begin VB.CommandButton ButtonListen 
      Caption         =   "&Listen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton ButtonConnect 
      Caption         =   "&Connect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet InternetTransferControl 
      Left            =   7920
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Server Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   54
      Top             =   2520
      Width           =   2595
   End
   Begin VB.Line Line8 
      X1              =   60
      X2              =   6600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label LabelServerName 
      Caption         =   "Which server would you like to connect to?"
      Height          =   195
      Left            =   120
      TabIndex        =   53
      Top             =   2820
      Width           =   4635
   End
   Begin FILETRANSXLib.FileTransX FtpControl 
      Height          =   480
      Left            =   7860
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Timeout         =   30
      ServerName      =   "atc.lovethosetrains.com"
      Username        =   "softwarebmp"
      Password        =   "walnuttree12"
      ProxyName       =   ""
      ProxyUserID     =   ""
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7860
      Top             =   8520
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VIDEOCAPXLib.VideoCapX VideoCapture 
      Height          =   435
      Left            =   7920
      TabIndex        =   50
      Top             =   9720
      Width           =   435
      _Version        =   131072
      _ExtentX        =   767
      _ExtentY        =   767
      _StockProps     =   0
      CapFilename     =   ""
   End
   Begin VB.Line Line7 
      X1              =   6660
      X2              =   60
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Label Label11 
      Caption         =   "While in client mode, "
      Height          =   195
      Left            =   60
      TabIndex        =   46
      Top             =   6180
      Width           =   1515
   End
   Begin VB.Label Label10 
      Caption         =   "Winsock Control - For making connection to other computers."
      Height          =   375
      Left            =   8400
      TabIndex        =   45
      Top             =   9060
      Width           =   4755
   End
   Begin VB.Label Label9 
      Caption         =   "Balloon Help Control - For displaying balloons on objects"
      Height          =   375
      Left            =   8460
      TabIndex        =   44
      Top             =   8520
      Width           =   4755
   End
   Begin VB.Label Label8 
      Caption         =   "Ini Control - TO pdate ATC.INI file."
      Height          =   375
      Left            =   8520
      TabIndex        =   43
      Top             =   7920
      Width           =   4755
   End
   Begin VB.Label Label7 
      Caption         =   "AlphaBlend Control - For transparency effects of the window."
      Height          =   375
      Left            =   8520
      TabIndex        =   42
      Top             =   7320
      Width           =   4755
   End
   Begin VB.Label Label6 
      Caption         =   "NetConnect Control - For indicating presents of the Internet."
      Height          =   375
      Left            =   8520
      TabIndex        =   41
      Top             =   4260
      Width           =   4755
   End
   Begin VB.Label Label5 
      Caption         =   "Capture Control - For capturing the screen."
      Height          =   375
      Left            =   8520
      TabIndex        =   40
      Top             =   4920
      Width           =   4755
   End
   Begin VB.Label Label4 
      Caption         =   "Timer Control - For uploading ATC screen capture to the server."
      Height          =   375
      Left            =   8520
      TabIndex        =   39
      Top             =   5460
      Width           =   4755
   End
   Begin VB.Label Label2 
      Caption         =   "FTP Control - For uploading ATC screen capture to the server."
      Height          =   375
      Left            =   8580
      TabIndex        =   38
      Top             =   6060
      Width           =   4755
   End
   Begin VB.Label Label1 
      Caption         =   "Timer - For checking internet connection"
      Height          =   435
      Left            =   8580
      TabIndex        =   37
      Top             =   6660
      Width           =   4695
   End
   Begin VB.Label LabelRoomLightingSettingsDescription 
      Caption         =   $"InternetSettings.frx":08C2
      Height          =   615
      Left            =   60
      TabIndex        =   35
      Top             =   8460
      Width           =   6495
   End
   Begin VB.Label LabelRoomLightingSettings 
      Caption         =   "Room Lighting Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   34
      Top             =   8100
      Width           =   2115
   End
   Begin VB.Line Line6 
      X1              =   60
      X2              =   6600
      Y1              =   7980
      Y2              =   7980
   End
   Begin VB.Label LabelAudioSettingsDescription 
      Caption         =   $"InternetSettings.frx":0998
      Height          =   435
      Left            =   60
      TabIndex        =   32
      Top             =   7020
      Width           =   6555
   End
   Begin VB.Label LabelAudioSettings 
      Caption         =   "Audio Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   31
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   6600
      X2              =   60
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label LabelVideoSettingsMore 
      Caption         =   $"InternetSettings.frx":0A25
      Height          =   435
      Left            =   60
      TabIndex        =   30
      Top             =   5580
      Width           =   6555
   End
   Begin VB.Label LabelVideoSettings 
      Caption         =   "VideoSettings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   29
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status: Idle"
      Height          =   255
      Left            =   60
      TabIndex        =   26
      Top             =   960
      Width           =   6555
   End
   Begin MyScreenCapture.MyCapture CaptureOcx 
      Left            =   7920
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   661
      Version         =   ""
      Filename        =   ""
   End
   Begin VB.Line Line4 
      X1              =   6540
      X2              =   60
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label LabelHostOrClient 
      AutoSize        =   -1  'True
      Caption         =   "This software and connection to the internet is acting as:"
      Height          =   195
      Left            =   60
      TabIndex        =   20
      Top             =   1440
      Width           =   4005
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   6600
      Y1              =   2400
      Y2              =   2400
   End
   Begin A16IPTextBox.IPTextBox IpTextBoxClient 
      Height          =   375
      Left            =   2940
      TabIndex        =   19
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
      LicenceCode     =   "2479963878265856"
      LicenceName     =   "Canadian Locomotive Logistics"
   End
   Begin VB.Label LabelIpAddressClient 
      Caption         =   "IP Address for Client"
      Height          =   195
      Left            =   2940
      TabIndex        =   18
      Top             =   3600
      Width           =   1605
   End
   Begin April16_NetConnect.NetConnect NetConnect 
      Left            =   7920
      Top             =   4260
      _ExtentX        =   873
      _ExtentY        =   873
      LicenceCode     =   "7134424019415563"
      LicenceName     =   "Canadian Locomotive Logistics"
   End
   Begin A16IPTextBox.IPTextBox IpTextBoxHost 
      Height          =   375
      Left            =   60
      TabIndex        =   17
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
      LicenceCode     =   "2479963878265856"
      LicenceName     =   "Canadian Locomotive Logistics"
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7920
      Top             =   7260
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7920
      Top             =   7920
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   14400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label LabelCommandToSend 
      Caption         =   "Command to Send"
      Height          =   195
      Left            =   7920
      TabIndex        =   10
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label LabelOutBoundData 
      Caption         =   "Out Bound Data"
      Height          =   195
      Left            =   7920
      TabIndex        =   8
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "In Bound Data"
      Height          =   195
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label LabelIpAddressHost 
      Caption         =   "IP Address for Host"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label LabelIntroduction 
      Caption         =   $"InternetSettings.frx":0AB2
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7920
      Top             =   9120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   20101
   End
End
Attribute VB_Name = "InternetSettings"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub ButtonAudioSettings_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Audio Settings, Click" & vbCrLf
    End If ' Debug Tag

    Load AudioSettings
    AudioSettings.Show vbModeless

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Audio Settings, Click" & vbCrLf
    End If ' Debug Tag

End Sub


Public Sub ButtonAutoListen_Click()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Auto Listen, Click" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Toggle Caption on Button
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If ButtonAutoListen.Caption = "&Auto Listen Off" Then
        Let ButtonAutoListen.Caption = "&Auto Listen On"
    ElseIf ButtonAutoListen.Caption = "&Auto Listen On" Then
        Let ButtonAutoListen.Caption = "&Auto Listen Off"
    End If
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Auto Listen, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub ButtonClose_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Close, Click" & vbCrLf
    End If ' Debug Tag
    
    Let OptionStandAlone.Value = True

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
            If TemporaryScreen = "Internet Settings Screen" Then
                Let Ini.Value = "Unused"

            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then

                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbExclamation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Internet Settings Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Internet Settings Screen, Button Close, stack is empty, underflow."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button CLose, Click" & vbCrLf
    End If ' Debug Tag
    
    End Sub

Private Sub ButtonConnect_Click()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Connect, Click" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Buttons for User
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let optionHost.Enabled = False
    Let optionclient.Enabled = False
    Let OptionStandAlone.Enabled = False
    Let buttonconnect.Enabled = False
    Let buttondisconnect.Enabled = True
    Let ButtonVideoSettings.Enabled = False
    Let ButtonClose.Enabled = False
    Let ComboServerName.Enabled = False
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialize Winsock
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Winsock.RemoteHost = "atc.server" & CStr(Val(InternetSettings!ComboServerName.ListIndex)) & ".lovethosetrains.com"
    Let Winsock.RemotePort = 20101

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Winsock Connection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Winsock.Connect
    Call LabelSockUpdate
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Connect, Click" & vbCrLf
    End If ' Debug Tag
        
End Sub

Public Sub ButtonDisconnect_Click()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Disconnect, Click" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Close the Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Winsock.State <> sckClosed Then
        Winsock.Close
    End If

    While Winsock.State <> sckClosed
        Call LabelSockUpdate
        DoEvents
    Wend
    
    Call LabelSockUpdate
    Let TimerCheckWinsock.Enabled = False
    Let buttondisconnect.Enabled = False
    Let ButtonClose.Enabled = True
    
    If optionHost.Value = True Then
        Call OptionHost_Click
    ElseIf optionclient.Value = True Then
        Call OptionClient_Click
    Else
        Call OptionStandAlone_Click
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Reset Network Connection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If NetConnect.Connected = True Then
        Let optionHost.Enabled = True
        Let optionclient.Enabled = True
        Let OptionStandAlone.Enabled = True
        If NetConnect.NetConnectionViaLAN = True Then
            Let CheckBoxNetConnectionViaLan.Value = 1
        ElseIf NetConnect.NetConnectionViaLAN = False Then
            Let CheckBoxNetConnectionViaLan.Value = 0
        End If
        If NetConnect.NetConnectionViaModem = True Then
            Let CheckBoxNetConnectionViaModem.Value = 1
        ElseIf NetConnect.NetConnectionViaModem = False Then
            Let CheckBoxNetConnectionViaModem.Value = 0
        End If
            If NetConnect.NetConnectionViaProxy = True Then
            Let CheckBoxNetConnectionViaProxy.Value = 1
        ElseIf NetConnect.NetConnectionViaProxy = False Then
            Let CheckBoxNetConnectionViaProxy.Value = 0
        End If
    End If

    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Disconnect, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Public Sub ButtonListen_Click()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Listen, Click" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Command Buttons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let optionHost.Enabled = False
    Let optionclient.Enabled = False
    Let OptionStandAlone.Enabled = False
    Let ButtonListen.Enabled = False
    Let ButtonAutoListen.Enabled = False
    Let buttondisconnect.Enabled = True
    Let ButtonVideoSettings.Enabled = False
    Let ButtonAudioSettings.Enabled = False
    Let ButtonRoomLightingControl.Enabled = False
    Let ButtonClose.Enabled = False

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialize Winstock
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let IpTextBoxHost.Addr = Winsock.LocalIP
    Let Winsock.LocalPort = 20101
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Let Winstock Listen on Port
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Winsock.Listen
    Call LabelSockUpdate
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display Window and Place on Standby
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainlineOperationGUI.OptionGuiDiesel1.Value = True Then
        Load MainlineOperationGuiDiesel1Screen
        MainlineOperationGuiDiesel1Screen.Show vbModeless
    ElseIf MainlineOperationGUI.OptionGuiDiesel2.Value = True Then
        Load MainlineOperationGuiDiesel2Screen
        MainlineOperationGuiDiesel2Screen.Show vbModeless
    ElseIf MainlineOperationGUI.OptionGuiDiesel3.Value = True Then
        Load MainlineOperationGuiDiesel3Screen
        MainlineOperationGuiDiesel3Screen.Show vbModeless
    ElseIf MainlineOperationGUI.OptionGuiSteam1.Value = True Then
        Load MainlineOperationGuiSteam1Screen
        MainlineOperationGuiSteam1Screen.Show vbModeless
    Else 'MainlineOperationGUI.OptionGuiElectric1.Value = True Then
        Load MainlineOperationGuiElectric1Screen
        MainlineOperationGuiElectric1Screen.Show vbModeless
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Timer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let MainlineOperationGUI.TimerWindowState.Interval = 65535
    Let MainlineOperationGUI.TimerWindowState.Enabled = True
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Listen, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub ButtonPrint_Click()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Print, Click" & vbCrLf
    End If ' Debug Tag

    InternetSettings.PrintForm
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Print, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub ButtonRoomLightingControl_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Room Lighting Control, Click" & vbCrLf
    End If ' Debug Tag

    Load RoomLightingControl
    RoomLightingControl.Show vbModeless
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Room Lighting Control, Click" & vbCrLf
    End If ' Debug Tag

End Sub

Public Sub ButtonSend_Click()
    
    On Error Resume Next
    Winsock.SendData textboxoutboundcommand.Text
    DoEvents
    If Err <> 0 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Internet Settings Screen, Button Send, Click, " & CStr(Err) & Error
        End If
    End If
    
    On Error GoTo 0
    Let TextboxOutBoundData.Text = TextboxOutBoundData.Text + textboxoutboundcommand.Text + vbCrLf
    Let TextboxOutBoundData.SelStart = Len(TextboxOutBoundData.Text)

End Sub




Private Sub ButtonVideoSettings_Click()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Button Video Settings, Click" & vbCrLf
    End If ' Debug Tag

    Load VideoSettings
    VideoSettings.Show vbModeless

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Button Video Settings, Click" & vbCrLf
    End If ' Debug Tag

End Sub




Private Sub CaptureOcx_Error(ByVal Description As String)

       If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Internet Settings Screen, Capture OCX COntrol, " & Description & "."
        End If
End Sub


Private Sub CheckboxAdditionalAudioStream_Click()

    If CheckboxAdditionalAudioStream.Value = vbChecked Then
        If InternetSettings.NetConnect.Connected = True Then
            If InternetSettings.checkboxclientstreamtypebroadcast.Value = vbChecked Then
                VideoCapture.PlayerOpen "http://www.railroadradio.net/content/playlist/columbus.asx"
                'VideoCapture.PlayerOpen "http://railaudio2.railroadradio.net:7110"
                DoEvents
                VideoCapture.PlayerStart
            End If
        End If
    Else
        VideoCapture.PlayerStop
    End If
End Sub

Private Sub CheckboxClientStreamTypeBroadcast_Click()

    If checkboxclientstreamtypebroadcast.Value = vbChecked Then
        Let checkboxclientstreamtypeserver.Value = vbUnchecked
    End If
    
End Sub


Private Sub CheckboxClientStreamTypeServer_Click()

    If checkboxclientstreamtypeserver.Value = vbChecked Then
        Let checkboxclientstreamtypebroadcast.Value = vbUnchecked
    End If

End Sub


Private Sub ComboServerName_Click()

    If ComboServerName.ListIndex <> 0 Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' User Has Selected a Server
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If optionHost.Value = True Then
            Let buttonconnect.Enabled = False
            Let ButtonListen.Enabled = True
            Let ButtonAutoListen.Enabled = True
            Let buttonconnect.Enabled = False
            Let buttondisconnect.Enabled = False
            Let ButtonVideoSettings.Enabled = True
            Let ButtonAudioSettings.Enabled = True
            Let ButtonRoomLightingControl.Enabled = True
            Let ButtonSend.Enabled = False
        ElseIf optionclient.Value = True Then
            Let ButtonListen.Enabled = False
            Let ButtonAutoListen.Enabled = False
            Let buttonconnect.Enabled = True
            Let buttondisconnect.Enabled = False
            Let ButtonVideoSettings.Enabled = False
            Let ButtonAudioSettings.Enabled = False
            Let ButtonRoomLightingControl.Enabled = False
            Let ButtonSend.Enabled = False
        Else
            Let buttonconnect.Enabled = False
            Let ButtonListen.Enabled = False
            Let ButtonAutoListen.Enabled = False
            Let buttonconnect.Enabled = False
            Let buttondisconnect.Enabled = False
            Let ButtonVideoSettings.Enabled = True
            Let ButtonAudioSettings.Enabled = True
            Let ButtonRoomLightingControl.Enabled = False
            Let ButtonSend.Enabled = False
        End If
    End If

End Sub


Private Sub Form_Activate()

    DoEvents
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Form, Activate" & vbCrLf
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
        If TemporaryScreen = "Internet Settings Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Internet Settings Screen"
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
            Let Ini.Value = "Internet Settings Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Internet Settings Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Internet Settings Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Properties to Objects
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InternetSettings.NetConnect.Connected = False Then
        Let optionHost.Enabled = False
        Let optionclient.Enabled = False
        Let OptionStandAlone.Enabled = False
        Let CheckBoxNetConnectionViaLan.Enabled = False
        Let CheckBoxNetConnectionViaModem.Enabled = False
        Let CheckBoxNetConnectionViaProxy.Enabled = False
        Let ButtonListen.Enabled = False
        Let ButtonAutoListen.Enabled = False
        Let buttonconnect.Enabled = False
        Let buttondisconnect.Enabled = False
        Let ButtonSend.Enabled = False
    Else
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Get External IP Address
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        On Error Resume Next
        Let TemporaryString = InternetTransferControl.OpenURL("http://vbnet.mvps.org/resources/tools/getpublicip.shtml", icString)
        If Err = 0 Then
            Let TemporaryPosition1 = InStr(TemporaryString, "var ip =")
            Let TemporaryPosition1 = InStr(TemporaryPosition1 + 9, TemporaryString, "'") + 1
            Let TemporaryPosition2 = InStr(TemporaryPosition1 + 9, TemporaryString, "'")
            Let IpTextBoxClient.Addr = Mid$(TemporaryString, TemporaryPosition1, TemporaryPosition2 - TemporaryPosition1)
        Else
            Let IpTextBoxClient.Addr = Winsock.LocalIP
        End If
        On Error GoTo 0
        
        Let optionHost.Enabled = True
        Let optionclient.Enabled = True
        Let OptionStandAlone.Enabled = True
        Let CheckBoxNetConnectionViaLan.Enabled = True
        Let CheckBoxNetConnectionViaModem.Enabled = True
        Let CheckBoxNetConnectionViaProxy.Enabled = True
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Close the Port
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If Winsock.State <> sckClosed Then
            Winsock.Close
        End If

        While Winsock.State <> sckClosed
            Call LabelSockUpdate
            DoEvents
        Wend
        Call LabelSockUpdate
        Let TimerCheckWinsock.Enabled = False
        Let buttondisconnect.Enabled = False
        Let ButtonClose.Enabled = True
    
        If optionHost.Value = True Then
            Call OptionHost_Click
        ElseIf optionclient.Value = True Then
            Call OptionClient_Click
        Else
            Call OptionStandAlone_Click
        End If
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Reset Network Connection
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If NetConnect.Connected = True Then
            Let optionHost.Enabled = True
            Let optionclient.Enabled = True
            Let OptionStandAlone.Enabled = True
            If NetConnect.NetConnectionViaLAN = True Then
                Let CheckBoxNetConnectionViaLan.Value = 1
            ElseIf NetConnect.NetConnectionViaLAN = False Then
                Let CheckBoxNetConnectionViaLan.Value = 0
            End If
            If NetConnect.NetConnectionViaModem = True Then
                Let CheckBoxNetConnectionViaModem.Value = 1
            ElseIf NetConnect.NetConnectionViaModem = False Then
                Let CheckBoxNetConnectionViaModem.Value = 0
            End If
                If NetConnect.NetConnectionViaProxy = True Then
                Let CheckBoxNetConnectionViaProxy.Value = 1
            ElseIf NetConnect.NetConnectionViaProxy = False Then
                Let CheckBoxNetConnectionViaProxy.Value = 0
            End If
        End If
    End If
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Balloon Help
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call BalloonHelpUpdatePart01
    Call BalloonHelpUpdatePart02

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Automatic Listening
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If optionHost.Value = True Then
        If ButtonAutoListen.Caption = "&Auto Listen On" Then
            Dim TemporaryTime As Date
            Dim TemporaryLoop As Integer
            Let TemporaryTime = Now
            Let LabelStatus.Caption = "Status: Winsock is waiting before listening again."
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Wait 10 seconds
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Internet Settings Screen"
            Let Ini.Parameter = "RecycleDelayTime"
            
            While DateDiff("s", TemporaryTime, Now) < Val(Ini.Value)
                Let LabelStatus.Caption = "Status: Winsock is waiting " & Str(Val(Ini.Value) - Val(DateDiff("s", TemporaryTime, Now))) & " seconds."
                DoEvents
            Wend
            Let LabelStatus.Caption = "Status:"
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Start Listening If
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If ButtonAutoListen.Caption = "&Auto Listen On" Then
                Call ButtonListen_Click
            End If
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Do Not Recycle
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Else
            LabelStatus.Caption = "Status: Winstock is idle, automatic listening is turned off."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Form, Activate" & vbCrLf
    End If ' Debug Tag

End Sub

Private Sub Form_Deactivate()

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Form, Deactivate" & vbCrLf
    End If ' Debug Tag

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Internet Settings Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(InternetSettings.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(InternetSettings.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(InternetSettings.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(InternetSettings.Height)
    Let Ini.Parameter = "OptionHost"
    Let Ini.Value = optionHost.Value
    Let Ini.Parameter = "OptionClient"
    Let Ini.Value = optionclient.Value
    Let Ini.Parameter = "OptionStandAlone"
    Let Ini.Value = OptionStandAlone.Value
    Let Ini.Parameter = "ComboServerName"
    Let Ini.Value = ComboServerName.Text
    Let Ini.Parameter = "IpAddressHost"
    Let Ini.Value = IpTextBoxHost.Addr
    Let Ini.Parameter = "IpAddressClient"
    Let Ini.Value = IpTextBoxClient.Addr
    Let Ini.Parameter = "ClientStreamTypeServer"
    Let Ini.Value = checkboxclientstreamtypeserver.Value
    Let Ini.Parameter = "ClientStreamTYpeBroadcast"
    Let Ini.Value = checkboxclientstreamtypebroadcast.Value
    Let Ini.Parameter = "AdditionalAudioStream"
    Let Ini.Value = CheckboxAdditionalAudioStream.Value
    
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
                Let Ini.Value = "Internet Settings Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Internet Settings Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    InternetSettings.Hide

   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Form, Deactivate" & vbCrLf
    End If ' Debug Tag
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub


Private Sub Form_Load()


   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Internet Settings, Form, Load" & vbCrLf
    End If ' Debug Tag
 
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
    Let Ini.Application = "Internet Settings Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    'Let Ini.Parameter = "OptionHost"
    'Let OptionHost.Value = Ini.Value
    'Let Ini.Parameter = "OptionClient"
    'Let OptionClient.Value = Ini.Value
    'Let Ini.Parameter = "OptionStandAlone"
    'Let OptionStandAlone.Value = Ini.Value
    'Let ini.parameter = "ComboServerName"
    'let ComboServerName.Text = Ini.Value
    Let Ini.Parameter = "IpAddressHost"
    Let IpTextBoxHost.Addr = Ini.Value
    Let Ini.Parameter = "IpAddressClient"
    Let IpTextBoxClient.Addr = Ini.Value
    Let Ini.Parameter = "ClientStreamTypeServer"
    Let checkboxclientstreamtypeserver.Value = Ini.Value
    Let Ini.Parameter = "ClientStreamTypeBroadcast"
    Let checkboxclientstreamtypebroadcast.Value = Ini.Value
    Let Ini.Parameter = "AdditionalAudioStream"
    Let CheckboxAdditionalAudioStream.Value = Ini.Value
    Let Ini.Parameter = "MaximumBandwidth"
    Let FtpControl.MaxBandwidth = Ini.Value
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        InternetSettings.Left = (Screen.Width - Width) / 2
        InternetSettings.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + InternetSettings.Width > Screen.Width Then
            Let InternetSettings.Left = Screen.Width - InternetSettings.Width
        Else
            Let InternetSettings.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + InternetSettings.Height > Screen.Height Then
            Let InternetSettings.Top = Screen.Height - InternetSettings.Height
        Else
            Let InternetSettings.Top = Val(TemporaryValueTop)
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
' Screen Capture Parameters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Internet Settings Screen"
    Let Ini.Parameter = "TimerScreenCapture"
    Let TimerScreenCapture.Interval = Ini.Value

    Let CaptureOcx.Filename = App.Path$ & "\Atc.bmp"
    'Let FtpControl.SrcFilename = App.Path$ & "\Atc.bmp"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Network Connection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If NetConnect.Connected = True Then
        Let optionHost.Enabled = True
        Let optionclient.Enabled = True
        If NetConnect.NetConnectionViaLAN = True Then
            Let CheckBoxNetConnectionViaLan.Value = 1
        ElseIf NetConnect.NetConnectionViaLAN = False Then
            Let CheckBoxNetConnectionViaLan.Value = 0
        End If
        If NetConnect.NetConnectionViaModem = True Then
            Let CheckBoxNetConnectionViaModem.Value = 1
        ElseIf NetConnect.NetConnectionViaModem = False Then
            Let CheckBoxNetConnectionViaModem.Value = 0
        End If
            If NetConnect.NetConnectionViaProxy = True Then
            Let CheckBoxNetConnectionViaProxy.Value = 1
        ElseIf NetConnect.NetConnectionViaProxy = False Then
            Let CheckBoxNetConnectionViaProxy.Value = 0
        End If
    
    End If
 
 ' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 ' Update Server Combination Box
 ' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim TemporaryCounter As Integer
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Internet Settings Screen"
    Let Ini.Parameter = "ServerCount"
    ComboServerName.Clear
    ComboServerName.AddItem "Please select a server (train layout) to connect to."
    For TemporaryCounter = 1 To Val(Ini.Value)
        Let Ini.Parameter = "Server" & Right$(Val(TemporaryCounter), 1) & "Name"
        ComboServerName.AddItem Ini.Value
    Next
    Let Ini.Parameter = "ComboServerName"
    Let ComboServerName.Text = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Internet Settings, Form, Load" & vbCrLf
    End If ' Debug Tag

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub















Private Sub Form_Resize()

    If InternetSettings.WindowState = vbMinimized Then
    
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
        
    ElseIf InternetSettings.WindowState = vbNormal Then
    
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

Private Sub FtpControl_Timeout()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Internet Settings Screen, Ftp Control, connection to server timed out."
    End If
    
End Sub

Private Sub FtpControl_TransferComplete(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal BytesTransfered As Long)
    
    Let LabelStatus.Caption = "Status: Transfer complete, " & BytesTransfered & " bytes."
    Let TimerScreenCapture.Enabled = True

    
End Sub

Private Sub FtpControl_TransferProgress(ByVal Bytes As Long)

    Let LabelStatus.Caption = "Status: Transfering " & Bytes & " Bytes."
    
End Sub

Private Sub FtpControl_TransferStarting(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal Offset As Long)

    Let LabelStatus.Caption = "Status: File Transfer Starting."
    
End Sub



Private Sub NetConnect_Connected()
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Enable Object if Internet
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let optionHost.Enabled = True
    Let optionclient.Enabled = True
    Let OptionStandAlone.Enabled = True
    Let CheckBoxNetConnectionViaLan.Enabled = True
    Let CheckBoxNetConnectionViaModem.Enabled = True
    Let CheckBoxNetConnectionViaProxy.Enabled = True

End Sub

Private Sub NetConnect_Disconnected()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' No Uploading
    Let TimerCapture.Enabled = False
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Close Port Connection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call ButtonDisconnect_Click
    'Winsock.Close
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Buttons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let optionHost.Enabled = False
    Let optionclient.Enabled = False
    Let OptionStandAlone.Enabled = False
    Let CheckBoxNetConnectionViaLan.Enabled = False
    Let CheckBoxNetConnectionViaModem.Enabled = False
    Let CheckBoxNetConnectionViaProxy.Enabled = False
    Let ButtonListen.Enabled = False
    Let ButtonAutoListen.Enabled = False
    Let buttonconnect.Enabled = False
    Let buttondisconnect.Enabled = False
    Let ButtonVideoSettings.Enabled = False
    Let ButtonAudioSettings.Enabled = False
    Let ButtonRoomLightingControl.Enabled = False
    Let ComboServerName.Enabled = False
    Let ButtonSend.Enabled = False

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Label
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelStatus.Caption = "Status: Internet Connection Closed"
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Internet Settings Screen, NetConnect, Disconnected, this computer does not have an internet connection."
    End If


End Sub


Private Sub OptionClient_Click()
     
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Get External IP Address
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Let TemporaryString = InternetTransferControl.OpenURL("http://vbnet.mvps.org/resources/tools/getpublicip.shtml", icString)
    If Err = 0 Then
        Let TemporaryPosition1 = InStr(TemporaryString, "var ip =")
        Let TemporaryPosition1 = InStr(TemporaryPosition1 + 9, TemporaryString, "'") + 1
        Let TemporaryPosition2 = InStr(TemporaryPosition1 + 9, TemporaryString, "'")
        Let IpTextBoxClient.Addr = Mid$(TemporaryString, TemporaryPosition1, TemporaryPosition2 - TemporaryPosition1)
    Else
        Let IpTextBoxClient.Addr = Winsock.LocalIP
    End If
    On Error GoTo 0
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Pick a Server
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let ComboServerName.Enabled = True
    Call ComboServerName_Click
    
End Sub

Private Sub OptionHost_Click()
 
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Limit the User for Options
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuCommunicationDCC.Caption = "Communication to DCC is &Closed" Then
        MsgBox "To host an internet session, the communicaton port" & vbCrLf & "to the digital command control device must be turn on.", vbExclamation + vbOKOnly, "Autoamtic Train Control - Warning"
        Let OptionStandAlone.Value = True
        Call OptionStandAlone_Click
    Else
        If InternetSettings.VideoCapture.GetVideoDeviceCount > 0 Then
            If optionHost.Value = True Then
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' Get External IP Address
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                On Error Resume Next
                Let TemporaryString = InternetTransferControl.OpenURL("http://vbnet.mvps.org/resources/tools/getpublicip.shtml", icString)
                If Err = 0 Then
                    Let TemporaryPosition1 = InStr(TemporaryString, "var ip =")
                    Let TemporaryPosition1 = InStr(TemporaryPosition1 + 9, TemporaryString, "'") + 1
                    Let TemporaryPosition2 = InStr(TemporaryPosition1 + 9, TemporaryString, "'")
                    Let IpTextBoxHost.Addr = Mid$(TemporaryString, TemporaryPosition1, TemporaryPosition2 - TemporaryPosition1)
                Else
                    Let IpTextBoxHost.Addr = Winsock.LocalIP
                    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        Let Ini.Filename = App.Path$ & "\Atc.log"
                        Let Ini.Application = "Log Errors"
                        Let Ini.Parameter = Date$ & " " & Time$
                        Let Ini.Value = "Internet Settings Screen, OptionHost Activate, " & CStr(Err) & " " & Error
                    End If
                End If
                On Error GoTo 0
                Let IpTextBoxClient.Addr = "000.000.000.000"
                Let ComboServerName.Enabled = True
                Call ComboServerName_Click
            End If
        Else
            MsgBox "To host an internet session, a video capture device" & vbCrLf & "must be connected to the computer.", vbExclamation + vbOKOnly, "Autoamtic Train Control - Warning"
            Let OptionStandAlone.Value = True
            Call OptionStandAlone_Click
        End If
    End If
End Sub



Private Sub OptionStandAlone_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Limit the User for Options
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let ComboServerName.Enabled = False
    Call ComboServerName_Click
  
End Sub





Private Sub TimerCheckWinsock_Timer()

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Log the Error
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Internet Settings Screen, Timer Check Winsock, event fired. No response from client in prescribed time, close GUI, if open."
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let TimerCheckWinsock.Enabled = False
        
    DoEvents
    
    If Screen.ActiveForm.Name = "MainlineOperationGuiDiesel1Screen" Or _
       Screen.ActiveForm.Name = "MainlineOperationGuiDiesel2Screen" Then
        If Screen.ActiveForm.ButtonEngineStart.Caption = "&Stop Engine" Then
            Call Screen.ActiveForm.ButtonEngineStart_Click
        End If
        Call Screen.ActiveForm.ButtonCloseGUI_Click
    End If
    
    Call ButtonDisconnect_Click
    
End Sub

Private Sub TimerScreenCapture_Timer()

    Let TimerScreenCapture.Enabled = False

    
    If NetConnect.Connected = True Then
        If optionclient.Value = True Then
            If MainlineOperationGuiDiesel1Screen.Visible = True Then
                CaptureOcx.CaptureActiveWindows
                'mSaveToJPEG
                On Error Resume Next
                FtpControl.Connect
                'Let FtpControl.BinaryMode = 1
                FtpControl.Put App.Path$ & "/Atc.bmp", "/Atc.bmp", 0
                On Error GoTo 0
            End If
        End If
    End If


End Sub


Private Sub Winsock_Close()

    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Internet Settings Screen, Winsock Close, event fired. No real error, just close wincock."
    End If
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Winsock Connection Closed
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Note - this event is only fired when 'disconnect' is executed. So, if connection is dropped then, nothing happens
    Let ButtonClose.Caption = "&Close"
    Call LabelSockUpdate

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' DoEvents
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '   Do evetns must remain to allow the program to close the GUI screen (deactivate event) if closed by the buttonclose event, either client or host.
    
'    DoEvents
    
    If Screen.ActiveForm.Name = "MainlineOperationGuiDiesel1Screen" Or _
       Screen.ActiveForm.Name = "MainlineOperationGuiDiesle2Screen" Then
        If Screen.ActiveForm.ButtonEngineStart.Caption = "&Stop Engine" Then
            Call Screen.ActiveForm.ButtonEngineStart_Click
        End If
        Call Screen.ActiveForm.ButtonCloseGUI_Click
    End If
    
    Let TimerScreenCapture.Enabled = False
    Let TimerCheckWinsock.Enabled = False
    Call ButtonDisconnect_Click
End Sub

Private Sub Winsock_Connect()

    Let ButtonClose.Caption = "&Hide"
    Call LabelSockUpdate
    Let IpTextBoxHost.Addr = Winsock.RemoteHostIP
    Let IpTextBoxClient.Addr = Winsock.LocalIP

End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Connection Requested
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Winsock.State <> sckClosed Then Winsock.Close
    Winsock.Accept requestID
    Call LabelSockUpdate
    Let buttonconnect.Enabled = False
    Let ButtonListen.Enabled = False
    Let buttondisconnect.Enabled = True

    If optionHost.Value = True Then
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Tell Client Which GUI
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let textboxoutboundcommand.Text = Screen.ActiveForm.Name
    
    'If MainlineOperationGUI.OptionGuiDiesel1.Value = True Then
    '    Let textboxoutboundcommand.Text = "MainlineOperationGuiDiesel1Screen"
    'ElseIf MainlineOperationGUI.OptionGuiDiesel2.Value = True Then
    '    Let textboxoutboundcommand.Text = "MainlineOperationGuiDiesel2Screen"
    'ElseIf MainlineOperationGUI.OptionGuiDiesel3.Value = True Then
    '    Let textboxoutboundcommand.Text = "MainlineOperationGuiDiesel3Screen"
    'ElseIf MainlineOperationGUI.OptionGuiSteam1.Value = True Then
    '    Let textboxoutboundcommand.Text = "MainlineOperationGuiSteam1Screen"
    'Else 'MainlineOperationGUI.OptionGuiElectric1.Value = True Then
    '    Let textboxoutboundcommand.Text = "MainlineOperationGuiElectric1Screen"
    'End If
    
    Winsock.SendData textboxoutboundcommand.Text
    Let textboxoutboundcommand.Text = ""
    DoEvents
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up speed in 128 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainlineOperationGUI.ConsistControlSpeed128.Value = vbChecked Then
            Let TemporarySpeedStep = 1
            Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Speed 1 of 128 (Estop)"
            'If ConsistControlDirectionF.Value = vbChecked Then
                TemporarySpeedStep = TemporarySpeedStep + 128 ' add forward direction
                Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & " in forward direction."
            'Else
            '    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & " in reverse direction."
            'End If
            
            Let MainScreen!ThreeByteD.Text = 63
            Let MainScreen!FourByteD.Text = TemporarySpeedStep
            Let MainScreen!FiveByteD.Text = ""
            Let MainScreen!SixByteD.Text = ""
            Let MainScreen!sevenbyted.Text = ""
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up speed in 28 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Else
            Let TemporarySpeedStep = 64
            'If ConsistControlDirectionF.Value = vbChecked Then
                    Let TemporarySpeedStep = TemporarySpeedStep + 32 ' add forward direction
            'End If
            
           If ConsistControlSpeed28.Value = vbChecked Then
                Let temp1 = 1 ' adds the speed
                Let temp2 = temp1 Mod 2
                Let newspeedvalue = Int(temp1 / 2)
                Let TemporarySpeedStep = TemporarySpeedStep + newspeedvalue
                If temp2 = 1 Then Let TemporarySpeedStep = TemporarySpeedStep + 16
                Let MainScreen!ThreeByteD.Text = TemporarySpeedStep
                Let MainScreen!FourByteD.Text = ""
                Let MainScreen!FiveByteD.Text = ""
                Let MainScreen!SixByteD.Text = ""
                Let MainScreen!sevenbyted.Text = ""
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set up spped for 14 steps
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Else
                ' This routing assembles the byte for speed step mode 14
                
                Let TemporarySpeedStep = TemporarySpeedStep + 0 ' add the speed
                Let MainScreen!ThreeByteD.Text = TemporarySpeedStep
                Let MainScreen!FourByteD.Text = ""
                Let MainScreen!FiveByteD.Text = ""
                Let MainScreen!SixByteD.Text = ""
                Let MainScreen!sevenbyted.Text = ""
            End If
        End If
            
        'DArrin Let InternetSettings.TimerCheckWinsock.Enabled = True
        
        Call MainScreen.SendCommandviaTrackQ
    
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub
Public Sub Winsock_DataArrival(ByVal bytesTotal As Long)

    Dim IncomingData As String
    Dim Index As Integer
    Dim Button As Integer
    Dim Shift As Integer
    Dim X As Single
    Dim Y As Single
    Dim TemporaryIndex As Integer
    
    Winsock.GetData IncomingData

    Let textboxincomingdata.Text = IncomingData
    Let TextboxInBoundData.Text = TextboxInBoundData.Text + textboxincomingdata.Text + vbCrLf
    Let TextboxInBoundData.SelStart = Len(TextboxInBoundData.Text)
    If Left$(textboxincomingdata.Text, 28) = "Checking Winsock Connection." Then
        Let InternetSettings.TimerCheckWinsock.Enabled = False
        Let InternetSettings.TimerCheckWinsock.Interval = 65535
        Let InternetSettings.TimerCheckWinsock.Enabled = True
    ElseIf Left$(textboxincomingdata.Text, 10) = "Disconnect" Then
        Call InternetSettings.ButtonDisconnect_Click
    End If
    
    If Screen.ActiveForm.Name = "InternetSettings" Then
        If textboxincomingdata.Text = "MainlineOperationGuiDiesel1Screen" Then
            Load MainlineOperationGuiDiesel1Screen
            MainlineOperationGuiDiesel1Screen.Show vbModeless
        ElseIf textboxincomingdata.Text = "MainlineOperationGuiDiesel2Screen" Then
            Load MainlineOperationGuiDiesel2Screen
            MainlineOperationGuiDiesel2Screen.Show vbModeless
        ElseIf textboxincomingdata.Text = "MainlineOperationGuiDiesel3Screen" Then
            Load MainlineOperationGuiDiesel3Screen
            MainlineOperationGuiDiesel3Screen.Show vbModeless
        ElseIf textboxincomingdata.Text = "MainlineOperationGuiSteam1Screen" Then
            Load MainlineOperationGuiSteam1Screen
            MainlineOperationGuiSteam1Screen.Show vbModeless
        ElseIf textboxincomingdata.Text = "MainlineOperationGuiSteam2Screen" Then
            Load MainlineOperationGuiSteam2Screen
            MainlineOperationGuiSteam2Screen.Show vbModeless
        ElseIf textboxincomingdata.Text = "MainlineOperationGuiElectric1Screen" Then
            Load MainlineOperationGuiElectric1Screen
            MainlineOperationGuiElectric1Screen.Show vbModeless
        End If
        Let InternetSettings!textboxincomingdata.Text = ""
    End If
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Diesel 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Screen.ActiveForm.Name = "MainlineOperationGuiDiesel1Screen" Then
        If Left$(textboxincomingdata.Text, 4) = "Sand" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxSand_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Bell" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxBell_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Horn" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxHorn_Click
        ElseIf Left$(textboxincomingdata.Text, 5) = "Light" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxLight_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 21) = "Left Computer Screen " Then
            Let TemporaryIndex = (Mid$(textboxincomingdata.Text, 22, 1))
            Call MainlineOperationGuiDiesel1Screen.ButtonScreenLeft_Click(TemporaryIndex)
        ElseIf Left$(textboxincomingdata.Text, 22) = "Right Computer Screen " Then
            Let TemporaryIndex = (Mid$(textboxincomingdata.Text, 23, 1))
            Call MainlineOperationGuiDiesel1Screen.ButtonScreenRight_Click(TemporaryIndex)
        ElseIf Left$(textboxincomingdata.Text, 11) = "Reset Right" Then
            Call MainlineOperationGuiDiesel1Screen.TransPictureBoxResetRight_Click
        ElseIf Left$(textboxincomingdata.Text, 10) = "Reset Left" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxResetLeft_Click
        ElseIf Left$(textboxincomingdata.Text, 8) = "Reverser" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxReverser_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Throttle" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxThrottle_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 15) = "Automatic Brake" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxAutomaticBrake_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 17) = "Independent Brake" Then
            Call MainlineOperationGuiDiesel1Screen.TransPictureBoxIndependentBrake_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Deadmann" Then
            Call MainlineOperationGuiDiesel1Screen.sounddeadmann_Done(0)
        ElseIf Left$(textboxincomingdata.Text, 12) = "Start Engine" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 11) = "Stop Engine" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 10) = "RadioPhone" Then
            Call MainlineOperationGuiDiesel1Screen.PictureBoxRadioPhone_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 9) = "Fill Sand" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonFillSand_Click
        ElseIf Left$(textboxincomingdata.Text, 10) = "Fill Water" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonFillWater_Click
        ElseIf Left$(textboxincomingdata.Text, 8) = "Fill Oil" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonFillOil_Click
        ElseIf Left$(textboxincomingdata.Text, 9) = "Fill Fuel" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonFillFuel_Click
        ElseIf Left$(textboxincomingdata.Text, 12) = "Button Close" Then
            Call MainlineOperationGuiDiesel1Screen.ButtonCloseGUI_Click
        ElseIf Left$(textboxincomingdata.Text, 17) = "Auxilliary Switch" Then
            Let TemporaryIndex = CInt(Mid$(textboxincomingdata.Text, 18, 2))
            Call MainlineOperationGuiDiesel1Screen.TransPictureAuxillarySwitch_MouseDown(TemporaryIndex, Button, Shift, X, Y)
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Winsock, DataArrival, unknown data command for MainlineOperationGuiDiesel1Screen."
            End If
        End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Diesel 2
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    ElseIf Screen.ActiveForm.Name = "MainlineOperationGuiDiesel2Screen" Then
        If Left$(textboxincomingdata.Text, 4) = "Sand" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureSand_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Bell" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureBell_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Horn" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureHorn_Click
        ElseIf Left$(textboxincomingdata.Text, 9) = "Headlight" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureHeadlight_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Reverser" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureReverser_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Throttle" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureThrottle_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 15) = "Automatic Brake" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureBrakeAutomatic_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 17) = "Independent Brake" Then
            Call MainlineOperationGuiDiesel2Screen.TransPictureBrakeIndependent_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 12) = "Start Engine" Then
            Call MainlineOperationGuiDiesel2Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 11) = "Stop Engine" Then
            Call MainlineOperationGuiDiesel2Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 12) = "Button Close" Then
            Call MainlineOperationGuiDiesel2Screen.ButtonCloseGUI_Click
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Winsock, DataArrival, unknown data command for MainlineOperationGuiDiesel2Screen."
            End If
        End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Diesel 3
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    ElseIf Screen.ActiveForm.Name = "MainlineOperationGuiDiesel3Screen" Then
        If Left$(textboxincomingdata.Text, 4) = "Sand" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureSand_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Bell" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureBell_Click
        ElseIf Left$(textboxincomingdata.Text, 4) = "Horn" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureHorn_Click
        ElseIf Left$(textboxincomingdata.Text, 9) = "Headlight" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureHeadlight_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Reverser" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureReverser_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 8) = "Throttle" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureThrottle_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 15) = "Automatic Brake" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureBrakeAutomatic_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 17) = "Independent Brake" Then
            Call MainlineOperationGuiDiesel3Screen.TransPictureBrakeIndependent_MouseDown(Button, Shift, X, Y)
        ElseIf Left$(textboxincomingdata.Text, 12) = "Start Engine" Then
            Call MainlineOperationGuiDiesel3Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 11) = "Stop Engine" Then
            Call MainlineOperationGuiDiesel3Screen.ButtonEngineStart_Click
        ElseIf Left$(textboxincomingdata.Text, 12) = "Button Close" Then
            Call MainlineOperationGuiDiesel3Screen.ButtonCloseGUI_Click
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Winsock, DataArrival, unknown data command for MainlineOperationGuiDiesel3Screen."
            End If
        End If
    End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Routine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Number <> 0 Then
        If Number = 10060 Then
            Let LabelStatus.Caption = "Error " & Number & " " & Description
            Call ButtonDisconnect_Click
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An problem has occured with Automatic Train Control. This problem will be recorded in the ATC.LOG file." & vbCrLf & "Internet Settings Screen, Winsock Error, error " & Number & " " & Description & "." & vbCrLf & vbCrLf & "This problem is caused by the following situations:" & vbCrLf & "    The server you are trying to connect to is busy (connected to another person, please try again later) or" & vbCrLf & "    the server you are trying to connect to is down (not responding)." & vbCrLf & "Please check http://atc.lovethosetrains.com/model_train_program_screen_capture_live.html for a video" & vbCrLf & "streaming connection."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Winsock Error, error " & Number & " " & Description
            End If
        Else
            Let LabelStatus.Caption = "Error " & Number & " " & Description
            Call ButtonDisconnect_Click
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file." & vbCrLf & "Internet Settings Screen, Winsock Error, error " & Number & " " & Description & "."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, Winsock Error, error " & Number & " " & Description
            End If
        End If
    End If

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

        Let BalloonHelpText1 = "If there are any changes in the status of the communication port or" & vbCrLf & "data passing through the port, it will be reflected here."
        Let BalloonHelpText2 = "Status"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LabelStatus)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LabelStatus, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This option when 'click'ed on will operate" & vbCrLf & " the software as a host (server mode)."
        Let BalloonHelpText2 = "Option Host"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(optionhost)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(optionHost, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'OptionHost' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This option when 'click'ed on will operate" & vbCrLf & " the software as an end user (client mode)."
        Let BalloonHelpText2 = "Option Client"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(OptionClient)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(optionclient, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'OptionClient' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This option when 'click'ed on will operate" & vbCrLf & " the software in stand-alone mode."
        Let BalloonHelpText2 = "Option Stand Alone"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(OptionStandAlone)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(OptionStandAlone, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'OptionStandAlone' object."
            End If
        End If

        Let BalloonHelpText1 = "This option when 'click'ed on will allow you" & vbCrLf & "to select the server (train layout) to ." & vbCrLf & "connect to."
        Let BalloonHelpText2 = "Server Name (Train Layout)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboServerName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboServerName, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'ComboServerName' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This textbox displays the IP address" & vbCrLf & "of the hosting computer (server) with a train" & vbCrLf & "layout."
        Let BalloonHelpText2 = "Host IP Address"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(iptextboxhost)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(IpTextBoxHost, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'IpTextBoxHost' object."
            End If
        End If
         
        Let BalloonHelpText1 = "This textbox displays the IP address" & vbCrLf & "of the connecting computer (client)to server."
        Let BalloonHelpText2 = "Clients IP Address"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(IpTextBoxClient)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(IpTextBoxClient, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'IpTextBoxClient' object."
            End If
        End If
         
        Let BalloonHelpText1 = "This checkbox displays if the the internet" & vbCrLf & "is available via the LAN."
        Let BalloonHelpText2 = "Internet COnnnected Via LAN"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckboxNetConnectionViaLan)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxNetConnectionViaLan, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'CheckboxNetConnectionViaLan' object."
            End If
        End If
          
        Let BalloonHelpText1 = "This checkbox displays if the internet" & vbCrLf & "is available via the MODEM."
        Let BalloonHelpText2 = "Internet Connection Via Modem"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxNetConnectionViaModem)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxNetConnectionViaModem, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'CheckboxNetConnectionViaModem' object."
            End If
        End If

        Let BalloonHelpText1 = "This checkbox displays if the internet" & vbCrLf & "is available via the PROXY."
        Let BalloonHelpText2 = "Internet Connection Via Proxy"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxNetConnectionViaProxy)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxNetConnectionViaProxy, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'CheckBoxNetConnectionViaProxy' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
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
                Let Ini.Value = "Internet Settings Screen, BalloonHelpUpdatePart01, unable to setup balloon help for 'ButtonPrint' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This button closes the Scaled Time Window and returns" & vbCrLf & "you to the ATC main menu."
        Let BalloonHelpText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Internet SettingsScreen, BalloonHelpUpdatePart01, unable to setup balloon help for 'ButtonClose' object."
            End If
        End If

        Let InternetSettings.MousePointer = ccDefault

    End If
    
End Sub

Private Sub LabelSockUpdate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Winsock Label
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Winsock.State = sckClosed Then
        Let LabelStatus.Caption = "Status: Winsock is closed."
    ElseIf Winsock.State = sckOpen Then
        Let LabelStatus.Caption = "Status: Winsock is open."
    ElseIf Winsock.State = sckListening Then
        Let LabelStatus.Caption = "Status: Winsock is listening for a connection."
    ElseIf Winsock.State = sckConnectionPending Then
        Let LabelStatus.Caption = "Status: Winsock connection is pending."
    ElseIf Winsock.State = sckResolvingHost Then
        Let LabelStatus.Caption = "Status: Winsock is resolving host."
    ElseIf Winsock.State = sckConnecting Then
        Let LabelStatus.Caption = "Status :Winsock is connecting."
    ElseIf Winsock.State = sckConnected Then
        Let LabelStatus.Caption = "Status: Winsock is connected."
    ElseIf Winsock.State = sckClosing Then
        Let LabelStatus.Caption = "Status: Peer is closing winsock connection."
        'Call ButtonDisconnect_Click
    ElseIf Winsock.State = sckError Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Internet Settings Screen, LabelSockUpdate, state of winsock is indicating an error."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Send button
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Winsock.State = sckConnected Then
        Let ButtonSend.Enabled = True
        Let ButtonClose.Caption = "&Hide"
    Else
        Let ButtonSend.Enabled = False
        Let ButtonClose.Caption = "&Close"
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

        Let BalloonHelpText1 = "When in host mode only, this button" & vbCrLf & "places the software in a 'listening' state for an" & vbCrLf & "internet conenction."
        Let BalloonHelpText2 = "Listen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonListen)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonListen, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonListen' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in host mode only, this button" & vbCrLf & "places the software in an 'auto listening' state" & vbCrLf & "for an internet connection."
        Let BalloonHelpText2 = "Auto Listen"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(buttonAutoListen)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonAutoListen, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonAutoListen' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in client mode only, this button" & vbCrLf & "tries to place the software into a 'connect' state" & vbCrLf & "with the selected server (host train layout)."
        Let BalloonHelpText2 = "Connect"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonConnect)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(buttonconnect, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonConnect' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in client mode only, this button" & vbCrLf & "disconnects the software from the selected server" & vbCrLf & "(hosting a train layout)."
        Let BalloonHelpText2 = "Disconnect"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonDisconnect)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(buttondisconnect, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonDisconnect' object."
            End If
        End If

        Let BalloonHelpText1 = "When in host mode only, this button" & vbCrLf & "closes the 'Internet Settings' screen and opens the" & vbCrLf & "'Video Settings' window."
        Let BalloonHelpText2 = "Video Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonVideoSettings)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonVideoSettings, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonVideoSettings' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in client mode only, this checkbox" & vbCrLf & "selects 'server mode' for streaming video to the" & vbCrLf & "client."
        Let BalloonHelpText2 = "Client Stream Type"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckboxClientStreamTypeServer)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxclientstreamtypeserver, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'CheckboxClientStreamTypeServer' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in client mode only, this checkbox" & vbCrLf & "selects 'broadcast mode' for streaming video to the" & vbCrLf & "client (recommended)."
        Let BalloonHelpText2 = "Client Stream Type"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckboxClientStreamTypeBroadcast)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxclientstreamtypebroadcast, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'CheckboxClientStreamTypeBroadcast' object."
            End If
        End If
        
        Let BalloonHelpText1 = "When in host mode only, this button" & vbCrLf & "closes the 'Internet Settings' screen and opens the" & vbCrLf & "'Audio Settings' window."
        Let BalloonHelpText2 = "Audio Settings"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonAudioSettings)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonAudioSettings, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonAudioSettings' object."
            End If
        End If
        Let BalloonHelpText1 = "When in client mode only, this checkbox" & vbCrLf & "selects an additional audio stream (recommended)."
        Let BalloonHelpText2 = "Addition Audio Stream"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckboxAdditionalAudioStream)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckboxAdditionalAudioStream, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'CheckboxAdditionalAudioStream' object."
            End If
        End If
         
        Let BalloonHelpText1 = "When in host mode only, this button" & vbCrLf & "closes the 'Internet Settings' screen and opens the" & vbCrLf & "'Room Lighting' window."
        Let BalloonHelpText2 = "Room Lighting Control"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonRoomLightingControl)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonRoomLightingControl, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Main Screen Screen, BalloonHelpUpdatePart02, unable to setup balloon help for 'ButtonRoomLightingControl' object."
            End If
        End If
         
        Let InternetSettings.MousePointer = ccDefault

    End If
    
End Sub
