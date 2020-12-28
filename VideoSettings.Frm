VERSION 4.00
Begin VB.Form VideoSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Video Settings"
   ClientHeight    =   10560
   ClientLeft      =   14265
   ClientTop       =   1410
   ClientWidth     =   5775
   Height          =   10965
   Icon            =   "VideoSettings.frx":0000
   Left            =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   704
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Top             =   1065
   Width           =   5895
   Begin VB.CommandButton ButtonVideoTest 
      Caption         =   "Test &Video"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   4740
      Width           =   1275
   End
   Begin VB.ComboBox ComboBoxVideoInputIndex 
      Height          =   315
      ItemData        =   "VideoSettings.frx":0442
      Left            =   60
      List            =   "VideoSettings.frx":0452
      TabIndex        =   25
      Text            =   "ComboBoxVideoInputIndex"
      Top             =   2340
      Width           =   5655
   End
   Begin VB.ComboBox ComboBoxVideoCaptureFrameRate 
      Height          =   315
      ItemData        =   "VideoSettings.frx":04C4
      Left            =   60
      List            =   "VideoSettings.frx":04E3
      TabIndex        =   24
      Text            =   "ComboBoxVideoCaptureFrameRate"
      Top             =   3600
      Width           =   5655
   End
   Begin VB.ComboBox ComboBoxVideoCaptureFormats 
      Height          =   315
      ItemData        =   "VideoSettings.frx":05DA
      Left            =   60
      List            =   "VideoSettings.frx":05DC
      TabIndex        =   21
      Text            =   "ComboBoxVideoCaptureFormats"
      Top             =   3000
      Width           =   5655
   End
   Begin VB.ComboBox ComboBoxVideoCodec 
      Height          =   315
      ItemData        =   "VideoSettings.frx":05DE
      Left            =   60
      List            =   "VideoSettings.frx":05E0
      TabIndex        =   17
      Text            =   "ComboBoxVideoCodec"
      Top             =   4140
      Width           =   5655
   End
   Begin VB.CommandButton ButtonStartCasting 
      Caption         =   "&Start Casting"
      Enabled         =   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   10260
      Width           =   1215
   End
   Begin VB.CommandButton ButtonCaptureFilename 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   300
      Left            =   5400
      TabIndex        =   14
      Top             =   11880
      Width           =   300
   End
   Begin VB.CommandButton ButtonMoreWindowsMediaStreamingProfile 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   320
      Left            =   5460
      TabIndex        =   13
      Top             =   8880
      Width           =   320
   End
   Begin VB.TextBox TextBoxCaptureFilename 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   11880
      Width           =   5355
   End
   Begin VB.ComboBox ComboBoxWindowsMediaStreamingProfile 
      Height          =   315
      ItemData        =   "VideoSettings.frx":05E2
      Left            =   60
      List            =   "VideoSettings.frx":05E4
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   8820
      Width           =   5355
   End
   Begin VB.ComboBox ComboBoxVideoCaptureDevices 
      Height          =   315
      ItemData        =   "VideoSettings.frx":05E6
      Left            =   60
      List            =   "VideoSettings.frx":05E8
      TabIndex        =   3
      Text            =   "ComboBoxVideoCaptureDevices"
      Top             =   1740
      Width           =   5655
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   10260
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   60
      Picture         =   "VideoSettings.frx":05EA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4500
      TabIndex        =   0
      Top             =   10260
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Video Input"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   2100
      Width           =   810
   End
   Begin VB.Line Line4 
      X1              =   380
      X2              =   0
      Y1              =   820
      Y2              =   820
   End
   Begin VB.Line Line3 
      X1              =   4
      X2              =   380
      Y1              =   308
      Y2              =   308
   End
   Begin VB.Line Line2 
      X1              =   4
      X2              =   380
      Y1              =   676
      Y2              =   676
   End
   Begin VB.Line Line1 
      X1              =   4
      X2              =   384
      Y1              =   528
      Y2              =   528
   End
   Begin VB.Label Label4 
      Caption         =   "Video Capture Frame Rate"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Video Capture Format Size"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   1830
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   8040
      Top             =   2340
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VIDEOCAPXLib.VideoCapX VideoCapture 
      Height          =   2775
      Left            =   60
      TabIndex        =   20
      Top             =   5040
      Width           =   5655
      _Version        =   131072
      _ExtentX        =   9975
      _ExtentY        =   4895
      _StockProps     =   1
      CapFilename     =   ""
      CaptureRate     =   10
      VideoDeviceIndex=   -1
      AudioDeviceIndex=   -1
      VideoCodecQuality=   100
      ServerPort      =   20103
      ServerQuality   =   100
      SyncUsingStreamOffset=   -1  'True
      WMAttributes    =   $"VideoSettings.frx":0A2C
      StretchMode     =   1
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label LabelWindowsMediaStreamingProfileComment 
      Caption         =   "Currently selected windows media streaming profile is"
      Height          =   855
      Left            =   60
      TabIndex        =   18
      Top             =   9240
      Width           =   5655
   End
   Begin VB.Label LabelVideoCodec 
      Caption         =   "Video Codec"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   8040
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelVideoCapture 
      Caption         =   $"VideoSettings.frx":0AF7
      Height          =   855
      Left            =   660
      TabIndex        =   12
      Top             =   120
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   $"VideoSettings.frx":0BE6
      Height          =   435
      Left            =   60
      TabIndex        =   10
      Top             =   11400
      Width           =   5715
   End
   Begin VB.Label LabelCaptureFilename 
      Caption         =   "Capture Filename"
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
      TabIndex        =   9
      Top             =   11100
      Width           =   1635
   End
   Begin VB.Label LabelWindowsMediaStreamingProfileNote 
      Caption         =   $"VideoSettings.frx":0C76
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   8340
      Width           =   5715
   End
   Begin VB.Label LabelWindowsMediaStreamingProfile 
      Caption         =   "Windows Media Streaming Profile"
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
      TabIndex        =   6
      Top             =   8040
      Width           =   5115
   End
   Begin VB.Label LabelPreviewDevice 
      Caption         =   "Video Preview of Device"
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
      TabIndex        =   5
      Top             =   4740
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "Video Capture Device "
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
      TabIndex        =   4
      Top             =   1500
      Width           =   1395
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   8040
      Top             =   2880
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   8040
      Top             =   1740
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
End
Attribute VB_Name = "VideoSettings"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub buttoncapturefilename_Click()

    Let CommonDialog.Filter = "WMV files (*.wmv)|*.wmv|All fiels (*.*)|*.*"
    Let CommonDialog.DefaultExt = ".wmv"
    Let CommonDialog.Filename = ""
    CommonDialog.ShowSave
    If CommonDialog.Filename <> "" Then
        TextBoxCaptureFilename.Text = CommonDialog.Filename
    End If

End Sub

Private Sub ButtonClose_Click()
    
    VideoCapture.Overlay = False
    VideoCapture.Preview = False
    VideoCapture.Connected = False
    VideoCapture.VideoRenderer = False
    
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
            If TemporaryScreen = "Video Settings Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Setting Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
            End If
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' End Loop Premature
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TemporaryCounter = -2
        End If
    Next TemporaryCounter
    
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Screen Stack is Empty
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryCounter = -1 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings Screen, Button Close, stack empty, underflow."
        End If
    End If

End Sub







Private Sub ButtonMoreWindowsMediaStreamingProfile_Click()

    Let CommonDialog.Filter = "PRX files (*.prx)|*.prx|All fiels (*.*)|*.*"
    Let CommonDialog.DefaultExt = ".prx"
    Let CommonDialog.Filename = ""
    CommonDialog.ShowOpen
    If CommonDialog.Filename <> "" Then
        Dim TemporaryX As String
        Let TemporaryX = FileToString(CommonDialog.Filename)
        'MsgBox x
        VideoCapture.ProfileData = TemporaryX
    ComboBoxWindowsMediaStreamingProfile.ListIndex = -1
    ComboBoxWindowsMediaStreamingProfile.Text = VideoCapture.GetProfileName(-1)
    Call comboboxvideocapturedevices_Click
    
    End If
End Sub
























Private Sub ButtonStartCasting_Click()
    buttonstartcasting.Enabled = False
    If buttonstartcasting.Caption = "&Start Casting" Then
        Let buttonstartcasting.Caption = "&Stop Casting"
        'Let VideoCapture.ServerMode = True
        'DoEvents
        VideoCapture.StartBroadcast 20103, 2
        Let VideoCapture.Preview = True
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Enable and Disable Buttons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let comboboxvideocapturedevices.Enabled = False
        Let ComboBoxVideoCodec.Enabled = False
        Let ComboBoxWindowsMediaStreamingProfile.Enabled = False
        Let ButtonMoreWindowsMediaStreamingProfile.Enabled = False
        Let TextBoxCaptureFilename.Enabled = False
        Let buttoncapturefilename.Enabled = False
        Let ButtonClose.Enabled = False
    ElseIf buttonstartcasting.Caption = "&Stop Casting" Then
        Let buttonstartcasting.Caption = "&Start Casting"
        VideoCapture.StopBroadcast
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Enable and Disable Buttons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let comboboxvideocapturedevices.Enabled = True
        Let ComboBoxVideoCodec.Enabled = True
        Let ComboBoxWindowsMediaStreamingProfile.Enabled = True
        Let ButtonMoreWindowsMediaStreamingProfile.Enabled = True
        Let TextBoxCaptureFilename.Enabled = True
        Let buttoncapturefilename.Enabled = True
        Let ButtonClose.Enabled = True
    End If
    buttonstartcasting.Enabled = True
End Sub



Private Sub ButtonVideoTest_Click()

    If comboboxvideocapturedevices.ListIndex <> 0 Then
        If ComboBoxVideoInputIndex.ListIndex <> 0 Then
            If comboboxvideocaptureformats.ListIndex <> 0 Then
                If ComboBoxVideoCaptureFrameRate.ListIndex <> 0 Then
                    If ComboBoxVideoCodec.ListIndex <> 0 Then
                        If ComboBoxWindowsMediaStreamingProfile.ListIndex <> 0 Then
                        
                            Let ButtonVideoTest.Enabled = False
                            Let VideoSettings.MousePointer = vbHourglass
                            Let LabelStatus.Caption = "Status: Initializing Video"
                        
                            Let VideoCapture.VideoRenderer = False
                            Let VideoCapture.Connected = False
                            Let VideoCapture.Overlay = False
                            Let VideoCapture.Preview = False

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Video On
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            If VideoCapture.GetVideoDeviceCount <> 0 Then
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Video Device
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Filename = App.Path$ & "\Atc.ini"
                                Let Ini.Application = "Video Settings Screen"
                                Let Ini.Parameter = "VideoCaptureDeviceIndex"
                                Let VideoCapture.VideoDeviceIndex = Val(Ini.Value)
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Input Index
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "VideoInputIndex"
                                Let VideoCapture.VideoInputIndex = Val(Left$(Ini.Value, 1))
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Video Codec Index
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "VideoCodecIndex"
                                Let VideoCapture.VideoCodecIndex = Val(Ini.Value)
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Frame Rate (set before conecting)
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "VideoCaptureFrameRate"
                                Let VideoCapture.CaptureRate = Val(Left$(Ini.Value, 1))
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Streming Profiles
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "WindowsMediaStreamingProfileIndex"
                                Let VideoCapture.ProfileIndex = Val(Ini.Value)
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Audio Capture Device
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Application = "Audio Settings Screen"
                                Let Ini.Parameter = "AudioCaptureDeviceIndex"
                                If Val(Ini.Value) <> -1 Then
                                    Let VideoCapture.CaptureAudio = True
                                    Let VideoCapture.AudioDeviceIndex = Val(Ini.Value)
                                Else
                                    Let VideoCapture.CaptureAudio = False
                                End If
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Audio Codec Index
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "AudioCodecIndex"
                                Let VideoCapture.AudioCodecIndex = Val(Ini.Value)
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                ' Audio Input Index
                                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                Let Ini.Parameter = "AudioInputIndex"
                                Let VideoCapture.AudioInputIndex = Val(Ini.Value)
                                
                                If VideoCapture.VideoDeviceIndex > -1 Then
                                
                                    Let VideoCapture.VideoRenderer = True
                                            
                                    On Error Resume Next
                                    Let VideoCapture.Connected = True
                                    
                                    If Err.Number <> 0 Then
                                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                            Let Ini.Filename = App.Path$ & "/Atc.log"
                                            Let Ini.Application = "Log Errors"
                                            Let Ini.Parameter = Date$ & " " & Time$
                                            Let Ini.Value = "Video Settings Screen, ButtonVideoTest, no video capture device found, cannot connect."
                                        End If
                                    Else
                                        Let piu = VideoCapture.GetVMR9IUnknown
                                    
                                        Call GetVideoDeviceFormats
                                        
                                        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                        ' Set Maximum Resolution
                                        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                        
                                        If Left$(Ini.Value, 10) = "720 by 480" Then
                                            Let MaximumVideoCapture = VideoCapture.SetVideoFormat(720, 480)
                                            If MaximumVideoCapture = "False" Then
                                                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                    Let Ini.Filename = App.Path$ & "/Atc.log"
                                                    Let Ini.Application = "Log Errors"
                                                    Let Ini.Parameter = Date$ & " " & Time$
                                                    Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, cannot set video to 720 by 480 pixels. "
                                                End If
                                            End If
                                        ElseIf Left$(Ini.Value, 10) = "640 by 480" Then
                                            Let MaximumVideoCapture = VideoCapture.SetVideoFormat(640, 480)
                                            If MaximumVideoCapture = "False" Then
                                                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                    Let Ini.Filename = App.Path$ & "/Atc.log"
                                                    Let Ini.Application = "Log Errors"
                                                    Let Ini.Parameter = Date$ & " " & Time$
                                                    Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, cannot set video to 640 by 480 pixels."
                                                End If
                                            End If
                                        ElseIf Left$(Ini.Value, 10) = "320 by 240" Then
                                            Let MaximumVideoCapture = VideoCapture.SetVideoFormat(320, 240)
                                            If MaximumVideoCapture = "False" Then
                                                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                    Let Ini.Filename = App.Path$ & "/Atc.log"
                                                    Let Ini.Application = "Log Errors"
                                                    Let Ini.Parameter = Date$ & " " & Time$
                                                    Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, cannot set video to 320 by 240 pixels."
                                                End If
                                            End If
                                        ElseIf Left$(Ini.Value, 10) = "240 by 120" Then
                                            Let MaximumVideoCapture = VideoCapture.SetVideoFormat(240, 120)
                                            If MaximumVideoCapture = "False" Then
                                                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                    Let Ini.Filename = App.Path$ & "/Atc.log"
                                                    Let Ini.Application = "Log Errors"
                                                    Let Ini.Parameter = Date$ & " " & Time$
                                                    Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, cannot set video to 240 by 120 pixels."
                                                End If
                                            End If
                                        Else
                                            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                Let Ini.Filename = App.Path$ & "/Atc.log"
                                                Let Ini.Application = "Log Errors"
                                                Let Ini.Parameter = Date$ & " " & Time$
                                                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, cannot find the maximim video resolution for video source. "
                                            End If
                                        End If
                                    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn on
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                    If VideoCapture.HasOverlay = True Then
                                        Let VideoCapture.Preview = False
                                        Let VideoCapture.Overlay = True
                                    Else
                                       Let VideoCapture.Overlay = False
                                        Let VideoCapture.Preview = True
                                    End If
                            
                                    On Error GoTo 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Video Filter Settings
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                    Let Ini.Filename = App.Path$ & "/ATC.ini"
                                    Let Ini.Application = "Video Settings Screen"
                                    Let Ini.Parameter = "UseVideoFilter"
                                    If Ini.Value = "True" Then
                                        Let VideoCapture.UseVideoFilter = True
                                        Let Ini.Parameter = "SetTextOverlayIndex"
                                        'Dim TemporarySetTextOverlayIndex As Long
                                        Let TemporarySetTextOverlayIndex = CLng(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayCaption"
                                        'Dim TemporarySetTextOverlayCaption As String
                                        Let TemporarySetTextOverlayCaption = CStr(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayX"
                                        'Dim TemporarySetTextOverlayX As Long
                                        Let TemporarySetTextOverlayX = CLng(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayY"
                                        'Dim TemporarySetTextOverlayY As Long
                                        Let TemporarySetTextOverlayY = CLng(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayFontName"
                                        'Dim TemporarySetTextOverlayFontName As String
                                        Let TemporarySetTextOverlayFontName = Ini.Value
                                        Let Ini.Parameter = "SetTextOverlayFontSize"
                                        'Dim TemporarySetTextOverlayFontSize As Long
                                        Let TemporarySetTextOverlayFontSize = CLng(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayFColor"
                                        'Dim TemporarySetTextOverlayFColor As Long
                                        Let TemporarySetTextOverlayFColor = CLng(Ini.Value)
                                        Let Ini.Parameter = "SetTextOverlayBColor"
                                        'Dim TemporarySetTextOverlayBColor As Long
                                        Let TemporarySetTextOverlayBColor = CLng(Ini.Value)
                                        VideoCapture.SetTextOverlay TemporarySetTextOverlayIndex, "Video Streaming Test", TemporarySetTextOverlayX, TemporarySetTextOverlayY, TemporarySetTextOverlayFontName, TemporarySetTextOverlayFontSize, TemporarySetTextOverlayFColor, TemporarySetTextOverlayBColor
                                   Else
                                        Let VideoCapture.UseVideoFilter = False
                                        If Ini.Value <> "False" Then
                                            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                                Let Ini.Filename = App.Path$ & "/Atc.log"
                                                Let Ini.Application = "Log Errors"
                                                Let Ini.Parameter = Date$ & " " & Time$
                                                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, invalied value for 'UseVideoFilter' setting."
                                            End If
                                        End If
                                    End If
                            
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Video Capture Proc Amp
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                                    Call VideoProcAmp
                                End If
                            End If
                            Let LabelStatus.Caption = "Status: Finished Initializing Video"
                            VideoSettings.MousePointer = vbDefault
                            ButtonVideoTest.Enabled = True

                        Else
                            Let TemporaryMessage = "No video streaming profile was selected. Before testing the video stream, select a video streaming profile."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        End If
                    Else
                        Let TemporaryMessage = "No video codec was selected. Before testing the video stream, select a video codec."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    End If
                Else
                    Let TemporaryMessage = "No video frame rate was selected. Before testing the video stream, select a video frame rate."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                End If
            Else
                Let TemporaryMessage = "No video capture format was selected. Before testing the video stream, select a video capture format."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            End If
        Else
            Let TemporaryMessage = "No video input was selected. Before testing the video stream, select a video input."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        End If
    Else
        Let TemporaryMessage = "No video capture device was selected. Before testing the video stream, select a video capture device."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
    End If
    
End Sub

Private Sub comboboxvideocapturedevices_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCaptureDevice"
    Let Ini.Value = comboboxvideocapturedevices.Text
    Let Ini.Parameter = "VideoCaptureDeviceIndex"
    Let Ini.Value = comboboxvideocapturedevices.ListIndex - 1
    
    If comboboxvideocapturedevices.ListIndex = 0 Then
        Let ComboBoxVideoInputIndex.Enabled = False
        Let comboboxvideocaptureformats.Enabled = False
        Let ComboBoxVideoCaptureFrameRate.Enabled = False
        Let ComboBoxVideoCodec.Enabled = False
        Let ComboBoxWindowsMediaStreamingProfile.Enabled = False
    Else
        Let ComboBoxVideoInputIndex.Enabled = True
        Let comboboxvideocaptureformats.Enabled = True
        Let ComboBoxVideoCaptureFrameRate.Enabled = True
        Let ComboBoxVideoCodec.Enabled = True
        Let ComboBoxWindowsMediaStreamingProfile.Enabled = True
    End If

    
End Sub


Private Sub comboboxvideocaptureformats_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCaptureFormat"
    Let Ini.Value = comboboxvideocaptureformats.Text

End Sub


Private Sub ComboBoxVideoCaptureFrameRate_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCaptureFrameRate"
    Let Ini.Value = ComboBoxVideoCaptureFrameRate.Text

End Sub


Private Sub ComboBoxVideoCodec_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCodecName"
    Let Ini.Value = ComboBoxVideoCodec.Text
    Let Ini.Parameter = "VideoCodecIndex"
    Let Ini.Value = ComboBoxVideoCodec.ListIndex - 1

End Sub


Private Sub ComboBoxVideoInputIndex_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoInputIndex"
    Let Ini.Value = ComboBoxVideoInputIndex.Text
    
End Sub

Private Sub ComboBoxWindowsMediaStreamingProfile_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "WindowsMediaStreamingProfile"
    Let Ini.Value = ComboBoxWindowsMediaStreamingProfile.Text
    Let Ini.Parameter = "WindowsMediaStreamingProfileIndex"
    Let Ini.Value = ComboBoxWindowsMediaStreamingProfile.ListIndex - 1

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
        If TemporaryScreen = "Video Settings Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Video Settings Screen"
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
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. "
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Balloon Help Update
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call BalloonHelpUpdatePart01
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn on the Video (First Time)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call GetVideoDevices
    Call GetVideoInput
    Call GetVideoDeviceFormats
    Call GetVideoFrameRate
    Call GetVideoCodec
    Call GetStreamingProfile
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub

Private Sub Form_Deactivate()
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning Video Settings, Form, Deactivate" & vbCrLf
    End If 'Darrin

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(VideoSettings.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(VideoSettings.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(VideoSettings.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(VideoSettings.Height)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Video Capture Device Selection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Let Ini.Parameter = "VideoCaptureDevice"
'    Let Ini.Value = comboboxvideocapturedevices.Text
'    Let Ini.Parameter = "VideoCaptureDeviceIndex"
'    Let Ini.Value = VideoCapture.VideoDeviceIndex
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Video Capture Device Selection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Let Ini.Parameter = "VideoInputindex"
'    Let Ini.Value = ComboBoxVideoInputIndex.Text
'    Let Ini.Parameter = "VideoCaptureFormat"
'    Let Ini.Value = comboboxvideocaptureformats.Text
'    Let Ini.Parameter = "VideoCaptureFrameRate"
'    Let Ini.Value = ComboBoxVideoCaptureFrameRate.Text
'    Let Ini.Parameter = "WindowsMediaStreamingProfile"
'    Let Ini.Value = ComboBoxWindowsMediaStreamingProfile.Text
'    Let Ini.Parameter = "WindowsMediaStreamingProfileIndex"
'    Let Ini.Value = ComboBoxWindowsMediaStreamingProfile.ListIndex
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Video Codec
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Let Ini.Parameter = "VideoCodecName"
'    Let Ini.Value = ComboBoxVideoCodec.Text
'    Let Ini.Parameter = "VideoCodecIndex"
'    Let Ini.Value = ComboBoxVideoCodec.ListIndex
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Video Capture Device Selection
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    'Let Ini.Filename = App.Path$ & "\Atc.ini"
'    'Let Ini.Application = "Video Settings Screen"
'    Let Ini.Parameter = "CaptureFilename"
'    Let Ini.Value = TextBoxCaptureFilename.Text
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
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    VideoSettings.Hide

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending Video Settings, Form, Deativate" & vbCrLf
    End If
  
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
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        VideoSettings.Left = (Screen.Width - Width) / 2
        VideoSettings.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + VideoSettings.Width > Screen.Width Then
            Let VideoSettings.Left = Screen.Width - VideoSettings.Width
        Else
            Let VideoSettings.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + VideoSettings.Height > Screen.Height Then
            Let VideoSettings.Top = Screen.Height - VideoSettings.Height
        Else
            Let VideoSettings.Top = Val(TemporaryValueTop)
        End If
    End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'No databases to declare
    
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



Private Function FileToString(Filename)

    Dim TemporaryX
    Open Filename For Input As #1
    While Not EOF(1)
        Let TemporaryX = TemporaryX + Input(1, #1)
    Wend
    Close #1
    Let FileToString = TemporaryX

End Function

Private Sub Form_Resize()

    If VideoSettings.WindowState = vbMinimized Then
    
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
        
    ElseIf VideoSettings.WindowState = vbNormal Then
    
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




Private Sub TextBoxCaptureFilename_Click()

    VideoCapture.CapFilename = TextBoxCaptureFilename.Text

End Sub


Private Sub Timer1_Timer()

    Let Timer1.Interval = 0
    Let Timer1.Enabled = False
    
    Call ButtonVideoTest_Click
    
End Sub

Private Sub VideoCapture_CaptureEnd()

    Let LabelStatus.Caption = "Status: Video capture has ended."

End Sub

Private Sub VideoCapture_CaptureReady()
    
    Let LabelStatus.Caption = "Status: Video capture is ready."

End Sub


Private Sub VideoCapture_CaptureStart()
    
    Let LabelStatus.Caption = "Status: Video capture has started."

End Sub


Private Sub VideoCapture_ConnectionClosed(ByVal RemoteAddress As String, ByVal ErrorCode As Long)
    
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let Ini.Filename = App.Path$ & "/Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Video streaming connection closed from " & FromAddress & " with error code of " & ErrorCode
    End If
    
    Let LabelStatus.Caption = "Status: Video streaming connection closed from " & FromAddress

End Sub

Private Sub VideoCapture_ConnectionRequest(ByVal FromAddress As String, Cancel As Long)
   
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let Ini.Filename = App.Path$ & "/Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Video streaming connection requested from " & FromAddress
    End If
    Let LabelStatus.Caption = "Status: Video streaming connection requested from " & FromAddress

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
        Else 'If mainscreen!MenuTransparency.Caption = "&Transparency is On" Then
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
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'LabelStatus' object."
            End If
        End If

        Let BalloonHelpText1 = "Select a video capture device from the" & vbCrLf & "drop-down combination box to stream the video into" & vbCrLf & "the computer."
        Let BalloonHelpText2 = "Video Capture Devices"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboboxVideoCaptureDevices)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(comboboxvideocapturedevices, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ComboboxVideoCaptureDevices' object."
            End If
        End If

        Let BalloonHelpText1 = "Select a video format to capture  video" & vbCrLf & "from the device selected above."
        Let BalloonHelpText2 = "Video Capture Formats"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboboxVideoCaptureFormats)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(comboboxvideocaptureformats, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ComboboxVideoCaptureFormats' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Select a video frame rate to capture  video" & vbCrLf & "from the device selected above."
        Let BalloonHelpText2 = "Video Capture Frame Rate"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboboxVideoCaptureFrameRate)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxVideoCaptureFrameRate, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ComboboxVideoCaptureFrameRate' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Select a video codec to stream  video" & vbCrLf & "from the device selected above."
        Let BalloonHelpText2 = "Video Capture Formats"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBoxVideoCodec)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxVideoCodec, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ComboBoxVideoCodec' object."
            End If
        End If
        
        Let BalloonHelpText1 = "After selecting a video capture device" & vbCrLf & "a preview of the video stream can be seen here."
        Let BalloonHelpText2 = "Video Capture Preview"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(VideoCapture)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(VideoCapture, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'VideoCapture' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Setting the profile for streaming the" & vbCrLf & "video from the host to the clients is set here."
        Let BalloonHelpText2 = "Video Streaming Profile"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBoxWindowsMediaStreamingProfile)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBoxWindowsMediaStreamingProfile, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ComboBoxWindowsMediaStreamingProfile' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Do you have your own video streaming" & vbCrLf & "profile file? You can manual load it here."
        Let BalloonHelpText2 = "Additional Streaming Profiles"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonMoreWindowsMediaStreamingProfile)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonMoreWindowsMediaStreamingProfile, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ButtonMoreWindowsMediaStreamingProfile' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Enter a valid filename to save the" & vbCrLf & "video stream."
        Let BalloonHelpText2 = "Capture Filename"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCaptureFilename)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCaptureFilename, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'TextBoxCaptureFilename' object."
            End If
        End If
        
        Let BalloonHelpText1 = "Enter a valid filename to save the" & vbCrLf & "video stream."
        Let BalloonHelpText2 = "Additional Filenames"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(buttoncapturefilename)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(buttoncapturefilename, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ButtonCaptureFilename' object."
            End If
        End If

        Let BalloonHelpText1 = "To test out the streaming video, turn" & vbCrLf & "the video broadcasting on."
        Let BalloonHelpText2 = "Broadcast"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonStartCasting)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(buttonstartcasting, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ButtonStartCasting' object."
            End If
        End If

        Let BalloonHelpText1 = "Click on this button to print the" & vbCrLf & "Video Settings sceeen."
        Let BalloonHelpText2 = "Print"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ButtonPrint' object."
            End If
        End If

        Let BalloonHelpText1 = "Close this screen and return the the" & vbCrLf & "previous screen."
        Let BalloonHelpText2 = "Close"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings Screen, Balloon Help Update Part 01, unable to setup balloon help for 'ButtonClose' object."
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
                Let Ini.Value = "Video Settings Screen, Balloon Help Update, unable to destroy all balloon help."
            End If
        End If
    End If

End Sub

Private Sub VideoProcAmp()

    Let LabelStatus.Caption = "Status: Initializing Video Proc Amp"

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
                
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Video Capture Proc Amp
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Parameter = "VideoCaptureBrightness"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(0)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureBrightness' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureContrast"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(1)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureContrast' setting."
            End If
        End If
    End If
        
    Let Ini.Parameter = "VideoCaptureHue"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(2)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureBrightness' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureSaturation"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(3)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureSaturation' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureSharpness"
    If Val(Ini.Value) = -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(4)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureSharpness' setting."
            End If
        End If
    End If
        
    Let Ini.Parameter = "VideoCaptureGamma"
    If Val(Ini.Value) = -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(5)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureGamma' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureColorEnable"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(6)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureColorEnable' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureWhiteBalance"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(7)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureWhiteBalance' setting."
            End If
        End If
    End If

    Let Ini.Parameter = "VideoCaptureBacklightCompensation"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(8)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureBackLightCompensation' setting."
            End If
        End If
    End If
    
    Let Ini.Parameter = "VideoCaptureGain"
    If Val(Ini.Value) <> -1 Then
        On Error Resume Next
        Let TemporaryValue = VideoCapture.VideoProcAmp(9)
        If Err <> 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, " & Error
            End If
        End If
        On Error GoTo 0
        If TemporaryValue = -1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, Form Activate, variable error in ATC.INI file for 'VideoCaptureGain' setting."
            End If
        End If
    End If

End Sub

Private Sub GetVideoDeviceFormats()

    Let LabelStatus.Caption = "Status: Getting Video Formats"

    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Get Video Formats
    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Connect to the video cature device, prior get the information
    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let VideoCapture.Connected = True
    
    
    Dim TemporaryFunction, TemporaryFormats
    Let TemporaryFunction = VideoCapture.GetVideoCaps()

    comboboxvideocaptureformats.Clear
    comboboxvideocaptureformats.AddItem "After selecting a video capture device, select a video format."
    
    For TemporaryFormats = LBound(TemporaryFunction) To UBound(TemporaryFunction)
        Let TemporaryString = TemporaryFunction(TemporaryFormats, 0) & " by " & TemporaryFunction(TemporaryFormats, 1) & " pixels with " & TemporaryFunction(TemporaryFormats, 2) & " bits per pixel value and with "
        Select Case TemporaryFunction(TemporaryFormats, 3)
        Case 0
            Let TemporaryString = TemporaryString & "RGB24 Colour Format"
        Case 1
            Let TemporaryString = TemporaryString & "CLPL Colour Format"
        Case 2
            Let TemporaryString = TemporaryString & "YUYV Colour Format"
        Case 3
            Let TemporaryString = TemporaryString & "IYUV Colour Format"
        Case 4
            Let TemporaryString = TemporaryString & "YUV9 Colour Format"
        Case 5 '
            Let TemporaryString = TemporaryString & "Y411 Colour Format"
        Case 6
            Let TemporaryString = TemporaryString & "Y41P Colour Format"
        Case 7
            Let TemporaryString = TemporaryString & "YUY2 Colour Format"
        Case 8
            Let TemporaryString = TemporaryString & "YVYU Colour Format"
        Case 9
            Let TemporaryString = TemporaryString & "UYVY Colour Format"
        Case 10
            Let TemporaryString = TemporaryString & "Y211 Colour Format"
        Case 11
            Let TemporaryString = TemporaryString & "YV12 Colour Format"
        Case 12
            Let TemporaryString = TemporaryString & "CLJR Colour Format"
        Case 13
            Let TemporaryString = TemporaryString & "IF09 Colour Format"
        Case 14
            Let TemporaryString = TemporaryString & "CPLA Colour Format"
        Case 15
            Let TemporaryString = TemporaryString & "MJPG Colour Format"
        Case 16
            Let TemporaryString = TemporaryString & "TVMJ Colour Format"
        Case 17
            Let TemporaryString = TemporaryString & "WAKE Colour Format"
        Case 18
            Let TemporaryString = TemporaryString & "CFCC Colour Format"
        Case 19
            Let TemporaryString = TemporaryString & "IJPG Colour Format"
        Case 20
            Let TemporaryString = TemporaryString & "PLUM Colour Format"
        Case 21
            Let TemporaryString = TemporaryString & "DVCS Colour Format"
        Case 22
            Let TemporaryString = TemporaryString & "MDVF Colour Format"
        Case 23
            Let TemporaryString = TemporaryString & "RGB1 Colour Format"
        Case 24
            Let TemporaryString = TemporaryString & "RGB4 Colour Format"
        Case 25
            Let TemporaryString = TemporaryString & "RGB8 Colour Format"
        Case 26
            Let TemporaryString = TemporaryString & "RGB565 Colour Format"
        Case 28
            Let TemporaryString = TemporaryString & "RGB555 Colour Format"
        Case 29
            Let TemporaryString = TemporaryString & "RGB32 Colour Format"
        Case Else
            Let TemporaryString = TemporaryString & "No Colour Format"
        End Select
        comboboxvideocaptureformats.AddItem TemporaryString
    Next TemporaryFormats

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCaptureFormat"
    Let comboboxvideocaptureformats.Text = Ini.Value

End Sub


Private Sub GetStreamingProfile()

    Let LabelStatus.Caption = "Status: Getting Streaming Profile"

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    ComboBoxWindowsMediaStreamingProfile.Clear
    ComboBoxWindowsMediaStreamingProfile.AddItem "After selecting a video device, select a media streaming profile."
    Let TemporaryPointer = 0
    Let Ini.Parameter = "WindowsMediaStreamingProfile"
    For TemporaryF = 0 To VideoCapture.GetProfileCount - 1
        ComboBoxWindowsMediaStreamingProfile.AddItem VideoCapture.GetProfileName(TemporaryF) ' + " " + videocapture.GetProfileDesc(temporaryf)
        If ComboBoxWindowsMediaStreamingProfile.List(TemporaryF + 1) = Ini.Value Then
            Let TemporaryPointer = TemporaryF + 1
        End If
    Next TemporaryF
    Let ComboBoxWindowsMediaStreamingProfile.ListIndex = TemporaryPointer

    Let Ini.Parameter = "WindowsMediaStreamingProfileIndex"
    If Val(Ini.Value) = ComboBoxWindowsMediaStreamingProfile.ListIndex - 1 Then
        Let VideoCapture.ProfileIndex = Val(Ini.Value)
    Else
        Let VideoCapture.ProfileIndex = -1
        Let ComboBoxWindowsMediaStreamingProfile.ListIndex = 0
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings, GetStreamingProfile, streaming profile is missmatched to streaming profile name."
        End If
    End If

    Let labelwindowsmediastreamingprofilecomment.Caption = VideoCapture.GetProfileDesc(ComboBoxWindowsMediaStreamingProfile.ListIndex - 1)

End Sub

Private Sub GetVideoCodec()
    
    Let LabelStatus.Caption = "Status: Getting Video Codecs"
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Get Video Codec
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ComboBoxVideoCodec.Clear
    ComboBoxVideoCodec.AddItem "After selecting a video capture device, please select a codec."
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCodecName"
    Let TemporaryPointer = 0
    For TemporaryF = 0 To VideoCapture.GetVideoCodecCount - 1
        ComboBoxVideoCodec.AddItem VideoCapture.GetVideoCodecName(TemporaryF)
        If ComboBoxVideoCodec.List(TemporaryF + 1) = Ini.Value Then
            Let TemporaryPointer = TemporaryF + 1
        End If
    Next TemporaryF
    Let ComboBoxVideoCodec.ListIndex = TemporaryPointer
    Let Ini.Parameter = "VideoCodecIndex"
    If Val(Ini.Value) = ComboBoxVideoCodec.ListIndex - 1 Then
        Let VideoCapture.VideoCodecIndex = Val(Ini.Value)
     Else
        Let VideoCapture.VideoCodecIndex = -1
        Let ComboBoxVideoCodec.ListIndex = 0
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Video Settings, GetVideoCodec, video index is missmatched to video codec name in INI file."
        End If
    End If

End Sub
Private Sub GetVideoDevices()

    Let LabelStatus.Caption = "Status: Getting Video Formats"

    If VideoCapture.GetVideoDeviceCount <> 0 Then
        
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Video Settings Screen"
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Poll for Video Devices
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        comboboxvideocapturedevices.Clear
        comboboxvideocapturedevices.AddItem "Please select a video capture device."
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let TemporaryPointer = 0
        Let Ini.Parameter = "VideoCaptureDevice"
        For TemporaryF = 0 To VideoCapture.GetVideoDeviceCount - 1
            comboboxvideocapturedevices.AddItem VideoCapture.GetVideoDeviceName(TemporaryF)
            If comboboxvideocapturedevices.List(TemporaryF + 1) = Ini.Value Then
                Let TemporaryPointer = TemporaryF + 1
            End If
        Next TemporaryF

        Let comboboxvideocapturedevices.ListIndex = TemporaryPointer
        Let Ini.Parameter = "VideoCaptureDeviceIndex"
        If Val(Ini.Value) = comboboxvideocapturedevices.ListIndex - 1 Then
            Let VideoCapture.VideoDeviceIndex = Val(Ini.Value)
        Else
            Let VideoCapture.VideoDeviceIndex = -1
            Let comboboxvideocapturedevices.ListIndex = 0
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Video Settings, GetVideoDevices, video device index is missmatched to video device name in INI file."
            End If
        End If
    End If
    
End Sub




Private Sub GetVideoFrameRate()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoCaptureFrameRate"
    Let ComboBoxVideoCaptureFrameRate.Text = Ini.Value
    
End Sub

Private Sub GetVideoInput()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Video Settings Screen"
    Let Ini.Parameter = "VideoInputIndex"
    Let ComboBoxVideoInputIndex.Text = Ini.Value

End Sub
