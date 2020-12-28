VERSION 4.00
Begin VB.Form AudioSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Audio Settings"
   ClientHeight    =   8190
   ClientLeft      =   4215
   ClientTop       =   2235
   ClientWidth     =   6225
   Height          =   8595
   Icon            =   "AudioSettings.frx":0000
   Left            =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   Top             =   1890
   Width           =   6345
   Begin VB.CommandButton ButtonOutPutFileName 
      Caption         =   "..."
      Height          =   315
      Left            =   5820
      TabIndex        =   27
      Top             =   7320
      Width           =   315
   End
   Begin VB.CommandButton ButtonAudioCodec 
      Caption         =   "..."
      Height          =   315
      Left            =   5820
      TabIndex        =   26
      Top             =   6000
      Width           =   315
   End
   Begin VB.ComboBox ComboBoxAudioCodec 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   6000
      Width           =   5655
   End
   Begin VB.TextBox TextBoxOutputFileName 
      Height          =   315
      Left            =   60
      TabIndex        =   23
      Top             =   7320
      Width           =   5655
   End
   Begin VB.ComboBox ComboBoxSamplesPerSecond 
      Height          =   315
      ItemData        =   "AudioSettings.frx":0442
      Left            =   60
      List            =   "AudioSettings.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5220
      Width           =   6075
   End
   Begin VB.Frame FrameAudioBitRate 
      Caption         =   "Audio Bit Rate"
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   4020
      Width           =   6075
      Begin VB.OptionButton OptionAudioBitRate32 
         Caption         =   "Thirty-Two (32) Bits"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1995
      End
      Begin VB.OptionButton OptionAudioBitRate16 
         Caption         =   "Sixteen (16) Bits"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptionAudioBitRate8 
         Caption         =   "Eight (8) Bits"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame FrameAudioChannelNumbers 
      Caption         =   "Audio Channel Numbers"
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   3120
      Width           =   6075
      Begin VB.OptionButton OptionAudioChannelNumbers5 
         Caption         =   "Option5.1 (Cinema)"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   1875
      End
      Begin VB.OptionButton OptionAudioChannelNumbers2 
         Caption         =   "2.1 (Stereo)"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton OptionAudioChannelNumbers1 
         Caption         =   "1.1 (Mono)"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton ButtonAudioInput 
      Caption         =   "..."
      Height          =   315
      Left            =   5820
      TabIndex        =   11
      Top             =   2520
      Width           =   315
   End
   Begin VB.CommandButton ButtonAudioSource 
      Caption         =   "..."
      Height          =   315
      Left            =   5820
      TabIndex        =   10
      Top             =   1680
      Width           =   315
   End
   Begin VB.ComboBox ComboBoxAudioInput 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2520
      Width           =   5655
   End
   Begin VB.ComboBox ComboBoxAudioDevice 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   3660
      TabIndex        =   2
      Top             =   7860
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "AudioSettings.frx":04B5
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
      Left            =   4920
      TabIndex        =   0
      Top             =   7860
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   6480
      Top             =   4260
      _ExtentX        =   873
      _ExtentY        =   556
   End
   Begin VB.Label Label1 
      Caption         =   "Select an output file name if you would like the audio captured, otherwise leave the textbox blank."
      Height          =   435
      Left            =   60
      TabIndex        =   28
      Top             =   6840
      Width           =   6075
   End
   Begin VB.Label LabelAudioCodec 
      Caption         =   "Audio Codec"
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   5760
      Width           =   6015
   End
   Begin VB.Label LabelOutputFileName 
      Caption         =   "Output File Name"
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   6540
      Width           =   6075
   End
   Begin VB.Label LabelSamplesPerSecond 
      Caption         =   "Samples Per Second"
      Height          =   195
      Left            =   60
      TabIndex        =   21
      Top             =   4920
      Width           =   5955
   End
   Begin VB.Label LabelAudioInput 
      Caption         =   "Audio Input"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2220
      Width           =   5655
   End
   Begin VB.Label LabelAudioDevice 
      Caption         =   "Audio Device"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1440
      Width           =   5715
   End
   Begin VIDEOCAPXLib.VideoCapX VideoCapture 
      Height          =   315
      Left            =   5820
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
      _Version        =   131072
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   0
      CapFilename     =   ""
      CaptureRate     =   10
      PreviewScale    =   -1  'True
      VideoDeviceIndex=   -1
      ServerPort      =   10103
      ServerQuality   =   100
      SyncUsingStreamOffset=   -1  'True
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6480
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelAudoCapture 
      Caption         =   $"AudioSettings.frx":08F7
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   6480
      Top             =   4680
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   6480
      Top             =   3720
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
Attribute VB_Name = "AudioSettings"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub ButtonClose_Click()
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Audio Channel Number
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let Ini.Filename = App.Path$ & "\Atc.ini"
    'Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioChannelNumbers"
    If OptionAudioChannelNumbers1.Value = True Then
        Let Ini.Value = "1.1 (Mono)"
    ElseIf OptionAudioChannelNumbers2.Value = True Then
        Let Ini.Value = "2.1 (Stereo)"
    ElseIf OptionAudioChannelNumbers5.Value = True Then
        Let Ini.Value = "5.1 (Cinema)"
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Audio Bit Rate
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let Ini.Filename = App.Path$ & "\Atc.ini"
    'Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioBitRate"
    If OptionAudioBitRate8.Value = True Then
        Let Ini.Value = "Eight (8) Bit"
    ElseIf OptionAudioBitRate16.Value = True Then
        Let Ini.Value = "Sixteen (16) Bit"
    ElseIf OptionAudioBitRate32.Value = True Then
        Let Ini.Value = "Thirty-Two (32) Bit"
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save Audio Output Filename
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let Ini.Filename = App.Path$ & "\Atc.ini"
    'Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "OutputFileName"
    Let Ini.Value = TextBoxOutputFileName.Text
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check for Audio Settings
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Dim nFormatTag As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Waveform-audio format type. Format tags are registered with Microsoft Corporation for many compression algorithms. A complete list of format tags can be found in
    ' the MMREG.H header file. WAVE_FORMAT_PCM = 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let FormatTag = 1
    
    'Dim nChannels As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Number of channels in the waveform-audio data. Monaural data uses one channel and stereo data uses two channels.
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let nChannels = IIf(OptionAudioChannelNumbers1.Value = True, 1, 2)
    
    'Dim nSamplesPerSecond As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Sample rate, in samples per second (hertz), that each channel should be played or recorded. If wFormatTag is WAVE_FORMAT_PCM, then common values for nSamplesPerSec
    ' are 8.0 kHz, 11.025 kHz, 22.05 kHz, and 44.1 kHz. For non-PCM formats, this member must be computed according to the manufacturer's specification of the format tag.
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let nSamplesPerSecond = ComboBoxSamplesPerSecond.Text
    
    'Dim nAverageBytesPerSecond
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Required average data-transfer rate, in bytes per second, for the format tag. If wFormatTag is WAVE_FORMAT_PCM, nAvgBytesPerSec should be equal to the product
    ' of nSamplesPerSec and nBlockAlign. For non-PCM formats, this member must be computed according to the manufacturer's specification of the format tag.
    ' Playback and record software can estimate buffer sizes by using the nAvgBytesPerSec member.
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let nAverageBytesPerSecond = (IIf(OptionAudioBitRate8.Value = True, 1, 2)) * (IIf(OptionAudioChannelNumbers1.Value = True, 1, 2)) * (IIf(OptionAudioBitRate8.Value = True, 8, 16) / 8)
    
    'Dim nBlockAlign As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data for the wFormatTag format type. If wFormatTag is WAVE_FORMAT_PCM, nBlockAlign should
    ' be equal to the product of nChannels and wBitsPerSample divided by 8 (bits per byte). For non-PCM formats, this member must be computed according to the manufacturer's
    ' specification of the format tag.
    ' Playback and record software must process a multiple of nBlockAlign bytes of data at a time. Data written and read from a device must always start at the beginning of
    ' a block. For example, it is illegal to start playback of PCM data in the middle of a sample (that is, on a non-block-aligned boundary).
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Let nBlockAlign = (IIf(OptionAudioChannelNumbers1.Value = True, 1, 2)) * (IIf(OptionAudioBitRate8.Value = True, 8, 16) / 8)
    
   ' Dim nBitsPerSample As Long
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Bits per sample for the wFormatTag format type. If wFormatTag is WAVE_FORMAT_PCM, then wBitsPerSample should be equal to 8 or 16. For non-PCM formats, this member must
    ' be set according to the manufacturer's specification of the format tag. Note that some compression schemes cannot define a value for wBitsPerSample, so this member can
    ' be zero.
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Let nBitsPerSample = IIf(OptionAudioBitRate8.Value = True, 8, 16)
    'If VideoCapture.SetAudioFormat(1, nChannels, nSamplesPerSecond, nAverageBytesPerSecond, nBlockAlign, nBitsPerSample) = False Then
    '    MsgBox "The current selection of audio settings is not valid. Please change some of the setting to make the audio format valid.", vbExclamation + vbOKOnly, "Automatic Train Control - Invalid Audio Format"
    '    Exit Sub
    'End If

    VideoCapture.Connected = False



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
            If TemporaryScreen = "Audio Settings Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Audio Setting Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Audio Settings Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Audio Settings Screen, Button Close, stack empty, underflow."
        End If
    End If

End Sub






Private Sub ButtonPrint_Click()

    AudioSettings.PrintForm

End Sub








































Private Sub ComboBoxAudioCodec_Click()

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioCodec"
    Let Ini.Value = ComboBoxAudioCodec.Text
    Let Ini.Parameter = "AudioCodecIndex"
    Let Ini.Value = ComboBoxAudioCodec.ListIndex - 1
    Let VideoCapture.AudioCodecIndex = ComboBoxAudioCodec.ListIndex - 1

End Sub


Private Sub ComboBoxAudioDevice_Click()
        
    Let Ini.Filename = App.Path$ & "/Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioCaptureDevice"
    Let Ini.Value = ComboBoxAudioDevice.Text
    Let Ini.Parameter = "AudioCaptureDeviceIndex"
    Let Ini.Value = ComboBoxAudioDevice.ListIndex - 1
    Let VideoCapture.AudioDeviceIndex = Ini.Value
    Call GetAudioInput
    
End Sub


Private Sub ComboBoxAudioInput_Click()
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioInput"
    Let Ini.Value = ComboBoxAudioInput.Text
    Let Ini.Parameter = "AudioInputIndex"
    Let Ini.Value = ComboBoxAudioInput.ListIndex - 1
    Let VideoCapture.AudioInputIndex = Ini.Value
    
End Sub

Private Sub ComboBoxSamplesPerSecond_Click()
        
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "SamplesPerSecond"
    Let Ini.Value = ComboBoxSamplesPerSecond.Text
    Let Ini.Parameter = "SamplesPerSecondIndex"
    Let Ini.Value = ComboBoxSamplesPerSecond.ListIndex - 1
    
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
        If TemporaryScreen = "Audio Settings Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Audio Settings Screen"
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
            Let Ini.Value = "Audio Settings Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Audio Settings Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. "
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
            Let Ini.Value = "Audio Settings Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Audio Device and Set
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryF As Integer
    Let VideoCapture.CaptureAudio = True
    Let VideoCapture.Connected = True
    
    Call GetAudioDevices
    Call GetAudioSamplesPerSecond
    Call GetAudioCodec
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Audio Settings
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim FormatTag As Long
    Dim nChannel As Long
    Dim nSamplePerSecond As Long
    Dim nAverageRate As Long
    Dim nBlock As Long
    Dim nBits As Long
    
    If VideoCapture.GetAudioFormat(FormatTag, nChannel, nSamplePerSecond, nAverageRate, nBlock, nBits) = False Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Audio Settings Screen, Form Activate, GetAudioFormat failed. The program is unable to get the current setting of the capture device."
        End If
    Else
        If nChannel = 1 Then
            Let OptionAudioChannelNumbers1.Value = True
        ElseIf nChannel = 2 Then
            Let OptionAudioChannelNumbers2.Value = True
        ElseIf nchanel1 = 5 Then
            Let OptionAudioChannelNumbers5.Value = True
        End If
        If nBits = 8 Then
            Let OptionAudioBitRate8.Value = True
        ElseIf nBits = 16 Then
            Let OptionAudioBitRate16.Value = True
        ElseIf nBits = 32 Then
            Let OptionAudioBitRate32.Value = True
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
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(AudioSettings.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(AudioSettings.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(AudioSettings.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(AudioSettings.Height)

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
                Let Ini.Value = "Audio Settings Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Audio Settings Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    AudioSettings.Hide

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
    Let Ini.Application = "Audio Settings Screen"
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
        AudioSettings.Left = (Screen.Width - Width) / 2
        AudioSettings.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + AudioSettings.Width > Screen.Width Then
            Let AudioSettings.Left = Screen.Width - AudioSettings.Width
        Else
            Let AudioSettings.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + AudioSettings.Height > Screen.Height Then
            Let AudioSettings.Top = Screen.Height - AudioSettings.Height
        Else
            Let AudioSettings.Top = Val(TemporaryValueTop)
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
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
'        Dim TemporaryText1 As String
'        Dim TemporaryText2 As String
'        Dim i As Long
'        Dim BalloonFont As New StdFont
'
'        Let Ini.Filename = App.Path$ & "\Atc.ini"
'        Let Ini.Application = "All Screens"
'            Ini.Parameter = "BalloonHelpFontName"
'            Ini.Value = BalloonFont.Name
'            Ini.Parameter = "BalloonHelpFontSize"
'            Ini.Value = BalloonFont.Size
'            Ini.Parameter = "BalloonHelpColour1"
'            Colour1 = ini.value
'            Ini.Parameter = "BalloonHelpColour2"
'            Colour2 = ini.value
'            Ini.Parameter = "BalloonHelpColour3"
'            Colour3 = ini.value
'
'        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
'        Let TemporaryText2 = "Print Button"
'        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'
'        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "close the Audio Settings Window and return control" & vbCrLf & "to the main screen."
'        Let TemporaryText2 = "Close Button"
'        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
'    Else
'        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
'    End If

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

    If AudioSettings.WindowState = vbMinimized Then
    
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
        
    ElseIf AudioSettings.WindowState = vbNormal Then
    
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
        Let Ini.Value = "Video streaming connection closed from " & FromAddress & " with an error code of " & ErrorCode
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


Private Sub VideoCapture_DeviceLost()
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Log the Error
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Mainline Operation GUI Diesel1 Screen, video capture device lost."
    End If
    
    MsgBox "The software has lost the connection to the video device. Please check you cables to make sure the video device is securely attached.", vbOKOnly + vbInformation, "Automatic Train Control - Warning"
 
End Sub


Private Sub GetAudioDevices()

    Let LabelStatus.Caption = "Status: Getting Audio Devices"

    If VideoCapture.GetAudioDeviceCount <> 0 Then
        
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Audio Settings Screen"
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Poll for Audio Devices
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ComboBoxAudioDevice.Clear
        ComboBoxAudioDevice.AddItem "Please select an audio capture device."
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let TemporaryPointer = 0
        Let Ini.Parameter = "AudioCaptureDevice"
        For TemporaryF = 0 To VideoCapture.GetAudioDeviceCount - 1
            ComboBoxAudioDevice.AddItem VideoCapture.GetAudioDeviceName(TemporaryF)
            If ComboBoxAudioDevice.List(TemporaryF + 1) = Ini.Value Then
                Let TemporaryPointer = TemporaryF + 1
            End If
        Next TemporaryF

        Let ComboBoxAudioDevice.ListIndex = TemporaryPointer
        Let Ini.Parameter = "AudioCaptureDeviceIndex"
        If Val(Ini.Value) = ComboBoxAudioDevice.ListIndex - 1 Then
            Let VideoCapture.AudioDeviceIndex = Val(Ini.Value)
            Call GetAudioInput
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Audio Settings Screen, GetAudioDevices, unable to find last saved audio device."
            End If
            Let VideoCapture.AudioDeviceIndex = -1
            Let ComboBoxAudioDevice.ListIndex = 0
        End If
    End If
End Sub

Private Sub GetAudioInput()
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Audio Input and Set
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
    Let LabelStatus.Caption = "Status: Getting Audio Inputs"

    If VideoCapture.GetAudioDeviceCount <> 0 Then
        
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Audio Settings Screen"
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Poll for Audio Devices
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ComboBoxAudioInput.Clear
        ComboBoxAudioInput.AddItem "After selecting and Audio device, select a audio input."
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let TemporaryPointer = 0
        Let Ini.Parameter = "AudioInput"
        For TemporaryF = 0 To VideoCapture.GetAudioInputCount - 1
            ComboBoxAudioInput.AddItem VideoCapture.GetAudioInputName(TemporaryF)
            If ComboBoxAudioInput.List(TemporaryF + 1) = Ini.Value Then
                Let TemporaryPointer = TemporaryF + 1
            End If
        Next TemporaryF

        Let ComboBoxAudioInput.ListIndex = TemporaryPointer
        Let Ini.Parameter = "AudioInputIndex"
        If Val(Ini.Value) = ComboBoxAudioInput.ListIndex - 1 Then
            Let VideoCapture.AudioDeviceIndex = Val(Ini.Value)
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "/Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Audio Settings Screen, GetAudioInput, unable to find last saved audio input."
            End If
            Let VideoCapture.AudioDeviceIndex = -1
            Let ComboBoxAudioInput.ListIndex = 0
        End If
    End If
End Sub

Private Sub GetAudioSamplesPerSecond()

' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Audio Samples Per Second
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "/Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "SamplesPerSecond"
    On Error Resume Next
    Let ComboBoxSamplesPerSecond.Text = Ini.Value
    If Err <> 0 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Audio Settings Screen, GetAudioSamplesPerSecond, invalid value saved in INI file for index."
        End If
    End If
    On Error GoTo 0
    
End Sub

Private Sub GetAudioCodec()

    Let LabelStatus.Caption = "Status: Getting Audio Codecs"
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Get Audio Codec
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ComboBoxAudioCodec.Clear
    ComboBoxAudioCodec.AddItem "After selecting an audio capture device, please select an audio codec."
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Audio Settings Screen"
    Let Ini.Parameter = "AudioCodec"
    Let TemporaryPointer = 0
    For TemporaryF = 0 To VideoCapture.GetAudioCodecCount - 1
        ComboBoxAudioCodec.AddItem VideoCapture.GetAudioCodecName(TemporaryF)
        If ComboBoxAudioCodec.List(TemporaryF + 1) = Ini.Value Then
            Let TemporaryPointer = TemporaryF + 1
        End If
    Next TemporaryF
    Let ComboBoxAudioCodec.ListIndex = TemporaryPointer
    Let Ini.Parameter = "AudioCodecIndex"
    If Val(Ini.Value) = ComboBoxAudioCodec.ListIndex - 1 Then
        Let VideoCapture.AudioCodecIndex = Val(Ini.Value)
    Else
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "/Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Audio Settings Screen, GetAudioCodec, unable to find last saved audio codec."
        End If
        Let VideoCapture.AudioCodecIndex = -1
        Let ComboBoxAudioCodec.ListIndex = 0
    End If

End Sub
