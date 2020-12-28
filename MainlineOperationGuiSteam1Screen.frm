VERSION 4.00
Begin VB.Form MainlineOperationGuiSteam1Screen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Automatic Train Control - Mainline Operation - GUI for Steam"
   ClientHeight    =   11895
   ClientLeft      =   2865
   ClientTop       =   945
   ClientWidth     =   16560
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Height          =   12300
   Icon            =   "MainlineOperationGuiSteam1Screen.frx":0000
   Left            =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12328.42
   ScaleMode       =   0  'User
   ScaleWidth      =   16690.39
   Top             =   600
   Width           =   16680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictureBoxLocomotiveCab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   783.299
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
      Begin VB.CommandButton ButtonStartEngine 
         Caption         =   "&Start Engine"
         Height          =   255
         Left            =   13920
         TabIndex        =   17
         Top             =   74
         Width           =   1230
      End
      Begin VB.PictureBox PictureBoxInjectorSteamValveLive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   30
         ScaleHeight     =   1575
         ScaleWidth      =   2010
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   3270
         Width           =   2010
      End
      Begin VB.PictureBox PictureBoxInjectorSteamValveExhaust 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   5040
         ScaleHeight     =   1545
         ScaleWidth      =   2010
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   3285
         Width           =   2010
      End
      Begin VB.PictureBox PictureBoxDamper 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   6510
         ScaleHeight     =   1890
         ScaleWidth      =   1335
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   9615
         Width           =   1335
      End
      Begin VB.PictureBox PictureBoxFireBoxDoor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   300
         ScaleHeight     =   3210
         ScaleWidth      =   4455
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8205
         Width           =   4455
      End
      Begin VB.PictureBox PictureBoxCylinderCock 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   6000
         ScaleHeight     =   3060
         ScaleWidth      =   405
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8460
         Width           =   405
      End
      Begin VB.PictureBox PictureBoxRegulator 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5475
         Left            =   10035
         ScaleHeight     =   5475
         ScaleWidth      =   1695
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   2475
         Width           =   1695
      End
      Begin VB.PictureBox PictureBoxSmallInjectorCompressor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   12180
         ScaleHeight     =   960
         ScaleWidth      =   645
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   5550
         Width           =   645
      End
      Begin VB.PictureBox PictureBoxAutomaticBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   13230
         ScaleHeight     =   3120
         ScaleWidth      =   1365
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   6045
         Width           =   1365
      End
      Begin VB.PictureBox PictureBoxSand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   2790
         ScaleHeight     =   1320
         ScaleWidth      =   165
         TabIndex        =   8
         Tag             =   "0"
         Top             =   5805
         Width           =   165
      End
      Begin VB.PictureBox PictureBoxInjectorWaterValveLive 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   13260
         ScaleHeight     =   420
         ScaleWidth      =   1005
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   10440
         Width           =   1005
      End
      Begin VB.PictureBox PictureBoxInjectorWaterValveExhaust 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   13860
         ScaleHeight     =   420
         ScaleWidth      =   1005
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   10845
         Width           =   1005
      End
      Begin VB.PictureBox PictureBoxBlower 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   120
         ScaleHeight     =   1905
         ScaleWidth      =   1785
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   735
         Width           =   1785
      End
      Begin VB.CommandButton ButtonHelp 
         Caption         =   "&Help is Off"
         Height          =   255
         Left            =   13920
         TabIndex        =   4
         Top             =   1059
         Width           =   1230
      End
      Begin VB.PictureBox PictureBoxPointer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   9656
         ScaleHeight     =   345
         ScaleWidth      =   390
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4860
         Width           =   390
      End
      Begin VB.PictureBox PictureBoxReverser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   8295
         ScaleHeight     =   1440
         ScaleWidth      =   4650
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8670
         Width           =   4650
      End
      Begin VB.CommandButton ButtonClose 
         Caption         =   "&Close"
         Height          =   255
         Left            =   13920
         TabIndex        =   1
         Top             =   1912
         Width           =   1230
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   15960
      Top             =   3060
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   15960
      Top             =   2520
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   15960
      Top             =   1920
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Menu menuCaptureDevice 
      Caption         =   "Capture Device"
      Visible         =   0   'False
      Begin VB.Menu menuCaptureDeviceVideoSource 
         Caption         =   "Video Source"
      End
      Begin VB.Menu menuCaptureDeviceAudioSetting 
         Caption         =   "Audio Setting"
      End
      Begin VB.Menu menuCaptureDeviceVideoFormat 
         Caption         =   "Video Format"
      End
      Begin VB.Menu menuCaptureDeviceVideoCompression 
         Caption         =   "Video Compression"
      End
      Begin VB.Menu menuCaptureDeviceVideoDisplay 
         Caption         =   "Video Display"
      End
   End
End
Attribute VB_Name = "MainlineOperationGuiSteam1Screen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


















Private Sub ButtonCaption_Click()

If ButtonCaption.Caption = "&Caption is On" Then
    Let ButtonCaption.Caption = "&Caption is Off"
Else
    Let ButtonCaption.Caption = "&Caption is On"
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
            If TemporaryScreen = "Mainline Operation GUI Steam1 Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonHelp_Click()

If ButtonHelp.Caption = "&Help is Off" Then
    Let ButtonHelp.Caption = "&Help is On"
Else
    Let ButtonHelp.Caption = "&Help is Off"
End If

End Sub

Private Sub ButtonVideo_Click()

If ButtonVideo.Caption = "Video is Off" Then
    Let ButtonVideo.Caption = "Video is On"
    Let VideoCapture.Visible = True
Else
    Let ButtonVideo.Caption = "Video is Off"
    Let VideoCapture.Visible = False
End If

End Sub

Private Sub ButtonVideoSettings_Click()

With VideoCapture
    If .HasAudio Then
        Let menuCaptureDeviceAudioSetting.Enabled = False
    Else
        Let menuCaptureDeviceAudioSetting.Enabled = True
    End If
    If .HasDlgFormat Then
        Let menuCaptureDeviceVideoFormat.Enabled = True
    Else
        Let menuCaptureDeviceVideoFormat.Enabled = False
    End If
    If .HasDlgDisplay Then
        Let menuCaptureDeviceVideoDisplay.Enabled = True
    Else
        Let menuCaptureDeviceVideoDisplay.Enabled = False
    End If
    If .HasDlgSource Then
        Let menuCaptureDeviceVideoSource.Enabled = True
    Else
        Let menuCaptureDeviceVideoSource.Enabled = False
    End If
End With

MainlineOperationGuiSteam1Screen.PopupMenu menuCaptureDevice

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
        If TemporaryScreen = "Mainline Operation GUI Steam1 Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Operation GUI Steam1 Screen"
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
            Let Ini.Value = "Mainline Operation GI Steam1 Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Operation GUI Steam1 Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineOperationGuiSteam1Screen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineOperationGuiSteam1Screen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineOperationGuiSteam1Screen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineOperationGuiSteam1Screen.Height)

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
                Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
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
            Let Ini.Value = "Mainline Operation GUI Steam1 Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineOperationGuiSteam1Screen.Hide
    'unload Mainlineoperationguisteam1screen

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub


Private Sub Form_Load()


TemporaryLocomotivePath$ = "\Graphics\Locomotive Steam1\"
Let PictureBoxLocomotiveCab.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "CabScreen(s1).bmp")
Let PictureBoxAutomaticBrake.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "AutomaticBrake" + PictureBoxAutomaticBrake.Tag + "(s1).bmp")
Let PictureBoxBlower.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "Blower" + PictureBoxBlower.Tag + "(s1).bmp")
Let PictureBoxCylinderCock.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "CylinderCock" + PictureBoxCylinderCock.Tag + "(s1).bmp")
Let PictureBoxDamper.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "Damper" + PictureBoxDamper.Tag + "(s1).bmp")
Let PictureBoxFireBoxDoor.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "FireBoxDoor" + PictureBoxFireBoxDoor.Tag + "(s1).bmp")
Let PictureBoxInjectorSteamValveExhaust.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "InjectorSteamValveExhaust" + PictureBoxInjectorSteamValveExhaust.Tag + "(s1).bmp")
Let PictureBoxInjectorSteamValveLive.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "InjectorSteamValveLive" + PictureBoxInjectorSteamValveLive.Tag + "(s1).bmp")
Let PictureBoxInjectorWaterValveExhaust.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "InjectorWaterValveExhaust" + PictureBoxInjectorWaterValveExhaust.Tag + "(s1).bmp")
Let PictureBoxInjectorWaterValveLive.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "InjectorWaterValveLive" + PictureBoxInjectorWaterValveLive.Tag + "(s1).bmp")
Let PictureBoxRegulator.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "Regulator" + PictureBoxRegulator.Tag + "(s1).bmp")
Let PictureBoxReverser.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "Reverser" + PictureBoxReverser.Tag + "(s1).bmp")
Let PictureBoxSand.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "Sand" + PictureBoxSand.Tag + "(s1).bmp")
Let PictureBoxSmallInjectorCompressor.Picture = LoadPicture(App.Path$ & TemporaryLocomotivePath$ + "SmallInjectorCompressor" + PictureBoxSmallInjectorCompressor.Tag + "(s1).bmp")

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If
 
If Screen.Width > PictureBoxLocomotiveCab.ScaleWidth And Screen.Height > PictureBoxLocomotiveCab.Height Then
   Let MainlineOperationGuiSteam1Screen.Caption = "Automatic Train Control - Mainline Mode - GUI for Flying Scottsman"
   Let MainlineOperationGuiSteam1Screen.WindowState = vbNormal
   Let MainlineOperationGuiSteam1Screen.Width = (PictureBoxLocomotiveCab.ScaleWidth + 6) * 15
   Let MainlineOperationGuiSteam1Screen.Height = (PictureBoxLocomotiveCab.ScaleHeight + 25) * 15
   Let MainlineOperationGuiSteam1Screen.Left = (Screen.Width - MainlineOperationGuiSteam1Screen.Width) / 2   ' Center form horizontally.
   Let MainlineOperationGuiSteam1Screen.Top = (Screen.Height - MainlineOperationGuiSteam1Screen.Height) / 2  ' Center form vertiCally.
End If

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



























Private Sub Form_Resize()

    If MainlineOperationGuiSteam1Screen.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineOperationGuiSteam1Screen.WindowState = vbNormal Then
    
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

Private Sub PictureBoxAutomaticBrake_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) + 1

End Sub

Private Sub PictureBoxAutomaticBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxAutomaticBrake.Tag) > 0 Then
        Let PictureBoxAutomaticBrake.Tag = Trim$(Str$(Val(PictureBoxAutomaticBrake.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the automatic brake (trainline brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxAutomaticBrake.Tag) < 11 Then
        Let PictureBoxAutomaticBrake.Tag = Trim$(Str$(Val(PictureBoxAutomaticBrake.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the automatic brake (trainline brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\AutomaticBrake"
Let Temporary$ = Temporary$ + PictureBoxAutomaticBrake.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxAutomaticBrake.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxBlower_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxBlower.Left = Val(PictureBoxBlower.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxBlower.Left = Val(PictureBoxBlower.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxBlower.Top = Val(PictureBoxBlower.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxBlower.Top = Val(PictureBoxBlower.Top) + 1

End Sub

Private Sub PictureBoxBlower_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxBlower.Tag) > 0 Then
        Let PictureBoxBlower.Tag = Trim$(Str$(Val(PictureBoxBlower.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the blower."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxBlower.Tag) < 20 Then
        Let PictureBoxBlower.Tag = Trim$(Str$(Val(PictureBoxBlower.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the blower."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Blower"
Let Temporary$ = Temporary$ + PictureBoxBlower.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxBlower.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxCylinderCock_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxCylinderCock.Left = Val(PictureBoxCylinderCock.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxCylinderCock.Left = Val(PictureBoxCylinderCock.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxCylinderCock.Top = Val(PictureBoxCylinderCock.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxCylinderCock.Top = Val(PictureBoxCylinderCock.Top) + 1

End Sub

Private Sub PictureBoxCylinderCock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxCylinderCock.Tag) > 0 Then
        Let PictureBoxCylinderCock.Tag = Trim$(Str$(Val(PictureBoxCylinderCock.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the cylinder cock."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxCylinderCock.Tag) < 1 Then
        Let PictureBoxCylinderCock.Tag = Trim$(Str$(Val(PictureBoxCylinderCock.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the cylinder cock."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\CylinderCock"
Let Temporary$ = Temporary$ + PictureBoxCylinderCock.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxCylinderCock.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxDamper_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxDamper.Left = Val(PictureBoxDamper.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxDamper.Left = Val(PictureBoxDamper.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxDamper.Top = Val(PictureBoxDamper.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxDamper.Top = Val(PictureBoxDamper.Top) + 1

End Sub


Private Sub PictureBoxDamper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxDamper.Tag) > 0 Then
        Let PictureBoxDamper.Tag = Trim$(Str$(Val(PictureBoxDamper.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the damper."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxDamper.Tag) < 3 Then
        Let PictureBoxDamper.Tag = Trim$(Str$(Val(PictureBoxDamper.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the damper."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Damper"
Let Temporary$ = Temporary$ + PictureBoxDamper.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxDamper.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxFireBoxDoor_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxFireBoxDoor.Left = Val(PictureBoxFireBoxDoor.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxFireBoxDoor.Left = Val(PictureBoxFireBoxDoor.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxFireBoxDoor.Top = Val(PictureBoxFireBoxDoor.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxFireBoxDoor.Top = Val(PictureBoxFireBoxDoor.Top) + 1

End Sub


Private Sub PictureBoxFireBoxDoor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxFireBoxDoor.Tag) > 0 Then
        Let PictureBoxFireBoxDoor.Tag = Trim$(Str$(Val(PictureBoxFireBoxDoor.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the fire box door."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxFireBoxDoor.Tag) < 4 Then
        Let PictureBoxFireBoxDoor.Tag = Trim$(Str$(Val(PictureBoxFireBoxDoor.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the fire box door."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\FireBoxDoor"
Let Temporary$ = Temporary$ + PictureBoxFireBoxDoor.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxFireBoxDoor.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorSteamValveExhaust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxInjectorSteamValveExhaust.Tag) > 0 Then
        Let PictureBoxInjectorSteamValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveExhaust.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the injector steam valve (exhaust)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxInjectorSteamValveExhaust.Tag) < 1 Then
        Let PictureBoxInjectorSteamValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveExhaust.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the injector steam valve (exhaust)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorSteamValveExhaust"
Let Temporary$ = Temporary$ + PictureBoxInjectorSteamValveExhaust.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorSteamValveExhaust.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorSteamValveLive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxInjectorSteamValveLive.Tag) > 0 Then
        Let PictureBoxInjectorSteamValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveLive.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the injector steam valve (live)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxInjectorSteamValveLive.Tag) < 1 Then
        Let PictureBoxInjectorSteamValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveLive.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the injector steam valve (live)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorSteamValveLive"
Let Temporary$ = Temporary$ + PictureBoxInjectorSteamValveLive.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorSteamValveLive.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorWaterValveExhaust_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxInjectorWaterValveExhaust.Left = Val(PictureBoxInjectorWaterValveExhaust.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxInjectorWaterValveExhaust.Left = Val(PictureBoxInjectorWaterValveExhaust.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxInjectorWaterValveExhaust.Top = Val(PictureBoxInjectorWaterValveExhaust.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxInjectorWaterValveExhaust.Top = Val(PictureBoxInjectorWaterValveExhaust.Top) + 1

End Sub

Private Sub PictureBoxInjectorWaterValveExhaust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxInjectorWaterValveExhaust.Tag) > 0 Then
        Let PictureBoxInjectorWaterValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveExhaust.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the injector water valve (exhaust)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxInjectorWaterValveExhaust.Tag) < 9 Then
        Let PictureBoxInjectorWaterValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveExhaust.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the injector water valve (exhaust)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorWaterValveExhaust"
Let Temporary$ = Temporary$ + PictureBoxInjectorWaterValveExhaust.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorWaterValveExhaust.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorWaterValveLive_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxInjectorWaterValveLive.Left = Val(PictureBoxInjectorWaterValveLive.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxInjectorWaterValveLive.Left = Val(PictureBoxInjectorWaterValveLive.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxInjectorWaterValveLive.Top = Val(PictureBoxInjectorWaterValveLive.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxInjectorWaterValveLive.Top = Val(PictureBoxInjectorWaterValveLive.Top) + 1

End Sub

Private Sub PictureBoxInjectorWaterValveLive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxInjectorWaterValveLive.Tag) > 0 Then
        Let PictureBoxInjectorWaterValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveLive.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the injector water valve (live)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxInjectorWaterValveLive.Tag) < 9 Then
        Let PictureBoxInjectorWaterValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveLive.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the injector water valve (live)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorWaterValveLive"
Let Temporary$ = Temporary$ + PictureBoxInjectorWaterValveLive.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorWaterValveLive.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxPointer_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxPointer.Left = Val(PictureBoxPointer.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxPointer.Left = Val(PictureBoxPointer.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxPointer.Top = Val(PictureBoxPointer.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxPointer.Top = Val(PictureBoxPointer.Top) + 1

End Sub


Private Sub PictureBoxRegulator_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxRegulator.Left = Val(PictureBoxRegulator.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxRegulator.Left = Val(PictureBoxRegulator.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxRegulator.Top = Val(PictureBoxRegulator.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxRegulator.Top = Val(PictureBoxRegulator.Top) + 1

End Sub

Private Sub PictureBoxRegulator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxRegulator.Tag) > 0 Then
        Let PictureBoxRegulator.Tag = Trim$(Str$(Val(PictureBoxRegulator.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the regulator."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxRegulator.Tag) < 11 Then
        Let PictureBoxRegulator.Tag = Trim$(Str$(Val(PictureBoxRegulator.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the regulator."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Regulator"
Let Temporary$ = Temporary$ + PictureBoxRegulator.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxRegulator.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxReverser_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) + 1

End Sub

Private Sub PictureBoxReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxReverser.Tag) > -15 Then
        Let PictureBoxReverser.Tag = Trim$(Str$(Val(PictureBoxReverser.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxReverser.Tag) < 21 Then
        Let PictureBoxReverser.Tag = Trim$(Str$(Val(PictureBoxReverser.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the independent brake (locomotive brake)."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Reverser"
Let Temporary$ = Temporary$ + PictureBoxReverser.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxReverser.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxSand_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) + 1

End Sub

Private Sub PictureBoxSand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxSand.Tag) > 0 Then
        Let PictureBoxSand.Tag = Trim$(Str$(Val(PictureBoxSand.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the sand lever."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxSand.Tag) < 1 Then
        Let PictureBoxSand.Tag = Trim$(Str$(Val(PictureBoxSand.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the sand lever."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Sand"
Let Temporary$ = Temporary$ + PictureBoxSand.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxSand.Picture = LoadPicture(Temporary$)

End Sub



Private Sub PictureBoxSmallInjectorCompressor_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxSmallInjectorCompressor.Left = Val(PictureBoxSmallInjectorCompressor.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxSmallInjectorCompressor.Left = Val(PictureBoxSmallInjectorCompressor.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxSmallInjectorCompressor.Top = Val(PictureBoxSmallInjectorCompressor.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxSmallInjectorCompressor.Top = Val(PictureBoxSmallInjectorCompressor.Top) + 1

End Sub


Private Sub PictureBoxSmallInjectorCompressor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) > 0 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressor.Tag) - 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the minimum application of the small injector compressor."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) < 1 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressor.Tag) + 1))
        Let MainlineOperationGUI!Wave1.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        If ButtonHelp.Caption = "&Help is On" Then
            Let TemporaryPrompt = "You have reached the maximum application of the small injector compressor."
            MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
        End If
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\SmallInjectorCompressor"
Let Temporary$ = Temporary$ + PictureBoxSmallInjectorCompressor.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxSmallInjectorCompressor.Picture = LoadPicture(Temporary$)

End Sub


