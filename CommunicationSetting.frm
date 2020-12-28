VERSION 4.00
Begin VB.Form CommunicationSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Communication Setting"
   ClientHeight    =   5130
   ClientLeft      =   6105
   ClientTop       =   2910
   ClientWidth     =   6855
   Height          =   5535
   Icon            =   "CommunicationSetting.frx":0000
   Left            =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6855
   Top             =   2565
   Width           =   6975
   Begin VB.OptionButton OptionDCCsystemDigitrax 
      Caption         =   "Digitrax System"
      Height          =   195
      Left            =   4440
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton OptionDCCsystemNCE 
      Caption         =   "NCE or Wangraow System or"
      Height          =   195
      Left            =   1860
      TabIndex        =   18
      Top             =   1560
      Value           =   -1  'True
      Width           =   2475
   End
   Begin VB.ComboBox ComboMode 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":0442
      Left            =   3900
      List            =   "CommunicationSetting.frx":044C
      TabIndex        =   16
      Text            =   "Standard Mode"
      Top             =   2700
      Width           =   2895
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4260
      TabIndex        =   15
      Top             =   4860
      Width           =   1215
   End
   Begin VB.ComboBox ComboBaudRateSetting3 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":0472
      Left            =   3900
      List            =   "CommunicationSetting.frx":047C
      TabIndex        =   6
      Text            =   "9600 bits per second"
      Top             =   4380
      Width           =   2895
   End
   Begin VB.ComboBox ComboCommunicationPortSetting3 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":04AD
      Left            =   3900
      List            =   "CommunicationSetting.frx":04C0
      TabIndex        =   5
      Text            =   "Not Used"
      Top             =   4020
      Width           =   2895
   End
   Begin VB.ComboBox ComboBaudRateSetting2 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":0551
      Left            =   3900
      List            =   "CommunicationSetting.frx":055B
      TabIndex        =   4
      Text            =   "9600 bits per second"
      Top             =   3540
      Width           =   2895
   End
   Begin VB.ComboBox ComboCommunicationPortSetting2 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":058C
      Left            =   3900
      List            =   "CommunicationSetting.frx":059F
      TabIndex        =   3
      Text            =   "Not Used"
      Top             =   3180
      Width           =   2895
   End
   Begin VB.ComboBox ComboBaudRateSetting1 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":0630
      Left            =   3900
      List            =   "CommunicationSetting.frx":064F
      TabIndex        =   2
      Text            =   "9600 bits per second"
      Top             =   2340
      Width           =   2895
   End
   Begin VB.ComboBox ComboCommunicationPortSetting1 
      Height          =   315
      ItemData        =   "CommunicationSetting.frx":06AD
      Left            =   3900
      List            =   "CommunicationSetting.frx":06EB
      TabIndex        =   1
      Text            =   "Not Used"
      Top             =   1980
      Width           =   2895
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "CommunicationSetting.frx":0921
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   5580
      TabIndex        =   8
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "I'm connected to a "
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   1395
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   3720
      Top             =   4920
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "using digital commands in"
      Height          =   195
      Left            =   1920
      TabIndex        =   17
      Top             =   2700
      Width           =   1800
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6780
      X2              =   60
      Y1              =   1320
      Y2              =   1320
   End
   Begin IniconLib.Init Ini 
      Left            =   2580
      Top             =   4860
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
      Left            =   3180
      Top             =   4860
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "and the baud rate is set to"
      Height          =   195
      Left            =   1860
      TabIndex        =   14
      Top             =   4440
      Width           =   1845
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Communication Port (not used at this time)."
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   4080
      Width           =   3000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "and the baud rate is set to"
      Height          =   195
      Left            =   1860
      TabIndex        =   12
      Top             =   3600
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Communication Port for your CMRI system is set to"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   3240
      Width           =   3540
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6780
      X2              =   0
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6780
      X2              =   0
      Y1              =   3930
      Y2              =   3930
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6780
      X2              =   0
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Label Label5 
      Caption         =   $"CommunicationSetting.frx":0D63
      Height          =   855
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Communication Port for your DCC System is set to"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "and the baud rate is set to"
      Height          =   195
      Left            =   1860
      TabIndex        =   10
      Top             =   2400
      Width           =   1845
   End
End
Attribute VB_Name = "CommunicationSetting"
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
            If TemporaryScreen = "Communication Setting Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Button Close, current window is not listed in the stack to remove it and hide."
            End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Open Previous Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Screen Stack"
            Let Ini.Parameter = CStr(TemporaryCounter - 1)
            Let TemporaryScreen = Ini.Value
            If TemporaryScreen = "Communication Setting Screen" Then
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
                Let Ini.Value = "Communication Setting Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Communication Setting Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub






Private Sub ButtonPrint_Click()

    CommunicationSetting.PrintForm
    
End Sub

Private Sub ComboCommunicationPortSetting1_Change()
4
    If ComboCommunicationPortSetting1.Text <> "Not Used" Then
        If ComboCommunicationPortSetting1.Text = ComboCommunicationPortSetting2.Text Then
            MsgBox "You have selected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        ElseIf ComboCommunicationPortSetting1.Text = ComboCommunicationPortSetting3.Text Then
            MsgBox "You have selected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        Else
            Let ButtonClose.Enabled = True
        End If
    Else
        Let ButtonClose.Enabled = True
    End If
    
End Sub

Private Sub ComboCommunicationPortSetting2_Click()

    If ComboCommunicationPortSetting2.Text <> "Not Used" Then
        If ComboCommunicationPortSetting2.Text = ComboCommunicationPortSetting1.Text Then
            MsgBox "You have selected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        ElseIf ComboCommunicationPortSetting2.Text = ComboCommunicationPortSetting3.Text Then
            MsgBox "You have selected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        Else
            Let ButtonClose.Enabled = True
        End If
    Else
        Let ButtonClose.Enabled = True
    End If

End Sub


Private Sub ComboCommunicationPortSetting3_Click()

    If ComboCommunicationPortSetting3.Text <> "Not Used" Then
        If ComboCommunicationPortSetting3.Text = ComboCommunicationPortSetting1.Text Then
            MsgBox "You have slected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        ElseIf ComboCommunicationPortSetting3.Text = ComboCommunicationPortSetting2.Text Then
            MsgBox "You have slected a communication port (serial port) that has already been" & vbCrLf & "assigned. Please selected another port or change the other.", vbOKOnly + vbExclamation, "Automatic Train Control - User Selected Communication Port"
            ButtonClose.Enabled = False
        Else
            Let ButtonClose.Enabled = True
        End If
    Else
        Let ButtonClose.Enabled = True
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
        If TemporaryScreen = "Communication Setting Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Communication Setting Screen"
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
            Let Ini.Value = "Communication Setting Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Communication Setting Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Communication Setting Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    Call BalloonHelpUpdatePart01
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Communication Settings Screen"
    Let Ini.Parameter = "CommunicationPortForDCC"
    Let Ini.Value = ComboCommunicationPortSetting1.Text
    Let Ini.Parameter = "BaudRateForDCC"
    Let Ini.Value = ComboBaudRateSetting1.Text
    Let Ini.Parameter = "Mode"
    Let Ini.Value = ComboMode.Text
    Let Ini.Parameter = "CommunicationPortForCMRI"
    Let Ini.Value = ComboCommunicationPortSetting2.Text
    Let Ini.Parameter = "BaudRateForCMRI"
    Let Ini.Value = ComboBaudRateSetting2.Text
    Let Ini.Parameter = "CommunicationPortForOTHER"
    Let Ini.Value = ComboCommunicationPortSetting3.Text
    Let Ini.Parameter = "BaudRateForOTHER"
    Let Ini.Value = ComboBaudRateSetting3.Text
    Let Ini.Parameter = "OptionDCCsystemNCE"
    Let Ini.Value = optiondccsystemnce.Value
    Let Ini.Parameter = "OptionDCCSystemDigitrax"
    Let Ini.Value = optiondccsystemDigitrax.Value

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transfer Values to MainScreen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainScreen!LabelCommunicationPortSetting1.Text = ComboCommunicationPortSetting1.Text
    Let MainScreen!LabelBaudRateSetting1.Text = ComboBaudRateSetting1.Text
    Let MainScreen!Labelmode.Text = ComboMode.Text
    Let MainScreen!LabelCommunicationPortSetting2.Text = ComboCommunicationPortSetting2.Text
    Let MainScreen!LabelBaudRateSetting2.Text = ComboBaudRateSetting2.Text
    Let MainScreen!LabelCommunicationPortSetting3.Text = ComboCommunicationPortSetting3.Text
    Let MainScreen!LabelBaudRateSetting3.Text = ComboBaudRateSetting3.Text

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Communication Settings Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(CommunicationSetting.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(CommunicationSetting.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(CommunicationSetting.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(CommunicationSetting.Height)

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
                Let Ini.Value = "Communication Setting Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Communication Setting Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    CommunicationSetting.Hide


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
    Let Ini.Application = "Communication Settings Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "CommunicationPortForDCC"
    Let ComboCommunicationPortSetting1.Text = Ini.Value
    Let Ini.Parameter = "BaudRateForDCC"
    Let ComboBaudRateSetting1.Text = Ini.Value
    Let Ini.Parameter = "Mode"
    Let ComboMode.Text = Ini.Value
    Let Ini.Parameter = "CommunicationPortForCMRI"
    Let ComboCommunicationPortSetting2.Text = Ini.Value
    Let Ini.Parameter = "BaudRateForCMRI"
    Let ComboBaudRateSetting2.Text = Ini.Value
    Let Ini.Parameter = "CommunicationPortForOTHER"
    Let ComboCommunicationPortSetting3.Text = Ini.Value
    Let Ini.Parameter = "BaudRateForOTHER"
    Let ComboBaudRateSetting3.Text = Ini.Value
    Let Ini.Parameter = "OptionDCCsystemNCE"
    Let optiondccsystemnce.Value = Ini.Value
    Let Ini.Parameter = "OptionDCCSystemDigitrax"
    Let optiondccsystemDigitrax.Value = Ini.Value

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        CommunicationSetting.Left = (Screen.Width - Width) / 2
        CommunicationSetting.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + CommunicationSetting.Width > Screen.Width Then
            Let CommunicationSetting.Left = Screen.Width - CommunicationSetting.Width
        Else
            Let CommunicationSetting.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + CommunicationSetting.Height > Screen.Height Then
            Let CommunicationSetting.Top = Screen.Height - CommunicationSetting.Height
        Else
            Let CommunicationSetting.Top = Val(TemporaryValueTop)
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
    'No Databases
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
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let CommunicationSetting.MousePointer = ccHourglass
       
        Let BalloonHelpText1 = "This drop down combintion box is to select what communication" & vbCrLf & "(serial) port is connected to your DCC station."
        Let BalloonHelpText2 = "Communication Port for DCC"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboCommunicationPortSetting1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboCommunicationPortSetting1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboCommunicationPortSetting1' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This drop down combintion box is to select what baud rate the communication" & vbCrLf & "(serial) port is connected to your DCC station."
        Let BalloonHelpText2 = "Communication Setting  for DCC"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBaudRateSetting1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBaudRateSetting1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboBaudRateSetting1' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This drop down combintion box is to select what communication" & vbCrLf & "mode is used with your DCC station."
        Let BalloonHelpText2 = "Communication Mode for CMRI"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboMode)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboMode, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboMode' object."
            End If
        End If

        Let BalloonHelpText1 = "This drop down combintion box is to select what communication" & vbCrLf & "(serial) port is connected to your CMRI."
        Let BalloonHelpText2 = "Communication Port for CMRI"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboCommunicationPortSetting2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboCommunicationPortSetting2, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboCommunicationPortSetting2' object."
            End If
        End If
                
        Let BalloonHelpText1 = "This drop down combintion box is to select what baud rate the communication" & vbCrLf & "(serial) port is connected to your CMRI."
        Let BalloonHelpText2 = "Communication Setting for CMRI"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBaudRateSetting2)
                Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBaudRateSetting2, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboBaudRateSettting2' object."
            End If
        End If
        
        Let BalloonHelpText1 = "This drop down combintion box is to select what communication" & vbCrLf & "(serial) port is connected to your OTHER device."
        Let BalloonHelpText2 = "Communication Port for OTHER"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboCommunicationPortSetting3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboCommunicationPortSetting3, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboCommunicationPortSetting3' object."
            End If
        End If
                
        Let BalloonHelpText1 = "This drop down combintion box is to select what baud rate the communication" & vbCrLf & "(serial) port is connected to your OTHER device."
        Let BalloonHelpText2 = "Communication Setting for OTHER"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ComboBaudRateSetting3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ComboBaudRateSetting3, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ComboBaudRateSetting3' object."
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
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ButtonPrint' object."
            End If
        End If
                
        Let BalloonHelpText1 = "This button will close the Communication Setting window and" & vbCrLf & "return control to the main screen."
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
                Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ButtonClose' object."
            End If
        End If
        
        Let CommunicationSetting.MousePointer = ccDefault

    Else 'If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is Off" Then
    
        Let ballonhelpsetup = balloonhelp.DestroyAllToolTips
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Communication Setting Screen, Balloon Help Update, unable to setup balloon help for 'ButtonClose' object."
        End If
    End If
End Sub

Private Sub Form_Resize()

    If CommunicationSetting.WindowState = vbMinimized Then
    
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
        
    ElseIf CommunicationSetting.WindowState = vbNormal Then
    
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




Private Sub OptionDCCsystemDigitrax_Click()
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Digitrax unit does not support standard mode
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let ComboMode.Text = "Non-standard Mode"
    Let ComboMode.Enabled = False
    Let ComboBaudRateSetting1.Text = "16457 bits per second"

End Sub


Private Sub optiondccsystemnce_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' NCE unit does both standard mode and non-standard
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let ComboMode.Enabled = True

End Sub


