VERSION 4.00
Begin VB.Form UtilitiesForCommandControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Utilities for Command Control"
   ClientHeight    =   4905
   ClientLeft      =   3045
   ClientTop       =   7845
   ClientWidth     =   6945
   Height          =   5310
   Icon            =   "UtilitiesForCommandControl.frx":0000
   Left            =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6945
   Top             =   7500
   Width           =   7065
   Begin VB.OptionButton OptionNewerEPROM 
      Caption         =   "Newer EPROM version (6.0.0)"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   2535
   End
   Begin VB.OptionButton OptionOlderEPROM 
      Caption         =   "Older EPROM version (4.0.1)"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton ButtonLeadLocomotiveBackup 
      Caption         =   "Backup"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton ButtonLeadLocomotiveRestore 
      Caption         =   "Restore"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton ButtonMacroBackup 
      Caption         =   "Backup"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton ButtonMacroRestore 
      Caption         =   "Restore"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "UtilitiesForCommandControl.frx":0442
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
      Left            =   5640
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7200
      Top             =   1140
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Line Line3 
      X1              =   6840
      X2              =   120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Caption         =   $"UtilitiesForCommandControl.frx":0884
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   6735
   End
   Begin VB.Label LabelLeadLocomotives 
      Caption         =   "Lead Locomotives"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3300
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   $"UtilitiesForCommandControl.frx":0942
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status: Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Your digital command control unit can store up to two hundred macros, each entry is twenty bytes."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label LabelMacroTitle 
      Caption         =   "Macros"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   525
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7200
      Top             =   600
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7200
      Top             =   0
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
Attribute VB_Name = "UtilitiesForCommandControl"
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
            If TemporaryScreen = "Utilities for Command Control Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Utilities for Command Control Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
            ElseIf temporaryscereen = "Utilities for Command Control" Then
                UtilitiesForCommandControl.Show vbModeless
            ElseIf TemporaryScreen = "Video Settings Screen" Then
                VideoSettings.Show vbModeless
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Utilities for Command Control Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Utilities for Command Control Screen, Button Close, stack empty, underflow."
        End If
    End If

End Sub


Private Sub ButtonLeadLocomotiveBackup_Click()

    Call ObjectsDisable
    
    Dim TemporaryLimitOne As Long
    Dim TemporaryLimitTwo As Long
    Dim TemporaryLocomotiveCounter As Integer
    Dim TemporaryString As String
    Dim TemporaryVersion As String
    Dim TemporaryLocomotiveAddress As Integer
    Dim TemporaryLocomotiveMemoryInside As Long
    Dim TemporaryLocomotiveMemoryOutside As Long
    
    Let LabelStatus.Caption = "Status: Checking for EPROM version."
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command to DCC system
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainScreen!MSComm1.InBufferCount = 0
    MainScreen!MSComm1.Output = Chr$(&HAA)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While MainScreen!MSComm1.InBufferCount < 3
        Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
        DoEvents
    Loop
    Let TemporaryInput = MainScreen!MSComm1.Input
    If Len(TemporaryInput) <> 3 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Lead Locomotive Backup, wrong number of characters received from command uit after sending &HAA command."
        End If
    End If
    Let TemporaryVersion = Asc(Left$(TemporaryInput, 1)) & "." & Asc(Mid$(TemporaryInput, 2, 1)) & "." & Asc(Right$(TemporaryInput, 1))
    Let LabelStatus.Caption = "Status: EPROM version is " & TemporaryVersion
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Values for EPROM version
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryVersion = "4.0.1" Then
        OptionOlderEPROM.Value = True
        Let TemporaryLimitOne = 61952
        Let TemporaryLimitTwo = 62207
    ElseIf TemporaryVersion = "6.0.0" Then
        OptionNewerEPROM.Value = True
        Let TemporaryLimitOne = 62720
        Let TemporaryLimitTwo = 62975
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Lead Locomotive Backup, unknown version number of EPROM or command unit."
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Reading
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Utilities for Command Control Screen"
    If TemporaryVersion = "4.0.1" Then
        For TemporaryLocomotiveMemoryOutside = TemporaryLimitOne To TemporaryLimitTwo Step 2
            Let TemporaryString = ""
            Let TemporaryLocomotiveAddress = 0
            For TemporaryLocomotiveMemoryInside = TemporaryLocomotiveMemoryOutside To TemporaryLocomotiveMemoryOutside + 1
                Let MainScreen!MSComm1.InBufferCount = 0
                Let a = Int(TemporaryLocomotiveMemoryInside / 256)
                Let b = TemporaryLocomotiveMemoryInside - (a * 256)
                Let MainScreen!MSComm1.Output = Chr$(&H9D) & Chr$(a) & Chr$(b)
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Wait for Response
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Do While MainScreen!MSComm1.InBufferCount < 1
                    DoEvents
                Loop
                Let TemporaryInput = MainScreen!MSComm1.Input
                If Len(TemporaryInput) <> 1 Then
                    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        Let Ini.Filename = App.Path$ & "\Atc.log"
                        Let Ini.Application = "Log Errors"
                        Let Ini.Parameter = Date$ & " " & Time$
                        Let Ini.Value = "Utility Screen, Button Lead Locomotive Backup, wrong number of bytes returned from digital command control unit after sending &H9D command."
                    End If
                End If
                LabelStatus.Caption = "Status: Reading memory location " & Hex(TemporaryLocomotiveMemoryInside)
                DoEvents
                If Hex(Asc(TemporaryInput)) = "0" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "1" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "2" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "3" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "4" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "5" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "6" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "7" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "8" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "9" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "A" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "B" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "C" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "D" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "E" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                ElseIf Hex(Asc(TemporaryInput)) = "F" Then
                    Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
                Else
                    Let TemporaryString = TemporaryString & Hex(Asc(TemporaryInput)) & " "
                End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate the locomotive address
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                If TemporaryLocomotiveMemoryInside = TemporaryLocomotiveMemoryOutside Then
                    If Asc(TemporaryInput) <> 0 Then
                        Let TemporaryLocomotiveAddress = (Asc(TemporaryInput) - 128 - 64) * 256
                    End If
                Else
                    Let TemporaryLocomotiveAddress = TemporaryLocomotiveAddress + Asc(TemporaryInput)
                End If
                
            Next TemporaryLocomotiveMemoryInside
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Utilities for Command Control Screen"
            Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryLocomotiveMemoryOutside) & " to 0x" & Hex(TemporaryLocomotiveMemoryOutside + 1)
            Let Ini.Value = TemporaryString & ";Consist number" & Str$(TemporaryLocomotiveCounter) & " programmed with lead locomotive of " & TemporaryLocomotiveAddress
            Let TemporaryLocomotiveCounter = TemporaryLocomotiveCounter + 1
        Next TemporaryLocomotiveMemoryOutside
    End If
    
    Let LabelStatus.Caption = "Status: Idle"
    Call ObjectsEnable

End Sub

Private Sub ButtonLeadLocomotiveRestore_Click()
    
    Call ObjectsDisable

    Dim TemporaryInput As String
    Dim TemporaryLimitOne As Long
    Dim TemporaryLimitTwo As Long
    Dim TemporaryLimitThree As Long
    Dim TemporaryLimitFour As Long
    Dim TemporaryDecimalValue As Integer
    Dim TemporaryLocomotiveMemoryInside As Long
    Dim TemporaryLocomotiveMemoryOutside As Long
    
    Let LabelStatus.Caption = "Status: Checking for EPROM version."
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command to DCC system
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainScreen!MSComm1.InBufferCount = 0
    MainScreen!MSComm1.Output = Chr$(&HAA)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While MainScreen!MSComm1.InBufferCount < 3
        Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
        DoEvents
    Loop
    Let TemporaryInput = MainScreen!MSComm1.Input
    If Len(TemporaryInput) <> 3 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Lead Locomotive Restore, wrong number of character returned by command unit after sending &HAA command."
        End If
    End If
    Let TemporaryVersion = Asc(Left$(TemporaryInput, 1)) & "." & Asc(Mid$(TemporaryInput, 2, 1)) & "." & Asc(Right$(TemporaryInput, 1))
    Let LabelStatus.Caption = "Status: EPROM version is " & TemporaryVersion
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Values for EPROM version
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryVersion = "4.0.1" Then
        OptionOlderEPROM.Value = True
        Let TemporaryLimitOne = 61952
        Let TemporaryLimitTwo = 62207
    ElseIf TemporaryVersion = "6.0.0" Then
        OptionNewerEPROM.Value = True
        Let TemporaryLimitOne = 62720
        Let TemporaryLimitTwo = 62975
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Lead Locomotive Restore, unknown version number of EPROM or command unit."
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Reading
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Utilities for Command Control Screen"
    For TemporaryLocomotiveMemoryOutside = TemporaryLimitOne To TemporaryLimitTwo Step 2
    Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryLocomotiveMemoryOutside) & " to 0x" & Hex(TemporaryLocomotiveMemoryOutside + 1)
        Let TemporaryString = Ini.Value
        For TemporaryLocomotiveMemoryInside = TemporaryLocomotiveMemoryOutside To TemporaryLocomotiveMemoryOutside + 1
            Let a = Int(TemporaryLocomotiveMemoryInside / 256)
            Let b = TemporaryLocomotiveMemoryInside - (a * 256)
            Debug.Print Mid$(TemporaryString, (TemporaryLocomotiveMemoryInside - TemporaryLocomotiveMemoryOutside) * 3 + 1, 2)
            Call Hexidecimal(Mid$(TemporaryString, (TemporaryLocomotiveMemoryInside - TemporaryLocomotiveMemoryOutside) * 3 + 1, 2), TemporaryDecimalValue)
            Let MainScreen!MSComm1.InBufferCount = 0
            Let MainScreen!MSComm1.Output = Chr$(&H97) & Chr$(a) & Chr$(b) & Chr$(TemporaryDecimalValue)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Do While MainScreen!MSComm1.InBufferCount < 1
                DoEvents
            Loop
            Let TemporaryInput = MainScreen!MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Clock Screen, Button Stop Clock, wrong number of bytes returned from digital command control unitafter sending &H97 command."
                End If
            End If
            LabelStatus.Caption = "Status: Writing memory location " & Hex(TemporaryLocomotiveMemoryInside)
        Next TemporaryLocomotiveMemoryInside
    Next TemporaryLocomotiveMemoryOutside

    Let LabelStatus.Caption = "Status: Idle"
    Call ObjectsEnable

End Sub

Private Sub ButtonMacroBackup_Click()

    Call ObjectsDisable
    Dim TemporaryLimitOne As Long
    Dim TemporaryLimitTwo As Long
    Dim TemporaryLimitThree As Long
    Dim TemporaryLimitFour As Long
    Dim TemporaryMacroCounter As Integer
    Dim TemporaryString As String
    Dim TemporaryVersion As String
    Dim TemporaryMacroMemoryInside As Long
    Dim TemporaryMacroMemoryOutside As Long
    
    Let LabelStatus.Caption = "Status: Checking for EPROM version."
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command to DCC system
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainScreen!MSComm1.InBufferCount = 0
    MainScreen!MSComm1.Output = Chr$(&HAA)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While MainScreen!MSComm1.InBufferCount < 3
        Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
        DoEvents
    Loop
    Let TemporaryInput = MainScreen!MSComm1.Input
    If Len(TemporaryInput) <> 3 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Macro Backup, wrong number of character returned by command control unit after sending &HAA command."
        End If
    End If
    Let TemporaryVersion = Asc(Left$(TemporaryInput, 1)) & "." & Asc(Mid$(TemporaryInput, 2, 1)) & "." & Asc(Right$(TemporaryInput, 1))
    Let LabelStatus.Caption = "Status: EPROM version is " & TemporaryVersion
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Values for EPROM version
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryVersion = "4.0.1" Then
        OptionOlderEPROM.Value = True
        Let TemporaryLimitOne = 51200
        Let TemporaryLimitTwo = 53199
        Let TemporaryLimitThree = 62208
        Let TemporaryLimitFour = 64207
    ElseIf TemporaryVersion = "6.0.0" Then
        OptionNewerEPROM.Value = True
        Let TemporaryLimitOne = 51200
        Let TemporaryLimitTwo = 53199
        Let TemporaryLimitThree = 53200
        Let TemporaryLimitFour = 55199
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Macro Backup, unknown version number of EPROM or command unit."
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Reading
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Utilities for Command Control Screen"
    For TemporaryMacroMemoryOutside = TemporaryLimitOne To TemporaryLimitTwo Step 20
        Let TemporaryString = ""
        For TemporaryMacroMemoryInside = TemporaryMacroMemoryOutside To TemporaryMacroMemoryOutside + 19
            Let MainScreen!MSComm1.InBufferCount = 0
            Let a = Int(TemporaryMacroMemoryInside / 256)
            Let b = TemporaryMacroMemoryInside - (a * 256)
            Let MainScreen!MSComm1.Output = Chr$(&H9D) & Chr$(a) & Chr$(b)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Do While MainScreen!MSComm1.InBufferCount < 1
                DoEvents
            Loop
            Let TemporaryInput = MainScreen!MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Utility Screen, Button Macro Backup, wrong number of bytes returned from digital command control unit after sending &H9D command."
                End If
            End If
            LabelStatus.Caption = "Status: Reading memory location " & Hex(TemporaryMacroMemoryInside)
            DoEvents
            If Hex(Asc(TemporaryInput)) = "0" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "1" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "2" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "3" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "4" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "5" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "6" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "7" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "8" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "9" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "A" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "B" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "C" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "D" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "E" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "F" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            Else
                Let TemporaryString = TemporaryString & Hex(Asc(TemporaryInput)) & " "
            End If
        Next TemporaryMacroMemoryInside
        Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryMacroMemoryOutside) & " to 0x" & Hex(TemporaryMacroMemoryOutside + 19)
        Let Ini.Value = TemporaryString & ";Macro number" & Str$(TemporaryMacroCounter)
        Let TemporaryMacroCounter = TemporaryMacroCounter + 1
    Next TemporaryMacroMemoryOutside
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Second Memory Location
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryMacroMemoryOutside = TemporaryLimitThree To TemporaryLimitFour Step 20
        Let TemporaryString = ""
        For TemporaryMacroMemoryInside = TemporaryMacroMemoryOutside To TemporaryMacroMemoryOutside + 19
            Let MainScreen!MSComm1.InBufferCount = 0
            Let a = Int(TemporaryMacroMemoryInside / 256)
            Let b = TemporaryMacroMemoryInside - (a * 256)
            Let MainScreen!MSComm1.Output = Chr$(&H9D) & Chr$(a) & Chr$(b)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Do While MainScreen!MSComm1.InBufferCount < 1
                DoEvents
            Loop
            Let TemporaryInput = MainScreen!MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Utility Screen, Button Macro Backup, wrong number of bytes returned from digital command control unit after sending &H9D command."
                End If
            End If
            LabelStatus.Caption = "Status: Reading memory location " & Hex(TemporaryMacroMemoryInside)
            DoEvents
            If Hex(Asc(TemporaryInput)) = "0" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "1" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "2" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "3" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "4" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "5" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "6" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "7" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "8" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "9" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "A" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "B" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "C" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "D" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "E" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            ElseIf Hex(Asc(TemporaryInput)) = "F" Then
                Let TemporaryString = TemporaryString & "0" & Hex(Asc(TemporaryInput)) & " "
            Else
                Let TemporaryString = TemporaryString & Hex(Asc(TemporaryInput)) & " "
            End If
        Next TemporaryMacroMemoryInside
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Utilities for Command Control Screen"
        Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryMacroMemoryOutside) & " to 0x" & Hex(TemporaryMacroMemoryOutside + 19)
        Let Ini.Value = TemporaryString & ";Macro number" & Str$(TemporaryMacroCounter)
        Let TemporaryMacroCounter = TemporaryMacroCounter + 1
    Next TemporaryMacroMemoryOutside
    
    Let LabelStatus.Caption = "Status: Idle"
    Call ObjectsEnable

End Sub

Private Sub ButtonMacroRestore_Click()
    
    Call ObjectsDisable

    Dim TemporaryInput As String
    Dim TemporaryLimitOne As Long
    Dim TemporaryLimitTwo As Long
    Dim TemporaryLimitThree As Long
    Dim TemporaryLimitFour As Long
    Dim TemporaryDecimalValue As Integer
    Dim TemporaryMacroMemoryInside As Long
    Dim TemporaryMacroMemoryOutside As Long
    
    Let LabelStatus.Caption = "Status: Checking for EPROM version."
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Send Command to DCC system
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainScreen!MSComm1.InBufferCount = 0
    MainScreen!MSComm1.Output = Chr$(&HAA)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While MainScreen!MSComm1.InBufferCount < 3
        Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
        DoEvents
    Loop
    Let TemporaryInput = MainScreen!MSComm1.Input
    If Len(TemporaryInput) <> 3 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Macro Restore, wrong number of character returned by command unit after sending &HAA command."
        End If
    End If
    Let TemporaryVersion = Asc(Left$(TemporaryInput, 1)) & "." & Asc(Mid$(TemporaryInput, 2, 1)) & "." & Asc(Right$(TemporaryInput, 1))
    Let LabelStatus.Caption = "Status: EPROM version is " & TemporaryVersion
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Values for EPROM version
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryVersion = "4.0.1" Then
        OptionOlderEPROM.Value = True
        Let TemporaryLimitOne = 51200
        Let TemporaryLimitTwo = 53199
        Let TemporaryLimitThree = 62208
        Let TemporaryLimitFour = 64207
    ElseIf TemporaryVersion = "6.0.0" Then
        OptionNewerEPROM.Value = True
        Let TemporaryLimitOne = 51200
        Let TemporaryLimitTwo = 53199
        Let TemporaryLimitThree = 53200
        Let TemporaryLimitFour = 55199
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Utility Screen, Button Macro Restore, unknown version number of EPROM or command unit."
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Reading
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Utilities for Command Control Screen"
    For TemporaryMacroMemoryOutside = TemporaryLimitOne To TemporaryLimitTwo Step 20
        Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryMacroMemoryOutside) & " to 0x" & Hex(TemporaryMacroMemoryOutside + 19)
        Let TemporaryString = Ini.Value
        For TemporaryMacroMemoryInside = TemporaryMacroMemoryOutside To TemporaryMacroMemoryOutside + 19
            Let a = Int(TemporaryMacroMemoryInside / 256)
            Let b = TemporaryMacroMemoryInside - (a * 256)
            Call Hexidecimal(Mid$(TemporaryString, (TemporaryMacroMemoryInside - TemporaryMacroMemoryOutside) * 3 + 1, 2), TemporaryDecimalValue)
            Let MainScreen!MSComm1.InBufferCount = 0
            Let MainScreen!MSComm1.Output = Chr$(&H97) & Chr$(a) & Chr$(b) & Chr$(TemporaryDecimalValue)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Do While MainScreen!MSComm1.InBufferCount < 1
                DoEvents
            Loop
            Let TemporaryInput = MainScreen!MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Utility Screen, Button Macro Restore, wrong number of bytes returned from digital command control unit after sending &H97 command."
                End If
            End If
            LabelStatus.Caption = "Status: Writing memory location " & Hex(TemporaryMacroMemoryInside)
        Next TemporaryMacroMemoryInside
    Next TemporaryMacroMemoryOutside
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start Reading
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Utilities for Command Control Screen"
    For TemporaryMacroMemoryOutside = TemporaryLimitThree To TemporaryLimitFour Step 20
        Let Ini.Parameter = "Memory location from 0x" & Hex(TemporaryMacroMemoryOutside) & " to 0x" & Hex(TemporaryMacroMemoryOutside + 19)
        Let TemporaryString = Ini.Value
        For TemporaryMacroMemoryInside = TemporaryMacroMemoryOutside To TemporaryMacroMemoryOutside + 19
            Let a = Int(TemporaryMacroMemoryInside / 256)
            Let b = TemporaryMacroMemoryInside - (a * 256)
            Call Hexidecimal(Mid$(TemporaryString, (TemporaryMacroMemoryInside - TemporaryMacroMemoryOutside) * 3 + 1, 2), TemporaryDecimalValue)
            Let MainScreen!MSComm1.InBufferCount = 0
            Let MainScreen!MSComm1.Output = Chr$(&H97) & Chr$(a) & Chr$(b) & Chr$(TemporaryDecimalValue)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Do While MainScreen!MSComm1.InBufferCount < 1
                DoEvents
            Loop
            Let TemporaryInput = MainScreen!MSComm1.Input
            If Len(TemporaryInput) <> 1 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Utility Screen, Button Macro Restore, wrong number of bytes returned from digital command control unit after sending &H97 command."
                End If
            End If
            LabelStatus.Caption = "Status: Writing memory location " & Hex(TemporaryMacroMemoryInside)
        Next TemporaryMacroMemoryInside
    Next TemporaryMacroMemoryOutside
    
    Let LabelStatus.Caption = "Status: Idle"
    Call ObjectsEnable

End Sub

Private Sub ButtonPrint_Click()

    UtilitiesForCommandControl.PrintForm

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
        If TemporaryScreen = "Utilities for Command Control Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Utilities for Command Control Screen"
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
            Let Ini.Value = "Utilities for Command Control Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Utilities for Command Control Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. "
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
            Let Ini.Value = "Utilities for Command Control Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Utilities for Command Control Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(UtilitiesForCommandControl.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(UtilitiesForCommandControl.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(UtilitiesForCommandControl.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(UtilitiesForCommandControl.Height)

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
                Let Ini.Value = "Utilities for Command Control Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Utilities for Command Control Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    UtilitiesForCommandControl.Hide

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
    Let Ini.Application = "Utilities for Command Control Screen"
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
        UtilitiesForCommandControl.Left = (Screen.Width - Width) / 2
        UtilitiesForCommandControl.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + UtilitiesForCommandControl.Width > Screen.Width Then
            Let UtilitiesForCommandControl.Left = Screen.Width - UtilitiesForCommandControl.Width
        Else
            Let UtilitiesForCommandControl.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + UtilitiesForCommandControl.Height > Screen.Height Then
            Let UtilitiesForCommandControl.Top = Screen.Height - UtilitiesForCommandControl.Height
        Else
            Let UtilitiesForCommandControl.Top = Val(TemporaryValueTop)
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
        
        Let TemporaryText1 = "This highlighted text when 'click'ed will" & vbCrLf & "display my home page in your web browser."
        Let TemporaryText2 = "Universal Resource Link"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(HyperLabelOcx1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(HyperLabelOcx1, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This highlighted text when 'click'ed will" & vbCrLf & "display my email address in your email program."
        Let TemporaryText2 = "Universal Resource Link"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(HyperLabelOcx2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(HyperLabelOcx2, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This highlighted text when 'click'ed will" & vbCrLf & "display my yahoo group in your web browser."
        Let TemporaryText2 = "Universal Resource Link"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(HyperLabelOcx3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(HyperLabelOcx3, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "close the Utilities for Command Control Window and return control" & vbCrLf & "to the main screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    Else
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
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



Private Sub ObjectsDisable()

    ButtonMacroBackup.Enabled = False
    ButtonMacroRestore.Enabled = False
    ButtonPrint.Enabled = False
    ButtonClose.Enabled = False

End Sub

Private Sub ObjectsEnable()

    ButtonMacroBackup.Enabled = True
    ButtonMacroRestore.Enabled = True
    ButtonPrint.Enabled = True
    ButtonClose.Enabled = True

End Sub

Private Sub Hexidecimal(TemporaryHexidecimalValue As String, TemporaryDecimalValue As Integer)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert to Decimal
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Left$(TemporaryHexidecimalValue, 1) = "0" Then
        Let TemporaryDecimalValue = 0
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "1" Then
        Let TemporaryDecimalValue = 16
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "2" Then
        Let TemporaryDecimalValue = 32
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "3" Then
        Let TemporaryDecimalValue = 48
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "4" Then
        Let TemporaryDecimalValue = 64
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "5" Then
        Let TemporaryDecimalValue = 80
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "6" Then
        Let TemporaryDecimalValue = 96
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "7" Then
        Let TemporaryDecimalValue = 112
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "8" Then
        Let TemporaryDecimalValue = 128
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "9" Then
        Let TemporaryDecimalValue = 144
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "A" Then
        Let TemporaryDecimalValue = 160
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "B" Then
        Let TemporaryDecimalValue = 176
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "C" Then
        Let TemporaryDecimalValue = 192
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "D" Then
        Let TemporaryDecimalValue = 208
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "E" Then
        Let TemporaryDecimalValue = 224
    ElseIf Left$(TemporaryHexidecimalValue, 1) = "F" Then
        Let TemporaryDecimalValue = 240
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Clock Screen, Button Stop Clock, wrong number of bytes returned from digital command control unit."
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert to Decimal
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Right$(TemporaryHexidecimalValue, 1) = "0" Then
        'Let TemporaryDecimalValue = TemporaryDecimalValue
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "1" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 1
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "2" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 2
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "3" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 3
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "4" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 4
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "5" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 5
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "6" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 6
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "7" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 7
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "8" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 8
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "9" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 9
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "A" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 10
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "B" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 11
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "C" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 12
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "D" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 13
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "E" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 14
    ElseIf Right$(TemporaryHexidecimalValue, 1) = "F" Then
        Let TemporaryDecimalValue = TemporaryDecimalValue + 15
    ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "Clock Screen, Button Stop Clock, wrong number of bytes returned from digital command control unit."
    End If

End Sub

Private Sub Form_Resize()

    If UtilitiesForCommandControl.WindowState = vbMinimized Then
    
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
        
    ElseIf UtilitiesForCommandControl.WindowState = vbNormal Then
    
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


