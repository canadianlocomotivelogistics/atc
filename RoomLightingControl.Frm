VERSION 4.00
Begin VB.Form RoomLightingControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Room Lighting Control"
   ClientHeight    =   6195
   ClientLeft      =   3675
   ClientTop       =   2685
   ClientWidth     =   6240
   Height          =   6600
   Icon            =   "RoomLightingControl.frx":0000
   Left            =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6240
   Top             =   2340
   Width           =   6360
   Begin VB.TextBox TextBoxFunctionForRoomLightingControl 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4740
      TabIndex        =   31
      Text            =   "2"
      Top             =   3900
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function Controls - Group One"
      Height          =   1575
      Left            =   6480
      TabIndex        =   25
      Top             =   180
      Width           =   2535
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Function Controls - Group Two"
      Height          =   1455
      Left            =   6480
      TabIndex        =   20
      Top             =   1860
      Width           =   2535
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 7"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 8"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Function Controls - Group Three"
      Height          =   1575
      Left            =   6480
      TabIndex        =   15
      Top             =   3420
      Width           =   2535
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 9"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 10"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 11"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox CheckBoxRoomLightingControlWithFunction 
         Caption         =   "Function 12"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox TextBoxDecoderNumber 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4740
      TabIndex        =   14
      Text            =   "3"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton ButtonManualControl 
      Caption         =   "&Lights are Off"
      Height          =   255
      Left            =   4980
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxShortAddressUsed 
      Caption         =   "Decoder being used to control room light uses a short address."
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Value           =   1  'Checked
      Width           =   5355
   End
   Begin VB.CheckBox CheckBoxUseRoomLightingControl 
      Caption         =   "I want to use room lighting control."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4980
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   3660
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "RoomLightingControl.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7500
      Top             =   5100
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Label Label9 
      Caption         =   "What function button will be used to control the room lights?"
      Height          =   195
      Left            =   180
      TabIndex        =   30
      Top             =   3900
      Width           =   4230
   End
   Begin VB.Label Label8 
      Caption         =   "If you want to mannually control the relay or test it out, you can use the button to toggle the room lights."
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   4980
      Width           =   6075
   End
   Begin VB.Label Label7 
      Caption         =   "Manual Control"
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
      Left            =   120
      TabIndex        =   12
      Top             =   4620
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Room Lighting Control"
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
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "Identifiy the Deocder for Room Lighting Control"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label4 
      Caption         =   "What is the decoder number used to control the room lights?"
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   3600
      Width           =   4395
   End
   Begin VB.Label Label3 
      Caption         =   "To control the room lighting, I need to know the number of the accessory decoder."
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   5955
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6120
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label2 
      Caption         =   $"RoomLightingControl.frx":0884
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "If you are hosting a 'intenet train session' you might want the room light controled by the software."
      Height          =   435
      Left            =   720
      TabIndex        =   2
      Top             =   180
      Width           =   5415
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   8520
      Top             =   5040
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   7980
      Top             =   5040
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
Attribute VB_Name = "RoomLightingControl"
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
            If TemporaryScreen = "Room Lighting Control Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Room Lighting Control Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Room Lighting Control Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Room Lighting Control Screen, Button Close, stack empty, underflow."
        End If
    End If

End Sub


Public Sub ButtonManualControl_Click()

    Dim FunctionForRoomLightingControl As Integer
    
    If ButtonManualControl.Caption = "&Lights are Off" Then
        Let ButtonManualControl.Caption = "&Lights are On"
        Let MainScreen!onebyted.Text = Int(Val(RoomLightingControl.TextBoxDecoderNumber.Text) / 256)
        Let MainScreen!twoByteD.Text = Val(RoomLightingControl.TextBoxDecoderNumber.Text) - (Val(MainScreen!onebyted.Text) * 256)
        If CheckboxShortAddressUsed.Value = vbUnchecked Then
             Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64
        Else
            Let MainScreen!onebyted.Text = ""
        End If
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Room Lighting Decoder using decoder number " & RoomLightingControl.TextBoxDecoderNumber.Text & " is turned on." & vbCrLf
    Else
        Let ButtonManualControl.Caption = "&Lights are Off"
        Let MainScreen!onebyted.Text = Int(Val(RoomLightingControl.TextBoxDecoderNumber.Text) / 256)
        Let MainScreen!twoByteD.Text = Val(RoomLightingControl.TextBoxDecoderNumber.Text) - (Val(MainScreen!onebyted.Text) * 256)
        If CheckboxShortAddressUsed.Value = vbUnchecked Then
             Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64
        Else
            Let MainScreen!onebyted.Text = ""
        End If
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Room Lighting Decoder using decoder number " & RoomLightingControl.TextBoxDecoderNumber.Text & " is turned off." & vbCrLf
    End If
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    For FunctionForRoomLightingControl = 0 To 4
        If Val(TextBoxFunctionForRoomLightingControl.Text) = FunctionForRoomLightingControl Then
            If CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = False Then
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbChecked
            Else
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbUnchecked
            End If
        Call SetFunctionForRoomLightingControl01234
        End If
    Next FunctionForRoomLightingControl
    For FunctionForRoomLightingControl = 5 To 8
        If Val(TextBoxFunctionForRoomLightingControl.Text) = FunctionForRoomLightingControl Then
            If CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = False Then
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbChecked
            Else
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbUnchecked
            End If
        Call SetFunctionForRoomLightingControl5678
        End If
    Next FunctionForRoomLightingControl
    For FunctionForRoomLightingControl = 9 To 12
        If Val(TextBoxFunctionForRoomLightingControl.Text) = FunctionForRoomLightingControl Then
            If CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = False Then
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbChecked
            Else
                CheckBoxRoomLightingControlWithFunction(FunctionForRoomLightingControl).Value = vbUnchecked
            End If
        Call SetFunctionForRoomLightingControl01234
        End If
    Next FunctionForRoomLightingControl

    Call MainScreen.SendCommandviaTrackQ

End Sub

Private Sub ButtonPrint_Click()

    RoomLightingControl.PrintForm

End Sub



Private Sub CheckBoxUseRoomLightingControl_Click()

    If CheckBoxUseRoomLightingControl = vbChecked Then
        Let TextBoxDecoderNumber.Enabled = True
        Let CheckboxShortAddressUsed.Enabled = True
        Let ButtonManualControl.Enabled = True
    Else
        Let TextBoxDecoderNumber.Enabled = False
        Let CheckboxShortAddressUsed.Enabled = False
        Let ButtonManualControl.Enabled = False
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
        If TemporaryScreen = "Room Lighting Control Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Room Lighting Control Screen"
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
            Let Ini.Value = "Room Lighting Control Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Room Lighting Control Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. "
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
            Let Ini.Value = "Room Lighting Control Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Getting Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Room Lighting Control Screen"
    
    Let Ini.Parameter = "CheckBoxUseRoomLightingControl"
    Let CheckBoxUseRoomLightingControl.Value = Ini.Value
    Let Ini.Parameter = "TextBoxDecoderNumber"
    Let TextBoxDecoderNumber.Text = Ini.Value
    Let Ini.Parameter = "TextBoxFunctionNumber"
    Let TextBoxFunctionForRoomLightingControl.Text = Ini.Value
    Let Ini.Parameter = "CheckBoxShortAddressUsed"
    Let CheckboxShortAddressUsed.Value = Ini.Value

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Room Lighting Control Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(RoomLightingControl.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(RoomLightingControl.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(RoomLightingControl.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(RoomLightingControl.Height)
    
    Let Ini.Parameter = "CheckBoxUseRoomLightingControl"
    Let Ini.Value = CheckBoxUseRoomLightingControl.Value
    Let Ini.Parameter = "TextBoxDecoderNumber"
    Let Ini.Value = TextBoxDecoderNumber.Text
    Let Ini.Parameter = "TextBoxFunctionNumber"
    Let Ini.Value = TextBoxFunctionForRoomLightingControl.Text
    Let Ini.Parameter = "CheckBoxShortAddressUsed"
    Let Ini.Value = CheckboxShortAddressUsed.Value
    
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
                Let Ini.Value = "Room Lighting Control Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Room Lighting Control Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    RoomLightingControl.Hide

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
    Let Ini.Application = "Room Lighting Control Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "CheckBoxUseRoomLightingControl"
    Let CheckBoxUseRoomLightingControl.Value = Ini.Value
    Let Ini.Parameter = "DecoderNumber"
    Let TextBoxDecoderNumber.Text = Ini.Value
    Let Ini.Parameter = "FunctionNumber"
    Let TextBoxFunctionForRoomLightingControl.Text = Ini.Value
    Let Ini.Parameter = "CheckBoxShortAddressUsed"
    Let CheckboxShortAddressUsed.Value = Ini.Value
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        RoomLightingControl.Left = (Screen.Width - Width) / 2
        RoomLightingControl.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + RoomLightingControl.Width > Screen.Width Then
            Let RoomLightingControl.Left = Screen.Width - RoomLightingControl.Width
        Else
            Let RoomLightingControl.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + RoomLightingControl.Height > Screen.Height Then
            Let RoomLightingControl.Top = Screen.Height - RoomLightingControl.Height
        Else
            Let RoomLightingControl.Top = Val(TemporaryValueTop)
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



Public Sub SetFunctionForRoomLightingControl01234()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function for Lighting Control "
    
    Let TemporaryByte = 128
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(0).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 16
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "0 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(1).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "1 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(2).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "2 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(3).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "3 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(4).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "4 Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""

End Sub

Public Sub SetFunctionForRoomLightingControl5678()

    Dim TemporaryByte As Integer
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function for Lighting Control "
    
    Let TemporaryByte = 128 + 32
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(5).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "5 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(6).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "6 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(7).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "7 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(8).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 On; "
    Else
        Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "8 Off; "
    End If
    Let MainScreen!ThreeByteD.Text = TemporaryByte
    Let MainScreen!FourByteD.Text = ""
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
End Sub

Public Sub SetFunctionForRoomLightingControl9101112()

    Dim TemporaryByte As Integer
    
    Let TemporaryByte = 128 + 32
    
    Let MainScreen!TextBoxCommunicationWindowDCC.Text = MainScreen!TextBoxCommunicationWindowDCC.Text & "Function for Lighting Control "

    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(9).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 1
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "9 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(10).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 2
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "10 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(11).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 4
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "11 Off; "
    End If
    If RoomLightingControl!CheckBoxRoomLightingControlWithFunction(12).Value = vbChecked Then
        Let TemporaryByte = TemporaryByte + 8
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 On; "
    Else
        Let MainScreen.TextBoxCommunicationWindowDCC.Text = MainScreen.TextBoxCommunicationWindowDCC.Text + "12 Off; "
    End If
    
    Let ThreeByteD.Text = TemporaryByte
    Let FourByteD.Text = ""
    Let FiveByteD.Text = ""
    Let SixByteD.Text = ""
    Let sevenbyted.Text = ""
    
End Sub

Private Sub Form_Resize()

    If RoomLightingControl.WindowState = vbMinimized Then
    
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
        
    ElseIf RoomLightingControl.WindowState = vbNormal Then
    
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


Private Sub TextBoxDecoderNumber_Change()

End Sub


