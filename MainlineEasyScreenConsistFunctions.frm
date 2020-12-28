VERSION 4.00
Begin VB.Form MainlineEasyScreenConsistFunctions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Easy Screen - Consist Functions"
   ClientHeight    =   2715
   ClientLeft      =   6330
   ClientTop       =   5595
   ClientWidth     =   7080
   Height          =   3120
   Icon            =   "MainlineEasyScreenConsistFunctions.frx":0000
   Left            =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7080
   Top             =   5250
   Width           =   7200
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4440
      TabIndex        =   28
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxCV22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   20
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   19
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV22 
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   13
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV21 
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   22
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   21
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7440
      Top             =   1680
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7440
      Top             =   1200
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7440
      Top             =   600
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label LabelConsistHeadlightControl 
      Caption         =   "No active headlight control is available in consist mode."
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label LabelConsistFunctions 
      Caption         =   "No active functions are available in consist mode."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label5 
      Caption         =   $"MainlineEasyScreenConsistFunctions.frx":0442
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MainlineEasyScreenConsistFunctions.frx":04E8
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   6960
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Consist Headlight Control"
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Consist Control for "
      Height          =   195
      Left            =   525
      TabIndex        =   14
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Function Number      8      7      6      5      4      3      2      1"
      Height          =   195
      Left            =   645
      TabIndex        =   6
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "CV22"
      Height          =   195
      Left            =   5160
      TabIndex        =   3
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CV21"
      Height          =   195
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   390
   End
End
Attribute VB_Name = "MainlineEasyScreenConsistFunctions"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButtonClose_Click()

Let MainlineDecoder!LocomotiveDecoderCVd(21) = textboxcvvalue(21).Text
Let MainlineDecoder!LocomotiveDecoderCVd(22) = textboxcvvalue(22).Text
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
            If TemporaryScreen = "Mainline Easy Screen Consist Functions Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonPrint_Click()

    MainlineEasyScreenConsistFunctions.PrintForm

End Sub

Private Sub CheckBoxCV21_Click(Index As Integer)

If CheckBoxCV21(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 1
    If Index = 2 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 2
    If Index = 3 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 4
    If Index = 4 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 8
    If Index = 5 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 16
    If Index = 6 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 32
    If Index = 7 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 64
    If Index = 8 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) - 128
End If

If CheckBoxCV21(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 1
    If Index = 2 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 2
    If Index = 3 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 4
    If Index = 4 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 8
    If Index = 5 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 16
    If Index = 6 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 32
    If Index = 7 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 64
    If Index = 8 Then textboxcvvalue(21).Text = Val(textboxcvvalue(21).Text) + 128
End If

Call UpdateConsistFunctionLabel

End Sub


Private Sub CheckBoxCV22_Click(Index As Integer)

If CheckBoxCV22(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 1
    If Index = 2 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 2
    If Index = 3 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 4
    If Index = 4 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 8
    If Index = 5 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 16
    If Index = 6 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 32
    If Index = 7 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 64
    If Index = 8 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) - 128
End If

If CheckBoxCV22(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 1
    If Index = 2 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 2
    If Index = 3 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 4
    If Index = 4 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 8
    If Index = 5 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 16
    If Index = 6 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 32
    If Index = 7 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 64
    If Index = 8 Then textboxcvvalue(22).Text = Val(textboxcvvalue(22).Text) + 128
End If

Call UpdateConsistHeadlightControl

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
        If TemporaryScreen = "Mainline Easy Screen Consist Functions Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Easy Screen Consist Functions Screen"
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
            Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Easy Screen Consist Functions Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineEasyScreenConsistFunctions.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineEasyScreenConsistFunctions.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineEasyScreenConsistFunctions.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineEasyScreenConsistFunctions.Height)

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
                Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Easy Screen Consist Functions Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineEasyScreenConsistFunctions.Hide
    'unload Mainlineeasyscreenconsistfunctions

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
    Let Ini.Application = "Mainline Easy Screen Consist Function Screen"
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
        MainlineEasyScreenConsistFunctions.Left = (Screen.Width - Width) / 2
        MainlineEasyScreenConsistFunctions.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineEasyScreenConsistFunctions.Width > Screen.Width Then
            Let MainlineEasyScreenConsistFunctions.Left = Screen.Width - MainlineEasyScreenConsistFunctions.Width
        Else
            Let MainlineEasyScreenConsistFunctions.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineEasyScreenConsistFunctions.Height > Screen.Height Then
            Let MainlineEasyScreenConsistFunctions.Top = Screen.Height - MainlineEasyScreenConsistFunctions.Height
        Else
            Let MainlineEasyScreenConsistFunctions.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This text box displays the value of configuration" & vbCrLf & "variable twenty-one, the function(s) which can be controlled" & vbCrLf & "while the decoder is consisted with another decoder." & vbCrLf & "When the 'Easy Screen for Consist Functions' is 'Closed' the" & vbCrLf & "value of the configuration varianle will be copied to the 'Mainline Programming" & vbCrLf & "for Decoder' screen."
        Let TemporaryText2 = "Controlling Functions in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(21))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(21), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
       
        Let TemporaryText1 = "This text box displays the value of configuration" & vbCrLf & "variable twenty-two, headlights which can be controlled" & vbCrLf & "while the decoder is consisted with another decoder." & vbCrLf & "When the 'Easy Screen for Consist Functions' is 'Closed' the" & vbCrLf & "value of the configuration varianle will be copied to the 'Mainline Programming" & vbCrLf & "for Decoder' screen."
        Let TemporaryText2 = "Controlling Headights in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(22))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(22), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "one while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function One in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "two while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Two in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "three while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Three in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "four while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Four in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "five while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Five in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "six while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Six in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "seven while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Seven in a Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "eight while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Function Eight in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV21(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV21(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
      
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "zero while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates function" & vbCrLf & "zero while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates a function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box activates or de-activates a function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates a function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling a Headight Five in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates a function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight Six in a Consist."
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or de-activates a function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling a Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box activates or a deactivates function" & vbCrLf & "while the decoder is consisted with another" & vbCrLf & "decoder."
        Let TemporaryText2 = "Controlling Headlight(s) in a Consist"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV22(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV22(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button closes the current window and returns" & vbCrLf & "you to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update CheckBoxes
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim TemporaryVariableT As Integer
    Dim TemporaryVariableY As Integer
    Dim TemporaryVariableZ As Integer
    Dim TemporaryVariableBitWeight As Integer

    For TemporaryVariableT = 21 To 22
    
    Let TemporaryVariableZ = Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text)
        
        For TemporaryVariableY = 8 To 1 Step -1
        
        If TemporaryVariableY = 8 Then TemporaryVariableBitWeight = 128
        If TemporaryVariableY = 7 Then TemporaryVariableBitWeight = 64
        If TemporaryVariableY = 6 Then TemporaryVariableBitWeight = 32
        If TemporaryVariableY = 5 Then TemporaryVariableBitWeight = 16
        If TemporaryVariableY = 4 Then TemporaryVariableBitWeight = 8
        If TemporaryVariableY = 3 Then TemporaryVariableBitWeight = 4
        If TemporaryVariableY = 2 Then TemporaryVariableBitWeight = 2
        If TemporaryVariableY = 1 Then TemporaryVariableBitWeight = 1
        
        
        If Val(TemporaryVariableZ) / TemporaryVariableBitWeight >= 1 Then
            Let TemporaryVariableZ = TemporaryVariableZ - TemporaryVariableBitWeight
            If TemporaryVariableT = 21 Then Let CheckBoxCV21(TemporaryVariableY).Value = vbChecked
            If TemporaryVariableT = 22 Then Let CheckBoxCV22(TemporaryVariableY).Value = vbChecked
        Else
            If TemporaryVariableT = 21 Then Let CheckBoxCV21(TemporaryVariableY).Value = vbUnchecked
            If TemporaryVariableT = 22 Then Let CheckBoxCV22(TemporaryVariableY).Value = vbUnchecked
         End If
        
        Next TemporaryVariableY
        
    Next TemporaryVariableT


    Call UpdateConsistFunctionLabel
    Call UpdateConsistHeadlightControl

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
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub



Private Sub UpdateConsistFunctionLabel()

Let Temporary$ = ""
Let LabelConsistFunctions.Caption = "No active functions are available in consist mode."
For j = 1 To 8
    If CheckBoxCV21(j).Value = vbChecked Then
        Let LabelConsistFunctions.Caption = "In consist mode active functions available are "
        Let Temporary$ = Temporary$ + Str$(j) + " "
    End If
Next j
Let LabelConsistFunctions.Caption = LabelConsistFunctions.Caption + Temporary$

End Sub

Private Sub UpdateConsistHeadlightControl()

Let Temporary$ = ""
Let LabelConsistHeadlightControl.Caption = "No active headlight control is available in consist mode."

    If CheckBoxCV22(1).Value = vbChecked Then
        Let LabelConsistHeadlightControl.Caption = "In consist mode active headlight control is "
        Let Temporary$ = Temporary$ + "1 (White Wire [Front]) "
    End If
    If CheckBoxCV22(2).Value = vbChecked Then
        Let LabelConsistHeadlightControl.Caption = "In consist mode active headlight control is "
        Let Temporary$ = Temporary$ + "2 (Yellow Wire [Rear])"
    End If

Let LabelConsistHeadlightControl.Caption = LabelConsistHeadlightControl.Caption + Temporary$

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    If MainlineEasyScreenConsistFunctions.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineEasyScreenConsistFunctions.WindowState = vbNormal Then
    
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

Private Sub Form_Unload(Cancel As Integer)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unloading the Form
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Saving the screen size
'

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Easy Screen Consist Function"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineEasyScreenConsistFunctions.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineEasyScreenConsistFunctions.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineEasyScreenConsistFunctions.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineEasyScreenConsistFunctions.Height)
 
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


