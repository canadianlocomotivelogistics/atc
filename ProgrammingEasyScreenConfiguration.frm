VERSION 4.00
Begin VB.Form ProgrammingEasyScreenConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Programming Mode - Easy Screen - Configuration CV29"
   ClientHeight    =   4290
   ClientLeft      =   6810
   ClientTop       =   4545
   ClientWidth     =   7065
   Height          =   4695
   Icon            =   "ProgrammingEasyScreenConfiguration.frx":0000
   Left            =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7065
   Top             =   4200
   Width           =   7185
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Caption         =   "6"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCV29 
      Caption         =   "8"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   29
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7560
      Top             =   240
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7560
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7560
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
   Begin VB.Label LabelConfigurationBit8 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   6975
   End
   Begin VB.Label LabelConfigurationBit7 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   6975
   End
   Begin VB.Label LabelConfigurationBit6 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label LabelConfigurationBit5 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   6975
   End
   Begin VB.Label LabelConfigurationBit4 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label LabelConfigurationBit3 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label LabelNotes 
      Caption         =   "Notes:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label LabelConfigurationBit2 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Label LabelConfigurationBit1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label5 
      Caption         =   $"ProgrammingEasyScreenConfiguration.frx":0442
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "ProgrammingEasyScreenConfiguration.frx":04CF
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   6960
      X2              =   120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "CV29  - Configuration"
      Height          =   195
      Left            =   345
      TabIndex        =   12
      Top             =   1080
      Width           =   1500
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
      TabIndex        =   11
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CV29"
      Height          =   195
      Left            =   5160
      TabIndex        =   9
      Top             =   1080
      Width           =   390
   End
End
Attribute VB_Name = "ProgrammingEasyScreenConfiguration"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub ButtonClose_Click()

Let ProgrammingDecoder!LocomotiveDecoderCVd(29) = textboxcvvalue(29).Text
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
            If TemporaryScreen = "Programming Easy Screen Configuration Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Programming Easy Screen Configuration Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Programming Easy Screen Configuration Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Programming Easy Screen Configuration Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub CheckBoxCV29_Click(Index As Integer)

If CheckBoxCV29(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 1
    If Index = 2 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 2
    If Index = 3 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 4
    If Index = 4 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 8
    If Index = 5 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 16
    If Index = 6 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 32
    If Index = 7 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 64
    If Index = 8 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) - 128
End If

If CheckBoxCV29(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 1
    If Index = 2 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 2
    If Index = 3 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 4
    If Index = 4 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 8
    If Index = 5 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 16
    If Index = 6 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 32
    If Index = 7 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 64
    If Index = 8 Then textboxcvvalue(29).Text = Val(textboxcvvalue(29).Text) + 128
End If

Call UpdateConfigurationLabel


End Sub

Private Sub Form_Activate()

    DoEvents
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Add to Screen Stack
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
        If TemporaryScreen = "Programming Easy Screen Configuration Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Programming Easy Screen Configuration Screen"
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
            Let Ini.Value = "Programming Easy Screen Configuration Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Programming Easy Screen Configuration Screen , Form Activate, variable error in ATC.INI file for 'Trnsparency' setting."
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
            Let Ini.Value = "Programming Easy Screen Configuration Screen , Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Programming Easy Screen Configuration Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenConfiguration.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenConfiguration.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenConfiguration.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenConfiguration.Height)

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
                Let Ini.Value = "Programming Easy Screen Configuration Screen , Form Deactivate, variable error in ATC.INI file for 'Trnsparency' setting."
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
            Let Ini.Value = "Programming Easy Screen Configuration Screen , Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ProgrammingEasyScreenConfiguration.Hide
    'unload ProgrammingEasyScreenConfiguration

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
    Let Ini.Application = "Programming Easy Screen Configuration Screen"
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
        ProgrammingEasyScreenConfiguration.Left = (Screen.Width - Width) / 2
        ProgrammingEasyScreenConfiguration.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + ProgrammingEasyScreenConfiguration.Width > Screen.Width Then
            Let ProgrammingEasyScreenConfiguration.Left = Screen.Width - ProgrammingEasyScreenConfiguration.Width
        Else
            Let ProgrammingEasyScreenConfiguration.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + ProgrammingEasyScreenConfiguration.Height > Screen.Height Then
            Let ProgrammingEasyScreenConfiguration.Top = Screen.Height - ProgrammingEasyScreenConfiguration.Height
        Else
            Let ProgrammingEasyScreenConfiguration.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit eight of configuration variable twenty-nine is" & vbCrLf & "on or off. Bit eight indicates if the decoder" & vbCrLf & "is a multifuncton or an acessory decoder."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Eight"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit seven of configuration variable twenty-nine" & vbCrLf & "is on or off. Bit seven does not indicate any option" & vbCrLf & "in the decoder. It is not used."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Seven"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit six of configuratin variable twenty-nine " & vbCrLf & "is on or off. Bit six indicates if the extended" & vbCrLf & "addresses, configuration variable seventeen" & vbCrLf & "and eighteen (four digit) are used; otherwise configuration" & vbCrLf & "variable one is used (two digits)."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Six"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit five of configuration variable twenty-nine" & vbCrLf & "is on or off. Bit five indicates if the decoder" & vbCrLf & "uses configuraton variables two, five and" & vbCrLf & "six for spped control; otherwise configuartion variable sixty-six" & vbCrLf & "through ninety-five (speed table) are used."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Five"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit four of configuration variable, twenty-nine" & vbCrLf & "is on or off. Bit four indicates if the decoder" & vbCrLf & "is non-advanced; otherwise it is an advanced" & vbCrLf & "decoder."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit three of configuration variable twenty-nine" & vbCrLf & "is on or off. Bit three indicates if the decoder" & vbCrLf & "responds to digital signals only; otherwise direct current too."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit two of configuration variable twenty-nine" & vbCrLf & "is on or off. Bit two indicates if the decoder" & vbCrLf & "responds to fourteen speed steps commands; otherwise" & vbCrLf & "twenty-eight and one hundred and twenty-eight speed" & vbCrLf & "step commands."
        Let TemporaryText2 = "Configuration Variable 29 - Bit Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox allows you to indicate whether" & vbCrLf & "bit one of configuratin variable one" & vbCrLf & "is on or off. Bit one indicates if the decoder is" & vbCrLf & "operating in normal; otherwise the decoder is in a" & vbCrLf & "reversed direction."
        Let TemporaryText2 = "Configuration Variable 29 - Bit One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCV29(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCV29(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")


        Let TemporaryText1 = "This textbox displays the value of configuraton" & vbCrLf & "variable twenty-nine. This is the total, in" & vbCrLf & "decimal notation, for each bit turned on."
        Let TemporaryText2 = "Configuration Variable 29"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue(29))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue(29), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button closes the window (Programming Easy" & vbCrLf & "Screen Configuration)and returns you to " & vbCrLf & "Programming Decoder Screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

Let temp = Val(ProgrammingDecoder!LocomotiveDecoderCVd(29).Text)

For Y = 8 To 1 Step -1

If Y = 8 Then bitweight = 128
If Y = 7 Then bitweight = 64
If Y = 6 Then bitweight = 32
If Y = 5 Then bitweight = 16
If Y = 4 Then bitweight = 8
If Y = 3 Then bitweight = 4
If Y = 2 Then bitweight = 2
If Y = 1 Then bitweight = 1


If Val(temp) / bitweight >= 1 Then
    Let temp = temp - bitweight
    Let CheckBoxCV29(Y).Value = vbChecked
Else
    Let CheckBoxCV29(Y).Value = vbUnchecked
End If

Next Y

Call UpdateConfigurationLabel

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





Private Sub UpdateConfigurationLabel()

If CheckBoxCV29(1).Value = vbUnchecked Then
    Let LabelConfigurationBit1.Caption = "Decoder is configured for normal direction, via bit one."
Else
    Let LabelConfigurationBit1.Caption = "Decoder is configured for reverse direction, via bit one."
End If

If CheckBoxCV29(2).Value = vbUnchecked Then
    Let LabelConfigurationBit2.Caption = "Decoder is configured for fourteen step mode, via bit two."
Else
    Let LabelConfigurationBit2.Caption = "Decoder is configured for twenty-eight step mode, via bit two."
End If

If CheckBoxCV29(3).Value = vbUnchecked Then
    Let LabelConfigurationBit3.Caption = "Decoder is configured for digital source only, via bit three."
Else
    Let LabelConfigurationBit3.Caption = "Decoder is configured for alternative source, CV12, via bit three."
End If

If CheckBoxCV29(4).Value = vbUnchecked Then
    Let LabelConfigurationBit4.Caption = "Decoder is configured for non-advanced acknowledgement, via bit four."
Else
    Let LabelConfigurationBit4.Caption = "Decoder is configured for advanced acknowledgement, via bit four."
End If

If CheckBoxCV29(5).Value = vbUnchecked Then
    Let LabelConfigurationBit5.Caption = "Decoder is configured for speed set by CV2, CV5 and CV6, via bit five."
Else
    Let LabelConfigurationBit5.Caption = "Decoder is configured for speed set by table CV66 though CV95, via bit five."
End If

If CheckBoxCV29(6).Value = vbUnchecked Then
    Let LabelConfigurationBit6.Caption = "Decoder is configured for use of primary address CV1, via bit six."
Else
    Let LabelConfigurationBit6.Caption = "Decoder is configured for use of extended address CV17 and CV18, via bit six."
End If

If CheckBoxCV29(7).Value = vbUnchecked Then
    Let LabelConfigurationBit7.Caption = "Decoder is configured for reserved feature, via bit seven."
Else
    Let LabelConfigurationBit7.Caption = "Decoder is configured for reserved feature, via bit seven."
End If

If CheckBoxCV29(8).Value = vbUnchecked Then
    Let LabelConfigurationBit8.Caption = "Decoder is configured for multifunction decoder, via bit eight."
Else
    Let LabelConfigurationBit8.Caption = "Decoder is configured for accessory decoder, via bit eight."
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    If ProgrammingEasyScreenConfiguration.WindowState = vbMinimized Then
    
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
        
    ElseIf ProgrammingEasyScreenConfiguration.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Programming Easy Screen Configuration Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(ProgrammingEasyScreenConfiguration.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(ProgrammingEasyScreenConfiguration.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(ProgrammingEasyScreenConfiguration.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(ProgrammingEasyScreenConfiguration.Height)
 
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


