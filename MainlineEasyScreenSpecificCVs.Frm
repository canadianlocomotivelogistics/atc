VERSION 4.00
Begin VB.Form MainlineEasyScreenSpecificCvs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Easy Screen Specific CVs"
   ClientHeight    =   4035
   ClientLeft      =   3390
   ClientTop       =   7575
   ClientWidth     =   7080
   Height          =   4440
   Icon            =   "MainlineEasyScreenSpecificCVs.frx":0000
   Left            =   3330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7080
   Top             =   7230
   Width           =   7200
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckBoxCVBit 
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox TextBoxCVValue 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7440
      Top             =   1800
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7440
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7440
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
   Begin VB.Label LabelBit 
      Caption         =   "Bit 8 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   30
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 7 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 6 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 5 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 4 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 3 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 2 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label LabelBit 
      Caption         =   "Bit 1 -"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   21
      Top             =   3480
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   20
      Top             =   3240
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   19
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   18
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   17
      Top             =   2520
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   16
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label LabelBitDescription 
      Caption         =   "No Data"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   "This easy screen helps you properly set the bits of the specific configuration variable as set by the manufactuer of this decoder."
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MainlineEasyScreenSpecificCVs.frx":0442
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
      Caption         =   "Bit Value"
      Height          =   195
      Left            =   525
      TabIndex        =   11
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
      Caption         =   "Bit Number      8      7      6      5      4      3      2      1"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3660
   End
   Begin VB.Label LabelCVNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "CV"
      Height          =   195
      Left            =   5160
      TabIndex        =   0
      Top             =   1080
      Width           =   390
   End
End
Attribute VB_Name = "MainlineEasyScreenSpecificCvs"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButtonClose_Click()

For X = 49 To 64
    If MainlineDecoder!Option1(X).Value = True Then
        Let MainlineDecoder!LocomotiveDecoderCVd(X).Text = textboxcvvalue.Text
    End If
Next X
    
For X = 112 To 128
    If MainlineDecoder!Option1(X).Value = ture Then
        Let MainlineDecoder!LocomotiveDecoderCVd(X).Text = textboxcvvalue.Text
    End If
Next X
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
            If TemporaryScreen = "Mainline Easy Screen Specific CVs Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
            ElseIf TemporaryScreen = "Screen Attribute Setting Screen" Then
                ScreenAttributeSetting.Show vbModeless
            ElseIf TemporaryScreen = "Sound Screen" Then
                SoundScreen.Show vbModeless
            ElseIf TemporaryScreen = "Sound Screen Edit List Screen" Then
                SoundScreenEditList.Show vbModeless
            ElseIf TemporaryScreen = "Room Lighting Control Screen" Then
                RoomLightingControl.Show vbModeless
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
                Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub





Private Sub ButtonPrint_Click()

    MainlineEasyScreenSpecificCvs.PrintForm
    
End Sub

Private Sub CheckBoxCVBit_Click(Index As Integer)

If CheckBoxCVBit(Index).Value = vbUnchecked Then
    If Index = 1 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 1
    If Index = 2 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 2
    If Index = 3 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 4
    If Index = 4 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 8
    If Index = 5 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 16
    If Index = 6 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 32
    If Index = 7 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 64
    If Index = 8 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) - 128
End If

If CheckBoxCVBit(Index).Value = vbChecked Then
    If Index = 1 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 1
    If Index = 2 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 2
    If Index = 3 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 4
    If Index = 4 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 8
    If Index = 5 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 16
    If Index = 6 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 32
    If Index = 7 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 64
    If Index = 8 Then textboxcvvalue.Text = Val(textboxcvvalue.Text) + 128
End If

Call UpdateLabels

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
        If TemporaryScreen = "Mainline Easy Screen Specific CVs Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Easy Screen Specific CVs Screen"
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
            Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
            Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Easy Screen Specific CVs Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineEasyScreenSpecificCvs.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineEasyScreenSpecificCvs.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineEasyScreenSpecificCvs.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineEasyScreenSpecificCvs.Height)

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
                Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
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
            Let Ini.Value = "Mainline Easy Screen Specific CVs Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineEasyScreenSpecificCvs.Hide
    'unload Mainlineeasyscreenspecificcvs

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
    Let Ini.Application = "Mainline Easy Screen Specific CVs Screen"
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
        MainlineEasyScreenSpecificCvs.Left = (Screen.Width - Width) / 2
        MainlineEasyScreenSpecificCvs.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineEasyScreenSpecificCvs.Width > Screen.Width Then
            Let MainlineEasyScreenSpecificCvs.Left = Screen.Width - MainlineEasyScreenSpecificCvs.Width
        Else
            Let MainlineEasyScreenSpecificCvs.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineEasyScreenSpecificCvs.Height > Screen.Height Then
            Let MainlineEasyScreenSpecificCvs.Top = Screen.Height - MainlineEasyScreenSpecificCvs.Height
        Else
            Let MainlineEasyScreenSpecificCvs.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This text box displays the current value of the" & vbCrLf & "configuration variable. When this window is closed" & vbCrLf & "the value of the configuration variable will be" & vbCrLf & "saved to the 'Mainline Mode - Decoder' screen."
        Let TemporaryText2 = "Configuration Value"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxcvvalue)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxcvvalue, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit one of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "one changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit One of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(1))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(1), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit two of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "two changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Two of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(2))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(2), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit three of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "three changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Three of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(3))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(3), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit four of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "four changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Four of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(4))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(4), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit five of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "five changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Five of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(5))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(5), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit six of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "six changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Six of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(6))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(6), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit seven of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "seven changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Seven of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(7))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(7), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box enables or disables bit eight of this" & vbCrLf & "configuration variable. Enabling or disabling bit" & vbCrLf & "eight changes specific function(s) of the" & vbCrLf & "decoder. See the notes below."
        Let TemporaryText2 = "Bit Eight of Specifc Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxCVBit(8))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxCVBit(8), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button closes this window and returns you" & vbCrLf & "to the 'Mainline Mode - Decoder' window."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Check Boxes to Bits
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim TemporaryVariableT As Integer
    
    If MainlineDecoder!SSTab2.Tab = 0 Then
        For TemporaryVariableT = 49 To 64
            If MainlineDecoder!Option1(TemporaryVariableT).Value = True Then
        
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 128) Then
                    Let CheckBoxCVBit(8).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 64) Then
                    Let CheckBoxCVBit(7).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 32) Then
                    Let CheckBoxCVBit(6).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 16) Then
                    Let CheckBoxCVBit(5).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 8) Then
                    Let CheckBoxCVBit(4).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 4) Then
                    Let CheckBoxCVBit(3).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 2) Then
                    Let CheckBoxCVBit(2).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 1) Then
                    Let CheckBoxCVBit(1).Value = vbChecked
                End If
                Call UpdateLabels
            End If
        Next TemporaryVariableT
    End If
    
    If MainlineDecoder!SSTab2.Tab = 1 Then
        For TemporaryVariableT = 112 To 128
            If MainlineDecoder!Option1(t).Value = True Then
        
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 128) Then
                    Let CheckBoxCVBit(8).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 64) Then
                    Let CheckBoxCVBit(7).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 32) Then
                    Let CheckBoxCVBit(6).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 16) Then
                    Let CheckBoxCVBit(5).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 8) Then
                    Let CheckBoxCVBit(4).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 4) Then
                    Let CheckBoxCVBit(3).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 2) Then
                    Let CheckBoxCVBit(2).Value = vbChecked
                End If
                If (Val(MainlineDecoder!LocomotiveDecoderCVd(TemporaryVariableT).Text) And 1) Then
                    Let CheckBoxCVBit(1).Value = vbChecked
                End If
                Call UpdateLabels
            End If
        Next TemporaryVariableT
    End If

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





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



Private Sub UpdateLabels()
         
If MainlineDecoder!SSTab2.Tab = 0 Then
    For t = 49 To 64
    If MainlineDecoder!Option1(t).Value = True Then
        Let LabelCVNumber.Caption = "CV" & CStr(t)
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1Used") = True Then
            Let CheckBoxCVBit(1).Enabled = True
            Let LabelBit(1).Enabled = True
            Let LabelBitDescription(1).Enabled = True
        Else
            Let CheckBoxCVBit(1).Enabled = False
            Let LabelBit(1).Enabled = False
            Let LabelBitDescription(1).Enabled = False
        End If
        If CheckBoxCVBit(1) = vbChecked Then
            Let LabelBitDescription(1).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1On")
        Else
            Let LabelBitDescription(1).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
                If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2Used") = True Then
            Let CheckBoxCVBit(2).Enabled = True
            Let LabelBit(2).Enabled = True
            Let LabelBitDescription(2).Enabled = True
        Else
            Let CheckBoxCVBit(2).Enabled = False
            Let LabelBit(2).Enabled = False
            Let LabelBitDescription(2).Enabled = False
        End If
                If CheckBoxCVBit(2) = vbChecked Then
            Let LabelBitDescription(2).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2On")
        Else
            Let LabelBitDescription(2).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3Used") = True Then
            Let CheckBoxCVBit(3).Enabled = True
            Let LabelBit(3).Enabled = True
            Let LabelBitDescription(3).Enabled = True
        Else
            Let CheckBoxCVBit(3).Enabled = False
            Let LabelBit(3).Enabled = False
            Let LabelBitDescription(3).Enabled = False
        End If
                If CheckBoxCVBit(3) = vbChecked Then
            Let LabelBitDescription(3).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3On")
        Else
            Let LabelBitDescription(3).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4Used") = True Then
            Let CheckBoxCVBit(4).Enabled = True
            Let LabelBit(4).Enabled = True
            Let LabelBitDescription(4).Enabled = True
        Else
            Let CheckBoxCVBit(4).Enabled = False
            Let LabelBit(4).Enabled = False
            Let LabelBitDescription(4).Enabled = False
        End If
        If CheckBoxCVBit(4) = vbChecked Then
            Let LabelBitDescription(4).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4On")
        Else
            Let LabelBitDescription(4).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4Off")
        End If
        '------------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5Used") = True Then
            Let CheckBoxCVBit(5).Enabled = True
            Let LabelBit(5).Enabled = True
            Let LabelBitDescription(5).Enabled = True
        Else
            Let CheckBoxCVBit(5).Enabled = False
            Let LabelBit(5).Enabled = False
            Let LabelBitDescription(5).Enabled = False
        End If
        If CheckBoxCVBit(5) = vbChecked Then
            Let LabelBitDescription(5).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5On")
        Else
            Let LabelBitDescription(5).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6Used") = True Then
            Let CheckBoxCVBit(6).Enabled = True
            Let LabelBit(6).Enabled = True
            Let LabelBitDescription(6).Enabled = True
        Else
            Let CheckBoxCVBit(6).Enabled = False
            Let LabelBit(6).Enabled = False
            Let LabelBitDescription(6).Enabled = False
        End If
        If CheckBoxCVBit(6) = vbChecked Then
            Let LabelBitDescription(6).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6On")
        Else
            Let LabelBitDescription(6).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7Used") = True Then
            Let CheckBoxCVBit(7).Enabled = True
            Let LabelBit(7).Enabled = True
            Let LabelBitDescription(7).Enabled = True
        Else
            Let CheckBoxCVBit(7).Enabled = False
            Let LabelBit(7).Enabled = False
            Let LabelBitDescription(7).Enabled = False
        End If
        If CheckBoxCVBit(7) = vbChecked Then
            Let LabelBitDescription(7).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7On")
        Else
            Let LabelBitDescription(7).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
        If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8Used") = True Then
            Let CheckBoxCVBit(8).Enabled = True
            Let LabelBit(8).Enabled = True
            Let LabelBitDescription(8).Enabled = True
        Else
            Let CheckBoxCVBit(8).Enabled = False
            Let LabelBit(8).Enabled = False
            Let LabelBitDescription(8).Enabled = False
        End If
        
        If CheckBoxCVBit(8) = vbChecked Then
            Let LabelBitDescription(8).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8On")
        Else
            Let LabelBitDescription(8).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8Off")
        End If
        ' -----------------------------------------------------------------------------------------------------------------
    End If
    Next t
Else
    If MainlineDecoder!SSTab2.Tab = 1 Then
        For t = 49 To 64
        If MainlineDecoder!Option1(t).Value = True Then
            Let LabelCVNumber.Caption = "CV" & CStr(t)
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1Used") = True Then
                Let CheckBoxCVBit(1).Enabled = True
                Let LabelBit(1).Enabled = True
                Let LabelBitDescription(1).Enabled = True
            Else
                Let CheckBoxCVBit(1).Enabled = False
                Let LabelBit(1).Enabled = False
                Let LabelBitDescription(1).Enabled = False
            End If
            If CheckBoxCVBit(1) = vbChecked Then
                Let LabelBitDescription(1).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1On")
            Else
                Let LabelBitDescription(1).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit1Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2Used") = True Then
               Let CheckBoxCVBit(2).Enabled = True
                Let LabelBit(2).Enabled = True
                Let LabelBitDescription(2).Enabled = True
            Else
                Let CheckBoxCVBit(2).Enabled = False
                Let LabelBit(2).Enabled = False
                Let LabelBitDescription(2).Enabled = False
            End If
                    If CheckBoxCVBit(2) = vbChecked Then
                Let LabelBitDescription(2).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2On")
            Else
                Let LabelBitDescription(2).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit2Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3Used") = True Then
                Let CheckBoxCVBit(3).Enabled = True
                Let LabelBit(3).Enabled = True
                Let LabelBitDescription(3).Enabled = True
            Else
                Let CheckBoxCVBit(3).Enabled = False
                Let LabelBit(3).Enabled = False
                Let LabelBitDescription(3).Enabled = False
            End If
            If CheckBoxCVBit(3) = vbChecked Then
                Let LabelBitDescription(3).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3On")
            Else
                Let LabelBitDescription(3).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit3Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4Used") = True Then
                Let CheckBoxCVBit(4).Enabled = True
                Let LabelBit(4).Enabled = True
                Let LabelBitDescription(4).Enabled = True
            Else
                Let CheckBoxCVBit(4).Enabled = False
                Let LabelBit(4).Enabled = False
                Let LabelBitDescription(4).Enabled = False
            End If
            If CheckBoxCVBit(4) = vbChecked Then
                Let LabelBitDescription(4).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4On")
            Else
                Let LabelBitDescription(4).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit4Off")
            End If
            '------------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5Used") = True Then
                Let CheckBoxCVBit(5).Enabled = True
                Let LabelBit(5).Enabled = True
                Let LabelBitDescription(5).Enabled = True
            Else
                Let CheckBoxCVBit(5).Enabled = False
                Let LabelBit(5).Enabled = False
                Let LabelBitDescription(5).Enabled = False
            End If
            If CheckBoxCVBit(5) = vbChecked Then
                Let LabelBitDescription(5).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5On")
            Else
                Let LabelBitDescription(5).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit5Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6Used") = True Then
                Let CheckBoxCVBit(6).Enabled = True
                Let LabelBit(6).Enabled = True
                Let LabelBitDescription(6).Enabled = True
            Else
                Let CheckBoxCVBit(6).Enabled = False
                Let LabelBit(6).Enabled = False
                Let LabelBitDescription(6).Enabled = False
            End If
            If CheckBoxCVBit(6) = vbChecked Then
                Let LabelBitDescription(6).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6On")
            Else
                Let LabelBitDescription(6).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit6Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7Used") = True Then
                Let CheckBoxCVBit(7).Enabled = True
                Let LabelBit(7).Enabled = True
                Let LabelBitDescription(7).Enabled = True
            Else
                Let CheckBoxCVBit(7).Enabled = False
                Let LabelBit(7).Enabled = False
                Let LabelBitDescription(7).Enabled = False
            End If
            If CheckBoxCVBit(7) = vbChecked Then
                Let LabelBitDescription(7).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7On")
            Else
                Let LabelBitDescription(7).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit7Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
            If MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8Used") = True Then
                Let CheckBoxCVBit(8).Enabled = True
                Let LabelBit(8).Enabled = True
                Let LabelBitDescription(8).Enabled = True
            Else
                Let CheckBoxCVBit(8).Enabled = False
                Let LabelBit(8).Enabled = False
                Let LabelBitDescription(8).Enabled = False
            End If
            If CheckBoxCVBit(8) = vbChecked Then
                Let LabelBitDescription(8).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8On")
            Else
                Let LabelBitDescription(8).Caption = MainlineDecoder!SpecificCVs(t).Recordset.Fields("Bit8Off")
            End If
            ' -----------------------------------------------------------------------------------------------------------------
        End If
        Next t
    End If
End If

End Sub

Private Sub Form_Resize()

    If MainlineEasyScreenSpecificCvs.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineEasyScreenSpecificCvs.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Easy Screen Specific CVs"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineEasyScreenSpecificCvs.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineEasyScreenSpecificCvs.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineEasyScreenSpecificCvs.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineEasyScreenSpecificCvs.Height)
 
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


