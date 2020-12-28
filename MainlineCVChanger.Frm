VERSION 4.00
Begin VB.Form MainlineCVChanger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoamtic Train Control - CV Changer"
   ClientHeight    =   5940
   ClientLeft      =   6765
   ClientTop       =   5010
   ClientWidth     =   5205
   Height          =   6345
   Icon            =   "MainlineCVChanger.frx":0000
   Left            =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5205
   Top             =   4665
   Width           =   5325
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton ButtonStart3 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDatabaseNumberSearch 
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox LocomotiveDatabaseDecoderSearch 
      Caption         =   "Decoder Equiped"
      DataField       =   "LocomotiveDecoderEquiped"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   255
      Left            =   6720
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox CheckBoxAddressShort 
      Caption         =   "Short Address"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   3900
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton ButtonStart2 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Data LocomotiveDatabaseSearch 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBase"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TextBoxLocomotiveNumberCurrent 
      Height          =   285
      Left            =   6720
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton ButtonStart1 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox TextBoxLocomotiveNumberStop 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TextBoxStartLocomotiveNumber 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   150
      Picture         =   "MainlineCVChanger.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   480
   End
   Begin VB.Frame Frame7 
      Caption         =   "Changing Configuration Variables"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   5055
      Begin VB.CheckBox ConsistControlCVWright 
         Caption         =   "Wright Configuration Variable"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox ConsistControlCVRead 
         Caption         =   "Read or"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox ConsistControlCVValue 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox ConsistControlCV 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Configuration Variable"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   6720
      Top             =   4380
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   6720
      Top             =   5520
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   6720
      Top             =   4920
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label18 
      Caption         =   "Configure all locomotive in the current consist"
      Height          =   270
      Left            =   135
      TabIndex        =   25
      Top             =   3390
      Width           =   3495
   End
   Begin VB.Label Label17 
      Caption         =   "Configure all locomotive active in your database."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label16 
      Caption         =   "Configure all locomotives in the specified search or"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Currently at"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Locomotive Number to Stop at for changing CVs"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "Locomotive Number to Start at for changing CVs"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   $"MainlineCVChanger.frx":0884
      Height          =   615
      Left            =   810
      TabIndex        =   15
      Top             =   120
      Width           =   4365
   End
   Begin VB.Label LocomotiveCommunicationStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "MainlineCVChanger"
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
            If TemporaryScreen = "Mainline CV Changer Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline CV Changer Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline CV Changer Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline CV Changer Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub






Private Sub ButtonPrint_Click()

    MainlineCVChanger.PrintForm
    
End Sub

Private Sub ButtonStart1_Click()

If ButtonStart1.Caption = "&Start" Then

    Let ButtonClose.Enabled = False
    Let ButtonStart1.Caption = "&Stop"
    Let ButtonStart2.Enabled = False
    Let ButtonStart3.Enabled = False

    While ButtonStart1.Caption = "&Stop"
    
    Let consistcontrolmacrolocomotivenumber.Text = Val(TextBoxLocomotiveNumberCurrent.Text) + 1
    
    Call SetLocomotiveNumber
    Call SetConfigurationVariable
    Call MainScreen.SendCommandviaTrackQ
    
    Wend
Else
    Let ButtonClose.Enabled = True
    Let ButtonStart1.Caption = "&Start"
    Let ButtonStart2.Enabled = True
    Let ButtonStart3.Enabled = True
End If

End Sub

Private Sub ButtonStart2_Click()
    Let LocomotiveCommunicationStatus.Caption = "Changing Configuration Variables"
    Let MousePointer = vbArrowHourglass
    LocomotiveDatabaseSearch.Recordset.MoveFirst
    If ButtonStart2.Caption = "&Start" Then
        Let ButtonClose.Enabled = False
        Let ButtonStart1.Enabled = False
        Let ButtonStart2.Caption = "&Stop"
        Let ButtonStart3.Enabled = False
            While ButtonStart2.Caption = "&Stop" And Not LocomotiveDatabaseSearch.Recordset.EOF
                LocomotiveDatabaseSearch.Recordset.MoveNext
                If LocomotiveDatabaseDecoderSearch.Value = vbChecked Then
                    Let TextBoxLocomotiveNumberCurrent.Text = Val(LocomotiveDatabaseNumberSearch.Text)
                    Call SetLocomotiveNumber
                    Call SetConfigurationVariable
                    Call MainScreen.SendCommandviaTrackQ
                End If
            Wend
        Let ButtonClose.Enabled = True
        Let ButtonStart1.Enabled = True
        Let ButtonStart2.Caption = "&Start"
        Let ButtonStart3.Enabled = True
        Let MousePointer = vbDefault
    Else
        Let ButtonClose.Enabled = True
        Let ButtonStart1.Enabled = True
        Let ButtonStart2.Caption = "&Start"
        Let ButtonStart3.Enabled = True
        Let MousePointer = vbDefault
    End If
    Let LocomotiveCommunicationStatus.Caption = "Status: Clear"

End Sub



Private Sub ConsistControlCV_Change()

If Val(ConsistControlCV.Text) <> 0 Then
    If Val(ConsistControlCVValue.Text) > 0 And _
       Val(ConsistControlCVValue.Text) < 255 Then
        Let ButtonStart1.Enabled = True
        Let ButtonStart2.Enabled = True
        Let ButtonStart3.Enabled = True
    Else
        Let ButtonStart1.Enabled = False
        Let ButtonStart2.Enabled = False
        Let ButtonStart3.Enabled = False
    End If
Else
    Let ButtonStart1.Enabled = False
    Let ButtonStart2.Enabled = False
    Let ButtonStart3.Enabled = False
End If

End Sub

Private Sub ConsistControlCVRead_Click()

If ConsistControlCVRead.Value = vbChecked Then
    Let ConsistControlCVWright.Value = vbUnchecked
Else
    Let ConsistControlCVWright.Value = vbChecked
End If

End Sub

Private Sub ConsistControlCVValue_Change()

If Val(ConsistControlCV.Text) <> 0 Then
    
    If Val(Str$(ConsistControlCVValue.Text)) >= 0 And _
       Val(Str$(ConsistControlCVValue.Text)) < 255 And _
       Str$(ConsistControlCVValue.Text) <> "" Then
       
        Let ButtonStart1.Enabled = True
        Let ButtonStart2.Enabled = True
        Let ButtonStart3.Enabled = True
    Else
        Let ButtonStart1.Enabled = False
        Let ButtonStart2.Enabled = False
        Let ButtonStart3.Enabled = False
    End If
Else
    Let ButtonStart1.Enabled = False
    Let ButtonStart2.Enabled = False
    Let ButtonStart3.Enabled = False
End If

End Sub

Private Sub ConsistControlCVWright_Click()

If ConsistControlCVWright.Value = vbChecked Then
    Let ConsistControlCVRead.Value = vbUnchecked
Else
    Let ConsistControlCVRead.Value = vbChecked
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
        If TemporaryScreen = "Mainline CV Changer Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline CV Changer Screen"
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
            Let Ini.Value = "Mainline CV Changer Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline CV Changer Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline CV Changer Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unloading the Form
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline CV Changer Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlineCVChanger.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlineCVChanger.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlineCVChanger.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlineCVChanger.Height)
    Let Ini.Parameter = "TextBoxStartLocomotiveNumber"
    Let Ini.Value = TextBoxStartLocomotiveNumber.Text
    Let Ini.Parameter = "TextBoxLocomotiveNumberStop"
    Let Ini.Value = TextBoxLocomotiveNumberStop.Text
    Let Ini.Parameter = "ConsistControlCV"
    Let Ini.Value = ConsistControlCV.Text
    Let Ini.Parameter = "ConsistControlCVValue"
    Let Ini.Value = ConsistControlCVValue.Text
    Let Ini.Parameter = "ConsistControlCVRead"
    Let Ini.Value = Str(ConsistControlCVRead.Value)
    Let Ini.Parameter = "ConsistControlCVWright"
    Let Ini.Value = Str(ConsistControlCVWright.Value)
    
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
                Let Ini.Value = "Mainline CV Changer Screen, Form Dectivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline CV Changer Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineCVChanger.Hide
    'unload Mainlinecvchanger

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
    Let Ini.Application = "Mainline CV Changer Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "TextBoxStartLocomotiveNumber"
    Let TextBoxStartLocomotiveNumber.Text = Ini.Value
    Let Ini.Parameter = "TextBoxLocomotiveNumberStop"
    Let TextBoxLocomotiveNumberStop.Text = Ini.Value
    Let Ini.Parameter = "ConsistControlCV"
    Let ConsistControlCV.Text = Ini.Value
    Let Ini.Parameter = "ConsistControlCVValue"
    Let ConsistControlCVValue.Text = Ini.Value
    Let Ini.Parameter = "ConsistControlCVRead"
    Let ConsistControlCVRead.Value = Val(Ini.Value)
    Let Ini.Parameter = "ConsistControlCVWright"
    Let ConsistControlCVWright.Value = Val(Ini.Value)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        MainlineCVChanger.Left = (Screen.Width - Width) / 2
        MainlineCVChanger.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineCVChanger.Width > Screen.Width Then
            Let MainlineCVChanger.Left = Screen.Width - MainlineCVChanger.Width
        Else
            Let MainlineCVChanger.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineCVChanger.Height > Screen.Height Then
            Let MainlineCVChanger.Top = Screen.Height - MainlineCVChanger.Height
        Else
            Let MainlineCVChanger.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This text box contains the low locomotive number" & vbCrLf & "to change a configuration variable."
        Let TemporaryText2 = "Low Locomotive Number"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxStartLocomotiveNumber)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxStartLocomotiveNumber, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This text box contains the high locomotive number" & vbCrLf & "to change a configuration variable."
        Let TemporaryText2 = "High Locomotive Number"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxLocomotiveNumberStop)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxLocomotiveNumberStop, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button will start changing the configuration" & vbCrLf & "variable for all the locomotives in the specified area" & vbCrLf & "(low locomotive number to high)."
        Let TemporaryText2 = "Start Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonStart1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonStart1, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button will start changing the configuration" & vbCrLf & "variable for active locomotives in the specified area" & vbCrLf & "(low locomotive number to high)."
        Let TemporaryText2 = "Start Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonStart2)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonStart2, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button will start changing the configuration" & vbCrLf & "variable for locomotives in a specific consist only."
        Let TemporaryText2 = "Start Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonStart3)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonStart3, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This text box contain the configuration variable" & vbCrLf & "that you wish to change."
        Let TemporaryText2 = "Configuration Variable"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ConsistControlCV)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ConsistControlCV, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This text box contain the value of the configuration" & vbCrLf & "variable that you wish to change."
        Let TemporaryText2 = "Configuration Variable Value"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ConsistControlCVValue)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ConsistControlCVValue, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box confirms reading or writing of the" & vbCrLf & "configuration variable to the command station."
        Let TemporaryText2 = "Configuration Variable Value"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ConsistControlCVRead)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ConsistControlCVRead, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box confirms reading or writing of the" & vbCrLf & "configuration variable to the command station."
        Let TemporaryText2 = "Configuration Variable Value"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ConsistControlCVWright)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ConsistControlCVWright, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button closes the Mainline Configuration Variable" & vbCrLf & "Changer Window and returns you to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
 
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




Private Sub SevenByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let SevenByteH.Text = Hex(Val(sevenbyted.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SevenByteH.Text) = 1 Then
        Let SevenByteH.Text = "0" + SevenByteH.Text
    End If

    
Rem Binary Conversion
    
        Let temp = Val(sevenbyted.Text)
        
        If temp / 128 >= 1 Then
            Let SevenByteB.Text = "1"
            Let temp = temp - 128
            Else: SevenByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 64
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 32
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 16
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 8
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 4
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 2
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let SevenByteB.Text = SevenByteB.Text + "1"
            Let temp = temp - 1
            Else: SevenByteB.Text = SevenByteB.Text + "0"
        End If

End Sub


Private Sub ShortAdDress_Click()

End Sub

Private Sub SixByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let SixByteH.Text = Hex(Val(SixByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(SixByteH.Text) = 1 Then
        Let SixByteH.Text = "0" + SixByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(SixByteD.Text)
        
        If temp / 128 >= 1 Then
            Let SixByteB.Text = "1"
            Let temp = temp - 128
            Else: SixByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 64
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 32
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 16
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 8
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 4
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 2
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let SixByteB.Text = SixByteB.Text + "1"
            Let temp = temp - 1
            Else: SixByteB.Text = SixByteB.Text + "0"
        End If

End Sub


Private Sub ThreeByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let ThreeByteH.Text = Hex(Val(ThreeByteD.Text))
    
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(ThreeByteH.Text) = 1 Then
        Let ThreeByteH.Text = "0" + ThreeByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(ThreeByteD.Text)
        
        If temp / 128 >= 1 Then
            Let ThreeByteB.Text = "1"
            Let temp = temp - 128
            Else: ThreeByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 64
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 32
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 16
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 8
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 4
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 2
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let ThreeByteB.Text = ThreeByteB.Text + "1"
            Let temp = temp - 1
            Else: ThreeByteB.Text = ThreeByteB.Text + "0"
        End If

End Sub


Private Sub TwoByteD_Change()
    
Rem Hexidecimal Converstion
        
    Let TwoByteH.Text = Hex(Val(twoByteD.Text))
            
' The following routine is required by the system becuase up the previous
' statement above does not correctly add zeros to straings that are
' less then two in length. Hexidecimal converstion does not have a leading zero.

    If Len(TwoByteH.Text) = 1 Then
        Let TwoByteH.Text = "0" + TwoByteH.Text
    End If

Rem Binary Conversion
    
        Let temp = Val(twoByteD.Text)
        
        If temp / 128 >= 1 Then
            Let TwoByteB.Text = "1"
            Let temp = temp - 128
            Else: TwoByteB.Text = "0"
        End If
        
        If temp / 64 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 64
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 32 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 32
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 16 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 16
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 8 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 8
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 4 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 4
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 2 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 2
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If
        
        If temp / 1 >= 1 Then
            Let TwoByteB.Text = TwoByteB.Text + "1"
            Let temp = temp - 1
            Else: TwoByteB.Text = TwoByteB.Text + "0"
        End If

End Sub



Private Sub SetLocomotiveNumber()

If ShortAdDress.Value = vbUnchecked Then
    Let MainScreen!onebyted.Text = Int(Val(TextBoxLocomotiveNumberCurrent.Text) / 256)
    Let MainScreen!twoByteD.Text = Val(TextBoxLocomotiveNumberCurrent.Text) - (Val(onebyted.Text) * 256)
    Let MainScreen!onebyted.Text = Val(onebyted.Text) + 128 + 64
    Let TextBoxCommunicationWindowDCC.Text = "Locomotive (long address) number " + TextBoxLocomotiveNumberCurrent.Text + " "
ElseIf ShortAdDress.Value = vbChecked Then
    Let MainScreen!onebyted.Text = ""
    Let MainScreen!twoByteD.Text = Int(Val(TextBoxLocomotiveNumberCurrent.Text))
    Let TextBoxCommunicationWindowDCC.Text = TextBoxCommunicationWindowDCC.Text & " Locomotive (short address) or Consist number " + TextBoxLocomotiveNumberCurrent.Text + "; "
End If

End Sub

Private Sub SetConfigurationVariable()
    
        Let TemporaryByteOne = 0
        Let TemporaryByteTwo = Val(ConsistControlCV.Text) - 1
        
        If TemporaryByteTwo / 512 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 2
            Let TemporaryByteTwo = TemporaryByteTwo - 512
        End If
        
        If TemporaryByteTwo / 256 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 1
            Let TemporaryByteTwo = TemporaryByteTwo - 256
        End If
        
        Let TemporaryByteOne = TemporaryByteOne + 128
        Let TemporaryByteOne = TemporaryByteOne + 64
        Let TemporaryByteOne = TemporaryByteOne + 32
        
       If ConsistControlCVRead = vbChecked Then
              Let TemporaryByteOne = TemporaryByteOne + 4
        Else
            Let TemporaryByteOne = TemporaryByteOne + 8 + 4
        End If
        
    Let MainScreen!ThreeByteD.Text = TemporaryByteOne
    Let MainScreen!FourByteD.Text = TemporaryByteTwo
    Let MainScreen!FiveByteD.Text = Val(ConsistControlCVValue.Text)
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    
 Let TextBoxCommunicationWindowDCC = TextBoxCommunicationWindowDCC + "Change CV" + ConsistControlCV.Text + " to " + ConsistControlCVValue.Text

 
End Sub


Private Sub Form_Resize()

    If MainlineCVChanger.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineCVChanger.WindowState = vbNormal Then
    
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


