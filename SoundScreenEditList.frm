VERSION 4.00
Begin VB.Form SoundScreenEditList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Editing Sounds"
   ClientHeight    =   1905
   ClientLeft      =   5805
   ClientTop       =   10155
   ClientWidth     =   6345
   Height          =   2310
   Icon            =   "SoundScreenEditList.frx":0000
   Left            =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6345
   Top             =   9810
   Width           =   6465
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   1260
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "SoundScreenEditList.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonEdit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton ButtonDelete 
      Caption         =   "&Delete"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton ButtonAddNew 
      Caption         =   "&Add New"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox SoundFileDisplay 
      Alignment       =   2  'Center
      DataField       =   "SoundFileName"
      DataSource      =   "SoundData"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Additional\716dupont.wav"
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Data SoundData 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SoundList"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   6720
      Top             =   600
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Current selected wave file to edit or delete."
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   3015
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   6720
      Top             =   1680
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   6720
      Top             =   1080
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin MCI.MMControl SoundRecordings 
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      AutoEnable      =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Try adding more wave files to your playlist database or removing the ones you do;nt want to hear."
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   1680
   End
End
Attribute VB_Name = "SoundScreenEditList"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButtonAddNew_Click()

If ButtonAddNew.Caption = "&Add New" Then
    Let SoundRecordings.Visible = False
    Let ButtonEdit.Caption = "&Update"
    Let ButtonDelete.Enabled = False
    Let ButtonAddNew.Caption = "&Cancel"
    Let ButtonClose.Enabled = False
    SoundData.Recordset.AddNew
Else
    Let SoundRecordings.Visible = True
    Let ButtonEdit.Caption = "&Edit"
    Let ButtonDelete.Enabled = True
    Let ButtonAddNew.Caption = "&Add New"
    Let ButtonClose.Enabled = True
    
    SoundData.Recordset.AbsolutePosition = 1
        
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
            If TemporaryScreen = "Sound Screen Edit List Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Sound Screen Edit List Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Sound Screen Edit List Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Sound Screen Edit List Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


Private Sub ButtonDelete_Click()


    Let ButtonEdit.Enabled = False
    Let ButtonDelete.Enabled = False
    Let ButtonAddNew.Enabled = False
    Let ButtonClose.Enabled = False
    
    SoundData.Recordset.Delete
    SoundData.Refresh

    Let ButtonEdit.Enabled = True
    Let ButtonDelete.Enabled = True
    Let ButtonAddNew.Enabled = True
    Let ButtonClose.Enabled = True
    
End Sub

Private Sub ButtonEdit_Click()

If ButtonEdit.Caption = "&Edit" Then
    Let ButtonEdit.Caption = "&Update"
    Let ButtonDelete.Enabled = False
    Let ButtonAddNew.Enabled = False
    Let ButtonClose.Enabled = False
    SoundData.Recordset.Edit
    SoundFileDisplay.Enabled = True
    SoundFileDisplay.SetFocus
    SoundRecordings.PlayEnabled = False
    SoundRecordings.PrevEnabled = False
    SoundRecordings.NextEnabled = False
Else
    Let ButtonEdit.Caption = "&Edit"
    Let ButtonDelete.Enabled = True
    Let ButtonAddNew.Caption = "&Add New"
    Let ButtonAddNew.Enabled = True
    Let ButtonClose.Enabled = True
    SoundData.Recordset.Update
    SoundFileDisplay.Enabled = False
    SoundRecordings.PlayEnabled = True
    SoundRecordings.PrevEnabled = True
    SoundRecordings.NextEnabled = True

End If

End Sub

Private Sub ButtonForward_Click()

SoundData.Recordset.MoveNext

If SoundData.Recordset.EOF Then SoundData.Recordset.MoveFirst

End Sub


Private Sub ButtonPlay_Click()

If ButtonPlay.Caption = "&Play" Then
    Let ButtonPlay.Caption = "&Stop"
    Let ButtonEdit.Enabled = False
    Let ButtonDelete.Enabled = False
    Let ButtonAddNew.Enabled = False
    Let ButtonClose.Enabled = False
    Let SoundRecordings.Notify = False
    Let SoundRecordings.Wait = False
    Let SoundRecordings.Shareable = False
    Let SoundRecordings.Filename = App.Path$ & "\Sounds\" + SoundFileDisplay.Text
    Let SoundRecordings.Command = "Open"
    Let SoundRecordings.Command = "Play"
Else
    Let ButtonPlay.Caption = "&Play"
    Let ButtonEdit.Enabled = True
    Let ButtonDelete.Enabled = True
    Let ButtonAddNew.Enabled = True
    Let ButtonClose.Enabled = True
    Let SoundRecordings.Command = "Stop"
    Let SoundRecordings.Command = "Close"
End If

End Sub




Private Sub ButtonReverse_Click()

SoundData.Recordset.MovePrevious

If SoundData.Recordset.BOF Then SoundData.Recordset.MoveLast

End Sub


Private Sub ButtonPrint_Click()

    SoundScreenEditList.PrintForm
    
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
        If TemporaryScreen = "Sound Screen Edit List Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Sound Screen Edit List Screen"
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
            Let Ini.Value = "Sound Screen Edit List Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Sound Screen Edit List Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Sound Screen Edit List Screen Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Sound Screen Edit List Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(SoundScreenEditList.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(SoundScreenEditList.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(SoundScreenEditList.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(SoundScreenEditList.Height)

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
                Let Ini.Value = "Sound Screen Edit List Screen, Form Deactivate, variable error in ATC.INI file for 'Trnsparency' setting."
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
            Let Ini.Value = "Sound Screen Edit List Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    SoundScreenEditList.Hide

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
    Let Ini.Application = "Sound Screen Edit List Screen"
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
        SoundScreenEditList.Left = (Screen.Width - Width) / 2
        SoundScreenEditList.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + SoundScreenEditList.Width > Screen.Width Then
            Let SoundScreenEditList.Left = Screen.Width - SoundScreenEditList.Width
        Else
            Let SoundScreenEditList.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + SoundScreenEditList.Height > Screen.Height Then
            Let SoundScreenEditList.Top = Screen.Height - SoundScreenEditList.Height
        Else
            Let SoundScreenEditList.Top = Val(TemporaryValueTop)
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
       
        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This text box is where all information from your" & vbCrLf & "serial port is displayed. Commands given by the" & vbCrLf & "program are displayed here. You can also type your" & vbCrLf & "own commands, providing the port is not busy."
        Let TemporaryText2 = "Communication Window"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxCommunicationWindowDCC)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxCommunicationWindowDCC, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let SoundData.DatabaseName = App.Path$ & "\Databases\SoundFileDatabase.mdb"
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




Private Sub Form_Resize()

    If SoundScreenEditList.WindowState = vbMinimized Then
    
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
        
    ElseIf SoundScreenEditList.WindowState = vbNormal Then
    
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

Private Sub SoundRecordings_Done(NotifyCode As Integer)
    
'    Let SoundRecordings.Command = "Stop"
'    Let SoundRecordings.Command = "Close"
'    Let SoundRecordings.PrevEnabled = True
'    Let SoundRecordings.NextEnabled = True

'    Let ButtonEdit.Enabled = True
'    Let ButtonDelete.Enabled = True
'    Let ButtonAddNew.Enabled = True
'    Let ButtonClose.Enabled = True

End Sub


Private Sub SoundRecordings_NextClick(Cancel As Integer)

    SoundData.Recordset.MoveNext

    If SoundData.Recordset.EOF Then
        SoundData.Recordset.MoveFirst
    End If

End Sub

Private Sub SoundRecordings_PauseClick(Cancel As Integer)

    'If SoundRecordings.PlayEnabled = True Then
    '    SoundRecordings.PlayEnabled = False
    'ElseIf SoundRecordings.PlayEnabled = False Then
    '    SoundRecordings.PlayEnabled = True
    'End If
    
End Sub

Private Sub SoundRecordings_PlayClick(Cancel As Integer)
    
    Let SoundRecordings.Filename = App.Path$ & "\Sounds\" + SoundFileDisplay.Text
    
    Let SoundRecordings.PlayEnabled = False
    Let SoundRecordings.NextEnabled = False
    Let SoundRecordings.PrevEnabled = False
    Let SoundRecordings.PauseEnabled = True
    Let SoundRecordings.StopEnabled = True
    Let SoundRecordings.Notify = False
    Let SoundRecordings.Wait = False
    Let SoundRecordings.Shareable = False
    Let SoundRecordings.Command = "Open"
    Let SoundRecordings.Command = "Play"
    
    Let ButtonEdit.Enabled = False
    Let ButtonDelete.Enabled = False
    Let ButtonAddNew.Enabled = False
    Let ButtonClose.Enabled = False

End Sub


Private Sub SoundRecordings_PrevClick(Cancel As Integer)

    SoundData.Recordset.MovePrevious

    If SoundData.Recordset.BOF Then
        SoundData.Recordset.MoveLast
    End If

End Sub

Private Sub SoundRecordings_StopClick(Cancel As Integer)
    
    
    Let SoundRecordings.Command = "Close"
    Let SoundRecordings.StopEnabled = False
    Let SoundRecordings.NextEnabled = True
    Let SoundRecordings.PrevEnabled = True
    Let SoundRecordings.PlayEnabled = True
    Let SoundRecordings.PauseEnabled = False
    Let ButtonEdit.Enabled = True
    Let ButtonDelete.Enabled = True
    Let ButtonAddNew.Enabled = True
    Let ButtonClose.Enabled = True

End Sub






Private Sub BalloonHelpUpdate()

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

        Let TemporaryText1 = "This text box displays the currently selected wave file" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "to edit or delete from the database."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SoundFileDisplay)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SoundFileDisplay, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This control allows you to play wave files or move" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "forward or backward in the database to edit or" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "delete from the database."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(SoundRecordings)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(SoundRecordings, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This button allows you to edit the file name of" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "the currently selected wave in the database."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonEdit)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonEdit, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
            
        Let TemporaryText1 = "This button deletes the currently selected file" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "from the play list database."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonDelete)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonDelete, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This button allows you to add a new file" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "to the play list database."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonDelete)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonDelete, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This button closes the Editing Sound" + vbCrLf
        Let TemporaryText1 = TemporaryText1 + "window and returns you to the main window of ATC."
        Let TemporaryText2 = "Sound Screen Edit"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonDelete)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonDelete, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    Else
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
    End If
End Sub
