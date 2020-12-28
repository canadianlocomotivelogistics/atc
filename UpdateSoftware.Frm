VERSION 4.00
Begin VB.Form UpdateSoftware 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Update Software"
   ClientHeight    =   5340
   ClientLeft      =   5340
   ClientTop       =   975
   ClientWidth     =   8925
   Height          =   5745
   Icon            =   "UpdateSoftware.frx":0000
   Left            =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8925
   Top             =   630
   Width           =   9045
   Begin VB.CommandButton ButtonCheck 
      Caption         =   "Check"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "UpdateSoftware.frx":0442
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
      Left            =   5940
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin FATHFTPLibCtl.FathFTPCtrl FtpControl 
      Left            =   7320
      OleObjectBlob   =   "UpdateSoftware.frx":0884
      Top             =   2940
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   1995
      Left            =   240
      TabIndex        =   8
      Top             =   2940
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "for additional help"
      Height          =   195
      Left            =   4440
      TabIndex        =   12
      Top             =   2100
      Width           =   1245
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status: Idle"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   6795
   End
   Begin SystemInfoControl.MSysInfo SysInfo 
      Left            =   7260
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label8 
      Caption         =   $"UpdateSoftware.frx":093E
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   6255
   End
   Begin April16_NetConnect.NetConnect NetConnect 
      Left            =   7320
      Top             =   1740
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7080
      Y1              =   2820
      Y2              =   2820
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx3 
      Height          =   195
      Left            =   240
      Top             =   2100
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   344
      Registered      =   -1  'True
      ProjectKey      =   "et0945B"
      ForeColor       =   8388608
      Target          =   "http://groups.yahoo.com/group/Automatic_Train_Control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "http://groups.yahoo.com/group/Automatic_Train_Control"
      Autosize        =   0   'False
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx2 
      Height          =   195
      Left            =   240
      Top             =   1860
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   344
      Registered      =   -1  'True
      ProjectKey      =   "et0945B"
      ForeColor       =   8388608
      Target          =   "dcalcutt@rogers.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "dcalcutt@rogers.com"
      TargetType      =   "1"
      EmailSubject    =   "Darrin J. Calcutt, I'm writing you to ask..."
      EmailBody       =   "Automatic Train Control - About"
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx1 
      Height          =   195
      Left            =   2340
      Top             =   1620
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   344
      Registered      =   -1  'True
      ProjectKey      =   "et0945B"
      ForeColor       =   8388608
      Target          =   "http://atc.lovethosetrains.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "http://atc.lovethosetrains.com"
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7320
      Top             =   120
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7260
      Top             =   660
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin IniconLib.Init Ini 
      Left            =   7320
      Top             =   1140
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
      AutoSize        =   -1  'True
      Caption         =   "and user input."
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2340
      Width           =   1050
   End
   Begin URLLabelCtl.URLLabel UrlLabel2 
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2100
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      ForeColor       =   8388608
      URL             =   "http://groups.yahoo.com/group/Automatic_Train_Control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "and don't forget to join the Yahoo group at "
      Height          =   195
      Left            =   1860
      TabIndex        =   4
      Top             =   1860
      Width           =   3030
   End
   Begin VB.Label Label2 
      Caption         =   "or email me at"
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   1620
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Please visit my home page at"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1620
      Width           =   2175
   End
End
Attribute VB_Name = "UpdateSoftware"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub ButtonCheck_Click()

    Let ButtonClose.Enabled = False
    If NetConnect.Connected Then
        FtpControl.Connect
        ListView.ListItems.Clear
            FtpControl.EnumFiles ("/Automatic Train Control/")
        'FtpControl.Disconnect
    End If
    
    Let ButtonClose.Enabled = True
End Sub


Private Sub ButtonClose_Click()
End
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
            If TemporaryScreen = "About Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!MenuLogFile.Caption = "&Log File is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "About Screen, Button Close, current window is not listed in the stack to remove it and hide."
            End If
            
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Open Previous Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Screen Stack"
            Let Ini.Parameter = CStr(TemporaryCounter - 1)
            Let TemporaryScreen = Ini.Value
            If TemporaryScreen = "About Screen" Then
                About.Show
            ElseIf TemporaryScreen = "Clock Screen" Then
                ClockScreen.Show
            ElseIf TemporaryScreen = "Communication Setting Screen " Then
                CommunicationSetting.Show
            ElseIf TemporaryScreen = "Define Block Properties Screen" Then
                DefineBlockProperties.Show
            ElseIf TemporaryScreen = "Define Blocks Screen" Then
                DefineBlocks.Show
            ElseIf TemporaryScreen = "Define Blocks Spreadsheet Screen" Then
                DefineBlocksSpreadsheet.Show
            'ElseIf TemporaryScreen = "Fun Screen" Then
            '    FunScreen.Show
            ElseIf TemporaryScreen = "Internet Settings Screen" Then
                InternetSettings.Show
            ElseIf TemporaryScreen = "Locomotive CV Spreadsheet Screen" Then
                LocomotiveCVSpreadsheet.Show
            ElseIf TemporaryScreen = "Locomotive Spreadsheet Screen" Then
                LocomotiveSpreadsheet.Show
            ElseIf TemporaryScreen = "Mainline Consist Screen" Then
                MainlineConsist.Show
            ElseIf TemporaryScreen = "Mainline CV Changer Screen" Then
                MainlineCVChanger.Show
            ElseIf TemporaryScreen = "Mainline Decoder Screen" Then
                MainlineDecoder.Show
            ElseIf TemporaryScreen = "Mainline Diesel Screen" Then
                MainlineDiesel.Show
            ElseIf TemporaryScreen = "Mainline Easy Screen Configuration Screen" Then
                MainlineEasyScreenConfiguration.Show
            ElseIf TemporaryScreen = "Mainline Easy Screen Consist Functions Screen" Then
                MainlineEasyScreenConsistFunctions.Show
            ElseIf TemporaryScreen = "Mainline Easy Screen Functions Screen" Then
                MainlineEasyScreenFunctions.Show
            ElseIf TemporaryScreen = "Mainline Easy Screen Specific CVs Screen" Then
                MainlineEasyScreenSpecificCvs.Show
            ElseIf TemporaryScreen = "Mainline Easy Screen Speed Table Screen" Then
                MainlineEasyScreenSpeedTable.Show
            ElseIf TemporaryScreen = "Mainline Macro Maker Screen" Then
                MainlineMacroMaker.Show
            ElseIf TemporaryScreen = "Mainline Operation ATC Screen" Then
                MainlineOperationATC.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Screen" Then
                MainlineOperationGUI.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel1 Screen" Then
                MainlineOperationGuiDiesel1Screen.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel2 Screen" Then
                MainlineOperationGuiDiesel2Screen.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Diesel3 Screen" Then
                MainlineOperationGuiDiesel3Screen.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Electric1 Screen" Then
                MainlineOperationGuiElectric1Screen.Show
            ElseIf TemporaryScreen = "Mainline Operation GUI Steam1 Screen" Then
                MainlineOperationGuiSteam1Screen.Show
            ElseIf TemporaryScreen = "Mainline Other Screen" Then
                MainlineOther.Show
            ElseIf TemporaryScreen = "Mainline Prototype Info Screen" Then
                MainlinePrototypeInfo.Show
            ElseIf TemporaryScreen = "Mainline Rolling Stock Screen" Then
                MainlineRollingStock.Show
            ElseIf TemporaryScreen = "Mainline Scale Speed Operation Screen" Then
                MainlineScaleSpeedOperation.Show
            ElseIf TemporaryScreen = "Mainline Scale Speed Setting Screen" Then
                MainlineScaleSpeedSetting.Show
            ElseIf TemporaryScreen = "Mainline Speed Table Screen" Then
                MainlineSpeedTable.Show
            ElseIf TemporaryScreen = "Mainline Steam Screen" Then
                MainlineSteam.Show
            ElseIf TemporaryScreen = "Main Screen" Then
                MainScreen.Show
            ElseIf TemporaryScreen = "Opening Screen" Then
                OpeningScreen.Show
            ElseIf TemporaryScreen = "Password Screen" Then
                Password.Show
            ElseIf TemporaryScreen = "Programming Decoder Screen" Then
                ProgrammingDecoder.Show
            ElseIf TemporaryScreen = "Programming Diesel Screen" Then
                ProgrammingDiesel.Show
            ElseIf TemporaryScreen = "Programming Easy Screen Configuration Screen" Then
                ProgrammingEasyScreenConfiguration.Show
            ElseIf TemporaryScreen = "Programming Easy Screen Consist Functions Screen" Then
                ProgrammingEasyScreenConsistFunctions.Show
            ElseIf TemporaryScreen = "Programming Easy Screen Functions Screen" Then
                ProgrammingEasyScreenFunctions.Show
            ElseIf TemporaryScreen = "Programming Easy Screen Specific CVs Screen" Then
                ProgrammingEasyScreenSpecificCvs.Show
            ElseIf TemporaryScreen = "Programming Easy Screen Speed Table Screen" Then
                ProgrammingEasyScreenSpeedTable.Show
            ElseIf TemporaryScreen = "Programming Other Screen" Then
                ProgrammingOther.Show
            ElseIf TemporaryScreen = "Programming Prototype Info Screen" Then
                ProgrammingPrototypeInfo.Show
            ElseIf TemporaryScreen = "Programming Rolling Stock Screen" Then
                ProgrammingRollingStock.Show
            ElseIf TemporaryScreen = "Programming Speed Table Screen" Then
                ProgrammingSpeedTable.Show
            ElseIf TemporaryScreen = "Programming Steam Screen" Then
                ProgrammingSteam.Show
            ElseIf TemporaryScreen = "Room Lighting Control Screen" Then
                RoomLightingControl.Show
            ElseIf TemporaryScreen = "Screen Attribute Setting Screen" Then
                ScreenAttributeSetting.Show
            ElseIf TemporaryScreen = "Sound Screen" Then
                SoundScreen.Show
            ElseIf TemporaryScreen = "Sound Screen Edit List Screen" Then
                SoundScreenEditList.Show
            ElseIf TemporaryScreen = "System Information Screen" Then
                SystemInformation.Show
            ElseIf TemporaryScreen = "Update Software Screen" Then
                UpdateSoftware.Show vbModeless
            ElseIf TemporaryScreen = "Update Software Screen" Then
                UpdateSoftware.Show vbModeless
            ElseIf temporaryscereen = "Utilities for Command Control" Then
                UtilitiesForCommandControl.Show vbModeless
            ElseIf TemporaryScreen = "Video Settings Screen" Then
                VideoSettings.Show vbModeless
            ElseIf MainScreen!MenuLogFile.Caption = "&Log File is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "About Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "About Screen, Button Close, stack empty, underflow."
        End If
    End If

End Sub


Private Sub ButtonPrint_Click()

    UpdateSoftware.PrintForm

End Sub

Private Sub ButtonUpdateLocomoticeDatabaseConsist_Click()

End Sub

Private Sub Form_Activate()

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
        If TemporaryScreen = "About Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "About Screen"
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
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "About Screen, Form Activate, stack is full, overflow."
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
                For InsideLoop = 0 To Val(TemporaryScreenDelay)
                    DoEvents
                Next InsideLoop
            Next OutsideLoop
        ElseIf TemporaryTransparency = "Off" Then
           Let AlphaBlend.Enabled = False
        Else
            If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "About Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. "
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Update Software Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "About Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(UpdateSoftware.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(UpdateSoftware.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(UpdateSoftware.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(UpdateSoftware.Height)

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
                For InsideLoop = Val(TemporaryScreenDelay) To 0 Step -1
                    DoEvents
                Next InsideLoop
            Next OutsideLoop
        ElseIf TemporaryTransparency = "Off" Then
           Let AlphaBlend.Enabled = False
        Else
            If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "About Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "About Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    UpdateSoftware.Hide

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub

Private Sub Form_Load()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Do While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbRetry Then
            Load ScreenAttributeSetting
            ScreenAttributeSetting.Show vbModal
        ElseIf TemporaryResponse = vbCancel Then
            End
        End If
    Loop
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Initialization of Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Update Software Screen"
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
        UpdateSoftware.Left = (Screen.Width - Width) / 2
        UpdateSoftware.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + UpdateSoftware.Width > Screen.Width Then
            Let UpdateSoftware.Left = Screen.Width - UpdateSoftware.Width
        Else
            Let UpdateSoftware.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + UpdateSoftware.Height > Screen.Height Then
            Let UpdateSoftware.Top = Screen.Height - UpdateSoftware.Height
        Else
            Let UpdateSoftware.Top = Val(TemporaryValueTop)
        End If
    End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    ElseIf MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Integer
        Dim t As Boolean
        Dim f As Boolean
        Let t = True
        Let f = False

'        Let TemporaryText1 = "This highlighted text when 'click'ed will" + vbCrLf + "display my home page in your web browser."
'        Let TemporaryText2 = "Universal Resource Link"
'        i = BalloonHelp.DestroyToolTip(HyperLabelOcx1.hWnd)
'        i = BalloonHelp.AddToolTip(HyperLabelOcx1.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

'        Let TemporaryText1 = "This highlighted text when 'click'ed will" + vbCrLf + "display my email address in your email program."
'        Let TemporaryText2 = "Universal Resource Link"
'        i = BalloonHelp.DestroyToolTip(HyperLabelOcx2.hWnd)
'        i = BalloonHelp.AddToolTip(HyperLabelOcx2.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

'        Let TemporaryText1 = "This highlighted text when 'click'ed will" + vbCrLf + "display my yahoo group in your web browser."
'        Let TemporaryText2 = "Universal Resource Link"
'        i = BalloonHelp.DestroyToolTip(HyperLabelOcx3.hWnd)
'        i = BalloonHelp.AddToolTip(HyperLabelOcx3.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

        Let TemporaryText1 = "This button when 'click'ed on will" + vbCrLf + "print the current screen."
        Let TemporaryText2 = "Print Button"
        i = BalloonHelp.DestroyToolTip(ButtonPrint.hWnd)
        i = BalloonHelp.AddToolTip(ButtonPrint.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

        Let TemporaryText1 = "This button when 'click'ed on will" + vbCrLf + "close the About Window and return control" + vbCrLf + "to the main screen."
        Let TemporaryText2 = "Close Button"
        i = BalloonHelp.DestroyToolTip(ButtonClose.hWnd)
        i = BalloonHelp.AddToolTip(ButtonClose.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)
    Else
        i = BalloonHelp.DestroyAllToolTips
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


ListView.ColumnHeaders.Add , , "Name"
ListView.ColumnHeaders.Add , , "Size"
ListView.ColumnHeaders.Add , , "Date"
ListView.ColumnHeaders.Add , , "Flags"
ListView.ColumnHeaders.Add , , "User"
ListView.ColumnHeaders.Add , , "Group"

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


Private Sub NetConnect1_Connected()

End Sub


Private Sub FtpControl_EnumFile(ByVal Filename As String, ByVal FileSize As Double, ByVal User As String, ByVal Group As String, ByVal FileDate As Date, ByVal FileFlags As String)
        
    Dim l As ListItem
    'On Error Resume Next
    If Filename <> "." And Filename <> ".." Then
        Set l = ListView.ListItems.Add(, , Filename)
        l.SubItems(1) = Str(FileSize)
        l.SubItems(2) = FileDate
        l.SubItems(3) = FileFlags
        l.SubItems(4) = User
        l.SubItems(5) = Group
    End If
                
End Sub

Private Sub FtpControl_Timeout()

       If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Update Software Screen, FFtp COntrol, timed out."
        End If
        
End Sub


Private Sub FtpControl_TransferComplete(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal BytesTransfered As Long)
    
    Let LabelStatus.Caption = "File transfer protocol, last transfer is complete (" & BytesTransfered & " bytes)."

End Sub

Private Sub FtpControl_TransferProgress(ByVal Bytes As Long)
    
    Let LabelStatus.Caption = "File transfer protocol, transfer in progress, " & Bytes & " bytes recieved."

End Sub


Private Sub FtpControl_TransferStarting(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal Offset As Long)

    Let LabelStatus.Caption = "File transfer protocol, starting transfer."

End Sub


Private Sub NetConnect_DownloadComplete()

    Let TextBoxUpdate.Text = TextBoxUpdate.Text & "Download Complete." & Chr$(13)
    
End Sub


Private Sub NetConnect_DownloadFailed()

    Let TextBoxUpdate.Text = TextBoxUpdate.Text + "Download failed. Please contact the authour." & Chr$(13)

End Sub


Private Sub SysInfo1_ConfigChangeCancelled()

End Sub


