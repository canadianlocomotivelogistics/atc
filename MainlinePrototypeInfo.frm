VERSION 4.00
Begin VB.Form MainlinePrototypeInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Prototype Information"
   ClientHeight    =   7785
   ClientLeft      =   3045
   ClientTop       =   3960
   ClientWidth     =   6690
   Height          =   8190
   Icon            =   "MainlinePrototypeInfo.frx":0000
   Left            =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6690
   Top             =   3615
   Width           =   6810
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonAdoptInfo 
      Caption         =   "Adopt Info"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Data PrototypeInfoDatabase 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   -1  'True
      Height          =   270
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PrototypeInfo"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "MainlinePrototypeInfo.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox PrototypeUnitsBuilt 
      DataField       =   "PrototypeUnitsBuilt"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Text            =   "Units Built"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox PrototypeDateManufactured 
      DataField       =   "PrototypeDateManufactured"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Text            =   "Date Manufactured"
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox PrototypeCylinders 
      DataField       =   "PrototypeCylinders"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Text            =   "Cylinders"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox PrototypeLength 
      DataField       =   "PrototypeLength"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "Length"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Picture"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox PrototypeImageFilename 
      DataField       =   "PrototypeImageFilename"
      DataSource      =   "PrototypeInfoDatabase"
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox PrototypeDrawBarPull 
      DataField       =   "PrototypeDrawBarPull"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Text            =   "Draw Bar Pull"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox PrototypeTractionEffort 
      DataField       =   "PrototypeTractionEffort"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Text            =   "Traction Effort"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox PrototypeAdhesionFactor 
      DataField       =   "PrototypeAdhesionFactor"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Text            =   "Adhesion Factor"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox PrototypeWeight 
      DataField       =   "PrototypeWeight"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "Weight"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox PrototypeHorsePower 
      DataField       =   "PrototypeHorsePower"
      DataSource      =   "PrototypeInfoDatabase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "Horse Power"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox PrototypeFacts 
      Enabled         =   0   'False
      Height          =   5055
      Left            =   120
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "MainlinePrototypeInfo.frx":0884
      Top             =   2280
      Width           =   3015
   End
   Begin VB.ComboBox PrototypeModel 
      Height          =   315
      ItemData        =   "MainlinePrototypeInfo.frx":0896
      Left            =   120
      List            =   "MainlinePrototypeInfo.frx":0898
      TabIndex        =   3
      Text            =   "Prototype Model"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox PrototypeManufacturer 
      Height          =   315
      ItemData        =   "MainlinePrototypeInfo.frx":089A
      Left            =   120
      List            =   "MainlinePrototypeInfo.frx":08B6
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Prototype Manufacturer"
      Top             =   960
      Width           =   3015
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   7080
      Top             =   1980
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   7080
      Top             =   1440
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   7080
      Top             =   840
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label LabelDetails 
      Caption         =   "Locomotive Details"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   3360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label LabelFacts 
      Caption         =   "Locomotive Facts"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label10 
      Caption         =   "Units Built (number)"
      Height          =   195
      Left            =   4920
      TabIndex        =   27
      Top             =   6960
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "Date (year-year)"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Cylinders (number)"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Draw Bar Pull (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Traction Effort (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Adhesion (percentage)"
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Weight (lbs)"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Horse Power (hp)"
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Length (feet, Inches)"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "To search for a specific model:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label LabelPrototypeInfo 
      Caption         =   "Here to can look at the prototype information and place this information into your locomotive database."
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   120
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog PictureGet 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open an Image"
      Filter          =   "Picture Files (*.gif)|*.gif|Picture FIles (*.jpg)|*.jpg"
   End
   Begin VB.Image PrototypeImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "MainlinePrototypeInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub buttonAdoptInfo_Click()

Let MainlineDiesel!LocomotivePrototypeHorsePower.Text = PrototypeHorsePower.Text
Let MainlineDiesel!LocomotivePrototypeWeight.Text = PrototypeWeight.Text
Let MainlineDiesel!LocomotivePrototypeAdhesionFactor.Text = PrototypeAdhesionFactor.Text
Let MainlineDiesel!LocomotivePrototypeTractionEffort.Text = PrototypeTractionEffort.Text
Let MainlineDiesel!LocomotivePrototypeDrawBarPull.Text = PrototypeDrawBarPull.Text
Let MainlineDiesel!LocomotiveManufacturer.Text = PrototypeManufacturer.Text
Let MainlineDiesel!LocomotiveModel.Text = PrototypeModel.Text
Let MainlineDiesel!LocomotiveFacts.Text = PrototypeFacts.Text

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
            If TemporaryScreen = "Mainline Prototype Info Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Prototype Information Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Prototype Information Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Prototype Information Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonPrint_Click()

    MainlinePrototypeInfo.PrintForm
    
End Sub

Private Sub Command3_Click()

    PictureGet.ShowOpen
    Let PrototypeImageFilename.Text = PictureGet.Filename

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
        If TemporaryScreen = "Mainline Prototype Info Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Prototype Info Screen"
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
            Let Ini.Value = "Mainline Prototype Infomation Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Prototype Info Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Prototype Info Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Mainline Prototype Info Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlinePrototypeInfo.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlinePrototypeInfo.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlinePrototypeInfo.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlinePrototypeInfo.Height)

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
                Let Ini.Value = "Mainline Prototype Info Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Prototype Info Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlinePrototypeInfo.Hide
    'unload Mainlineprototypeinfo

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
    Let Ini.Application = "Mainline Prototype Info Screen"
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
        MainlinePrototypeInfo.Left = (Screen.Width - Width) / 2
        MainlinePrototypeInfo.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlinePrototypeInfo.Width > Screen.Width Then
            Let MainlinePrototypeInfo.Left = Screen.Width - MainlinePrototypeInfo.Width
        Else
            Let MainlinePrototypeInfo.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlinePrototypeInfo.Height > Screen.Height Then
            Let MainlinePrototypeInfo.Top = Screen.Height - MainlinePrototypeInfo.Height
        Else
            Let MainlinePrototypeInfo.Top = Val(TemporaryValueTop)
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

    Let PrototypeInfoDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabasePrototypeInfo.mdb"
    PrototypeInfoDatabase.Refresh

    If MainlineDiesel.LocomotiveManufacturer.Text <> "Locomotive Manufacturer" Then
        Let PrototypeManufacturer.Text = MainlineDiesel!LocomotiveManufacturer.Text
    End If
    If MainlineDiesel.LocomotiveModel.Text <> "Locomotive Model" Then
        Let PrototypeModel.Text = MainlineDiesel!LocomotiveModel.Text
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



Private Sub Form_Resize()

    If MainlinePrototypeInfo.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlinePrototypeInfo.WindowState = vbNormal Then
    
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
    Let Ini.Application = "Mainline Prototype Info Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(MainlinePrototypeInfo.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(MainlinePrototypeInfo.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(MainlinePrototypeInfo.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(MainlinePrototypeInfo.Height)
 
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

Private Sub PrototypeImageFilename_Change()

' What if there is No Picture?
'
' If you don't use an On Error statement, any run-time error that occurs is fatal; that is, an error message is displayed
' and execution stops.An "enabled" error handler is one that has been turned on by an On Error statement; an "active"
' error handler is an enabled handler that is in the process of handling an error.  If an error occurs while an error
' handler is active (between the occurrence of the error and a Resume, Exit Sub, Exit Function, or Exit Property
' statement), the current procedure's error handler can't handle the error.  Control returns to the Calling procedure;
' if the Calling procedure has an enabled error handler, it is activated to handle the error.  If the Calling
' procedure's error handler is also active, control passes back through previous Calling procedures until an enabled,
' but inactive, error handler is found.  If no inactive, enabled error handler is found, the error is fatal at the
' point at which it actually occurred.  Each time the error handler passes control back to the Calling procedure,
' that procedure becomes the current procedure.  Once an error is handled by an error handler in any procedure, ex
' current procedure at the point designated by the Resume statement.
'
' Note   An error-handling routine is not a Sub or Function procedure.  It is a section of code marked by a line label
' or line number.
'
' Error-handling routines rely on the value in the Err object's Number property to determine the cause of the error.
' The error-handling routine should test or save relevant Err object property  values before any other error can occur
' or before a procedure that could cause an error is Called.  The values in the Err object's properties reflect only the
' most recent error.  The error message associated with Err.Number is contained in Err.Description.
' On Error Resume Next causes execution to continue with the statement immediately following the statement that caused
' the run-time error, or with the statement immediately following the most recent Call out of the procedure containing
' the On Error Resume Next statement. This allows execution to continue despite a run-time error.  You can then build
' the error-handling routine inline within the procedure, rather than transfer control to another location within the
' procedure.  An On Error Resume Next statement becomes inactive when another procedure is Called, so you should execute
' an On Error Resume Next statement in each Called routine if you want inline error handling within that routine.
'
' Note   The On Error Resume Next construct may be preferable to On Error GoTo when dealing with errors generated
' during access to other objects, since it permits unambiguous identification of the object whose error code is being
' returned.  Checking Err after each interaction with an object removes ambiguity about which object your code was
' accessing when the error occurred because the context is immediate.  Thus, you can be sure of which object placed the
' error code in Err.Number, as well as which object originally generated the error (the one specified in Err.Source).

    On Error Resume Next

    Let MainlinePrototypeInfo!PROTOTYPEIMAGE.Picture = LoadPicture(PrototypeImageFilename.Text)

    If Err.Number = 53 Then

'Displays a message in a dialog box, waits for the user to choose a button, and returns a value indicating which button
' the user has chosen.

        MainlinePrototypeInfo!PROTOTYPEIMAGE.Picture = LoadPicture()
        'MsgBox "Your picture listed on file was not found." + Chr$(13) + "Please update this record.", vbExclamation, "Locomotive Picture not Found"

    End If

' -------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub



Private Sub PrototypeInfoDatabase_Reposition()
Let PrototypeInfoDatabase.Caption = PrototypeInfoDatabase.Recordset.AbsolutePosition
End Sub

Private Sub PrototypeManufacturer_Change()

If PrototypeManufacturer.Text = "American Locomotive Company" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "C-636"
    PrototypeModel.AddItem "C-630"
    PrototypeModel.AddItem "C-628"
    PrototypeModel.AddItem "C-430"
    PrototypeModel.AddItem "C-425"
    PrototypeModel.AddItem "C-424"
    PrototypeModel.AddItem "C-420"
    PrototypeModel.AddItem "RS36"
    PrototypeModel.AddItem "RS32"
    PrototypeModel.AddItem "RS27"
    PrototypeModel.AddItem "RSD15"
    PrototypeModel.AddItem "RSD12"
    PrototypeModel.AddItem "RD11"
    PrototypeModel.AddItem "RS3"
    PrototypeModel.AddItem "RS2"
    PrototypeModel.AddItem "RS1"
    PrototypeModel.AddItem "S-1"
    PrototypeModel.AddItem "S-2"
    PrototypeModel.AddItem "S-3"
    PrototypeModel.AddItem "S-4"
    PrototypeModel.AddItem "S-6"
    PrototypeModel.AddItem "T-6"
    PrototypeModel.AddItem "PA-3"
    PrototypeModel.AddItem "PA-2"
    PrototypeModel.AddItem "PA-1"
    PrototypeModel.AddItem "FA-2"
    PrototypeModel.AddItem "FA-1"
End If

If PrototypeManufacturer.Text = "Baldwin" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "VO 1000"
    PrototypeModel.AddItem "DRS-6-6-15 or AS-616"
    PrototypeModel.AddItem "DRS-6-4-15 or AS-416"
    PrototypeModel.AddItem "DRS-4-4-15 or AS-16"
    PrototypeModel.AddItem "DS-4-4-10"
    PrototypeModel.AddItem "S-12"
End If

If PrototypeManufacturer.Text = "Bombardier" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "HR616"
    End If
    
If PrototypeManufacturer.Text = "FairBanks-Morse" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "H-12-44"
End If

If PrototypeManufacturer.Text = "General Electric" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "AC6000"
    PrototypeModel.AddItem "AC4400CW"
    PrototypeModel.AddItem "9-44CW"
    PrototypeModel.AddItem "9-44C"
    PrototypeModel.AddItem "9-44BW"
    PrototypeModel.AddItem "9-44B"
    PrototypeModel.AddItem "9-40CW"
    PrototypeModel.AddItem "9-40C"
    PrototypeModel.AddItem "9-40BW"
    PrototypeModel.AddItem "9-40B"
    PrototypeModel.AddItem "8-41CW"
    PrototypeModel.AddItem "8-41C"
    PrototypeModel.AddItem "8-40CW"
    PrototypeModel.AddItem "8-40CM"
    PrototypeModel.AddItem "8-40C"
    PrototypeModel.AddItem "8-40BW"
    PrototypeModel.AddItem "8-40B"
    PrototypeModel.AddItem "8-39CE"
    PrototypeModel.AddItem "8-39C"
    PrototypeModel.AddItem "8-39B"
    PrototypeModel.AddItem "8-32C"
    PrototypeModel.AddItem "8-32BWH or P32-8BWH"
    PrototypeModel.AddItem "8-32B"
    PrototypeModel.AddItem "7-36B"
    PrototypeModel.AddItem "7-36C"
    PrototypeModel.AddItem "7-33C"
    PrototypeModel.AddItem "7-30B-A1"
    PrototypeModel.AddItem "7-30C-A"
    PrototypeModel.AddItem "7-30B-A"
    PrototypeModel.AddItem "7-30C"
    PrototypeModel.AddItem "7-30B"
    PrototypeModel.AddItem "7-23BQ"
    PrototypeModel.AddItem "7-23B"
    PrototypeModel.AddItem "U36C"
    PrototypeModel.AddItem "U36B"
    PrototypeModel.AddItem "U33C"
    PrototypeModel.AddItem "U33B"
    PrototypeModel.AddItem "SF30C"
    PrototypeModel.AddItem "U30C"
    PrototypeModel.AddItem "U30B"
    PrototypeModel.AddItem "U28B"
    PrototypeModel.AddItem "U25C"
    PrototypeModel.AddItem "U25B"
    PrototypeModel.AddItem "U23C"
    PrototypeModel.AddItem "U23B"
    PrototypeModel.AddItem "U18B"
End If

If PrototypeManufacturer.Text = "General Motors - Electromotive Division" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "SD90/43MAC"
    PrototypeModel.AddItem "SD80MAC"
    PrototypeModel.AddItem "SD80MC"
    PrototypeModel.AddItem "SD75M"
    PrototypeModel.AddItem "SD70MAC"
    PrototypeModel.AddItem "SD70M"
    PrototypeModel.AddItem "SD70I"
    PrototypeModel.AddItem "SD70"
    PrototypeModel.AddItem "F69PH-AC"
    PrototypeModel.AddItem "SD60MAC"
    PrototypeModel.AddItem "SD60M"
    PrototypeModel.AddItem "SD60F"
    PrototypeModel.AddItem "SD60I"
    PrototypeModel.AddItem "SD60"
    PrototypeModel.AddItem "SD50M"
    PrototypeModel.AddItem "SD50F"
    PrototypeModel.AddItem "SD50I"
    PrototypeModel.AddItem "SD50"
    PrototypeModel.AddItem "GP60M"
    PrototypeModel.AddItem "GP60"
    PrototypeModel.AddItem "AMD-103"
    PrototypeModel.AddItem "F59PHI"
    PrototypeModel.AddItem "F59PH"
    PrototypeModel.AddItem "GP59"
    PrototypeModel.AddItem "GP50"
    PrototypeModel.AddItem "DD40AX"
    PrototypeModel.AddItem "SD45-2T"
    PrototypeModel.AddItem "SD45-2"
    PrototypeModel.AddItem "SDP45"
    PrototypeModel.AddItem "FP45"
    PrototypeModel.AddItem "F45"
    PrototypeModel.AddItem "SD45"
    PrototypeModel.AddItem "SD40-2W"
    PrototypeModel.AddItem "SD40-2T"
    PrototypeModel.AddItem "SD40-2F"
    PrototypeModel.AddItem "SD40-2"
    PrototypeModel.AddItem "SDP40F"
    PrototypeModel.AddItem "SDP40"
    PrototypeModel.AddItem "F40PHM-2"
    PrototypeModel.AddItem "F40PH-2C"
    PrototypeModel.AddItem "F40PH"
    PrototypeModel.AddItem "F40C"
    PrototypeModel.AddItem "SD40"
    PrototypeModel.AddItem "GP40-2"
    PrototypeModel.AddItem "GP40X"
    PrototypeModel.AddItem "GP40P"
    PrototypeModel.AddItem "GP40W"
    PrototypeModel.AddItem "GP40"
    PrototypeModel.AddItem "GP39-2"
    PrototypeModel.AddItem "SDL39"
    PrototypeModel.AddItem "SD39"
    PrototypeModel.AddItem "GP39"
    PrototypeModel.AddItem "SD38-2"
    PrototypeModel.AddItem "GP39-2W"
    PrototypeModel.AddItem "GP38-2"
    PrototypeModel.AddItem "SD38"
    PrototypeModel.AddItem "GP38"
    PrototypeModel.AddItem "SDP35"
    PrototypeModel.AddItem "SD35"
    PrototypeModel.AddItem "GP35"
    PrototypeModel.AddItem "GP30"
    PrototypeModel.AddItem "SD24"
    PrototypeModel.AddItem "GP20"
    PrototypeModel.AddItem "GP15-1"
    PrototypeModel.AddItem "GP15T"
    PrototypeModel.AddItem "GP15"
    PrototypeModel.AddItem "SD18"
    PrototypeModel.AddItem "GP18"
    PrototypeModel.AddItem "SD9"
    PrototypeModel.AddItem "GP9"
    PrototypeModel.AddItem "SD7"
    PrototypeModel.AddItem "GP7"
    PrototypeModel.AddItem "CF7"
    PrototypeModel.AddItem "GMD1"
    PrototypeModel.AddItem "RS1325"
    PrototypeModel.AddItem "NM5"
    PrototypeModel.AddItem "NW2"
    PrototypeModel.AddItem "TR5"
    PrototypeModel.AddItem "TR4"
    PrototypeModel.AddItem "MP15T"
    PrototypeModel.AddItem "MP15AC"
    PrototypeModel.AddItem "MP15(DC)"
    PrototypeModel.AddItem "SW1504"
    PrototypeModel.AddItem "SW1500"
    PrototypeModel.AddItem "SW1200"
    PrototypeModel.AddItem "SW1001"
    PrototypeModel.AddItem "SW1000"
    PrototypeModel.AddItem "SW900"
    PrototypeModel.AddItem "SW600"
    PrototypeModel.AddItem "SW9"
    PrototypeModel.AddItem "SW8"
    PrototypeModel.AddItem "SW7"
    PrototypeModel.AddItem "SW1"
    PrototypeModel.AddItem "FL9"
    PrototypeModel.AddItem "FP9"
    PrototypeModel.AddItem "F9"
    PrototypeModel.AddItem "FP7"
    PrototypeModel.AddItem "F7"
    PrototypeModel.AddItem "F3"
    PrototypeModel.AddItem "E9A"
    PrototypeModel.AddItem "E8A"
    PrototypeModel.AddItem "BL2"

End If

If PrototypeManufacturer.Text = "Montreal Locomotive Works" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "S-13"
        PrototypeModel.AddItem "M636"
        PrototypeModel.AddItem "M630"
        PrototypeModel.AddItem "M420"
        PrototypeModel.AddItem "M420R"
End If

If PrototypeManufacturer.Text = "Morrison Knudsen" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "MK5000C"
        PrototypeModel.AddItem "MK1200G"
        PrototypeModel.AddItem "MK-F40PHL-2"
        PrototypeModel.AddItem "MKGP40FH-2"
End If

Let PrototypeManufacturer.SelStart = 0
Let PrototypeManufacturer.SelLength = 0

End Sub

Private Sub PrototypeManufacturer_Click()

If PrototypeManufacturer.Text = "American Locomotive Company" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "C-636"
    PrototypeModel.AddItem "C-630"
    PrototypeModel.AddItem "C-628"
    PrototypeModel.AddItem "C-430"
    PrototypeModel.AddItem "C-425"
    PrototypeModel.AddItem "C-424"
    PrototypeModel.AddItem "C-420"
    PrototypeModel.AddItem "RS36"
    PrototypeModel.AddItem "RS32"
    PrototypeModel.AddItem "RS27"
    PrototypeModel.AddItem "RSD15"
    PrototypeModel.AddItem "RSD12"
    PrototypeModel.AddItem "RD11"
    PrototypeModel.AddItem "RS3"
    PrototypeModel.AddItem "RS2"
    PrototypeModel.AddItem "RS1"
    PrototypeModel.AddItem "S-1"
    PrototypeModel.AddItem "S-2"
    PrototypeModel.AddItem "S-3"
    PrototypeModel.AddItem "S-4"
    PrototypeModel.AddItem "S-6"
    PrototypeModel.AddItem "T-6"
    PrototypeModel.AddItem "PA-3"
    PrototypeModel.AddItem "PA-2"
    PrototypeModel.AddItem "PA-1"
    PrototypeModel.AddItem "FA-2"
    PrototypeModel.AddItem "FA-1"
End If

If PrototypeManufacturer.Text = "Baldwin" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "VO 1000"
    PrototypeModel.AddItem "DRS-6-6-15 or AS-616"
    PrototypeModel.AddItem "DRS-6-4-15 or AS-416"
    PrototypeModel.AddItem "DRS-4-4-15 or AS-16"
    PrototypeModel.AddItem "DS-4-4-10"
    PrototypeModel.AddItem "S-12"
End If

If PrototypeManufacturer.Text = "Bombardier" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "HR616"
    End If
    
If PrototypeManufacturer.Text = "FairBanks-Morse" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "H-12-44"
End If

If PrototypeManufacturer.Text = "General Electric" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "AC6000"
    PrototypeModel.AddItem "AC4400CW"
    PrototypeModel.AddItem "9-44CW"
    PrototypeModel.AddItem "9-44C"
    PrototypeModel.AddItem "9-44BW"
    PrototypeModel.AddItem "9-44B"
    PrototypeModel.AddItem "9-40CW"
    PrototypeModel.AddItem "9-40C"
    PrototypeModel.AddItem "9-40BW"
    PrototypeModel.AddItem "9-40B"
    PrototypeModel.AddItem "8-41CW"
    PrototypeModel.AddItem "8-41C"
    PrototypeModel.AddItem "8-40CW"
    PrototypeModel.AddItem "8-40CM"
    PrototypeModel.AddItem "8-40C"
    PrototypeModel.AddItem "8-40BW"
    PrototypeModel.AddItem "8-40B"
    PrototypeModel.AddItem "8-39CE"
    PrototypeModel.AddItem "8-39C"
    PrototypeModel.AddItem "8-39B"
    PrototypeModel.AddItem "8-32C"
    PrototypeModel.AddItem "8-32BWH or P32-8BWH"
    PrototypeModel.AddItem "8-32B"
    PrototypeModel.AddItem "7-36B"
    PrototypeModel.AddItem "7-36C"
    PrototypeModel.AddItem "7-33C"
    PrototypeModel.AddItem "7-30B-A1"
    PrototypeModel.AddItem "7-30C-A"
    PrototypeModel.AddItem "7-30B-A"
    PrototypeModel.AddItem "7-30C"
    PrototypeModel.AddItem "7-30B"
    PrototypeModel.AddItem "7-23BQ"
    PrototypeModel.AddItem "7-23B"
    PrototypeModel.AddItem "U36C"
    PrototypeModel.AddItem "U36B"
    PrototypeModel.AddItem "U33C"
    PrototypeModel.AddItem "U33B"
    PrototypeModel.AddItem "SF30C"
    PrototypeModel.AddItem "U30C"
    PrototypeModel.AddItem "U30B"
    PrototypeModel.AddItem "U28B"
    PrototypeModel.AddItem "U25C"
    PrototypeModel.AddItem "U25B"
    PrototypeModel.AddItem "U23C"
    PrototypeModel.AddItem "U23B"
    PrototypeModel.AddItem "U18B"
End If

If PrototypeManufacturer.Text = "General Motors - Electromotive Division" Then
    PrototypeModel.Clear
    PrototypeModel.AddItem "SD90/43MAC"
    PrototypeModel.AddItem "SD80MAC"
    PrototypeModel.AddItem "SD80MC"
    PrototypeModel.AddItem "SD75M"
    PrototypeModel.AddItem "SD70MAC"
    PrototypeModel.AddItem "SD70M"
    PrototypeModel.AddItem "SD70I"
    PrototypeModel.AddItem "SD70"
    PrototypeModel.AddItem "F69PH-AC"
    PrototypeModel.AddItem "SD60MAC"
    PrototypeModel.AddItem "SD60M"
    PrototypeModel.AddItem "SD60F"
    PrototypeModel.AddItem "SD60I"
    PrototypeModel.AddItem "SD60"
    PrototypeModel.AddItem "SD50M"
    PrototypeModel.AddItem "SD50F"
    PrototypeModel.AddItem "SD50I"
    PrototypeModel.AddItem "SD50"
    PrototypeModel.AddItem "GP60M"
    PrototypeModel.AddItem "GP60"
    PrototypeModel.AddItem "AMD-103"
    PrototypeModel.AddItem "F59PHI"
    PrototypeModel.AddItem "F59PH"
    PrototypeModel.AddItem "GP59"
    PrototypeModel.AddItem "GP50"
    PrototypeModel.AddItem "DD40AX"
    PrototypeModel.AddItem "SD45-2T"
    PrototypeModel.AddItem "SD45-2"
    PrototypeModel.AddItem "SDP45"
    PrototypeModel.AddItem "FP45"
    PrototypeModel.AddItem "F45"
    PrototypeModel.AddItem "SD45"
    PrototypeModel.AddItem "SD40-2W"
    PrototypeModel.AddItem "SD40-2T"
    PrototypeModel.AddItem "SD40-2F"
    PrototypeModel.AddItem "SD40-2"
    PrototypeModel.AddItem "SDP40F"
    PrototypeModel.AddItem "SDP40"
    PrototypeModel.AddItem "F40PHM-2"
    PrototypeModel.AddItem "F40PH-2C"
    PrototypeModel.AddItem "F40PH"
    PrototypeModel.AddItem "F40C"
    PrototypeModel.AddItem "SD40"
    PrototypeModel.AddItem "GP40-2"
    PrototypeModel.AddItem "GP40X"
    PrototypeModel.AddItem "GP40P"
    PrototypeModel.AddItem "GP40W"
    PrototypeModel.AddItem "GP40"
    PrototypeModel.AddItem "GP39-2"
    PrototypeModel.AddItem "SDL39"
    PrototypeModel.AddItem "SD39"
    PrototypeModel.AddItem "GP39"
    PrototypeModel.AddItem "SD38-2"
    PrototypeModel.AddItem "GP39-2W"
    PrototypeModel.AddItem "GP38-2"
    PrototypeModel.AddItem "SD38"
    PrototypeModel.AddItem "GP38"
    PrototypeModel.AddItem "SDP35"
    PrototypeModel.AddItem "SD35"
    PrototypeModel.AddItem "GP35"
    PrototypeModel.AddItem "GP30"
    PrototypeModel.AddItem "SD24"
    PrototypeModel.AddItem "GP20"
    PrototypeModel.AddItem "GP15-1"
    PrototypeModel.AddItem "GP15T"
    PrototypeModel.AddItem "GP15"
    PrototypeModel.AddItem "SD18"
    PrototypeModel.AddItem "GP18"
    PrototypeModel.AddItem "SD9"
    PrototypeModel.AddItem "GP9"
    PrototypeModel.AddItem "SD7"
    PrototypeModel.AddItem "GP7"
    PrototypeModel.AddItem "CF7"
    PrototypeModel.AddItem "GMD1"
    PrototypeModel.AddItem "RS1325"
    PrototypeModel.AddItem "NM5"
    PrototypeModel.AddItem "NW2"
    PrototypeModel.AddItem "TR5"
    PrototypeModel.AddItem "TR4"
    PrototypeModel.AddItem "MP15T"
    PrototypeModel.AddItem "MP15AC"
    PrototypeModel.AddItem "MP15(DC)"
    PrototypeModel.AddItem "SW1504"
    PrototypeModel.AddItem "SW1500"
    PrototypeModel.AddItem "SW1200"
    PrototypeModel.AddItem "SW1001"
    PrototypeModel.AddItem "SW1000"
    PrototypeModel.AddItem "SW900"
    PrototypeModel.AddItem "SW600"
    PrototypeModel.AddItem "SW9"
    PrototypeModel.AddItem "SW8"
    PrototypeModel.AddItem "SW7"
    PrototypeModel.AddItem "SW1"
    PrototypeModel.AddItem "FL9"
    PrototypeModel.AddItem "FP9"
    PrototypeModel.AddItem "F9"
    PrototypeModel.AddItem "FP7"
    PrototypeModel.AddItem "F7"
    PrototypeModel.AddItem "F3"
    PrototypeModel.AddItem "E9A"
    PrototypeModel.AddItem "E8A"
    PrototypeModel.AddItem "BL2"

End If

If PrototypeManufacturer.Text = "Montreal Locomotive Works" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "S-13"
        PrototypeModel.AddItem "M636"
        PrototypeModel.AddItem "M630"
        PrototypeModel.AddItem "M420"
        PrototypeModel.AddItem "M420R"
End If

If PrototypeManufacturer.Text = "Morrison Knudsen" Then
        PrototypeModel.Clear
        PrototypeModel.AddItem "MK5000C"
        PrototypeModel.AddItem "MK1200G"
        PrototypeModel.AddItem "MK-F40PHL-2"
        PrototypeModel.AddItem "MKGP40FH-2"
End If

Let PrototypeManufacturer.SelStart = 0
Let PrototypeManufacturer.SelLength = 0

End Sub


Private Sub PrototypeModel_Change()

    Let MainlinePrototypeInfo.MousePointer = vbArrowHourglass
    
    Let ButtonClose.Enabled = False
    Let buttonAdoptInfo.Enabled = False

'Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

    PrototypeInfoDatabase.Recordset.MoveFirst

' Use the RecordCount property to find out how many records in a Recordset or TableDef object have been accessed.
' RecordCount doesn't indicate how many records are contained in a dynaset- or snapshot-type Recordset until all records have been accessed.
' Once the last record has been accessed, the RecordCount property indicates the total number of undeleted records in the Recordset or TableDef.
' To force the last record to be accessed, use the MoveLast or FindLast method on the Recordset.
' You can also use an SQL Count function to determine the approximate number of records your query will return.

    Do While Not PrototypeInfoDatabase.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    PrototypeInfoDatabase.Recordset.MoveNext
    
    If Not PrototypeInfoDatabase.Recordset.EOF Then
    
        If PrototypeInfoDatabase.Recordset.Fields("PrototypeModel") = PrototypeModel.Text Then
            Let PrototypeFacts.Text = PrototypeInfoDatabase.Recordset.Fields("PrototypeFacts")
            Exit Do
        End If
    End If

    ' Just in cause the user of the program closes the window while the database is still open searching
    ' This will terminate the search and allow the window to be close withou gerating and error.
         
    Loop
    
    If PrototypeInfoDatabase.Recordset.AbsolutePosition = -1 Then
        Let PrototypeFacts.Text = "There is no data available for this type of locomotive. Please email the author to have specific information added to this database for your roster."
        Let buttonAdoptInfo.Enabled = False
    Else
        Let buttonAdoptInfo.Enabled = True
    End If
    
    Let PrototypeModel.SelStart = 0
    Let PrototypeModel.SelLength = 0
    Let PrototypeFacts.SelStart = 0
    Let PrototypeFacts.SelLength = 0
    
    Let ButtonClose.Enabled = True
    
    Let MainlinePrototypeInfo.MousePointer = vbDefault


End Sub

Private Sub PrototypeModel_Click()

    Let MainlinePrototypeInfo.MousePointer = vbArrowHourglass
    
    Let ButtonClose.Enabled = False
    Let buttonAdoptInfo.Enabled = False

'Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

    PrototypeInfoDatabase.Recordset.MoveFirst

' Use the RecordCount property to find out how many records in a Recordset or TableDef object have been accessed.
' RecordCount doesn't indicate how many records are contained in a dynaset- or snapshot-type Recordset until all records have been accessed.
' Once the last record has been accessed, the RecordCount property indicates the total number of undeleted records in the Recordset or TableDef.
' To force the last record to be accessed, use the MoveLast or FindLast method on the Recordset.
' You can also use an SQL Count function to determine the approximate number of records your query will return.

    Do While Not PrototypeInfoDatabase.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    PrototypeInfoDatabase.Recordset.MoveNext
    
    If Not PrototypeInfoDatabase.Recordset.EOF Then
    
        If PrototypeInfoDatabase.Recordset.Fields("PrototypeModel") = PrototypeModel.Text Then
            Let PrototypeFacts.Text = PrototypeInfoDatabase.Recordset.Fields("PrototypeFacts")
            Exit Do
        End If
    End If

    ' Just in cause the user of the program closes the window while the database is still open searching
    ' This will terminate the search and allow the window to be close withou gerating and error.
         
    Loop
    
    If PrototypeInfoDatabase.Recordset.AbsolutePosition = -1 Then
        Let PrototypeFacts.Text = "There is no data available for this type of locomotive. Please email the author to have specific information added to this database for your roster."
        Let buttonAdoptInfo.Enabled = False
    Else
        Let buttonAdoptInfo.Enabled = True
    End If
    
    Let PrototypeModel.SelStart = 0
    Let PrototypeModel.SelLength = 0
    Let PrototypeFacts.SelStart = 0
    Let PrototypeFacts.SelLength = 0
    
    Let ButtonClose.Enabled = True
    
    Let MainlinePrototypeInfo.MousePointer = vbDefault

End Sub


