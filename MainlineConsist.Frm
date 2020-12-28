VERSION 4.00
Begin VB.Form MainlineConsist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Mainline Mode - Consist Control"
   ClientHeight    =   5595
   ClientLeft      =   2145
   ClientTop       =   2085
   ClientWidth     =   11880
   Height          =   6000
   Icon            =   "MainlineConsist.frx":0000
   Left            =   2085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11880
   Top             =   1740
   Width           =   12000
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "List"
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   4990
      Width           =   1215
   End
   Begin VB.CommandButton ButtonBulkCV 
      Caption         =   "Bulk CV"
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   4990
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxAddressShort 
      Caption         =   "Short Address in"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   195
      Picture         =   "MainlineConsist.frx":0442
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   510
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveConsist"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox CheckBoxUniqueConsistFailed 
      Caption         =   "Unique Consist Failed"
      Height          =   255
      Left            =   14880
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox CheckBoxUniqueConsist 
      Caption         =   "Unique"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   5160
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox LocomotiveNumberConsist 
      Height          =   285
      Left            =   14880
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1515
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveConsistNumber 
      Height          =   285
      Left            =   14880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveConsistTableCol 
      Height          =   285
      Left            =   14880
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveConsistTableRow 
      Height          =   285
      Left            =   14880
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Width           =   1215
   End
   Begin VB.CheckBox LocomotiveDirectionReversed 
      Caption         =   "Reverse Direction with"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox LocomotiveDirectionNormal 
      Caption         =   "Forward or"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   10560
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Data LocomotiveDatabase 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBaseLong"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox LocomotiveConsistSelect 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Data LocomotiveDatabaseConsist 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveConsist"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   12180
      Top             =   2280
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   12120
      Top             =   1800
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   12120
      Top             =   1200
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin VB.Label Label2 
      Caption         =   $"MainlineConsist.frx":0884
      Height          =   615
      Left            =   945
      TabIndex        =   20
      Top             =   120
      Width           =   10830
   End
   Begin MSDBGrid.DBGrid LocomotiveConsistTable 
      Bindings        =   "MainlineConsist.frx":0A4D
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "MainlineConsist.frx":0A6F
      TabIndex        =   2
      Top             =   1200
      Width           =   11655
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Locomotive"
      Height          =   195
      Left            =   13950
      TabIndex        =   19
      Top             =   1590
      Width           =   825
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Consist Assigned"
      Height          =   195
      Left            =   13575
      TabIndex        =   17
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Label Lable12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Column"
      Height          =   195
      Left            =   14295
      TabIndex        =   15
      Top             =   780
      Width           =   525
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Row"
      Height          =   195
      Left            =   14490
      TabIndex        =   13
      Top             =   420
      Width           =   330
   End
   Begin VB.Label LabelStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status: Clear"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Locomotive to Add or Delete in a Consist"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   2880
   End
   Begin VB.Menu menuLoad 
      Caption         =   "Load"
      Visible         =   0   'False
      Begin VB.Menu menuLoadAssign 
         Caption         =   "Assign to Cab"
      End
      Begin VB.Menu menuLoadATC 
         Caption         =   "ATC with Consist"
      End
      Begin VB.Menu menuLoadGUI 
         Caption         =   "GUI with Consist"
      End
   End
   Begin VB.Menu menuLocomotive 
      Caption         =   "Locomotive"
      Visible         =   0   'False
      Begin VB.Menu menuLocomotiveAdd 
         Caption         =   "Add Selected Locomotive"
      End
      Begin VB.Menu menuLocomotiveDelete 
         Caption         =   "Delete Selected Locomotive"
      End
      Begin VB.Menu menuLocomotiveEdit 
         Caption         =   "Edit Selected Locomotive"
      End
   End
   Begin VB.Menu menuDirection 
      Caption         =   "MenuDirection"
      Visible         =   0   'False
      Begin VB.Menu menuDirectionForward 
         Caption         =   "Forward in Consist"
      End
      Begin VB.Menu menuDirectionReverse 
         Caption         =   "Reverse in Consist"
      End
   End
End
Attribute VB_Name = "MainlineConsist"
Attribute VB_Creatable = False
Attribute VB_Exposed = False











Private Sub ButtonBulkCV_Click()
    
    Load MainlineCVChanger
    MainlineCVChanger.Show vbModeless

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
            If TemporaryScreen = "Mainline Consist Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Mainline Consist Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Mainline Consist Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   End Sub




Private Sub ButtonList_Click()
    
Load LocomotiveSpreadsheet
LocomotiveSpreadsheet.Show vbModeless

End Sub



Private Sub ButtonPrint_Click()

    MainlineConsist.PrintForm
    
End Sub

Private Sub CheckAddressShort_Click()


End Sub




Private Sub CheckBoxAddressShort_Click()
    
    Let ButtonClose.Enabled = False
    Let MainlineDiesel.MousePointer = vbArrowHourglass
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the recordsource property
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If CheckBoxAddressShort.Value = vbChecked Then
        Let LocomotiveDatabase.RecordSource = "LocomotiveDatabaseShortV030002"
    Else
        Let LocomotiveDatabase.RecordSource = "LocomotiveDatabaseLongV030002"
    End If
    LocomotiveDatabase.Refresh
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the locomotive selection from database
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveConsistSelect.Clear
    LocomotiveDatabase.Recordset.MoveFirst
    Do
        If LocomotiveDatabase.Recordset.Fields("LocomotiveDecoderEquiped").Value = "True" Then
             LocomotiveConsistSelect.AddItem Format(LocomotiveDatabase.Recordset.Fields("LocomotiveNumber").Value, "0000")
        End If
        LocomotiveDatabase.Recordset.MoveNext
        DoEvents
    Loop While Not LocomotiveDatabase.Recordset.EOF
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Return properties to
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let ButtonClose.Enabled = True
    Let MainlineDiesel.MousePointer = vbDefault
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
        If TemporaryScreen = "Mainline Consist Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Mainline Consist Screen"
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
            Let Ini.Value = "Mainline Consist Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Mainline Consist Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Consist Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    Call CheckBoxAddressShort_Click
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unloading the Form
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Mainline Consist Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(LocomotiveCVSpreadsheet.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(LocomotiveCVSpreadsheet.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(LocomotiveCVSpreadsheet.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(LocomotiveCVSpreadsheet.Height)
    Let Ini.Parameter = "CheckBoxAddressShort"
    Let Ini.Value = Str(CheckBoxAddressShort.Value)
    Let Ini.Parameter = "LocomotiveDirectionNormal"
    Let Ini.Value = Str(LocomotiveDirectionNormal.Value)
    Let Ini.Parameter = "LocomotiveDirectionReverse"
    Let Ini.Value = Str(LocomotiveDirectionReversed.Value)
    Let Ini.Parameter = "CheckBoxUniqueConsist"
    Let Ini.Value = Str(CheckBoxUniqueConsist.Value)
    
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
                Let Ini.Value = "Mainline Consist Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Mainline Consist Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    MainlineConsist.Hide
    'unload MainlineConsist

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
    Let Ini.Application = "Mainline Consist Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "CheckBoxAddressShort"
    Let CheckBoxAddressShort.Value = Val(Ini.Value)
    Let Ini.Parameter = "LocomotiveDirectionNormal"
    Let LocomotiveDirectionNormal.Value = Val(Ini.Value)
    Let Ini.Parameter = "LocomotiveDirectionReverse"
    Let LocomotiveDirectionReversed.Value = Val(Ini.Value)
    Let Ini.Parameter = "CheckBoxUniqueConsist"
    Let CheckBoxUniqueConsist.Value = Val(Ini.Value)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        MainlineConsist.Left = (Screen.Width - Width) / 2
        MainlineConsist.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + MainlineConsist.Width > Screen.Width Then
            Let MainlineConsist.Left = Screen.Width - MainlineConsist.Width
        Else
            Let MainlineConsist.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + MainlineConsist.Height > Screen.Height Then
            Let MainlineConsist.Top = Screen.Height - MainlineConsist.Height
        Else
            Let MainlineConsist.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This database grid displays all consists with the" & vbCrLf & "locomotives assigned to it."
        Let TemporaryText2 = "Consist Spreadsheet"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveConsistTable)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveConsistTable, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This combonation box is used to select the" & vbCrLf & "locomotives in your database. You can then" & vbCrLf & "add or delete it form a consist by 'click'ing" & vbCrLf & "the grid location."
        Let TemporaryText2 = "Locomotive Number"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveConsistSelect)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveConsistSelect, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This checkbox is used to select the" & vbCrLf & "locomotives database for short addresses."
        Let TemporaryText2 = "Short Addresses"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxAddressShort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxAddressShort, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box indicates if the locomotive" & vbCrLf & "you have selected will be added to a consist in the" & vbCrLf & "forward or reverse direction."
        Let TemporaryText2 = "Forward Direction Check Box"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDirectionNormal)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDirectionNormal, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This check box indicates if the locomotive" & vbCrLf & "you have selected will be added to a consist in the" & vbCrLf & "forward or reverse direction."
        Let TemporaryText2 = "Reverse Direction Check Box"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(LocomotiveDirectionReversed)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(LocomotiveDirectionReversed, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This check box indicates if the locomotive" & vbCrLf & "you have selected is unique (one and only). The" & vbCrLf & "program will to validate if this locomotive was" & vbCrLf & "previously entered in a consist."
        Let TemporaryText2 = "Unique Check Box"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxUniqueConsist)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxUniqueConsist, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button will display the Locomotive" & vbCrLf & "Spreadsheet Window. You can look at your whole database" & vbCrLf & "(locomotive roster)"
        Let TemporaryText2 = "List Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonList, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button will display the Mainline" & vbCrLf & "Configuration Variable Changer Window. It is used to make" & vbCrLf & "changes to configuration variable to all locomotives on the" & vbCrLf & "mainline (eg. to change all locomtive's acceleration)."
        Let TemporaryText2 = "Bulk CV Changer Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonBulkCV)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonBulkCV, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button will close the Mainline Consist" & vbCrLf & "Window and return you to the previous screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    Else
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let LocomotiveDatabaseConsist.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseConsist.mdb"
    Let LocomotiveDatabase.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
    Let Data1.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseConsist.mdb"

' For the Columns object of the DBGrid control, Width is specified in the unit of measure of
' the object that contains the DBGrid.  The default value for Width is the value of the
' DefColWidth property of DBGrid. For the Picture object, use the ScaleX and ScaleY methods
' to convert HiMetric units into the scale you need.

    Let temp1 = 550
    Let temp2 = 400
    Let temp3 = 820
    Let temp4 = 295
    
' ---------------------------------------------------------------------------------------------
    
    Let LocomotiveConsistTable.Columns(0).Width = temp1
    Let LocomotiveConsistTable.Columns(1).Width = temp2
    Let LocomotiveConsistTable.Columns(2).Width = temp3
    Let LocomotiveConsistTable.Columns(3).Width = temp4
    Let LocomotiveConsistTable.Columns(4).Width = temp3
    Let LocomotiveConsistTable.Columns(5).Width = temp4
    Let LocomotiveConsistTable.Columns(6).Width = temp3
    Let LocomotiveConsistTable.Columns(7).Width = temp4
    Let LocomotiveConsistTable.Columns(8).Width = temp3
    Let LocomotiveConsistTable.Columns(9).Width = temp4
    Let LocomotiveConsistTable.Columns(10).Width = temp3
    Let LocomotiveConsistTable.Columns(11).Width = temp4
    Let LocomotiveConsistTable.Columns(12).Width = temp3
    Let LocomotiveConsistTable.Columns(13).Width = temp4
    Let LocomotiveConsistTable.Columns(14).Width = temp3
    Let LocomotiveConsistTable.Columns(15).Width = temp4
    Let LocomotiveConsistTable.Columns(16).Width = temp3
    Let LocomotiveConsistTable.Columns(17).Width = temp4
    Let LocomotiveConsistTable.Columns(18).Width = temp3
    Let LocomotiveConsistTable.Columns(19).Width = temp4
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



Private Sub Form_Resize()

    If MainlineConsist.WindowState = vbMinimized Then
    
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
        
    ElseIf MainlineConsist.WindowState = vbNormal Then
    
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

Private Sub LocomotiveConsistSelect_Change()

    Let CheckBoxUniqueConsist.Value = vbUnchecked

End Sub

Private Sub LocomotiveConsistTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Table Controls
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If LocomotiveConsistTable.RowContaining(Y) < 0 Then Exit Sub
    If LocomotiveConsistTable.ColContaining(X) < 0 Then Exit Sub
    Let LocomotiveConsistTableRow.Text = LocomotiveConsistTable.RowContaining(Y)
    Let LocomotiveConsistTableCol.Text = LocomotiveConsistTable.ColContaining(X)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Start of Changing Locomotive CV on the Mainline
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If LocomotiveConsistTableCol.Text = "0" Then MainlineConsist.PopupMenu menuLoad
    If LocomotiveConsistTableCol.Text = "2" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "3" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "4" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "5" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "6" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "7" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "8" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "9" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "10" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "11" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "12" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "13" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "14" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "15" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "16" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "17" Then MainlineConsist.PopupMenu menuDirection
    If LocomotiveConsistTableCol.Text = "18" Then MainlineConsist.PopupMenu menuLocomotive
    If LocomotiveConsistTableCol.Text = "19" Then MainlineConsist.PopupMenu menuDirection

End Sub

Private Sub LocomotiveDirectionReversed_Click()

If LocomotiveDirectionReversed = vbChecked Then
    Let LocomotiveDirectionNormal = vbUnchecked
Else
    Let LocomotiveDirectionNormal = vbChecked
End If

EXITSTAGELEFT:

End Sub




Private Sub LocomotiveDirectionNormal_Click()

If LocomotiveDirectionNormal = vbChecked Then
    LocomotiveDirectionReversed = vbUnchecked
Else
    LocomotiveDirectionReversed = vbChecked
End If


EXITSUB:

End Sub


Private Sub menuDirection_Click()

    Let LocomotiveConsistTable.Col = LocomotiveConsistTableCol.Text
    Let LocomotiveConsistTable.Row = LocomotiveConsistTableRow.Text
    
    
    If Val(LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text) - 1)) = 0 Then
        Let menuDirectionForward.Enabled = False
        Let menuDirectionReverse.Enabled = False
    Else
        Let menuDirectionForward.Enabled = True
        Let menuDirectionReverse.Enabled = True
    End If
    
End Sub

Private Sub menuDirectionForward_Click()
    
' -------------------------------------------------------------------------------------------------------------------------
' Setup command to Command Station for Consist Updates (edit direction locomotive)
' -------------------------------------------------------------------------------------------------------------------------
    Let LocomotiveCommunicationStatus.Caption = "Status: Changing Locomotive Direction to Foward in Consist"
    LocomotiveDatabaseConsist.Recordset.Edit
    Let MainScreen!onebyted.Text = Int(Val(LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text) - 1)) / 256)
    Let MainScreen!twoByteD.Text = Val(LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text) - 1)) - (Val(MainScreen!onebyted.Text) * 256)
    Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64 ' add first two bits to address one byte for long address
    Let MainScreen!ThreeByteD.Text = 16 ' instruction byte
    Let LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text)) = "F"
    Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 ' activate consist in forward direction
    Let MainScreen!FourByteD.Text = LocomotiveDatabaseConsist.Recordset.Fields(0) ' which is the consist number
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyte.Text = ""
    LocomotiveDatabaseConsist.Recordset.Update
' -------------------------------------------------------------------------------------------------------------------------
    Call MainScreen.SendCommandviaTrackQ
        
End Sub

Private Sub menuDirectionReverse_Click()

' -------------------------------------------------------------------------------------------------------------------------
' Setup command to Command Station for Consist Updates (edit direction locomotive)
' -------------------------------------------------------------------------------------------------------------------------
    Let LocomotiveCommunicationStatus.Caption = "Status: Changing Locomotive Direction in Consist"
    LocomotiveDatabaseConsist.Recordset.Edit
    Let MainScreen!onebyted.Text = Int(Val(LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text) - 1)) / 256)
    Let MainScreen!twoByteD.Text = Val(LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text) - 1)) - (Val(MainScreen!onebyted.Text) * 256)
    Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64 ' add first two bytes to address one byte for long address
    Let MainScreen!ThreeByteD.Text = 16 ' instruction byte
    Let LocomotiveDatabaseConsist.Recordset.Fields(Val(LocomotiveConsistTableCol.Text)) = "R"
    Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 + 1 ' activate consist in reverse direction
    Let MainScreen!FourByteD.Text = LocomotiveDatabaseConsist.Recordset.Fields(0) ' which is the consist number
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    LocomotiveDatabaseConsist.Recordset.Update
' -------------------------------------------------------------------------------------------------------------------------
    Call MainScreen.SendCommandviaTrackQ

End Sub

Private Sub menuLoad_Click()
   
    Let LocomotiveConsistTable.Col = Val(LocomotiveConsistTableCol.Text)
    Let LocomotiveConsistTable.Row = Val(LocomotiveConsistTableRow.Text)
 
If Val(LocomotiveDatabaseConsist.Recordset.Fields(2)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(4)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(6)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(8)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(10)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(12)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(14)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(16)) = 0 And _
   Val(LocomotiveDatabaseConsist.Recordset.Fields(18)) = 0 Then
    
    Let menuLoadAssign.Enabled = False
    Let menuLoadATC.Enabled = False
    Let menuLoadGUI.Enabled = False
Else
    Let menuLoadAssign.Enabled = True
    Let menuLoadATC.Enabled = True
    Let menuLoadGUI.Enabled = True
End If

End Sub

Private Sub menuLoadAssign_Click()
    
    Dim TemporaryLocomotiveCabAssignment As String
    Dim TemporaryTimeOutCounter As Long
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Assignment of a Consist
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'    Let LocomotiveConsistNumber.Text = LocomotiveDatabaseConsist.Recordset.Fields(0)
    Let temp = "What handheld cab would you like consist number " & Str$(LocomotiveDatabaseConsist.Recordset.Fields(0)) & Chr$(13) & "assigned to? From one to sixty-four are valid options."
    Let TemporaryLocomotiveCabAssignment = InputBox(temp, "Assignment of Consist", "2")
    If TemporaryLocomotiveCabAssignment = "" Or _
       Val(TemporaryLocomotiveCabAssignment) < 1 Or _
       Val(TemporaryLocomotiveCabAssignment) > 64 Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, time out occured by digital command control unit after sending command."
        End If
        Let LabelStatus.Caption = "Status: Command aborted due to error."
        Exit Sub
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelStatus.Caption = "Status: Assigning Consist to Cab"
    Let MainScreen!MSComm1.InBufferCount = 0
    Let MainScreen!MSComm1.Output = Chr$(&H81) & Chr$(0) & Chr$(Val(LocomotiveDatabaseConsist.Recordset.Fields(0))) & Chr$(Val(TemporaryLocomotiveCabAssignment))
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveDatabaseConsist.Recordset.Edit
    Let LocomotiveDatabaseConsist.Recordset.Fields(1) = TemporaryLocomotiveCabAssignment
    LocomotiveDatabaseConsist.Recordset.Update
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.Labelmode.Text = "Standard Mode" Then
        Do While MainScreen.MSComm1.InBufferCount < 1
            Let TemporaryTimeOutCounter = TemporaryTimeOutCounter + 1
            If TemporaryTimeOutCounter > 750000 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, time out occured by digital command control unit after sending command."
                End If
                Let LabelStatus.Caption = "Status: Command aborted due to error."
                Exit Sub
            End If
            Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
            DoEvents
        Loop
        Let TemporaryInput = MainScreen.MSComm1.Input
        If Len(TemporaryInput) <> 1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, wrong number of bytes returned from digital command control unit after sending command to cab screen."
            End If
                ElseIf TemporaryInput <> "!" Then
            If TemporaryInput = "0" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, command not supported by digital command control unit after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "2" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, cab address out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "3" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, data out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "4" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, byte count out of range after sending command to cab screen."
                End If
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, unknown reply by digital command control unit after sending command to cab screen."
            End If
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update the Cab Screen with Prompts
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainScreen!MSComm1.InBufferCount = 0
    Let MainScreen!MSComm1.Output = Chr$(&H90) & Chr$(Val(LocomotiveCabAssignment)) & "CONSIST " & Format(Str$(LocomotiveDatabaseConsist.Recordset.Fields(0)), "0000") & " HAS"
    'Chr$ (Val(LocomotiveConsistAssignment)) & " Consist " & Temporary$ & "    "
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.Labelmode.Text = "Standard Mode" Then
        Let TemporaryTimeOutCounter = 0
        Do While MainScreen.MSComm1.InBufferCount < 1
            Let TemporaryTimeOutCounter = TemporaryTimeOutCounter + 1
            If TemporaryTimeOutCounter > 750000 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, time out occured by digital command control unit after sending command."
                End If
                Let MousePointer = vbDefault
                Let LabelStatus.Caption = "Status: Digital command control timed out."
                Exit Sub
            End If
            Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
            DoEvents
        Loop
        Let TemporaryInput = MainScreen.MSComm1.Input
        If Len(TemporaryInput) <> 1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, wrong number of bytes returned from digital command control unit after send command to cab screen."
            End If
        ElseIf TemporaryInput <> "!" Then
            If TemporaryInput = "0" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, command not supported by digital command control unit after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "2" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, cab address out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "3" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, data out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "4" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, byte count out of range after sending command to cab screen."
                End If
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, unknown reply by digital command control unit after sending command to cab screen."
            End If
        End If
    Let LabelStatus.Caption = "Status: Idle"
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Output
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MainScreen.MSComm1.InBufferCount = 0
    Let MainScreen.MSComm1.Output = Chr$(&H91) & Chr$(Val(LocomotiveConsistAssignment)) & "BEEN ASSIGNED   "
    '& Chr$(Val(LocomotiveConsistAssignment)) & " Assigned to Cab "
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Wait for Response
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.Labelmode.Text = "Standard Mode" Then
        Let TemporaryTimeOutCounter = 0
        Do While MainScreen.MSComm1.InBufferCount < 1
            Let TemporaryTimeOutCounter = TemporaryTimeOutCounter + 1
            If TemporaryTimeOutCounter > 750000 Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, time out occured by digital command control unit after sending command."
                End If
                Let LabelStatus.Caption = "Status: Command aborted due to error."
                Exit Sub
            End If
            Let LabelStatus.Caption = "Status: Wait for response from digitial command control unit."
            DoEvents
        Loop
        Let TemporaryInput = MainScreen.MSComm1.Input
        If Len(TemporaryInput) <> 1 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, wrong number of bytes returned from digital command control unit after sending command to cab screen."
            End If
                ElseIf TemporaryInput <> "!" Then
            If TemporaryInput = "0" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, command not supported by digital command control unit after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "2" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, cab address out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "3" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, data out of range after sending command to cab screen."
                End If
            ElseIf TemporaryInput = "4" Then
                If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                    Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                    MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                    Let Ini.Filename = App.Path$ & "\Atc.log"
                    Let Ini.Application = "Log Errors"
                    Let Ini.Parameter = Date$ & " " & Time$
                    Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, byte count out of range after sending command to cab screen."
                End If
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Mainline Consist Screen, Menu Load Assign, unknown reply by digital command control unit after sending command to cab screen."
            End If
        End If
    Let LabelStatus.Caption = "Status: Idle"
    End If

End Sub

Private Sub menuLoadATC_Click()

    Let LocomotiveConsistNumber.Text = LocomotiveDatabaseConsist.Recordset.Fields(0)
    
    Let MousePointer = vbHourglass
    Load MainlineOperation
    MainlineOperationATC.Show vbModeless
    Let MousePointer = vbDefault

End Sub

Private Sub menuLoadGUI_Click()

    Let LocomotiveConsistNumber.Text = LocomotiveDatabaseConsist.Recordset.Fields(0)
    
    Let MousePointer = vbHourglass
    Load MainlineOperation
    MainlineOperationGUI.Show vbModeless
    Let MousePointer = vbDefault

End Sub


Private Sub menuLocomotive_Click()
   
    Let LocomotiveConsistTable.Col = Val(LocomotiveConsistTableCol.Text)
    Let LocomotiveConsistTable.Row = Val(LocomotiveConsistTableRow.Text)
    
    If Val(LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col)) = 0 Then
        Let menuLocomotiveDelete.Enabled = False
        If CheckBoxAddressShort.Value = False Then
            If Val(LocomotiveConsistSelect.Text) > 0 And Val(LocomotiveConsistSelect.Text) < 9999 Then
                Let menuLocomotiveAdd.Enabled = True
            Else
                Let menuLocomotiveAdd.Enabled = False
            End If
        ElseIf CheckBoxAddressShort.Value = True Then
            If Val(LocomotiveConsistSelect.Text) > 0 And Val(LocomotiveConsistSelect.Text) < 127 Then
                Let menuLocomotiveAdd.Enabled = True
            Else
                Let menuLocomotiveAdd.Enabled = False
            End If
        End If
        Let menuLocomotiveEdit.Enabled = False
    Else
        Let menuLocomotiveDelete.Enabled = True
        Let menuLocomotiveAdd.Enabled = False
        Let menuLocomotiveEdit.Enabled = True
    End If

End Sub

Private Sub menuLocomotiveAdd_Click()
    
    Dim TemporaryCommand As String
    Dim TemporaryByteCount As Integer

    Let MousePointer = vbHourglass

    If Val(LocomotiveConsistSelect.Text) <= 0 Or Val(LocomotiveConsistSelect.Text) >= 9999 Then
        Let LocomotiveConsistSelect.Text = ""
    Else
        If CheckBoxUniqueConsist.Value = vbChecked Then
            Call UniqueConsist
        End If
    End If

    If CheckBoxUniqueConsistfailed.Value = vbUnchecked Then
        Let LocomotiveConsistTable.Col = Val(LocomotiveConsistTableCol.Text)
        Let LocomotiveConsistTable.Row = Val(LocomotiveConsistTableRow.Text)
        LocomotiveDatabaseConsist.Recordset.Edit
        Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col) = LocomotiveConsistSelect.Text
        If LocomotiveConsistSelect.Text <> "" Then
            If LocomotiveDirectionNormal.Value = vbChecked Then
                Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col + 1) = "F"
            End If
            If LocomotiveDirectionReversed.Value = vbChecked Then
                Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col + 1) = "R"
            End If
        Else
            Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col + 1) = " "
        End If
            
        LocomotiveDatabaseConsist.Recordset.Update
    End If
    Let CheckBoxUniqueConsistfailed.Value = vbUnchecked
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Setup Command Station for Consist Updates
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LabelStatus.Caption = "Status: Adding Locomotive to Consist"
    Let MainScreen!onebyted.Text = Int(Val(LocomotiveConsistSelect.Text) / 256)
    Let MainScreen!twoByteD.Text = Val(LocomotiveConsistSelect.Text) - (Val(MainScreen!onebyted.Text) * 256)
    Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64 ' add first two bytes to address one byte for long address
    Let MainScreen!ThreeByteD.Text = 16 ' instruction byte
    If LocomotiveDirectionNormal.Value = vbChecked Then
        Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 ' activate consist in forward direction
    ElseIf LocomotiveDirectionReversed.Value = vbChecked Then
        Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 + 1 ' activate consist in reverse direction
    End If
    Let MainScreen!FourByteD.Text = LocomotiveDatabaseConsist.Recordset.Fields(0) ' which is the consist number
    Let MainScreen!FiveByteD.Text = ""
    Let MainScreen!SixByteD.Text = ""
    Let MainScreen!sevenbyted.Text = ""
    Call MainScreen.SendCommandviaTempQ
    Let MousePointer = Default
End Sub

Private Sub menuLocomotiveDelete_Click()

    Dim TemporaryLocomotiveDelete As String
    Let MousePointer = vbHourglass
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Poll Information from Database
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveDatabaseConsist.Recordset.Edit
    If LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col) <> "" Then
        Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col) = ""
        Let LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col + 1) = " "
        LocomotiveDatabaseConsist.Recordset.Update
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Setup Command Station for Consist Updates (remove locomotive)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let LabelStatus.Caption = "Status: Removing Locomotive from Consist"
        Let TemporaryLocomotiveDelete = LocomotiveDatabaseConsist.Recordset.Fields(LocomotiveConsistTable.Col)
        Let MainScreen!onebyted.Text = Int(Val(TemporaryLocomotiveDelete) / 256) '                                      MSB of locomotive address
        Let MainScreen!twoByteD.Text = Val(TemporaryLocomotiveDelete) - (Val(MainScreen!onebyted.Text) * 256) '          LSB of locomotive address
        Let MainScreen!onebyted.Text = Val(MainScreen!onebyted.Text) + 128 + 64 '                                       Modify the MSB for the command Station
        Let MainScreen!ThreeByteD.Text = 16 '                                                                           Instruction byte
        If LocomotiveDirectionNormal.Value = vbChecked Then
            Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 '                                      Activate consist in forward direction
        ElseIf LocomotiveDirectionReversed.Value = vbChecked Then
            Let MainScreen!ThreeByteD.Text = Val(MainScreen!ThreeByteD.Text) + 2 + 1 '                                  Activate consist in reverse direction
        End If
        Let MainScreen!FourByteD.Text = 0 '                                                                             Zero to remove the consist
        Let MainScreen!FiveByteD.Text = ""
        Let MainScreen!SixByteD.Text = ""
        Let MainScreen!sevenbyted.Text = ""
        
        Call MainScreen.SendCommandviaTempQ
    End If
    Let MousePointer = vbDefault

End Sub


Private Sub menuLocomotiveEdit_Click()

    Load MainlineDiesel
    MainlineDiesel.Show vbModeless

End Sub







Private Sub UniqueConsist()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unique Locomotive in Consist
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let MousePointer = vbHourglass
    Data1.Recordset.MoveFirst
    For OutsideLoop = 0 To 126
        For InsideLoop = 2 To 18 Step 2
            If LocomotiveConsistSelect.Text = Data1.Recordset.Fields(InsideLoop) Then
                MsgBox "The Locomotive you have selcted is currently assigned to a consist. Please check consist number " + Str$(OutsideLoop + 1) + ".", vbOKOnly, "Consist Safety"
                Let CheckBoxUniqueConsistfailed.Value = vbChecked
            End If
        Next InsideLoop
        Data1.Recordset.MoveNext
    Next OutsideLoop
    Data1.Recordset.MoveFirst
    Let MousePointer = Default
    

End Sub













Public Sub ModifyLocomotiveDatabaseLong()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef

Set OldWs = DBEngine.Workspaces(0)
    Let TemporaryFileName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
Set NewDb = OldWs.OpenDatabase(TemporaryFileName)
Set myTable = NewDb.CreateTableDef("LocomotiveDatabaseLong")

Set myF1 = myTable.CreateField("LocomotiveNumber", dbInteger)
'Let myF1.AllowZeroLength = True 'Only for text or memo types only
Let myF1.Attributes = dbUpdatableField
Let myF1.DefaultValue = "0000"

Set myF2 = myTable.CreateField("LocomotivePrototype", dbMemo)
Let myF2.AllowZeroLength = True
Let myF2.Attributes = dbUpdatableField
Let myF2.DefaultValue = "Locomotive Prototype Owner"

Set myF3 = myTable.CreateField("LocomotiveActiveList", dbBoolean)
'Let myF3.AllowZeroLength = True
Let myF3.Attributes = dbUpdatableField
Let myF3.DefaultValue = False

Set myF4 = myTable.CreateField("LocomotiveDecoderEquiped", dbBoolean)
'Let myF4.AllowZeroLength = True
Let myF4.Attributes = dbUpdatableField
Let myF4.DefaultValue = False

Set myF5 = myTable.CreateField("LocomotiveManufacturer", dbMemo)
Let myF5.AllowZeroLength = True
Let myF5.Attributes = dbUpdatableField
Let myF5.DefaultValue = "Locomotive Manufacturer"

Set myF6 = myTable.CreateField("LocomotiveModel", dbMemo)
Let myF6.AllowZeroLength = True
Let myF6.Attributes = dbUpdatableField
Let myF6.DefaultValue = "Locomotive Model"

Set myF7 = myTable.CreateField("LocomotiveModelManufacturer", dbMemo)
Let myF7.AllowZeroLength = True
Let myF7.Attributes = dbUpdatableField
Let myF7.DefaultValue = "Locomotive Model Manufacturer"

Set myF8 = myTable.CreateField("LocomotiveFacts", dbMemo)
Let myF8.AllowZeroLength = True
Let myF8.Attributes = dbUpdatableField
Let myF8.DefaultValue = "Locomotive Facts"

Set myF9 = myTable.CreateField("LocomotiveImage", dbMemo)
Let myF9.AllowZeroLength = True
Let myF9.Attributes = dbUpdatableField
Let myF9.DefaultValue = ""

Set myF10 = myTable.CreateField("LocomotivePrototypeAdhesionFactor", dbMemo)
Let myF10.AllowZeroLength = True
Let myF10.Attributes = dbUpdatableField
Let myF10.DefaultValue = "0"

Set myF11 = myTable.CreateField("LocomotivePrototypeHorsePower", dbMemo)
Let myF11.AllowZeroLength = True
Let myF11.Attributes = dbUpdatableField
Let myF11.DefaultValue = "0"

Set myF12 = myTable.CreateField("LocomotivePrototypeTractionEffort", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotivePrototypeDrawBarPull", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"

Set myF14 = myTable.CreateField("LocomotivePrototypeWeight", dbMemo)
Let myF14.AllowZeroLength = True
Let myF14.Attributes = dbUpdatableField
Let myF14.DefaultValue = "0"

Set myF15 = myTable.CreateField("LocomotiveModelAdhesionFactor", dbMemo)
Let myF15.AllowZeroLength = True
Let myF15.Attributes = dbUpdatableField
Let myF15.DefaultValue = "0"

Set myF16 = myTable.CreateField("LocomotiveModelHorsePower", dbMemo)
Let myF16.AllowZeroLength = True
Let myF16.Attributes = dbUpdatableField
Let myF16.DefaultValue = "0"

Set myF17 = myTable.CreateField("LocomotiveModelTractionEffort", dbMemo)
Let myF17.AllowZeroLength = True
Let myF17.Attributes = dbUpdatableField
Let myF17.DefaultValue = "0"

Set myF18 = myTable.CreateField("LocomotiveModelDrawBarPull", dbMemo)
Let myF18.AllowZeroLength = True
Let myF18.Attributes = dbUpdatableField
Let myF18.DefaultValue = "0"

Set myf19 = myTable.CreateField("LocomotiveModelWeight", dbMemo)
Let myf19.AllowZeroLength = True
Let myf19.Attributes = dbUpdatableField
Let myf19.DefaultValue = "0"

Set myF20 = myTable.CreateField("LocomotiveCouplerTypeFront", dbMemo)
Let myF20.AllowZeroLength = True
Let myF20.Attributes = dbUpdatableField
Let myF20.DefaultValue = "Choose a Front Coupler"

Set myF21 = myTable.CreateField("LocomotiveCouplerTypeRear", dbMemo)
Let myF21.AllowZeroLength = True
Let myF21.Attributes = dbUpdatableField
Let myF21.DefaultValue = "Choose a Rear Coupler"

myTable.Fields.Append myF1
myTable.Fields.Append myF2
myTable.Fields.Append myF3
myTable.Fields.Append myF4
myTable.Fields.Append myF5
myTable.Fields.Append myF6
myTable.Fields.Append myF7
myTable.Fields.Append myF8
myTable.Fields.Append myF9
myTable.Fields.Append myF10
myTable.Fields.Append myF11
myTable.Fields.Append myF12
myTable.Fields.Append myF13
myTable.Fields.Append myF14
myTable.Fields.Append myF15
myTable.Fields.Append myF16
myTable.Fields.Append myF17
myTable.Fields.Append myF18
myTable.Fields.Append myf19
myTable.Fields.Append myF20
myTable.Fields.Append myF21

NewDb.TableDefs.Append myTable

Set myRecord = NewDb.OpenRecordset("LocomotiveDatabaseLong")

For t = 0 To 10000
    myRecord.AddNew
    myRecord.Fields("LocomotiveNumber") = Str$(t)
    myRecord.Update
    DoEvents
Next

NewDb.Close

End Sub

Public Sub ModifyLocomotiveDatabaseShort()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef

Set OldWs = DBEngine.Workspaces(0)
    Let TemporaryFileName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
Set NewDb = OldWs.OpenDatabase(TemporaryFileName)
Set myTable = NewDb.CreateTableDef("LocomotiveDatabaseShort")

Set myF1 = myTable.CreateField("LocomotiveNumber", dbInteger)
'Let myF1.AllowZeroLength = True 'Only for text or memo types only
Let myF1.Attributes = dbUpdatableField
Let myF1.DefaultValue = "0000"

Set myF2 = myTable.CreateField("LocomotivePrototype", dbMemo)
Let myF2.AllowZeroLength = True
Let myF2.Attributes = dbUpdatableField
Let myF2.DefaultValue = "Locomotive Prototype Owner"

Set myF3 = myTable.CreateField("LocomotiveActiveList", dbBoolean)
'Let myF3.AllowZeroLength = True
Let myF3.Attributes = dbUpdatableField
Let myF3.DefaultValue = False

Set myF4 = myTable.CreateField("LocomotiveDecoderEquiped", dbBoolean)
'Let myF4.AllowZeroLength = True
Let myF4.Attributes = dbUpdatableField
Let myF4.DefaultValue = False

Set myF5 = myTable.CreateField("LocomotiveManufacturer", dbMemo)
Let myF5.AllowZeroLength = True
Let myF5.Attributes = dbUpdatableField
Let myF5.DefaultValue = "Locomotive Manufacturer"

Set myF6 = myTable.CreateField("LocomotiveModel", dbMemo)
Let myF6.AllowZeroLength = True
Let myF6.Attributes = dbUpdatableField
Let myF6.DefaultValue = "Locomotive Model"

Set myF7 = myTable.CreateField("LocomotiveModelManufacturer", dbMemo)
Let myF7.AllowZeroLength = True
Let myF7.Attributes = dbUpdatableField
Let myF7.DefaultValue = "Locomotive Model Manufacturer"

Set myF8 = myTable.CreateField("LocomotiveFacts", dbMemo)
Let myF8.AllowZeroLength = True
Let myF8.Attributes = dbUpdatableField
Let myF8.DefaultValue = "Locomotive Facts"

Set myF9 = myTable.CreateField("LocomotiveImage", dbMemo)
Let myF9.AllowZeroLength = True
Let myF9.Attributes = dbUpdatableField
Let myF9.DefaultValue = ""

Set myF10 = myTable.CreateField("LocomotivePrototypeAdhesionFactor", dbMemo)
Let myF10.AllowZeroLength = True
Let myF10.Attributes = dbUpdatableField
Let myF10.DefaultValue = "0"

Set myF11 = myTable.CreateField("LocomotivePrototypeHorsePower", dbMemo)
Let myF11.AllowZeroLength = True
Let myF11.Attributes = dbUpdatableField
Let myF11.DefaultValue = "0"

Set myF12 = myTable.CreateField("LocomotivePrototypeTractionEffort", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotivePrototypeDrawBarPull", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"


Set myF14 = myTable.CreateField("LocomotivePrototypeWeight", dbMemo)
Let myF14.AllowZeroLength = True
Let myF14.Attributes = dbUpdatableField
Let myF14.DefaultValue = "0"

Set myF15 = myTable.CreateField("LocomotiveModelAdhesionFactor", dbMemo)
Let myF15.AllowZeroLength = True
Let myF15.Attributes = dbUpdatableField
Let myF15.DefaultValue = "0"

Set myF16 = myTable.CreateField("LocomotiveModelHorsePower", dbMemo)
Let myF16.AllowZeroLength = True
Let myF16.Attributes = dbUpdatableField
Let myF16.DefaultValue = "0"

Set myF17 = myTable.CreateField("LocomotiveModelTractionEffort", dbMemo)
Let myF17.AllowZeroLength = True
Let myF17.Attributes = dbUpdatableField
Let myF17.DefaultValue = "0"

Set myF18 = myTable.CreateField("LocomotiveModelDrawBarPull", dbMemo)
Let myF18.AllowZeroLength = True
Let myF18.Attributes = dbUpdatableField
Let myF18.DefaultValue = "0"

Set myf19 = myTable.CreateField("LocomotiveModelWeight", dbMemo)
Let myf19.AllowZeroLength = True
Let myf19.Attributes = dbUpdatableField
Let myf19.DefaultValue = "0"

Set myF20 = myTable.CreateField("LocomotiveCouplerTypeFront", dbMemo)
Let myF20.AllowZeroLength = True
Let myF20.Attributes = dbUpdatableField
Let myF20.DefaultValue = "Choose a Front Coupler"

Set myF21 = myTable.CreateField("LocomotiveCouplerTypeRear", dbMemo)
Let myF21.AllowZeroLength = True
Let myF21.Attributes = dbUpdatableField
Let myF21.DefaultValue = "Choose a Rear Coupler"

myTable.Fields.Append myF1
myTable.Fields.Append myF2
myTable.Fields.Append myF3
myTable.Fields.Append myF4
myTable.Fields.Append myF5
myTable.Fields.Append myF6
myTable.Fields.Append myF7
myTable.Fields.Append myF8
myTable.Fields.Append myF9
myTable.Fields.Append myF10
myTable.Fields.Append myF11
myTable.Fields.Append myF12
myTable.Fields.Append myF13
myTable.Fields.Append myF14
myTable.Fields.Append myF15
myTable.Fields.Append myF16
myTable.Fields.Append myF17
myTable.Fields.Append myF18
myTable.Fields.Append myf19
myTable.Fields.Append myF20
myTable.Fields.Append myF21

NewDb.TableDefs.Append myTable

Set myRecord = NewDb.OpenRecordset("LocomotiveDatabaseShort")

For t = 0 To 127
    myRecord.AddNew
    myRecord.Fields("LocomotiveNumber") = Str$(t)
    myRecord.Update
    DoEvents
Next

NewDb.Close

End Sub

Public Sub ModifyLocomotiveDatabaseCopy()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef

Set OldWs = DBEngine.Workspaces(0)

Let TemporaryFileName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"

Set NewDb = OldWs.OpenDatabase(TemporaryFileName)
Set myrecordset1 = NewDb.OpenRecordset("LocomotiveDatabase")
Set myrecordset2 = NewDb.OpenRecordset("LocomotiveDatabaseLong")
'Set myRecord = NewDb.OpenRecordset("LocomotiveDatabaseLong")

myrecordset1.MoveFirst
myrecordset2.MoveFirst

While X < 9999
myrecordset2.Edit

Let myrecordset2.Fields("LocomotiveNumber") = myrecordset1.Fields("LocomotiveNumber")
If myrecordset1.Fields("LocomotivePrototype") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototype") = myrecordset1.Fields("LocomotivePrototype")
End If

Let myrecordset2.Fields("LocomotiveActiveList") = myrecordset1.Fields("LocomotiveActiveList")
Let myrecordset2.Fields("LocomotiveDecoderEquiped") = myrecordset1.Fields("LocomotiveDecoderEquiped")

If myrecordset1.Fields("LocomotiveManufacturer") <> "" Then
    Let myrecordset2.Fields("LocomotiveManufacturer") = myrecordset1.Fields("LocomotiveManufacturer")
End If
If myrecordset1.Fields("LocomotiveModel") <> "" Then
    Let myrecordset2.Fields("LocomotiveModel") = myrecordset1.Fields("LocomotiveModel")
End If

If myrecordset1.Fields("LocomotiveModelManufacturer") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelManufacturer") = myrecordset1.Fields("LocomotiveModelManufacturer")
End If
If myrecordset1.Fields("LocomotiveFacts") <> "" Then
    Let myrecordset2.Fields("LocomotiveFacts") = myrecordset1.Fields("LocomotiveFacts")
End If
If myrecordset1.Fields("LocomotiveImage") <> "" Then
    Let myrecordset2.Fields("LocomotiveImage") = myrecordset1.Fields("LocomotiveImage")
End If
If myrecordset1.Fields("LocomotivePrototypeAdhesionFactor") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototypeAdhesionFactor") = myrecordset1.Fields("LocomotivePrototypeAdhesionFactor")
End If
If myrecordset1.Fields("LocomotivePrototypeHorsePower") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototypeHorsePower") = myrecordset1.Fields("LocomotivePrototypeHorsePower")
End If
If myrecordset1.Fields("LocomotivePrototypeHorsePower") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototypeTractionEffort") = myrecordset1.Fields("LocomotivePrototypeHorsePower")
End If
If myrecordset1.Fields("LocomotivePrototypeDrawBarPull") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototypeDrawBarPull") = myrecordset1.Fields("LocomotivePrototypeDrawBarPull")
End If
If myrecordset1.Fields("LocomotivePrototypeWeight") <> "" Then
    Let myrecordset2.Fields("LocomotivePrototypeWeight") = myrecordset1.Fields("LocomotivePrototypeWeight")
End If
If myrecordset1.Fields("LocomotiveModelAdhesionFactor") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelAdhesionFactor") = myrecordset1.Fields("LocomotiveModelAdhesionFactor")
End If
If myrecordset1.Fields("LocomotiveModelHorsePower") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelHorsePower") = myrecordset1.Fields("LocomotiveModelHorsePower")
End If
If myrecordset1.Fields("LocomotiveModelTractionEffort") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelTractionEffort") = myrecordset1.Fields("LocomotiveModelTractionEffort")
End If
If myrecordset1.Fields("LocomotiveModelDrawBarPull") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelDrawBarPull") = myrecordset1.Fields("LocomotiveModelDrawBarPull")
End If
If myrecordset1.Fields("LocomotiveModelWeight") <> "" Then
    Let myrecordset2.Fields("LocomotiveModelWeight") = myrecordset1.Fields("LocomotiveModelWeight")
End If

myrecordset2.Update
myrecordset1.MoveNext
myrecordset2.MoveNext
X = X + 1
Wend

NewDb.Close

End Sub
