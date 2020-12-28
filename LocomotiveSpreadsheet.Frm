VERSION 4.00
Begin VB.Form LocomotiveSpreadsheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Locomotive Spreadsheet"
   ClientHeight    =   4050
   ClientLeft      =   1860
   ClientTop       =   2010
   ClientWidth     =   10530
   Height          =   4455
   Icon            =   "LocomotiveSpreadsheet.frx":0000
   Left            =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10530
   Top             =   1665
   Width           =   10650
   Begin VB.CommandButton ButtonPrintList 
      Caption         =   "Print List"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxAddressShort 
      Caption         =   "Short Addresses"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonCompile 
      Caption         =   "Compile"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxDecoderEquiped 
      Caption         =   " Decoder Equiped for"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CheckBox CheckBoxActiveList 
      Caption         =   "Active List and/or"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "LocomotiveSpreadsheet.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.Data LocomotiveDatabaseSearch 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBaseLong"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   10800
      Top             =   2100
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   10800
      Top             =   1560
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   10800
      Top             =   960
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin REPORTOCXLib.ReportOcx Report 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   239
   End
   Begin VB.Label Label2 
      Caption         =   $"LocomotiveSpreadsheet.frx":0884
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   9735
   End
   Begin MSGrid.Grid Grid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   4895
      _StockProps     =   77
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
   End
End
Attribute VB_Name = "LocomotiveSpreadsheet"
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
            If TemporaryScreen = "Locomotive Spreadsheet Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.INI file. Please email the author reporting the error and attach a copy of the file called ATC.INI for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.ini"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, Button Close, locating the current screen in the ATC.INI file (screen stack) counting from 9 to 0 step -1"
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
             '   FunScreen.Show vbModeless
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
                Let Ini.Value = "Locomotive Spreadsheet Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Locomotive Spreadsheet Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

Private Sub ButtonCompile_Click()

Let ButtonClose.Enabled = False
Let ButtonPrint.Enabled = False
Let MousePointer = vbArrowHourglass
Let Grid1.Rows = 2

' Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

LocomotiveDatabaseSearch.Recordset.MoveFirst
Let temporaryRow = 1
Let TemporaryCol = 0
    

' Use the RecordCount property to find out how many records in a Recordset or TableDef object have been accessed.
' RecordCount doesn't indicate how many records are contained in a dynaset- or snapshot-type Recordset until all records have been accessed.
' Once the last record has been accessed, the RecordCount property indicates the total number of undeleted records in the Recordset or TableDef.
' To force the last record to be accessed, use the MoveLast or FindLast method on the Recordset.
' You can also use an SQL Count function to determine the approximate number of records your query will return.

While Not LocomotiveDatabaseSearch.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    If checkboxactivelist = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(2).Value = "True" Or _
       checkboxdecoderequiped.Value = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(3).Value = "True" Then
            Let Grid1.Rows = temporaryRow + 1
            Let Grid1.Row = temporaryRow
        
            If LocomotiveDatabaseSearch.Recordset.Fields(0).Value <> "" Then
                Let Grid1.Col = 0
                Let Grid1.Text = LocomotiveDatabaseSearch.Recordset.Fields(0).Value
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(1).Value <> "" Then
                Let Grid1.Col = 1
                Let Grid1.Text = LocomotiveDatabaseSearch.Recordset.Fields(1).Value
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(4).Value <> "" Then
                Let Grid1.Col = 2
                Let Grid1.Text = LocomotiveDatabaseSearch.Recordset.Fields(4).Value
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(5).Value <> "" Then
                Let Grid1.Col = 3
                Let Grid1.Text = LocomotiveDatabaseSearch.Recordset.Fields(5).Value
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(6).Value <> "" Then
                Let Grid1.Col = 4
                Let Grid1.Text = LocomotiveDatabaseSearch.Recordset.Fields(6).Value
            End If
            Let temporaryRow = temporaryRow + 1
        
    End If
    LocomotiveDatabaseSearch.Recordset.MoveNext
    DoEvents
Wend

Let MousePointer = Default
Let ButtonClose.Enabled = True
Let ButtonPrint.Enabled = True

End Sub

Private Sub ButtonPrint_Click()

    LocomotiveSpreadsheet.PrintForm
    
End Sub

Private Sub ButtonPrintList_Click()

MousePointer = vbHourglass

With Report
        .Clear
        .SetColNums 5
        .SetFixedLeftMargin 0
        
        .HeadLine = False
        .SetPageHeadFormat "Arial", 24, 0, RGB(0, 0, 255), False, False, False
        .SetPageHead "Automatic Train Control - Locomotive List                Date:" & Date
        
        .SetReportTitleHeight 10
        .SetReportSubTitleHeight 0
                
        .SetWidthString "50|140|220|120|170"
        
        
        .SetTopString "Number|Road Name|Manufacturer|Model|Model Manufacturer"
        .SetBottomString "Number|Road Name|Manufacturer|Model|Model Manufacturer"
        .SetReportHeadHeight 0
    End With
        
LocomotiveDatabaseSearch.Recordset.MoveFirst
        
While Not LocomotiveDatabaseSearch.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.
    
    If checkboxactivelist = vbChecked And _
       LocomotiveDatabaseSearch.Recordset.Fields(2).Value = "True" Or _
       checkboxdecoderequiped.Value = vbChecked And _
       LocomotiveDatabaseSearch.Recordset.Fields(3).Value = "True" Then
                   
            If LocomotiveDatabaseSearch.Recordset.Fields(0).Value <> "" Then
                Let temp$ = LocomotiveDatabaseSearch.Recordset.Fields(0).Value
            Else
                Let temp$ = " "
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(1).Value <> "" Then
                Let temp$ = temp$ + "|" + LocomotiveDatabaseSearch.Recordset.Fields(1).Value
            Else
                Let temp$ = temp$ + "| "
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(4).Value <> "" Then
                Let temp$ = temp$ + "|" + LocomotiveDatabaseSearch.Recordset.Fields(4).Value
            Else
                Let temp$ = temp$ + "| "
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(5).Value <> "" Then
                Let temp$ = temp$ + "|" + LocomotiveDatabaseSearch.Recordset.Fields(5).Value
            Else
                Let temp$ = temp$ + "| "
            End If
            If LocomotiveDatabaseSearch.Recordset.Fields(6).Value <> "" Then
                Let temp$ = temp$ + "|" + LocomotiveDatabaseSearch.Recordset.Fields(6).Value
            Else
                Let temp$ = temp$ + "| "
            End If
    Report.AddRecord temp$
    End If
    
    LocomotiveDatabaseSearch.Recordset.MoveNext
    DoEvents
Wend
     
Report.DoPrintPreview
     
MousePointer = Default

End Sub

Private Sub checkboxactivelist_Click()

If checkboxdecoderequiped.Value = vbChecked Or _
   checkboxactivelist.Value = vbChecked Then
    Let ButtonPrint.Enabled = False
    Let ButtonCompile.Enabled = True
Else
    Let ButtonPrint.Enabled = False
    Let ButtonCompile.Enabled = False
End If

End Sub

Private Sub CheckBoxAddressShort_Click()

If CheckBoxAddressShort.Value = vbChecked Then
    LocomotiveDatabaseSearch.RecordSource = "LocomotiveDatabaseShort"
    LocomotiveDatabaseSearch.Refresh
    LocomotiveDatabaseSearch.Recordset.MoveFirst
Else
    LocomotiveDatabaseSearch.RecordSource = "LocomotiveDatabaseLong"
    LocomotiveDatabaseSearch.Refresh
    LocomotiveDatabaseSearch.Recordset.MoveFirst
End If

End Sub

Private Sub checkboxdecoderequiped_Click()

If checkboxdecoderequiped.Value = vbChecked Or _
   checkboxactivelist.Value = vbChecked Then
    Let ButtonPrint.Enabled = False
    Let ButtonCompile.Enabled = True
Else
    Let ButtonPrint.Enabled = False
    Let ButtonCompile.Enabled = False
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
        If TemporaryScreen = "Locomotive Spreadsheet Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Locomotive Spreadsheet Screen"
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
            Let Ini.Value = "Locomotive Spredsheet Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Locomotive Spreadsheet Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Locomotive Spreadsheet Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call BalloonHelpUpdatePart01
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Parameters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Locomotive Spreadsheet Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(LocomotiveCVSpreadsheet.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(LocomotiveCVSpreadsheet.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(LocomotiveCVSpreadsheet.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(LocomotiveCVSpreadsheet.Height)
    Let Ini.Parameter = "CheckBoxActiveList"
    Let Ini.Value = Str(checkboxactivelist.Value)
    Let Ini.Parameter = "CheckBoxDecoderEquiped"
    Let Ini.Value = Str(checkboxdecoderequiped.Value)
    Let Ini.Parameter = "CheckBoxAddressShort"
    Let Ini.Value = Str(CheckBoxAddressShort.Value)
    
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
                Let Ini.Value = "Locomotive Spreadsheet Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Locomotive Spreadsheet Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveSpreadsheet.Hide
    'unload LocomotiveSpreadsheet

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
    Let Ini.Application = "LocomotiveSpreadsheet Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "CheckBoxActiveList"
    Let checkboxactivelist.Value = Val(Ini.Value)
    Let Ini.Parameter = "CheckBoxDecoderEquiped"
    Let checkboxdecoderequiped.Value = Val(Ini.Value)
    Let Ini.Parameter = "CheckBoxAddressShort"
    Let CheckBoxAddressShort.Value = Val(Ini.Value)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        LocomotiveSpreadsheet.Left = (Screen.Width - Width) / 2
        LocomotiveSpreadsheet.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + LocomotiveSpreadsheet.Width > Screen.Width Then
            Let LocomotiveSpreadsheet.Left = Screen.Width - LocomotiveSpreadsheet.Width
        Else
            Let LocomotiveSpreadsheet.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + LocomotiveSpreadsheet.Height > Screen.Height Then
            Let LocomotiveSpreadsheet.Top = Screen.Height - LocomotiveSpreadsheet.Height
        Else
            Let LocomotiveSpreadsheet.Top = Val(TemporaryValueTop)
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
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Let Grid1.ColWidth(0) = 800
Let Grid1.ColWidth(1) = 3000
Let Grid1.ColWidth(2) = 2500
Let Grid1.ColWidth(3) = 1500
Let Grid1.ColWidth(4) = 2500

Let Grid1.Row = 0
Let Grid1.Col = 0
Let Grid1.Text = "Locomotive"
Let Grid1.Col = 1
Let Grid1.Text = "Road Name"
Let Grid1.Col = 2
Let Grid1.Text = "Manufacturer"
Let Grid1.Col = 3
Let Grid1.Text = "Model"
Let Grid1.Col = 4
Let Grid1.Text = "Model Manufacturer"

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

    If LocomotiveSpreadsheet.WindowState = vbMinimized Then
    
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
        
    ElseIf LocomotiveSpreadsheet.WindowState = vbNormal Then
    
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

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Locomotive Spreadsheet Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(LocomotiveSpreadsheet.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(LocomotiveSpreadsheet.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(LocomotiveSpreadsheet.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(LocomotiveSpreadsheet.Height)
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
End Sub



Private Sub BalloonHelpUpdatePart01()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is On" Then
        Dim BalloonHelpText1 As String
        Dim BalloonHelpText2 As String
        Dim BalloonHelpSetup As Long
        Dim BalloonHelpFont As New StdFont
        Dim BalloonHelpVisibleTime As Long
        Dim BalloonHelpTimeDelay As Long
        Dim BalloonHelpShadow As Boolean
        Dim BalloonHelpCenter As Boolean
        Dim BalloonHelpShowOnDemand As Boolean
        Dim BalloonHelpOpacity As Byte
        Dim BalloonHelpWaveFile As String '

        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let BalloonHelpFont.Name = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonHelpFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontBold"
        Let BalloonHelpFont.Bold = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontItalic"
        Let BalloonHelpFont.Italic = Ini.Value
        Let Ini.Parameter = "BalloonHelpFontUnderline"
        Let BalloonHelpFont.Underline = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let BalloonHelpColour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let BalloonHelpColour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let BalloonHelpColour3 = Ini.Value
        Let Ini.Parameter = "BalloonHelpVisibleTime"
        Let BalloonHelpVisibleTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpDelayTime"
        Let BalloonHelpDelayTime = Ini.Value
        Let Ini.Parameter = "BalloonHelpShadow"
        Let BalloonHelpShadow = Ini.Value
        Let Ini.Parameter = "BalloonHelpCenter"
        Let BalloonHelpCenter = Ini.Value
        Let Ini.Parameter = "BalloonHelpShowOnDemand"
        Let BalloonHelpShowOnDemand = Ini.Value
        Let Ini.Parameter = "BalloonHelpWaveFile"
        'Let balloonhelp.SoundFile = App.Path$ & "\Help\" & Ini.Value
        Let BalloonHelpWaveFile = App.Path$ & "\Help\" & Ini.Value
        If MainScreen!MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuspeechHelp.Caption = "&Speech Help is Off" Then
                Let balloonhelp.Speech = False
        Else 'If mainscreen!menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let LocomotiveSpreadsheet.MousePointer = ccHourglass
        
        Let BalloonHelpText1 = "This checkbox is used to select parameters." & vbCrLf & "Clicking this check box will include all locomotives" & vbCrLf & "active in your database."
        Let BalloonHelpText2 = "CheckBox Active Locomotive"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxActiveList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxactivelist, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'CheckBoxActiveList' control."
            End If
        End If
            
        Let BalloonHelpText1 = "This checkbox is used to select parameters." & vbCrLf & "Clicking this check box will include all locomotives" & vbCrLf & "atcive in your database."
        Let BalloonHelpText2 = "CheckBox Active Locomotives"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxActiveList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxactivelist, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'CheckBoxActiveList' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This checkbox is used to select parameters." & vbCrLf & "Clicking this check box will include all locomotives" & vbCrLf & "with decoders installed in your database."
        Let BalloonHelpText2 = "CheckBox Decoder Equiped"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxDecoderEquiped)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxdecoderequiped, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'CheckBoxDecoderEquiped' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This checkbox is used to select the current database." & vbCrLf & "You can toggle the database by simply 'click'ing on" & vbCrLf & "this checkbox to indicate the decoder address type, short or long."
        Let BalloonHelpText2 = "Address Type (Short or Long)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxAddressShort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxAddressShort, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'CheckBoxAdddressShort' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This grid displays the database information with" & vbCrLf & "your selected options (active locomotive, decoder equipped" & vbCrLf & "using the short or long address)."
        Let BalloonHelpText2 = "Locomotive Spreadsheet"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(Grid1)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(Grid1, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'Grid1' control."
            End If
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Buttons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Let BalloonHelpText1 = "This button will print the contents of the grid" & vbCrLf & "displaying the database information with your selected options" & vbCrLf & "(active locomotive, decoder equipped using the short or long address)."
        Let BalloonHelpText2 = "Print List Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrintList)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrintList, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'ButtonPrintList' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This button will print the Locomotive" & vbCrLf & "Spreadsheet window."
        Let BalloonHelpText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'ButtonPrint' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This compiles the data and displays the" & vbCrLf & "database information with your selected" & vbCrLf & "options (active locomotive, decoder equipped" & vbCrLf & "using the short or long address)."
        Let BalloonHelpText2 = "Complie Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonCompile)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonCompile, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'ButtonCompile' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This button will close the Locomotive" & vbCrLf & "Spreadsheet window and return you to the" & vbCrLf & "previous window."
        Let BalloonHelpText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive Spreadsheet, BalloonHelpUpdate, unable to setup balloon help for 'ButtonClose' control."
            End If
        End If
        
        Let LocomotiveSpreadsheet.MousePointer = ccDefault

    End If

End Sub
