VERSION 4.00
Begin VB.Form LocomotiveCVSpreadsheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Locomotive Configuration Variable SpreadSheet"
   ClientHeight    =   5100
   ClientLeft      =   3495
   ClientTop       =   2385
   ClientWidth     =   10530
   Height          =   5505
   Icon            =   "LocomotiveCVSpreadsheet.frx":0000
   Left            =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   10530
   Top             =   2040
   Width           =   10650
   Begin VB.CommandButton ButtonPrintList 
      Caption         =   "Print List"
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxAddressShort 
      Caption         =   "Short Addresses"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Data LocomotiveDatabaseDecoder 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDecoders"
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton ButtonCompile 
      Caption         =   "Compile"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox CheckBoxDecoderEquiped 
      Caption         =   "Decoder Equiped for"
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CheckBox CheckBoxActiveList 
      Caption         =   "Active List and/or"
      Height          =   255
      Left            =   3840
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
      Picture         =   "LocomotiveCVSpreadsheet.frx":0442
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
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBaseLong"
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab TabCVList 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "CV1-16"
      TabPicture(0)   =   "LocomotiveCVSpreadsheet.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CV17-32"
      TabPicture(1)   =   "LocomotiveCVSpreadsheet.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CV32-48"
      TabPicture(2)   =   "LocomotiveCVSpreadsheet.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "CV49-64"
      TabPicture(3)   =   "LocomotiveCVSpreadsheet.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CV65-80"
      TabPicture(4)   =   "LocomotiveCVSpreadsheet.frx":08F4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "CV81-96"
      TabPicture(5)   =   "LocomotiveCVSpreadsheet.frx":0910
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "CV97-112"
      TabPicture(6)   =   "LocomotiveCVSpreadsheet.frx":092C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Grid(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "CV113-128"
      TabPicture(7)   =   "LocomotiveCVSpreadsheet.frx":0948
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Grid(7)"
      Tab(7).ControlCount=   1
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   7
         Left            =   -74880
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   6
         Left            =   -74880
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   5
         Left            =   -74880
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   4
         Left            =   -74880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   3
         Left            =   -74880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   2
         Left            =   -74880
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   1
         Left            =   -74880
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
      Begin MSGrid.Grid Grid 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   17
         FixedCols       =   0
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   10800
      Top             =   2520
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   10800
      Top             =   2040
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   10800
      Top             =   1440
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
      Left            =   5280
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   239
   End
   Begin VB.Label Label2 
      Caption         =   $"LocomotiveCVSpreadsheet.frx":0964
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "LocomotiveCVSpreadsheet"
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
            If TemporaryScreen = "Locomotive CV Spreadsheet Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Locomotive CV Spreadsheet Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Locomotive CV Spreadsheet Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Locomotive CV Spreadsheet Screen, Button Close, stack is empty, underflow."
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

For TemporaryX = 0 To TabCVList.TabIndex
    Let Grid(TemproaryX).Rows = 2
Next TemporaryX

' Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

LocomotiveDatabaseSearch.Recordset.MoveFirst
LocomotiveDatabaseDecoder.Recordset.MoveFirst
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

    If LocomotiveDatabaseSearch.Recordset.Fields(0).Value <> "" Then

        If (checkboxactivelist = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(2).Value = "True") Or _
           (checkboxdecoderequiped.Value = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(3).Value = "True") Then
            
            For Y = 0 To TabCVList.TabIndex + 2
            
                Let Grid(Y).Rows = temporaryRow + 1
                Let Grid(Y).Row = temporaryRow
        
                If LocomotiveDatabaseSearch.Recordset.Fields(0).Value <> "" Then
                    Let Grid(Y).Col = 0
                    Let Grid(Y).Text = LocomotiveDatabaseSearch.Recordset.Fields(0).Value
                    ' -------------------------------------------------------------------------------------------------------------
                    For X = 1 To 16
                        Let z = Y * 16 + X
                        If LocomotiveDatabaseDecoder.Recordset.Fields(z).Value <> "" Then
                            Let Grid(Y).Col = X
                            Let Grid(Y).Text = CStr(LocomotiveDatabaseDecoder.Recordset.Fields(z).Value)
                        End If
                    Next X
                    
                End If
            Next Y
            Let temporaryRow = temporaryRow + 1
        End If
    End If
    LocomotiveDatabaseSearch.Recordset.MoveNext
    LocomotiveDatabaseDecoder.Recordset.MoveNext
    
Wend

Let MousePointer = Default
Let ButtonClose.Enabled = True
Let ButtonPrint.Enabled = True

End Sub

Private Sub ButtonPrint_Click()

    LocomotiveCVSpreadsheet.PrintForm
    
End Sub



Private Sub ButtonPrintList_Click()

MousePointer = vbHourglass

Report.Clear
Report.SetColNums 17
Report.SetFixedLeftMargin 0

Report.HeadLine = True
Report.SetPageHeadFormat "Arial", 24, 0, RGB(0, 0, 255), False, False, False
Report.SetPageHead "Automatic Train Control - Configuration Variable List     Date:" & Date


Report.SetReportTitleHeight 10
Report.SetReportSubTitleHeight 0
                        
Report.SetWidthString "49|39|39|39|39|39|39|39|39|39|39|39|39|39|39|39|39"

Let Temporary$ = "Number"
For X = Val(TabCVList.Tab) * 16 + 1 To Val(TabCVList.Tab) * 16 + 16
    Let Temporary$ = Temporary$ + "|CV" + CStr(X)
Next X
    
Report.SetTopString Temporary$
Report.SetBottomString Temporary$
Report.SetReportHeadHeight 20
             
LocomotiveDatabaseSearch.Recordset.MoveFirst
LocomotiveDatabaseDecoder.Recordset.MoveFirst

While Not LocomotiveDatabaseSearch.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    If LocomotiveDatabaseSearch.Recordset.Fields(0).Value <> "" Then
        
        If (checkboxactivelist = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(2).Value = "True") Or _
           (checkboxdecoderequiped.Value = vbChecked And LocomotiveDatabaseSearch.Recordset.Fields(3).Value = "True") Then
           
            Let temp$ = LocomotiveDatabaseSearch.Recordset.Fields(0).Value
                            
            For X = Val(TabCVList.Tab) * 16 + 1 To Val(TabCVList.Tab) + 1 * 16
                 If LocomotiveDatabaseDecoder.Recordset.Fields(X).Value <> "" Then
                    Let temp$ = temp$ & "|" & LocomotiveDatabaseDecoder.Recordset.Fields(X).Value
                 'Else
                 '    Let temp$ = temp$ = "| "
                 End If
            Next X
            Report.AddRecord temp$
        End If
    End If
    LocomotiveDatabaseSearch.Recordset.MoveNext
    LocomotiveDatabaseDecoder.Recordset.MoveNext
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
        If TemporaryScreen = "Locomotive CV Spreadsheet Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Locomotive CV Spreadsheet Screen"
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
            Let Ini.Value = "Locomotive CV Spreadsheet Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Locomotive CV Spreadsheet Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Locomotive CV Spreadsheet Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Locomotive CV Spreadsheet Screen"
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
    Let Ini.Parameter = "CheckboxAddressShort"
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
                Let Ini.Value = "Locomotive CV Spreadsheet Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Locomotive CV Spreadsheet Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    LocomotiveCVSpreadsheet.Hide
    'unload LocomotiveCVSpreadsheet

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
    Let Ini.Application = "Locomotive CV Spreadsheet Screen"
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
        LocomotiveCVSpreadsheet.Left = (Screen.Width - Width) / 2
        LocomotiveCVSpreadsheet.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + LocomotiveCVSpreadsheet.Width > Screen.Width Then
            Let LocomotiveCVSpreadsheet.Left = Screen.Width - LocomotiveCVSpreadsheet.Width
        Else
            Let LocomotiveCVSpreadsheet.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + LocomotiveCVSpreadsheet.Height > Screen.Height Then
            Let LocomotiveCVSpreadsheet.Top = Screen.Height - LocomotiveCVSpreadsheet.Height
        Else
            Let LocomotiveCVSpreadsheet.Top = Val(TemporaryValueTop)
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
        
        Let TemporaryText1 = "This checkbox allows you to select locomotives" & vbCrLf & "(numbers) which are flagged as being active in the database."
        Let TemporaryText2 = "Active Locomotive List"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(checkboxactivelist)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxactivelist, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This checkbox allows you to select locomotives" & vbCrLf & "(numbers) which are flagged as being decoder equipped" & vbCrLf & "in the database."
        Let TemporaryText2 = "Decoder Equipped Locomotive List"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(checkboxdecoderequiped)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(checkboxdecoderequiped, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This checkbox allows you to select locomotives" & vbCrLf & "(numbers) with a short address; otherwise a long" & vbCrLf & "address database will be used."
        Let TemporaryText2 = "Short Addresses"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(CheckBoxAddressShort)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(CheckBoxAddressShort, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This tab allows you to select the group" & vbCrLf & "of configuration variables values to view."
        Let TemporaryText2 = "Tab Group of CVs"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TabCVList)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(TabCVList, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This button compiles the list of locomotives" & vbCrLf & "(numbers) in your database on screen."
        Let TemporaryText2 = "Compile Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonCompile)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonCompile, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        Let TemporaryText1 = "This button closes the Locomotive Configuration" & vbCrLf & "Speadsheet window and returns you to the main screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
'        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let LocomotiveDatabaseSearch.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDiesels.mdb"
    Let LocomotiveDatabaseDecoder.DatabaseName = App.Path$ & "\Databases\LocomotiveDatabaseDecoders.mdb"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Set Grid
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
For Y = 0 To TabCVList.TabIndex + 2
    Let Grid(Y).ColWidth(0) = 800
    Let Grid(Y).Row = 0
    Let Grid(Y).Col = 0
    Let Grid(Y).Text = "Locomotive"
    For X = 1 To 16
        Let Grid(Y).Col = X
        Let Grid(Y).Text = "CV" + CStr(X + Y * 16)
    Next X
Next Y
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

    If LocomotiveCVSpreadsheet.WindowState = vbMinimized Then
    
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
        
    ElseIf LocomotiveCVSpreadsheet.WindowState = vbNormal Then
    
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

Private Sub LocomotiveDatabaseDecoder_Reposition()
Let LocomotiveDatabaseDecoder.Caption = LocomotiveDatabaseDecoder.Recordset.AbsolutePosition
End Sub

Private Sub LocomotiveDatabaseSearch_Reposition()
    Let LocomotiveDatabaseSearch.Caption = LocomotiveDatabaseSearch.Recordset.AbsolutePosition
End Sub

