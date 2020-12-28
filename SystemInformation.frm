VERSION 4.00
Begin VB.Form SystemInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - System Information"
   ClientHeight    =   8775
   ClientLeft      =   1725
   ClientTop       =   2205
   ClientWidth     =   4185
   Height          =   9180
   Icon            =   "SystemInformation.frx":0000
   Left            =   1665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   4185
   Top             =   1860
   Width           =   4305
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   1800
      TabIndex        =   64
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox TextBoxUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   63
      Top             =   7920
      Width           =   2040
   End
   Begin VB.TextBox TextBoxOSBuildOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   61
      Top             =   5040
      Width           =   2040
   End
   Begin VB.TextBox TextBoxOSBuild 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   59
      Top             =   4800
      Width           =   2040
   End
   Begin VB.TextBox TextBoxComputerName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   57
      Top             =   1920
      Width           =   2040
   End
   Begin VB.TextBox TextBoxWindowsPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   55
      Top             =   8160
      Width           =   2040
   End
   Begin VB.TextBox TextBoxTotalVirtual 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   54
      Top             =   7680
      Width           =   2040
   End
   Begin VB.TextBox TextBoxTotalPhysical 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   53
      Top             =   7440
      Width           =   2040
   End
   Begin VB.TextBox TextBoxTotalPage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   52
      Top             =   7200
      Width           =   2040
   End
   Begin VB.TextBox TextBoxTotalDiskSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   51
      Top             =   6960
      Width           =   2040
   End
   Begin VB.TextBox TextBoxTempPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   50
      Top             =   6720
      Width           =   2040
   End
   Begin VB.TextBox TextBoxSystemPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   49
      Top             =   6480
      Width           =   2040
   End
   Begin VB.TextBox TextBoxProcessorType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   48
      Top             =   6240
      Width           =   2040
   End
   Begin VB.TextBox TextBoxProcessorCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   47
      Top             =   6000
      Width           =   2040
   End
   Begin VB.TextBox TextBoxOSVersionMinor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   46
      Top             =   5760
      Width           =   2040
   End
   Begin VB.TextBox TextBoxOSVersionMajor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   45
      Top             =   5520
      Width           =   2040
   End
   Begin VB.TextBox TextBoxOSPlatform 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   44
      Top             =   5280
      Width           =   2040
   End
   Begin VB.TextBox TextBoxIsFileSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   43
      Top             =   4560
      Width           =   2040
   End
   Begin VB.TextBox TextBoxIsFileReadOnly 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   42
      Top             =   4320
      Width           =   2040
   End
   Begin VB.TextBox TextBoxIsFileHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   41
      Top             =   4080
      Width           =   2040
   End
   Begin VB.TextBox TextBoxIsFileArchived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   40
      Top             =   3840
      Width           =   2040
   End
   Begin VB.TextBox TextBoxFileVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   39
      Top             =   3600
      Width           =   2040
   End
   Begin VB.TextBox TextBoxFileTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   38
      Top             =   3360
      Width           =   2040
   End
   Begin VB.TextBox TextBoxFileSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1965
      TabIndex        =   37
      Top             =   3120
      Width           =   2040
   End
   Begin VB.TextBox TextBoxFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   36
      Top             =   2850
      Width           =   2040
   End
   Begin VB.TextBox TextBoxFileDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   35
      Top             =   2610
      Width           =   2040
   End
   Begin VB.TextBox TextBoxDriveType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   34
      Top             =   2355
      Width           =   2040
   End
   Begin VB.TextBox TextBoxDriveLetter 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   33
      Top             =   2130
      Width           =   2040
   End
   Begin VB.TextBox TextboxAvailableVirtual 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   32
      Top             =   1665
      Width           =   2040
   End
   Begin VB.TextBox TextBoxAvailablePhysical 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   31
      Top             =   1410
      Width           =   2040
   End
   Begin VB.TextBox TextBoxAvailablePage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   5
      Top             =   1200
      Width           =   2040
   End
   Begin VB.TextBox TextBoxAvailableDiskSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1935
      TabIndex        =   3
      Top             =   960
      Width           =   2040
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   8520
      Width           =   1095
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   4380
      Top             =   600
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin SystemInfoControl.MSysInfo SystemInformationOCX 
      Left            =   4320
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   4320
      Top             =   1800
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   4320
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
   Begin VB.Label LabelUserName 
      Caption         =   "User Name"
      Height          =   195
      Left            =   240
      TabIndex        =   62
      Top             =   7920
      Width           =   795
   End
   Begin VB.Label LabelOSBuildOptions 
      Caption         =   "OS Build Options"
      Height          =   195
      Left            =   240
      TabIndex        =   60
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label LabelOSBuild 
      Caption         =   "OS Build"
      Height          =   195
      Left            =   240
      TabIndex        =   58
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label LabelComputerName 
      Caption         =   "Computer Name"
      Height          =   195
      Left            =   240
      TabIndex        =   56
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Windows Path"
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   8160
      Width           =   1035
   End
   Begin VB.Label LabelTotalVirtual 
      Caption         =   "Total Virtual"
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   7680
      Width           =   840
   End
   Begin VB.Label LabelTotalPysical 
      Caption         =   "Total Physical"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   7440
      Width           =   990
   End
   Begin VB.Label LabelTotalPage 
      AutoSize        =   -1  'True
      Caption         =   "Total Page"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   7200
      Width           =   780
   End
   Begin VB.Label LabelTotalDiskSpace 
      Caption         =   "Total Disk Space"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   6960
      Width           =   1230
   End
   Begin VB.Label LabelTempPath 
      Caption         =   "Temp Path"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   6720
      Width           =   780
   End
   Begin VB.Label LabelSystemPath 
      Caption         =   "System Path"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   6480
      Width           =   885
   End
   Begin VB.Label LabelProcessorType 
      Caption         =   "Processor Type"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   6240
      Width           =   1110
   End
   Begin VB.Label LabelProcessorCount 
      Caption         =   "Processor Count"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   1170
   End
   Begin VB.Label LabelOSVersionMinor 
      Caption         =   "OS Version Minor"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Width           =   1230
   End
   Begin VB.Label LabelOSVersionMajor 
      Caption         =   "OS Version Major"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Width           =   1230
   End
   Begin VB.Label LabelOSPlatform 
      Caption         =   "OS Platform"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   5280
      Width           =   840
   End
   Begin VB.Label LabelSystem 
      Caption         =   "Is File System"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   4560
      Width           =   960
   End
   Begin VB.Label LabelReadOnly 
      Caption         =   "Is File Read Only"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label LabelIsHidden 
      Caption         =   "Is File Hidden"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   960
   End
   Begin VB.Label LabelIsArchived 
      Caption         =   "Is File Archived"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Label LabelFileVersion 
      Caption         =   "File Version"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label LabelFileTime 
      Caption         =   "File Time"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   630
   End
   Begin VB.Label LabelFileSize 
      Caption         =   "File Size"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label LabelFilename 
      Caption         =   "File Name"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   705
   End
   Begin VB.Label LabelFileDate 
      Caption         =   "File Date"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label LabelDriveType 
      Caption         =   "Drive Type"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label LabelDriveLetter 
      Caption         =   "Drive Letter"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label LabelAvailableVirtual 
      Caption         =   "Available Virtual"
      Height          =   195
      Left            =   225
      TabIndex        =   7
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label LabelAvailablePhysical 
      Caption         =   "Available Physical"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Available Page"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label LabelAvailableDiskSpace 
      Caption         =   "Available Disk Space"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   990
      Width           =   1515
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "SystemInformation.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"SystemInformation.frx":0884
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "SystemInformation"
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
            If TemporaryScreen = "System Information Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "System Information Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "System Information Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub



Private Sub ButtonPrint_Click()

    SystemInformation.PrintForm
    
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
        If TemporaryScreen = "System Information Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "System Information Screen"
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
            Let Ini.Value = "System Information Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "System Information Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "System Information Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    Call BalloonHelpUpdatePart01
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "System Information Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(SystemInformation.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(SystemInformation.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(SystemInformation.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(SystemInformation.Height)

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
                Let Ini.Value = "System Information Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "System Information Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    SystemInformation.Hide
    'unload systeminformation

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
    Let Ini.Application = "System Information Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        SystemInformation.Left = (Screen.Width - Width) / 2
        SystemInformation.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + SystemInformation.Width > Screen.Width Then
            Let SystemInformation.Left = Screen.Width - SystemInformation.Width
        Else
            Let SystemInformation.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + SystemInformation.Height > Screen.Height Then
            Let SystemInformation.Top = Screen.Height - SystemInformation.Height
        Else
            Let SystemInformation.Top = Val(TemporaryValueTop)
        End If
    End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    Else 'If MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Textboxes
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Application = "System Information Screen"
    Let SystemInformationOcx.Filename = App.Path$ & "\Atc.exe"
    Let SystemInformationOcx.Drive = "C:"
    
    Let textboxavailablediskspace.Text = SystemInformationOcx.AvailableDiskSpace
    Let Ini.Parameter = "AvailableDiskspace"
    Let Ini.Value = textboxavailablediskspace.Text
    
    Let TextBoxAvailablePage.Text = SystemInformationOcx.AvailablePage
    Let Ini.Parameter = "AvailablePage"
    Let Ini.Value = TextBoxAvailablePage.Text
    
    Let textboxavailablephysical.Text = SystemInformationOcx.AvailablePhysical
    Let Ini.Parameter = "AvailablePhysical"
    Let Ini.Value = textboxavailablephysical.Text
    
    Let TextboxAvailableVirtual.Text = SystemInformationOcx.AvailableVirtual
    Let Ini.Parameter = "AvailableVirtual"
    Let Ini.Value = TextboxAvailableVirtual.Text
    
    Let TextBoxComputerName.Text = SystemInformationOcx.ComputerName
    Let Ini.Parameter = "ComputerName"
    Let Ini.Value = TextBoxComputerName.Text
    
    Let TextBoxDriveLetter.Text = SystemInformationOcx.Drive
    Let Ini.Parameter = "Drive"
    Let Ini.Value = TextBoxDriveLetter.Text
    
     Let TextBoxDriveType.Text = SystemInformationOcx.DriveType
    Let Ini.Parameter = "DriveType"
    Let Ini.Value = TextBoxDriveType.Text
    
     Let textboxfiledate.Text = SystemInformationOcx.FileDate
    Let Ini.Parameter = "FileDate"
    Let Ini.Value = textboxfiledate.Text
    
     Let TextBoxFileName.Text = SystemInformationOcx.Filename
    Let Ini.Parameter = "Filename"
    Let Ini.Value = TextBoxFileName.Text
    
    Let TextBoxFileSize.Text = SystemInformationOcx.FileSize
    Let Ini.Parameter = "Filesize"
    Let Ini.Value = TextBoxFileSize.Text
    
     Let TextBoxFileTime.Text = SystemInformationOcx.FileTime
    Let Ini.Parameter = "FileTime"
    Let Ini.Value = TextBoxFileTime.Text
    
    Let TextBoxFileVersion.Text = SystemInformationOcx.FileVersion
    Let Ini.Parameter = "FileVersion"
    Let Ini.Value = TextBoxFileVersion.Text
    
     Let textboxisfilearchived.Text = SystemInformationOcx.IsArchived
    Let Ini.Parameter = "IsArchived"
    Let Ini.Value = textboxisfilearchived.Text
    
    Let TextBoxIsFileHidden.Text = SystemInformationOcx.IsHidden
    Let Ini.Parameter = "IsHidden"
    Let Ini.Value = TextBoxIsFileHidden.Text
    
     Let TextBoxIsFileReadOnly.Text = SystemInformationOcx.IsReadOnly
    Let Ini.Parameter = "IsReadOnly"
    Let Ini.Value = TextBoxIsFileReadOnly.Text
    
    Let TextBoxIsFileSystem.Text = SystemInformationOcx.IsSystem
    Let Ini.Parameter = "IsSystem"
    Let Ini.Value = TextBoxIsFileSystem.Text
    
    Let TextBoxOSBuild.Text = SystemInformationOcx.OSBuild
    Let Ini.Parameter = "OSBuild"
    Let Ini.Value = TextBoxOSBuild.Text
    
    Let TextBoxOSBuildOptions.Text = SystemInformationOcx.OSBuildOptions
    Let Ini.Parameter = "OSBuildOptions"
    Let Ini.Value = TextBoxOSBuildOptions.Text
    
    Let TextBoxOSPlatform.Text = SystemInformationOcx.OSPlatform
    Let Ini.Parameter = "OsPlatform"
    Let Ini.Value = TextBoxOSPlatform.Text
    
    Let TextBoxOSVersionMajor.Text = SystemInformationOcx.OSVersionMajor
    Let Ini.Parameter = "OSVersionMajor"
    Let Ini.Value = TextBoxOSVersionMajor.Text
    
    Let TextBoxOSVersionMinor.Text = SystemInformationOcx.OSVersionMinor
    Let Ini.Parameter = "OSVersionMinor"
    Let Ini.Value = TextBoxOSVersionMinor.Text
    
    Let TextBoxProcessorCount.Text = SystemInformationOcx.ProcessorCount
    Let Ini.Parameter = "ProcessorCount"
    Let Ini.Value = TextBoxProcessorCount.Text
    
    Let textboxprocessortype.Text = SystemInformationOcx.ProcessorType
    Let Ini.Parameter = "ProcessorType"
    Let Ini.Value = textboxprocessortype.Text
    
    Let TextBoxSystemPath.Text = SystemInformationOcx.SystemPath
    Let Ini.Parameter = "SystemPath"
    Let Ini.Value = TextBoxSystemPath.Text
    
    Let TextBoxTempPath.Text = SystemInformationOcx.TempPath
    Let Ini.Parameter = "TempPath"
    Let Ini.Value = TextBoxTempPath.Text
    
    Let TextBoxTotalDiskSpace.Text = SystemInformationOcx.TotalDiskSpace
    Let Ini.Parameter = "TotalDiskSpace"
    Let Ini.Value = TextBoxTotalDiskSpace.Text
    
    Let textboxtotalpage.Text = SystemInformationOcx.TotalPage
    Let Ini.Parameter = "TotalPage"
    Let Ini.Value = textboxtotalpage.Text
    
    Let textboxtotalphysical.Text = SystemInformationOcx.TotalPhysical
    Let Ini.Parameter = "TotalPhysical"
    Let Ini.Value = textboxtotalphysical.Text
    
    Let TextBoxTotalVirtual.Text = SystemInformationOcx.TotalVirtual
    Let Ini.Parameter = "TotalVirtual"
    Let Ini.Value = TextBoxTotalVirtual.Text
    
    Let TextBoxUserName.Text = SystemInformationOcx.UserName
    Let Ini.Parameter = "UserName"
    Let Ini.Value = TextBoxUserName.Text
    
    Let TextBoxWindowsPath.Text = SystemInformationOcx.WindowsPath
    Let Ini.Parameter = "WindowsPath"
    Let Ini.Value = TextBoxWindowsPath.Text

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

    If SystemInformation.WindowState = vbMinimized Then
    
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
        
    ElseIf SystemInformation.WindowState = vbNormal Then
    
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
    Let Ini.Application = "System Information Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(SystemInformation.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(SystemInformation.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(SystemInformation.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(SystemInformation.Height)
 
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
        Dim BalloonHelpWaveFile As String

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
        Else 'If mainscreen!MenuTransparency.Caption = "&Transparency is On" Then
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
        
        Let SystemInformation.MousePointer = ccHourglass
        
        Let TemporaryText1 = "This text box displays the available" & vbCrLf & "disk space on your computer."
        Let TemporaryText2 = "Available Disk Space"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxavailablediskspace)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxavailablediskspace, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxAvailableDiskSpace' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the available" & vbCrLf & "page space on your computer."
        Let TemporaryText2 = "Available Page"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxAvailablePage)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxAvailablePage, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxAvailablePage' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the available" & vbCrLf & "physical space on your computer."
        Let TemporaryText2 = "Available Physical Space"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxavailablephysical)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxavailablephysical, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxAvailablePysical' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays the available" & vbCrLf & "virtual space on your computer."
        Let TemporaryText2 = "Available Virtual Space"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxAvailableVirtual)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxAvailableVirtual, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxAvailableVirtual' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays your computer's" & vbCrLf & "name."
        Let TemporaryText2 = "Computer Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxComputerName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxComputerName, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxComputerName' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays the name of the computer drive" & vbCrLf & "this software in installed on."
        Let TemporaryText2 = "Drive Letter"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxDriveLetter)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxDriveLetter, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxDriveLeteter' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays the type of hard drive." & vbCrLf & "this software is installed on."
        Let TemporaryText2 = "Drive Type"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxDriveType)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxDriveType, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxDriveType' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays the date of the executable" & vbCrLf & "file for this software."
        Let TemporaryText2 = "File Date"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxfiledate)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxfiledate, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxFileDate' object."
            End If
        End If
        
        Let TemporaryText1 = "This text box displays the name of the executable" & vbCrLf & "file for this software."
        Let TemporaryText2 = "File Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxFileName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxFileName, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxFileName' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the size of the executable" & vbCrLf & "file for this software."
        Let TemporaryText2 = "File Size"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxFileSize)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxFileSize, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxFileSize' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the time of the executable" & vbCrLf & "file for this software."
        Let TemporaryText2 = "File Time"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxFileTime)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxFileTime, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxFileTime' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the version of the executable" & vbCrLf & "file for this software."
        Let TemporaryText2 = "File Version"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxFileVersion)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxFileVersion, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxFileVersion' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays if the executable file is " & vbCrLf & "archived."
        Let TemporaryText2 = "Is File Archived"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxisfilearchived)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxisfilearchived, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxIsFileArchived' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays if the executable file is " & vbCrLf & "hidden."
        Let TemporaryText2 = "Is File Hidden"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxIsFileHidden)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxIsFileHidden, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxIsFileHidden' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays if the executable file is " & vbCrLf & "read only."
        Let TemporaryText2 = "Is File Read Only"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxIsFileReadOnly)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxIsFileReadOnly, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxIsFileReadOnly' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays if the executable file is " & vbCrLf & "a system file."
        Let TemporaryText2 = "Is File System"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxIsFileSystem)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxIsFileSystem, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxIsFileSystem' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the operation system built" & vbCrLf & "number."
        Let TemporaryText2 = "Operating System Build"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxOSBuild)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxOSBuild, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxOSBuild' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the operation system options."
        Let TemporaryText2 = "Operating System Options"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxOSBuildOptions)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxOSBuildOptions, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxOSBuildOptions' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the operation system platform."
        Let TemporaryText2 = "Operating System Platform"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxOSPlatform)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxOSPlatform, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxOSPlatform' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the operation system major" & vbCrLf & "version number."
        Let TemporaryText2 = "Operating System Version (Major)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxOSVersionMajor)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxOSVersionMajor, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxOSVersionMajor' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the operation system minor" & vbCrLf & "version number."
        Let TemporaryText2 = "Operating System Version (Minor)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxOSVersionMinor)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxOSVersionMinor, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxOSVersionMinor' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the processor count."
        Let TemporaryText2 = "Processor Count"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxProcessorCount)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxProcessorCount, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxProcessorCount' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the processor type."
        Let TemporaryText2 = "Processor Type"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxprocessortype)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxprocessortype, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxProcessorType' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the system path on your" & vbCrLf & "computer."
        Let TemporaryText2 = "System Path"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxSystemPath)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxSystemPath, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxSystemPath' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the temporary path on your" & vbCrLf & "computer."
        Let TemporaryText2 = "Temporary Path"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxTempPath)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxTempPath, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxTempPath' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the total disk space" & vbCrLf & "on your computer."
        Let TemporaryText2 = "Total Disk Space"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxTotalDiskSpace)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxTotalDiskSpace, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxTotalDiskSpace' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the total page memory" & vbCrLf & "on your computer."
        Let TemporaryText2 = "Total Page Memory"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxtotalpage)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxtotalpage, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxTotalPage' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the total physical memory" & vbCrLf & "on your computer."
        Let TemporaryText2 = "Total Pysical Memory"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(textboxtotalphysical)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(textboxtotalphysical, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxTotalPysical' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the total virtual memory" & vbCrLf & "on your computer."
        Let TemporaryText2 = "Total Virtual Memory"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxTotalVirtual)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxTotalVirtual, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextBoxTotalVirtual' object."
            End If
        End If

        Let TemporaryText1 = "This text box displays the user name" & vbCrLf & "on your computer."
        Let TemporaryText2 = "User Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxUserName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxUserName, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TextboxUserName' object."
            End If
        End If

        Let SystemInformation.MousePointer = ccDefault
    
    Else 'If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is Off" Then
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
         If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'ButtonClose' object."
        End If
    End If

End Sub

Private Sub BalloonHelpUpdatePart02()

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
        Dim BalloonHelpWaveFile As String

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
        If MenuTransparency.Caption = "&Transparency is Off" Then
            BalloonHelpOpacity = 255
        Else 'If MenuTransparency.Caption = "&Transparency is On" Then
            Let Ini.Parameter = "BalloonHelpOpacity"
            Let BalloonHelpOpacity = Ini.Value
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If menuspeechHelp.Caption = "&Speech Help is Off" Then
            Let balloonhelp.Speech = False
        Else 'If menuspeechHelp.Caption = "&Speech Help is On" Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let SystemInformation.MousePointer = ccHourglass
        
        Let TemporaryText1 = "This text box displays the path to the windows operation" & vbCrLf & "system on your computer."
        Let TemporaryText2 = "Windows Path"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxWindowsPath)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxWindowsPath, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'TestboxWindowPath' object."
            End If
        End If
        
        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'ButtonPrint' object."
            End If
        End If

        Let TemporaryText1 = "This button closes the system information window" & vbCrLf & "and returns control to the main screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'ButtonClose' object."
            End If
        End If

        Let SystemInformation.MousePointer = ccDefault
    
    Else 'If MainScreen!menuBalloonHelp.Caption = "&Balloon Help is Off" Then
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
         If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "System Information, Balloon Help Update, unable to setup balloon help for 'ButtonClose' object."
        End If
    End If


End Sub
