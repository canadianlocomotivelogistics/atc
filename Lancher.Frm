VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Launcher"
   ClientHeight    =   4380
   ClientLeft      =   5625
   ClientTop       =   5880
   ClientWidth     =   9270
   Height          =   4785
   Left            =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9270
   Top             =   5535
   Width           =   9390
   Begin VB.CheckBox CheckboxTransferInProgress 
      Caption         =   "Transfer in Progress"
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton ButtonAbort 
      Caption         =   "&Abort"
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox TextBoxStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Text            =   "Status:"
      Top             =   840
      Width           =   9015
   End
   Begin VB.Timer TimerLancher 
      Interval        =   1000
      Left            =   7980
      Top             =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "Lancher.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin Balloon_OCX.BalloonOCX BalloonOCX 
      Left            =   5880
      Top             =   600
      _ExtentX        =   450
      _ExtentY        =   238
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   9000
      Top             =   60
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin FATHFTPLibCtl.FathFTPCtrl FtpControl2 
      Left            =   5820
      OleObjectBlob   =   "Lancher.frx":0442
      Top             =   60
   End
   Begin ComctlLib.ListView ListViewFtp 
      Height          =   2715
      Left            =   180
      TabIndex        =   4
      Top             =   1260
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   4789
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filesize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Date"
         Object.Width           =   3068
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   4764
      EndProperty
   End
   Begin SystemInfoControl.MSysInfo SystemInfoControl 
      Left            =   8460
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      Caption         =   $"Lancher.frx":0512
      Height          =   615
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   5055
   End
   Begin April16_NetConnect.NetConnect NetConnectControl 
      Left            =   6900
      Top             =   60
      _ExtentX        =   873
      _ExtentY        =   873
      LicenceCode     =   "7134424019415563"
      LicenceName     =   "Canadian Locomotive Logistics"
   End
   Begin IniconLib.Init IniControl 
      Left            =   7440
      Top             =   60
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin FATHFTPLibCtl.FathFTPCtrl FtpControl 
      Left            =   6360
      OleObjectBlob   =   "Lancher.frx":05B9
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub


Private Sub ButtonAbort_Click()

    End
    
End Sub

Private Sub Form_Activate()

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
                Let Ini.Value = "Atc Updater, Form, Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Atc Updater, Form, Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

End Sub

Private Sub Form_Load()

End Sub

Private Sub FtpControl_EnumFile(ByVal Filename As String, ByVal FileSize As Double, ByVal User As String, ByVal Group As String, ByVal FileDate As Date, ByVal FileFlags As String)

Dim TemporaryBytes As Double

Set TemporaryList = listviewftp.ListItems.Add(, , Filename)
    TemporaryList.SubItems(1) = Str(FileSize)
    TemporaryList.SubItems(2) = FileDate
    TemporaryList.SubItems(3) = "Pending"

    Dim TemporaryDate As String
    Let TemporaryDate = CDate(SystemInfoControl.FileDate & " " & SystemInfoControl.FileTime)
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Compare Dates
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If IsDate(TemporaryDate) = True Then
        If FileDate > TemporaryDate Then
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Enumerated File is Newer Version
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TemporaryList.SubItems(3) = "Update Needed"
            FtpControl2.Connect
            Let TemporaryList.SubItems(3) = "Downloading Update"
            FtpControl2.Get Filename, SystemInfoControl.Filename, 0
            While CheckboxTransferInProgress.Value = vbChecked
                Debug.Print x: x = x + 1
                DoEvents
            Wend
            Let TemporaryList.SubItems(3) = "Update Completed"
        Else
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Enumerated File is Older
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TemporaryList.SubItems(3) = "No Update Needed"
        End If
    Else
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' TemporaryDate is not a valid Date
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let IniControl.Filename = App.Path$ & "\ATC.log"
        Let IniControl.Application = "Main Screen"
        Let IniControl.Parameter = "Log File"
        If Ini.Value = "On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\ATC.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Launcher, FTPControl, EnumFile, date of file is an invalid format."
        End If
    End If
    listviewftp.ListItems.Item(listviewftp.ListItems.Count).Selected = True


End Sub

Private Sub FtpControl2_TransferComplete(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal BytesTransfered As Long)

    Let CheckboxTransferInProgress.Value = vbUnchecked
    
End Sub

Private Sub FtpControl2_TransferStarting(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal Offset As Long)

    Let CheckboxTransferInProgress.Value = vbChecked
    
End Sub


Private Sub listviewftp_ColumnClick(ByVal ColumnHeader As ColumnHeader)

    listviewftp.SortKey = ColumnHeader.Index - 1
    listviewftp.Sorted = True

End Sub


Private Sub listviewftp_ItemClick(ByVal Item As ListItem)

Debug.Print Item



End Sub


Private Sub FtpControl_StateChange(ByVal State As Long)
    
    Select Case State
        Case FFTP_STATE_AUTHORIZING
            Let TextBoxStatus.Text = "Status: Authorizing"
        Case FFTP_STATE_CONNECTED
            Let TextBoxStatus.Text = "Status: Connected"
        Case FFTP_STATE_CONNECTING
            Let TextBoxStatus.Text = "Status: Connecting"
        Case FFTP_STATE_DOWNLOADING
            Let TextBoxStatus.Text = "Status: Downloading"
        Case FFTP_STATE_LISTING
            Let TextBoxStatus.Text = "Status: Listing"
        Case FFTP_STATE_LOGGEDIN
            Let TextBoxStatus.Text = "Status: Logged In"
        Case FFTP_STATE_NOTCONNECTED
            Let TextBoxStatus.Text = "Status: Not Connected"
        Case FFTP_STATE_PROCESSING
            Let TextBoxStatus.Text = "Status: Processing"
        Case FFTP_STATE_UPLOADING
            Let TextBoxStatus.Text = "Status: Uploading"
    End Select

End Sub

Private Sub TimerLancher_Timer()

    Let TimerLancher.Interval = 0

    If NetConnectControl.Connected Then
        
        Let IniControl.Filename = App.Path$ & "ATC.ini"
        Let IniControl.Application = "Main Screen"
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Connect to the Server
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        On Error Resume Next
        FtpControl.Connect
        If Err <> 0 Then
        On Error GoTo 0
            Let IniControl.Filename = App.Path$ & "\ATC.log"
            Let IniControl.Application = "Main Screen"
            Let IniControl.Parameter = "Log File"
            Let TemporaryMessage = "An problem has occured with Automatic Train Control. This problem will be recorded in the ATC.LOG file." & vbCrLf & "Internet Settings Screen, Winsock Error, error " & Number & " " & Description & "." & vbCrLf & vbCrLf & "This problem is caused by the following situations:" & vbCrLf & "    The server you are trying to connect to is busy (connected to another person, please try again later) or" & vbCrLf & "    the server you are trying to connect to is down (not responding)."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\ATC.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Launcher, TimerLauncher, Timer, server timed out."
        Else
            On Error GoTo 0
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Files and Enumerated Files
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If FtpControl.FileExists("Atc.exe") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.exe"
                FtpControl.EnumFiles ("Atc.exe")
            End If
            If FtpControl.FileExists("Atc.ini") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.ini"
                FtpControl.EnumFiles ("Atc.ini")
            End If
            If FtpControl.FileExists("Atc.tip") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.tip"
                FtpControl.EnumFiles ("Atc.tip")
            End If
            If FtpControl.FileExists("Atc.log") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.log"
                FtpControl.EnumFiles ("Atc.log")
            End If
            If FtpControl.FileExists("Atc.bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.bmp"
                FtpControl.EnumFiles ("Atc.bmp")
            End If
            If FtpControl.FileExists("\Help\Atc.chm") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Help\Atc.chm"
                FtpControl.EnumFiles ("\Help\Atc.chm")
            End If
            If FtpControl.FileExists("Atc.prx") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.prx"
                FtpControl.EnumFiles ("Atc.prx")
            End If
            If FtpControl.FileExists("Atc.mac") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Atc.mac"
                FtpControl.EnumFiles ("Atc.mac")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake3(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake3(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake3(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake4(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake4(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake4(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake5(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake5(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake5(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake6(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake6(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake6(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake7(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake7(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake7(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake8(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake8(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake8(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake9(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake9(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake9(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake10(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake10(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake10(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake11(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake11(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake11(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake12(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake12(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake12(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake13(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake13(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake13(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake14(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake14(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake14(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\AutomaticBrake15(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\AutomaticBrake15(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\AutomaticBrake15(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Bell0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Bell0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Bell0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Bell1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Bell1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Bell1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\CabScreen(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\CabScreen(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\CabScreen(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\CabWindow(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\CabWindow(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\CabWindow(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake3(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake3(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake3(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake4(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake4(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake4(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake5(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake5(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake5(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake6(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake6(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake6(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake7(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake7(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake7(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\DynamicBrake8(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\DynamicBrake8(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\DynamicBrake8(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Horn0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Horn0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Horn0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Horn1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Horn1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Horn1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake3(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake3(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake3(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake4(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake4(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake4(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake5(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake5(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake5(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake6(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake6(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake6(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake7(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake7(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake7(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake8(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake8(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake8(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake9(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake9(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake9(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake10(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake10(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake10(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake11(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake11(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake11(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake12(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake12(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake12(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake13(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake13(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake13(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake14(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake14(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake14(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake15(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake15(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake15(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake16(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake16(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake16(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake17(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake17(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake17(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake18(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake18(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake18(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake19(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake19(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake19(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake20(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake20(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake20(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\IndependentBrake21(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\IndependentBrake21(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\IndependentBrake21(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Light0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Light0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Light0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Light1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Light1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Light1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Light2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Light2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Light2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\RadioPhone0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\RadioPhone0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\RadioPhone0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\RadioPhone1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\RadioPhone1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\RadioPhone1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\ResetLeft0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\ResetLeft0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\ResetLeft0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\ResetRight0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\ResetRight0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\ResetRight0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\ResetRight1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\ResetRight1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\ResetRight1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Reverser0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Reverser0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Reverser0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Reverser1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Reverser1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Reverser1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Reverser2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Reverser2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Reverser2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Sand0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Sand0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Sand0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Sand1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Sand1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Sand1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Speedometer(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Speedometer(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Speedometer(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle0(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle0(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle0(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle1(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle1(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle1(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle2(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle2(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle2(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle3(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle3(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle3(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle4(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle4(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle4(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle5(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle5(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle5(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle6(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle6(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle6(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle7(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle7(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle7(d1).bmp")
            End If
            If FtpControl.FileExists("\Graphics\Locomotive Diesel1\Throttle8(d1).bmp") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Graphics\Locomotive Diesel1\Throttle8(d1).bmp"
                FtpControl.EnumFiles ("\Graphics\Locomotive Diesel1\Throttle8(d1).bmp")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Bell.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Bell.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Bell.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\BrakeAutomaticClose.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticClose.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\BrakeAutomaticClose.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\BrakeAutomaticOpen.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\BrakeAutomaticOpen.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\BrakeAutomaticOpen.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\BrakeDynamic.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\BrakeDynamic.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\BrakeDynamic.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\BrakeIndependentClose.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentClose.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\BrakeIndependentClose.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\BrakeIndependentOpen.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\BrakeIndependentOpen.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\BrakeIndependentOpen.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Control.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Control.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Control.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\ControlBrakeDynamic.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\ControlBrakeDynamic.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\ControlDynamicBrake.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\ControlLight.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\ControlLight.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\ControlLight.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\ControlReverser.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\ControlReverser.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\ControlReverser.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\ControlSand.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\ControlSand.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\ControlSand.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\ControlThrottle.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\ControlThrottle.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\ControlThrottle.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Deadman.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Deadman.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Deadman.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Ge_p3.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Ge_p3.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Ge_p3.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Horn.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Horn.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Horn.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\PermissionDenied.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\PermissionDenied.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\PermissionDenied.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\PermissionGranted.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\PermissionGranted.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\PermissionGranted.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\PermissionOkProceed.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\PermissionOkProceed.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\PermissionOkProceed.wav")
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Sand.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Sand.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Sand.wav")
            End If
            End If
            If FtpControl.FileExists("\Sounds\Graphics\Sand.wav") Then
                Let SystemInfoControl.Filename = App.Path$ & "\Sounds\Graphics\Sand.wav"
                FtpControl.EnumFiles ("\Sounds\Graphics\Sand.wav")
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Disconnect from the Server
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            FtpControl.Disconnect
        End If
    End If
    
    Let ButtonAbort.Caption = "Close"
    
    Shell ("ATC.exe"), vbNormalFocus
    
    End
    
End Sub


