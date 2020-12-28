VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{43D53FA6-A170-4F41-9834-B92A67EDB899}#1.0#0"; "filetransx.ocx"
Object = "{625E24A3-B09D-101D-85F5-6EBA1EE93AF4}#3.3#0"; "INICON32.OCX"
Object = "{50DCCD2C-5336-44E7-836E-520996088209}#1.0#0"; "A16_NetConnect.ocx"
Object = "{6E3BBBCF-5A0D-11D1-B659-0060976C699F}#1.1#0"; "MSYSINFO.OCX"
Object = "{F7C7DD64-246B-4E83-B41C-F900FAFE437A}#104.0#0"; "Balloon.ocx"
Object = "{68EC193C-7292-41C9-853A-63CDCA210B3B}#5.0#0"; "AlphaBlend.ocx"
Begin VB.Form Updater 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Launcher"
   ClientHeight    =   4380
   ClientLeft      =   7935
   ClientTop       =   7305
   ClientWidth     =   9420
   Icon            =   "Updater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   9420
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
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
      Caption         =   "&Start"
      Height          =   255
      Left            =   6540
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
      Left            =   7980
      Top             =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "Updater.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin FILETRANSXLib.FileTransX FileTransControl 
      Height          =   480
      Left            =   6360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Timeout         =   30
      ServerName      =   "atc.developer.lovethosetrains.com"
      Username        =   "softwareupdate"
      Password        =   "walnuttree12"
      ProxyName       =   ""
      ProxyUserID     =   ""
   End
   Begin FILETRANSXLib.FileTransX FileTransControlAgain 
      Height          =   480
      Left            =   5820
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Timeout         =   30
      ServerName      =   "atc.developer.lovethosetrains.com"
      Username        =   "softwareupdate"
      Password        =   "walnuttree12"
      ProxyName       =   ""
      ProxyUserID     =   ""
   End
   Begin Balloon_OCX.BalloonOCX BalloonOCX 
      Left            =   5880
      Top             =   600
      _ExtentX        =   661
      _ExtentY        =   661
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
   Begin April16_NetConnect.NetConnect NetConnectControl 
      Left            =   6900
      Top             =   60
      _ExtentX        =   873
      _ExtentY        =   873
      LicenceCode     =   "7134424019415563"
      LicenceName     =   "Canadian Locomotive Logistics"
   End
   Begin IniconLib.Init Ini 
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
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   9000
      Top             =   60
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Label Label1 
      Caption         =   $"Updater.frx":0884
      Height          =   615
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   5055
   End
End
Attribute VB_Name = "Updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub


Private Sub ButtonAbort_Click()
    
    If ButtonAbort.Caption = "&Start" Then
        Let TimerLancher.Interval = 1
        Let ButtonAbort.Caption = "&Abort"
        Let ButtonClose.Enabled = False
    Else 'If ButtonAbort.Caption = "&Abort" Then
        Let ButtonAbort.Enabled = False
        Let ButtonClose.Enabled = True
    End If
    
End Sub

Private Sub ButtonClose_Click()

    End
    
    Call Form_Deactivate
    Updater.Hide
    Shell (App.Path$ & "\ATC.exe"), vbNormalFocus
    End

End Sub

Private Sub FileTransControl_EnumFile(ByVal Filename As String, ByVal FileSize As Double, ByVal User As String, ByVal Group As String, ByVal FileDate As Date, ByVal FileFlags As String)

Dim TemporaryBytes As Double

Set TemporaryList = ListViewFtp.ListItems.Add(, , Filename)
    TemporaryList.SubItems(1) = CStr(FileSize)
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
            If ButtonAbort.Enabled = False Then Exit Sub
            FileTransControlAgain.Connect
            If ButtonAbort.Enabled = False Then Exit Sub
            Let TemporaryList.SubItems(3) = "Downloading Update"
            If ButtonAbort.Enabled = False Then Exit Sub
            If SystemInfoControl.Filename = App.Path$ & "\AtcUpdater.exe" Then
                Let TemporaryPath = Right$(SystemInfoControl.Filename, Len(SystemInfoControl.Filename) - Len(App.Path$))
                FileTransControlAgain.Get TemporaryPath, App.Path$ & "\AtcUpdater.new", 0
            Else
                Let TemporaryPath = Right$(SystemInfoControl.Filename, Len(SystemInfoControl.Filename) - Len(App.Path$))
                FileTransControlAgain.Get TemporaryPath, SystemInfoControl.Filename, 0
            End If
            If ButtonAbort.Enabled = False Then Exit Sub
            While CheckboxTransferInProgress.Value = vbChecked
                If ButtonAbort.Enabled = False Then Exit Sub
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
        Let Ini.Filename = App.Path$ & "\ATC.log"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "Log File"
        If Ini.Value = "On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\ATC.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Launcher, FTPControl, EnumFile, date of file is an invalid format."
        End If
    End If
    
    ListViewFtp.ListItems.Item(ListViewFtp.ListItems.Count).Selected = True


End Sub

Private Sub FileTransControl_StateChange(ByVal State As Long)

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

Private Sub FileTransControlAgain_TransferComplete(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal BytesTransfered As Long)

    Let CheckboxTransferInProgress.Value = vbUnchecked
    
End Sub


Private Sub FileTransControlAgain_TransferStarting(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal Offset As Long)

    Let CheckboxTransferInProgress.Value = vbChecked
    
End Sub


Private Sub Form_Activate()

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
        If TemporaryScreen = "Updater Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Updater Screen"
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
            Let Ini.Value = "Updater Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Updater, Form, Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Updater, Form, Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Updater Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(Updater.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(Updater.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(Updater.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(Updater.Height)

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
                Let Ini.Value = "Updater Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Updater Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub


Private Sub Form_Load()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Updater Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        Updater.Left = (Screen.Width - Width) / 2
        Updater.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + Updater.Width > Screen.Width Then
            Let Updater.Left = Screen.Width - Updater.Width
        Else
            Let Updater.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + Updater.Height > Screen.Height Then
            Let Updater.Top = Screen.Height - Updater.Height
        Else
            Let Updater.Top = Val(TemporaryValueTop)
        End If
    End If

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the log file in the ini file and adjust the menu accordingly
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "Transparency"
    Dim TemporaryTransparency As String
    Let TemporaryTransparency = Ini.Value
    If TemporaryTransparency = "On" Then
        'Let MenuTransparency.Caption = "&Transparency is On"
        Let AlphaBlend.Enabled = True
    ElseIf TemporaryTransparency = "Off" Then
        'Let MenuTransparency.Caption = "&Transparency is Off"
        Let AlphaBlend.Enabled = False
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Updater Screen, Form Load, an incorrect value is stored in ATC.INI for the 'Transparency' file."
        End If
    End If
       
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the balloon help file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "BalloonHelp"
    Dim TemporaryBalloon As String
    Let TemporaryBalloon = Ini.Value
    If TemporaryBalloon = "True" Then
        'Let menuBalloonHelp.Caption = "&Balloon Help is On"
    ElseIf TemporaryBalloon = "False" Then
        'Let menuBalloonHelp.Caption = "&Balloon Help is Off"
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Updater Screen, Form Load, an incorrect value is stored in ATC.INI for the 'BalloonHelp' file."
        End If
    End If
    
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the Speech help file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "SpeechHelp"
    Dim TemporarySpeech As String
    Let TemporarySpeech = Ini.Value
    If TemporarySpeech = "True" Then
        'Let menuspeechHelp.Caption = "&Speech Help is On"
    ElseIf TemporarySpeech = "False" Then
        'Let menuspeechHelp.Caption = "&Speech Help is Off"
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Updater Screen, Form Load, an incorrect value is stored in ATC.INI for the 'SpeechHelp' file."
        End If
    End If
        
End Sub

Private Sub FtpControl_EnumFile(ByVal Filename As String, ByVal FileSize As Double, ByVal User As String, ByVal Group As String, ByVal FileDate As Date, ByVal FileFlags As String)

End Sub

Private Sub FtpControl2_TransferComplete(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal BytesTransfered As Long)

End Sub

Private Sub FtpControl2_TransferStarting(ByVal LocalFilename As String, ByVal RemoteFilename As String, ByVal Offset As Long)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub

Private Sub listviewftp_ColumnClick(ByVal ColumnHeader As ColumnHeader)

    ListViewFtp.SortKey = ColumnHeader.Index - 1
    ListViewFtp.Sorted = True

End Sub


Private Sub listviewftp_ItemClick(ByVal Item As ListItem)

Debug.Print Item



End Sub


Private Sub FtpControl_StateChange(ByVal State As Long)
    
End Sub

Private Sub TimerLancher_Timer()

    Let TimerLancher.Interval = 0

    If NetConnectControl.Connected Then
        
        Let Ini.Filename = App.Path$ & "ATC.ini"
        Let Ini.Application = "Main Screen"
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Connect to the Server
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        On Error Resume Next
        FileTransControl.Connect
        If Err <> 0 Then
            On Error GoTo 0
            Let Ini.Filename = App.Path$ & "\ATC.log"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "Log File"
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
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.exe": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.ini": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.tip": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.ini": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.ini": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.log": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.mac": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Atc.prx": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Help\Atc.chm": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake2(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake3(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake4(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake5(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake6(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake7(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake8(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake9(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake10(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake11(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake12(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake13(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake14(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\AutomaticBrake15(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Bell0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Bell1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ButtonSide0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ButtonSide1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\CabScreen(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\CabWindow(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake1(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake2(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake3(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake4(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake5(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake6(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake7(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\DynamicBrake8(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Horn0(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Horn1(d1).bmp"": GoSub Check"
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake3(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake4(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake5(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake6(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake7(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake8(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake9(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake10d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake11(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake12d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake13(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake14(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake15(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake16(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake17(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake18(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake19(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake20(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\IndependentBrake21(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Light0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Light1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Light2(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\RadioPhone0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\RadioPhone1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ResetLeft0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ResetLeft1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ResetRight0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\ResetRight1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Reverser0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Reverser1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Reverser2(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Sand0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Sand1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\SignalLightGreen(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\SignalLightYellow(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\SignalLightRed(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Speedometer(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\SSunVisor(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle0(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle1(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle2(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle3(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle4(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle5(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle6(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle7(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Throttle8(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Graphics\Locomotive Diesel1\Wipers(d1).bmp": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            
            Let tempFile = "Sounds\Graphics\Bell.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\BrakeAutomaticClose.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\BrakeAutomaticOpen.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\BrakeDynamic.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\BrakeIndependentClose.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\BrakeIndepententOpen.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\Control.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\ControlBrakeDynamic.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\ControlLight.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\ControlReverser.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\ControlSand.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\ControlThrottle.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\Deadman.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\Ge_p3.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\Horn.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\PermissionDenied.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\PermissionGranted.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\PermissionOkProceed.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\Sand.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            
            Let tempFile = "Sounds\Graphics\TrackJoint1.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint2.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint3.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint4.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint5.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint6.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint7.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint8.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint9.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint10.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint11.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint12.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint13.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint14.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint15.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint16.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint17.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint18.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint19.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint20.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint21.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint22.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            Let tempFile = "Sounds\Graphics\TrackJoint23.wav": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast
            
            Let tempFile = "AtcUpdateer": GoSub Check
            If ButtonAbort.Enabled = False Then GoTo ExitFast

            '------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Disconnect from the Server
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            FileTransControl.Disconnect
ExitFast:
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Reset Date Installed
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Opening Screen"
    Let Ini.Parameter = "InstallationDate"
    Dim TemporaryValueDate As String
    Let TemporaryValueDate = Ini.Value
    Let Ini.Value = Date
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Reset Email Addresses Passwords and Sponsors
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Password Screen"
    Let Ini.Parameter = "UsersEmailAddress"
    Let Ini.Value = ""
    Let Ini.Parameter = "SponsorsName"
    Let Ini.Value = "Type in your sponsor's name here."
    Let Ini.Parameter = "SponsorsID"
    Let Ini.Value = "Type in your sponsor's ID here."
    Let Ini.Parameter = "Password"
    Let Ini.Value = "Type in your password here."

    Let ButtonAbort.Enabled = False
    Let ButtonClose.Enabled = True
    Exit Sub
    
Check:
        If FileTransControl.FileExists(TempFilemame) Then
            Let SystemInfoControl.Filename = App.Path$ & "\" & tempFilename
            FileTransControl.EnumFiles (tempFilename)
        End If
    Return
End Sub


