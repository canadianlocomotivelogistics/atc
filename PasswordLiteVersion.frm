VERSION 4.00
Begin VB.Form PasswordLiteVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATC - Password Authenitcation"
   ClientHeight    =   7815
   ClientLeft      =   6990
   ClientTop       =   2265
   ClientWidth     =   5490
   Height          =   8220
   Icon            =   "PasswordLiteVersion.frx":0000
   Left            =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5490
   Top             =   1920
   Width           =   5610
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox TextboxUserName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6060
      Top             =   1560
   End
   Begin VB.PictureBox PictureIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "PasswordLiteVersion.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonAbort 
      Caption         =   "&Abort"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox TextboxPassword 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   6960
      Width           =   3495
   End
   Begin VB.TextBox TextboxDate 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox TextboxVersion 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox TextboxComputerName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3495
   End
   Begin FILETRANSXLib.FileTransX FtpControl 
      Height          =   480
      Left            =   6060
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Timeout         =   30
      ServerName      =   "192.168.1.100"
      Username        =   "softwarepasswords"
      Password        =   "walnuttree12"
      ProxyName       =   ""
      ProxyUserID     =   ""
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   6120
      Top             =   2640
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx2 
      Height          =   195
      Left            =   780
      Top             =   2820
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   344
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
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx1 
      Height          =   195
      Left            =   1740
      Top             =   1380
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   344
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
      EmailSubject    =   "Automatic Train Control - Password Authenication"
   End
   Begin SystemInfoControl.MSysInfo SystemInformationOCX 
      Left            =   6060
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label12 
      Caption         =   "User Authentication"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   5400
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   5400
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   5400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Name"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   5400
      X2              =   120
      Y1              =   3840
      Y2              =   3840
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   6060
      Top             =   1080
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin IniconLib.Init Ini 
      Left            =   6060
      Top             =   540
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
   Begin URLLabelCtl.URLLabel UrlLabel2 
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   2805
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
   Begin VB.Label Label10 
      Caption         =   $"PasswordLiteVersion.frx":0884
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   2220
      Width           =   5295
   End
   Begin VB.Label Label9 
      Caption         =   $"PasswordLiteVersion.frx":0967
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label8 
      Caption         =   ", please include all the"
      Height          =   255
      Left            =   3300
      TabIndex        =   16
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "When you email me at"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label LabelWindowDescription 
      Caption         =   $"PasswordLiteVersion.frx":0A0F
      Height          =   1155
      Left            =   720
      TabIndex        =   14
      Tag             =   "0"
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5400
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Password"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Computer's Name"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "User Indentification"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date and Time of File"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Software Version"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Automatic Train Control "
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   5295
   End
End
Attribute VB_Name = "PasswordLiteVersion"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub ButtonAbort_Click()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Password Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(Password.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(Password.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(Password.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(Password.Height)
    Let Ini.Parameter = "SoftwareVersion"
    Let Ini.Value = TextboxVersion.Text
    Let Ini.Parameter = "DateTime"
    Let Ini.Value = TextboxDate.Text
    Let Ini.Parameter = "ComputerName"
    Let Ini.Value = TextBoxComputerName.Text
    Let Ini.Parameter = "UserName"
    Let Ini.Value = TextBoxUserName.Text
    Let Ini.Parameter = "Password"
    Let Ini.Value = TextboxPassword.Text

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
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unload Current Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Password.Hide

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Terminate Program
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


Private Sub ButtonClose_Click()
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Password
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryNumber As Integer
    Dim TemporaryAsciiTotal As Long
    Let TemporaryAsciiTotal = 0
    
    For TemporaryNumber = 1 To Len(TextBoxUserName.Text)
        Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextBoxUserName.Text, TemporaryNumber, 1))
    Next TemporaryNumber
    For TemporaryNumber = 1 To Len(TextBoxComputerName.Text)
        Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextBoxComputerName.Text, TemporaryNumber, 1))
    Next TemporaryNumber
    For TemporaryNumber = 1 To Len(TextboxDate.Text)
        Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextboxDate.Text, TemporaryNumber, 1))
    Next TemporaryNumber
    For TemporaryNumber = 1 To Len(TextboxVersion.Text)
        Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextboxVersion.Text, TemporaryNumber, 1))
    Next TemporaryNumber
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Password
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If TemporaryAsciiTotal ^ 4 <> Val(TextboxPassword.Text) Then
        Let TextboxPassword.Text = "Type your new password here."
        Let LabelStatus.Caption = "Status: Updated Password Needed"
        TemporaryResponse = MsgBox("This software is not registered with Canadian Locomotive Logistics. Please email the author of this program to request your password. You may use the link provided on this window for a quick response. Do you wish to register at this time?", vbExclamation + vbYesNo, "Passowrd Authenication Request")
        If TemporaryResponse = vbNo Then
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Allow for ninety days grace
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If DateDiff("d", SystemInformationOcx.FileDate, Date) < 0 Or DateDiff("d", SystemInformationOcx.FileDate, Date) > 90 Then
                MsgBox "This version of software has expired. Please email" & vbCrLf & "dcalcutt@rogers.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Version"
                Let ButtonClose.Enabled = False
            Else
                MsgBox "A trial period of" & Str(90 - (DateDiff("d", SystemInformationOcx.FileDate, Date))) & " days is granted. After the expiration you must register the software.", vbOKOnly, "Automatic Train Control - Thirty day trail period"
                Load MainScreen
                MainScreen.Show vbModeless
            End If
        End If
        
        
        
        Let HyperLabelOcx1.EmailSubject = "Automatic Train Control - Registration"
        Let HyperLabelOcx1.EmailBody = "For registration of Automatic Train Control software," & Chr$(13) & "the software version is " & Chr$(34) & TextboxVersion.Text & Chr$(34) & "," & Chr$(13) & "the Date and Time of File is " & Chr$(34) + TextboxDate.Text + Chr$(34) & "," & Chr$(13) & "the computer's name is" + Chr$(34) & TextBoxComputerName.Text & Chr$(34) & "," & Chr$(13) & "the users' name is" & Chr$(34) & TextBoxUserName.Text
    Else
        LabelStatus.Caption = "Status: Accepted"
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' If registered, software only valid for 365 days
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DateDiff("d", SystemInformationOcx.FileDate, Date) < 0 Or DateDiff("d", SystemInformationOcx.FileDate, Date) > 365 Then
            MsgBox "This version of software has expired. Please email" & vbCrLf & "dcalcutt@rogers.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Software"
            Let ButtonClose.Enabled = False
        End If
        
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display Main Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Load MainScreen
        MainScreen.Show vbModeless
    End If
    
End Sub

Private Sub ButtonPrint_Click()

    Password.PrintForm
    
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
                Let Ini.Value = "Password Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Passowrd Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
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
    Let Ini.Application = "Password Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(Password.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(Password.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(Password.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(Password.Height)
    Let Ini.Parameter = "SoftwareVersion"
    Let Ini.Value = TextboxVersion.Text
    Let Ini.Parameter = "DateTime"
    Let Ini.Value = TextboxDate.Text
    Let Ini.Parameter = "ComputerName"
    Let Ini.Value = TextBoxComputerName.Text
    Let Ini.Parameter = "UserName"
    Let Ini.Value = TextBoxUserName.Text
    Let Ini.Parameter = "Password"
    Let Ini.Value = TextboxPassword.Text

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
                Let Ini.Value = "Password Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Password Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Unload Current Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Password.Hide

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
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
    Let Ini.Application = "Password Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "Password"
    Let TextboxPassword.Text = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        Password.Left = (Screen.Width - Width) / 2
        Password.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + Password.Width > Screen.Width Then
            Let Password.Left = Screen.Width - Password.Width
        Else
            Let Password.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + Password.Height > Screen.Height Then
            Let Password.Top = Screen.Height - Password.Height
        Else
            Let Password.Top = Val(TemporaryValueTop)
        End If
    End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check the parameter of the log file in the ini file and adjust the menu accordingly
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "Transparency"
    Dim TemporaryTransparency As String
    Let TemporaryTransparency = Ini.Value
    If TemporaryTransparency = "On" Then
        Let AlphaBlend.Enabled = True
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
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "BalloonHelp"
    Dim TemporaryBalloonHelp As String
    Let TemporaryBalloonHelp = Ini.Value

    If TemporaryBalloonHelp = "True" Then
        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim BalloonFont As New StdFont
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "BalloonHelpFontName"
        Let Ini.Value = BalloonFont.Name
        Let Ini.Parameter = "BalloonHelpFontSize"
        Let BalloonFont.Size = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour1"
        Let Colour1 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour2"
        Let Colour2 = Ini.Value
        Let Ini.Parameter = "BalloonHelpColour3"
        Let Colour3 = Ini.Value
        
        Let TemporaryText1 = "This textbox displays your current version of" & vbCrLf & "software for Automatic Train Control. It is" & vbCrLf & "needed to make a password."
        Let TemporaryText2 = "Software Version"
        'i = BalloonHelp.DestroyToolTip(TextboxVersion)
        i = BalloonHelp.AddToolTip(TextboxVersion, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(225, 225, 128), 0, 5000, 100, True, False, False, 200, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This textbox displays the time/date stamp for the " & vbCrLf & "software, Automatic Train Control. It is" & vbCrLf & "needed to make a password."
        Let TemporaryText2 = "Date and Time of File"
        'i = BalloonHelp.DestroyToolTip(TextboxDate)
        i = BalloonHelp.AddToolTip(TextboxDate, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This textbox displays your current computer name" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let TemporaryText2 = "Your Computer's Name"
        'i = BalloonHelp.DestroyToolTip(TextBoxComputerName)
        i = BalloonHelp.AddToolTip(TextBoxComputerName, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This textbox displays your current user name" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let TemporaryText2 = "Your User Name"
        'i = BalloonHelp.DestroyToolTip(TextBoxUserName)
        i = BalloonHelp.AddToolTip(TextBoxUserName, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
         
        Let TemporaryText1 = "This textbox displays your for Automatic Train" & vbCrLf & "Control. Type in the here and 'Close' the" & vbCrLf & "window or 'Abort' to exit the program."
        Let TemporaryText2 = "Password"
        'i = BalloonHelp.DestroyToolTip(TextboxPassword)
        i = BalloonHelp.AddToolTip(TextboxPassword, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "This button aborts the program. If you do not have" & vbCrLf & "a passowrd, please 'click' on my email address for" & vbCrLf & "a password."
        Let TemporaryText2 = "Abort Button"
        'i = BalloonHelp.DestroyToolTip(buttonabort)
        i = BalloonHelp.AddToolTip(buttonabort, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button prints the current window to your printer."
        Let TemporaryText2 = "Print Button"
        'i = BalloonHelp.DestroyToolTip(ButtonPrint)
        i = BalloonHelp.AddToolTip(ButtonPrint, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
        Let TemporaryText1 = "This button close the Authenication" & vbCrLf & "window and returns you to the main screen, if you have" & vbCrLf & "a correct password."
        Let TemporaryText2 = "Close Button"
        'i = BalloonHelp.DestroyToolTip(ButtonClose)
        i = BalloonHelp.AddToolTip(ButtonClose, TemporaryText1, balBalloon, TemporaryText2, balInfo, RGB(255, 255, 128), 0, 5000, 100, True, False, False, 255, BalloonFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
    ElseIf TemporaryBalloonHelp = "False" Then
        i = BalloonHelp.DestroyAllToolTips
    Else
        If MainScreen!MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Password Screen, Load Form, variable error in ATC.INI file for 'Balloon Help' setting."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fill Textboxes with Paramters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let SystemInformationOcx.Drive = "c:"
    Let SystemInformationOcx.Filename = App.Path$ & "\Atc.exe"
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryType As String
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Type"
    Let TemporaryType = Ini.Value
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let TextboxVersion = SystemInformationOcx.FileVersion & " (" & TemporaryType & " Version)"
    Let TextboxDate = SystemInformationOcx.FileDate & " " & SystemInformationOcx.FileTime
    Let TextBoxComputerName.Text = SystemInformationOcx.ComputerName
    Let TextBoxUserName.Text = SystemInformationOcx.UserName

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


Private Sub HyperLabelOcx1_Click()

    Let HyperLabelOcx1.EmailSubject = "Automatic Train Control - Registration"
    Let HyperLabelOcx1.EmailBody = "For registration of Automatic Train Control software," & Chr$(13) & "the software version is " & Chr$(34) & TextboxVersion.Text & Chr$(34) & "," & Chr$(13) & "the Date and Time of File is " & Chr$(34) + TextboxDate.Text + Chr$(34) & "," & Chr$(13) & "the computer's name is" + Chr$(34) & TextBoxComputerName.Text & Chr$(34) & "," & Chr$(13) & "the users' name is" & Chr$(34) & TextBoxUserName.Text

End Sub

Private Sub TextboxDate_Change()

'    Let AsciiTotal = 0
    
'    For X = 1 To Len(TextboxName.Text)
'        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxName.Text, X, 1))
'    Next X
'
'    For X = 1 To Len(TextboxDate.Text)
'        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, X, 1))
'    Next X
'
'    For X = 1 To Len(TextboxVersion.Text)
'        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, X, 1))
'    Next X
  
End Sub

Private Sub TextboxName_Change()

    Let AsciiTotal = 0
    
    For x = 1 To Len(TextboxName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, x, 1))
    Next x
    
End Sub


Private Sub TextboxVersion_Change()

    'Let AsciiTotal = 0
    
    'For X = 1 To Len(TextBoxUserName.Text)
    '    Let AsciiTotal = AsciiTotal + Asc(Mid$(TextBoxUserName.Text, X, 1))
    'Next X
    
    'For X = 1 To Len(TextBoxComputerName.Text)
    '    Let AsciiTotal = AsciiTotal + Asc(Mid$(TextBoxComputerName.Text, X, 1))
    'Next X
    
    'For X = 1 To Len(TextboxDate.Text)
    '    Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, X, 1))
    'Next X
    
    'For X = 1 To Len(TextboxVersion.Text)
    '    Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, X, 1))
    'Next X
  
End Sub


Private Sub Timer1_Timer()

    Let Timer1.Interval = 0

    Call ButtonClose_Click
    
End Sub





