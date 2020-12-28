VERSION 4.00
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATC - Password Authenitcation"
   ClientHeight    =   7365
   ClientLeft      =   8895
   ClientTop       =   3270
   ClientWidth     =   5445
   Height          =   7770
   Icon            =   "Password.frx":0000
   Left            =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   5445
   Top             =   2925
   Width           =   5565
   Begin VB.TextBox TextBoxUsersEmailAddress 
      Height          =   285
      Left            =   1860
      TabIndex        =   27
      Top             =   4500
      Width           =   3495
   End
   Begin VB.TextBox TextboxSponsorsID 
      Height          =   285
      Left            =   1860
      TabIndex        =   25
      Top             =   5700
      Width           =   3495
   End
   Begin VB.TextBox TextBoxSponsorsName 
      Height          =   285
      Left            =   1860
      TabIndex        =   23
      Top             =   5340
      Width           =   3495
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox TextboxUserName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4140
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6060
      Top             =   1560
   End
   Begin VB.PictureBox PictureIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "Password.frx":0442
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
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox TextboxPassword 
      Height          =   285
      Left            =   1860
      TabIndex        =   4
      Top             =   6600
      Width           =   3495
   End
   Begin VB.TextBox TextboxDate 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox TextboxVersion 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox TextboxComputerName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3495
   End
   Begin VB.Label LabelEmailAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Your Email Address"
      Height          =   195
      Left            =   60
      TabIndex        =   26
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Line Line5 
      X1              =   60
      X2              =   5400
      Y1              =   4860
      Y2              =   4860
   End
   Begin FATHMAILOCXLib.Message MessageOcx 
      Left            =   6000
      Top             =   4380
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin FATHMAILOCXLib.SMTP SmtpOcx 
      Left            =   6060
      Top             =   3840
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin VB.Label LabelSponsorsID 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Sponsor's ID"
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label LabelSponsorsName 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Sponsor's Name"
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   5340
      Width           =   1710
   End
   Begin VB.Label LabelSponsor 
      Caption         =   "Sponsor Identification"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   21
      Top             =   4980
      Width           =   2010
   End
   Begin FILETRANSXLib.FileTransX FtpControl 
      Height          =   480
      Left            =   6060
      TabIndex        =   20
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
      Left            =   60
      Top             =   1440
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   344
      ProjectKey      =   "et0B49E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SystemInfoControl.MSysInfo SystemInformationOCX 
      Left            =   6060
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label12 
      Caption         =   "User Authentication"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   5400
      Y1              =   6060
      Y2              =   6060
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Name"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   5340
      X2              =   60
      Y1              =   2220
      Y2              =   2220
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   6060
      Top             =   1080
      _ExtentX        =   767
      _ExtentY        =   767
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
   Begin VB.Label Label10 
      Caption         =   $"Password.frx":0884
      Height          =   615
      Left            =   60
      TabIndex        =   15
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label LabelWindowDescription 
      Caption         =   $"Password.frx":0962
      Height          =   675
      Left            =   720
      TabIndex        =   14
      Tag             =   "0"
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   5400
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Password"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   6660
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Computer's Name"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "User Identification"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5340
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date and Time of File"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Software Version"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2640
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Automatic Train Control "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   2280
      Width           =   2115
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   5235
   End
End
Attribute VB_Name = "Password"
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
                Let Ini.Value = "About Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
' Temporary Bypass Password Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load MainScreen
    MainScreen.Show vbModeless
    GoTo ExitOut
    
    Password.Enabled = False
    DoEvents
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check for valid email address
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(TextBoxUsersEmailAddress.Text, "@") = 0 Or _
       InStr(TextBoxUsersEmailAddress.Text, ".") = 0 Then _
       MsgBox "Please enter a valid email address for completion of registration.", vbexclaimation + vbOKOnly, "Automatic Train Control - Invalid Email Address": _
       GoTo ExitOut

    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Type"
    Dim TemporaryType As String
    Let TemporaryType = Ini.Value
    If TemporaryType <> "Pro" Then

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Password
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim TemporaryNumber As Integer
        Dim TemporaryAsciiTotal As Long
        Let TemporaryAsciiTotal = 0
        
        Let TemporaryAsciiTotal = 0
        
        If TemporaryType = "Full" Then
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
        Else ' if temporarytype = "Lite" then
            For TemporaryNumber = 1 To Len(TextboxSponsorsID.Text)
                Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextboxSponsorsID.Text, TemporaryNumber, 1))
            Next TemporaryNumber
            For TemporaryNumber = 1 To Len(TextBoxSponsorsName.Text)
                Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextBoxSponsorsName.Text, TemporaryNumber, 1))
            Next TemporaryNumber
            For TemporaryNumber = 1 To Len(TextboxDate.Text)
                Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextboxDate.Text, TemporaryNumber, 1))
            Next TemporaryNumber
            For TemporaryNumber = 1 To Len(TextboxVersion.Text)
                Let TemporaryAsciiTotal = TemporaryAsciiTotal + Asc(Mid$(TextboxVersion.Text, TemporaryNumber, 1))
            Next TemporaryNumber
        End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Password
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If TemporaryAsciiTotal ^ 3 <> Val(TextboxPassword.Text) Then
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Incorrect Password
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TextboxPassword.Text = "Type your new password here."
            Let LabelStatus.Caption = "Status: Updated Password Needed"
            
            If TemporaryType = "Full" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Full Version with Incorrect Password
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Let TemporaryResponse = MsgBox("The password for Automatic Train Control is incorrect, we at Canadian" & vbCrLf & _
                                               "Locomotive Logistics assume that Automatic Train Control (Full Version)" & vbCrLf & _
                                               "is not registered. Do you wish to register at this time?" & vbCrLf & vbCrLf & _
                                               "Registering the software with Canadian Locomotive Logistics will provide" & vbCrLf & _
                                               "you with a password.", vbExclamation + vbYesNo, "Password Authenication Request")
                If TemporaryResponse = vbNo Then
                ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' No Registration
                ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    'If DateDiff("d", format (SystemInformationOcx.FileDate,"mm-dd/yyy"), Date$) < 0 Or DateDiff("d", format(SystemInformationOcx.FileDate,"mm-dd-yyyy"), Date$) > 30 Then
                    '    MsgBox "This version of software has expired. Please email" & vbCrLf & "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Version"
                    '    Let ButtonClose.Enabled = False
                    'Else
                    '    MsgBox "A trial period of" & Str(90 - (DateDiff("d", SystemInformationOcx.FileDate, Date$))) & " days is granted. After the expiration you must register the software.", vbExclamation + vbOKOnly, "Automatic Train Control - Thirty day trail period"
                    '    Load MainScreen
                    '    MainScreen.Show vbModeless
                    'End If
                Else
                ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' With Registration
                ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let TemporaryResponse = MsgBox("Automatic Train Control (Full Version) will now send an email for" & vbCrLf & _
                                                   "registration of the software. Please wait.", vbExclamation + vbOKOnly, "Password Authenication Request")
                ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' Send SMS of Server Status
                ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let Password.MousePointer = ccHourglass
                    Let MessageOcx.Text = "Automatic Train Control - Software Registration" & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Software Version is " & Chr$(34) & TextboxVersion.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Software Date is " & Chr$(34) & TextboxDate.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Computer's Name is " & Chr$(34) & TextBoxComputerName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "User's Name is " & Chr$(34) & TextBoxUserName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "User's Email Address is " & Chr$(34) & TextBoxUsersEmailAddress.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Sponsor's Name is  " & Chr$(34) & TextBoxSponsorsName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Sponsor's ID is " & Chr$(34) & TextboxSponsorsID.Text & Chr$(34) & vbCrLf
                    
                    Let MessageOcx.Sender = "canadianlocomotivelogistics@gmail.com"
                    Let MessageOcx.Recipients = "canadianlocomotivelogistics@gmail.com"
                    Let MessageOcx.Subject = "Software Registration"
                    Let MessageOcx.ReplyTo = TextBoxUsersEmailAddress.Text
                    Let SmtpOcx.TimeoutMS = 12000
                    Let SmtpOcx.UserName = "CanadianLocomotiveLogisitcs@gmail.com"
                    Let SmtpOcx.Password = "walnuttree12"
                    Let SmtpOcx.LoginMethod = AuthLoginMethod
                    Let SmtpOcx.ServerAddr = "smtp.gmail.com"
                    Let SmtpOcx.ServerPort = 587
                    Let TemporaryValue = SmtpOcx.Send(MessageOcx.GetRaw)
                    Let Password.MousePointer = ccDefault
                    If TemporaryValue <> 0 Then
                        Let TemporaryInput = MsgBox("Automatic Train Control software tired to send an email for registration." & vbCrLf & _
                                                    "Sending the email was unsucessful. Do you wish to try again?", vbExclamation + vbYesNo, "Automatic Train Control - Unsucessful Email Attempt")
                        If TemporaryInput = vbYes Then
                            Let Password.MousePointer = ccHourglass
                            Let TemporaryValue = SmtpOcx.Send(MessageOcx.GetRaw)
                            Let Password.MousePointer = ccDefault
                            If TemporaryValue <> 0 Then
                                Let TemporaryInput = MsgBox("Automatic Train Control software tired to send an email for registration." & vbCrLf & _
                                                            "Sending the email was unsucessful.", vbExclamation + vbOKOnly, "Automatic Train Control - Unsucessful Email Attempt")
                            End If
                        End If
                    Else
                        Let TemporaryInput = MsgBox("Automatic Train Control software has sucessfully sent an email for registration.", vbExclamation + vbOKOnly, "Automatic Train Control - Sucessful Email Attempt")
                    End If
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Not Registred Grace Period
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let LabelStatus.Caption = "Status: Unregistred with Grace"
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Get Installation Date
                    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let Ini.Filename = App.Path$ & "\Atc.ini"
                    Let Ini.Application = "Opening Screen"
                    Let Ini.Parameter = "InstallationDate"
                    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Remmember Date$ vs Date
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    If DateDiff("d", Ini.Value, Date) < 0 Or DateDiff("d", Ini.Value, Date) > 30 Then
                        Let TemporaryResponse = MsgBox("You have exceeded the thirty days grace to register the software." & vbCrLf & _
                                                       "This version of software is still not registred. Please email" & vbCrLf & _
                                                       "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Software")
                    Else
                        Let TemporaryResponse = MsgBox("Until registration is complete, you have " & _
                                                       CStr(6 - DateDiff("d", Ini.Value, Date)) & " day(s) grace.", vbOKOnly, "Automatic Train Control - Unregistered Software")
                        Load MainScreen
                        MainScreen.Show vbModeless
                    End If
                End If
            Else ' If TemporaryType = "Lite" then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Lite Version with Incorrect Password
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Let TemporaryResponse = MsgBox("The password for Automatic Train Control is incorrect, we at Canadian" & vbCrLf & _
                                               "Locomotive Logistics assume that Automatic Train Control (Lite Version)" & vbCrLf & _
                                               "is not registered. Do you wish to register at this time?" & vbCrLf & vbCrLf & _
                                               "Registering the software with Canadian Locomotive Logistics will provide" & vbCrLf & _
                                               "you with a password.", vbExclamation + vbYesNo, "Password Authenication Request")
                If TemporaryResponse = vbNo Then
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' No Registration
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    'If DateDiff("d", format (SystemInformationOcx.FileDate, "mm-dd-yyyy", Date$) < 0 Or DateDiff("d", format(SystemInformationOcx.FileDate,"mm-dd-yyyy), Date$) > 60 Then
                    '    MsgBox "This version of software has expired. Please email" & vbCrLf & "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Version"
                    '    Let ButtonClose.Enabled = False
                    'Else
                    '    MsgBox "A trial period of" & Str(90 - (DateDiff("d", SystemInformationOcx.FileDate, Date))) & " days is granted. After the expiration you must register the software.", vbexclamation + vbOKOnly, "Automatic Train Control - Thirty day trail period"
                    '    Load MainScreen
                    '    MainScreen.Show vbModeless
                    'End If
                Else
                    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' With Registration
                    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let TemporaryResponse = MsgBox("Automatic Train Control (Lite Version) will now send an email for" & vbCrLf & _
                                                   "registration of the software. Please wait.", vbExclamation + vbOKOnly, "Password Authenication Request")
                    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Send SMS of Server Status
                    ' ------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let Password.MousePointer = ccHourglass
                    Let MessageOcx.Text = "Automatic Train Control - Software Registration" & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Software Version is " & Chr$(34) & TextboxVersion.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Software Date is " & Chr$(34) & TextboxDate.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Computer's Name is " & Chr$(34) & TextBoxComputerName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "User's Name is " & Chr$(34) & TextBoxUserName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "User's Email Address is " & Chr$(34) & TextBoxUsersEmailAddress.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Sponsor's Name is  " & Chr$(34) & TextBoxSponsorsName.Text & Chr$(34) & vbCrLf
                    Let MessageOcx.Text = MessageOcx.Text & "Sponsor's ID is " & Chr$(34) & TextboxSponsorsID.Text & Chr$(34) & vbCrLf
                    
                    Let MessageOcx.Sender = "canadianlocomotivelogistics@gmail.com"
                    Let MessageOcx.Recipients = "canadianlocomotivelogistics@gmail.com"
                    Let MessageOcx.Subject = "Software Registration"
                    Let MessageOcx.ReplyTo = TextBoxUsersEmailAddress.Text
                    Let SmtpOcx.TimeoutMS = 12000
                    Let SmtpOcx.UserName = "CanadianLocomotiveLogisitics@gmail.com"
                    Let SmtpOcx.Password = "walnuttree12"
                    Let SmtpOcx.LoginMethod = AuthLoginMethod
                    Let SmtpOcx.ServerAddr = "smtp.gmail.com"
                    Let SmtpOcx.ServerPort = 587
                    Let TemporaryValue = SmtpOcx.Send(MessageOcx.GetRaw)
                    Let Password.MousePointer = ccDefault
                    If TemporaryValue <> 0 Then
                        Let TemporaryInput = MsgBox("Automatic Train Control software tried to send an email for registration." & vbCrLf & _
                                                    "Sending the email was unsucessful. Do you wish to try again?", vbExclamation + vbYesNo, "Automatic Train Control - Unsucessful Email Attempt")
                        If TemporaryInput = vbYes Then
                            Let Password.MousePointer = ccHourglass
                            Let TemporaryValue = SmtpOcx.Send(MessageOcx.GetRaw)
                            Let Password.MousePointer = ccDefault
                            If TemporaryValue <> 0 Then
                                Let TemporaryInput = MsgBox("Automatic Train Control software tried to send an email for registration." & vbCrLf & _
                                                            "Sending the email was unsucessful.", vbExclamation + vbOKOnly, "Automatic Train Control - Unsucessful Email Attempt")
                            End If
                        End If
                    Else
                        Let TextboxPassword.Text = TemporaryAsciiTotal ^ 3
                    End If
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Not Registred Grace Period
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let LabelStatus.Caption = "Status: Unregistred with Grace"
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Get Installation Date
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Let Ini.Filename = App.Path$ & "\Atc.ini"
                    Let Ini.Application = "Opening Screen"
                    Let Ini.Parameter = "InstallationDate"
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    ' Remember Date$ vs Date
                    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
                    If DateDiff("d", Ini.Value, Date) < 0 Or DateDiff("d", Ini.Value, Date) > 30 Then
                        Let TemporaryResponse = MsgBox("You have exceeded the thirty days grace to register the software." & vbCrLf & _
                                                       "This version of software is still not registred. Please email" & vbCrLf & _
                                                       "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Software")
                    Else
                        Let TemporaryResponse = MsgBox("Until registration is complete, you have " & _
                                                       CStr(6 - DateDiff("d", Ini.Value, Date)) & " day(s) grace.", vbOKOnly, "Automatic Train Control - Unregistered Software")
                        Load MainScreen
                        MainScreen.Show vbModeless
                    End If
                End If
            End If
        Else 'If TemporaryAsciiTotal ^ 3 = Val(TextboxPassword.Text) Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Correct Password
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let LabelStatus.Caption = "Status: Accepted"
            If TemporaryType = "Full" Then
                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' Correct Password with Full Version
                ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                If DateDiff("d", SystemInformationOcx.FileDate, Date) < 0 Or DateDiff("d", SystemInformationOcx.FileDate, Date) > 365 Then
                    Let TemporaryResponse = MsgBox("This version of software has expired. Please email" & vbCrLf & _
                                                    "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Software")
                Else
                    Load MainScreen
                    MainScreen.Show vbModeless
                End If
            Else 'If TemporaryType = "Lite" Then
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' Correct Password with Lite Version
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
                If DateDiff("d", SystemInformationOcx.FileDate, Date) < 0 Or DateDiff("d", SystemInformationOcx.FileDate, Date) > 365 Then
                    Let TemporaryResponse = MsgBox("This version of software has expired. Please email" & vbCrLf & _
                                                    "canadianlocomotivelogistics@gmail.com for an updated version of this software.", vbOKOnly, "Automatic Train Control - Expired Software")
                Else
                    Load MainScreen
                    MainScreen.Show vbModeless
                End If
            End If
        End If
    Else ' Must be "Pro" version then
        Load MainScreen
        MainScreen.Show vbModeless
    End If
    
ExitOut:
    Password.Enabled = True
    DoEvents
    
End Sub

Private Sub ButtonPrint_Click()

    Password.PrintForm
    
End Sub

Private Sub Form_Activate()

    DoEvents
    
'   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Password Screen, Form, Activate" & vbcrlf
'    End If

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
                Let Ini.Value = "Password Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Passowrd Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

 Call BalloonHelpUpdatePart01
 Call BalloonHelpUpdatePart02
 
 
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
     
'    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Password Screen, Form, Activate" & vbcrlf
'    End If
   
    End Sub

Private Sub Form_Deactivate()
    
'   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Password Screen, Form, Deactivate" & vbcrlf
'    End If

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
    Let Ini.Parameter = "UsersEmailAddress"
    Let Ini.Value = TextBoxUsersEmailAddress.Text
    Let Ini.Parameter = "SponsorsName"
    Let Ini.Value = TextBoxSponsorsName.Text
    Let Ini.Parameter = "SponsorsID"
    Let Ini.Value = TextboxSponsorsID.Text
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
                Let Ini.Value = "Password Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
    
'   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Password Screen, Form, Deactivate" & vbcrlf
'    End If

End Sub


Private Sub Form_Load()
    
'   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Beginning,  Password Screen, Form, Load" & vbcrlf
'    End If

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
    Let Ini.Application = "Password Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "UsersEmailAddress"
    Let TextBoxUsersEmailAddress.Text = Ini.Value
    Let Ini.Parameter = "SponsorsName"
    Let TextBoxSponsorsName.Text = Ini.Value
    Let Ini.Parameter = "SponsorsID"
    Let TextboxSponsorsID.Text = Ini.Value
    Let Ini.Parameter = "Password"
    Let TextboxPassword.Text = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "LogFile"
        If Ini.Value = "On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "About Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting."
        End If
    End If
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No databases to declare

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fill Textboxes with parameters
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let SystemInformationOcx.Drive = "c:"
    Let SystemInformationOcx.Filename = App.Path$ & "\Atc.exe"
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Type"
    Dim TemporaryType As String
    Let TemporaryType = Ini.Value
    If TemporaryType = "Full" Then
        Let TextBoxSponsorsName.Text = "Canadian Locomotive Logistics"
        Let TextBoxSponsorsName.Locked = True
        Let TextBoxSponsorsName.BackColor = &H8000000F
        'Let TextBoxSponsorsName.ForeColor = &H80000008
        Let TextBoxSponsorsName.TabStop = False
        Let TextboxSponsorsID.Text = "CLL010"
        Let TextboxSponsorsID.Locked = True
        Let TextboxSponsorsID.BackColor = &H8000000F
        'Let TextboxSponsorsID.ForeColor = &H80000008
        Let TextboxSponsorsID.TabStop = False
    End If
    

    Let TextboxVersion.Text = SystemInformationOcx.FileVersion & " (" & TemporaryType & " Version)"
    'Format the date and time variables to be consistant. Windows changes the way the dat and time are displayed
    Let TextboxDate.Text = Format(SystemInformationOcx.FileDate, "mm-dd-yyyy") & " " & Format(SystemInformationOcx.FileTime, "hh:mm:ss")
    Let TextBoxComputerName.Text = SystemInformationOcx.ComputerName
    Let TextBoxUserName.Text = SystemInformationOcx.UserName

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'   If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then 'Darrin
'        Let DebugMode!TextBoxDebugMode.Text = DebugMode!TextBoxDebugMode.Text & "Ending, Password Screen, Form, Load" & vbcrlf
'    End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub



Private Sub Form_Resize()

    If Password.WindowState = vbMinimized Then
    
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
        
    ElseIf Password.WindowState = vbNormal Then
    
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

Private Sub SmtpOcx_SendProgress(ByVal CurrentBytes As Long, ByVal TotalBytes As Long, Cancel As Long)

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






Private Sub BalloonHelpUpdatePart01()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "BalloonHelp"
    Let TemporaryBalloonHelp = Ini.Value
   
    If TemporaryBalloonHelp = True Then
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
        Let Ini.Parameter = "BalloonHelpOpacity"
        Let BalloonHelpOpacity = Ini.Value
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "Transparency"
        If Ini.Value = "Off" Then
            BalloonHelpOpacity = 255
        ElseIf Ini.Value = "On" Then
            'Do Nothing
        Else
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, invalid value for 'Transparency' in ATC.INI file."
            End If
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "BalloonHelp"
        Dim TemporarySpeechHelp As Boolean
        Let TemporarySpeechHelp = Ini.Value
        
        If TemporarySpeechHelp = False Then
            Let balloonhelp.Speech = False
        ElseIf TemporarySpeechHelp = True Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let Password.MousePointer = ccHourglass
                
        Let BalloonHelpText1 = "This textbox displays your current version of" & vbCrLf & "software for Automatic Train Control. It is" & vbCrLf & "needed to make a password."
        Let BalloonHelpText2 = "Software Version"
        'Let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxVersion)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxVersion, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, BalloonHelpWaveFile)
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxVersion' control."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays the time/date stamp for the " & vbCrLf & "software, Automatic Train Control. It is" & vbCrLf & "needed to make a password."
        Let BalloonHelpText2 = "Date and Time of File"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxDate)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxDate, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxDate' control."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays your current computer name" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let BalloonHelpText2 = "Your Computer's Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxComputerName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxComputerName, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxComputerName' control."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays your current user name" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let BalloonHelpText2 = "Your User Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxUserName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxUserName, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxUserName' control."
            End If
        End If
        
        Let BalloonHelpText1 = "This textbox displays your email address" & vbCrLf & " for registration. It is needed" & vbCrLf & "to make a password."
        Let BalloonHelpText2 = "Your Email Address"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxUserEmailAddress)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxUsersEmailAddress, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxUsersEmailAddress' control."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays your sponsor's name" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let BalloonHelpText2 = "Your Sponsor's Name"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxComputerName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxSponsorsName, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxComputerName' control."
            End If
        End If

        Let BalloonHelpText1 = "This textbox displays your sponsor's ID number" & vbCrLf & " for Automatic Train Control. It is needed" & vbCrLf & "to make a password."
        Let BalloonHelpText2 = "Your Sponsor's ID"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxUserName)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxSponsorsID, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxUserName' control."
            End If
        End If
         
        Let BalloonHelpText1 = "This textbox displays your for Automatic Train" & vbCrLf & "Control. Type in the here and 'Close' the" & vbCrLf & "window or 'Abort' to exit the program."
        Let BalloonHelpText2 = "Password"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextboxPassword)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextboxPassword, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'TextboxPassword' control."
            End If
        End If
    
        Let BalloonHelpText1 = "This button aborts the program. If you do not have" & vbCrLf & "a passowrd, please 'click' on my email address for" & vbCrLf & "a password."
        Let BalloonHelpText2 = "Abort Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(buttonabort)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonAbort, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'buttonAbort' control."
            End If
        End If
        
        Let Password.MousePointer = ccDefault
        
    ElseIf TemporaryBalloonHelp = False Then
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'ButtonClose' control."
            End If
        End If
    Else
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "LogFile"
        If Ini.Value = "On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Password Screen, Load Form, variable error in ATC.INI file for 'Balloon Help' setting."
        End If
    
    End If
    
End Sub

Public Sub BalloonHelpUpdatePart02()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "BalloonHelp"
    Let TemporaryBalloonHelp = Ini.Value
   
    If TemporaryBalloonHelp = True Then
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
        Let Ini.Parameter = "BalloonHelpOpacity"
        Let BalloonHelpOpacity = Ini.Value
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "All Screens"
        Let Ini.Parameter = "Transparency"
        If Ini.Value = "Off" Then
            BalloonHelpOpacity = 255
        ElseIf Ini.Value = "On" Then
            'Do Nothing
        Else
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, invalid value for 'Transparency' in ATC.INI file."
            End If
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Speech On if
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "BalloonHelp"
        Dim TemporarySpeechHelp As Boolean
        Let TemporarySpeechHelp = Ini.Value
        
        If TemporarySpeechHelp = False Then
            Let balloonhelp.Speech = False
        ElseIf TemporarySpeechHelp = True Then
            Let balloonhelp.Speech = True
            Let balloonhelp.Voice = 0
            Let BalloonHelpWaveFile = ""
        End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Each Element
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Let Password.MousePointer = ccHourglass
                
        Let BalloonHelpText1 = "This button prints the current window to your printer."
        Let BalloonHelpText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'ButtonPrint' control."
            End If
        End If
     
        Let BalloonHelpText1 = "This button close the Authenication" & vbCrLf & "window and returns you to the main screen, if you have" & vbCrLf & "a correct password."
        Let BalloonHelpText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonHelpText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'ButtonClose' control."
            End If
        End If
        
        Let Password.MousePointer = ccDefault
        
    ElseIf TemporaryBalloonHelp = False Then
        Let BalloonHelpSetup = balloonhelp.DestroyAllToolTips
        If BalloonHelpSetup = 0 Then
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Password Authenication, Balloon Help Update Part 01, unable to update balloon help for 'ButtonClose' control."
            End If
        End If
    Else
        Let Ini.Filename = App.Path$ & "\Atc.ini"
        Let Ini.Application = "Main Screen"
        Let Ini.Parameter = "LogFile"
        If Ini.Value = "On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Password Screen, Load Form, variable error in ATC.INI file for 'Balloon Help' setting."
        End If
    
    End If
    
End Sub
