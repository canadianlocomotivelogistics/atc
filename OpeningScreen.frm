VERSION 4.00
Begin VB.Form OpeningScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Contol - Opening Screen"
   ClientHeight    =   4395
   ClientLeft      =   1110
   ClientTop       =   8400
   ClientWidth     =   10545
   FillStyle       =   0  'Solid
   Height          =   4800
   Icon            =   "OpeningScreen.frx":0000
   Left            =   1050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   10545
   Top             =   8055
   Width           =   10665
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "OpeningScreen.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox PictureLogo 
      Height          =   1725
      Left            =   3540
      ScaleHeight     =   107.205
      ScaleMode       =   0  'User
      ScaleWidth      =   453
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   6855
   End
   Begin VB.Timer OpeningScreenTimer 
      Interval        =   6000
      Left            =   960
      Top             =   1380
   End
   Begin etHyperLabel.HyperLabel HyperLabel1 
      Height          =   195
      Left            =   3480
      Top             =   3720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   344
      ProjectKey      =   "et0B49E"
      ForeColor       =   8388608
      Target          =   "canadianlocomotivelogistics@gmail.com"
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
      Caption         =   "canadianlocomotivelogistics@gmail.com"
      Autosize        =   0   'False
      HoverUnderline  =   0   'False
      TargetType      =   "1"
   End
   Begin VB.Label Label6 
      Caption         =   ". And if"
      Height          =   195
      Left            =   8640
      TabIndex        =   14
      Top             =   3240
      Width           =   735
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   3000
      Top             =   180
      _ExtentX        =   767
      _ExtentY        =   661
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx3 
      Height          =   195
      Left            =   5520
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   344
      ProjectKey      =   "et0B49E"
      ForeColor       =   8388608
      Target          =   "https://groups.io/g/AutomaticTrainControl"
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
      Caption         =   "https://groups.io/g/AutomaticTrainControl"
      Autosize        =   0   'False
      HoverUnderline  =   0   'False
   End
   Begin etHyperLabel.HyperLabel HyperLabelOcx1 
      Height          =   195
      Left            =   5700
      Top             =   3000
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   344
      ProjectKey      =   "et0B49E"
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
      Autosize        =   0   'False
      HoverUnderline  =   0   'False
   End
   Begin SystemInfoControl.MSysInfo SystemInformationOCX 
      Left            =   3000
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin IniconLib.Init Ini 
      Left            =   2460
      Top             =   1320
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
      Left            =   1980
      Top             =   1320
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Line Line2 
      X1              =   10440
      X2              =   3540
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label8 
      Caption         =   "you still need additional help and user input please email me at ."
      Height          =   195
      Left            =   3540
      TabIndex        =   13
      Top             =   3480
      Width           =   4470
   End
   Begin VB.Label Label5 
      Caption         =   "to join the Yahoo groups at"
      Height          =   195
      Left            =   3540
      TabIndex        =   12
      Top             =   3240
      Width           =   1905
   End
   Begin VB.Label Label4 
      Caption         =   "or email me at and do not forget"
      Height          =   195
      Left            =   7920
      TabIndex        =   11
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Please visit my home page at"
      Height          =   255
      Left            =   3540
      TabIndex        =   10
      Top             =   3000
      Width           =   2115
   End
   Begin WaveLib.Wave Wave1 
      Left            =   1440
      Top             =   1380
      _Version        =   65537
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Exclusive       =   0   'False
      Filename        =   ""
      FileLength      =   -1
      Loop            =   0   'False
      PlayEnd         =   -1
      PlayStart       =   -1
   End
   Begin VB.Label Label2 
      Caption         =   "If you can assist financially, I would appreciate it . Some of the OCX files cost me to register them."
      Height          =   255
      Left            =   3540
      TabIndex        =   7
      Top             =   2040
      Width           =   6915
   End
   Begin VB.Label Label1 
      Caption         =   $"OpeningScreen.frx":0884
      Height          =   435
      Left            =   3540
      TabIndex        =   6
      Top             =   2340
      Width           =   6915
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label OpeningScreenLastModified 
      Caption         =   "Last Modified,"
      Height          =   495
      Left            =   660
      TabIndex        =   5
      Top             =   780
      Width           =   2790
   End
   Begin VB.Label OpeningScreenCompanyName 
      Caption         =   $"OpeningScreen.frx":0923
      Height          =   1575
      Left            =   660
      TabIndex        =   4
      Top             =   2340
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label OpeningScreenCopywrite 
      AutoSize        =   -1  'True
      Caption         =   "Copywrite 2001-2021 by"
      Height          =   195
      Left            =   660
      TabIndex        =   3
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Label OpeningScreenAuthor 
      AutoSize        =   -1  'True
      Caption         =   "Written by Darrin J. Calcutt"
      Height          =   195
      Left            =   3540
      TabIndex        =   2
      Top             =   4080
      Width           =   1890
   End
   Begin VB.Label OpeningScreenVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
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
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   645
   End
   Begin VB.Label OpeningScreenTitle 
      AutoSize        =   -1  'True
      Caption         =   "Automatic Train Control"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "OpeningScreen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub Form_Activate()

DoEvents

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
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Opening Screen, Form Activate, variable error in ATC.INI file for 'Transparency' setting. Varible is currently set to '" & TemporaryTransparency & "'."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
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
            Let Ini.Value = "Opening Screen, Form Activate, variable error in ATC.INI file for 'Background' setting. Variable is currently set to '" & TemporaryBackgroundImage & "'."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Attributes
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Opening Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(OpeningScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(OpeningScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(OpeningScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(OpeningScreen.Height)

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
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "LogFile"
            If Ini.Value = "On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Opening Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
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
            Let Ini.Value = "Opening Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

    OpeningScreen.Hide

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
    Let Ini.Application = "Opening Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "Counter"
    Dim TemporaryValueCounter As String
    Let TemporaryValueCounter = Ini.Value
    Let Ini.Value = Str$(Val(TemporaryValueCounter) + 1)
    Let Ini.Parameter = "InstallationDate"
    Dim TemporaryValueDate As String
    Let TemporaryValueDate = Ini.Value
    If TemporaryValueDate = "" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Date$ Vs Date
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Let Ini.Value = Date
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        OpeningScreen.Left = (Screen.Width - Width) / 2
        OpeningScreen.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + OpeningScreen.Width > Screen.Width Then
            Let OpeningScreen.Left = Screen.Width - OpeningScreen.Width
        Else
            Let OpeningScreen.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + OpeningScreen.Height > Screen.Height Then
            Let OpeningScreen.Top = Screen.Height - OpeningScreen.Height
        Else
            Let OpeningScreen.Top = Val(TemporaryValueTop)
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
            Let Ini.Value = "Opening Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If
       
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Opening Screen Sound File
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Wave1.Filename = App.Path$ & "\Sounds\Graphics\Ge_p3.wav"
    Let Wave1.Action = wAPlay

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Logo
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let PictureLogo.Picture = LoadPicture(App.Path$ & "\Graphics\Logo.gif")

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update System Information
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let SystemInformationOcx.Filename = App.Path$ & "\Atc.exe"
    Let SystemInformationOcx.Drive = "C:"
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim TemporaryType As String
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Main Screen"
    Let Ini.Parameter = "Type"
    Let TemporaryType = Ini.Value
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let OpeningScreenVersion.Caption = "Version " & SystemInformationOcx.FileVersion & " (" & TemporaryType & " Version)"
    Let OpeningScreenLastModified = "LastModified, " & Chr$(13) & Format(SystemInformationOcx.FileDate, "mm-dd-yyyy") & " at " & Format(SystemInformationOcx.FileTime, "hh:mm:ss")
    
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


Private Sub HyperLabel1_RightClick()

End Sub

Private Sub HyperLabelOcx1_RightClick()

End Sub


Private Sub HyperLabelOcx3_Click()

End Sub

Private Sub HyperLabelOcx3_RightClick()

End Sub


Private Sub Label1_Click()

End Sub

Private Sub OpeningScreenAuthor_Click()

End Sub


Private Sub OpeningScreenCompanyName_Click()

End Sub

Private Sub OpeningScreenCopywrite_Click()

End Sub

Private Sub OpeningScreenTimer_Timer()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Turn Timer Object Off
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let OpeningScreenTimer.Interval = 0

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Display Password Form
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load MainScreen
    MainScreen.Show vbModeless
    'Load Password
    'Password.Show vbModeless

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Statement
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub


