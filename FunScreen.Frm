VERSION 4.00
Begin VB.Form FunScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - FunScreen"
   ClientHeight    =   8085
   ClientLeft      =   2985
   ClientTop       =   3165
   ClientWidth     =   7755
   Height          =   8490
   Icon            =   "FunScreen.frx":0000
   Left            =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Top             =   2820
   Width           =   7875
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5400
      Top             =   1560
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   17
      Left            =   6720
      TabIndex        =   35
      Text            =   "-2000"
      Top             =   6240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   17
      Left            =   240
      Picture         =   "FunScreen.frx":0442
      ScaleHeight     =   360
      ScaleWidth      =   10860
      TabIndex        =   34
      Top             =   7200
      Width           =   10860
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   16
      Left            =   6720
      TabIndex        =   33
      Text            =   "-2000"
      Top             =   5880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   3255
      Picture         =   "FunScreen.frx":16AB
      ScaleHeight     =   315
      ScaleWidth      =   1440
      TabIndex        =   32
      Top             =   1350
      Width           =   1440
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   15
      Left            =   6720
      TabIndex        =   31
      Text            =   "-2000"
      Top             =   5520
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   14
      Left            =   6720
      TabIndex        =   30
      Text            =   "-2000"
      Top             =   5160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   13
      Left            =   6720
      TabIndex        =   29
      Text            =   "-2000"
      Top             =   4800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   12
      Left            =   6720
      TabIndex        =   28
      Text            =   "-2000"
      Top             =   4440
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   11
      Left            =   6720
      TabIndex        =   27
      Text            =   "-2000"
      Top             =   4080
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   16
      Left            =   135
      Picture         =   "FunScreen.frx":191F
      ScaleHeight     =   315
      ScaleWidth      =   1500
      TabIndex        =   26
      Top             =   6540
      Width           =   1500
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   135
      Picture         =   "FunScreen.frx":1ACE
      ScaleHeight     =   300
      ScaleWidth      =   1320
      TabIndex        =   25
      Top             =   6120
      Width           =   1320
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   135
      Picture         =   "FunScreen.frx":1C0C
      ScaleHeight     =   300
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   5685
      Width           =   1110
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   135
      Picture         =   "FunScreen.frx":20F8
      ScaleHeight     =   300
      ScaleWidth      =   1080
      TabIndex        =   23
      Top             =   5280
      Width           =   1080
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   150
      Picture         =   "FunScreen.frx":2291
      ScaleHeight     =   300
      ScaleWidth      =   1080
      TabIndex        =   22
      Top             =   4875
      Width           =   1080
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   120
      Picture         =   "FunScreen.frx":2476
      ScaleHeight     =   300
      ScaleWidth      =   1125
      TabIndex        =   21
      Top             =   4500
      Width           =   1125
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   135
      Picture         =   "FunScreen.frx":263B
      ScaleHeight     =   300
      ScaleWidth      =   1125
      TabIndex        =   20
      Top             =   4140
      Width           =   1125
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   120
      Picture         =   "FunScreen.frx":2872
      ScaleHeight     =   315
      ScaleWidth      =   1125
      TabIndex        =   19
      Top             =   3720
      Width           =   1125
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   90
      Picture         =   "FunScreen.frx":2ABF
      ScaleHeight     =   315
      ScaleWidth      =   1125
      TabIndex        =   18
      Top             =   3285
      Width           =   1125
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   10
      Left            =   6720
      TabIndex        =   17
      Text            =   "-2000"
      Top             =   3720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   9
      Left            =   6720
      TabIndex        =   16
      Text            =   "-2000"
      Top             =   3360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   8
      Left            =   6720
      TabIndex        =   15
      Text            =   "-2000"
      Top             =   3000
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   7
      Left            =   6720
      TabIndex        =   14
      Text            =   "-2000"
      Top             =   2640
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   6
      Left            =   6720
      TabIndex        =   13
      Text            =   "-2000"
      Top             =   2265
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   5
      Left            =   6720
      TabIndex        =   12
      Text            =   "-2000"
      Top             =   1905
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   4
      Left            =   6720
      TabIndex        =   11
      Text            =   "-2000"
      Top             =   1545
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   3
      Left            =   6720
      TabIndex        =   10
      Text            =   "-2000"
      Top             =   1185
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   2
      Left            =   6720
      TabIndex        =   9
      Text            =   "-2000"
      Top             =   810
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   315
      Index           =   1
      Left            =   6720
      TabIndex        =   8
      Text            =   "-3500"
      Top             =   480
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   120
      Picture         =   "FunScreen.frx":2C8C
      ScaleHeight     =   300
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   2880
      Width           =   1140
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   3
      Left            =   45
      Picture         =   "FunScreen.frx":2E3B
      ScaleHeight     =   390
      ScaleWidth      =   1950
      TabIndex        =   6
      Top             =   1770
      Width           =   1950
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   4
      Left            =   2070
      Picture         =   "FunScreen.frx":3403
      ScaleHeight     =   390
      ScaleWidth      =   1950
      TabIndex        =   5
      Top             =   2040
      Width           =   1950
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   6
      Left            =   1770
      Picture         =   "FunScreen.frx":39CC
      ScaleHeight     =   390
      ScaleWidth      =   1950
      TabIndex        =   4
      Top             =   2520
      Width           =   1950
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   165
      Picture         =   "FunScreen.frx":3C52
      ScaleHeight     =   315
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   450
      Width           =   1500
   End
   Begin VB.TextBox TrianObjectLocation 
      Height          =   300
      Index           =   0
      Left            =   6720
      TabIndex        =   2
      Text            =   "-2000"
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   105
      Picture         =   "FunScreen.frx":3EDA
      ScaleHeight     =   315
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   60
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   1560
   End
   Begin VB.PictureBox TrainObject 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   75
      Picture         =   "FunScreen.frx":417B
      ScaleHeight     =   315
      ScaleWidth      =   1125
      TabIndex        =   0
      Top             =   2250
      Width           =   1125
   End
   Begin IniconLib.Init Ini 
      Left            =   4920
      Top             =   2160
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
      Left            =   6000
      Top             =   1560
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin VB.Line TrackTies 
      BorderStyle     =   3  'Dot
      X1              =   2400
      X2              =   5760
      Y1              =   3720
      Y2              =   5520
   End
   Begin VB.Line TrackLine 
      X1              =   2520
      X2              =   5880
      Y1              =   3480
      Y2              =   5160
   End
End
Attribute VB_Name = "FunScreen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()

End Sub


Private Sub ButtonClose_Click()

Let Timer1.Interval = 0
Let Timer1.Enabled = False
FunScreen.Hide
Unload FunScreen

End Sub

Private Sub Form_Activate()

' =============================================================================================================================================================================
' Add to Screen Stack
' =============================================================================================================================================================================
    Let Ini.filename = App.Path$ & "\Atc.ini"
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
        If TemporaryScreen = "Fun Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Fun Screen"
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
        If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This" & Chr$(13) & "error will be recored in the ATC.LOG file. Please" & Chr$(13) & "email the author reporting the error and attach a copy" & Chr$(13) & "of the file called ATC.LOG for detailed information. This" & Chr$(13) & "program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Fun Screen, Form Activate, stack is full, overflow."
        End If
    End If
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transparency Screen Delay
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.filename = App.Path$ & "\Atc.ini"
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
            If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
                Stop
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
                Stop
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
    Let Ini.filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Fun Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(FunScreen.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(FunScreen.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(FunScreen.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(FunScreen.Height)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Transparency Screen Delay
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.filename = App.Path$ & "\Atc.ini"
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
            If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
                Stop
            End If
        End If
    ElseIf TemporaryBackgroundImage = "Off" Then
        AlphaBlend.Enabled = False
    Else
        If MainScreen.MenuLogFile.Caption = "&Log File is On" Then
                Stop
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    FunScreen.Hide
    'unload funscreen

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
End Sub


Private Sub Form_Load()

Let FunScreen.Width = MainScreen.Width
Let FunScreen.Left = MainScreen.Left
Let FunScreen.Top = Val(MainScreen.Top) + Val(MainScreen.Height)
Let FunScreen.Height = 1000

Let trackties.X1 = 5
Let trackties.Y1 = 510
Let trackties.X2 = Val(FunScreen.Width) - 5
Let trackties.Y2 = 510


Let trackline.X1 = 5
Let trackline.Y1 = 501
Let trackline.X2 = Val(FunScreen.Width) - 5
Let trackline.Y2 = 501

Let temp = 0

For t = 0 To 16
Let TrianObjectLocation(t) = 0 - TrainObject(t).Width - temp
Let temp = temp + Val(TrainObject(t).Width)
Let TrainObject(t).Top = 500 - Val(TrainObject(t).Height)
Next t

Let TrianObjectLocation(17) = Val(FunScreen.Width) + 5
Let TrainObject(17).Top = 500 - Val(TrainObject(17).Height)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'   Every time a new window is opened in Automatic Train Control we check the screen size and compare it to the window screen size.
' If the window cannot be displayed in the current screen size a message box is displayed. This allows time for the user to change
' the screen attributes to correct size.

    Do While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    Loop

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Initialization of Screen
'
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Let Ini.filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Fun Screen"
    Let Ini.Parameter = "Top"
    Dim temporaryvalueTop As String
    Let temporaryvalueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
    If Val(TemporaryValueLeft) = 0 And Val(temporaryvalueTop) = 0 Then
        FunScreen.Left = (Screen.Width - Width) / 2   ' Center form horizontally.
        FunScreen.Top = (Screen.Height - Height) / 2  ' Center form vertiCally.
    Else
        If Val(TemporaryValueLeft) + FunScreen.Width > Screen.Width Then
            Let FunScreen.Left = Screen.Width - FunScreen.Width
        Else
            Let FunScreen.Left = Val(TemporaryValueLeft)
        End If
        If Val(temporaryvalueTop) + FunScreen.Height > Screen.Height Then
            Let FunScreen.Top = Screen.Height - FunScreen.Height
        Else
            Let FunScreen.Top = Val(temporaryvalueTop)
        End If
    End If
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Check Status of Transparency
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If MainScreen.MenuTransparency.Caption = "&Transparency is On" Then
        Let AlphaBlend.Enabled = True
    ElseIf MainScreen.MenuTransparency.Caption = "&Transparency is Off" Then
        Let AlphaBlend.Enabled = False
    End If
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Adding Balloons
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
        'Dim TemporaryText1 As String
        'Dim TemporaryText2 As String
        'Dim i As Integer
        'Dim t As Boolean
        'Dim f As Boolean
        'Let t = True
        'Let f = False

        'Let TemporaryText1 = "This text box is where all information from your" + vbCrLf + "serial port is displayed. Commands given by the" + vbCrLf + "program are displayed here. You can also type your" + vbCrLf + "own commands, providing the port is not busy."
        'Let TemporaryText2 = "Communication Window"
        'i = BalloonHelp.DestroyToolTip(TextBoxCommunicationWindowDCC.hWnd)
        'i = BalloonHelp.AddToolTip(TextBoxCommunicationWindowDCC.hWnd, TemporaryText1, IIf(t, balBalloon, balStandard), TemporaryText2, IIf(t, balInfo, IIf(f, balWarning, balError)), &HC0FFFF, &H0)

    'End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'No database to declare
    
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Timer1_Timer()

For t = 0 To 16
    Let TrianObjectLocation(t).Text = Val(TrianObjectLocation(t).Text) + 5
Next t

If TrianObjectLocation(16).Text = Val(FunScreen.Width) + 500 Then
    For t = 0 To 16
        Let TrianObjectLocation(t) = 0 - TrainObject(t).Width - temp
        Let temp = temp + Val(TrainObject(t).Width)
    Next t
End If

End Sub


Private Sub Timer2_Timer()

Let TrianObjectLocation(17).Text = Val(TrianObjectLocation(17).Text) - 5

If Val(TrianObjectLocation(17).Text) = 0 - Val(TrainObject(17).Width) - 100 Then
    Let TrianObjectLocation(17).Text = Val(FunScreen.Width) + 200
End If

End Sub

Private Sub TrianObjectLocation_Change(Index As Integer)

Let TrainObject(Index).Left = Val(TrianObjectLocation(Index).Text)

End Sub


