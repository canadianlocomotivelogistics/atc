VERSION 4.00
Begin VB.Form BackGround 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   1275
   ClientTop       =   1830
   ClientWidth     =   7920
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Height          =   4410
   Icon            =   "BackGround.frx":0000
   Left            =   1215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   Top             =   1485
   Visible         =   0   'False
   Width           =   8040
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7200
      Top             =   480
   End
   Begin IniconLib.Init Ini 
      Left            =   7200
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
   Begin TransPicture.TransPictureCtl TransparentPicture 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      TransparentColor=   16777215
   End
   Begin VB.Image ImageBoxBackGround 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "BackGround"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Load()

    ImageBoxBackGround.Picture = LoadPicture(App.Path$ + "\Graphics\BackGroundImage.bmp")
    TransparentPicture.Picture = LoadPicture(App.Path$ + "\Graphics\BackgroundTitle.bmp")

End Sub


Private Sub Timer1_Timer()

    Let Timer1.Interval = 0
    
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "All Screens"
    Let Ini.Parameter = "BackgroundImage"
    Dim TemporaryValue As String
    Let TemporaryValue = Ini.Value
    
    Let ImageBoxBackGround.Width = Screen.Width / 15
    Let ImageBoxBackGround.Height = Screen.Height / 15
    
    If TemporaryValue = "On" Then
        BackGround.ZOrder 1
        BackGround.WindowState = 2
        Let BackGround.Visible = True
    ElseIf TemporaryValue = "Off" Then
        Let BackGround.Visible = False
        BackGround.ZOrder 1
        BackGround.WindowState = 2
    Else 'If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
        Let Ini.Filename = App.Path$ & "\Atc.log"
        Let Ini.Application = "Log Errors"
        Let Ini.Parameter = Date$ & " " & Time$
        Let Ini.Value = "BackGround Screen, Timer1 Timer, variable not set correctly for 'BackGround Image' in ATC.INI file."
    End If
    
    For tt = 1 To 5000
        DoEvents
    Next tt
    
    Load OpeningScreen
    OpeningScreen.Show vbModeless

End Sub


Private Sub TransparentPicture_Click()

End Sub


