VERSION 5.00
Object = "{625E24A3-B09D-101D-85F5-6EBA1EE93AF4}#3.3#0"; "INICON32.OCX"
Object = "{5025865D-07B8-11D4-AC0B-00E07D76E465}#1.0#0"; "TransPicture.ocx"
Begin VB.Form UpdaterBackGround 
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   1275
   ClientTop       =   1830
   ClientWidth     =   7920
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "UpdaterBackGround.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
Attribute VB_Name = "UpdaterBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
        UpdaterBackGround.ZOrder 1
        UpdaterBackGround.WindowState = 2
        Let UpdaterBackGround.Visible = True
    ElseIf TemporaryValue = "Off" Then
        Let UpdaterBackGround.Visible = False
        UpdaterBackGround.ZOrder 1
        UpdaterBackGround.WindowState = 2
    ElseIf MainScreen!MenuLogFile.Caption = "&Log File is On" Then
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
    
    Load Updater
    Updater.Show vbModeless

End Sub


