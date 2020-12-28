VERSION 4.00
Begin VB.Form DefineBlocks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Defining Layout"
   ClientHeight    =   9075
   ClientLeft      =   3030
   ClientTop       =   3465
   ClientWidth     =   7170
   Height          =   9480
   Icon            =   "DefineBlocks.frx":0000
   Left            =   2970
   LinkTopic       =   "Form1"
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   Top             =   3120
   Width           =   7290
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   4560
      TabIndex        =   78
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton ButtonList 
      Caption         =   "&List"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10080
      Top             =   840
   End
   Begin VB.TextBox TextBoxMapLabelCounter 
      Height          =   285
      Left            =   9720
      TabIndex        =   75
      Text            =   "0"
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox TextboxMapLabelPositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   74
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox TextBoxMapLabelPositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   73
      Top             =   6960
      Width           =   855
   End
   Begin VB.HScrollBar ScrollHorizontal 
      Height          =   255
      LargeChange     =   36
      Left            =   120
      TabIndex        =   65
      Top             =   5640
      Width           =   6615
   End
   Begin VB.VScrollBar ScrollVertical 
      Height          =   4215
      LargeChange     =   36
      Left            =   6840
      TabIndex        =   64
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox PictureBoxTrack 
      Height          =   4335
      Left            =   120
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   62
      Top             =   1200
      Width           =   6615
      Begin VB.PictureBox PictureBoxTrackMap 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7500
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   500
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   63
         Top             =   120
         Width           =   15360
      End
   End
   Begin VB.TextBox TextBoxSignal333PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   61
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal222PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   60
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal111PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   59
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal33PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   58
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal22PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   57
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal11PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   56
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal3PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   55
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal2PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   54
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal1PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   53
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal333Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   52
      Text            =   "0"
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal222Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   51
      Text            =   "0"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal111Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   50
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal33Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   49
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal22Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   48
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal11Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   47
      Text            =   "0"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal3Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   46
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal1Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   45
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal333PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   44
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal222PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   43
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal111PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   42
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal33PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   41
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal22PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   40
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal11PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   39
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal3PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   38
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal1PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   37
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignal2Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   36
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignal2PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   35
      Top             =   3720
      Width           =   855
   End
   Begin VB.Data DataBaseTrackPlan 
      Connect         =   "Access"
      DatabaseName    =   ""
      EOFAction       =   1  'EOF
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TrackPlan"
      Top             =   120
      Width           =   1980
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   20
      Text            =   "0"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   19
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   18
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1Counter 
      Height          =   285
      Left            =   9720
      TabIndex        =   17
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconStraightCounter 
      Height          =   285
      Left            =   9720
      TabIndex        =   16
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1PositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1PositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconStraightPositionTop 
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconStraightPositionLeft 
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   8760
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "DefineBlocks.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin TabDlg.SSTab TabTrackIcon 
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2990
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Track Icons"
      TabPicture(0)   =   "DefineBlocks.frx":0884
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "PictureBoxTrackIconSwitch4(0)"
      Tab(0).Control(2)=   "PictureBoxTrackIconSwitch3(0)"
      Tab(0).Control(3)=   "PictureBoxTrackIconSwitch2(0)"
      Tab(0).Control(4)=   "PictureBoxTrackIconSwitch1(0)"
      Tab(0).Control(5)=   "PictureIconTrack"
      Tab(0).Control(6)=   "PictureBoxTrackIconStraight(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Signal Icons"
      TabPicture(1)   =   "DefineBlocks.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelDescription5"
      Tab(1).Control(1)=   "PictureBoxSignal1(0)"
      Tab(1).Control(2)=   "PictureBoxSignal2(0)"
      Tab(1).Control(3)=   "PictureBoxSignal3(0)"
      Tab(1).Control(4)=   "PictureBoxSignal22(0)"
      Tab(1).Control(5)=   "PictureBoxSignal11(0)"
      Tab(1).Control(6)=   "PictureBoxSignal33(0)"
      Tab(1).Control(7)=   "PictureBoxSignal111(0)"
      Tab(1).Control(8)=   "PictureBoxSignal222(0)"
      Tab(1).Control(9)=   "PictureBoxSignal333(0)"
      Tab(1).Control(10)=   "Picture2"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Other Icons"
      TabPicture(2)   =   "DefineBlocks.frx":08BC
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "LabelDecription4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TextBoxMapLabel(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Picture1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.PictureBox PictureBoxTrackIconStraight 
         AutoSize        =   -1  'True
         Height          =   330
         Index           =   0
         Left            =   -71640
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   76
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -74760
         Picture         =   "DefineBlocks.frx":08D8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   71
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox PictureIconTrack 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -74760
         Picture         =   "DefineBlocks.frx":0D1A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   69
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "DefineBlocks.frx":115C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   68
         Top             =   600
         Width           =   480
      End
      Begin VB.TextBox TextBoxMapLabel 
         Height          =   615
         Index           =   0
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   66
         Text            =   "DefineBlocks.frx":159E
         Top             =   600
         Width           =   2175
      End
      Begin VB.PictureBox PictureBoxSignal333 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -68400
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   34
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal222 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -68640
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   33
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal111 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -68880
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   32
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal33 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -69240
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   31
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal11 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -69720
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   30
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal22 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -69480
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   29
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal3 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70080
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   28
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal2 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70320
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   27
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignal1 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -70560
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   26
         Top             =   600
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch1 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   -70920
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   25
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch2 
         AutoSize        =   -1  'True
         Height          =   600
         Index           =   0
         Left            =   -70200
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   24
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch3 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   -68760
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   23
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch4 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   -69480
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   22
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LabelDescription5 
         Caption         =   "Again, drag and drop any of these track signals to the track map (picture box) above to replicate your personal layout."
         Height          =   975
         Left            =   -74160
         TabIndex        =   72
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Drag and drop any one of these track icon onto the track map (picture box) above to rplicate your personal layout."
         Height          =   855
         Left            =   -74160
         TabIndex        =   70
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label LabelDecription4 
         Caption         =   "Type in the information you would like to see in a lable. Then drag and drop it on the track map (picture box) above."
         Height          =   855
         Left            =   960
         TabIndex        =   67
         Top             =   600
         Width           =   3375
      End
   End
   Begin Balloon_OCX.BalloonOCX BalloonHelp 
      Left            =   8280
      Top             =   720
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   9480
      Top             =   840
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
      Left            =   8880
      Top             =   720
      _ExtentX        =   767
      _ExtentY        =   767
      Opacity         =   0
   End
   Begin VB.Label LabelDescription3 
      Caption         =   "When you are finished drawing out your track diagram, 'update' the database before 'close'ing the window."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   6975
   End
   Begin VB.Label LabelStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label LabelDescription2 
      Caption         =   $"DefineBlocks.frx":15BD
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   6975
   End
   Begin VB.Label LabelDescription 
      Caption         =   $"DefineBlocks.frx":1654
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu menuTrackIcon 
      Caption         =   "TrackIcon"
      Visible         =   0   'False
      Begin VB.Menu menuTrackIconProperties 
         Caption         =   "&Track Icon Properties"
      End
      Begin VB.Menu menuLine 
         Caption         =   ""
      End
      Begin VB.Menu menuTrackIconDelete 
         Caption         =   "&Delete Track Icon"
      End
      Begin VB.Menu menuTrackIconDeleteAll 
         Caption         =   "Delete &All Track Icons"
      End
   End
End
Attribute VB_Name = "DefineBlocks"
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
            If TemporaryScreen = "Define Blocks Screen" Then
                Let Ini.Value = "Unused"
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks Screen, Button Close, current window is not listed in the stack to remove it and hide."
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
                Let Ini.Value = "Define Blocks Screen, Button Close, trying to display the previous window using the screen stack, window not recognized."
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
            Let Ini.Value = "Define Blocks Screen, Button Close, stack is empty, underflow."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Subroutine
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub






Private Sub ButtonPrint_Click()

    DefineBlocks.PrintForm
    
End Sub

Private Sub DataBaseTrackPlan_Reposition()
    
    'Let DataBaseTrackPlan.Caption = DataBaseTrackPlan.Recordset.AbsolutePosition

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
        If TemporaryScreen = "Define Blocks Screen" Then
            Let TemporaryCounter = 11
        ElseIf TemporaryScreen = "Unused" Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add to INI if not Present
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let Ini.Value = "Define Blocks Screen"
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
            Let Ini.Value = "Define Blocks Screen, Form Activate, stack is full, overflow."
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
                Let Ini.Value = "Define Blocks Screen, Form Activate, variable error in ATC.INI file for 'Trnsparency' setting."
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
            Let Ini.Value = "Define Blocks Screen, Form Activate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Save PictureBox Locations
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Let TabTrackIcon.Tab = 0
    DoEvents
    Let TextBoxTrackIconStraightPositionLeft.Text = PictureBoxTrackIconStraight(0).Left
    Let TextBoxTrackIconStraightPositionTop.Text = PictureBoxTrackIconStraight(0).Top
    Let TextBoxTrackIconSwitch1PositionLeft.Text = PictureBoxTrackIconSwitch1(0).Left
    Let TextBoxTrackIconSwitch1PositionTop.Text = PictureBoxTrackIconSwitch1(0).Top
    Let TextBoxTrackIconSwitch2PositionLeft.Text = PictureBoxTrackIconSwitch2(0).Left
    Let TextBoxTrackIconSwitch2PositionTop.Text = PictureBoxTrackIconSwitch2(0).Top
    Let TextBoxTrackIconSwitch3PositionLeft.Text = PictureBoxTrackIconSwitch3(0).Left
    Let TextBoxTrackIconSwitch3PositionTop.Text = PictureBoxTrackIconSwitch3(0).Top
    Let TextBoxTrackIconSwitch4PositionLeft.Text = PictureBoxTrackIconSwitch4(0).Left
    Let TextBoxTrackIconSwitch4PositionTop.Text = PictureBoxTrackIconSwitch4(0).Top
    
    Let TabTrackIcon.Tab = 1
    DoEvents
    Let TextBoxSignal1PositionTop.Text = PictureBoxSignal1(0).Top
    Let TextBoxSignal1PositionLeft.Text = PictureBoxSignal1(0).Left
    Let TextBoxSignal11PositionTop.Text = PictureBoxSignal11(0).Top
    Let TextBoxSignal11PositionLeft.Text = PictureBoxSignal11(0).Left
    Let TextBoxSignal111PositionTop.Text = PictureBoxSignal111(0).Top
    Let TextBoxSignal111PositionLeft.Text = PictureBoxSignal111(0).Left
    Let TextBoxSignal2PositionTop.Text = PictureBoxSignal2(0).Top
    Let TextBoxSignal2PositionLeft.Text = PictureBoxSignal2(0).Left
    Let TextBoxSignal22PositionTop.Text = PictureBoxSignal22(0).Top
    Let TextBoxSignal22PositionLeft.Text = PictureBoxSignal22(0).Left
    Let TextBoxSignal222PositionTop.Text = PictureBoxSignal222(0).Top
    Let TextBoxSignal222PositionLeft.Text = PictureBoxSignal222(0).Left
    Let TextBoxSignal3PositionTop.Text = PictureBoxSignal3(0).Top
    Let TextBoxSignal3PositionLeft.Text = PictureBoxSignal3(0).Left
    Let TextBoxSignal33PositionTop.Text = PictureBoxSignal33(0).Top
    Let TextBoxSignal33PositionLeft.Text = PictureBoxSignal33(0).Left
    Let TextBoxSignal333PositionTop.Text = PictureBoxSignal333(0).Top
    Let TextBoxSignal333PositionLeft.Text = PictureBoxSignal333(0).Left
    
    
    Let TabTrackIcon.Tab = 2
    DoEvents
    Let TextboxMapLabelPositionTop.Text = TextBoxMapLabel(0).Top
    Let TextBoxMapLabelPositionLeft.Text = TextBoxMapLabel(0).Left

    Let TabTrackIcon.Tab = 0

End Sub

Private Sub Form_Deactivate()

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Saving Variables
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Let Ini.Filename = App.Path$ & "\Atc.ini"
    Let Ini.Application = "Define Blocks Screen"
    Let Ini.Parameter = "Top"
    Let Ini.Value = Str$(DefineBlocks.Top)
    Let Ini.Parameter = "Left"
    Let Ini.Value = Str$(DefineBlocks.Left)
    Let Ini.Parameter = "Width"
    Let Ini.Value = Str(DefineBlocks.Width)
    Let Ini.Parameter = "Height"
    Let Ini.Value = Str(DefineBlocks.Height)

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
                Let Ini.Value = "Define Blocks Screen, Form Deactivate, variable error in ATC.INI file for 'Transparency' setting."
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
            Let Ini.Value = "Define Blocks Screen, Form Deactivate, variable error in ATC.INI file for 'Background' setting."
        End If
    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hide Screen
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    DefineBlocks.Hide
    'unload defineblocks

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
    Let Ini.Application = "Define Blocks Screen"
    Let Ini.Parameter = "Top"
    Dim TemporaryValueTop As String
    Let TemporaryValueTop = Ini.Value
    Let Ini.Parameter = "Left"
    Dim TemporaryValueLeft As String
    Let TemporaryValueLeft = Ini.Value
    Let Ini.Parameter = "WindowState"
    Dim TemporaryValueWindowState As String
    Let TemporaryValueWindowState = Ini.Value
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Positioning the Screen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Val(TemporaryValueLeft) = 0 And Val(TemporaryValueTop) = 0 Then
        DefineBlocks.Left = (Screen.Width - Width) / 2
        DefineBlocks.Top = (Screen.Height - Height) / 2
    Else
        If Val(TemporaryValueLeft) + DefineBlockst.Width > Screen.Width Then
            Let DefineBlocks.Left = Screen.Width - DefineBlocks.Width
        Else
            Let DefineBlocks.Left = Val(TemporaryValueLeft)
        End If
        If Val(TemporaryValueTop) + DefineBlocks.Height > Screen.Height Then
            Let DefineBlocks.Top = Screen.Height - DefineBlocks.Height
        Else
            Let DefineBlocks.Top = Val(TemporaryValueTop)
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

        Let TemporaryText1 = "This picture box is where your track plan is" & vbCrLf & "displayed. You can drag other track objects onto" & vbCrLf & "this map."
        Let TemporaryText2 = "Track Map Picture"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackMap)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackMap, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This vertical scroll moves your track map up" & vbCrLf & "or down."
        Let TemporaryText2 = "Veritcle Scroll"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ScrollVertical)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ScrollVertical, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This horizontal scroll moves your track map left" & vbCrLf & "or right."
        Let TemporaryText2 = "Horizontal Scroll"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ScrollHorizontal)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ScrollHorizontal, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This tab allow for different selections of track" & vbCrLf & "pictures to drag onto the track plan."
        Let TemporaryText2 = "Tab Track Icon Groups"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TabTrackIcon)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TabTrackIcon, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Straight No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the type of straight" & vbCrLf & "track. Right 'click'ing gives you more options."
        Let TemporaryText2 = "Straight Track"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconStraight(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconStraight(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Switch No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the type of switch" & vbCrLf & "track. Right 'click'ing gives you more options."
        Let TemporaryText2 = "Switch Type One"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch1(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch1(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "Switch No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the type of switch" & vbCrLf & "track. Right 'click'ing gives you more options."
        Let TemporaryText2 = "Switch Type Two"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch2(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch2(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Switch No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the type of switch" & vbCrLf & "track. Right 'click'ing gives you more options."
        Let TemporaryText2 = "Switch Type Three"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch3(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch3(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Switch No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the type of switch" & vbCrLf & "track. Right 'click'ing gives you more options."
        Let TemporaryText2 = "Switch Type Four"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch4(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch4(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (single target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal1(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal1(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (single over single target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal11(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal11(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (single over single over single target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal111(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal111(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (double target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal2(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal2(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (double over double target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal22(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal22(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
         
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (double over double over double target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal222(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal222(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
     
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (triple target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal3(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal3(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (triple over triple target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal33(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal33(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "Signal No. 0." & vbCrLf & "This picture box can be dragged to the track" & vbCrLf & "map. Double 'click'ing changes the direction of signal." & vbCrLf & "Right 'click'ing gives you more options."
        Let TemporaryText2 = "Signal (triple over triple over triple target)"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal333(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal333(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
        Let TemporaryText1 = "TextBox No. 0." & vbCrLf & "This text box can be dragged to the track" & vbCrLf & "map. Try editing the text in the textbox first," & vbCrLf & "to label areas on your track plan."
        Let TemporaryText2 = "Label"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxMapLabel(0))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxMapLabel(0), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

        Let TemporaryText1 = "This button when 'click'ed on will" & vbCrLf & "print the current screen."
        Let TemporaryText2 = "Print Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonPrint)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonPrint, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        
        Let TemporaryText1 = "This button closes th Define Track window" & vbCrLf & "and returns you to the main screen."
        Let TemporaryText2 = "Close Button"
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(ButtonClose)
        Let BalloonHelpSetup = balloonhelp.AddToolTip(ButtonClose, BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

    End If

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Defining Databases and files
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Let DataBaseTrackPlan.DatabaseName = App.Path$ & "\Databases\TrackPlanDatabase.mdb"
PictureBoxTrackMap.Picture = LoadPicture(App.Path$ & "\Graphics\TrackPlan.bmp")
PictureBoxTrackIconStraight(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight1.bmp")
PictureBoxTrackIconSwitch1(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType1Normal.bmp")
PictureBoxTrackIconSwitch2(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType2Normal.bmp")
PictureBoxTrackIconSwitch3(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType3Normal.bmp")
PictureBoxTrackIconSwitch4(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType4Normal.bmp")
PictureBoxSignal1(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal1.bmp")
PictureBoxSignal11(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal11.bmp")
PictureBoxSignal111(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal111.bmp")
PictureBoxSignal2(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal2.bmp")
PictureBoxSignal22(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal22.bmp")
PictureBoxSignal222(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal222.bmp")
PictureBoxSignal3(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal3.bmp")
PictureBoxSignal33(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal33.bmp")
PictureBoxSignal333(0).Picture = LoadPicture(App.Path$ & "\Graphics\signal333.bmp")

' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub Form_Resize()

    Let DataBaseTrackPlan.Left = DefineBlocks.ScaleWidth + 100
    'Let Timer1.Left = DefineBlocks.ScaleWidth + 100
    'Let Timer2.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconStraightPositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch1PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch2PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch3PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch4PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconStraightPositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch1PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch2PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch3PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch4PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconStraightCounter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch1Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch2Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch3Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxTrackIconSwitch4Counter.Left = DefineBlocks.ScaleWidth + 100
      
    Let TextBoxSignal1PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal11PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal111PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal2PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal22PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal222PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal3PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal33PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal333PositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal1PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal11PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal111PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal2PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal22PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal222PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal3PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal33PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal333PositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal1Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal11Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal111Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal2Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal22Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal222Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal3Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal33Counter.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxSignal333Counter.Left = DefineBlocks.ScaleWidth + 100
   
    Let PictureBoxTrack.Height = -88 + DefineBlocks.ScaleHeight - 261
    Let PictureBoxTrack.Width = DefineBlocks.ScaleWidth - 9 - 9 - ScrollVertical.Width - 9
    
    Let ScrollVertical.Left = DefineBlocks.ScaleWidth - 9 - ScrollVertical.Width
    Let ScrollVertical.Height = PictureBoxTrack.Height

    Let ScrollHorizontal.Top = PictureBoxTrack.Top + PictureBoxTrack.Height + 9
    Let ScrollHorizontal.Width = PictureBoxTrack.Width


    Let LabelDescription2.Top = DefineBlocks.ScaleHeight - 222
 
    Let TabTrackIcon.Top = DefineBlocks.ScaleHeight - 182
    Let TabTrackIcon.Left = Int(DefineBlocks.ScaleWidth / 2) - Int(TabTrackIcon.Width / 2)

    Let LabelDescription3.Top = DefineBlocks.ScaleHeight - 62

    Let ButtonClose.Left = DefineBlocks.ScaleWidth - ButtonClose.Width - 9
    Let ButtonClose.Top = DefineBlocks.ScaleHeight - ButtonClose.Height - 9
    Let ButtonPrint.Left = DefineBlocks.ScaleWidth - ButtonPrint.Width - 9
    Let ButtonPrint.Top = DefineBlocks.ScaleWidth - ButtonPrint.Width - 9
    Let ButtonList.Left = DefineBlocks.ScaleWidth - ButtonList.Width - 100
    Let ButtonList.Top = DefineBlocks.ScaleHeight - ButtonList.Height - 9

    Let TextBoxMapLabelPositionLeft.Left = DefineBlocks.ScaleWidth + 100
    Let TextboxMapLabelPositionTop.Left = DefineBlocks.ScaleWidth + 100
    Let TextBoxMapLabelCounter.Left = DefineBlocks.ScaleWidth + 100

    Call PictureBoxTrackMap_Change
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    If DefineBlocks.WindowState = vbMinimized Then
    
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
        
    ElseIf DefineBlocks.WindowState = vbNormal Then
    
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





Private Sub menuTrackIconProperties_Click()
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Who has the Focus?
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    DefineBlockProperties.DatabaseBlockProperties.Recordset.FindFirst "RecordCounter =" & DefineBlocks.ActiveControl.Tag
    If DataBaseTrackPlan.Recordset.NoMatch Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Menu Track Icon Properties, Click, no article found in the Database for Block Properties; tag property no matching any record."
        End If
    End If
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Load DefineBlockProperties
    DefineBlockProperties.Show vbModeless
  
End Sub

Private Sub menuTrackIconDeleteAll_Click()
    
    Dim TemporaryIndex As Integer
    On Error Resume Next
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Straight
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxTrackIconStraightCounter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxTrackIconStraight(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxTrackIconStraight(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxTrackIconStraightCounter = 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Switch1
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxTrackIconSwitch1Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxTrackIconSwitch1(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxTrackIconSwitch1(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxTrackIconSwitch1Counter = 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Switch2
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxTrackIconSwitch2Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxTrackIconSwitch2(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxTrackIconSwitch2(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxTrackIconSwitch2Counter = 0
        
   
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Switch3
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxTrackIconSwitch3Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxTrackIconSwitch3(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxTrackIconSwitch3(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxTrackIconSwitch3Counter = 0
    
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Switch4
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxTrackIconSwitch4Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxTrackIconSwitch4(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxTrackIconSwitch4(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxTrackIconSwitch4Counter = 0
  
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal1
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal1Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal1(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal1(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal1Counter = 0
     
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal11
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal11Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal11(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal11(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal11Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal111
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal111Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal111(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal111(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal111Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal2
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal2Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal2(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal2(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal2Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal22
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal22Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal22(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal22(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal22Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal222
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal222Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal222(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal222(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal222Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal3
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal3Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal3(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal3(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal3Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal33
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal33Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal33(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal33(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal33Counter = 0
     
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Signal333
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxSignal333Counter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & PictureBoxSignal333(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload PictureBoxSignal333(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxSignal333Counter = 0
 
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Track Icon Map Label
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    For TemporaryIndex = 1 To Val(TextBoxMapLabelCounter.Text)
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " & TextBoxMapLabel(TemporaryIndex).Tag
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Where in the Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.NoMatch Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Log If Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
            End If
        Else
            Unload TextBoxMapLabel(TemporaryIndex)
            DataBaseTrackPlan.Recordset.Delete
        End If
    Next TemporaryIndex
    
    Let TextBoxMapLabelCounter = 0
    
    On Error GoTo 0
    
End Sub

Private Sub menuTrackIconDelete_Click()
        
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Who has the Focus?
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter =" & DefineBlocks.ActiveControl.Tag
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Where in the Database
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If DataBaseTrackPlan.Recordset.NoMatch Then
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Log If Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Menu Track Icon Delete All, Click, no article found in the Database for Block Properties; tag property no matching any record."
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Remove the Record for Database
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.Delete
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Locate Icon On the Map
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DefineBlocks.ActiveControl.Name = "PictureBoxTrackIconStraight" Then
            Unload PictureBoxTrackIconStraight(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxTrackIconSwitch1" Then
            Unload PictureBoxTrackIconSwitch1(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxTrackIconSwitch2" Then
            Unload PictureBoxTrackIconSwitch2(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxTrackIconSwitch3" Then
            Unload PictureBoxTrackIconSwitch3(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxTrackIconSwitch4" Then
            Unload PictureBoxTrackIconSwitch4(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal1" Then
            Unload PictureBoxSignal1(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal11" Then
            Unload PictureBoxSignal11(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal111" Then
            Unload PictureBoxSignal111(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal2" Then
            Unload PictureBoxSignal2(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal22" Then
            Unload PictureBoxSignal22(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal222" Then
            Unload PictureBoxSignal222(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal3" Then
            Unload PictureBoxSignal3(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal33" Then
            Unload PictureBoxSignal33(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "PictureBoxSignal333" Then
            Unload PictureBoxSignal333(DefineBlocks.ActiveControl.Index)
        ElseIf DefineBlocks.ActiveControl.Name = "TextBoxMapLabel" Then
            Unload TextBoxMapLabel(DefineBlocks.ActiveControl.Index)
        Else
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Menu Track Icon Delete, Click, active control not recognized by program, cannot delete."
            End If
        End If
    End If
End Sub






Private Sub PictureBoxSignal1_DblClick(Index As Integer)
  
If PictureBoxSignal1(Index).Container.Name = "PictureBoxTrackMap" Then
  
    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal1(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal1West.bmp" Then
        PictureBoxSignal1(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal1East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal1East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single Target)" & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal1East.bmp" Then
        PictureBoxSignal1(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal1West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal1West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single Target)" & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Picture Box Signal1, Double Click, signal in the database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal1(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal1(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxSignal1_GotFocus(Index As Integer)

    Let PictureBoxSignal1(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (single target) No." + Str$(Index) + " of " + Str$(TextBoxSignal1Counter.Text)

End Sub

Private Sub PictureBoxSignal1_LostFocus(Index As Integer)

    If PictureBoxSignal1(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal1(Index).BorderStyle = 0
    ElseIf PictureBoxSignal1(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal1(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal1(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal1(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If


End Sub

Private Sub PictureBoxSignal11_DblClick(Index As Integer)

If PictureBoxSignal11(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal11(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal11West.bmp" Then
        PictureBoxSignal11(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal11East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal11East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single over single target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal11East.bmp" Then
        PictureBoxSignal11(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal11West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal11West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single over single target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal11, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal11(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal11(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxSignal11_GotFocus(Index As Integer)

    Let PictureBoxSignal11(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (single over single target) No." + Str$(Index) + " of " + Str$(TextBoxSignal2Counter.Text)

End Sub


Private Sub PictureBoxSignal11_LostFocus(Index As Integer)

    If PictureBoxSignal11(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal11(Index).BorderStyle = 0
    ElseIf PictureBoxSignal11(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal11(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        PictureBoxSignal11(Index).Drag vbBeginDrag
    ElseIf Button = vbRightButton Then
        If PictureBoxSignal11(Index).Container.Name = "PictureBoxTrackMap" Then
            DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub PictureBoxSignal111_DblClick(Index As Integer)
    
If PictureBoxSignal111(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal111(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal111West.bmp" Then
        PictureBoxSignal111(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal111East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal111East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single over single over single target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal111East.bmp" Then
        PictureBoxSignal111(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal111West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal111West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Single over single over single target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal111, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal111(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal111(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal111_GotFocus(Index As Integer)

    Let PictureBoxSignal111(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (single over single over single target) No." + Str$(Index) + " of " + Str$(TextBoxSignal3Counter.Text)

End Sub

Private Sub PictureBoxSignal111_LostFocus(Index As Integer)

    If PictureBoxSignal111(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal111(Index).BorderStyle = 0
    ElseIf PictureBoxSignal111(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal111(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal111_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal111(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal111(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal2_DblClick(Index As Integer)

If PictureBoxSignal2(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal2(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal2West.bmp" Then
        PictureBoxSignal2(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal2East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal2East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal2East.bmp" Then
        PictureBoxSignal2(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal2West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal2West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal2, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal2(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal2(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal2_GotFocus(Index As Integer)

    Let PictureBoxSignal2(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (double target) No." + Str$(Index) + " of " + Str$(TextBoxSignal11Counter.Text)

End Sub

Private Sub PictureBoxSignal2_LostFocus(Index As Integer)

    If PictureBoxSignal2(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal2(Index).BorderStyle = 0
    ElseIf PictureBoxSignal2(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal2(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"
    

End Sub

Private Sub PictureBoxSignal2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal2(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal2(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal22_DblClick(Index As Integer)

If PictureBoxSignal22(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal22(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal22West.bmp" Then
        PictureBoxSignal22(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal22East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal22East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double over double target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal22East.bmp" Then
        PictureBoxSignal22(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal22West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal22West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double over double target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal22, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal22(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal22(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal22_GotFocus(Index As Integer)

    Let PictureBoxSignal22(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (double over double traget) No." + Str$(Index) + " of " + Str$(TextBoxSignal22Counter.Text)

End Sub

Private Sub PictureBoxSignal22_LostFocus(Index As Integer)

    If PictureBoxSignal22(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal22(Index).BorderStyle = 0
    ElseIf PictureBoxSignal22(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal22(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal22_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal22(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal22(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal222_DblClick(Index As Integer)

If PictureBoxSignal222(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal222(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal222West.bmp" Then
        PictureBoxSignal222(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal222East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal222East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double over double over double target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal222East.bmp" Then
        PictureBoxSignal222(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal222West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal222West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Double over double over double target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal222, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal222(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal222(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal222_GotFocus(Index As Integer)

    Let PictureBoxSignal222(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (double over double over double target) No." + Str$(Index) + " of " + Str$(TextBoxSignal33Counter.Text)

End Sub

Private Sub PictureBoxSignal222_LostFocus(Index As Integer)

    If PictureBoxSignal222(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal222(Index).BorderStyle = 0
    ElseIf PictureBoxSignal222(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal222(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal222_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal222(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal222(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal3_DblClick(Index As Integer)

If PictureBoxSignal3(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal3(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal3West.bmp" Then
        PictureBoxSignal3(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal3East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal3East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal3East.bmp" Then
        PictureBoxSignal3(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal3West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal3West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal3, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal3(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal3(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal3_GotFocus(Index As Integer)

    Let PictureBoxSignal3(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (triple target) No." + Str$(Index) + " of " + Str$(TextBoxSignal111Counter.Text)

End Sub

Private Sub PictureBoxSignal3_LostFocus(Index As Integer)

    If PictureBoxSignal3(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal3(Index).BorderStyle = 0
    ElseIf PictureBoxSignal3(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal3(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal3(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal3(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal33_DblClick(Index As Integer)

If PictureBoxSignal33(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal33(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal33West.bmp" Then
        PictureBoxSignal33(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal33East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal33East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple over triple target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal33East.bmp" Then
        PictureBoxSignal33(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal33West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal33West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple over triple target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal33, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal33(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal33(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal33_GotFocus(Index As Integer)

    Let PictureBoxSignal33(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (triple over triple target) No." + Str$(Index) + " of " + Str$(TextBoxSignal222Counter.Text)

End Sub

Private Sub PictureBoxSignal33_LostFocus(Index As Integer)

    If PictureBoxSignal33(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal33(Index).BorderStyle = 0
    ElseIf PictureBoxSignal33(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal33(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal33_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal33(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal33(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxSignal333_DblClick(Index As Integer)

If PictureBoxSignal333(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal333(Index).Tag
    
    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal333West.bmp" Then
        PictureBoxSignal333(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal333East.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal333East.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple over triple over triple target)." & vbCrLf & "Signal is for EAST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal333East.bmp" Then
        PictureBoxSignal333(Index).Picture = LoadPicture(App.Path$ & "\Graphics\Signal333West.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "Signal333West.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Signal No." + Str$(Index) + ". (Triple over triple over triple target)." & vbCrLf & "Signal is for WEST bound trains only."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Signal333, Double Click, signal in database is not defined as east or west."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal333(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal333(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")

End If

End Sub


Private Sub PictureBoxSignal333_GotFocus(Index As Integer)

    Let PictureBoxSignal333(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Signal (triple over triple over triple target) No." + Str$(Index) + " of " + Str$(TextBoxSignal333Counter.Text)

End Sub

Private Sub PictureBoxSignal333_LostFocus(Index As Integer)

    If PictureBoxSignal333(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxSignal333(Index).BorderStyle = 0
    ElseIf PictureBoxSignal333(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxSignal333(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxSignal333_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    PictureBoxSignal333(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxSignal333(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxTrackIconStraight_DblClick(Index As Integer)
    
    If PictureBoxTrackIconStraight(Index).Container.Name = "PictureBoxTrackMap" Then
        
        Dim TemporaryText1 As String
        Dim TemporaryText2 As String
        Dim i As Long
        Dim t As Boolean
        Dim f As Boolean
        Let t = True
        Let f = False
    
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconStraight(Index).Tag
        
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight1.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight2.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight2.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block ends to the east."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight2.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight3.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight3.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block ends to the west."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight3.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight4.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight4.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block contains a block occupancy detector."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight4.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight5.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight5.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block contains a derail."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight5.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight6.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight6.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(Str$(DataBaseTrackPlan.Recordset.Fields("Length"))) + " feet."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight6.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight7.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight7.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(Str$(DataBaseTrackPlan.Recordset.Fields("Length"))) + " feet."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight7.bmp" Then
            PictureBoxTrackIconStraight(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight1.bmp")
            DataBaseTrackPlan.Recordset.Edit
            DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackStraight1.bmp"
            DataBaseTrackPlan.Recordset.Update
            If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
                Let TemporaryText1 = "Straight Block No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(Str$(DataBaseTrackPlan.Recordset.Fields("Length"))) + " feet."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            End If
        Else
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, Picture Box Track Icon Straight, Double Click, track type in database is not defined."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconStraight(Index))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconStraight(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    
    End If

End Sub


Private Sub PictureBoxTrackIconStraight_GotFocus(Index As Integer)

Let PictureBoxTrackIconStraight(Index).BorderStyle = 1

Let LabelStatus.Caption = "Status: Straight Block No." + Str$(Index) + " of " + Str$(TextBoxTrackIconStraightCounter.Text)

End Sub


Private Sub PictureBoxTrackIconStraight_LostFocus(Index As Integer)

If PictureBoxTrackIconStraight(Index).Container.Name = "PictureBoxTrackMap" Then
    PictureBoxTrackIconStraight(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconStraight(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconStraight(Index).BorderStyle = 1
End If

Let LabelStatus.Caption = "Status: "

End Sub

Private Sub PictureBoxTrackIconStraight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        PictureBoxTrackIconStraight(Index).Drag vbBeginDrag
    ElseIf Button = vbRightButton Then
        If PictureBoxTrackIconStraight(Index).Container.Name = "PictureBoxTrackMap" Then
            DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
        End If
    End If
    
End Sub

Private Sub PictureBoxTrackIconSwitch1_DblClick(Index As Integer)
     
If PictureBoxTrackIconSwitch1(Index).Container.Name = "PictureBoxTrackMap" Then
    
    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch1(Index).Tag

    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType1Normal.bmp" Then
        PictureBoxTrackIconSwitch1(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType1Reverse.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType1Reverse.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType1Reverse.bmp" Then
        PictureBoxTrackIconSwitch1(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType1Normal.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackSwitchType1Normal.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Track Icon Switch1, Double Click, switch in database is not defined as normal or reversed."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch1(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch1(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxTrackIconSwitch1_GotFocus(Index As Integer)

    PictureBoxTrackIconSwitch1(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Switch Type1 No." + Str$(Index) + " of " + Str$(TextBoxTrackIconSwitch1Counter.Text)

End Sub


Private Sub PictureBoxTrackIconSwitch1_LostFocus(Index As Integer)

    If PictureBoxTrackIconSwitch1(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxTrackIconSwitch1(Index).BorderStyle = 0
    ElseIf PictureBoxTrackIconSwitch1(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxTrackIconSwitch1(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub


Private Sub PictureBoxTrackIconSwitch1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = vbLeftButton Then
    PictureBoxTrackIconSwitch1(Index).Drag vbBeginDrag
ElseIf Button = vbRightButton Then
    If PictureBoxTrackIconSwitch1(Index).Container.Name = "PictureBoxTrackMap" Then
        DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
    End If
End If

End Sub

Private Sub PictureBoxTrackIconSwitch2_DblClick(Index As Integer)
    
If PictureBoxTrackIconSwitch2(Index).Container.Name = "PictureBoxTrackMap" Then
    
    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch2(Index).Tag

    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType2Normal.bmp" Then
        PictureBoxTrackIconSwitch2(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType2Reverse.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType2Reverse.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Two No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType2Reverse.bmp" Then
        PictureBoxTrackIconSwitch2(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType2Normal.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackSwitchType2Normal.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Two No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Track Icon Switch2, Double Click, switch in database is not defined as normal or reversed."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch2(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch2(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxTrackIconSwitch2_GotFocus(Index As Integer)

    Let PictureBoxTrackIconSwitch2(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Switch Type2 No." + Str$(Index) + " of " + Str$(TextBoxTrackIconSwitch2Counter.Text)

End Sub


Private Sub PictureBoxTrackIconSwitch2_LostFocus(Index As Integer)

    If PictureBoxTrackIconSwitch2(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxTrackIconSwitch2(Index).BorderStyle = 0
    ElseIf PictureBoxTrackIconSwitch2(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxTrackIconSwitch2(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub


Private Sub PictureBoxTrackIconSwitch2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        PictureBoxTrackIconSwitch2(Index).Drag vbBeginDrag
    ElseIf Button = vbRightButton Then
        If PictureBoxTrackIconSwitch2(Index).Container.Name = "PictureBoxTrackMap" Then
            DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
        End If
    End If

End Sub


Private Sub PictureBoxTrackIconSwitch3_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch3(Index).Container.Name = "PictureBoxTrackMap" Then
    
    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch3(Index).Tag

    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType3Normal.bmp" Then
        PictureBoxTrackIconSwitch3(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType3Reverse.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType3Reverse.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Three No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType3Reverse.bmp" Then
        PictureBoxTrackIconSwitch3(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType3Normal.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackSwitchType3Normal.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Three No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Track Icon Switch3, Double Click, switch in database is not defined as normal or reversed."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch3(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch3(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxTrackIconSwitch3_GotFocus(Index As Integer)

    Let PictureBoxTrackIconSwitch3(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Switch Type3 No." + Str$(Index) + " of " + Str$(TextBoxTrackIconSwitch3Counter.Text)

End Sub


Private Sub PictureBoxTrackIconSwitch3_LostFocus(Index As Integer)

    If PictureBoxTrackIconSwitch3(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxTrackIconSwitch3(Index).BorderStyle = 0
    ElseIf PictureBoxTrackIconSwitch3(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxTrackIconSwitch3(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub


Private Sub PictureBoxTrackIconSwitch3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        PictureBoxTrackIconSwitch3(Index).Drag vbBeginDrag
    ElseIf Button = vbRightButton Then
        If PictureBoxTrackIconSwitch3(Index).Container.Name = "PictureBoxTrackMap" Then
            DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
        End If
    End If

End Sub


Private Sub PictureBoxTrackIconSwitch4_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch4(Index).Container.Name = "PictureBoxTrackMap" Then

    Dim TemporaryText1 As String
    Dim TemporaryText2 As String
    Dim i As Long
    Dim t As Boolean
    Dim f As Boolean
    Let t = True
    Let f = False

    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch4(Index).Tag

    If DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType4Normal.bmp" Then
        PictureBoxTrackIconSwitch4(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType4Reverse.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType4Reverse.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Four No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType4Reverse.bmp" Then
        PictureBoxTrackIconSwitch4(Index).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType4Normal.bmp")
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("pictureBoxFilename") = "TrackSwitchType4Normal.bmp"
        DataBaseTrackPlan.Recordset.Update
        If MainScreen.menuBalloonHelp.Caption = "&Balloon Help is On" Then
            Let TemporaryText1 = "Switch Type Four No." + Str$(Index) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Error
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "Define Blocks, Picture Box Track Icon Switch1, Double Click, switch in database is not defined as normal or reversed."
        End If
    End If
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch4(Index))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch4(Index), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
 
End If

End Sub


Private Sub PictureBoxTrackIconSwitch4_GotFocus(Index As Integer)

    Let PictureBoxTrackIconSwitch4(Index).BorderStyle = 1
    
    Let LabelStatus.Caption = "Status: Switch Type4 No." + Str$(Index) + " of " + Str$(TextBoxTrackIconSwitch4Counter.Text)

End Sub


Private Sub PictureBoxTrackIconSwitch4_LostFocus(Index As Integer)

    If PictureBoxTrackIconSwitch4(Index).Container.Name = "PictureBoxTrackMap" Then
        PictureBoxTrackIconSwitch4(Index).BorderStyle = 0
    ElseIf PictureBoxTrackIconSwitch4(Index).Container.Name = "TabTrackIcon" Then
        PictureBoxTrackIconSwitch4(Index).BorderStyle = 1
    End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub PictureBoxTrackIconSwitch4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        PictureBoxTrackIconSwitch4(Index).Drag vbBeginDrag
    ElseIf Button = vbRightButton Then
        If PictureBoxTrackIconSwitch4(Index).Container.Name = "PictureBoxTrackMap" Then
            DefineBlocks.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign
        End If
    End If

End Sub



Private Sub PictureBoxTrackMap_Change()
    
    
    Let ScrollHorizontal.Max = 0 - Int(Val(PictureBoxTrackMap.ScaleWidth) / 2) + Int(Val(PictureBoxTrack.ScaleWidth) / 2)
    Let ScrollHorizontal.Min = 0 + Int(Val(PictureBoxTrackMap.ScaleWidth) / 2) - Int(Val(PictureBoxTrack.ScaleWidth) / 2)
    
    Let ScrollVertical.Max = 0 - Int(Val(PictureBoxTrackMap.ScaleHeight) / 2) + Int(Val(PictureBoxTrack.ScaleHeight) / 2)
    Let ScrollVertical.Min = 0 + Int(Val(PictureBoxTrackMap.ScaleHeight) / 2) - Int(Val(PictureBoxTrack.ScaleHeight) / 2)
    
    Let ScrollVertical.Value = 0
    Let ScrollHorizontal.Value = 0

    Call ScrollVertical_Change
    Call ScrollHorizontal_Change

End Sub


Private Sub PictureBoxTrackMap_DragDrop(Source As Control, X As Single, Y As Single)
            
    If Source.Container.Name = "TabTrackIcon" Then
        If Source.Name = "PictureBoxTrackIconStraight" Then
            
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddStraight(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set PictureBoxTrackIconStraight(0).Container = TabTrackIcon
            Let PictureBoxTrackIconStraight(0).Top = TextBoxTrackIconStraightPositionTop.Text
            Let PictureBoxTrackIconStraight(0).Left = TextBoxTrackIconStraightPositionLeft.Text
            PictureBoxTrackIconStraight(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight1.bmp")
            Let PictureBoxTrackIconStraight(0).BorderStyle = 1
            Let PictureBoxTrackIconStraight(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxTrackIconSwitch1" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSwitchType1(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set PictureBoxTrackIconSwitch1(0).Container = TabTrackIcon
            Let PictureBoxTrackIconSwitch1(0).Top = TextBoxTrackIconSwitch1PositionTop.Text
            Let PictureBoxTrackIconSwitch1(0).Left = TextBoxTrackIconSwitch1PositionLeft.Text
            PictureBoxTrackIconSwitch1(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType1Normal.bmp")
            Let PictureBoxTrackIconSwitch1(0).BorderStyle = 1
            Let PictureBoxTrackIconSwitch1(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxTrackIconSwitch2" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSwitchType2(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set PictureBoxTrackIconSwitch2(0).Container = TabTrackIcon
            Let PictureBoxTrackIconSwitch2(0).Top = TextBoxTrackIconSwitch2PositionTop.Text
            Let PictureBoxTrackIconSwitch2(0).Left = TextBoxTrackIconSwitch2PositionLeft.Text
            PictureBoxTrackIconSwitch2(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType2Normal.bmp")
            Let PictureBoxTrackIconSwitch2(0).BorderStyle = 1
            Let PictureBoxTrackIconSwitch2(0).Visible = True
        
        ElseIf Source.Name = "PictureBoxTrackIconSwitch3" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSwitchType3(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set PictureBoxTrackIconSwitch3(0).Container = TabTrackIcon
            Let PictureBoxTrackIconSwitch3(0).Top = TextBoxTrackIconSwitch3PositionTop.Text
            Let PictureBoxTrackIconSwitch3(0).Left = TextBoxTrackIconSwitch3PositionLeft.Text
            PictureBoxTrackIconSwitch3(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType3Normal.bmp")
            Let PictureBoxTrackIconSwitch3(0).BorderStyle = 1
            Let PictureBoxTrackIconSwitch3(0).Visible = True
        
        ElseIf Source.Name = "PictureBoxTrackIconSwitch4" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSwitchType4(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set PictureBoxTrackIconSwitch4(0).Container = TabTrackIcon
            Let PictureBoxTrackIconSwitch4(0).Top = TextBoxTrackIconSwitch4PositionTop.Text
            Let PictureBoxTrackIconSwitch4(0).Left = TextBoxTrackIconSwitch4PositionLeft.Text
            PictureBoxTrackIconSwitch4(0).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType4Normal.bmp")
            Let PictureBoxTrackIconSwitch4(0).BorderStyle = 1
            Let PictureBoxTrackIconSwitch4(0).Visible = True
        
        ElseIf Source.Name = "PictureBoxSignal1" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal1(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal1(0).Container = TabTrackIcon
            Let PictureBoxSignal1(0).Top = TextBoxSignal1PositionTop.Text
            Let PictureBoxSignal1(0).Left = TextBoxSignal1PositionLeft.Text
            PictureBoxSignal1(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal1.bmp")
            Let PictureBoxSignal1(0).BorderStyle = 1
            Let PictureBoxSignal1(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal11" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal11(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal11(0).Container = TabTrackIcon
            Let PictureBoxSignal11(0).Top = TextBoxSignal11PositionTop.Text
            Let PictureBoxSignal11(0).Left = TextBoxSignal11PositionLeft.Text
            PictureBoxSignal11(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal11.bmp")
            Let PictureBoxSignal11(0).BorderStyle = 1
            Let PictureBoxSignal11(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal111" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal111(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal111(0).Container = TabTrackIcon
            Let PictureBoxSignal111(0).Top = TextBoxSignal111PositionTop.Text
            Let PictureBoxSignal111(0).Left = TextBoxSignal111PositionLeft.Text
            PictureBoxSignal111(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal111.bmp")
            Let PictureBoxSignal111(0).BorderStyle = 1
            Let PictureBoxSignal111(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal2" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal2(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal2(0).Container = TabTrackIcon
            Let PictureBoxSignal2(0).Top = TextBoxSignal2PositionTop.Text
            Let PictureBoxSignal2(0).Left = TextBoxSignal2PositionLeft.Text
            PictureBoxSignal2(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal2.bmp")
            Let PictureBoxSignal2(0).BorderStyle = 1
            Let PictureBoxSignal2(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal22" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal22(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal22(0).Container = TabTrackIcon
            Let PictureBoxSignal22(0).Top = TextBoxSignal22PositionTop.Text
            Let PictureBoxSignal22(0).Left = TextBoxSignal22PositionLeft.Text
            PictureBoxSignal22(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal22.bmp")
            Let PictureBoxSignal22(0).BorderStyle = 1
            Let PictureBoxSignal22(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal222" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal222(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal222(0).Container = TabTrackIcon
            Let PictureBoxSignal222(0).Top = TextBoxSignal222PositionTop.Text
            Let PictureBoxSignal222(0).Left = TextBoxSignal222PositionLeft.Text
            PictureBoxSignal222(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal222.bmp")
            Let PictureBoxSignal222(0).BorderStyle = 1
            Let PictureBoxSignal222(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal3" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal3(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal3(0).Container = TabTrackIcon
            Let PictureBoxSignal3(0).Top = TextBoxSignal3PositionTop.Text
            Let PictureBoxSignal3(0).Left = TextBoxSignal3PositionLeft.Text
            PictureBoxSignal3(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal3.bmp")
            Let PictureBoxSignal3(0).BorderStyle = 1
            Let PictureBoxSignal3(0).Visible = True
            
        ElseIf Source.Name = "PictureBoxSignal33" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal33(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal33(0).Container = TabTrackIcon
            Let PictureBoxSignal33(0).Top = TextBoxSignal33PositionTop.Text
            Let PictureBoxSignal33(0).Left = TextBoxSignal33PositionLeft.Text
            PictureBoxSignal33(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal33.bmp")
            Let PictureBoxSignal33(0).BorderStyle = 1
            Let PictureBoxSignal33(0).Visible = True
        
        ElseIf Source.Name = "PictureBoxSignal333" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddSignal333(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 1
            Set PictureBoxSignal333(0).Container = TabTrackIcon
            Let PictureBoxSignal333(0).Top = TextBoxSignal333PositionTop.Text
            Let PictureBoxSignal333(0).Left = TextBoxSignal333PositionLeft.Text
            PictureBoxSignal333(0).Picture = LoadPicture(App.Path$ & "\Graphics\Signal333.bmp")
            Let PictureBoxSignal333(0).BorderStyle = 1
            Let PictureBoxSignal333(0).Visible = True
        
        ElseIf Source.Name = "TextBoxMapLabel" Then
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Add New Object and Return
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call AddMapLabel(X, Y)
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Move the original back
            ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Let TabTrackIcon.Tab = 0
            Set TextBoxMapLabel(0).Container = TabTrackIcon
            Let TextBoxMapLabel(0).Top = TextboxMapLabelPositionTop.Text
            Let TextBoxMapLabel(0).Left = TextBoxMapLabelPositionLeft.Text
            'TextBoxMapLabel(0).BorderStyle = 1
            Let TextBoxMapLabel(0).Visible = True
        
        ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "DefineBlocks, PictureBoxTrackMap, DragDrop, object not a valid picture box to be droped onto map."
        End If
    Else
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + Source.Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, PictureBoxTrackMap, DragDrop, wrong tab sourse."
            End If
        End If
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Object's Location
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.Edit
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        DataBaseTrackPlan.Recordset.Update
        Let Source.Left = (Int(X / 36) * 36)
        Let Source.Top = (Int(Y / 18) * 18)
    End If
End Sub


Private Sub PictureBoxTrackMap_LostFocus()
'
' Should update the database if object moved.


End Sub


Private Sub ScrollHorizontal_Change()
    
    Let PictureBoxTrackMap.Left = ScrollHorizontal.Value - Int(Val(PictureBoxTrackMap.ScaleWidth) / 2) + Int(Val(PictureBoxTrack.ScaleWidth) / 2)

End Sub

Private Sub ScrollVertical_Change()
    
    Let PictureBoxTrackMap.Top = ScrollVertical.Value - Int(Val(PictureBoxTrackMap.ScaleHeight) / 2) + Int(Val(PictureBoxTrack.ScaleHeight) / 2)

End Sub


Private Sub TabTrackIcon_Click(PreviousTab As Integer)

'Let TextBoxTrackIconStraightPositionLeft.Text = PictureBoxTrackIconStraight(0).Left
'Let TextBoxTrackIconStraightPositionTop.Text = PictureBoxTrackIconStraight(0).Top
'Let TextBoxTrackIconSwitch1PositionLeft.Text = PictureBoxTrackIconSwitch1(0).Left
'Let TextBoxTrackIconSwitch1PositionTop.Text = PictureBoxTrackIconSwitch1(0).Top
'Let TextBoxTrackIconSwitch2PositionLeft.Text = PictureBoxTrackIconSwitch2(0).Left
'Let TextBoxTrackIconSwitch2PositionTop.Text = PictureBoxTrackIconSwitch2(0).Top
'Let TextBoxTrackIconSwitch3PositionLeft.Text = PictureBoxTrackIconSwitch3(0).Left
'Let TextBoxTrackIconSwitch3PositionTop.Text = PictureBoxTrackIconSwitch3(0).Top
'Let TextBoxTrackIconSwitch4PositionLeft.Text = PictureBoxTrackIconSwitch4(0).Left
'Let TextBoxTrackIconSwitch4PositionTop.Text = PictureBoxTrackIconSwitch4(0).Top

'Let TextBoxSignalDoublePositionTop.Text = PictureBoxSignalDouble(0).Top
'Let TextBoxSignalDoublePositionLeft.Text = PictureBoxSignalDouble(0).Left

End Sub
























Private Sub LoadStraight()
    
    Let TextBoxTrackIconStraightCounter.Text = Val(TextBoxTrackIconStraightCounter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text)
    Set PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Container = PictureBoxTrackMap
    Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).BorderStyle = 0
    Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Visible = True
    Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
        
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight1.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(Str$(DataBaseTrackPlan.Recordset.Fields("Length"))) + " feet."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight2.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block ends to the east."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight3.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block ends to the west."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight4.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block contains a block occupancy detector."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight5.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This block contains a derail."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight6.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackStraight7.bmp" Then
            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadStraight, unknown straight block picture; checka database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSwitchType1()

    Let TextBoxTrackIconSwitch1Counter.Text = Val(TextBoxTrackIconSwitch1Counter.Text) + 1
    
    Load PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text)
    Set PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Top = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Left = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).BorderStyle = 0
    PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" + DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
    Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Visible = True
    DoEvents
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType1Normal.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch1Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType1Reverse.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch1Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSwitchType1, unknown switch type picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If

End Sub

Private Sub LoadSwitchType2()
       
    Let TextBoxTrackIconSwitch2Counter.Text = Val(TextBoxTrackIconSwitch2Counter.Text) + 1
    
    Load PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text)
    Set PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Top = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Left = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).BorderStyle = 0
    PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" + DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
    Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Visible = True
    DoEvents
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType2Normal.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch2Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType2Reverse.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch2Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSwitchType2, unknown switch type picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If


End Sub

Private Sub LoadSwitchType3()
       
    Let TextBoxTrackIconSwitch3Counter.Text = Val(TextBoxTrackIconSwitch3Counter.Text) + 1
    
    Load PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text)
    Set PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Top = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Left = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).BorderStyle = 0
    PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" + DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
    Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Visible = True
    DoEvents
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType3Normal.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch3Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType3Reverse.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch3Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSwitchType3, unknown switch type picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If

End Sub

Private Sub LoadSwitchType4()
       
    Let TextBoxTrackIconSwitch4Counter.Text = Val(TextBoxTrackIconSwitch4Counter.Text) + 1
    
    Load PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text)
    Set PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Top = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Left = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).BorderStyle = 0
    PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" + DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
    Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Visible = True
    DoEvents
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType4Normal.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch4Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType4Reverse.bmp" Then
            Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch4Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSwitchType4, unknown switch type picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If

End Sub

Private Sub LoadSignal1()
    
    Let TextBoxSignal1Counter.Text = Val(TextBoxSignal1Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal1(TextBoxSignal1Counter.Text)
    Set PictureBoxSignal1(TextBoxSignal1Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal1(TextBoxSignal1Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal1(TextBoxSignal1Counter.Text).BorderStyle = 0
    Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Visible = True
    Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal1East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal1Counter.Text) + " (single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal1West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal1Counter.Text) + " (single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal1, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal1(TextBoxSignal1Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal1(TextBoxSignal1Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If

End Sub

Private Sub LoadSignal11()
       
    Let TextBoxSignal11Counter.Text = Val(TextBoxSignal11Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal11(TextBoxSignal11Counter.Text)
    Set PictureBoxSignal11(TextBoxSignal11Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal11(TextBoxSignal11Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal11(TextBoxSignal11Counter.Text).BorderStyle = 0
    Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Visible = True
    Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal11East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal11Counter.Text) + " (single over single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal11West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal11Counter.Text) + " (single over single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal11, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal11(TextBoxSignal11Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal11(TextBoxSignal11Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal111()
       
    Let TextBoxSignal111Counter.Text = Val(TextBoxSignal111Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal111(TextBoxSignal111Counter.Text)
    Set PictureBoxSignal111(TextBoxSignal111Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal111(TextBoxSignal111Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal111(TextBoxSignal111Counter.Text).BorderStyle = 0
    Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Visible = True
    Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal111East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal111Counter.Text) + " (single over single over single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal111West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal111Counter.Text) + " (single over single over single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal111, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal111(TextBoxSignal111Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal111(TextBoxSignal111Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal2()
       
    Let TextBoxSignal2Counter.Text = Val(TextBoxSignal2Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal2(TextBoxSignal2Counter.Text)
    Set PictureBoxSignal2(TextBoxSignal2Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal2(TextBoxSignal2Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal2(TextBoxSignal2Counter.Text).BorderStyle = 0
    Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Visible = True
    Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal2East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal2Counter.Text) + " (double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal2West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal2Counter.Text) + " (double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal2, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal2(TextBoxSignal2Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal2(TextBoxSignal2Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal22()
       
    Let TextBoxSignal22Counter.Text = Val(TextBoxSignal22Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal22(TextBoxSignal22Counter.Text)
    Set PictureBoxSignal22(TextBoxSignal22Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal22(TextBoxSignal22Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal22(TextBoxSignal22Counter.Text).BorderStyle = 0
    Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Visible = True
    Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal22East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal22Counter.Text) + " (double over double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal22West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal22Counter.Text) + " (double over double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal22, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal22(TextBoxSignal22Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal22(TextBoxSignal22Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal222()
       
    Let TextBoxSignal222Counter.Text = Val(TextBoxSignal222Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal222(TextBoxSignal222Counter.Text)
    Set PictureBoxSignal222(TextBoxSignal222Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal222(TextBoxSignal222Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal222(TextBoxSignal222Counter.Text).BorderStyle = 0
    Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Visible = True
    Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal222East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal222Counter.Text) + " (double over double over double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal222West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal222Counter.Text) + " (double over double over double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal222, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal222(TextBoxSignal222Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal222(TextBoxSignal222Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal3()
       
    Let TextBoxSignal3Counter.Text = Val(TextBoxSignal3Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal3(TextBoxSignal3Counter.Text)
    Set PictureBoxSignal3(TextBoxSignal3Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal3(TextBoxSignal3Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal3(TextBoxSignal3Counter.Text).BorderStyle = 0
    Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Visible = True
    Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal3East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal3Counter.Text) + " (triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal3West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal3Counter.Text) + " (triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal3, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal3(TextBoxSignal3Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal3(TextBoxSignal3Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal33()
       
    Let TextBoxSignal33Counter.Text = Val(TextBoxSignal33Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal33(TextBoxSignal33Counter.Text)
    Set PictureBoxSignal33(TextBoxSignal33Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal33(TextBoxSignal33Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal33(TextBoxSignal33Counter.Text).BorderStyle = 0
    Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Visible = True
    Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal33East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal33Counter.Text) + " (triple over triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal33West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal33Counter.Text) + " (triple over triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal33, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal33(TextBoxSignal33Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal33(TextBoxSignal33Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub

Private Sub LoadSignal333()
       
    Let TextBoxSignal333Counter.Text = Val(TextBoxSignal333Counter.Text) + 1
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Load New Object with Values
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Load PictureBoxSignal333(TextBoxSignal333Counter.Text)
    Set PictureBoxSignal333(TextBoxSignal333Counter.Text).Container = PictureBoxTrackMap
    Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Top = DataBaseTrackPlan.Recordset.Fields("PictureBoxTop")
    Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Left = DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft")
    PictureBoxSignal333(TextBoxSignal333Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\" & DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename"))
    Let PictureBoxSignal333(TextBoxSignal333Counter.Text).BorderStyle = 0
    Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Visible = True
    Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")

    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
 
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal333East.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal333Counter.Text) + " (triple over triple over triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal333West.bmp" Then
            Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal333Counter.Text) + " (triple over triple over triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        Else
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, LoadSignal333, unknown signal picture; check database field 'PictureBoxFileName'."
            End If
        End If
    'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal333(TextBoxSignal333Counter.Text))
    Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal333(TextBoxSignal333Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If
    
End Sub


Private Sub AddStraight(X, Y)
       
    Let TextBoxTrackIconStraightCounter.Text = Val(TextBoxTrackIconStraightCounter.Text) + 1
    
    If PictureBoxTrackIconStraight(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text)
        Set PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Container = PictureBoxTrackMap
        Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Left = (Int(X / 36) * 36)
        PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\TrackStraight1.bmp")
        Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).BorderStyle = 1
        Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Visible = True
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconStraight"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackStraight1.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Block"
        DataBaseTrackPlan.Recordset.Update
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If mainscreenmenuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddStraight, rcord not found in database, with object tag as the recordcounter."
            End If
        End If
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Balloon Help
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
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

            Let TemporaryText1 = "Straight Block No." + Str$(TextBoxTrackIconStraightCounter.Text) + "." & vbCrLf & "Total length of " + Str$(Str$(DataBaseTrackPlan.Recordset.Fields("Length"))) + " feet."
            Let TemporaryText2 = "Unnamed Block"
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconStraight(TextBoxTrackIconStraightCounter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub AddSwitchType1(X, Y)
       
    Let TextBoxTrackIconSwitch1Counter.Text = Val(TextBoxTrackIconSwitch1Counter.Text) + 1
    
    If PictureBoxTrackIconSwitch1(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text)
        Set PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Left = (Int(X / 36) * 36)
        PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType1Normal.bmp")
        Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).BorderStyle = 1
        Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch1"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType1Normal.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Switch Type One"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If mainscreenmenuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSwitchType1, record not found in database, with object tag as the recordcounter."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType1Normal.bmp" Then
                Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch1Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
                ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType1Reverse.bmp" Then
                    Let TemporaryText1 = "Switch Type One No." + Str$(TextBoxTrackIconSwitch1Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
                    Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSwitchType1, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch1(TextBoxTrackIconSwitch1Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
    
End Sub

Private Sub AddSwitchType2(X, Y)
    
    Let TextBoxTrackIconSwitch2Counter.Text = Val(TextBoxTrackIconSwitch2Counter.Text) + 1
    
    If PictureBoxTrackIconSwitch2(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text)
        Set PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Left = (Int(X / 36) * 36)
        PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType2Normal.bmp")
        Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).BorderStyle = 1
        Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch2"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType2Normal.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Switch Type Two"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If mainscreenmenuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks Screen, General, AddSwitchType2, error in finding the record with the proerty tag"
            End If
        End If
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
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

            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType2Normal.bmp" Then
                Let TemporaryText1 = "Switch Type Two No." + Str$(TextBoxTrackIconSwitch2Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType2Reverse.bmp" Then
                Let TemporaryText1 = "Switch Type Two No." + Str$(TextBoxTrackIconSwitch2Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
                ' Error
                ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSwitchType2, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch2(TextBoxTrackIconSwitch2Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
        
End Sub

Private Sub AddSwitchType3(X, Y)
      
    Let TextBoxTrackIconSwitch3Counter.Text = Val(TextBoxTrackIconSwitch3Counter.Text) + 1
    
    If PictureBoxTrackIconSwitch3(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text)
        Set PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Left = (Int(X / 36) * 36)
        PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType3Normal.bmp")
        Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).BorderStyle = 1
        Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch3"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType3Normal.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Switch Type Three"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSwitchTpe3, record not found in the database."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType3Normal.bmp" Then
                Let TemporaryText1 = "Switch Type Three No." + Str$(TextBoxTrackIconSwitch3Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType3Reverse.bmp" Then
                Let TemporaryText1 = "Switch Type Three No." + Str$(TextBoxTrackIconSwitch3Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSwitchType3, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch3(TextBoxTrackIconSwitch3Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub AddSwitchType4(X, Y)
    
    Let TextBoxTrackIconSwitch4Counter.Text = Val(TextBoxTrackIconSwitch4Counter.Text) + 1
   
    If PictureBoxTrackIconSwitch4(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text)
        Set PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Left = (Int(X / 36) * 36)
        PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\TrackSwitchType4Normal.bmp")
        Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).BorderStyle = 1
        Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch4"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "TrackSwitchType4Normal.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Switch Type Four"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSwitchTpe4, record not found in the database."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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

            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType4Normal.bmp" Then
                Let TemporaryText1 = "Switch Type Four No." + Str$(TextBoxTrackIconSwitch4Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the NORMAL position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "TrackSwitchType4Reverse.bmp" Then
                Let TemporaryText1 = "Switch Type Four No." + Str$(TextBoxTrackIconSwitch4Counter.Text) + "." & vbCrLf & "Total length of " + Str$(DataBaseTrackPlan.Recordset.Fields("Length")) + " feet." & vbCrLf & "This switch is in the REVERSE position."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSwitchType4, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxTrackIconSwitch4(TextBoxTrackIconSwitch4Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If

End Sub

Private Sub AddSignal1(X, Y)
     
    Let TextBoxSignal1Counter.Text = Val(TextBoxSignal1Counter.Text) + 1
       
    If PictureBoxSignal1(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxSignal1(TextBoxSignal1Counter.Text)
        Set PictureBoxSignal1(TextBoxSignal1Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Left = (Int(X / 36) * 36)
        Let PictureBoxSignal1(TextBoxSignal1Counter.Text).BorderStyle = 1
        PictureBoxSignal1(TextBoxSignal1Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal1East.bmp")
        Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxSignal1(TextBoxSignal1Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal1"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal1East.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal1(TextBoxSignal1Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSwitchType1, record not found in the database."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal1East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal1Counter.Text) + "(single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal1West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal1Counter.Text) + "(single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal1, wrong picture for new picture box to intitialize help balloon."
            End If
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal1(TextBoxSignal1Counter.Text))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal1(TextBoxSignal1Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub AddSignal11(X, Y)
       
    Let TextBoxSignal11Counter.Text = Val(TextBoxSignal11Counter.Text) + 1

    If PictureBoxSignal11(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal11(TextBoxSignal11Counter.Text)
         Set PictureBoxSignal11(TextBoxSignal11Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal11(TextBoxSignal11Counter.Text).BorderStyle = 1
         PictureBoxSignal11(TextBoxSignal11Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal11East.bmp")
         Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal11(TextBoxSignal11Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal11"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal11East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal11(TextBoxSignal11Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal1, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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

            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal11East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal11Counter.Text) + "(single over single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal11West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal11Counter.Text) + "(single over single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal2, wrong picture for new picture box to intitialize help balloon."
             End If
             'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal11(TextBoxSignal11Counter.Text))
             Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal11(TextBoxSignal11Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
         End If
    End If
End Sub

Private Sub AddSignal111(X, Y)
     
    Let TextBoxSignal111Counter.Text = Val(TextBoxSignal111Counter.Text) + 1
    
    If PictureBoxSignal111(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal111(TextBoxSignal111Counter.Text)
         Set PictureBoxSignal111(TextBoxSignal111Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal111(TextBoxSignal111Counter.Text).BorderStyle = 1
         PictureBoxSignal111(TextBoxSignal111Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal111East.bmp")
         Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal111(TextBoxSignal111Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal111"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal111East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal111(TextBoxSignal111Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal111, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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

            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal111East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal111Counter.Text) + "(single over single over single target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal111West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal111Counter.Text) + "(single over single over single target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal111, wrong picture for new picture box to intitialize help balloon."
             End If
             'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal111(TextBoxSignal111Counter.Text))
             Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal111(TextBoxSignal111Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
         End If
     End If

End Sub

Private Sub AddSignal2(X, Y)
       
    Let TextBoxSignal2Counter.Text = Val(TextBoxSignal2Counter.Text) + 1
    
    If PictureBoxSignal2(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal2(TextBoxSignal2Counter.Text)
         Set PictureBoxSignal2(TextBoxSignal2Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal2(TextBoxSignal2Counter.Text).BorderStyle = 1
         PictureBoxSignal2(TextBoxSignal2Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal2East.bmp")
         Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal2(TextBoxSignal2Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal2"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal2East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal2(TextBoxSignal2Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSgnal2, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal2East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal2Counter.Text) + "(double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal2West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal2Counter.Text) + "(double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal2, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal2(TextBoxSignal2Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal2(TextBoxSignal2Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If

End Sub

Private Sub AddSignal22(X, Y)
       
    Let TextBoxSignal22Counter.Text = Val(TextBoxSignal22Counter.Text) + 1
    
    If PictureBoxSignal22(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxSignal22(TextBoxSignal22Counter.Text)
        Set PictureBoxSignal22(TextBoxSignal22Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Left = (Int(X / 36) * 36)
        Let PictureBoxSignal22(TextBoxSignal22Counter.Text).BorderStyle = 1
        PictureBoxSignal22(TextBoxSignal22Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal22East.bmp")
        Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxSignal22(TextBoxSignal22Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal22"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal22East.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal22(TextBoxSignal22Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal22, record not found in the database."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal22East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal22Counter.Text) + "(double over double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal22West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal22Counter.Text) + "(double over double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal22, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal22(TextBoxSignal22Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal22(TextBoxSignal22Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub AddSignal222(X, Y)
       
    Let TextBoxSignal222Counter.Text = Val(TextBoxSignal222Counter.Text) + 1
    
    If PictureBoxSignal222(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal222(TextBoxSignal222Counter.Text)
         Set PictureBoxSignal222(TextBoxSignal222Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal222(TextBoxSignal222Counter.Text).BorderStyle = 1
         PictureBoxSignal222(TextBoxSignal222Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal222East.bmp")
         Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal222(TextBoxSignal222Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal222"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal222East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal222(TextBoxSignal222Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal222, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal222East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal222Counter.Text) + "(double over double over double target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal222West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal222Counter.Text) + "(double over double over double target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal222, wrong picture for new picture box to intitialize help balloon."
             End If
             'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal222(TextBoxSignal222Counter.Text))
             Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal222(TextBoxSignal222Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
         End If
     End If

End Sub

Private Sub AddSignal3(X, Y)
       
    Let TextBoxSignal3Counter.Text = Val(TextBoxSignal3Counter.Text) + 1
   
    If PictureBoxSignal3(0).Container.Name <> "PictureBoxTrackMap" Then
        Load PictureBoxSignal3(TextBoxSignal3Counter.Text)
        Set PictureBoxSignal3(TextBoxSignal3Counter.Text).Container = PictureBoxTrackMap
        Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Top = (Int(Y / 18) * 18)
        Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Left = (Int(X / 36) * 36)
        Let PictureBoxSignal3(TextBoxSignal3Counter.Text).BorderStyle = 1
        PictureBoxSignal3(TextBoxSignal3Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal3East.bmp")
        Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Visible = True
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let PictureBoxSignal3(TextBoxSignal3Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal3"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal3East.bmp"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal3(TextBoxSignal3Counter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal3, record not found in the database."
            End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal3East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal3Counter.Text) + "(triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal3West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal3Counter.Text) + "(triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal3, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal3(TextBoxSignal3Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal3(TextBoxSignal3Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If

End Sub

Private Sub AddSignal33(X, Y)
 
    Let TextBoxSignal33Counter.Text = Val(TextBoxSignal33Counter.Text) + 1
    
    If PictureBoxSignal33(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal33(TextBoxSignal33Counter.Text)
         Set PictureBoxSignal33(TextBoxSignal33Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal33(TextBoxSignal33Counter.Text).BorderStyle = 1
         PictureBoxSignal33(TextBoxSignal33Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal33East.bmp")
         Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal33(TextBoxSignal33Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal33"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal33East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal33(TextBoxSignal33Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal33, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal33East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal33Counter.Text) + "(triple over triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signa33West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal33Counter.Text) + "(triple over triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal33, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal33(TextBoxSignal33Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal33(TextBoxSignal33Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub AddSignal333(X, Y)
       
    Let TextBoxSignal333Counter.Text = Val(TextBoxSignal333Counter.Text) + 1
    
    If PictureBoxSignal333(0).Container.Name <> "PictureBoxTrackMap" Then
         Load PictureBoxSignal333(TextBoxSignal333Counter.Text)
         Set PictureBoxSignal333(TextBoxSignal333Counter.Text).Container = PictureBoxTrackMap
         Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Top = (Int(Y / 18) * 18)
         Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Left = (Int(X / 36) * 36)
         Let PictureBoxSignal333(TextBoxSignal333Counter.Text).BorderStyle = 1
         PictureBoxSignal333(TextBoxSignal333Counter.Text).Picture = LoadPicture(App.Path$ & "\Graphics\Signal333East.bmp")
         Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Visible = True
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Update Database
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.AddNew
         Let PictureBoxSignal333(TextBoxSignal333Counter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal333"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "Signal333East.bmp"
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = (Int(X / 36) * 36)
         Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = (Int(Y / 18) * 18)
         Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
         Let DataBaseTrackPlan.Recordset.Fields("Name") = "Unnamed Signal"
         DataBaseTrackPlan.Recordset.Update
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Move Database Pointer Back
         ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
         DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + PictureBoxSignal333(TextBoxSignal333Counter.Text).Tag
         If DataBaseTrackPlan.Recordset.NoMatch Then
             If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddSignal333, record not found in the database."
             End If
         End If
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         ' Balloon Help
         ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            If DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signal333East.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal333Counter.Text) + "(triple over triple over triple target)." & vbCrLf & "This signal is for EAST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxFileName") = "Signa333West.bmp" Then
                Let TemporaryText1 = "Signal No." + Str$(TextBoxSignal333Counter.Text) + "(triple over triple over triple target)." & vbCrLf & "This signal is for WEST bound trains." & vbCrLf & "This signal is RED."
                Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Error
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "Define Blocks, General, AddSignal1, wrong picture for new picture box to intitialize help balloon."
            End If
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(PictureBoxSignal333(TextBoxSignal333Counter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(PictureBoxSignal333(TextBoxSignal333Counter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub
Private Sub TextboxMapLabel_GotFocus(Index As Integer)

    'Let TextBoxMapLabel(Index).BorderStyle = 1

    Let LabelStatus.Caption = "Status: Text Box (label) No. " + Str$(Index) + " of " + Str$(TextBoxMapLabelCounter.Text)

End Sub

Private Sub TextboxMapLabel_KeyPress(Index As Integer, KeyAscii As Integer)

' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Update Database
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Find Record
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + TextBoxMapLabel(Index).Tag
    If DataBaseTrackPlan.Recordset.NoMatch Then
        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "DefineBlocks, TextBoxMapLabel, Keypress, record not found in the database."
        End If
    Else
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update'
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.Edit
        DataBaseTrackPlan.Recordset.Fields("Comments") = TextBoxMapLabel(Index).Text
        DataBaseTrackPlan.Recordset.Update
    End If
    
End Sub


Private Sub TextboxMapLabel_LostFocus(Index As Integer)

    'If TextBoxLabel(Index).Container.Name = "PictureBoxTrackMap" Then
    '    TextBoxLabel(Index).BorderStyle = 0
    'ElseIf TextBoxLabel(Index).Container.Name = "TabTrackIcon" Then
    '    TextBoxLabel(Index).BorderStyle = 1
    'End If

    Let LabelStatus.Caption = "Status:"

End Sub

Private Sub TextboxMapLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    TextBoxMapLabel(Index).Drag vbBeginDrag

End Sub


Private Sub Timer1_Timer()

    Let LabelStatus = "Status: Loading track plan database. Please Wait."
    
    Timer1.Interval = 0
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Begin Reading Database
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    While Not DataBaseTrackPlan.Recordset.EOF
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' What Object is it?
        ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconStraight" Then
            Call LoadStraight
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch1" Then
            Call LoadSwitchType1
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch2" Then
            Call LoadSwitchType2
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch3" Then
            Call LoadSwitchType3
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch4" Then
            Call LoadSwitchType4
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal1" Then
            Call LoadSignal1
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal11" Then
            Call LoadSignal11
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal111" Then
            Call LoadSignal111
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal2" Then
            Call LoadSignal2
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal22" Then
            Call LoadSignal22
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal222" Then
            Call LoadSignal222
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal3" Then
            Call LoadSignal3
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal33" Then
            Call LoadSignal33
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "PictureBoxSignal333" Then
            Call LoadSignal333
        ElseIf DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "TextBoxMapLabel" Then
            Call LoadMapLabel
        ElseIf MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
            Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
            Let Ini.Filename = App.Path$ & "\Atc.log"
            Let Ini.Application = "Log Errors"
            Let Ini.Parameter = Date$ & " " & Time$
            Let Ini.Value = "DefineBlocks, Timer1, Timer, unknown type of object trying to load from database."
        End If
        DataBaseTrackPlan.Recordset.MoveNext
        DoEvents
    Wend

    Let LabelStatus = "Status: Finished loading track plan database."

End Sub






Private Sub AddMapLabel(X, Y)
       
    Let TextBoxMapLabelCounter.Text = Val(TextBoxMapLabelCounter.Text) + 1

    If TextBoxMapLabel(0).Container.Name <> "PictureBoxTrackMap" Then
        Load TextBoxMapLabel(TextBoxMapLabelCounter.Text)
        Set TextBoxMapLabel(TextBoxMapLabelCounter.Text).Container = PictureBoxTrackMap
        Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Top = Y
        Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Left = X
        Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Visible = True
       ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Update Database
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.AddNew
        Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxName") = "TextBoxMapLabel"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxFilename") = "None"
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft") = X
        Let DataBaseTrackPlan.Recordset.Fields("PictureBoxTop") = Y
        Let DataBaseTrackPlan.Recordset.Fields("Length") = 0
        Let DataBaseTrackPlan.Recordset.Fields("Name") = "Map Label"
        Let DataBaseTrackPlan.Recordset.Fields("Comments") = TextBoxMapLabel(TextBoxMapLabelCounter.Text).Text
        DataBaseTrackPlan.Recordset.Update
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Move Database Pointer Back
        ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        DataBaseTrackPlan.Recordset.FindFirst "RecordCounter = " + TextBoxMapLabel(TextBoxMapLabelCounter.Text).Tag
        If DataBaseTrackPlan.Recordset.NoMatch Then
            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                Let Ini.Filename = App.Path$ & "\Atc.log"
                Let Ini.Application = "Log Errors"
                Let Ini.Parameter = Date$ & " " & Time$
                Let Ini.Value = "DefineBlocks, General, AddMapLabel, record not found in the database."
             End If
        End If
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Balloon Help
        ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            
            Let TemporaryText1 = "Map Label No." + Str$(TemporaryIndex) + "."
            Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
            'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxMapLabel(TextBoxMapLabelCounter.Text))
            Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxMapLabel(TextBoxMapLabelCounter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
        End If
    End If
End Sub

Private Sub LoadMapLabel()
   
    Let TextBoxMapLabelCounter.Text = Val(TextBoxMapLabelCounter.Text) + 1
    
    Load TextBoxMapLabel(TextBoxMapLabelCounter.Text)
    Set TextBoxMapLabel(TextBoxMapLabelCounter.Text).Container = PictureBoxTrackMap
    Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Top = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxTop"))
    Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Left = Val(DataBaseTrackPlan.Recordset.Fields("PictureBoxLeft"))
    Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Text = DataBaseTrackPlan.Recordset.Fields("Comments")
    Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Tag = DataBaseTrackPlan.Recordset.Fields("RecordCounter")
    Let TextBoxMapLabel(TextBoxMapLabelCounter.Text).Visible = True
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Balloon Help
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
        
        Let TemporaryText1 = "Label No." + Str$(TextBoxMapLabelCounter.Text) + "."
        Let TemporaryText2 = DataBaseTrackPlan.Recordset.Fields("Name")
        'let BalloonHelpSetup = balloonhelp.DestroyToolTip(TextBoxMapLabel(TextBoxMapLabelCounter.Text))
        Let BalloonHelpSetup = balloonhelp.AddToolTip(TextBoxMapLabel(TextBoxMapLabelCounter.Text), BalloonText1, balBalloon, BalloonHelpText2, balInfo, RGB(BalloonHelpColour1, BalloonHelpColour2, BalloonHelpColour3), 0, BalloonHelpVisibleTime, BalloonHelpDelayTime, BalloonHelpShadow, BalloonHelpCenter, BalloonHelpShowOnDemand, BalloonHelpOpacity, BalloonHelpFont, App.Path$ & "\Help\info.ico", 10, 10, App.Path$ & "\Help\Balloon.wav")
    End If

End Sub


