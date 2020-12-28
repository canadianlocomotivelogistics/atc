VERSION 4.00
Begin VB.Form DebugMode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Debug Mode"
   ClientHeight    =   7755
   ClientLeft      =   3780
   ClientTop       =   5145
   ClientWidth     =   10290
   Enabled         =   0   'False
   Height          =   8160
   Icon            =   "DebugMode.frx":0000
   Left            =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Top             =   4800
   Width           =   10410
   Begin VB.CheckBox CheckBoxDebugMode 
      Caption         =   "Do you want to log to a file?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton ButtonPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   7740
      TabIndex        =   5
      Top             =   7440
      Width           =   1155
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox TextBoxDebugMode 
      Enabled         =   0   'False
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   10095
   End
   Begin VB.TextBox LabelDebugModeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "DebugMode.frx":0442
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox LabelDebugMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "DebugMode.frx":0456
      Top             =   120
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "DebugMode.frx":04D5
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin ctlAlphaBlend.AlphaBlend AlphaBlend 
      Left            =   9180
      Top             =   660
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin IniconLib.Init Ini 
      Left            =   9720
      Top             =   660
      _Version        =   196611
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Application     =   ""
      Parameter       =   ""
      Value           =   ""
      Filename        =   ""
   End
End
Attribute VB_Name = "DebugMode"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Activate()

    DoEvents
    
    MainScreen.Visible = True
    
End Sub

Private Sub Form_Load()

    DebugMode.Top = 1
    DebugMode.Left = 1

End Sub


Private Sub Form_Resize()

    If DebugMode.WindowState = vbMinimized Then
    
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
        
    ElseIf DebugMode.WindowState = vbNormal Then
    
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


Private Sub TextBoxDebugMode_Change()

    Dim tempExitEarly As Boolean
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Displaying Last 30,000 Characters
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Len(TextBoxDebugMode.Text) > 30000 Then
        Let TextBoxDebugMode.Text = Right$(TextBoxDebugMode.Text, 30000)
        Let tempExitEarly = True
    End If
    Let TextBoxDebugMode.SelStart = Len(TextBoxDebugMode.Text)

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Record last line to File
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If tempExitEarly = True Then Exit Sub: Rem stop duplication of event caused by truncating the text
    If CheckBoxDebugMode.Value = vbChecked Then
        If Right$(TextBoxDebugMode.Text, 2) = vbCrLf Then
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            ' Wright to Local File & Network File
            ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
            On Error Resume Next
            Let Ini.Filename = App.Path$ & "\Atc.ini"
            Let Ini.Application = "Main Screen"
            Let Ini.Parameter = "NetworkPath"
            If Ini.Value <> "Not Used" Then
                Open App.Path & "\AtcDebugMode.txt" For Append As #1
                If Err <> 0 Then
                    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        Let Ini.Filename = App.Path$ & "\Atc.log"
                        Let Ini.Application = "Log Errors"
                        Let Ini.Parameter = Date$ & " " & Time$
                        Let Ini.Value = "Debug Mode, TextBoxCommunicationWindowDCC, Change, invalid network path or filename (" & Err & "," & Error$ & ")."
                    End If
                End If
            End If
            Open App.Path$ & "\AtcDebugMode.Txt" For Append As #2
                If Err <> 0 Then
                    If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                        Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                        MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                        Let Ini.Filename = App.Path$ & "\Atc.log"
                        Let Ini.Application = "Log Errors"
                        Let Ini.Parameter = Date$ & " " & Time$
                        Let Ini.Value = "Debug Mode, TextBoxCommunicationWindowDCC, Change, unable to open file (" & Err & "," & Error$ & ")."
                    End If
                End If
            For TemporaryValue = Len(TextBoxDebugMode.Text) - 3 To 0 Step -1
                If Mid$(TextBoxDebugMode.Text, TemporaryValue, 2) = vbCrLf Then
                    If Ini.Value <> "Not Used" Then
                        Print #1, Mid$(TextBoxDebugMode.Text, TemporaryValue + 2, Len(TextBoxDebugMode.Text) - 2);
                        If Err <> 0 Then
                            If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                                Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                                MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                                Let Ini.Filename = App.Path$ & "\Atc.log"
                                Let Ini.Application = "Log Errors"
                                Let Ini.Parameter = Date$ & " " & Time$
                                Let Ini.Value = "Debug Mode, TextBoxCommunicationWindowDCC, Change, unable to append file (" & Err & "," & Error$ & ")."
                            End If
                        End If
                    End If
                    
                    Print #2, Mid$(TextBoxDebugMode.Text, TemporaryValue + 2, Len(TextBoxDebugMode.Text) - 2);
                    If Err <> 0 Then
                        If MainScreen!menuDebugMode.Caption = "&Debug Mode is On" Then
                           Let TemporaryMessage = "An error has occured with Automatic Train Control. This error will be recorded in the ATC.LOG file. Please email the author reporting the error and attach a copy of the file called ATC.LOG for detailed information. This program will continue, but it may not function correctly."
                            MsgBox TemporaryMessage, vbOKOnly + vbInformation, "Automatic Train Control - Warning"
                            Let Ini.Filename = App.Path$ & "\Atc.log"
                            Let Ini.Application = "Log Errors"
                            Let Ini.Parameter = Date$ & " " & Time$
                            Let Ini.Value = "Debug Mode, TextBoxCommunicationWindowDCC, Change, invalid network path in ATC.INI file (" & Err & "," & Error$ & ")."
                        End If
                    End If
                    Let TemporaryValue = -2
                End If
            Next TemporaryValue
            If TemporaryValue = -1 Then
                Print #1, TextBoxDebugMode.Text;
            End If
            Close #2
            Close #1
            On Error GoTo 0
        End If
    End If
End Sub


