VERSION 4.00
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATC - Password Authenitcation"
   ClientHeight    =   3345
   ClientLeft      =   14100
   ClientTop       =   8940
   ClientWidth     =   3705
   Height          =   3750
   Icon            =   "PasswordMakerVersion4.5.3.0.frx":0000
   Left            =   14040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3705
   Top             =   8595
   Width           =   3825
   Begin VB.TextBox TextboxUserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TextboxPassword 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox TextboxDate 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox TextboxVersion 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox TextboxComputerName 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User's Name"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   3600
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label7 
      Caption         =   "Password"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Computer's Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "User Indentification"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of File"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Automatic Train Control "
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Password"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub ButtonClose_Click()

End
    
End Sub

Private Sub Command1_Click()

End

End Sub

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.


    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub TextboxComputerName_Change()

    Let AsciiTotal = 0
    
    For x = 1 To Len(TextboxUserName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxUserName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxComputerName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxComputerName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, x, 1))
    Next x

    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

Private Sub TextboxDate_Change()

    Let AsciiTotal = 0
    
    For x = 1 To Len(TextboxUserName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxUserName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxComputerName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxComputerName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxDate.Text)
            Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, x, 1))
    Next x

    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

Private Sub Timer1_Timer()

End Sub
Private Function GetTheUserName() As String

End Function


Private Sub TextboxName_KeyUp(KeyCode As Integer, Shift As Integer)

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
    
    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

Private Sub TextboxUserName_Change()

    Let AsciiTotal = 0
    
    For x = 1 To Len(TextboxUserName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxUserName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxComputerName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxComputerName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, x, 1))
    Next x

    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

Private Sub TextboxVersion_Change()

    Let AsciiTotal = 0
    
    For x = 1 To Len(TextboxUserName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxUserName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxComputerName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxComputerName.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, x, 1))
    Next x
    
    For x = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, x, 1))
    Next x

    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

