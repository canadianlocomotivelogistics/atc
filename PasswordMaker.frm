VERSION 4.00
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATC - Password Authenitcation"
   ClientHeight    =   2985
   ClientLeft      =   3585
   ClientTop       =   4605
   ClientWidth     =   3705
   Height          =   3390
   Icon            =   "PasswordMaker.frx":0000
   Left            =   3525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3705
   Top             =   4260
   Width           =   3825
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TextboxPassword 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
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
   Begin VB.TextBox TextboxName 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
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
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Your Name"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "User Indentification"
      Height          =   255
      Left            =   240
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Automatic Train Control "
      Height          =   255
      Left            =   240
      TabIndex        =   6
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


Private Sub TextboxDate_KeyUp(KeyCode As Integer, Shift As Integer)

    Let AsciiTotal = 0
    
    For X = 1 To Len(TextboxName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxName.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, X, 1))
    Next X

    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub


Private Sub Timer1_Timer()

End Sub
Private Function GetTheUserName() As String

    Dim lngRetVal As Long
    Dim lpBuffer As String
    Dim nSize As Long
    
    lpBuffer = Space(255)
    nSize = 254
    lngRetVal = GetUserName(lpBuffer, nSize)
    GetTheUserName = lpBuffer

End Function


Private Sub TextboxName_KeyUp(KeyCode As Integer, Shift As Integer)

    Let AsciiTotal = 0
    
    For X = 1 To Len(TextboxName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxName.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, X, 1))
    Next X
    
    Let TextboxPassword.Text = AsciiTotal ^ 3

End Sub

Private Sub TextboxVersion_KeyUp(KeyCode As Integer, Shift As Integer)

    Let AsciiTotal = 0
    
    For X = 1 To Len(TextboxName.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxName.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxDate.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxDate.Text, X, 1))
    Next X
    
    For X = 1 To Len(TextboxVersion.Text)
        Let AsciiTotal = AsciiTotal + Asc(Mid$(TextboxVersion.Text, X, 1))
    Next X
  
    Let TextboxPassword.Text = AsciiTotal ^ 3
    
End Sub


