VERSION 4.00
Begin VB.Form ScreenResolutionTest 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   1320
   ClientTop       =   345
   ClientWidth     =   8475
   Height          =   11520
   Left            =   1260
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   8475
   Top             =   0
   Width           =   8595
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox TextBoxScreenHeightPixels 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox TextBoxScreenHeight 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   "0"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox TextBoxScreenWidthPixels 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox TextBoxScreenWidth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "twips or"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "pixels."
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "pixels."
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Screen Height is currently"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "twips or"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LabelScreenWidth 
      Alignment       =   1  'Right Justify
      Caption         =   "Screen Width is currently"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "ScreenResolutionTest"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub ButtonClose_Click()

End

End Sub

Private Sub Form_Load()

' =============================================================================================================================================================================
' Checking the Screen Resolution
' =============================================================================================================================================================================
'
'   Every time a new window is opened in Autoamtic Train Control we check the screen size and compare it to the window screen size.
' If the window cannot be displayed in the current screen size a message box is displayed. This allows time for the user to change
' the screen attributes to correct size.

Let TextBoxScreenWidth = Screen.Width
Let TextBoxScreenHeight = Screen.Height

Do
    Let TextBoxScreenWidth = Screen.Width
    Let TextBoxScreenHeight = Screen.Height
    
    If Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY Then
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    End If
Loop While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY

End Sub


Private Sub TextBoxScreenHeight_Change()


Let TextBoxScreenHeightPixels = TextBoxScreenHeight.Text / 15

End Sub

Private Sub TextBoxScreenWidth_Change()

Let TextBoxScreenWidthPixels = TextBoxScreenWidth.Text / 15

End Sub


