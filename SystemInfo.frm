VERSION 4.00
Begin VB.Form SystemInfo 
   Caption         =   "Automatic Train Control - System Information"
   ClientHeight    =   4680
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5745
   Height          =   5085
   Icon            =   "SystemInfo.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   5745
   Top             =   1170
   Width           =   5865
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label LabelDriveInformationC 
      AutoSize        =   -1  'True
      Caption         =   "Drive Information for 'C:'"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Label LabelMemoryInformation 
      AutoSize        =   -1  'True
      Caption         =   "Memory Information"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Label LabelApplicationDirectory 
      AutoSize        =   -1  'True
      Caption         =   "Application Directory"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label LabelSystemDirectory 
      AutoSize        =   -1  'True
      Caption         =   "System Directory"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Label LabelWindowsDirectory 
      AutoSize        =   -1  'True
      Caption         =   "Windows Directory"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label LabelWindowsVersion 
      AutoSize        =   -1  'True
      Caption         =   "Windows Version"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label LabelLastRebootState 
      AutoSize        =   -1  'True
      Caption         =   "Last Reboot State"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1290
   End
   Begin VB.Label LabelTimeSinceReboot 
      AutoSize        =   -1  'True
      Caption         =   "Time Since Reboot"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label LabelLogonServer 
      AutoSize        =   -1  'True
      Caption         =   "Logon Server"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label LabelDomainName 
      AutoSize        =   -1  'True
      Caption         =   "Domain Name"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label LabelNetwork 
      AutoSize        =   -1  'True
      Caption         =   "Network"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label LabelComputerName 
      AutoSize        =   -1  'True
      Caption         =   "Computer Name"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label LabelUserName 
      AutoSize        =   -1  'True
      Caption         =   "User Name "
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   840
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "SystemInfo.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"SystemInfo.frx":0884
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "SystemInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub ButtonClose_Click()

' =========================================================================================================================
' Close the SystemInfo Window
'
' =========================================================================================================================
' Hide Method
'
' Hides an MDIForm or Form object but doesn't unload it.
'
' The syntax is as follows,
'
' object.Hide
'
' The object placeholder represents an object expression that evaluates to an object in the Applies To list.  If object
' is omitted, the form with the focus is assumed to be object. When a form is hidden, it's removed from the screen and
' its Visible property is set to False.  A hidden form's controls aren't accessible to the user, but they are available
' to the running Visual Basic application, to other processes that may be communicating with the application
' through DDE, and to Timer control events. When a form is hidden, the user can't interact with the application
' until all code in the event procedure that caused the form to be hidden has finished executing. If the form isn't
' loaded when the Hide method is invoked, the Hide method loads the form but doesn't display it.

    SystemInfo.Hide
    
' =========================================================================================================================
' Uload Statement
'
' Unloads a form or control from memory.
'
' The syntax is as follows
'
'   Unload object
'
' The object placeholder is the name of a Form object or control array element to unload. Unloading a form or control may
' be necessary or expedient in some cases where the memory used is needed for something else or when you need to reset
' properties to their original values. Before a form is unloaded, the Query_Unload event procedure occurs, followed by
' the Form_Unload event procedure.  Setting the cancel argument to True in either of these events keeps the form from
' being unloaded.  For MDIForm objects, the MDIForm object's Query_Unload event procedure occurs, followed by
' the Query_Unload event procedure and Form_Unload event procedure for each MDI child form, and finally the MDIForm
' object's Form_Unload event procedure. When a form is unloaded, all controls placed on the form at run time are no
' longer accessible.  Controls placed on the form at design time remain intact; however, any run-time changes to
' those controls and their properties are lost when the form is reloaded.  All changes to form properties are
' also lost. When a form is unloaded, only the displayed component is unloaded.  The code associated with the form
' module remains in memory. Only control array elements added to a form at run time can be unloaded with the Unload
' statement.  The properties of unloaded controls are reinitialized when the controls are reloaded.

    Unload SystemInfo

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' Syntax is in the following format
'
'   End Sub
'
' End Sub Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions
' that apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form verti'cally.

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


