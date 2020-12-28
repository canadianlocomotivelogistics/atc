VERSION 4.00
Begin VB.Form LayoutCTC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Define CTC"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   1335
   ClientWidth     =   11595
   Height          =   7470
   Icon            =   "LayoutCTC.frx":0000
   Left            =   -15
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Top             =   990
   Width           =   11715
   Begin VB.CommandButton CommandLoadDatabase 
      Caption         =   "Load Database"
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\TrackPlanDatabase.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TrackPlan"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox ComboBoxTrack 
      Height          =   315
      ItemData        =   "LayoutCTC.frx":0442
      Left            =   120
      List            =   "LayoutCTC.frx":0470
      TabIndex        =   3
      Text            =   "Please select a track"
      Top             =   5040
      Width           =   2775
   End
   Begin VB.PictureBox PictureBoxTrack 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   1680
      Picture         =   "LayoutCTC.frx":05E8
      ScaleHeight     =   300
      ScaleWidth      =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1260
   End
   Begin VB.TextBox TextBoxTrackComment 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "No Information Available"
      Top             =   6480
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "LayoutCTC.frx":106A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   10200
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   120
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label3 
      Caption         =   "If you want text to be included with your section of track, type it here."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "You selected this piece of track."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "LayoutCTC.frx":14AC
      DragMode        =   1  'Automatic
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   10
      Cols            =   100
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      AllowBigSelection=   0   'False
      FillStyle       =   1
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Here we can run out centralized traffic control with the layout we have designed."
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "LayoutCTC"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ButtonClose_Click()

' =========================================================================================================================
' Close the Programming Diesel Window
'
   
' My Programming Notes
'
' Now that the commubication port is opened (hopefully to the correct communicaton port connected to the North Coast
' Enginerring Power HOuse Pro) the additional options on the menu bar can be enabled or disabled.
    
'    Let MainScreen.MenuCommunicationPortOpen.Enabled = False
'    Let MainScreen.MenuCommunicationPortClosed.Enabled = True
'    Let MainScreen.MenuCommunicationPortSetting.Enabled = False
'    Let MainScreen.MenuScaledTimeSetting.Enabled = True
'    Let MainScreen.MenuProgrammingModeDiesel.Enabled = True
'    Let MainScreen.MenuProgrammingModeSteam.Enabled = True
'    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = True
'    Let MainScreen.MenuProgrammingModeSteam.Enabled = True
'    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = True
'    Let MainScreen.MenuProgrammingModeOther.Enabled = True
'    Let MainScreen.MenuMainlineDieselProgramming.Enabled = True
'    Let MainScreen.MenuMainlineSteamProgramming.Enabled = True
'    Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = True
'    Let MainScreen.MenuMainlineOtherProgramming.Enabled = True
'    Let MainScreen.menumainlineconsist.Enabled = True
'    Let MainScreen.MenuMainlineOperation.Enabled = True
'    Let MainScreen.MenuMainlineOperationGUI.Enabled = True
'    Let MainScreen.MenuMainlineMacroMaker.Enabled = True
    Let MainScreen.MenuLayoutDefineBlocks.Enabled = True
    Let MainScreen.MenuLayoutCTC.Enabled = True

' =========================================================================================================================
' Hide Method
'
'Hides an MDIForm or Form object but doesn't unload it.
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

    LayoutCTC.Hide
    
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

    Unload LayoutCTC

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


Private Sub Command1_Click()

MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 1
Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\BellOn.bmp")

End Sub

Private Sub ComboBoxTrack_Click()

If ComboBoxTrack.Text = "None" Then
    PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackNone.bmp")
Else
    If ComboBoxTrack.Text = "Straight" Then
        PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackStraight.bmp")
    Else
        If ComboBoxTrack.Text = "Diagnal (top left, bottom right)" Then
            PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal0.bmp")
        Else
            If ComboBoxTrack.Text = "Diagnal (top right, bottom left)" Then
                PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal1.bmp")
            Else
                If ComboBoxTrack.Text = "Turnout (eastbound, right)" Then
                    PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutER.bmp")
                Else
                    If ComboBoxTrack.Text = "Turnout (eastbound, left)" Then
                        PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutEL.bmp")
                    Else
                        If ComboBoxTrack.Text = "Turnout (westbound, right)" Then
                            PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWR.bmp")
                        Else
                            If ComboBoxTrack.Text = "Turnout (westbound, left)" Then
                                PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWL.bmp")
                            Else
                                If ComboBoxTrack.Text = "Curve (eastbound, right)" Then
                                    PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackCurveER.bmp")
                                Else
                                    If ComboBoxTrack.Text = "Curve (eastbound, left)" Then
                                        PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackCurveEL.bmp")
                                    Else
                                        If ComboBoxTrack.Text = "Curve (westbound, right)" Then
                                            PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWR.bmp")
                                        Else
                                            If ComboBoxTrack.Text = "Curve (westbound, left)" Then
                                                PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWL.bmp")
                                            Else
                                                If ComboBoxTrack.Text = "Double Slip (top left, botton right)" Then
                                                    PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip0.bmp")
                                                Else
                                                    If ComboBoxTrack.Text = "Double Slip (top right, bottom left)" Then
                                                        PictureBoxTrack.Picture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip1.bmp")
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If


End Sub


Private Sub CommandLoadDatabase_Click()

For t = 0 To 99
    Let MSFlexGrid1.Col = t
    Let MSFlexGrid1.Row = 0
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    If Data1.Recordset.Fields("0TrackComments") <> "" Then
        Let MSFlexGrid1.Text = Data1.Recordset.Fields("0TrackComments")
    End If
    If Data1.Recordset.Fields("0TrackPicture") <> "" Then
    Let temp = Data1.Recordset.Fields("0TrackPicture")
    Set MSFlexGrid1.CellPicture = LoadPicture(temp)
    End If
    
    Let MSFlexGrid1.Row = 1
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("1TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("1TrackPicture")
    Let MSFlexGrid1.Row = 2
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("2TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("2TrackPicture")
    Let MSFlexGrid1.Row = 3
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("3TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("3TrackPicture")
    Let MSFlexGrid1.Row = 4
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("4TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("4TrackPicture")
    Let MSFlexGrid1.Row = 5
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("5TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("5TrackPicture")
    Let MSFlexGrid1.Row = 6
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("6TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("6TrackPicture")
    Let MSFlexGrid1.Row = 7
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("7TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("7TrackPicture")
    Let MSFlexGrid1.Row = 8
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("8TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("8TrackPicture")
    Let MSFlexGrid1.Row = 9
    Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    Let MSFlexGrid1.Text = Data1.Recordset.Fields("9TrackComments")
    Set MSFlexGrid1.CellPicture = Data1.Recordset.Fields("9TrackPicture")
Next t

End Sub

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertiCally.

' =========================================================================================================================
' Close the Programming Diesel Window
'
   
' My Programming Notes
'
' Now that the commubication port is opened (hopefully to the correct communicaton port connected to the North Coast
' Enginerring Power HOuse Pro) the additional options on the menu bar can be enabled or disabled.
    
    'Let MainScreen.MenuCommunicationPortOpen.Enabled = False
    'Let MainScreen.MenuCommunicationPortClosed.Enabled = False
    'Let MainScreen.MenuCommunicationPortSetting.Enabled = False
    'Let MainScreen.MenuScaledTimeSetting.Enabled = False
    'Let MainScreen.MenuProgrammingModeDiesel.Enabled = False
    'Let MainScreen.MenuProgrammingModeSteam.Enabled = False
    'Let MainScreen.MenuProgrammingModeRollingStock.Enabled = False
    'Let MainScreen.MenuProgrammingModeSteam.Enabled = False
    'Let MainScreen.MenuProgrammingModeRollingStock.Enabled = False
    'Let MainScreen.MenuProgrammingModeOther.Enabled = False
    'Let MainScreen.MenuMainlineDieselProgramming.Enabled = False
    'Let MainScreen.MenuMainlineSteamProgramming.Enabled = False
    'Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = False
    'Let MainScreen.MenuMainlineOtherProgramming.Enabled = False
    'Let MainScreen.menumainlineconsist.Enabled = False
    'Let MainScreen.MenuMainlineOperation.Enabled = False
    'Let MainScreen.MenuMainlineOperationGUI.Enabled = False
    'Let MainScreen.MenuMainlineMacroMaker.Enabled = False
    Let MainScreen.MenuLayoutDefineBlocks.Enabled = False
    Let MainScreen.MenuLayoutCTC.Enabled = True


For t = 0 To 99
    MSFlexGrid1.ColWidth(t) = 1270
Next t

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
'
' Every time a new window is opened in Autoamtic Train Control we check the screen size and compare it to the window screen size. If the window cannot be displayed in the current screen size a
' message box is displayed. This allows time for the user to change the screen attributes to correct size.

Do
    MainScreen!MonitorResolution.GetMonitorInfo = True
    If Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY Then
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window Called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    End If
Loop While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY
    
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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


Private Sub Grid1_RowColChange()

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub SubWizard1_GotFocus()

End Sub


Private Sub Option1_Click()

End Sub


Private Sub OptionButtonTrackCurvedUR_Click()

End Sub


Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    
If ComboBoxTrack.Text = "None" Then
    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackNone.bmp")
Else
    If ComboBoxTrack.Text = "Straight" Then
        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackStraight.bmp")
    Else
        If ComboBoxTrack.Text = "Diagnal (top left, bottom right)" Then
            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal0.bmp")
        Else
            If ComboBoxTrack.Text = "Diagnal (top right, bottom left)" Then
                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal1.bmp")
            Else
                If ComboBoxTrack.Text = "Turnout (eastbound, right)" Then
                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutER.bmp")
                Else
                    If ComboBoxTrack.Text = "Turnout (eastbound, left)" Then
                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutEL.bmp")
                    Else
                        If ComboBoxTrack.Text = "Turnout (westbound, right)" Then
                            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWR.bmp")
                        Else
                            If ComboBoxTrack.Text = "Turnout (westbound, left)" Then
                                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWL.bmp")
                            Else
                                If ComboBoxTrack.Text = "Curve (eastbound, right)" Then
                                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveER.bmp")
                                Else
                                    If ComboBoxTrack.Text = "Curve (eastbound, left)" Then
                                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveEL.bmp")
                                    Else
                                        If ComboBoxTrack.Text = "Curve (westbound, right)" Then
                                            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWR.bmp")
                                        Else
                                            If ComboBoxTrack.Text = "Curve (westbound, left)" Then
                                                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWL.bmp")
                                            Else
                                                If ComboBoxTrack.Text = "Double Slip (top left, botton right)" Then
                                                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip0.bmp")
                                                Else
                                                    If ComboBoxTrack.Text = "Double Slip (top right, bottom left)" Then
                                                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip1.bmp")
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

If TextBoxTrackComment.Text <> "No Information Available" Then
    Let MSFlexGrid1.CellAlignment = 0
    Let MSFlexGrid1.Text = TextBoxTrackComment.Text
End If

End Sub


Private Sub MSFlexGrid1_OLECompleteDrag(Effect As Long)

Let MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
    
If ComboBoxTrack.Text = "None" Then
    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackNone.bmp")
Else
    If ComboBoxTrack.Text = "Straight" Then
        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackStraight.bmp")
    Else
        If ComboBoxTrack.Text = "Diagnal (top left, bottom right)" Then
            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal0.bmp")
        Else
            If ComboBoxTrack.Text = "Diagnal (top right, bottom left)" Then
                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDiagnal1.bmp")
            Else
                If ComboBoxTrack.Text = "Turnout (eastbound, right)" Then
                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutER.bmp")
                Else
                    If ComboBoxTrack.Text = "Turnout (eastbound, left)" Then
                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutEL.bmp")
                    Else
                        If ComboBoxTrack.Text = "Turnout (westbound, right)" Then
                            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWR.bmp")
                        Else
                            If ComboBoxTrack.Text = "Turnout (westbound, left)" Then
                                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackTurnoutWL.bmp")
                            Else
                                If ComboBoxTrack.Text = "Curve (eastbound, right)" Then
                                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveER.bmp")
                                Else
                                    If ComboBoxTrack.Text = "Curve (eastbound, left)" Then
                                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveEL.bmp")
                                    Else
                                        If ComboBoxTrack.Text = "Curve (westbound, right)" Then
                                            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWR.bmp")
                                        Else
                                            If ComboBoxTrack.Text = "Curve (westbound, left)" Then
                                                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackCurveWL.bmp")
                                            Else
                                                If ComboBoxTrack.Text = "Double Slip (top left, botton right)" Then
                                                    Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip0.bmp")
                                                Else
                                                    If ComboBoxTrack.Text = "Double Slip (top right, bottom left)" Then
                                                        Set MSFlexGrid1.CellPicture = LoadPicture(App.Path$ + "\Graphics\TrackDoubleSlip1.bmp")
                                                    Else
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

If TextBoxTrackComment.Text <> "No Information Available" Then
    Let MSFlexGrid1.CellAlignment = 0
    Let MSFlexGrid1.Text = TextBoxTrackComment.Text
End If

End Sub


