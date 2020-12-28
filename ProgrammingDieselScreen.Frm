VERSION 4.00
Begin VB.Form ProgrammingDieselScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Train Control - Diesel Programming Database"
   ClientHeight    =   8685
   ClientLeft      =   8160
   ClientTop       =   2310
   ClientWidth     =   6795
   Height          =   9090
   Icon            =   "ProgrammingDieselScreen.frx":0000
   Left            =   8100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Top             =   1965
   Width           =   6915
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   8280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "ProgrammingDieselScreen.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      Top             =   180
      Width           =   480
   End
   Begin VB.CheckBox LocomotiveDatabaseActiveSearch 
      Caption         =   "Locomotive Active List Search"
      DataField       =   "LocomotiveActiveList"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   255
      Left            =   3720
      TabIndex        =   39
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data LocomotiveDatabaseSearch 
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBase"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox LocomotiveDatabaseNumberSearch 
      DataField       =   "LocomotiveNumber"
      DataSource      =   "LocomotiveDatabaseSearch"
      Height          =   285
      Left            =   5040
      TabIndex        =   38
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data LocomotiveNumber 
      Caption         =   "Locomotive Number"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LocomotiveDataBase"
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox LocomotiveImageFilename 
      DataField       =   "LocomotiveImage"
      DataSource      =   "LocomotiveNumber"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   36
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton ButtonChangePicture 
      Caption         =   "Change Picture"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   35
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonDecoderProg 
      Caption         =   "&Decoder Prog"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame LocomotiveType 
      Caption         =   "Locomotive Type"
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   3015
      Begin VB.ComboBox LocomotiveModelManufacturer 
         DataField       =   "LocomotiveModelManufacturer"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "Locomotive Model Manufacturer"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox LocomotiveFacts 
         DataField       =   "LocomotiveFacts"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   2655
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox LocomotiveModel 
         DataField       =   "LocomotiveModel"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Text            =   "Locomotive Model"
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox LocomotiveManufacturer 
         DataField       =   "LocomotiveManufacturer"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Locomotive Manufacturer"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   8310
      Width           =   1215
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame LocoID 
      Caption         =   "Locomotive Identification"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.ComboBox LocomotiveNumberGO 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Text            =   "0000"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TemporaryTextBox 
         Alignment       =   2  'Center
         DataField       =   "LocomotiveNumber"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Text            =   "0000"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton ButtonGo 
         Caption         =   "Go"
         Height          =   290
         Left            =   2040
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox LocomotiveDecoderEquiped 
         Caption         =   "Decoder Equiped"
         DataField       =   "LocomotiveDecoderEquiped"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox LocomotiveActiveList 
         Caption         =   "Locomotive Active in List"
         DataField       =   "LocomotiveActiveList"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox LocomotivePrototype 
         DataField       =   "LocomotivePrototype"
         DataSource      =   "LocomotiveNumber"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Proto Type Owner"
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.PictureBox SSTab1 
      Enabled         =   0   'False
      Height          =   3735
      Left            =   3480
      ScaleHeight     =   3675
      ScaleWidth      =   3075
      TabIndex        =   12
      Top             =   4440
      Width           =   3135
      Begin VB.TextBox LocomotiveModelDrawBarPull 
         DataField       =   "LocomotiveModelDrawBarPull"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74760
         TabIndex        =   31
         Text            =   "Draw Bar Pull"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelTractionEffort 
         DataField       =   "LocomotiveModelTractionEffort"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74760
         TabIndex        =   29
         Text            =   "Traction Effort"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelAdhesionFactor 
         DataField       =   "LocomotiveModelAdhesionFactor"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74760
         TabIndex        =   27
         Text            =   "Adhesion Factor"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelWeight 
         DataField       =   "LocomotiveModelWeight"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74760
         TabIndex        =   25
         Text            =   "Weight"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox LocomotiveModelHorsePower 
         DataField       =   "LocomotiveModelHorsePower"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   -74760
         TabIndex        =   23
         Text            =   "Horse Power"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeDrawBarPull 
         DataField       =   "LocomotivePrototypeDrawBarPull"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "Draw Bar Pull"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeTractionEffort 
         DataField       =   "LocomotivePrototypeTractionEffort"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "Traction Effort"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeAdhesionFactor 
         DataField       =   "LocomotivePrototypeAdhesionFactor"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "Adhesion Factor"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeWeight 
         DataField       =   "LocomotivePrototypeWeight"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "Weight"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox LocomotivePrototypeHorsePower 
         DataField       =   "LocomotivePrototypeHorsePower"
         DataSource      =   "LocomotiveNumber"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "Horse Power"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Dar Bar (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   32
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Traction (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Adhesion %"
         Height          =   195
         Left            =   -73320
         TabIndex        =   28
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Weight (oz)"
         Height          =   255
         Left            =   -73320
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Horse Power"
         Height          =   255
         Left            =   -73320
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Dar Bar (tons)"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Traction (tons)"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Adhesion %"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Weight (tons)"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horse Power"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   120
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog PictureGet 
      Left            =   3480
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Automaric Train Control - Picture"
      InitDir         =   "c:\Automatic Train Control\Graphics"
   End
   Begin VB.Label Label11 
      Caption         =   $"ProgrammingDieselScreen.frx":0884
      Height          =   615
      Left            =   840
      TabIndex        =   40
      Top             =   120
      Width           =   5775
   End
   Begin VB.Image LocomotiveImage 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3255
   End
End
Attribute VB_Name = "ProgrammingDieselScreen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub ButtonChangePicture_Click()

' =========================================================================================================================
' Loading the Picture File Name Screen
'
' Using the COmmon Dialog Box
'
' The CommonDialog control provides a standard set of dialog boxes for operations such as opening, saving, and printing
' files or selecting colors and fonts.
'
' The syntax is as follows:
'
'   CommonDialog
'
' The CommonDialog control provides an interface between Visual Basic and the routines in the Microsoft Windows
' dynamic-link library COMMDLG.DLL.  To create a dialog box using this control, COMMDLG.DLL must be in your Microsoft
' Windows SYSTEM directory.
' You create dialog boxes for your application by adding the CommonDialog control to a form and setting its properties.
' The type of dialog box displayed is determined by the methods of the control.  At run time, a dialog box is displayed
' when the appropriate method is invoked; at design time, the CommonDialog control is displayed as an icon on a form.
' This icon can't be sized. Under Microsoft Windows 95, the CommonDialog control automatically provides context sensitive
' help on the interface of the dialog boxes by clicking:
'
' The What 's This help button in the title bar then clicking the item for which you want more information. The right
' mouse button over the item for which you want more information then selecting the What's This command in the displayed
' context menu.
'
' The operating system provides the text shown in the Windows 95 Help popup.  However, no topic exists for the Help button
' displayed with the CommonDialog control by setting the Flags property. You can't specify where a dialog box is displayed.
'
' The CommonDialog control is a custom control, which is a separate file with an .OCX extension. To use the CommonDialog
' control in your application, you must add the COMDLG16.OCX or COMDLG32.OCX file to the project. To automatically include
' this custom control in new projects, add the file you need to AUTOLOAD.VBP.  When distributing your application, install
' the .OCX file in the user's Microsoft Windows SYSTEM directory.  For more information on how to add a custom control
' to a project, see the Programmer's Guide.

' =========================================================================================================================
' Show Open Common Dialog Control
'
' Displays the CommonDialog control's Open dialog box.
'
' Syntax
'
'   object.ShowOpen
'
' The object placeholder represents an object expression that evaluates to an object in the Applies To list.

    PictureGet.ShowOpen


' =========================================================================================================================
' FIle name Property
'
' Returns or sets the path and filename of a selected file.  Not available at design time for the FileListBox control
' and ProjectTemplate object.
'
' The syntax is as follows:
'
'   object.filename [=  pathname]
'
' The FileName property syntax has these parts: where 'object' is an object expression that evaluates to an object in
' the Applies To list; and where 'pathname' is a string expression that specifies the path and filename.
'
' When you create the control at run time, the FileName property is set to a zero-length string (""), meaning no file
' is currently selected. In the CommonDialog control, you can set the FileName property before opening a dialog box to
' set an initial filename. Reading this property returns the currently selected filename from the list.  The path is
' retrieved separately, using the Path property.  The value is functionally equivalent to List(ListIndex).  If no file
' is selected, FileName returns a zero-length string.
'
' When setting this property: Including a drive, path, or pattern in the string changes the settings of the Drive, Path,
' and Pattern properties accordingly. Including the name of an existing file (without wildcard characters) in the string
' selects the file. Changing the value of this property may also cause one or more of these events: PathChange (if you
' change the path), PatternChange (if you change the pattern), or DblClick (if you assign an existing filename).


    Let LocomotiveImageFilename.Text = PictureGet.filename

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub


Private Sub ButtonClose_Click()

' =========================================================================================================================
' Close the Programming Diesel Window
'
   
' My Programming Notes
'
' Now that the commubication port is opened (hopefully to the correct communicaton port connected to the North Coast
' Enginerring Power HOuse Pro) the additional options on the menu bar can be enabled or disabled.
    
    Let MainScreen.MenuCommunicationPortOpen.Enabled = False
    Let MainScreen.MenuCommunicationPortClosed.Enabled = True
    Let MainScreen.MenuCommunicationPortSetting.Enabled = False
    Let MainScreen.MenuScaledTimeSetting.Enabled = True
    Let MainScreen.MenuProgrammingModeDiesel.Enabled = True
    Let MainScreen.MenuProgrammingModeSteam.Enabled = True
    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = True
    Let MainScreen.MenuProgrammingModeOther.Enabled = True
    Let MainScreen.MenuMainlineDieselProgramming.Enabled = True
    Let MainScreen.MenuMainlineSteamProgramming.Enabled = True
    Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = True
    Let MainScreen.MenuMainlineOtherProgramming.Enabled = True
    Let MainScreen.menumainlineconsist.Enabled = True
    Let MainScreen.menumainlinescalespeedoperation.Enabled = True
    Let MainScreen.MenuMainlineOperation.Enabled = True
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

    ProgrammingDieselScreen.Hide
    
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

    Unload ProgrammingDieselScreen

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

Private Sub ButtonDecoderProg_Click()

' =========================================================================================================================
' Loading the Programming Decoder Window
'

' =========================================================================================================================
' Load Statement
'
' Loads a form or control into memory.
'
' Syntax is as the folowing:
'
'   Load object
'
' The object placeholder is the name of a Form object, MDIForm object, or control array element to load. You don 't need
' to use the Load statement with forms unless you want to load a form without displaying it.  Any reference to a form
' (except in a Set or If...TypeOf statement) automatically loads it if it's not already loaded.  For example, the Show
' method loads a form before displaying it.  Once the form is loaded, its properties and controls can be altered by the
' application, whether or not the form is actually visible.  Under some circumstances, you may want to load all your
' forms during initialization and display them later as they're needed. When more than one instance of a form is loaded,
' each form has its own set of properties and variables that can be changed independently from all other instances of the
' form. When Visual Basic loads a Form object, it sets form properties to their initial values and then performs the
' Load event procedure.  When an application starts, Visual Basic automatically loads and displays the application's
' startup form. If you load a Form with the MDIChild property set to True before loading the MDIForm, the MDIForm is
' automatically loaded before the child form.  Loading an MDI child form causes it to be visible immediately after
' the Form_Load event procedure ends because MDI child forms can't be hidden. The standard dialog boxes produced by
' Visual Basic functions such as MsgBox and InputBox do not need to be loaded, shown, or unloaded, but can simply be
' invoked directly.

    Load ProgrammingDecoder

' =========================================================================================================================
' Show Method
'
'Displays an MDIForm or Form object.  Doesn't support named arguments. For Visual Basic in-process OLE server (DLL)
' considerations and restrictions that apply to this topic, see Make OLE DLL File Command.
'
' Syntax is as the following:
'
'   object.Show Style
'
' The Show method syntax has these parts where 'object' is optional. An object expression that evaluates to an object in
' the Applies To list. If object is omitted, the form associated with the active form module is assumed to be object. The
' 'style' is optional. Integer that determines if the form is modal or modeless.  If style is 0, the form is modeless;
' if style is 1, the form is modal. If the specified form isn't loaded when the Show method is invoked, Visual Basic
' automatically loads it. When Show displays a modeless form, subsequent code is executed as it's
' encountered.  When Show displays a modal form, no subsequent code is executed until the form is hidden or
' unloaded. When Show displays a modal form, no input (keyboard or mouse click) can occur except to objects on the
' modal form.  The program must hide or unload a modal form (usually in response to some user action) before input
' to another form can occur.  An MDIForm can't be modal. Although other forms in your application are disabled when
' a modal form is displayed, other applications aren't. The startup form of an application is automatically shown
' after its Load event is invoked.

    ProgrammingDecoder.Show

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

Private Sub ButtonGo_Click()
    
' =========================================================================================================================
' Button Go Command
'
' The purpose of the GO command was to seach for a record in the data base equal to the locomotive number.
' Since there is nine thousand ninehundred ninety nine records to choose from that made clicking throu each one
' too slow for the user. Hence this speeds things up considerable.

' =========================================================================================================================
' My Programming Notes
'
' I needto make sure that the value of the locomotive to be searched is between 0 and 9999. If not then diplay an
' error message to the user and aborth the search.

    If Val(LocomotiveNumberGO.Text) < 1 Then
        MsgBox "The value entered for the Locomotive is not valid. Please" + Chr$(13) + "use values between one and nine thousand nine hundred and ninety-nine only.", vbExclamation, "Locomotive Entry Error"
        Let LocomotiveNumberGO.Text = "0"
    End If
    
    If Val(LocomotiveNumberGO.Text) > 9999 Then
        MsgBox "The value entered for the Locomotive is not valid. Please" + Chr$(13) + "use values between one and nine thousand nine hundred and ninety-nine only.", vbExclamation, "Locomotive Entry Error"
        Let LocomotiveNumberGO.Text = "0"
    End If
    
' =========================================================================================================================
' My Programming Notes
'
' Now that I know the user has enterd a valid locomotiv number, I can search for it in the data base and dispaly the
' information in the record.
    
    If Val(LocomotiveNumberGO.Text) <> 0 Then
    
' =========================================================================================================================
' Mouse Pointer Property
'
' Returns or sets a value indicating the type of mouse pointer displayed when the mouse is over a particular part of an
' object at run time.
'
' THe syntax is as follows:
'
'   object.MousePointer [= value]
'
'The MousePointer property syntax has these parts: where 'object'  is an object expression that evaluates to an object
' in the Applies To list; and 'value' is an integer specifying the type of mouse pointer displayed, as described in
' Settings. Now turn the mouse pointer to an hour glass so they know the program is searching for the locomotive.
'
' My Programming Notes
'
' I just wanted some way to show the user of the program that the search in the data base has been started; and to
' please wait unitl the results are displayed.

       Let MousePointer = vbHourglass
    
' =========================================================================================================================
' Searching for
'
' My Programming Notes
'
' I'm Placing the search string in a temporary string for execution.
    
       Let temporary = "LocomotiveNumber like " + LocomotiveNumberGO.Text
    
' =========================================================================================================================
' Findfirst Method
'
'Locates the first, last, next, or previous record in a dynaset- or snapshot-type Recordset object that satisfies the
' specified criteria and makes that record the current record.  The Find methods can also be used with the outdated
' Dynaset and Snapshot objects.
'
' The syntax is as follows:
'
'   recordset.{FindFirst | FindLast | FindNext | FindPrevious} criteria

' The Find methods have these parts: where 'recordset' is the name of an existing dynaset- or snapshot-type Recordset,
' Dynaset, or Snapshot object; and 'criteria' is a string expression (like the WHERE clause in an SQL statement without
' the word WHERE) used to locate the record.
'
' If you want to include all the records in your searchnot just those that meet a specific conditionuse the Move methods
' to move from record to record.  To locate a record in a table-type Recordset, use the Seek method. If a record matching
' the criteria isn't located, the current record pointer is positioned to the first record in the Recordset, and the
' NoMatch property is set to True.  If recordset contains more than one record that satisfies the criteria, FindFirst
' locates the first occurrence, FindNext locates the next occurrence, and so on. If you edit the current record, be sure
' you save the changes using the Update method before you move to another record.  If you move to another record
' without updating, your changes are lost without warning.
'
' Always check the value of the NoMatch property to determine whether the Find operation has succeeded.  If the search
' succeeds, NoMatch is set to False.  If it fails, NoMatch is set to True and the current record is set to the first
' record in the Recordset.
    
      LocomotiveNumber.Recordset.FindFirst temporary

' =========================================================================================================================
' Mouse Pointer Property
'
' Returns or sets a value indicating the type of mouse pointer displayed when the mouse is over a particular part of an
' object at run time.
'
' The syntax is as follows:
'
'   object.MousePointer [= value]
'
'The MousePointer property syntax has these parts: where 'object'  is an object expression that evaluates to an object
' in the Applies To list; and 'value' is an integer specifying the type of mouse pointer displayed, as described in
' Settings. Now turn the mouse pointer to an hour glass so they know the program is searching for the locomotive.
'
' My Programming Notes
'
' Return the mouse pointer to the default value

     Let MousePointer = vbDefault

' =========================================================================================================================
' Enabled Property
'
' Returns or sets a value that determines whether a form or control can respond to user-generated events.
'
' The syntax is as the following:
'
'   object.Enabled [= boolean]
'
' The Enabled property syntax has these parts: where 'object' is an object expression that evaluates to an object in the
' Applies To list.  If object is omitted, the form associated with the active form module is assumed to be object; and
' 'boolean' is a Boolean expression that specifies whether object can respond to user-generated events.
'
' The Enabled property allows forms and controls to be enabled or disabled at run time.  For example, you can disable
' objects that don't apply to the current state of the application.  You can also disable a control used purely for
' display purposes, such as a text box that provides read-only information.
' Disabling a Timer control by setting Enabled to False cancels the countdown set up by the control's Interval
' property.
' For a Menu control, Enabled is normally read/write at run time.  But Enabled is read-only for menu items that are
' exposed or supplied by Visual Basic to add-ins, such as the Add-In Manager command on the Add-Ins menu.
'
' My Programming Notes
'
' Set the command buttton 'OK' to enable updating of the record.

        Let LocomotivePrototype.Enabled = True
        Let LocomotiveActiveList.Enabled = True
        Let LocomotiveDecoderEquiped.Enabled = True
        Let LocomotiveManufacturer.Enabled = True
        Let LocomotiveModel.Enabled = True
        Let LocomotiveModelManufacturer.Enabled = True
        Let LocomotiveFacts.Enabled = True
        Let LocomotiveImage.Enabled = True
        Let LocomotiveImageFilename.Enabled = True
        Let SSTab1.Enabled = True
        Let LocomotivePrototypeHorsePower.Enabled = True
        Let LocomotivePrototypeWeight.Enabled = True
        Let LocomotivePrototypeAdhesionFactor.Enabled = True
        Let LocomotivePrototypeTractionEffort.Enabled = True
        Let LocomotivePrototypeDrawBarPull.Enabled = True
        Let LocomotiveModelHorsePower.Enabled = True
        Let LocomotiveModelWeight.Enabled = True
        Let LocomotiveModelAdhesionFactor.Enabled = True
        Let LocomotiveModelTractionEffort.Enabled = True
        Let LocomotiveModelDrawBarPull.Enabled = True
        Let ButtonDecoderProg.Visible = True
        Let ButtonUpdate.Enabled = True
        Let ButtonChangePicture.Enabled = True
    End If

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub



Private Sub ButtonUpdate_Click()
        
    LocomotiveNumber.Recordset.Edit
    LocomotiveNumber.Recordset.Update

' =========================================================================================================================
' End Sub Statement
'
'Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub






Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.

Let LocomotiveNumber.DatabaseName = App.Path + "\Databases\LocomotiveDatabaseDiesels.mdb"
Let LocomotiveDatabaseSearch.DatabaseName = App.Path + "\Databases\LocomotiveDatabaseDiesels.mdb"

' =========================================================================================================================
' Opening the Programming Diesel Window
'
   
' My Programming Notes
'
' Now that the commubication port is opened (hopefully to the correct communicaton port connected to the North Coast
' Enginerring Power HOuse Pro) the additional options on the menu bar can be enabled or disabled.
    
    Let MainScreen.MenuCommunicationPortOpen.Enabled = False
    Let MainScreen.MenuCommunicationPortClosed.Enabled = False
    Let MainScreen.MenuCommunicationPortSetting.Enabled = False
    Let MainScreen.MenuScaledTimeSetting.Enabled = False
    Let MainScreen.MenuProgrammingModeDiesel.Enabled = True
    Let MainScreen.MenuProgrammingModeSteam.Enabled = False
    Let MainScreen.MenuProgrammingModeRollingStock.Enabled = False
    Let MainScreen.MenuProgrammingModeOther.Enabled = False
    Let MainScreen.MenuMainlineDieselProgramming.Enabled = False
    Let MainScreen.MenuMainlineSteamProgramming.Enabled = False
    Let MainScreen.MenuMainlineRollingStockProgramming.Enabled = False
    Let MainScreen.MenuMainlineOtherProgramming.Enabled = False
    Let MainScreen.menumainlineconsist.Enabled = False
    Let MainScreen.MenuMainlineOperation.Enabled = False
    Let MainScreen.MenuLayoutDefineBlocks.Enabled = False
    Let MainScreen.MenuLayoutCTC.Enabled = False
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub


Private Sub LocomotiveActiveList_Click()

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' THe syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub


Private Sub LocomotiveDecoderEquiped_Click()

' =========================================================================================================================
' My Programming Notes
'
' When the user clicks on the check box to indicate that the locomotive has a decoder installed, this code changed
' the 'enabled' property of the button for decoder programming; thus, making it accessable to the user.
    
    If LocomotiveDecoderEquiped.Value = vbUnchecked Then Let ButtonDecoderProg.Enabled = False
    If LocomotiveDecoderEquiped.Value = vbChecked Then Let ButtonDecoderProg.Enabled = True
        
' =========================================================================================================================
' End Sub
'
' Ends a procedure or block.
'
' The syntax is as follows:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level variables
' and all static local variables in all modules.  If you need to preserve the value of these variables, use the Stop
' Statement instead.  You can then resume execution while preserving the value of those variables.

    End Sub


Private Sub LocomotiveFacts_Change()

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub LocomotiveImageFilename_Change()

' What if there is No Picture?
'
' If you don't use an On Error statement, any run-time error that occurs is fatal; that is, an error message is displayed
' and execution stops.An "enabled" error handler is one that has been turned on by an On Error statement; an "active"
' error handler is an enabled handler that is in the process of handling an error.  If an error occurs while an error
' handler is active (between the occurrence of the error and a Resume, Exit Sub, Exit Function, or Exit Property
' statement), the current procedure's error handler can't handle the error.  Control returns to the calling procedure;
' if the calling procedure has an enabled error handler, it is activated to handle the error.  If the calling
' procedure's error handler is also active, control passes back through previous calling procedures until an enabled,
' but inactive, error handler is found.  If no inactive, enabled error handler is found, the error is fatal at the
' point at which it actually occurred.  Each time the error handler passes control back to the calling procedure,
' that procedure becomes the current procedure.  Once an error is handled by an error handler in any procedure, ex
' current procedure at the point designated by the Resume statement.
'
' Note   An error-handling routine is not a Sub or Function procedure.  It is a section of code marked by a line label
' or line number.
'
' Error-handling routines rely on the value in the Err object's Number property to determine the cause of the error.
' The error-handling routine should test or save relevant Err object property  values before any other error can occur
' or before a procedure that could cause an error is called.  The values in the Err object's properties reflect only the
' most recent error.  The error message associated with Err.Number is contained in Err.Description.
' On Error Resume Next causes execution to continue with the statement immediately following the statement that caused
' the run-time error, or with the statement immediately following the most recent call out of the procedure containing
' the On Error Resume Next statement. This allows execution to continue despite a run-time error.  You can then build
' the error-handling routine inline within the procedure, rather than transfer control to another location within the
' procedure.  An On Error Resume Next statement becomes inactive when another procedure is called, so you should execute
' an On Error Resume Next statement in each called routine if you want inline error handling within that routine.
'
' Note   The On Error Resume Next construct may be preferable to On Error GoTo when dealing with errors generated
' during access to other objects, since it permits unambiguous identification of the object whose error code is being
' returned.  Checking Err after each interaction with an object removes ambiguity about which object your code was
' accessing when the error occurred because the context is immediate.  Thus, you can be sure of which object placed the
' error code in Err.Number, as well as which object originally generated the error (the one specified in Err.Source).

    On Error Resume Next

    Let LocomotiveImage.Picture = LoadPicture(LocomotiveImageFilename.Text)


    If Err.Number = 53 Then

'Displays a message in a dialog box, waits for the user to choose a button, and returns a value indicating which button
' the user has chosen.

    MsgBox "Your picture listed on file was not found." + Chr$(13) + "Please update this record.", vbExclamation, "Locomotive Picture not Found"

    End If

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub LocomotiveManufacturer_Click()

Let LocomotiveModel.Text = LocomotiveManufacturer.Text

If LocomotiveManufacturer.Text = "American Locomotive Company" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "C-636"
    LocomotiveModel.AddItem "C-630"
    LocomotiveModel.AddItem "C-628"
    LocomotiveModel.AddItem "C-430"
    LocomotiveModel.AddItem "C-425"
    LocomotiveModel.AddItem "C-424"
    LocomotiveModel.AddItem "C-420"
    LocomotiveModel.AddItem "RS36"
    LocomotiveModel.AddItem "RS32"
    LocomotiveModel.AddItem "RS27"
    LocomotiveModel.AddItem "RSD15"
    LocomotiveModel.AddItem "RSD12"
    LocomotiveModel.AddItem "RD11"
    LocomotiveModel.AddItem "RS3"
    LocomotiveModel.AddItem "RS2"
    LocomotiveModel.AddItem "RS1"
    LocomotiveModel.AddItem "S-1"
    LocomotiveModel.AddItem "S-2"
    LocomotiveModel.AddItem "S-3"
    LocomotiveModel.AddItem "S-4"
    LocomotiveModel.AddItem "S-6"
    LocomotiveModel.AddItem "T-6"
    LocomotiveModel.AddItem "PA-3"
    LocomotiveModel.AddItem "PA-2"
    LocomotiveModel.AddItem "PA-1"
    LocomotiveModel.AddItem "FA-2"
    LocomotiveModel.AddItem "FA-1"
End If

If LocomotiveManufacturer.Text = "Baldwin" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "VO 1000"
    LocomotiveModel.AddItem "DRS-6-6-15 or AS-616"
    LocomotiveModel.AddItem "DRS-6-4-15 or AS-416"
    LocomotiveModel.AddItem "DRS-4-4-15 or AS-16"
    LocomotiveModel.AddItem "DS-4-4-10"
    LocomotiveModel.AddItem "S-12"
End If

If LocomotiveManufacturer.Text = "Bombardier" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "HR616"
    End If
    
If LocomotiveManufacturer.Text = "FairBanks-Morse" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "H-12-44"
End If

If LocomotiveManufacturer.Text = "General Electric" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "AC6000"
    LocomotiveModel.AddItem "AC4400CW"
    LocomotiveModel.AddItem "9-44CW"
    LocomotiveModel.AddItem "9-44C"
    LocomotiveModel.AddItem "9-44BW"
    LocomotiveModel.AddItem "9-44B"
    LocomotiveModel.AddItem "9-40CW"
    LocomotiveModel.AddItem "9-40C"
    LocomotiveModel.AddItem "9-40BW"
    LocomotiveModel.AddItem "9-40B"
    LocomotiveModel.AddItem "8-41CW"
    LocomotiveModel.AddItem "8-41C"
    LocomotiveModel.AddItem "8-40CW"
    LocomotiveModel.AddItem "8-40CM"
    LocomotiveModel.AddItem "8-40C"
    LocomotiveModel.AddItem "8-40BW"
    LocomotiveModel.AddItem "8-40B"
    LocomotiveModel.AddItem "8-39CE"
    LocomotiveModel.AddItem "8-39C"
    LocomotiveModel.AddItem "8-39B"
    LocomotiveModel.AddItem "8-32C"
    LocomotiveModel.AddItem "8-32BWH or P32-8BWH"
    LocomotiveModel.AddItem "8-32B"
    LocomotiveModel.AddItem "7-36B"
    LocomotiveModel.AddItem "7-36C"
    LocomotiveModel.AddItem "7-33C"
    LocomotiveModel.AddItem "7-30B-A1"
    LocomotiveModel.AddItem "7-30C-A"
    LocomotiveModel.AddItem "7-30B-A"
    LocomotiveModel.AddItem "7-30C"
    LocomotiveModel.AddItem "7-30B"
    LocomotiveModel.AddItem "7-23BQ"
    LocomotiveModel.AddItem "7-23B"
    LocomotiveModel.AddItem "U36C"
    LocomotiveModel.AddItem "U36B"
    LocomotiveModel.AddItem "U33C"
    LocomotiveModel.AddItem "U33B"
    LocomotiveModel.AddItem "SF30C"
    LocomotiveModel.AddItem "U30C"
    LocomotiveModel.AddItem "U30B"
    LocomotiveModel.AddItem "U28B"
    LocomotiveModel.AddItem "U25C"
    LocomotiveModel.AddItem "U25B"
    LocomotiveModel.AddItem "U23C"
    LocomotiveModel.AddItem "U23B"
    LocomotiveModel.AddItem "U18B"
End If

If LocomotiveManufacturer.Text = "General Motors - Electromotive Division" Then
    LocomotiveModel.Clear
    LocomotiveModel.AddItem "SD90/43MAC"
    LocomotiveModel.AddItem "SD80MAC"
    LocomotiveModel.AddItem "SD80MC"
    LocomotiveModel.AddItem "SD75M"
    LocomotiveModel.AddItem "SD70MAC"
    LocomotiveModel.AddItem "SD70M"
    LocomotiveModel.AddItem "SD70I"
    LocomotiveModel.AddItem "SD70"
    LocomotiveModel.AddItem "F69PH-AC"
    LocomotiveModel.AddItem "SD60MAC"
    LocomotiveModel.AddItem "SD60M"
    LocomotiveModel.AddItem "SD60F"
    LocomotiveModel.AddItem "SD60I"
    LocomotiveModel.AddItem "SD60"
    LocomotiveModel.AddItem "SD50M"
    LocomotiveModel.AddItem "SD50F"
    LocomotiveModel.AddItem "SD50I"
    LocomotiveModel.AddItem "SD50"
    LocomotiveModel.AddItem "GP60M"
    LocomotiveModel.AddItem "GP60"
    LocomotiveModel.AddItem "AMD-103"
    LocomotiveModel.AddItem "F59PHI"
    LocomotiveModel.AddItem "F59PH"
    LocomotiveModel.AddItem "GP59"
    LocomotiveModel.AddItem "GP50"
    LocomotiveModel.AddItem "DD40AX"
    LocomotiveModel.AddItem "SD45-2T"
    LocomotiveModel.AddItem "SD45-2"
    LocomotiveModel.AddItem "SDP45"
    LocomotiveModel.AddItem "FP45"
    LocomotiveModel.AddItem "F45"
    LocomotiveModel.AddItem "SD45"
    LocomotiveModel.AddItem "SD40-2W"
    LocomotiveModel.AddItem "SD40-2T"
    LocomotiveModel.AddItem "SD40-2F"
    LocomotiveModel.AddItem "SD40-2"
    LocomotiveModel.AddItem "SDP40F"
    LocomotiveModel.AddItem "SDP40"
    LocomotiveModel.AddItem "F40PHM-2"
    LocomotiveModel.AddItem "F40PH-2C"
    LocomotiveModel.AddItem "F40PH"
    LocomotiveModel.AddItem "F40C"
    LocomotiveModel.AddItem "SD40"
    LocomotiveModel.AddItem "GP40-2"
    LocomotiveModel.AddItem "GP40X"
    LocomotiveModel.AddItem "GP40P"
    LocomotiveModel.AddItem "GP40W"
    LocomotiveModel.AddItem "GP40"
    LocomotiveModel.AddItem "GP39-2"
    LocomotiveModel.AddItem "SDL39"
    LocomotiveModel.AddItem "SD39"
    LocomotiveModel.AddItem "GP39"
    LocomotiveModel.AddItem "SD38-2"
    LocomotiveModel.AddItem "GP39-2W"
    LocomotiveModel.AddItem "GP38-2"
    LocomotiveModel.AddItem "SD38"
    LocomotiveModel.AddItem "GP38"
    LocomotiveModel.AddItem "SDP35"
    LocomotiveModel.AddItem "SD35"
    LocomotiveModel.AddItem "GP35"
    LocomotiveModel.AddItem "GP30"
    LocomotiveModel.AddItem "SD24"
    LocomotiveModel.AddItem "GP20"
    LocomotiveModel.AddItem "GP15-1"
    LocomotiveModel.AddItem "GP15T"
    LocomotiveModel.AddItem "GP15"
    LocomotiveModel.AddItem "SD18"
    LocomotiveModel.AddItem "GP18"
    LocomotiveModel.AddItem "SD9"
    LocomotiveModel.AddItem "GP9"
    LocomotiveModel.AddItem "SD7"
    LocomotiveModel.AddItem "GP7"
    LocomotiveModel.AddItem "CF7"
    LocomotiveModel.AddItem "GMD1"
    LocomotiveModel.AddItem "RS1325"
    LocomotiveModel.AddItem "NM5"
    LocomotiveModel.AddItem "NW2"
    LocomotiveModel.AddItem "TR5"
    LocomotiveModel.AddItem "TR4"
    LocomotiveModel.AddItem "MP15T"
    LocomotiveModel.AddItem "MP15AC"
    LocomotiveModel.AddItem "MP15(DC)"
    LocomotiveModel.AddItem "SW1504"
    LocomotiveModel.AddItem "SW1500"
    LocomotiveModel.AddItem "SW1200"
    LocomotiveModel.AddItem "SW1001"
    LocomotiveModel.AddItem "SW1000"
    LocomotiveModel.AddItem "SW900"
    LocomotiveModel.AddItem "SW600"
    LocomotiveModel.AddItem "SW9"
    LocomotiveModel.AddItem "SW8"
    LocomotiveModel.AddItem "SW7"
    LocomotiveModel.AddItem "SW1"
    LocomotiveModel.AddItem "FL9"
    LocomotiveModel.AddItem "FP9"
    LocomotiveModel.AddItem "F9"
    LocomotiveModel.AddItem "FP7"
    LocomotiveModel.AddItem "F7"
    LocomotiveModel.AddItem "F3"
    LocomotiveModel.AddItem "E9A"
    LocomotiveModel.AddItem "E8A"
    LocomotiveModel.AddItem "BL2"

End If

If LocomotiveManufacturer.Text = "Montreal Locomotive Works" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "S-13"
        LocomotiveModel.AddItem "M636"
        LocomotiveModel.AddItem "M630"
        LocomotiveModel.AddItem "M420"
        LocomotiveModel.AddItem "M420R"
End If

If LocomotiveManufacturer.Text = "Morrison Knudsen" Then
        LocomotiveModel.Clear
        LocomotiveModel.AddItem "MK5000C"
        LocomotiveModel.AddItem "MK1200G"
        LocomotiveModel.AddItem "MK-F40PHL-2"
        LocomotiveModel.AddItem "MKGP40FH-2"
End If

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' THe syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

    End Sub


Private Sub LocomotiveModel_Change()

' =========================================================================================================================
' End Sub Statement
'
'Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub LocomotiveModelManufacturer_Change()

' =========================================================================================================================
' End Sub Statement
'
'Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub LocomotiveNumberGO_Change()

' =========================================================================================================================
' Visible Property
'
'Returns or sets a value indicating whether an object is visible or hidden.
'
' The syntax is as follows:
'
'   object.Visible [= boolean]
'
' The Visible property syntax has these parts: where 'object' is an object expression that evaluates to an object in
' the Applies To list; and 'boolean' is a boolean expression specifying whether the object is visible or hidden. To hide
' an object at startup, set the Visible property to False at design time. Setting this property in code enables you to
' hide and later redisplay a control at run time in response to a particular event. Using the Show or Hide method
' on a form is the same as setting the form's Visible property in code to True or False, respectively.
'
' My Programming Notes
'
' Every time a new locomotive is selected I need to make the deocder button invisible. I don't waht the user going to
' the Decoder Programming screen if the locomotive does not support deocders.

    Let ButtonDecoderProg.Visible = False

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

    End Sub

Private Sub LocomotivePrototype_Change()

' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' THe syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub TemporaryTextBox_Change()

'Let LocomotiveDatabaseForLocomotive.LocomotiveNumberGO.Text = Format(LocomotiveNumber.Recordset("LocomotiveNumber"), "0000")
'Let LocomotiveDatabaseForLocomotive.LocomotiveNumber.Caption = Format(LocomotiveNumber.Recordset("LocomotiveNumber"), "0000")

End Sub


Private Sub Timer1_Timer()

    Let Timer1.Interval = 0

' Returns or sets a value indicating the type of mouse pointer displayed when the mouse is over a particular part of an object at run time.

    LocomotiveNumberGO.Clear

' Move to the first, last, next, or previous record in a specified Recordset object and make that record the current record.
' The Move methods can also be used with the outdated Dynaset, Snapshot, and Table objects.

    LocomotiveDatabaseSearch.Recordset.MoveFirst

' Use the RecordCount property to find out how many records in a Recordset or TableDef object have been accessed.
' RecordCount doesn't indicate how many records are contained in a dynaset- or snapshot-type Recordset until all records have been accessed.
' Once the last record has been accessed, the RecordCount property indicates the total number of undeleted records in the Recordset or TableDef.
' To force the last record to be accessed, use the MoveLast or FindLast method on the Recordset.
' You can also use an SQL Count function to determine the approximate number of records your query will return.

    Do While Not LocomotiveDatabaseSearch.Recordset.EOF

' The AbsolutePosition property enables you to position the current record pointer to a specific record based on its ordinal position in a dynaset- or snapshot-type Recordset.
' You can also determine the current record number by checking the AbsolutePosition property setting.
' The AbsolutePosition property value is zero-baseda setting of 0 refers to the first record in the Recordset.
' Setting a value greater than the number of populated records causes a trappable error.  You can determine the number of populated records in the Recordset by checking the RecordCount property setting.
' If there is no current record, as when there are no records in the Recordset, -1 is returned.
' If the current record is deleted, the AbsolutePosition property value isn't defined, and a trappable error occurs if it's referenced.
' New records are added to the end of the sequence.

    LocomotiveDatabaseSearch.Recordset.MoveNext
    If LocomotiveDatabaseActiveSearch.Value = vbChecked Then
         LocomotiveNumberGO.AddItem Format(Val(LocomotiveDatabaseNumberSearch.Text), "0000")
     End If

' DoEvents passes control to the operating system.
' Control is not returned until the operating system has finished processing the events in its queue and  all keys in the SendKeys queue have been sent.
' If parts of your code take up a lot of processor time, use DoEvents periodically to relinquish control to the operating system so that events, such as keyboard input and mouse clicks, can be processed without significant delay.
  
      Let temp = DoEvents
     
    ' Just in cause the user of the program closes the window while the database is still open searching
    ' This will terminate the search and allow the window to be close withou gerating and error.
         
Loop


' =========================================================================================================================
' End Sub Statement
'
' Ends a procedure or block.
'
' The syntax is a follwos:
'
'   End Sub
'
' Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions that
' apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.


End Sub


