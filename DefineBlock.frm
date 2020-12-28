VERSION 4.00
Begin VB.Form DefineBlocks 
   Caption         =   "Automatic Train Control - Defining Layout"
   ClientHeight    =   8040
   ClientLeft      =   1125
   ClientTop       =   2985
   ClientWidth     =   6600
   Height          =   8445
   Left            =   1065
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   6600
   Top             =   2640
   Width           =   6720
   Begin VB.TextBox TextBoxSignalDoubleCounter 
      Height          =   285
      Left            =   8280
      TabIndex        =   40
      Text            =   "0"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox TextBoxSignalDoublePositionTop 
      Height          =   285
      Left            =   8160
      TabIndex        =   39
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TextBoxSignalDoublePositionLeft 
      Height          =   285
      Left            =   8160
      TabIndex        =   38
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   7560
      Top             =   720
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\TrackPlanDatabase.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TrackPlan"
      Top             =   240
      Width           =   1140
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4Counter 
      Height          =   285
      Left            =   7200
      TabIndex        =   22
      Text            =   "0"
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3Counter 
      Height          =   285
      Left            =   7200
      TabIndex        =   21
      Text            =   "0"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2Counter 
      Height          =   285
      Left            =   7200
      TabIndex        =   20
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1Counter 
      Height          =   285
      Left            =   7200
      TabIndex        =   19
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox TextBoxTrackIconStraightCounter 
      Height          =   285
      Left            =   7200
      TabIndex        =   18
      Text            =   "0"
      Top             =   5040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   720
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4PositionTop 
      Height          =   285
      Left            =   7080
      TabIndex        =   17
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch4PositionLeft 
      Height          =   285
      Left            =   7080
      TabIndex        =   16
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3PositionTop 
      Height          =   285
      Left            =   7080
      TabIndex        =   15
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch3PositionLeft 
      Height          =   285
      Left            =   7080
      TabIndex        =   14
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2PositionTop 
      Height          =   285
      Left            =   7080
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch2PositionLeft 
      Height          =   285
      Left            =   7080
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1PositionTop 
      Height          =   285
      Left            =   7080
      TabIndex        =   11
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconSwitch1PositionLeft 
      Height          =   285
      Left            =   7080
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconStraightPositionTop 
      Height          =   285
      Left            =   7080
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TextBoxTrackIconStraightPositionLeft 
      Height          =   285
      Left            =   7080
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton ButtonUpdate 
      Caption         =   "&Update"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox PictureBoxIcon 
      ClipControls    =   0   'False
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   1320
      Width           =   6375
   End
   Begin TabDlg.SSTab TabTrackIcon 
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      _Version        =   393216
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
      TabPicture(0)   =   "DefineBlock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PictureBoxTrackIconSwitch4(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PictureBoxTrackIconSwitch3(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PictureBoxTrackIconSwitch2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PictureBoxTrackIconSwitch1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "PictureBoxTrackIconStraight(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Signal Icons"
      TabPicture(1)   =   "DefineBlock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PictureBoxSignalSingle(0)"
      Tab(1).Control(1)=   "PictureBoxSignalDouble(0)"
      Tab(1).Control(2)=   "PictureBoxSignalTriple(0)"
      Tab(1).Control(3)=   "PictureBoxSignalDoubleDouble(0)"
      Tab(1).Control(4)=   "PictureBoxSingleSingle(0)"
      Tab(1).Control(5)=   "PictureBoxSignalTripleTriple(0)"
      Tab(1).Control(6)=   "PictureBoxSingleSingleSingle(0)"
      Tab(1).Control(7)=   "PictureBoxDoubleDoubleDouble(0)"
      Tab(1).Control(8)=   "PictureBoxTripleTripleTriple(0)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Other Icons"
      TabPicture(2)   =   "DefineBlock.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox PictureBoxTripleTripleTriple 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -69000
         Picture         =   "DefineBlock.frx":0054
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   37
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxDoubleDoubleDouble 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -69240
         Picture         =   "DefineBlock.frx":0682
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   36
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSingleSingleSingle 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -69480
         Picture         =   "DefineBlock.frx":0CB0
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   35
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignalTripleTriple 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -69840
         Picture         =   "DefineBlock.frx":12DE
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   34
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSingleSingle 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70320
         Picture         =   "DefineBlock.frx":190C
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   33
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignalDoubleDouble 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70080
         Picture         =   "DefineBlock.frx":1F3A
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   32
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignalTriple 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70680
         Picture         =   "DefineBlock.frx":2568
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   31
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignalDouble 
         AutoSize        =   -1  'True
         Height          =   675
         Index           =   0
         Left            =   -70920
         Picture         =   "DefineBlock.frx":2B96
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   30
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxSignalSingle 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   675
         Index           =   0
         Left            =   -71160
         Picture         =   "DefineBlock.frx":31C4
         ScaleHeight     =   615
         ScaleWidth      =   135
         TabIndex        =   29
         Top             =   720
         Width           =   195
      End
      Begin VB.PictureBox PictureBoxTrackIconStraight 
         AutoSize        =   -1  'True
         Height          =   330
         Index           =   0
         Left            =   2640
         Picture         =   "DefineBlock.frx":37F2
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   28
         Tag             =   "c:\TrackStraight1.bmp"
         Top             =   840
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch1 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   3360
         Picture         =   "DefineBlock.frx":3FCC
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   27
         Tag             =   "c:\TrackSwitchType1Normal.bmp"
         Top             =   840
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch2 
         AutoSize        =   -1  'True
         Height          =   600
         Index           =   0
         Left            =   4080
         Picture         =   "DefineBlock.frx":4F3E
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   26
         Tag             =   "c:\TrackSwitchType2Normal.bmp"
         Top             =   840
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch3 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   5520
         Picture         =   "DefineBlock.frx":5EB0
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   25
         Tag             =   "c:\TrackSwitchType3Normal.bmp"
         Top             =   840
         Width           =   600
      End
      Begin VB.PictureBox PictureBoxTrackIconSwitch4 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   600
         Index           =   0
         Left            =   4800
         Picture         =   "DefineBlock.frx":6E22
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   24
         Tag             =   "c:\TrackSwitchType4Normal.bmp"
         Top             =   840
         Width           =   600
      End
   End
   Begin VB.Label LabelDescription3 
      Caption         =   "When you are finished drawing out your track diagram, 'update' the database before 'close'ing the window."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   6375
   End
   Begin VB.Label LabelStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label LabelDescription2 
      Caption         =   $"DefineBlock.frx":7D94
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   6375
   End
   Begin VB.Label LabelDescription 
      Caption         =   $"DefineBlock.frx":7E2B
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu menuTrackIcon 
      Caption         =   "TrackIcon"
      Visible         =   0   'False
      Begin VB.Menu menuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu menuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "DefineBlocks"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub ASPictureBox1_DragDrop(Source As Control, X As Single, Y As Single)
Set Source.Container = ASPictureBox1
End Sub


Private Sub ButtonClose_Click()

End

End Sub


Private Sub ButtonUpdate_Click()

Data1.Recordset.MoveLast

Let TemporaryLastRecord = Data1.Recordset.Fields("RecordCounter")

Data1.Recordset.MoveFirst

For X = 0 To Val(TextBoxTrackIconStraightCounter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
    Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconStraight"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxTrackIconStraight(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxTrackIconStraight(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxTrackIconStraight(X).Top
    Data1.Recordset.Update
    
Next X

For X = 0 To Val(TextBoxTrackIconSwitch1Counter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
      Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch1"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxTrackIconSwitch1(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxTrackIconSwitch1(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxTrackIconSwitch1(X).Top
    Data1.Recordset.Update
Next X

For X = 0 To Val(TextBoxTrackIconSwitch2Counter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
      Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch2"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxTrackIconSwitch2(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxTrackIconSwitch2(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxTrackIconSwitch2(X).Top
    Data1.Recordset.Update
Next X

For X = 0 To Val(TextBoxTrackIconSwitch3Counter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
      Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch3"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxTrackIconSwitch3(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxTrackIconSwitch3(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxTrackIconSwitch3(X).Top
    Data1.Recordset.Update
Next X

For X = 0 To Val(TextBoxTrackIconSwitch4Counter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
      Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch4"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxTrackIconSwitch4(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxTrackIconSwitch4(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxTrackIconSwitch4(X).Top
    Data1.Recordset.Update
Next X

If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
    Data1.Recordset.MoveNext
    Data1.Recordset.Edit
Else
    Data1.Recordset.AddNew
End If

For X = 0 To Val(TextBoxSignalDoubleCounter.Text) - 1
    If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
        Data1.Recordset.MoveNext
        Data1.Recordset.Edit
    Else
        Data1.Recordset.AddNew
    End If
      Let Data1.Recordset.Fields("PictureBoxName") = "PictureBoxSignalDouble"
    Let Data1.Recordset.Fields("PictureBoxFileName") = PictureBoxSignalDouble(X).Tag
    Let Data1.Recordset.Fields("PictureBoxLeft") = PictureBoxSignalDouble(X).Left
    Let Data1.Recordset.Fields("PictureBoxTop") = PictureBoxSignalDouble(X).Top
    Data1.Recordset.Update
Next X

If Val(Data1.Recordset.Fields("RecordCounter")) < TemporaryLastRecord Then
    Data1.Recordset.MoveNext
    Data1.Recordset.Edit
Else
    Data1.Recordset.AddNew
End If

Let Data1.Recordset.Fields("PictureBoxName") = "End"
Let Data1.Recordset.Fields("PictureBoxFileName") = "End"
Let Data1.Recordset.Fields("PictureBoxLeft") = 0
Let Data1.Recordset.Fields("PictureBoxTop") = 0
Data1.Recordset.Update
    
End Sub

Private Sub Form_Load()

Dim PictureBoxTrackIconStraight As Object
Dim PictureBoxTrackIconSwitch1 As Object
Dim PictureBoxTrackIconSwitch2 As Object
Dim PictureBoxTrackIconSwitch3 As Object
Dim PictureBoxTrackIconSwitch4 As Object

End Sub


Private Sub menuDelete_Click()

' ======================================================================================================================================================
' InStr Function
'
' Returns the position of the first occurrence of one string within another.
'
' Syntax
'
' InStr([start, ]string1, string2[, compare])
'
' The InStr function syntax has these named arguments:
'
' Part Description
'
' start is a numeric expression that sets the starting position for each search.  If omitted, search begins at the first character
' position.  If start contains Null, an error occurs.  The start argument is required if compare is specified.
' string1 is a string expression being searched.
' string2 is a string expression sought.
' compare specifies the type of string comparison.  The compare argument can be omitted, it can be 0 or 1, or it can be the
' value of the CollatingOrder property of a Field object.  Specify 0 (default) to perform a binary comparison.  Specify 1 to
' perform a textual, case-insensitive comparison.  Specify the return value of the CollatingOrder property of a Field object
' if you want to sort or compare values from a database in the same way the database itself would.  If compare is Null, an
' error occurs.  The start argument is required if compare is specified.  If compare is omitted, the Option Compare setting
' determines the type of comparison.
'
' Return Values
'
' If  InStr returns
'
' string1 is zero-length  0
' string1 is Null Null
' string2 is zero-length  start
' string2 is Null Null
' string2 is not found    0
' string2 is found within string1     Position at which match is found
' start > string2 0
'
' Remarks
'
' Note   Another function (InStrB) is provided for use with byte data contained in a string.  Instead of returning the character position
' of the first occurrence of one string within another, InStrB returns the byte position.
    
    Let TemporaryStartPosition = InStr(1, LabelStatus.Caption, "No.")
    Let TemporaryStopPosition = InStr(1, LabelStatus.Caption, " of ")
    Debug.Print Mid$(LabelStatus.Caption, TemporaryStartPosition + 3, TemporaryStopPosition - TemporaryStartPosition - 3)

End Sub


Private Sub Picture2_DragDrop(Source As Control, X As Single, Y As Single)

Let TemporaryIndex = Val(Source.Index)

If Source.Name = "PictureBoxTrackIconStraight" Then
    If PictureBoxTrackIconStraight(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxTrackIconStraight(TemporaryIndex + 1)
        Set PictureBoxTrackIconStraight(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Top = TextBoxTrackIconStraightPositionTop.Text
        Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Left = TextBoxTrackIconStraightPositionLeft.Text
        Let PictureBoxTrackIconStraight(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Visible = True
        Let TextBoxTrackIconStraightCounter.Text = Val(TextBoxTrackIconStraightCounter.Text) + 1
    End If
ElseIf Source.Name = "PictureBoxTrackIconSwitch1" Then
    If PictureBoxTrackIconSwitch1(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxTrackIconSwitch1(TemporaryIndex + 1)
        Set PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch1PositionTop.Text
        Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch1PositionLeft.Text
        Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Visible = True
        Let TextBoxTrackIconSwitch1Counter.Text = Val(TextBoxTrackIconSwitch1Counter.Text) + 1
    End If
ElseIf Source.Name = "PictureBoxTrackIconSwitch2" Then
    If PictureBoxTrackIconSwitch2(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxTrackIconSwitch2(TemporaryIndex + 1)
        Set PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch2PositionTop.Text
        Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch2PositionLeft.Text
        Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Visible = True
        Let TextBoxTrackIconSwitch2Counter.Text = Val(TextBoxTrackIconSwitch2Counter.Text) + 1
    End If
ElseIf Source.Name = "PictureBoxTrackIconSwitch3" Then
    If PictureBoxTrackIconSwitch3(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxTrackIconSwitch3(TemporaryIndex + 1)
        Set PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch3PositionTop.Text
        Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch3PositionLeft.Text
        Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Visible = True
        Let TextBoxTrackIconSwitch3Counter.Text = Val(TextBoxTrackIconSwitch3Counter.Text) + 1
    End If
ElseIf Source.Name = "PictureBoxTrackIconSwitch4" Then
    If PictureBoxTrackIconSwitch4(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxTrackIconSwitch4(TemporaryIndex + 1)
        Set PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch4PositionTop.Text
        Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch4PositionLeft.Text
        Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Visible = True
        Let TextBoxTrackIconSwitch4Counter.Text = Val(TextBoxTrackIconSwitch4Counter.Text) + 1
    End If
ElseIf Source.Name = "PictureBoxSignalDouble" Then
   If PictureBoxSignalDouble(TemporaryIndex).Container.Name <> "Picture2" Then
        Load PictureBoxSignalDouble(TemporaryIndex + 1)
        Set PictureBoxSignalDouble(TemporaryIndex + 1).Container = TabTrackIcon
        Let PictureBoxSignalDouble(TemporaryIndex + 1).Top = TextBoxSignalDoublePositionTop.Text
        Let PictureBoxSignalDouble(TemporaryIndex + 1).Left = TextBoxSignalDoublePositionLeft.Text
        Let PictureBoxSignalDouble(TemporaryIndex + 1).BorderStyle = 1
        Let PictureBoxSignalDouble(TemporaryIndex + 1).Picture = LoadPicture("c:\SignalDouble.bmp")
        Let PictureBoxSignalDouble(TemporaryIndex + 1).Visible = True
        Let TextBoxSignalDoubleCounter.Text = Val(TextBoxSignalDoubleCounter.Text) + 1
        Source.Picture = LoadPicture("c:\SignalDoubleWest.bmp")
        Source.Tag = "c:\SignalDoubleWest.bmp"
    End If

End If

Set Source.Container = Picture2
Let Source.Left = X
Let Source.Top = Y
Let NewX = (Int(X / 36) * 36)
Let NewY = (Int(Y / 18) * 18)
Let Source.Left = NewX
Let Source.Top = NewY
Let Source.BorderStyle = 1

End Sub

Private Sub PictureBoxTrackIcon_Click(Index As Integer)

End Sub

Private Sub PictureBoxTrackIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub PictureBoxDoubleDoubleDouble_Click(Index As Integer)

Let PictureBoxSignalDoubleDoubleDouble(Index).BorderStyle = 1

End Sub

Private Sub PictureBoxSignalDouble_DblClick(Index As Integer)


If PictureBoxSignalDouble(Index).Container.Name = "Picture2" Then
    If PictureBoxSignalDouble(Index).Tag = "c:\SignalDoubleEast.bmp" Then
        Let PictureBoxSignalDouble(Index).Tag = "c:\SignalDoubleWest.bmp"
        PictureBoxSignalDouble(Index).Picture = LoadPicture("c:\SignalDoubleWest.bmp")
    ElseIf PictureBoxSignalDouble(Index).Tag = "c:\SignalDoubleWest.bmp" Then
        Let PictureBoxSignalDouble(Index).Tag = "c:\SignalDoubleEast.bmp"
        PictureBoxSignalDouble(Index).Picture = LoadPicture("c:\SignalDoubleEast.bmp")
    End If
End If

End Sub

Private Sub PictureBoxSignalDouble_GotFocus(Index As Integer)

Let PictureBoxSignalDouble(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxSignalDouble_LostFocus(Index As Integer)

If PictureBoxSignalDouble(Index).Container.Name = "Picture2" Then
    PictureBoxSignalDouble(Index).BorderStyle = 0
ElseIf PictureBoxSignalDouble(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxSignalDouble(Index).BorderStyle = 1
End If

End Sub


Private Sub PictureBoxSignalDouble_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PictureBoxSignalDouble(Index).Drag vbBeginDrag

End Sub

Private Sub PictureBoxSignalDoubleDouble_Click(Index As Integer)

Let PictureBoxSignalDoubleDouble(Index).BorderStyle = 1

End Sub

Private Sub PictureBoxSignalTriple_GotFocus(Index As Integer)

Let PictureBoxSignalTriple(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxSignalTripleTriple_Click(Index As Integer)

Let PictureBoxSignalTripleTriple(Index).BorderStyle = 1

End Sub

Private Sub PictureBoxSingle_GotFocus(Index As Integer)

Let PictureBoxSignalSingle(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxSingle_LostFocus(Index As Integer)

If PictureBoxTrackIconStraight(Index).Container.Name = "Picture1" Then
    PictureBoxTrackIconStraight(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconStraight(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconStraight(Index).BorderStyle = 1
End If

End Sub


Private Sub PictureBoxSingleSingle_GotFocus(Index As Integer)

Let PictureBoxSignalSingleSingle(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxSingleSingleSingle_GotFocus(Index As Integer)

Let PictureBoxSignalSingleSingleSingle(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTrackIconStraight_DblClick(Index As Integer)

If PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight1.bmp" Then
    Let PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight2.bmp"
    PictureBoxTrackIconStraight(Index).Picture = LoadPicture("c:\TrackStraight2.bmp")
ElseIf PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight2.bmp" Then
    Let PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight3.bmp"
    PictureBoxTrackIconStraight(Index).Picture = LoadPicture("c:\TrackStraight3.bmp")
ElseIf PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight3.bmp" Then
    Let PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight4.bmp"
    PictureBoxTrackIconStraight(Index).Picture = LoadPicture("c:\TrackStraight4.bmp")
ElseIf PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight4.bmp" Then
    Let PictureBoxTrackIconStraight(Index).Tag = "c:\TrackStraight1.bmp"
    PictureBoxTrackIconStraight(Index).Picture = LoadPicture("c:\TrackStraight1.bmp")
End If

End Sub

Private Sub PictureBoxTrackIconStraight_GotFocus(Index As Integer)

Let PictureBoxTrackIconStraight(Index).BorderStyle = 1

Let LabelStatus.Caption = "Status: Straight Block No." + Str$(Index) + " of " + Str$(TextBoxTrackIconStraightCounter.Text)

End Sub

Private Sub PictureBoxTrackIconStraight_LostFocus(Index As Integer)

If PictureBoxTrackIconStraight(Index).Container.Name = "Picture2" Then
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

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Displaying a Pop-up Menu
'
' Displays a pop-up menu on an MDIForm or Form object at the current mouse location or at specified coordinates.  Doesn't support named arguments.
'
' Syntax
'
'   object.PopupMenu menuname, Flags, X, Y, boldcommand
'
' The PopupMenu method syntax has these parts:
'
' Part Description
'
' object  is optional.  An object expression that evaluates to an object in the Applies To list.  If object is omitted, the form with the focus is assumed to be object.
'
' menuname is required.  The name of the pop-up menu to be displayed.  The specified menu must have at least one submenu.
'
' flags is optional.  A value or constant that specifies the location and behavior of a pop-up menu, as described in Settings.
' x is optional.  Specifies the x-coordinate where the pop-up menu is displayed.  If omitted, the mouse coordinate is used.
' y is optional.  Specifies the y-coordinate where the pop-up menu is displayed.  If omitted, the mouse coordinate is used.
' boldcommand is optional.  Specifies the name of a menu control in the pop-up menu to display its caption in bold text.  If
' omitted, no controls in the pop-up menu appear in bold. This argument works only for applications running under
' Windows 95. The application will ignore this argument when running under 16-bit versions of Windows or Windows
' NT 3.51 and earlier.
'
' Settings
'
' The settings for flags are:
'
' Constant (location) Value   Description
'
' vbPopupMenuLeftAlign    0   (Default) The left side of the pop-up menu is located at x.
' vbPopupMenuCenterAlign  4   The pop-up menu is centered at x.
' vbPopupMenuRightAlign   8   The right side of the pop-up menu is located at x.
'
' Constant (behavior) Value   Description
'
' vbPopupMenuLeftButton   0   (Default) An item on the pop-up menu reacts to a mouse click only when you use the left mouse button.
' vbPopupMenuRightButton  2   An item on the pop-up menu reacts to a mouse click when you use either the right or the left mouse button.
'
' Note   The flags parameter has no effect on applications running under Microsoft Windows version 3.0 or earlier.  To specify
' two flags, combine one constant from each group using the Or operator.
'
' Remarks
'
' These constants are listed in the Visual Basic (VB) object library in the Object Browser.
' You specify the unit of measure for the x and y coordinates using the ScaleMode property.  The x and y coordinates
' define where the pop-up is displayed relative to the specified form.  If the x and y coordinates aren't included, the pop-up
' menu is displayed at the current location of the mouse pointer. When you display a pop-up menu, the code following
' the call to the PopupMenu method isn't executed until the user either chooses a command from the menu (in which
' case the code for that command's Click event is executed before the code following the PopupMenu statement) or
' cancels the menu.  In addition, only one pop-up menu can be displayed at a time; therefore, calls to this method are
' ignored if a pop-up menu is already displayed or if a pull-down menu is open.
'
' A pop-up menu that isn't visible on the menu bar because its Visible property is set to False in the Menu Editor can
' still be displayed because the Visible property of the specified menu is ignored when Visual Basic displays a pop-up menu.

    Form1.PopupMenu menuTrackIcon, vbPopupMenuLeftAlign

End If

End Sub


Private Sub PictureBoxTrackIconSwitch1_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch1(Index).Tag = "c:\TrackSwitchType1Normal.bmp" Then
    Let PictureBoxTrackIconSwitch1(Index).Tag = "c:\TrackSwitchType1Reverse.bmp"
    PictureBoxTrackIconSwitch1(Index).Picture = LoadPicture("c:\TrackSwitchType1Reverse.bmp")
Else
    Let PictureBoxTrackIconSwitch1(Index).Tag = "c:\TrackSwitchType1Normal.bmp"
    PictureBoxTrackIconSwitch1(Index).Picture = LoadPicture("c:\TrackSwitchType1Normal.bmp")
End If

End Sub


Private Sub PictureBoxTrackIconSwitch1_GotFocus(Index As Integer)

PictureBoxTrackIconSwitch1(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTrackIconSwitch1_LostFocus(Index As Integer)

If PictureBoxTrackIconSwitch1(Index).Container.Name = "Picture2" Then
    PictureBoxTrackIconSwitch1(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconSwitch1(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconSwitch1(Index).BorderStyle = 1
End If

End Sub


Private Sub PictureBoxTrackIconSwitch1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PictureBoxTrackIconSwitch1(Index).Drag vbBeginDrag

End Sub

Private Sub PictureBoxTrackIconSwitch2_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch2(Index).Tag = "c:\TrackSwitchType2Normal.bmp" Then
    Let PictureBoxTrackIconSwitch2(Index).Tag = "c:\TrackSwitchType2Reverse.bmp"
    PictureBoxTrackIconSwitch2(Index).Picture = LoadPicture("c:\TrackSwitchType2Reverse.bmp")
Else
Let PictureBoxTrackIconSwitch2(Index).Tag = "c:\TrackSwitchType2Normal.bmp"
    PictureBoxTrackIconSwitch2(Index).Picture = LoadPicture("c:\TrackSwitchType2Normal.bmp")
End If

End Sub


Private Sub PictureBoxTrackIconSwitch2_GotFocus(Index As Integer)

Let PictureBoxTrackIconSwitch2(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTrackIconSwitch2_LostFocus(Index As Integer)

If PictureBoxTrackIconSwitch2(Index).Container.Name = "Picture2" Then
    PictureBoxTrackIconSwitch2(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconSwitch2(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconSwitch2(Index).BorderStyle = 1
End If

End Sub


Private Sub PictureBoxTrackIconSwitch2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PictureBoxTrackIconSwitch2(Index).Drag vbBeginDrag

End Sub


Private Sub PictureBoxTrackIconSwitch3_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch3(Index).Tag = "c:\TrackSwitchType3Normal.bmp" Then
    Let PictureBoxTrackIconSwitch3(Index).Tag = "c:\TrackSwitchType3Reverse.bmp"
    PictureBoxTrackIconSwitch3(Index).Picture = LoadPicture("c:\TrackSwitchType3Reverse.bmp")
Else
Let PictureBoxTrackIconSwitch3(Index).Tag = "c:\TrackSwitchType3Normal.bmp"
    PictureBoxTrackIconSwitch3(Index).Picture = LoadPicture("c:\TrackSwitchType3Normal.bmp")
End If

End Sub


Private Sub PictureBoxTrackIconSwitch3_GotFocus(Index As Integer)

Let PictureBoxTrackIconSwitch3(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTrackIconSwitch3_LostFocus(Index As Integer)

If PictureBoxTrackIconSwitch3(Index).Container.Name = "Picture2" Then
    PictureBoxTrackIconSwitch3(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconSwitch3(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconSwitch3(Index).BorderStyle = 1
End If

End Sub


Private Sub PictureBoxTrackIconSwitch3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PictureBoxTrackIconSwitch3(Index).Drag vbBeginDrag

End Sub


Private Sub PictureBoxTrackIconSwitch4_DblClick(Index As Integer)

If PictureBoxTrackIconSwitch4(Index).Tag = "c:\TrackSwitchType4Normal.bmp" Then
    Let PictureBoxTrackIconSwitch4(Index).Tag = "c:\TrackSwitchType4Reverse.bmp"
    PictureBoxTrackIconSwitch4(Index).Picture = LoadPicture("c:\TrackSwitchType4Reverse.bmp")
Else
Let PictureBoxTrackIconSwitch4(Index).Tag = "c:\TrackSwitchType4Normal.bmp"
    PictureBoxTrackIconSwitch4(Index).Picture = LoadPicture("c:\TrackSwitchType4Normal.bmp")
End If

End Sub


Private Sub PictureBoxTrackIconSwitch4_GotFocus(Index As Integer)

Let PictureBoxTrackIconSwitch4(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTrackIconSwitch4_LostFocus(Index As Integer)

If PictureBoxTrackIconSwitch4(Index).Container.Name = "Picture2" Then
    PictureBoxTrackIconSwitch4(Index).BorderStyle = 0
ElseIf PictureBoxTrackIconSwitch4(Index).Container.Name = "TabTrackIcon" Then
    PictureBoxTrackIconSwitch4(Index).BorderStyle = 1
End If

End Sub

Private Sub PictureBoxTrackIconSwitch4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

PictureBoxTrackIconSwitch4(Index).Drag vbBeginDrag

End Sub


Private Sub PicturePictureBox_Click(Index As Integer)

End Sub

Private Sub PicturePictureBox_GotFocus(Index As Integer)

Let PictureBoxTrackIconStraight(Index).BorderStyle = 1

End Sub


Private Sub PictureBoxTripleTripleTriple_Click(Index As Integer)

Let PictureBoxSignalTripleTripleTriple(Index).BorderStyle = 1

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

Private Sub Timer1_Timer()

Timer1.Interval = 0

Let TabTrackIcon.Tab = 0
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
Let TextBoxSignalDoublePositionTop.Text = PictureBoxSignalDouble(0).Top
Let TextBoxSignalDoublePositionLeft.Text = PictureBoxSignalDouble(0).Left

End Sub


Private Sub Timer2_Timer()

Timer2.Interval = 0

Data1.Recordset.MoveFirst

Do

If Data1.Recordset.Fields("PictureBoxName") = "End" Then Exit Do

If Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconStraight" Then
    Let TabTrackIcon.Tab = 0
    Let TemporaryIndex = Val(TextBoxTrackIconStraightCounter.Text)
    Set PictureBoxTrackIconStraight(TemporaryIndex).Container = Picture2
    Let PictureBoxTrackIconStraight(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconStraight(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconStraight(TemporaryIndex).BorderStyle = 0
    PictureBoxTrackIconStraight(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconStraight(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxTrackIconStraight(TemporaryIndex).Visible = True
    
    Load PictureBoxTrackIconStraight(TemporaryIndex + 1)
    Set PictureBoxTrackIconStraight(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Top = TextBoxTrackIconStraightPositionTop.Text
    Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Left = TextBoxTrackIconStraightPositionLeft.Text
    Let PictureBoxTrackIconStraight(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxTrackIconStraight(TemporaryIndex + 1).Visible = True
    Let TextBoxTrackIconStraightCounter.Text = Val(TextBoxTrackIconStraightCounter.Text) + 1
    
ElseIf Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch1" Then
    Let TabTrackIcon.Tab = 0
    Let TemporaryIndex = Val(TextBoxTrackIconSwitch1Counter.Text)
    Set PictureBoxTrackIconSwitch1(TemporaryIndex).Container = Picture2
    Let PictureBoxTrackIconSwitch1(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch1(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch1(TemporaryIndex).BorderStyle = 0
    PictureBoxTrackIconSwitch1(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch1(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxTrackIconSwitch1(TemporaryIndex).Visible = True

    Load PictureBoxTrackIconSwitch1(TemporaryIndex + 1)
    Set PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch1PositionTop.Text
    Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch1PositionLeft.Text
    Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxTrackIconSwitch1(TemporaryIndex + 1).Visible = True
    Let TextBoxTrackIconSwitch1Counter.Text = Val(TextBoxTrackIconSwitch1Counter.Text) + 1

ElseIf Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch2" Then
    Let TabTrackIcon.Tab = 0
    Let TemporaryIndex = Val(TextBoxTrackIconSwitch2Counter.Text)
    Set PictureBoxTrackIconSwitch2(TemporaryIndex).Container = Picture2
    Let PictureBoxTrackIconSwitch2(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch2(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch2(TemporaryIndex).BorderStyle = 0
    PictureBoxTrackIconSwitch2(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch2(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxTrackIconSwitch2(TemporaryIndex).Visible = True
    
    Load PictureBoxTrackIconSwitch2(TemporaryIndex + 1)
    Set PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch2PositionTop.Text
    Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch2PositionLeft.Text
    Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxTrackIconSwitch2(TemporaryIndex + 1).Visible = True
    Let TextBoxTrackIconSwitch2Counter.Text = Val(TextBoxTrackIconSwitch2Counter.Text) + 1

ElseIf Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch3" Then
    Let TabTrackIcon.Tab = 0
    Let TemporaryIndex = Val(TextBoxTrackIconSwitch3Counter.Text)
    Set PictureBoxTrackIconSwitch3(TemporaryIndex).Container = Picture2
    Let PictureBoxTrackIconSwitch3(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch3(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch3(TemporaryIndex).BorderStyle = 0
    PictureBoxTrackIconSwitch3(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch3(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxTrackIconSwitch3(TemporaryIndex).Visible = True

    Load PictureBoxTrackIconSwitch3(TemporaryIndex + 1)
    Set PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch3PositionTop.Text
    Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch3PositionLeft.Text
    Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxTrackIconSwitch3(TemporaryIndex + 1).Visible = True
    Let TextBoxTrackIconSwitch3Counter.Text = Val(TextBoxTrackIconSwitch3Counter.Text) + 1

ElseIf Data1.Recordset.Fields("PictureBoxName") = "PictureBoxTrackIconSwitch4" Then
    Let TabTrackIcon.Tab = 0
    Let TemporaryIndex = Val(TextBoxTrackIconSwitch4Counter.Text)
    Set PictureBoxTrackIconSwitch4(TemporaryIndex).Container = Picture2
    Let PictureBoxTrackIconSwitch4(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxTrackIconSwitch4(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxTrackIconSwitch4(TemporaryIndex).BorderStyle = 0
    PictureBoxTrackIconSwitch4(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxTrackIconSwitch4(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxTrackIconSwitch4(TemporaryIndex).Visible = True
    
    Load PictureBoxTrackIconSwitch4(TemporaryIndex + 1)
    Set PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Top = TextBoxTrackIconSwitch4PositionTop.Text
    Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Left = TextBoxTrackIconSwitch4PositionLeft.Text
    Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxTrackIconSwitch4(TemporaryIndex + 1).Visible = True
    Let TextBoxTrackIconSwitch4Counter.Text = Val(TextBoxTrackIconSwitch4Counter.Text) + 1

ElseIf Data1.Recordset.Fields("PictureBoxName") = "PictureBoxSignalDouble" Then
    Let TabTrackIcon.Tab = 1
    Let TemporaryIndex = Val(TextBoxSignalDoubleCounter.Text)
    Set PictureBoxSignalDouble(TemporaryIndex).Container = Picture2
    Let PictureBoxSignalDouble(TemporaryIndex).Top = Val(Data1.Recordset.Fields("PictureBoxTop"))
    Let PictureBoxSignalDouble(TemporaryIndex).Left = Val(Data1.Recordset.Fields("PictureBoxLeft"))
    Let PictureBoxSignalDouble(TemporaryIndex).BorderStyle = 0
    PictureBoxSignalDouble(TemporaryIndex).Picture = LoadPicture(Data1.Recordset.Fields("PictureBoxFileName"))
    Let PictureBoxSignalDouble(TemporaryIndex).Tag = Data1.Recordset.Fields("PictureBoxFileName")
    Let PictureBoxSignalDouble(TemporaryIndex).Visible = True
    
    Load PictureBoxSignalDouble(TemporaryIndex + 1)
    Set PictureBoxSignalDouble(TemporaryIndex + 1).Container = TabTrackIcon
    Let PictureBoxSignalDouble(TemporaryIndex + 1).Top = TextBoxSignalDoublePositionTop.Text
    Let PictureBoxSignalDouble(TemporaryIndex + 1).Left = TextBoxSignalDoublePositionLeft.Text
    Let PictureBoxSignalDouble(TemporaryIndex + 1).BorderStyle = 1
    Let PictureBoxSignalDouble(TemporaryIndex + 1).Picture = LoadPicture("c:\SignalDouble.bmp")
    Let PictureBoxSignalDouble(TemporaryIndex + 1).Tag = "c:\SignalDouble.bmp"
    Let PictureBoxSignalDouble(TemporaryIndex + 1).Visible = True
    Let TextBoxSignalDoubleCounter.Text = Val(TextBoxSignalDoubleCounter.Text) + 1
End If

Data1.Recordset.MoveNext

Loop While Not Data1.Recordset.EOF

End Sub


