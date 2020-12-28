VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3885
   Height          =   2040
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   3885
   Top             =   1170
   Width           =   4005
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\LocomotiveDatabaseDecoders.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Automatic Train Control\Databases\LocomotiveDatabaseDecoders.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Kill "C:\Automatic Train Control\Databases\Test.mdb"

Call DecoderLong
Call DecoderShort

End Sub



Public Sub DecoderLong()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef
Dim myF(0 To 128)

Set OldWs = DBEngine.Workspaces(0)
Set NewDb = OldWs.OpenDatabase("c:\Automatic Train Control\Databases\LocomotiveDatabaseDecoders.mdb")
Set myTable = NewDb.CreateTableDef("LocomotiveDecodersLong")

Let TemporaryFieldName = "LocomotiveNumber"

Set myF(0) = myTable.CreateField(TemporaryFieldName, dbInteger)
Let myF(0).Required = True
'Let myF(0).AllowZeroLength = False
Let myF(0).Attributes = dbUpdatableField
Let myF(0).DefaultValue = "0"
myTable.Fields.Append myF(0)

For t = 1 To 128

Let TemporaryFieldName = "LocomotiveDecoderCV" + Right$(Str$(t), Len(Str$(t)) - 1) + "D"

Set myF(t) = myTable.CreateField(TemporaryFieldName, dbInteger)
Let myF(t).Required = False
'Let myF(t).AllowZeroLength = True
Let myF(t).Attributes = dbUpdatableField
Let myF(t).DefaultValue = "0"
myTable.Fields.Append myF(t)

Next t

NewDb.TableDefs.Append myTable

Set myrecord = NewDb.OpenRecordset("LocomotiveDecodersLong")

For t = 0 To 10000
    myrecord.AddNew
    myrecord.Fields("LocomotiveNumber") = Str$(t)
    DoEvents
    myrecord.Update
    DoEvents
Next

NewDb.Close

End Sub

Public Sub DecoderShort()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef
Dim myF(0 To 128)

Set OldWs = DBEngine.Workspaces(0)
Set NewDb = OldWs.OpenDatabase("c:\Automatic Train Control\Databases\LocomotiveDatabaseDecoders.mdb")
Set myTable = NewDb.CreateTableDef("LocomotiveDecodersShort")

Let TemporaryFieldName = "LocomotiveNumber"

Set myF(0) = myTable.CreateField(TemporaryFieldName, dbInteger)
Let myF(0).Required = True
'Let myF(0).AllowZeroLength = False
Let myF(0).Attributes = dbUpdatableField
Let myF(0).DefaultValue = "0"
myTable.Fields.Append myF(0)

For t = 1 To 128

Let TemporaryFieldName = "LocomotiveDecoderCV" + Right$(Str$(t), Len(Str$(t)) - 1) + "D"

Set myF(t) = myTable.CreateField(TemporaryFieldName, dbInteger)
Let myF(t).Required = False
'Let myF(t).AllowZeroLength = True
Let myF(t).Attributes = dbUpdatableField
Let myF(t).DefaultValue = "0"
myTable.Fields.Append myF(t)

Next t

NewDb.TableDefs.Append myTable

Set myrecord = NewDb.OpenRecordset("LocomotiveDecodersShort")

For t = 0 To 127
    myrecord.AddNew
    myrecord.Fields("LocomotiveNumber") = Str$(t)
    DoEvents
    myrecord.Update
    DoEvents
Next



NewDb.Close

End Sub

Private Sub Command2_Click()

Let Data1.RecordSource = "LocomotiveDecoders"
Let Data2.RecordSource = "LocomotiveDecodersLong"
Data1.Refresh
Data2.Refresh
DoEvents

Data1.Recordset.MoveFirst
Data2.Recordset.MoveFirst

While x < 10001

Data2.Recordset.Edit

Let Data2.Recordset.Fields("LocomotiveNumber") = Data1.Recordset.Fields("LocomotiveNumber")

For t = 1 To 127

Let TemporaryFieldName = "LocomotiveDecoderCV"
Let TemporaryFieldName = TemporaryFieldName + Right$(Str$(t), Len(Str$(t)) - 1)
Let TemporaryFieldName = TemporaryFieldName + "D"

If Data1.Recordset.Fields(TemporaryFieldName) <> Null Then
    Let Data2.Recordset.Fields(TemporaryFieldName) = Val(Data1.Recordset.Fields(TemporaryFieldName))
Else
    Let Data2.Recordset.Fields(TemporaryFieldName) = 0
End If

Next t

Data2.Recordset.Update

DoEvents

Data1.Recordset.MoveNext
Data2.Recordset.MoveNext
x = x + 1

Wend

End Sub


Private Sub Data1_Reposition()

Data1.Caption = Data1.Recordset.AbsolutePosition

End Sub

Private Sub Data2_Reposition()

Data2.Caption = Data2.Recordset.AbsolutePosition

End Sub

