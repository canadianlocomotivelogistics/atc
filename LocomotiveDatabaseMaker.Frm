VERSION 4.00
Begin VB.Form LocomotiveDatabaseMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locomotive Database Maker"
   ClientHeight    =   6180
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3930
   Height          =   6585
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   3930
   Top             =   1170
   Width           =   4050
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "LocomotiveDatabaseMaker"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

Call AddressLong
Call AddressShort

End

End Sub


Private Sub Command2_Click()
Dim DefaultWorkspace As Workspace
    Dim MyDatabase As Database, TempDatabase As Database
    Dim I As Integer, J As Integer
    Set DefaultWorkspace = Workspaces(0)
    Set MyDatabase = DefaultWorkspace.OpenDatabase("c:\Automatic Train Control\Databases\LocomotiveDatabaseDiesels.MDB")
' Enumerate all open databases.
    For J = 0 To DefaultWorkspace.Databases.Count - 1
        Set TempDatabase = DefaultWorkspace.Databases(J)
        Debug.Print
        Debug.Print "Enumeration of Databases("; J; "): "; TempDatabase.Name
        Debug.Print
' Enumerate containers.
        Debug.Print "Container: Name, Owner"
        For I = 0 To TempDatabase.Containers.Count - 1

Debug.Print "  "; TempDatabase.Containers(I).Name;
            Debug.Print ", "; TempDatabase.Containers(I).Owner
        Next I
        Debug.Print
' Enumerate query definitions.
        Debug.Print "QueryDef: Name"
        For I = 0 To TempDatabase.QueryDefs.Count - 1
            Debug.Print "  "; TempDatabase.QueryDefs(I).Name
        Next I
        Debug.Print
' Enumerate recordsets.
' No output because no Recordset is open.
        Debug.Print "Recordset: Name"
        For I = 0 To TempDatabase.Recordsets.Count - 1
            Debug.Print "  "; TempDatabase.Recordsets(I).Name
        Next I
        Debug.Print
' Enumerate relationships.
        Debug.Print "Relation: Name, Table, ForeignTable"
        For I = 0 To TempDatabase.Relations.Count - 1

Debug.Print "  "; TempDatabase.Relations(I).Name;
            Debug.Print ", "; TempDatabase.Relations(I).Table;
            Debug.Print ", "; TempDatabase.Relations(I).ForeignTable
        Next I
        Debug.Print
' Enumerate table definitions.
        Debug.Print "TableDef: Name, DateCreated"
        For I = 0 To TempDatabase.TableDefs.Count - 1
            Debug.Print "  "; TempDatabase.TableDefs(I).Name;
            Debug.Print ", "; TempDatabase.TableDefs(I).DateCreated
        Next I
        Debug.Print
    Next J
' Enumerate built-in properties of MyDatabase.
    Debug.Print "MyDatabase.Name: "; MyDatabase.Name
    Debug.Print "MyDatabase.CollatingOrder: "; MyDatabase.CollatingOrder

Debug.Print "MyDatabase.Connect: "; MyDatabase.Connect
    Debug.Print "MyDatabase.QueryTimeout: "; MyDatabase.QueryTimeout
    Debug.Print "MyDatabase.Transactions: "; MyDatabase.Transactions
    Debug.Print "MyDatabase.Updatable: "; MyDatabase.Updatable
    'Debug.Print "MyDatabase.V1xNullBehavior: "; MyDatabase.V1xNullBehavior
    Debug.Print
    MyDatabase.Close    ' File remains on disk.
    EnumerateDatabase = True

End Sub



Public Sub AddressLong()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef

Set OldWs = DBEngine.Workspaces(0)
Set NewDb = OldWs.OpenDatabase("c:\Automatic Train Control\Databases\Test.mdb")
Set myTable = NewDb.CreateTableDef("LocomotiveAddressLong")

Set myF1 = myTable.CreateField("LocomotiveAddress", dbInteger)
'Let myF1.AllowZeroLength = True 'Only for text or memo types only
Let myF1.Attributes = dbUpdatableField
Let myF1.DefaultValue = "0000"

Set myF2 = myTable.CreateField("LocomotivePrototypeOwner", dbMemo)
Let myF2.AllowZeroLength = True
Let myF2.Attributes = dbUpdatableField
Let myF2.DefaultValue = "Locomotive Prototype Owner"

Set myF3 = myTable.CreateField("LocomotiveActiveList", dbBoolean)
'Let myF3.AllowZeroLength = True
Let myF3.Attributes = dbUpdatableField
Let myF3.DefaultValue = False

Set myF4 = myTable.CreateField("LocomotiveDecoderEquiped", dbBoolean)
'Let myF4.AllowZeroLength = True
Let myF4.Attributes = dbUpdatableField
Let myF4.DefaultValue = False

Set myF5 = myTable.CreateField("LocomotiveManufacturer", dbMemo)
Let myF5.AllowZeroLength = True
Let myF5.Attributes = dbUpdatableField
Let myF5.DefaultValue = "Locomotive Manufacturer"

Set myF6 = myTable.CreateField("LocomotiveModel", dbMemo)
Let myF6.AllowZeroLength = True
Let myF6.Attributes = dbUpdatableField
Let myF6.DefaultValue = "Locomotive Model"

Set myF7 = myTable.CreateField("LocomotiveModelManufacturer", dbMemo)
Let myF7.AllowZeroLength = True
Let myF7.Attributes = dbUpdatableField
Let myF7.DefaultValue = "Locomotive Model Manufacturer"

Set myF8 = myTable.CreateField("LocomotiveFacts", dbMemo)
Let myF8.AllowZeroLength = True
Let myF8.Attributes = dbUpdatableField
Let myF8.DefaultValue = "Locomotive Facts"

Set myF9 = myTable.CreateField("LocomotiveImage", dbMemo)
Let myF9.AllowZeroLength = True
Let myF9.Attributes = dbUpdatableField
Let myF9.DefaultValue = ""

Set myF10 = myTable.CreateField("LocomotivePrototypeAdhesionFactor", dbMemo)
Let myF10.AllowZeroLength = True
Let myF10.Attributes = dbUpdatableField
Let myF10.DefaultValue = "0"

Set myF11 = myTable.CreateField("LocomotivePrototypeHorsePower", dbMemo)
Let myF11.AllowZeroLength = True
Let myF11.Attributes = dbUpdatableField
Let myF11.DefaultValue = "0"

Set myF12 = myTable.CreateField("LocomotivePrototypeTractionEffort", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotivePrototypeDrawBarPull", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"


Set myF12 = myTable.CreateField("LocomotivePrototypeWeight", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotiveModelAdhesionFactor", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"

Set myF14 = myTable.CreateField("LocomotiveModelHorsePower", dbMemo)
Let myF14.AllowZeroLength = True
Let myF14.Attributes = dbUpdatableField
Let myF14.DefaultValue = "0"

Set myF15 = myTable.CreateField("LocomotiveModelTractionEffort", dbMemo)
Let myF15.AllowZeroLength = True
Let myF15.Attributes = dbUpdatableField
Let myF15.DefaultValue = "0"

Set myF16 = myTable.CreateField("LocomotiveModelDrawBarPull", dbMemo)
Let myF16.AllowZeroLength = True
Let myF16.Attributes = dbUpdatableField
Let myF16.DefaultValue = "0"


Set myF17 = myTable.CreateField("LocomotiveModelWeight", dbMemo)
Let myF17.AllowZeroLength = True
Let myF17.Attributes = dbUpdatableField
Let myF17.DefaultValue = "0"

myTable.Fields.Append myF1
myTable.Fields.Append myF2
myTable.Fields.Append myF3
myTable.Fields.Append myF4
myTable.Fields.Append myF5
myTable.Fields.Append myF6
myTable.Fields.Append myF7
myTable.Fields.Append myF8
myTable.Fields.Append myF9
myTable.Fields.Append myF10
myTable.Fields.Append myF11
myTable.Fields.Append myF12
myTable.Fields.Append myF13
myTable.Fields.Append myF14
myTable.Fields.Append myF15
myTable.Fields.Append myF16
myTable.Fields.Append myF17

NewDb.TableDefs.Append myTable

Set myRecord = NewDb.OpenRecordset("LocomotiveAddressLong")

For t = 1 To 1000
    myRecord.AddNew
    myRecord.Fields("LocomotiveAddress") = Str$(t)
    myRecord.Update
    DoEvents
Next



NewDb.Close

End Sub

Public Sub AddressShort()

Dim OldWs As Workspace
Dim NewDb As Database
Dim myTable As TableDef
'Dim myRecord As Recordset

Set OldWs = DBEngine.Workspaces(0)
Set NewDb = OldWs.OpenDatabase("c:\Automatic Train Control\Databases\Test.mdb")
Set myTable = NewDb.CreateTableDef("LocomotiveAddressShort")

Set myF1 = myTable.CreateField("LocomotiveAddress", dbInteger)
'Let myF1.AllowZeroLength = True 'Only for text or memo types only
Let myF1.Attributes = dbUpdatableField
Let myF1.DefaultValue = "0000"

Set myF2 = myTable.CreateField("LocomotivePrototypeOwner", dbMemo)
Let myF2.AllowZeroLength = True
Let myF2.Attributes = dbUpdatableField
Let myF2.DefaultValue = "Locomotive Prototype Owner"

Set myF3 = myTable.CreateField("LocomotiveActiveList", dbBoolean)
'Let myF3.AllowZeroLength = True
Let myF3.Attributes = dbUpdatableField
Let myF3.DefaultValue = False

Set myF4 = myTable.CreateField("LocomotiveDecoderEquiped", dbBoolean)
'Let myF4.AllowZeroLength = True
Let myF4.Attributes = dbUpdatableField
Let myF4.DefaultValue = False

Set myF5 = myTable.CreateField("LocomotiveManufacturer", dbMemo)
Let myF5.AllowZeroLength = True
Let myF5.Attributes = dbUpdatableField
Let myF5.DefaultValue = "Locomotive Manufacturer"

Set myF6 = myTable.CreateField("LocomotiveModel", dbMemo)
Let myF6.AllowZeroLength = True
Let myF6.Attributes = dbUpdatableField
Let myF6.DefaultValue = "Locomotive Model"

Set myF7 = myTable.CreateField("LocomotiveModelManufacturer", dbMemo)
Let myF7.AllowZeroLength = True
Let myF7.Attributes = dbUpdatableField
Let myF7.DefaultValue = "Locomotive Model Manufacturer"

Set myF8 = myTable.CreateField("LocomotiveFacts", dbMemo)
Let myF8.AllowZeroLength = True
Let myF8.Attributes = dbUpdatableField
Let myF8.DefaultValue = "Locomotive Facts"

Set myF9 = myTable.CreateField("LocomotiveImage", dbMemo)
Let myF9.AllowZeroLength = True
Let myF9.Attributes = dbUpdatableField
Let myF9.DefaultValue = ""

Set myF10 = myTable.CreateField("LocomotivePrototypeAdhesionFactor", dbMemo)
Let myF10.AllowZeroLength = True
Let myF10.Attributes = dbUpdatableField
Let myF10.DefaultValue = "0"

Set myF11 = myTable.CreateField("LocomotivePrototypeHorsePower", dbMemo)
Let myF11.AllowZeroLength = True
Let myF11.Attributes = dbUpdatableField
Let myF11.DefaultValue = "0"

Set myF12 = myTable.CreateField("LocomotivePrototypeTractionEffort", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotivePrototypeDrawBarPull", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"


Set myF12 = myTable.CreateField("LocomotivePrototypeWeight", dbMemo)
Let myF12.AllowZeroLength = True
Let myF12.Attributes = dbUpdatableField
Let myF12.DefaultValue = "0"

Set myF13 = myTable.CreateField("LocomotiveModelAdhesionFactor", dbMemo)
Let myF13.AllowZeroLength = True
Let myF13.Attributes = dbUpdatableField
Let myF13.DefaultValue = "0"

Set myF14 = myTable.CreateField("LocomotiveModelHorsePower", dbMemo)
Let myF14.AllowZeroLength = True
Let myF14.Attributes = dbUpdatableField
Let myF14.DefaultValue = "0"

Set myF15 = myTable.CreateField("LocomotiveModelTractionEffort", dbMemo)
Let myF15.AllowZeroLength = True
Let myF15.Attributes = dbUpdatableField
Let myF15.DefaultValue = "0"

Set myF16 = myTable.CreateField("LocomotiveModelDrawBarPull", dbMemo)
Let myF16.AllowZeroLength = True
Let myF16.Attributes = dbUpdatableField
Let myF16.DefaultValue = "0"


Set myF17 = myTable.CreateField("LocomotiveModelWeight", dbMemo)
Let myF17.AllowZeroLength = True
Let myF17.Attributes = dbUpdatableField
Let myF17.DefaultValue = "0"

myTable.Fields.Append myF1
myTable.Fields.Append myF2
myTable.Fields.Append myF3
myTable.Fields.Append myF4
myTable.Fields.Append myF5
myTable.Fields.Append myF6
myTable.Fields.Append myF7
myTable.Fields.Append myF8
myTable.Fields.Append myF9
myTable.Fields.Append myF10
myTable.Fields.Append myF11
myTable.Fields.Append myF12
myTable.Fields.Append myF13
myTable.Fields.Append myF14
myTable.Fields.Append myF15
myTable.Fields.Append myF16
myTable.Fields.Append myF17

NewDb.TableDefs.Append myTable

Set myRecord = NewDb.OpenRecordset("LocomotiveAddressShort")

For t = 1 To 127
    myRecord.AddNew
    myRecord.Fields("LocomotiveAddress") = Str$(t)
    myRecord.Update
    DoEvents
Next



NewDb.Close

End Sub
