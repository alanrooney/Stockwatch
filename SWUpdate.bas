Attribute VB_Name = "SWUpdate"
Option Explicit

''Public Function DoDatabaseUpdate()
''Dim rs As Recordset
''Dim Tbl As TableDef
''Dim iFlds As Integer
''Dim fld As Field
''
''Dim idx As Index, fldIndex As Field
''
''    ' This update added in ver 440
''
''
''    Screen.MousePointer = 11
''
''    On Error GoTo ErrorHandler
''
''    ' show message
''
''    ' make sure Stockwatch isnt running
''
''    On Error Resume Next
''
''    ' CREATE THE TblBARS TABLE
''
''    ' Create a new TableDef object.
''    Set Tbl = SWdb.CreateTableDef("tblBars")
''
''    Set fld = Tbl.CreateField("ID", dbLong)
''    fld.Attributes = fld.Attributes + dbAutoIncrField
''    Tbl.Fields.Append fld
''
''    Tbl.Fields.Refresh
''    Set idx = Tbl.CreateIndex("PrimaryKey")
''    Set fldIndex = idx.CreateField("ID", dbLong)
''    idx.Fields.Append fldIndex
''    idx.Primary = True
''    Tbl.Indexes.Append idx
''
''    ' Add the TableDef to the database.
''    SWdb.TableDefs.Append Tbl
''
''    Set Tbl = SWdb.TableDefs!tblBars
''    ' add tables required
''
''    Set fld = Tbl.CreateField("ClientID")
''    fld.Type = dbLong
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("Bar")
''    fld.Type = dbText
''    fld.Size = 20
''    Tbl.Fields.Append fld
''    Tbl.Fields("Bar").AllowZeroLength = True
''
''    ' CREATE THE TblBARCOUNT TABLE
''
''    ' Create a new TableDef object.
''    Set Tbl = SWdb.CreateTableDef("tblBarCount")
''
''    Set fld = Tbl.CreateField("ID", dbLong)
''    fld.Attributes = fld.Attributes + dbAutoIncrField
''    Tbl.Fields.Append fld
''
''    Tbl.Fields.Refresh
''    Set idx = Tbl.CreateIndex("PrimaryKey")
''    Set fldIndex = idx.CreateField("ID", dbLong)
''    idx.Fields.Append fldIndex
''    idx.Primary = True
''    Tbl.Indexes.Append idx
''    ' Add the TableDef to the database.
''    SWdb.TableDefs.Append Tbl
''
''    Set Tbl = SWdb.TableDefs!tblBarCount
''    ' add tables required
''
''    Set fld = Tbl.CreateField("ClientProdPLUID")
''    fld.Type = dbLong
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("ClientID")
''    fld.Type = dbLong
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("BarID")
''    fld.Type = dbInteger
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("BarFullQty")
''    fld.Type = dbDouble
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("BarOpen")
''    fld.Type = dbInteger
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("BarWeight")
''    fld.Type = dbSingle
''    Tbl.Fields.Append fld
''
''    Set fld = Tbl.CreateField("Verified")
''    fld.Type = dbBoolean
''    Tbl.Fields.Append fld
''
''    ' ADD chkMultipleBars to tblClients
''
''    Set Tbl = SWdb.TableDefs!tblClients
''    Set fld = Tbl.CreateField("chkMultipleBars")
''    fld.Type = dbBoolean
''    Tbl.Fields.Append fld
''
''    Screen.MousePointer = 0
''
''    Exit Function
''
''ErrorHandler:
''
''    MsgBox "Error " & Trim$(Error)
''    End
''
''End Function
''
