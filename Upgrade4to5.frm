VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D6BDA7&
   Caption         =   "Stockwatch Update"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1275
      TabIndex        =   0
      Top             =   855
      Width           =   1680
   End
   Begin VB.Label lblDone 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1905
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gbOk As Boolean
Public SWdb As Database
Public sDBLoc As String

Private Sub Command1_Click()
Dim Tbl As TableDef
Dim fld As Field
Dim colm As Object


'''        tblClientProductPLUs
'''
'''            GlassPrice      double
'''            GlassPriceDP        double
'''            GlassQty            integer
'''            GlassQtyDP      integer
'''
'''        tblPLUGroup
'''
'''            Glass           integer


'    Set Tbl = SWdb.TableDefs("tblTillDifference")
'    Set fld = Tbl.Fields("Difference")
'    fld.Name = "OldDiff"
'
'    Set Tbl = SWdb.TableDefs("tblTillDifference")
'    Set fld = Tbl.CreateField("Difference")
'    fld.Type = dbDouble
'    Tbl.Fields.Append fld
'
'    ' Copy contents from OldDiff to Difference
'    Dim rs As Recordset
'    Dim i As Long
'    Set rs = SWdb.OpenRecordset("tblTillDifference")
'    rs.Index = "PrimaryKey"
'        rs.MoveFirst
'    For i = 0 To rs.RecordCount - 1
'        rs.Edit
'        rs("Difference") = rs("OldDiff")
'        rs.Update
'        rs.MoveNext
'    Next
    
    ' THIS WORKS BUT SLOW
    
    Set Tbl = SWdb.TableDefs("tblClientProductPLUs")
    Set fld = Tbl.CreateField("GlassPrice")
    fld.Type = dbDouble
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblClientProductPLUs")
    Set fld = Tbl.CreateField("GlassPriceDP")
    fld.Type = dbDouble
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblClientProductPLUs")
    Set fld = Tbl.CreateField("GlassQty")
    fld.Type = dbInteger
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblPLUGroup")
    Set fld = Tbl.CreateField("Glass")
    fld.Type = dbInteger
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblClientProductPLUs")
    Set fld = Tbl.CreateField("GlassQtyDP")
    fld.Type = dbInteger
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblClientProductPLUs")
    Set fld = Tbl.CreateField("chkHistory")
    fld.Type = dbBoolean
    Tbl.Fields.Append fld
    
    
    Set Tbl = SWdb.TableDefs("tblDates")
    Set fld = Tbl.CreateField("Surplus")
    fld.Type = dbCurrency
    Tbl.Fields.Append fld
    
    Set Tbl = SWdb.TableDefs("tblDates")
    Set fld = Tbl.CreateField("SurplusTitle")
    fld.Type = dbText
    fld.Size = 25
    Tbl.Fields.Append fld
    
    ' Change tblTillDifference Difference data type from Integer to Double
    
    
    
    lblDone = "Done!"
    ' Add Field dtlCommPaid (Boolean)


End Sub

Private Sub Form_Load()
        
        sDBLoc = "" & GetSetting("Stockwatch", "DB", "StockwatchDB") & ""
        ' get the DB Location from the registry
        
        
        If gbOpenDB(Me) Then
        ' now open db
        
        End If

End Sub
Function gbOpenDB(mainfrm As Form) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' remove the extension since customer windows expl might not be showing extensions
    
    If sDBLoc = "" Then
        sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & ")", "Invalid Database Location: " & sDBLoc)
        If sDBLoc = "" Or sDBLoc = "exit" Then End
        SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
    End If
    
OpenDB:
'    Set ClientsDB = OpenDatabase("C:\KeyhouseMerge\fileDexTEST.mdb", False, False, ";PWD=wss1")

    Set SWdb = OpenDatabase("" & sDBLoc & "\Stockwatch.mdb", False, False, ";PWD=fran2012")
'    LogMsg frmSubMan, "DataBase Opened", " File: " & sDBLoc
    gbOpenDB = True

CleanExit:
    Exit Function
    
ErrorHandler:
    
    If Err = 3031 Then
        MsgBox "Not a Valid Password to open Database"
        End
    
    Else
    
        sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & "\" & App.Title & ".mdb)", "Invalid Database Location: " & sDBLoc)
    
        If sDBLoc = "" Or sDBLoc = "exit" Then End
        SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
        Resume OpenDB
    
    End If
    
    
End Function

