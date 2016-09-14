VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmMerge 
   BackColor       =   &H00C6DDD6&
   Caption         =   "Select PLU Description"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VSFlex8LCtl.VSFlexGrid grdGroups 
      Height          =   9345
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
      _cx             =   5106
      _cy             =   16484
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMerge.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   13440
      TabIndex        =   2
      Top             =   270
      Width           =   945
   End
   Begin VSFlex8LCtl.VSFlexGrid grdList 
      Height          =   9375
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   11745
      _cx             =   20717
      _cy             =   16536
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMerge.frx":0046
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label labelMsg 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   3150
      TabIndex        =   1
      Top             =   300
      Width           =   10005
   End
End
Attribute VB_Name = "frmMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lPLUID As Long
Public bAddItem As Boolean
Public sShowGroup As String

Private Sub Command1_Click()

End Sub

Private Sub cmdAdd_Click()
    bAddItem = True
    lPLUID = 0
    Unload Me

End Sub

Private Sub cmdRun_Click()
Dim rs As Recordset
Dim sDesc As String
Dim rsCl As Recordset
Dim rsPLU As Recordset
Dim sClient As String
Dim sPLUfile As String
Dim sSTOKfile As String
Dim rsPLUfile As Recordset
Dim iPLUGroup As Integer
Dim sPLUDesc As String
Dim iGap As Integer
Dim rsPLUfind As Recordset
Dim rsSTKfile As Recordset
Dim rsClientPLU As Recordset
Dim rsClientSTK As Recordset
Dim lClientID As Long
Dim lPLUNo As Long
Dim curPLUPrice As Currency
Dim lPLUID As Long
Dim bFirstTime As Boolean
Dim iFoundRow As Integer
Dim cnt As Integer
Dim iTot As Integer

    ' Read all client stock files and create the master PLU File
    
'    If MsgBox("Sure u want to run this program?", vbDefaultButton1 + vbYesNo + vbQuestion, "Maintenance program") = vbYes Then
    
    ' ask continue on or start again
        
        Screen.MousePointer = 11
        
        
        Set rsCl = SWdb.OpenRecordset("Select * FROM tblClients WHERE chkActive = true")
        rsCl.MoveFirst
        ' open Client Tbl
    
        Set rsPLU = SWdb.OpenRecordset("tblPLUs")
        rsPLU.Index = "PrimaryKey"
        
        bFirstTime = True
        Do
        ' Loop all Client records
        
'            If bFirstTime Then
'            '
'                lClientID = 6
'                sClient = "Freeneys_Bar"
'                ' Get Client Name
'
'
'            ElseIf rsCl("ID") = 6 Then
'            ' Freeneys bar so skip it
'
'                GoTo GetNextClient
'
'            Else
                ' For each client...
                lClientID = rsCl("ID") + 0
                sClient = Replace(Replace(rsCl("txtName") & "", " ", "_"), "'", "")
                ' Get Client Name
            
'            End If
            
            cnt = 0
            
            gbOk = GetGroups(lClientID)
            ' get all groups for this client
            
            
            DoEvents
            
            sPLUfile = "APLUFILE_" & sClient
            ' Concatinate with APLUFILE & 'name'
            
            Set rsPLUfile = SWdb.OpenRecordset("SELECT * FROM " & sPLUfile & " ORDER BY PLUDESCR")
            ' open this table
            ' create snapshot sorted by plu # ASC
            
            If Not (rsPLUfile.EOF And rsPLUfile.BOF) Then
                
                rsPLUfile.MoveLast
                iTot = rsPLUfile.RecordCount
                
                rsPLUfile.MoveFirst
                
                Do
                    cnt = cnt + 1
                    
                    sPLUDesc = UCase$(rsPLUfile("PLUDESCR") & "")
                    ' Get Description
                    
                    labelMsg.Caption = "Working on " & sClient & "              " & sPLUDesc & "            Item: " & Trim$(cnt) & " of " & Trim$(iTot)
                    DoEvents
                    
                    
                    iGap = InStr(sPLUDesc, " ")
                    ' Remove PLU group No from start
                    
                    iPLUGroup = Val(Left(sPLUDesc, iGap))
                    ' Get Group
                    
                    ' Match group no with one already stored
                    ' Search grdGroups for match on group no
                    ' if its there then ok, if not then ask for a description
                      
                    If grdGroups.FindRow(iPLUGroup, , 0) < 0 Then
                    
                        frmGroups.iGroupNo = iPLUGroup
                        frmGroups.lCLid = lClientID
                        frmGroups.Show vbModal
                    
                        gbOk = GetGroups(lClientID)
                        ' get all groups for this client
                    
                        If frmGroups.bStop Then GoTo stopfornow

                    
                    End If
                    
                    
                    
                    
                    sPLUDesc = Mid(sPLUDesc, iGap + 1, 300)
                    ' Get Description
                    
                    Set rsPLUfind = SWdb.OpenRecordset("SELECT * FROM tblPLUs WHERE txtDescription = """ & sPLUDesc & """")
                    ' Search The Master PLU table for a match on the description
            
                    If Not (rsPLUfind.EOF And rsPLUfind.BOF) Then
                    ' Found ...
            
                        grdList.TopRow = grdList.FindRow(rsPLU("ID"), , 0)
                                            
                    ElseIf FoundInOther(sPLUDesc, iFoundRow) Then
                            ' not found in other either
                            
                        grdList.TopRow = iFoundRow
                            
                    Else
                    
                            frmSelect.labelItem.Caption = UCase$(sPLUDesc)
                            
                            frmSelect.txtItem.Text = UCase$(sPLUDesc)
                            frmSelect.sFind = Left(UCase$(sPLUDesc), 1)
                            frmSelect.Show vbModal
                            
                            If frmSelect.bStop Then GoTo stopfornow
                            
                            If frmSelect.bAdd Then
                            ' Add it...
                            
                                ' first see if its there already..
                                Set rsPLUfind = SWdb.OpenRecordset("SELECT * FROM tblPLUs WHERE txtDescription = """ & frmSelect.sItem & """")
                                If Not (rsPLUfind.EOF And rsPLUfind.BOF) Then
                                
                                    rsPLUfind.MoveFirst
                                    rsPLUfind.Edit
                                    rsPLUfind("other") = rsPLUfind("other") & UCase$(sPLUDesc) & ", "
                                    rsPLUfind("chkActive") = True
                                    rsPLUfind.Update
                                    rsPLUfind.Bookmark = rsPLUfind.LastModified
                                    grdList.Cell(flexcpText, grdList.FindRow(rsPLUfind("ID"), , 0), 2) = UCase$(rsPLUfind("other")) & ""
                                    DoEvents
                            
                                Else
                                    
                                    rsPLU.AddNew
                                    rsPLU("txtDescription") = frmSelect.sItem
                                    If UCase$(frmSelect.sItem) <> UCase$(sPLUDesc) Then
                                        rsPLU("Other") = UCase$(sPLUDesc) & ", "
                                        grdList.AddItem rsPLU("ID") & vbTab & UCase$(frmSelect.sItem) & vbTab & UCase$(rsPLU("Other"))
                                    Else
                                        grdList.AddItem rsPLU("ID") & vbTab & UCase$(sPLUDesc)
                                
                                    End If
                                    
                                    rsPLU("chkActive") = True
                                    rsPLU.Update
                                    rsPLU.Bookmark = rsPLU.LastModified
                                
                                    grdList.RowData(grdList.Rows - 1) = rsPLU("ID") + 0
                                    DoEvents
                                
                                End If
                                
                            ElseIf frmSelect.lPLUID > 0 Then
                            ' Merge it...
                                
                                rsPLU.Seek "=", frmSelect.lPLUID
                                If Not (rsPLU.EOF And rsPLU.BOF) Then
                                    rsPLU.Edit
                                    rsPLU("other") = rsPLU("other") & UCase$(sPLUDesc) & ", "
                                    rsPLU("chkActive") = True
                                    rsPLU.Update
                                    rsPLU.Bookmark = rsPLU.LastModified
                                    grdList.Cell(flexcpText, grdList.FindRow(rsPLU("ID"), , 0), 2) = UCase$(rsPLU("other")) & ""
                                    DoEvents
                            
                                End If
                            
                            End If
                            
                            grdList.TopRow = grdList.FindRow(rsPLU("ID"), , 0)
                        
                        'End If
                      
                    End If
                    
                    rsPLUfile.MoveNext
                    
                    grdList.AutoSize 1, 2
                    grdList.Refresh
                    
                Loop While Not rsPLUfile.EOF
            
            End If
            
            rsPLUfile.Close
            ' close APLUFILE
    
GetNextClient:
            bFirstTime = False
            rsCl.MoveNext
        
        Loop While Not rsCl.EOF
        ' continue loop on other client files
        
            grdList.AutoSize 1, 2
            grdList.Refresh
            DoEvents
            
stopfornow:
        rsCl.Close

        Screen.MousePointer = 0
    
'    End If

End Sub

Private Sub Form_Activate()
    bAddItem = False

    
End Sub

Private Sub Form_Load()

    If gbOpenDB(Me) Then
        ' now open db

    
        gbOk = ShowPLUs()
    
    End If
    
    '    grdList.Row = grdList.FindRow(sShowGroup, 0)
    
End Sub
Public Function ShowPLUs()
Dim rs As Recordset
Dim sSql As String
    
        
    sSql = "WHERE chkActive = true"
    
    Screen.MousePointer = 11
    
    On Error GoTo ErrorHandler
    
    grdList.Rows = 1
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblPLUs " & sSql & " ORDER BY txtDescription", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("ID") & vbTab & rs("txtDescription") & vbTab & rs("Other")
            grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    grdList.AutoSize 1, 2
'        grdList.FrozenCols = 1

    ShowPLUs = True

CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    Screen.MousePointer = 0
    
    Exit Function

ErrorHandler:
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit


End Function

Private Sub grdList_Click()

    If grdList.Row > 0 Then
    
        If MsgBox("Delete " & grdList.Cell(flexcpTextDisplay, grdList.Row, 1), vbDefaultButton1 + vbYesNo + vbQuestion, "Delete?") = vbYes Then
            SWdb.Execute "DELETE FROM tblPLUs WHERE ID = " & Trim$(grdList.RowData(grdList.Row))
            grdList.RemoveItem grdList.Row
        End If
    End If
    
End Sub


Public Function FoundInOther(sDesc As String, iRow As Integer)

    On Error GoTo ErrorHandler
        
        For iRow = 1 To grdList.Rows - 1
            Debug.Print grdList.Cell(flexcpTextDisplay, iRow, 1)
'
'            Debug.Print grdList.Cell(flexcpTextDisplay, iRow, 2)
            
            If InStr(grdList.Cell(flexcpTextDisplay, iRow, 2), sDesc & ",") > 0 Then
                
                FoundInOther = True
                GoTo CleanExit
            End If
            
        Next
        
CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    MsgBox Trim$(Error)
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit

End Function


Public Function GetGroups(lCLid As Long)
Dim rs As Recordset

    grdGroups.Rows = 1
    
    Set rs = SWdb.OpenRecordset("Select * from tblPLUGroup WHERE ClientID = " & Trim$(lCLid) & " ORDER BY txtGroupNumber")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdGroups.AddItem rs("txtGroupNumber") & vbTab & rs("txtDescription")
            grdGroups.RowData(grdGroups.Rows - 1) = rs("ID") + 0
        
            rs.MoveNext
            
        Loop While Not rs.EOF
    End If
    
CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    MsgBox Trim$(Error)
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit

End Function
