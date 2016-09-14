Attribute VB_Name = "Generic"
Option Explicit
Public gbOk As Boolean
Public sngvatrate As Double
'Public bDebug As Boolean

Public endofpause As Double
Public lAUDf As Long
Public SWdb As Database
Public iErrCount As Integer
Public Const sBlack = "&H80000012"
Public Const sDarkBlue = &HFF0000
Public gbCnt As Integer
Public DBGf As Integer
Public lDatesID As Long
Public lSelClientID As Long
Public sDBLoc As String
Public WriteWord As Word.Application

Public bAllowMove As Boolean
Public MoveX As Integer
Public MoveY As Integer

'Public bPassGood As Boolean
Public NoLicense As Boolean
Public MsgReply As Integer
Public sMenuCtrl As String
Public sFranchiseEmail As String
Public gbRegion As String
Public gbSWEmail As String

Public gbSMTP As String
Public gbPort As Integer
Public gbEmailfromAddress As String
Public gbSSL As Integer
Public gbUsername As String
Public gbPassword As String
Public gbTestEnvironment As Boolean
Public bEvaluation As Boolean


Public Const sBlue = &HDECCB8
Public Const sLightBlue = &HF4EEE8
Public Const sGreen = &HD2DCCF
Public Const sLightGreen = &HEFF2EE
Public Const sOrange = &H92BFFC
Public Const sLightOrange = &HD8E8FE
Public Const sRed = &HC5DDFE
Public Const sWhite = &H80000005
Public Const sLightGrey = &H8000000A
Public Const sVryLtgGrey = &HE9EEF3
Public Const sYellow = &HAEF5FD
Public Const sLightYellow = &HDAFBFE
Public Const sDarkGrey = &H8000000C
Public Const iRepLines = 44 ' include header
Public Const sDarkRed = &H80&
Public Const sDarkGreen = &H4000&
Public Const sLightPurple = &HDFD7DF
Public Const sDarkPurple = &HAAA2AA
Public Const sKey = "Stockwatch Ireland Version 3.0"
Public Const sDMY = "dd/mm/yy"
Public Const sDMMYY = "dd mmm yyyy"


Public SW1 As Boolean   ' Stock Watch Head Office Flag
                        ' useful for allowing franchisees to update their own
                        ' added products (and not those originally setup by
                        ' head office

Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Public Declare Function SetLayeredWindowAttributes Lib "USER32" (ByVal hwnd As Long, ByVal color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public bDualPrice As Boolean
Public bMultipleBars As Boolean
Public Const iBoxWidth = 765

Private Declare Function getdrivetype Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub bHourGlass(bhow As Boolean)
    Select Case bhow
        Case True
         Screen.MousePointer = 11
    
        Case False
         Screen.MousePointer = 0
    End Select
    
End Sub

Public Sub ShowSplash(iWait As Integer)
    
    frmSplash.bShowSplash = True
    frmSplash.Show
    Pause iWait

End Sub

Public Function OpenAuditFile(tim As Date)

On Error GoTo ErrorHandler

    lAUDf = FreeFile
    Open CurDir & "\SW" & Trim$(Format(DateValue(Now) - (Format(tim, "w") - 2), "ddmmyy")) & ".CSV" For Append As #lAUDf
    OpenAuditFile = True
    ' use mondays date of this week for the file name
    
    Print #lAUDf, Format(Now, "dd/mm/yy hh:mm:ss") & "," & App.Title & " Started"
    ' send out a small header
    
    OpenAuditFile = True

    Exit Function

ErrorHandler:
  
End Function

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

    Set SWdb = OpenDatabase("" & sDBLoc & "\" & App.Title & ".mdb", False, False, ";PWD=fran2012")
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

Public Sub Pause(millisec)
Dim X As Integer

  endofpause# = Timer + millisec / 1000
  Do
      X% = DoEvents()
  Loop While Timer < endofpause#

End Sub

Public Sub LogMsg(frm As Form, sMsg As String, sAudMsg As String)

    On Error GoTo ErrorHandler
    
    If (sMsg & sAudMsg = "") Or sMsg <> "" Then
        frm.lblMsg.Caption = sMsg
        DoEvents
    
    End If
    ' display it on the form
    
    If sAudMsg <> "" Then
        Print #lAUDf, Format$(Now, "dd/mm/yy hh:mm:ss") & "," & Screen.ActiveForm.Name & "," & sMsg & "," & sAudMsg
    End If
    ' save it in the log file
    
    Exit Sub
    
ErrorHandler:
    Print #lAUDf, Format$(Now, "dd/mm/yy hh:mm:ss") & "," & sMsg & "," & sAudMsg
   
End Sub

Public Sub SetupHeaderField(fname As Form, sField As String)
Dim iCnt As Integer

    With fname
    
        If .grdList.Cols > 0 Then
            For iCnt = 0 To .grdList.Cols - 1
                If .grdList.ColKey(iCnt) = "" Then
                    Exit For
                End If
            Next
        End If
        ' point to next blank header field
        
        .grdList.Cols = iCnt + 1
        .grdList.Cell(flexcpText, 0, iCnt) = sField
        .grdList.ColKey(iCnt) = Trim$(Replace(Replace(sField, "&", ""), " ", ""))
        .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
        ' setup header
        
        Select Case sField
            
             Case "A"
             .grdList.ColWidth(iCnt) = 300
             .grdList.ColDataType(iCnt) = flexDTBoolean
            
            Case "Name"
             .grdList.ColWidth(iCnt) = 2500
             .grdList.ColAlignment(iCnt) = flexAlignLeftCenter
            
            Case "Join Date"
             .grdList.ColDataType(iCnt) = flexDTDate
            
            Case "Address"
             .grdList.ColWidth(iCnt) = 3000
             .grdList.ColAlignment(iCnt) = flexAlignLeftCenter
            
            Case "First Name", "Last Name"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Phone"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Contact"
             .grdList.ColWidth(iCnt) = 2500
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
             
            Case "Mobile"
             .grdList.ColWidth(iCnt) = 1500
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Email"
             .grdList.ColWidth(iCnt) = 3000
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Notes"
             .grdList.ColWidth(iCnt) = 1500
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Group ID"
             .grdList.ColWidth(iCnt) = 2000
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
        
            Case "Group", "Key Group"
             .grdList.ColWidth(iCnt) = 2000
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
        
            Case "Empty Weight"
             .grdList.ColWidth(iCnt) = 1500
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Full Verify", "Empty Verify"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
             .grdList.ColDataType(iCnt) = flexDTBoolean
            
            Case "Description"
             .grdList.ColWidth(iCnt) = 2700
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Key Description"
             .grdList.ColWidth(iCnt) = 2000
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "--"
             .grdList.ColWidth(iCnt) = 500
             
             Case "Active"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColDataType(iCnt) = flexDTBoolean
             
             Case "No"
             .grdList.ColWidth(iCnt) = 800
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Sell Price", "Cost"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
             
             Case "Sell 1", "Sell 2", "Gls 1", "Gls 2"
             .grdList.ColWidth(iCnt) = 900
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
             
'             Case "Code"
'             .grdList.ColWidth(iCnt) = 1000
'             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
'
'             Case "Cost"
'             .grdList.ColWidth(iCnt) = 1000
'             .grdList.ColAlignment(iCnt) = flexAlignRightCenter
        
             Case "Quantity", "PLU #"
             .grdList.ColWidth(iCnt) = 1000
             .grdList.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Open Item"
             .grdList.ColWidth(iCnt) = 1200
             .grdList.ColDataType(iCnt) = flexDTBoolean
             
             
        
        End Select
    End With

End Sub

Public Function SetColWidths(fname As Form, sCtrl As String, sCol As String, bAutoSize As Boolean)
Dim iScrollBarWidth As Integer

       If bAutoSize Then fname.Controls(sCtrl).AutoSize 0, fname.Controls(sCtrl).Cols - 1
        ' adjust all cols so contents are visible
        
        
        If fname.Controls(sCtrl).ColIsVisible(fname.Controls(sCtrl).Cols - 1) Then
            iScrollBarWidth = 150
        Else
            iScrollBarWidth = 300
        End If
    
        fname.Controls(sCtrl).ColWidth(fname.Controls(sCtrl).ColIndex(sCol)) = _
            (fname.Controls(sCtrl).Width - (fname.Controls(sCtrl).ColPos(fname.Controls(sCtrl).Cols - 1) + _
            fname.Controls(sCtrl).ColWidth(fname.Controls(sCtrl).Cols - 1) - _
            fname.Controls(sCtrl).ColWidth(fname.Controls(sCtrl).ColIndex(sCol)))) - iScrollBarWidth

End Function

Public Function CheckDBError(sSection As String)
  
  Dim endofpause As Double
  Dim errorloop As Error
  
  CheckDBError = True
  ' default to true
  ' only false when retries are exhausted
      
'  If bDebug Then
    LogMsg frmStockWatch, "", "error: " & Trim$(Error) & " in " & sSection
'  End If
  
  
  If (Err > 2999 And Err < 4000) Or Err = 75 Or Err = 55 Or Err = 57 Or Err = 71 Or Err = 76 Then
    ' if its a database error...
    iErrCount = iErrCount + 1
      
    If iErrCount = 5 Then
      
      CheckDBError = False
            
      Exit Function
      ' wait exhausted so leave routine with the bad news...
    End If
  
    endofpause# = Timer + 1
    
    Do
     
    
    Loop While Timer < endofpause#
    ' wait one
  
  Else
    SaveDebug sSection, "Error #" + Str$(Err) + " [" + Trim$(Error) + "] occured"
    
    MsgBox sSection & " Error #" & Str$(Err) & " [" & Trim$(Error) & "] occured"
    End
  End If

End Function

Public Function ReadDB(fname As Form, sTbl As String, lID As Long, ParamArray sFldValue())
Dim rs As Recordset
Dim iFld As Integer
Dim sOut As String

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tbl" & sTbl, dbOpenTable)
    rs.Index = "PrimaryKey"
    
    If lID <> 0 Then
    ' ID specified so its either an edit of a record
    ' or there is only one record in this table
      
      If lID > 0 Then
      ' edit record of passed ID
        rs.Seek "=", Trim$(lID)
      
      Else
      ' neg ID passed so there must be only one record to read
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            lID = rs("ID") + 0
            ' return with the right ID
            
        Else
            GoTo CleanExit
        ' not great prog but need to get out of here!
        End If
      End If
        
        If Not rs.NoMatch Then
    
            For iFld = 1 To Val(sFldValue(0))
                If TypeOf sFldValue(iFld) Is TextBox Then
                    sFldValue(iFld).Text = rs(sFldValue(iFld).Name) & ""
                
                ElseIf TypeOf sFldValue(iFld) Is Label Then
                    sFldValue(iFld).Caption = rs(sFldValue(iFld).Name) & ""
                
                ElseIf TypeOf sFldValue(iFld) Is ComboBox Then
                    If Not IsNull(rs(sFldValue(iFld).Name)) Then
                        If rs(sFldValue(iFld).Name).Type = 10 Or rs(sFldValue(iFld).Name).Type = 2 Then
                            
                            PointToEntry fname, sFldValue(iFld).Name, rs(sFldValue(iFld).Name), True
                        Else
                            PointToEntry fname, sFldValue(iFld).Name, rs(sFldValue(iFld).Name), False
                        End If
                    End If
                
                ElseIf TypeOf sFldValue(iFld) Is CheckBox Then
                    sFldValue(iFld).Value = Abs(rs(sFldValue(iFld).Name))
                
                ElseIf TypeOf sFldValue(iFld) Is OptionButton Then
                    sFldValue(iFld).Value = Abs(rs(sFldValue(iFld).Name))
                ElseIf TypeOf sFldValue(iFld) Is RichTextBox Then
                    sFldValue(iFld).Text = rs(sFldValue(iFld).Name)
                
                ElseIf TypeOf sFldValue(iFld) Is MyButton Then
                    sFldValue(iFld).Enabled = rs(sFldValue(iFld).Name)
                
                ElseIf TypeOf sFldValue(iFld) Is MaskEdBox Then
                    If Not IsNull(rs(sFldValue(iFld).Name)) Then
                        sFldValue(iFld).Text = rs(sFldValue(iFld).Name)
                    End If
                Else
                    MsgBox "Object Type not supported yet"
                End If
            Next
        
        Else
           ' bad news must report can't find record
        End If
        
    ElseIf Not rs.EOF Then
    
        rs.MoveFirst
        
        fname.Controls("grd" & sTbl).Rows = 0
        Do
            sOut = ""
            For iFld = 1 To Val(sFldValue(0))
                sOut = sOut & rs(sFldValue(iFld).Name)
                If iFld < Val(sFldValue(0)) Then
                    sOut = sOut & vbTab
                End If
            Next
            
            fname.Controls("grd" & sTbl).AddItem sOut
            fname.Controls("grd" & sTbl).RowData(fname.Controls("grd" & sTbl).Rows - 1) = rs("ID") + 0
            rs.MoveNext
        
        Loop While Not rs.EOF
    
    End If
    
    ReadDB = True
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ReadDB " & sTbl & " Form" & fname.Name & " " & Str$(lID)) Then Resume 0
    Resume CleanExit

End Function
Public Function WriteDB(fname As Form, sTbl As String, lID As Long, bRecOut As Boolean, ParamArray sFldValue())
Dim rs As Recordset
Dim iFld As Integer
Dim lTX As Long
Dim sOut As String
Dim sHdr As String
Dim sFile As String
Dim sCmd As String
Dim sFldIdx As String
Dim sShopList As String
Dim iShop As Integer
Dim sOldIndexValue As String

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tbl" & sTbl, , dbOpenTable)
    rs.Index = "PrimaryKey"
    
    If lID <> 0 Then
    ' updating a previous record....
    
      If lID > 0 Then
      ' edit record of passed ID
        rs.Seek "=", Trim$(lID)
        If Not rs.NoMatch Then
            rs.Edit
            sCmd = "UPD"
        End If
      
      Else
      ' neg ID passed so there must be only
      ' one record in the table
        
        If rs.RecordCount > 0 Then
        ' there's a record so we'll edit that one
            rs.MoveFirst
            rs.Edit
            sCmd = "UPD"
        Else
        ' no records so it must be first time in so we'll
        ' add a new one.
            rs.AddNew
            sCmd = "NEW"
        End If
        
      End If
    
    Else
    ' otherwise its a new record to be added
        rs.AddNew
        sCmd = "NEW"
    End If
    
    If bRecOut Then
        sFldIdx = GetIndexField(sTbl)
        ' Init the Out string with the index Field name

        sOut = rs(sFldIdx) & vbTab
        ' start out string with Index Field Name

    End If
    
    For iFld = 1 To Val(sFldValue(0))
        
       Select Case TypeName(sFldValue(iFld))
                
           Case "TextBox", "RichTextBox"
'            rs(sFldValue(iFld).Name) = Trim$(Replace(sFldValue(iFld).Text, vbCrLf, "|"))
            rs(sFldValue(iFld).Name) = Trim$(sFldValue(iFld).Text)
           
            sOut = sOut & Trim$(Replace(sFldValue(iFld).Text, vbCrLf, "|")) & vbTab
           
           Case "ComboBox", "ListBox", "Label"
'            rs(sFldValue(iFld).Name) = Trim$(Replace(sFldValue(iFld), vbCrLf, "|"))
            rs(sFldValue(iFld).Name) = Trim$(sFldValue(iFld))
           
            sOut = sOut & Trim$(sFldValue(iFld)) & vbTab
           
           Case "OptionButton", "CheckBox"
            rs(sFldValue(iFld).Name) = sFldValue(iFld).Value
           
            sOut = sOut & Trim$(sFldValue(iFld)) & vbTab
           
           Case "MaskEdBox"
            If Not IsNull(sFldValue(iFld)) Then
                rs(sFldValue(iFld).Name) = Val(sFldValue(iFld))
            End If
            
            sOut = sOut & Trim$(sFldValue(iFld)) & vbTab
           
           Case "MyButton"
            rs(sFldValue(iFld).Name) = sFldValue(iFld).Value

           
           Case Else
            If rs(sFldValue(iFld).Name).Type = 3 Or rs(sFldValue(iFld).Name).Type = 5 Or rs(sFldValue(iFld).Name).Type = 6 Then
            ' integer or Currency
                rs(sFldValue(iFld).Name) = Val(sFldValue(iFld))
            ElseIf rs(sFldValue(iFld).Name).Type = 8 Then
            ' string
                rs(sFldValue(iFld).Name) = Trim$(sFldValue(iFld))
            
            End If
        
            sOut = sOut & Trim$(sFldValue(iFld)) & vbTab
        
        End Select
    Next
    
    rs.Update
    rs.Bookmark = rs.LastModified
    lID = rs("ID")
    ' return ID of rec updated.

    LogMsg frmStockWatch, "Record Saved", "WriteDB:" & vbTab & sTbl & vbTab & sOut
    
    rs.Close

    WriteDB = True
    

CleanExit:
    'DBEngine.Idle dbRefreshCache
     ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If Err = 3022 Then
        rs("id") = rs("id") + 1
        Resume 0
    
    Else
    
        If CheckDBError("WriteDB") Then Resume 0
        Resume CleanExit
    End If
    
End Function
Public Sub bSetFocus(fname As Form, sCtrl As String)

    If fname.Controls(sCtrl).Enabled And fname.Controls(sCtrl).Visible Then
        fname.Controls(sCtrl).SetFocus
    End If

End Sub

Public Function PointToEntry(fname As Form, sCtrl As String, sEntry As String, bEntryIsText As Boolean)
 Dim lEntryID As Long

    If bEntryIsText Then
        For gbCnt = 0 To fname.Controls(sCtrl).ListCount - 1
            If UCase(Left(fname.Controls(sCtrl).List(gbCnt), Len(sEntry))) = UCase$(sEntry) Then
                fname.Controls(sCtrl).ListIndex = gbCnt
                PointToEntry = True
                Exit For
            End If
        
        Next
        
    Else
        lEntryID = Val(sEntry)
        If lEntryID > 0 Then
            For gbCnt = 0 To fname.Controls(sCtrl).ListCount - 1
                If fname.Controls(sCtrl).ItemData(gbCnt) = lEntryID Then
                    fname.Controls(sCtrl).ListIndex = gbCnt
                    PointToEntry = True
                    Exit For
                End If
            Next
        End If
    End If


End Function

Public Sub SaveDebug(sSect As String, dbgmsg As String)

    On Error GoTo ErrorHandler

    DBGf = FreeFile
    
    Open CurDir & "Debug.CSV" For Append As #DBGf
    'local debug file
 
    With Screen.ActiveForm
        Print #DBGf, Format(Now, "dd/mm/yy hh:mm:ss") & ",Form:" & .Name & ",Object:" & Screen.ActiveControl.Name & ", Routine: " & sSect & ", "; dbgmsg
    End With
    ' send out a small header

    Close #DBGf

Leave:
Exit Sub

ErrorHandler:
  
  MsgBox "Warning: Cannot save debug information"
  Resume Leave

End Sub

Public Function bSetupControl(fname As Form)
Dim iPreviousTab As Integer
Dim iTab As Integer
Dim sLastCtrl As String
    
    On Error GoTo ErrorHandler
    
    If Not fname.Controls(fname.ActiveControl.Name).Locked Then
    
        Select Case TypeName(fname.Controls(fname.ActiveControl.Name))
            Case "TextBox", "RichTextBox"
             If fname.Controls(fname.ActiveControl.Name).MultiLine Then
                 fname.Controls(fname.ActiveControl.Name).SelStart = Len(fname.Controls(fname.ActiveControl.Name).Text)
             Else
                 fname.Controls(fname.ActiveControl.Name).SelStart = 0
                 fname.Controls(fname.ActiveControl.Name).SelLength = Len(fname.Controls(fname.ActiveControl.Name).Text)
             End If
             
             fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
             
             fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
             
             
             If sLastCtrl <> "" Then
               
        ' grey if not filled
               
               fname.Controls("lbl" & Mid(sLastCtrl, 4, 50)).ForeColor = vbBlack
             End If
    
            Case "MaskEdBox", "DriveListBox"
             fname.Controls(fname.ActiveControl.Name).SelStart = 0
             fname.Controls(fname.ActiveControl.Name).SelLength = Len(fname.Controls(fname.ActiveControl.Name).Text)
             fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
    
             fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
             
             If sLastCtrl <> "" Then
                
        ' grey if not filled
                
                fname.Controls("lbl" & Mid(sLastCtrl, 4, 50)).ForeColor = vbBlack
             End If
            
            Case "HScrollBar", "ComboBox", "VSFlexGrid"
                fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
                fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
            
            Case ""
            Case Else
        End Select
        
             fname.lblGlow.Visible = True
             fname.lblGlow.Left = fname.ActiveControl.Left - 8
             fname.lblGlow.Top = fname.ActiveControl.Top - 8
             fname.lblGlow.Width = fname.ActiveControl.Width + 28
             fname.lblGlow.Height = fname.ActiveControl.Height + 30

'        fname.Controls(fname.ActiveControl.BackColor) = vbWhite
        
        sLastCtrl = fname.ActiveControl.Name
    
    End If
    
CleanExit:
    Exit Function
    
ErrorHandler:
    
    If Err = 340 Or Err = 730 Or Err = 438 Then
        Resume Next

        
    Else
        Resume CleanExit
    End If
    
End Function

Public Function GotoNextControl(frmname As Form, iStartCtl As Integer)
Dim iNextTabIndex As Integer
        
  On Error GoTo ErrorHandler
    
    If iStartCtl > 0 Then
        iNextTabIndex = iStartCtl
        
    ElseIf frmname.ActiveControl.TabIndex = Screen.ActiveForm.Count - 1 Then
        iNextTabIndex = iStartCtl
    Else
        iNextTabIndex = frmname.ActiveControl.TabIndex + 1
    End If
    ' if we're at the last control then point to the first control
    ' otherwise point to the next control
        
        
    On Error GoTo skipThisCtrl
GetNextControl:
    For gbCnt = 0 To Screen.ActiveForm.Count - 1
        
      If TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Skinner" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Line" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Image" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "CommonDialog" Then
         If frmname.Controls(gbCnt).TabIndex = iNextTabIndex Then
        
            If Screen.ActiveForm.Controls(gbCnt).Enabled And Screen.ActiveForm.Controls(gbCnt).Visible Then
            
                Select Case TypeName(Screen.ActiveForm.Controls(gbCnt))
                ' we only want to set focus to controls that will allow it
                
                    Case "ListBox", "TextBox", "ComboBox", "CommandButton", "MaskEdBox", "OptionButton", "RichTextBox", "VSFlexGrid", "DriveListBox", "CheckBox"
                     If Screen.ActiveForm.Controls(gbCnt).TabStop Then
                     ' make sure tab stop is enabled for the control...
                     
                        Screen.ActiveForm.Controls(gbCnt).SetFocus
                     Else
                         iNextTabIndex = iNextTabIndex + 1
                         GoTo GetNextControl
                     End If
                     Exit For
                    
                    Case Else
                     iNextTabIndex = iNextTabIndex + 1
                     GoTo GetNextControl
                     ' yugh! but it works
                     
                End Select
                    
            Else
                iNextTabIndex = iNextTabIndex + 1
                GoTo GetNextControl
            End If
        
        End If
      End If
      
NextCtrl:
    Next gbCnt
  
CleanExit:
    Exit Function
        
ErrorHandler:
    Resume CleanExit


skipThisCtrl:
    Resume NextCtrl
    
End Function

Public Function CharCheck(iChr As Integer)
    
    If (iChr > 47 And iChr < 58) Or (iChr > 64 And iChr < 91) Or (iChr > 96 And iChr < 123) Or iChr = 32 Or iChr = 39 Or iChr = 40 Or iChr = 41 Or iChr = 45 Or iChr = 8 Or iChr = 13 Or iChr = 27 Or iChr = 38 Or iChr = 22 Or iChr = 3 Or iChr = 24 Then CharCheck = True
        
End Function

Public Function GetIndexField(sTbl As String)

    Select Case sTbl
    
        Case "Clients"
         GetIndexField = "ID"

    End Select
    
End Function


Public Function GetGroups(fname As Form)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    With fname
    
        .cboGroups.Clear
    
        Set rs = SWdb.OpenRecordset("Select * FROM tblProductGroup WHERE chkActive = true ", dbOpenSnapshot)
        
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                .cboGroups.AddItem rs("txtDescription")
                .cboGroups.ItemData(.cboGroups.NewIndex) = rs("ID") + 0
                
                rs.MoveNext
            Loop While Not rs.EOF
        End If
        
    
    
        rs.Close
    
        GetGroups = True
    End With

CleanExit:
    'DBEngine.Idle dbRefreshCache
     ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetGroups") Then Resume 0
    Resume CleanExit
    

End Function

Public Function GetGroup(lGroupID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblProductGroup")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lGroupID
    If Not rs.NoMatch Then
        GetGroup = rs("txtDescription")
    End If
        
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
     ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetGroup") Then Resume 0
    Resume CleanExit
    

End Function

Public Function ShowCount(fname As Form, sGrd As String)

    ShowCount = fname(sGrd).Rows - 1

End Function
Public Function GetVatRate(sCode As String) As Double
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblVat", dbOpenTable)
    rs.Index = "vatCode"
    rs.Seek "=", sCode
    If Not rs.NoMatch Then
        GetVatRate = rs("txtRate")
    End If
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetVatRate ") Then Resume 0
    Resume CleanExit

End Function

Public Sub SetupCountField(fname As Form, sFld1 As String, sFld2 As String)
Dim iCnt As Integer

    With fname
    
        If .grdCount.Cols > 0 Then
            For iCnt = 0 To .grdCount.Cols - 1
                If .grdCount.ColKey(iCnt) = "" Then
                    Exit For
                End If
            Next
        End If
        ' point to next blank header field
        
        .grdCount.Cols = iCnt + 1
        
        If .grdCount.Rows > 0 Then
            .grdCount.Cell(flexcpText, 0, iCnt) = sFld1
            If sFld2 <> "" Then
                .grdCount.Rows = 2
                .grdCount.FixedRows = 2
                .grdCount.Cell(flexcpText, 1, iCnt) = sFld2
            
            End If
        End If
        
        .grdCount.ColKey(iCnt) = Trim$(Replace(Replace(sFld1 & sFld2, "&", ""), " ", ""))
        .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        ' setup header
        
        Select Case sFld1 & " " & sFld2
            Case "Name"
             .grdCount.ColWidth(iCnt) = 2500
             .grdCount.ColAlignment(iCnt) = flexAlignLeftCenter
            
            Case "Join Date"
             .grdCount.ColDataType(iCnt) = flexDTDate
            
            Case "Address"
             .grdCount.ColWidth(iCnt) = 3000
             .grdCount.ColAlignment(iCnt) = flexAlignLeftCenter
            
            Case "First Name", "Last Name"
             .grdCount.ColWidth(iCnt) = 1200
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Phone"
             .grdCount.ColWidth(iCnt) = 1200
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Contact"
             .grdCount.ColWidth(iCnt) = 2500
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
             
            Case "Mobile"
             .grdCount.ColWidth(iCnt) = 1500
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Email"
             .grdCount.ColWidth(iCnt) = 3000
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Notes"
             .grdCount.ColWidth(iCnt) = 1500
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Group ID"
             .grdCount.ColWidth(iCnt) = 2000
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
            Case "Group"
             .grdCount.ColWidth(iCnt) = 2400
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
            Case "Empty Weight"
             .grdCount.ColWidth(iCnt) = 1500
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "Full Verify", "Empty Verify"
             .grdCount.ColWidth(iCnt) = 1200
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
             .grdCount.ColDataType(iCnt) = flexDTBoolean
            
            Case "Description"
             .grdCount.ColWidth(iCnt) = 2000
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
            
            Case "----"
             .grdCount.ColWidth(iCnt) = 800
             
             Case "Active"
             .grdCount.ColWidth(iCnt) = 1200
             .grdCount.ColDataType(iCnt) = flexDTBoolean
             
             Case "Stock Connection"
             .grdCount.ColWidth(iCnt) = 2500
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
        
             Case "No"
             .grdCount.ColWidth(iCnt) = 1000
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Sell Price"
             .grdCount.ColWidth(iCnt) = 1400
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Quantity"
             .grdCount.ColWidth(iCnt) = 1600
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Code"
             .grdCount.ColWidth(iCnt) = 1400
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Size", "Full Qty", "Open items"
             .grdCount.ColWidth(iCnt) = 1400
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "Weight"
             .grdCount.ColWidth(iCnt) = 1800
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
             
             Case "Delivery Note"
             .grdCount.ColWidth(iCnt) = 2800
             .grdCount.ColAlignment(iCnt) = flexAlignCenterCenter
        
             Case "<  Total  >"
             .grdCount.ColWidth(iCnt) = 2200

        
            Case "Key Description"
             .grdCount.ColWidth(iCnt) = 4000
             .grdCount.ColAlignment(iCnt) = flexAlignLeftCenter
        
            Case " Value"
             .grdCount.ColWidth(iCnt) = 2000
             .grdCount.FixedAlignment(iCnt) = flexAlignCenterCenter
             .grdCount.ColAlignment(iCnt) = flexAlignRightCenter

             Case " Sel "
             .grdCount.ColWidth(iCnt) = 1000
             .grdCount.ColDataType(iCnt) = flexDTBoolean
            
             Case "Diff"
             .grdCount.ColWidth(iCnt) = 1000
             .grdCount.FixedAlignment(iCnt) = flexAlignRightCenter

            
            
'ver440 out for now
'            Case "Chk"
'            .grdCount.ColWidth(iCnt) = 1000
'            .grdCount.ColDataType(iCnt) = flexDTBoolean
'            .grdCount.FixedAlignment(iCnt) = flexAlignLeftCenter
        End Select
    End With

End Sub

Public Function RepDeliveries()
Dim rs As Recordset
Dim dbTotal As Double
Dim sDelNote As String

    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    frmStockWatch.grdCount.ForeColor = sBlack
    
    frmStockWatch.grdCount.Rows = 1
    frmStockWatch.grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "Item", ""
    SetupCountField frmStockWatch, "Size", ""
    SetupCountField frmStockWatch, "Delivery Note", ""
    SetupCountField frmStockWatch, "Qty", ""
    SetupCountField frmStockWatch, "Rate", ""
    SetupCountField frmStockWatch, "Value", ""

        
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

 
    Set rs = SWdb.OpenRecordset("SELECT * FROM (tblDeliveries INNER JOIN tblClientProductPLUs ON tblDeliveries.ClientProdPLUID = tblClientProductPLUs.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND (Quantity <> 0 or Free <> 0) ORDER BY ref, cboGroups, txtDescription", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            
'Ver 2.2.0
' The cost needs to be shown from the tblClientProductPLUs table
            
            If rs("Quantity") <> 0 Then
'                frmStockWatch.grdCount.AddItem rs("Ref") & " " & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & rs("DeliveryNote") & vbTab & rs("Quantity") & vbTab & Format(rs("cost"), "0.00") & vbTab & Format(rs("Quantity") * rs("cost"), "0.00")
                frmStockWatch.grdCount.AddItem rs("Ref") & " " & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & rs("DeliveryNote") & vbTab & rs("Quantity") & vbTab & Format(rs("purchasePrice"), "0.00") & vbTab & Format(rs("Quantity") * rs("PurchasePrice"), "0.00")
'                dbTotal = dbTotal + rs("Quantity") * rs("Cost")
                dbTotal = dbTotal + rs("Quantity") * rs("PurchasePrice")
            
' ver 2.3
                If rs("Free") <> 0 Then
                    frmStockWatch.grdCount.AddItem rs("Ref") & " " & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & "Free" & vbTab & rs("Free") & vbTab & "0.00" & vbTab & "0.00"
'                    dbTotal = dbTotal + rs("Quantity") * rs("Cost")
                End If
                
'            Else
'                frmStockWatch.grdCount.AddItem rs("Ref") & " " & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & "Free" & vbTab & rs("Free") & vbTab & Format(0, "0.00") & vbTab & Format(0, "0.00")
            End If
            
            frmStockWatch.grdCount.RowData(frmStockWatch.grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    
        frmStockWatch.grdCount.AddItem ""
        frmStockWatch.grdCount.AddItem vbTab & vbTab & "Total Value:"
        frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Value")) = Format(dbTotal, "Currency")
    
        frmStockWatch.grdCount.AddItem ""
        
        frmStockWatch.grdCount.AutoSize 0, 5
    
    End If
    
    gbOk = SetReportSize("")
    
    bHourGlass False
    
    RepDeliveries = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepDeliveries") Then Resume 0
    Resume CleanExit

End Function

Public Function RepStockAnalysis()
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbSellExVat As Double
Dim dbProfitMargin As Double
Dim dbSalesExVat As Double
Dim dbCostOfSales As Double
Dim dbGrossProfit As Double
Dim dbRetailValue As Double
Dim dbClosingValueTotal As Double
Dim dbRetailValueTotal As Double
Dim dbSalesExVatTotal As Double
Dim dbCostOfSalesTotal As Double
Dim dbGrossProfitTotal As Double
Dim dbSalesExVatGrandTotal As Double
Dim dbClosingValueGrandTotal As Double

Dim iRow As Integer

Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double

Dim curCost As Currency

Dim lLastProd As Long
Dim dbCalcAmount As Double

Dim dbFullQty As Double
Dim dbLastQty As Double

'ver530
Dim iIssue As Integer
Dim dbGlassExVat As Double
Dim dbGlassProfitMargin As Double
Dim dbGlassDPExVat As Double
Dim dbGlassDPProfitMargin As Double

Dim dbTotProfitMargin As Double
Dim dbTotSellExVat As Double
    
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With frmStockWatch
    
    .grdCount.ForeColor = sBlack
    
    .grdCount.Rows = 1
    .grdCount.Cols = 0
    
    .grdCount.Left = .picStatus.Width + 100
    .grdCount.Top = .picStatus.Top + 630
    .grdCount.Width = .Width - .picStatus.Width - 200
    .grdCount.Height = .Height - .picStatus.Top - 700

    
    SetupCountField frmStockWatch, "", "Description"
    SetupCountField frmStockWatch, "", "Size"
    SetupCountField frmStockWatch, "Last", "Cost"
    
    SetupCountField frmStockWatch, "Sell", "In-Vat"
'    SetupCountField frmStockWatch, "Profit", "Margin"
    SetupCountField frmStockWatch, "Sell", "P'cent"

' ver530
' added glass here and dual pricing

    SetupCountField frmStockWatch, "Glass", "In-Vat"
'    SetupCountField frmStockWatch, "Gls Prf", "Margin"
    SetupCountField frmStockWatch, "Gls", "P'cent"
    
    
    'DP
    If bDualPrice Then
    
        SetupCountField frmStockWatch, "Sell 2", "In-Vat"
'        SetupCountField frmStockWatch, "Profit 2", "Margin"
        SetupCountField frmStockWatch, "Sell 2", "P'cent"
' ver530
' added glass here
        
        SetupCountField frmStockWatch, "Gls 2", "In-Vat"
'        SetupCountField frmStockWatch, "Gls2 Prf", "Margin"
        SetupCountField frmStockWatch, "Gls2", "P'cent"
    
    End If
    
    SetupCountField frmStockWatch, "Avg GP", "P'cent"
    ' average GP Percentage margin
    
    SetupCountField frmStockWatch, "Open'g", "Stock"
    SetupCountField frmStockWatch, "Delivr", "Total"
    SetupCountField frmStockWatch, "Delivr", "Free"
    SetupCountField frmStockWatch, "Closng", "Stock"
    SetupCountField frmStockWatch, "Closng", "Value"
    SetupCountField frmStockWatch, "", "Sales"
    
' Ver530 removed
'    SetupCountField frmStockWatch, "Units", "Sold"
'    SetupCountField frmStockWatch, "Retail", "Value"
    
    .grdCount.FixedCols = 2
    
    ' Report by Group
    
    ' Description Size Last Avg  Sell   Sell   Profit G.P.  Open  Deliv Mvts   Close Close       Units Retail  Sales   Cost of  Gross  Sales
    '           Cost Cost Ex-Vat In-Vat Margin p'cnt Stock eries In/Out Stock Value Sales Sold  Value   Ex-Vat  Sales    Profit p'cnt

    ' Group totals
    
    ' Value of Closing Stock Ex-Vat
    
    ' sql get records by Client ID
    
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

    
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lSelClientID) & " AND Active = true Order By cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        
        Do

'Debug.Print rs("tblProductGroup.txtDescription")


'            If rs("tblProductGroup.txtDescription") = "Whiskies" Then Stop
            If iLastGroup <> rs("cboGroups") Then
                
                .grdCount.AddItem ""
                .grdCount.AddItem ">>>  " & rs("cboGroups") & "   " & rs("tblProductGroup.txtDescription")
                iLastGroup = rs("cboGroups")
                .grdCount.AddItem ""
            
                dbClosingValueGrandTotal = dbClosingValueGrandTotal + dbClosingValueTotal
                dbClosingValueTotal = 0
                dbRetailValueTotal = 0
                dbSalesExVatGrandTotal = dbSalesExVatGrandTotal + dbSalesExVatTotal
                dbSalesExVatTotal = 0
                dbCostOfSalesTotal = 0
                dbGrossProfitTotal = 0


            End If
            
            iIssue = getGlasses(rs("PLUGroupID"))
            
            If lLastProd <> rs("tblProducts.ID") Then
            
                lLastProd = rs("tblProducts.ID")
                
'                If rs("tblProducts.ID") = 2064 Then Stop
                
                If Not IsNull(rs("FullQty")) Then dbFullQty = rs("FullQty") Else dbFullQty = 0
                If Not IsNull(rs("LastQty")) Then dbLastQty = rs("LastQty") Else dbLastQty = 0
                
'                If rs("tblProducts.txtDescription") = "Black & White New" Then Stop
                
                .grdCount.AddItem rs("tblProducts.txtDescription")
                
                .grdCount.RowData(.grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
        
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Description")) = rs("tblProducts.txtDescription")
                ' Item
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Size")) = rs("txtSize")
                ' Size
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("LastCost")) = Format(rs("PurchasePrice"), "0.00")
                ' Last Cost
                
                        
                ' SELL EXVAT INVAT PROFIT1MARGIN GP1pCent
'ver530
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("SellIn-Vat")) = Format(rs("SellPrice"), "0.00")
                ' Sell In-Vat
                
                dbSellExVat = Format(rs("SellPrice") / (1 + (sngvatrate / 100)), "0.00")
                dbGlassExVat = 0
                If Not IsNull(rs("GlassPrice")) Then
                    dbGlassExVat = Format(rs("GlassPrice") / (1 + (sngvatrate / 100)), "0.00")
                End If
                ' Sell Ex-Vat
                  
                dbProfitMargin = dbSellExVat - rs("PurchasePrice") / rs("txtIssueUnits")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ProfitMargin")) = Format(dbProfitMargin, "0.00")
                ' Profit Margin
            
                If dbSellExVat > 0 Then
                    dbTotProfitMargin = dbProfitMargin
                    dbTotSellExVat = dbSellExVat
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("SellP'cent")) = Format(dbProfitMargin, "0.00")
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sellp'cent")) = Format((dbProfitMargin / dbSellExVat) * 100, "0.00")
                
                End If
                ' G.P. p'Cent
                
'                If rs("tblProducts.txtDescription") = "Bulmers LN" Then Stop
                
                If rs("GlassPrice") > 0 Then
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GlassIn-Vat")) = Format(rs("GlassPrice"), "0.00")
                    ' Glass In-Vat
                End If
                'ver 543
                  
                  
                If iIssue > 0 And dbGlassExVat > 0 Then
                    dbGlassProfitMargin = dbGlassExVat - rs("PurchasePrice") / (rs("txtIssueUnits") * iIssue)
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GlsPrfMargin")) = Format(dbGlassProfitMargin, "0.00")
                    ' Glass Profit Margin
                  
                    dbTotProfitMargin = dbTotProfitMargin + dbGlassProfitMargin
                    dbTotSellExVat = dbTotSellExVat + dbGlassExVat
                    
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GlsP'cent")) = Format((dbGlassProfitMargin / dbGlassExVat) * 100, "0.00")
                    
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GlsP'cent")) = Format(dbGlassProfitMargin, "0.00")
                    ' G.P. p'Cent
                 
                End If
'DP
                If bDualPrice Then
                    
                    ' SELL2 EXVAT2 INVAT2 PROFIT2MARGIN GP2pCent
                    
                    If Not IsNull(rs("SellPriceDP")) Then
                        
                        .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sell2In-Vat")) = Format(rs("SellPriceDP"), "0.00")
                        ' Sell2 In-Vat
                        
                        If rs("SellPriceDP") > 0 Then
                            dbSellExVat = Format(rs("SellPriceDP") / (1 + (sngvatrate / 100)), "0.00")
                            dbGlassDPExVat = 0
                            If Not IsNull(rs("GlassPriceDP")) Then
                                dbGlassDPExVat = Format(rs("GlassPriceDP") / (1 + (sngvatrate / 100)), "0.00")
                            End If
                            
                        End If
                        ' Sell 2 Ex-Vat
                        
                    
                        dbProfitMargin = dbSellExVat - rs("PurchasePrice") / (rs("txtIssueUnits"))
'                        .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Profit2Margin")) = Format(dbProfitMargin, "0.00")
                        ' Profit Margin (utilising same variables as above to work it out)
                    
                        If dbSellExVat > 0 Then
                            dbTotProfitMargin = dbTotProfitMargin + dbProfitMargin
                            dbTotSellExVat = dbTotSellExVat + dbSellExVat
                            
'                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sell2P'cent")) = Format(dbProfitMargin, "0.00")
                        
                        
                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sell2P'cent")) = Format((dbProfitMargin / dbSellExVat) * 100, "0.00")
                        
                        
                        End If
                        ' G.P. p'Cent
                        
                        If rs("GlassPriceDP") > 0 Then
                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Gls2In-Vat")) = Format(rs("GlassPriceDP"), "0.00")
                            ' Glass 2 In-Vat
                        End If
                        'ver 543
                        
                        If dbGlassDPExVat > 0 Then
                            dbGlassProfitMargin = dbGlassDPExVat - rs("PurchasePrice") / (rs("txtIssueUnits") * iIssue)
'                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Gls2PrfMargin")) = Format(dbGlassProfitMargin, "0.00")
                            ' Glass Profit Margin (utilising same variables as above to work it out)
                            
                            dbTotProfitMargin = dbTotProfitMargin + dbGlassProfitMargin
                            dbTotSellExVat = dbTotSellExVat + dbGlassDPExVat

'                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Gls2P'cent")) = Format(dbGlassProfitMargin, "0.00")
                            
                            
                            .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Gls2P'cent")) = Format((dbGlassProfitMargin / dbGlassDPExVat) * 100, "0.00")
                            
                            ' G.P. p'Cent
                        End If
                    
                    End If
                End If
                  
                If dbSellExVat > 0 Then
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("AvgGPp'cent")) = Format((dbTotProfitMargin / dbTotSellExVat) * 100, "0.00")
                    ' G.P. p'Cent
                End If


'
'                If iDivNumber > 0 Then
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("AvgGPp'cent")) = Format(((GPpercent + GPGlasspercent + GPDPpercent + GPGlassDPpercent) / iDivNumber), "0.00")
'                End If
                


'                If dbSellExVat > 0 Then
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sellp'cent")) = Format((dbProfitMargin / dbSellExVat) * 100, "0.00")
'                    ' G.P. p'Cent
'                End If
            
'                If dbGlassExVat > 0 And iIssue > 0 Then
'                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Glsp'cent")) = Format((dbGlassProfitMargin / dbGlassExVat) * 100, "0.00")
'                    ' G.P. p'Cent
'                End If
    
    'ver530
                  
          
'DP
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Open'gStock")) = Format(dbLastQty, "0.00")
                ' Opening Stock
                
                
                ' DELIVERIES
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
    '            lDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"))
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("DelivrTotal")) = dblDeliveries
                ' Deliveries
                
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("DelivrFree")) = dblFreeDel
                ' Free (replaces the Mvts In/Out column
                ' Mvts In/Out
                
                dbCalcAmount = CalcAmount(dbFullQty, rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngStock")) = Format(dbCalcAmount, "0.00")
                ' Closing Stock
            
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngValue")) = Format(dbCalcAmount * rs("PurchasePrice"), "0.00")
                dbClosingValueTotal = dbClosingValueTotal + (dbCalcAmount * rs("PurchasePrice"))
                ' Closing Value
            
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("Sales")) = Format((dbLastQty + dblDeliveries) - dbCalcAmount, "0.00")
                ' Sales
                
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("UnitsSold")) = Format((dbLastQty + dblDeliveries - dbCalcAmount) * rs("txtIssueUnits"), "0")
                ' Units Sold
                
'                dbRetailValue = Format((((dbLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("RetailValue")) = Format(dbRetailValue * rs("SellPrice"), "0.00")
'                dbRetailValueTotal = dbRetailValueTotal + (dbRetailValue * rs("SellPrice"))
                ' Retail Value
'Debug.Print dbRetailValueTotal
                    
            End If
            
            rs.MoveNext
        
            If rs.EOF Then
            
                .grdCount.AddItem ""
                .grdCount.AddItem ""
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("DelivrFree")) = "Group"
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngStock")) = "Totals >"
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngValue")) = Format(dbClosingValueTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("RetailValue")) = Format(dbRetailValueTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("SalesEx-Vat")) = Format(dbSalesExVatTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CostOfSales")) = Format(dbCostOfSalesTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GrossProfit")) = Format(dbGrossProfitTotal, "0.00")
                            
                dbSalesExVatGrandTotal = dbSalesExVatGrandTotal + dbSalesExVatTotal
                dbClosingValueGrandTotal = dbClosingValueGrandTotal + dbClosingValueTotal

            
            ElseIf iLastGroup <> rs("cboGroups") Then
                    
                .grdCount.AddItem ""
                .grdCount.AddItem ""
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("DelivrFree")) = "Group"
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngStock")) = "Totals >"
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("ClosngValue")) = Format(dbClosingValueTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("RetailValue")) = Format(dbRetailValueTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("SalesEx-Vat")) = Format(dbSalesExVatTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CostOfSales")) = Format(dbCostOfSalesTotal, "0.00")
'                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GrossProfit")) = Format(dbGrossProfitTotal, "0.00")
            
            End If
        
        
        Loop While Not rs.EOF
    End If
    
    ' PUT IN Sales % Figures Now
    
'    For iRow = 1 To .grdCount.Rows - 1
'        If .grdCount.RowData(iRow) > 0 Then
'        ' check for data row
'
'
'            If dbSalesExVatGrandTotal > 0 And Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("SalesEx-Vat"))) > 0 Then
'                .grdCount.Cell(flexcpText, iRow, .grdCount.ColIndex("SalesP'cnt")) = Format((.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("SalesEx-Vat"))) / (dbSalesExVatGrandTotal) * 100, "0.00")
'            End If
'        End If
'
'    Next
        
'Ver 2.1.1
    
    iRow = 1
    Do
        If .grdCount.RowData(iRow) > 0 Then
        ' check for data row

'If .grdCount.RowData(iRow) = 169726 Then Stop

'            If Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("Open'gStock"))) + _
'                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("ClosngStock"))) + _
'                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("DelivrFree"))) + _
'                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("Sales"))) + _
'                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("UnitsSold"))) = 0 Then
            If Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("Open'gStock"))) + _
                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("ClosngStock"))) + _
                Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("DelivrFree"))) = 0 Then
                
                
'-----------------------------------------------
' Rev 553
' Delivery total also check for zero
' This is as a result of Kates email:12/05/2016
                If Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("DelivrTotal"))) = 0 Then
'-----------------------------------------------
                    .grdCount.RemoveItem iRow
                
                Else
                    iRow = iRow + 1
                End If
                
            Else
                iRow = iRow + 1
            End If

        Else
            iRow = iRow + 1
        End If
    
    Loop While Not iRow = .grdCount.Rows - 1
    
    .grdCount.ScrollBars = flexScrollBarBoth


' Ver 2.1.0
'    .grdCount.AutoSize 0, 1

    .grdCount.AutoSize 0, .grdCount.Cols - 1
    .grdCount.ColWidth(.grdCount.ColIndex("DelivrFree")) = 400


    .grdCount.AddItem ""
    .grdCount.AddItem "Value of Closing Stock is:" ' & vbTab & Format(dbClosingValueGrandTotal, "Currency") & vbTab & "(Ex Vat)"
    .grdCount.AddItem Format(dbClosingValueGrandTotal, "Currency") & " (Ex Vat)"
    ' put final figure here after autosize
    
    .grdCount.AddItem ""
    
    
    End With
    
    bHourGlass False
    
    RepStockAnalysis = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepStockAnalysis") Then Resume 0
    Resume CleanExit

End Function

Public Function GetDeliveries(lID As Long, dblDeliveriesThatCost As Double, dblFreeDeliveries As Double, curCostDeliveries As Currency)
Dim rs As Recordset

    dblDeliveriesThatCost = 0
    dblFreeDeliveries = 0
    curCostDeliveries = 0
    
    On Error GoTo ErrorHandler
    
'Ver 2.2.0

'    Set rs = SWdb.OpenRecordset("SELECT SUM(Quantity) AS TotalQty, SUM(Free) AS FreeQty, Max(Cost) as FullCost FROM tblDeliveries WHERE ClientProdPLUID = " & Trim$(lID), dbOpenSnapshot)
    Set rs = SWdb.OpenRecordset("SELECT SUM(Quantity) AS TotalQty, SUM(Free) AS FreeQty FROM tblDeliveries WHERE ClientProdPLUID = " & Trim$(lID), dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        If Not IsNull(rs("TotalQty")) Then
            
            
            dblDeliveriesThatCost = rs("TotalQty")
            dblFreeDeliveries = rs("FreeQty")
            
'Ver 2.2.0
'            curCostDeliveries = rs("FullCost")
            curCostDeliveries = GetPurchasePrice(lID)
            
            
            GetDeliveries = dblDeliveriesThatCost + dblFreeDeliveries + 0
                        
        End If
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetDeliveries") Then Resume 0
    Resume CleanExit

End Function

Public Function CalcAmount(varFull As Variant, _
                            varOpen As Variant, _
                            varCombWt As Variant, _
                            iFullWt As Integer, _
                            iEmptyWt As Integer)

    If IsNull(varFull) Then varFull = 0

    If Not IsNull(varOpen) And Not IsNull(varCombWt) Then
    
        If varOpen > 0 And varCombWt > 0 And iEmptyWt > 0 Then

            CalcAmount = Format(varFull + ((varCombWt - (varOpen * iEmptyWt))) / (iFullWt - iEmptyWt), "0.00")
    
            ' No Full + ( Combined Wt - (No Open * Empty Wt)) / (Full Wt - Empty Wt)
    
    
        Else
            CalcAmount = varFull
        End If
    Else
        CalcAmount = varFull
    End If
    

End Function

Public Function RepTillReconciliation(sHistory As String)
Dim rs As Recordset
Dim rsDiff As Recordset
Dim rsHist As Recordset
Dim iLastGroupNo As Integer
Dim sGroupName As String
Dim dblStockSales As Double
Dim sngTillTotal As Double
Dim sngStockTotal As Double
Dim sDiffTotal As String
Dim iRow As Integer
Dim iLastPLUNo As Integer
Dim iCol As Integer
Dim lLastGroup As Long
Dim dbAddTotal As Double
Dim iGroupTotalHistCols As Integer  ' Ver 405
Dim sngSalesTotQty As Double
Dim lSalesDP As Long
'ver 531
Dim sngGlass As Double
Dim sngGlassDP As Double
Dim iGlass As Integer
Dim sMeasure As String

    ' iAddtotal was set to long to cater for Burkes - group total > 32767
    ' then set to dbl to cater for fractions
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    frmStockWatch.grdCount.ForeColor = sBlack
    
    frmStockWatch.grdCount.Rows = 1
    frmStockWatch.grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "", "Description"
    SetupCountField frmStockWatch, "Till", "Sales"
    SetupCountField frmStockWatch, "Stock", "Sales"
    SetupCountField frmStockWatch, "", "Diff"
    SetupCountField frmStockWatch, "", "%"

    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

'Ver 549   - had to reinclude active flag here
'ver548
' Remove the Active flag so that while products may be deactivated if they have history they will still show in the report.
    

' Ver550 use chkHistory checkbox to see whats going to be shown.
'    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) LEFT JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND Active = true ORDER BY txtGroupNumber, tblPLUs.txtDescription", dbOpenSnapshot)
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) LEFT JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE ((tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & ") AND (chkHistory = true)) ORDER BY txtGroupNumber, tblPLUs.txtDescription", dbOpenSnapshot)
'------------------------------------------

'    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) LEFT JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " ORDER BY txtGroupNumber, tblPLUs.txtDescription", dbOpenSnapshot)
'----------
'----------
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        iLastGroupNo = rs("txtGroupNumber")
        sGroupName = rs("tblPLUGroup.txtDescription")
        ' for the 1st one
        
        Do
            
           
            If iLastPLUNo <> Val(rs("PLUNumber")) Then
                
                iLastPLUNo = Val(rs("PLUNumber"))
                
                dblStockSales = GetStockSales(rs("PLUNumber"))
                
                
                'DP
                sngSalesTotQty = 0    ' Init
                lSalesDP = 0
                sngGlass = 0
                sngGlassDP = 0
            
            
                If Not IsNull(rs("SalesQty")) Then
                    sngSalesTotQty = rs("SalesQty")
                End If
            
                If Not IsNull(rs("SalesQtyDP")) Then
                    lSalesDP = rs("SalesQtyDP")
                End If
                
                sMeasure = "0"
                
                If Not IsNull(rs("Glass")) Then
                    iGlass = rs("Glass")
                
                    If iGlass > 0 Then
                        
' ver 556 ----------------------------------------
' remove zeros
'                        sMeasure = "0.00"

                        sMeasure = ""
'-------------------------------------------------
                        
                        If Not IsNull(rs("GlassQty")) Then
                            sngGlass = rs("GlassQty") / iGlass
                        End If
                
                        If Not IsNull(rs("GlassQtyDP")) Then
                            sngGlassDP = rs("GlassQtyDP") / iGlass
                        End If
                    End If
                End If
                
                sngSalesTotQty = sngSalesTotQty + lSalesDP + sngGlass + sngGlassDP
                ' Total Sales units

'                If rs("tblPLUs.txtDescription") = "SCOTCH" Then Stop
                    
                frmStockWatch.grdCount.AddItem rs("tblPLUs.txtDescription") & vbTab & sngSalesTotQty & vbTab & Format(dblStockSales, sMeasure)
        
                'DP
                
                
                ' Ver 550 zero fill diff before hand
                
                frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Diff")) = "0"
                '
                If dblStockSales <= sngSalesTotQty Then
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Diff")) = Trim$(Format(sngSalesTotQty - dblStockSales, sMeasure))
            
                Else
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Diff")) = "-" & Trim$(Format(dblStockSales - sngSalesTotQty, sMeasure))
                
                End If
        
                frmStockWatch.grdCount.RowData(frmStockWatch.grdCount.Rows - 1) = rs("tblPLUs.ID") + 0
                
                frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = rs("PLUGroupID") + 0
                ' save the group ID here for counting group totals later
                
                If Not sngSalesTotQty = 0 Then
                    sngTillTotal = sngTillTotal + sngSalesTotQty
                End If
                
                sngStockTotal = sngStockTotal + dblStockSales
                
            End If
    
            If sHistory <> "" Then
            
                ' CREATE HISTORY COLUMNS HERE
                
                Set rsHist = SWdb.OpenRecordset("SELECT [To] FROM tblTillDifference INNER JOIN tblDates ON tblTillDifference.DatesID = tblDates.ID WHERE ClientID = " & Trim$(lSelClientID) & " ORDER BY [To] DESC", dbOpenSnapshot)
                If Not (rsHist.EOF And rsHist.BOF) Then
                    rsHist.MoveFirst
    
                    Do
  
                        ' CREATE HISTORY COLUMNS HERE
    
                        If sHistory = " All" Or (frmStockWatch.grdCount.Cols - 5 < Val(sHistory)) Then
    
                            If frmStockWatch.grdCount.ColIndex(DateValue(rsHist("To"))) = -1 Then
                            ' date column not found... so add it
    
    
                                frmStockWatch.grdCount.Cols = frmStockWatch.grdCount.Cols + 1
                                frmStockWatch.grdCount.ColKey(frmStockWatch.grdCount.Cols - 1) = DateValue(rsHist("To"))
                                frmStockWatch.grdCount.Cell(flexcpText, 1, frmStockWatch.grdCount.Cols - 1) = Format(rsHist("to"), "dd/mm")
                                
                                frmStockWatch.grdCount.ColAlignment(frmStockWatch.grdCount.Cols - 1) = flexAlignCenterCenter
    
                            End If
    
                            
                        End If
    
                        rsHist.MoveNext
                    
                    Loop While Not rsHist.EOF
    
                End If
                
                
                
                Set rsDiff = SWdb.OpenRecordset("SELECT [To], [From], [Difference] FROM tblTillDifference INNER JOIN tblDates ON tblTillDifference.DatesID = tblDates.ID WHERE PLUID = " & Trim$(frmStockWatch.grdCount.RowData(frmStockWatch.grdCount.Rows - 1)) & " AND ClientID = " & Trim$(lSelClientID) & " ORDER BY [To] DESC", dbOpenSnapshot)
                If Not (rsDiff.EOF And rsDiff.BOF) Then
                    rsDiff.MoveFirst
    
                    'Ver 550
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, 5, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.Cols - 1) = "0"
                    ' Zero fill the history cells
                    
                    Do
  
                        If frmStockWatch.grdCount.ColIndex(DateValue(rsDiff("To"))) <> -1 Then
                            
                            frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex(DateValue(rsDiff("To")))) = rsDiff("Difference")
                        End If
                        
                        rsDiff.MoveNext
                    
                    Loop While Not rsDiff.EOF
    
                End If
            
            
            End If
           
            
            rs.MoveNext
        
            If Not rs.EOF Then
            
                If iLastGroupNo <> rs("txtGroupNumber") Then
                    iLastGroupNo = rs("txtGroupNumber")
                
                    If sngTillTotal > sngStockTotal Then
                        sDiffTotal = Trim$(Format(sngTillTotal - sngStockTotal, sMeasure))
                    ElseIf sngTillTotal < sngStockTotal Then
                        sDiffTotal = "-" & Trim$(Format(sngStockTotal - sngTillTotal, sMeasure))
                    Else
                        sDiffTotal = "0"
                    End If
                    
                    frmStockWatch.grdCount.AddItem ""
                    If sngStockTotal <> 0 Then
                        frmStockWatch.grdCount.AddItem "Group  " & sGroupName & " Total " & vbTab & Trim$(Format(sngTillTotal, sMeasure)) & vbTab & Trim$(Format(sngStockTotal, sMeasure)) & vbTab & sDiffTotal & vbTab & Format(Abs((Val(sDiffTotal) / sngStockTotal) * 100), "0.00")
                        frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = "GroupTotal"
                    
                    ElseIf frmStockWatch.grdCount.Cols > 5 Then
                    
                        For iGroupTotalHistCols = 5 To frmStockWatch.grdCount.Cols - 1

                            If Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 2, iGroupTotalHistCols)) <> 0 Then
                                frmStockWatch.grdCount.AddItem "Group  " & sGroupName & " Total "
                                frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = "GroupTotal"
                                Exit For
                            End If
                        Next
                    
                    End If
                    
                    frmStockWatch.grdCount.AddItem ""
            
                    sGroupName = rs("tblPLUGroup.txtDescription")
                    
                    sngTillTotal = 0
                    sngStockTotal = 0
                    sDiffTotal = ""
                    
                End If
            
            Else
                    If sngTillTotal > sngStockTotal Then
                        sDiffTotal = Trim$(sngTillTotal - sngStockTotal)
                    ElseIf sngTillTotal < sngStockTotal Then
                        sDiffTotal = "-" & Trim$(sngStockTotal - sngTillTotal)
                    Else
                        sDiffTotal = "0"
                    End If
                    
                    frmStockWatch.grdCount.AddItem ""
                    If sngStockTotal <> 0 Then
                        frmStockWatch.grdCount.AddItem "Group  " & sGroupName & " Total " & vbTab & Trim$(sngTillTotal) & vbTab & Trim$(Format(sngStockTotal, sMeasure)) & vbTab & Trim$(Format(sDiffTotal, sMeasure)) & vbTab & Format(Abs((Val(sDiffTotal) / sngStockTotal) * 100), "0.00")
                        frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = "GroupTotal"
                    
                    ElseIf frmStockWatch.grdCount.Cols > 5 Then
                    
                        If Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 2, 5)) <> 0 Then
                            frmStockWatch.grdCount.AddItem "Group  " & sGroupName & " Total "
                            frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = "GroupTotal"
                        End If
                        
                    End If
                    
                    frmStockWatch.grdCount.AddItem ""

                    frmStockWatch.grdCount.AddItem ""
            
            
                    sngTillTotal = 0
                    sngStockTotal = 0
                    sDiffTotal = ""
            
            End If
            
doloop:
' this line label was used for debug

        Loop While Not rs.EOF
    
    
    End If
    rs.Close
    
    
    With frmStockWatch
    
'Ver 2.1.1
 'TEST TEST
    iRow = 1

  'Ver550
  If .grdCount.Rows > 2 Then
  '------
  
    Do
      
        If .grdCount.RowData(iRow) > 0 Then
        ' check for data row

            If (Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("TillSales"))) = 0) And _
                (Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("StockSales"))) = 0) _
            Then
                If .grdCount.Cols > 5 Then
                    If Val(.grdCount.Cell(flexcpTextDisplay, iRow, 5)) = 0 Then
                        .grdCount.RemoveItem iRow
                    Else
                        iRow = iRow + 1
                    End If
                    ' see if last audit had zero as well
                Else
                    .grdCount.RemoveItem iRow
                ' no previous audit so remove line
                End If

            Else
                iRow = iRow + 1
            End If

        Else
            iRow = iRow + 1
        End If
      
    Loop While iRow <> .grdCount.Rows - 1
    
    
'Ver 2.4
    
    ' Add group totals for Previous Stock Takes
    
    For iCol = 5 To .grdCount.Cols - 1
    
        For iRow = 2 To .grdCount.Rows - 1
    
            If .grdCount.Cell(flexcpData, iRow, 0) = "" And .grdCount.RowData(iRow) = "" Then
            ' skip blank row
            
            ElseIf .grdCount.Cell(flexcpData, iRow, 0) = "GroupTotal" Then
            ' show total...
                .grdCount.Cell(flexcpText, iRow, iCol) = dbAddTotal
                dbAddTotal = 0
            
            ElseIf lLastGroup <> .grdCount.Cell(flexcpData, iRow, 0) Then
            ' start off...
                lLastGroup = .grdCount.Cell(flexcpData, iRow, 0)
                dbAddTotal = Val(.grdCount.Cell(flexcpTextDisplay, iRow, iCol))
    
            ElseIf lLastGroup = .grdCount.Cell(flexcpData, iRow, 0) Then
            ' add them...
                dbAddTotal = dbAddTotal + Val(.grdCount.Cell(flexcpTextDisplay, iRow, iCol))
    
            End If
    
        Next
    
    Next
    
    
    
    bHourGlass False
    
'==============
    
    .grdCount.Cell(flexcpAlignment, 2, 3, .grdCount.Rows - 1, 3) = flexAlignRightCenter

  End If
  
  End With
  
    RepTillReconciliation = True
    
    frmStockWatch.grdCount.AutoSize 0, frmStockWatch.grdCount.Cols - 1

    frmStockWatch.grdCount.FrozenCols = 5
        
    gbOk = SetReportSize("Till")



CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepDeliveries") Then Resume 0
    Resume CleanExit

    
End Function

Public Function GetStockSales(iPLUNumber As Integer) As Double
Dim rs As Recordset
Dim dblDeliveries As Double
Dim dblStockSales As Double
'Dim lStockSales As Long
Dim dbCalcAmount As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double
Dim curCost As Currency
Dim dblLastQty As Double

' Ver543 change lstocksales to dblstocksales
    
    On Error GoTo ErrorHandler
    
' ver 554  needed to include a check for active flag as well     as per email re: Ger and Athenry Golf Club
'    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.PLUNumber = " & Trim$(iPLUNumber) & " AND tblClientProductPLUs.ClientID = " & Trim$(lSelClientID), dbOpenSnapshot)
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.PLUNumber = " & Trim$(iPLUNumber) & " AND tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblClientProductPLUs.Active = true", dbOpenSnapshot)
'-------------------------------------------------------------------------------------
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
            
'            If Not IsNull(rs("FullQty")) And Not IsNull(rs("LastQty")) Then
'               lStockSales = lStockSales + (rs("LastQty") + dblDeliveries - dbCalcAmount) * rs("txtIssueUnits")
'            End If
'Ver 4.0.0
  
            If Not IsNull(rs("FullQty")) Then

                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    dblStockSales = dblStockSales + (dblLastQty + dblDeliveries - dbCalcAmount) * rs("txtIssueUnits")

            
' ver 544 - round values
'ver 545 - change to instr function to cater for 'Draught Beers'

                    
                    If InStr(rs("tblPLUGroup.txtDescription"), "Draught") > 0 Then

                        If (dblStockSales - Int(dblStockSales)) < 0.25 Then

                            dblStockSales = Int(dblStockSales)
                        ElseIf (dblStockSales - Int(dblStockSales)) < 0.5 Then
                            dblStockSales = Int(dblStockSales) + 0.5
                        ElseIf (dblStockSales - Int(dblStockSales)) < 0.75 Then
                            dblStockSales = Int(dblStockSales) + 0.5
                        Else
                            dblStockSales = Int(dblStockSales) + 1
                        End If

'''' ver 546 - round except wine

                    ElseIf InStr(rs("tblPLUGroup.txtDescription"), "Wine") = 0 Then
                        ' leave wine as is
                        ' round everything else
                        dblStockSales = Round(dblStockSales)

                    End If
                
            End If
            rs.MoveNext
        Loop While Not rs.EOF
        
        GetStockSales = dblStockSales
        
    End If
    
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetStockSales") Then Resume 0
    Resume CleanExit

End Function

Public Function RepClosingStock()
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbSellExVat As Double
Dim dbProfitMargin As Double
Dim dbSalesExVat As Double
Dim dbCostOfSales As Double
Dim dbGrossProfit As Double
Dim dbRetailValue As Double

Dim lLastProd As Long
Dim dbCalcAmount As Double
Dim dbTotal As Double
Dim dbFullQty As Double


    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With frmStockWatch
    
    .grdCount.ForeColor = sBlack
    
    
    .grdCount.Rows = 1
    .grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "", "Code"
    SetupCountField frmStockWatch, "", "Item"
    SetupCountField frmStockWatch, "", "Size"
    SetupCountField frmStockWatch, "Closing", "Stock"
    SetupCountField frmStockWatch, "Unit", "Cost"
    SetupCountField frmStockWatch, "Close", "Value"
    
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lSelClientID) & " AND Active = true Order By cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
            If Not IsNull(rs("FullQty")) Then dbFullQty = rs("FullQty") Else dbFullQty = 0
            
            
'Ver530 -----------------------------------------------------------------
' only report it if theres a closing value. Cut down on the report length
            If dbFullQty > 0 Or rs("Open") > 0 Then
'------------------------------------------------------------------------
               
               If iLastGroup <> rs("cboGroups") Then
                   .grdCount.AddItem ""
                   .grdCount.AddItem vbTab & ">>>  " & rs("cboGroups") & "   " & rs("tblProductGroup.txtDescription")
                   iLastGroup = rs("cboGroups")
                   .grdCount.AddItem ""
               End If
    
    
               If lLastProd <> rs("tblProducts.ID") Then
           
                   lLastProd = rs("tblProducts.ID")
               
                   dbCalcAmount = CalcAmount(dbFullQty, rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                   .grdCount.AddItem rs("txtCode") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtSize") & vbTab & Format(dbCalcAmount, "0.00") & vbTab & Format(rs("PurchasePrice"), "0.00") & vbTab & Format(dbCalcAmount * rs("PurchasePrice"), "0.00")
                   .grdCount.RowData(.grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
                   dbTotal = dbTotal + dbCalcAmount * rs("PurchasePrice")
               
               End If
            End If
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
        
        

    .grdCount.AddItem ""
    .grdCount.AddItem "Total" & vbTab & "(Ex-Vat)"
    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CloseValue")) = Format(dbTotal, "Currency")
    .grdCount.AddItem ""
    
    .grdCount.Cell(flexcpAlignment, 1, 1, .grdCount.Rows - 1, 1) = flexAlignCenterCenter
    .grdCount.Cell(flexcpAlignment, 1, 0, .grdCount.Rows - 1, 0) = flexAlignLeftCenter
    .grdCount.ScrollBars = flexScrollBarBoth
    .grdCount.AutoSize 1, 5

    End With
    
'    frmStockWatch.grdCount.Width = frmStockWatch.grdCount.ColPos(frmStockWatch.grdCount.Cols - 1) + frmStockWatch.grdCount.ColWidth(frmStockWatch.grdCount.Cols - 1) + 300
''    frmStockWatch.grdCount.Left = frmStockWatch.picStatus.Width + (frmStockWatch.Width - frmStockWatch.picStatus.Width - frmStockWatch.grdCount.Width) / 2
'    frmStockWatch.grdCount.Height = frmStockWatch.Height - frmStockWatch.picSelect.Height - frmStockWatch.picSelect.Top - 1200
'    frmStockWatch.grdCount.Top = (frmStockWatch.Height - frmStockWatch.picStatus.Top - frmStockWatch.grdCount.Height) / 2 + 1600
    
    gbOk = SetReportSize("")
    
    bHourGlass False
    
    RepClosingStock = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepClosingStock") Then Resume 0
    Resume CleanExit


End Function


Public Function CharOk(iAsc As Integer, i012 As Integer, sExp As String)
Dim iCnt As Integer

    Select Case i012
        Case 0  ' No Only
         If (iAsc > 47 And iAsc < 58) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
         
        Case 1  ' Alpha Only
         If (iAsc > 64 And iAsc < 91) Or (iAsc > 96 And iAsc < 123) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
         
        Case 2  ' Either Only
         If (iAsc > 47 And iAsc < 58) Or (iAsc > 64 And iAsc < 91) Or (iAsc > 96 And iAsc < 123) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
    End Select

End Function

Public Function DatesCheck(sFrom As String, sTo As String)

    If IsDate(sFrom) Then
        If IsDate(sTo) Then
            If DateValue(sFrom) <= DateValue(Now) Then
                If DateValue(sTo) <= DateValue(Now) Then
                    If DateValue(sFrom) <= DateValue(sTo) Then
                        
                        DatesCheck = True
            
                    Else
                        MsgBox "From Date must be Later than To date"
                        bSetFocus frmStockWatch, "tedFrom"
                    End If
                Else
                    MsgBox "To Date is in the Future"
                    bSetFocus frmStockWatch, "tedTo"
                End If
            Else
                MsgBox "From Date is in the Future"
                bSetFocus frmStockWatch, "tedFrom"
            End If
        Else
            MsgBox "Please Select valid To Date"
            bSetFocus frmStockWatch, "tedTo"
        End If
    Else
        MsgBox "Please Select valid From Date"
        bSetFocus frmStockWatch, "tedFrom"
    End If

End Function

Public Function GetClientName(lCLId As Long, sAddr As String, bRemoveCRLF As Boolean) As String
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClients")
    rs.Index = "ID"
    rs.Seek "=", lCLId
    If Not rs.NoMatch Then
                    
        If bRemoveCRLF Then
            sAddr = Replace(rs("rtfAddress"), vbCrLf, " ")
        Else
            sAddr = rs("rtfAddress") & ""
        End If
        GetClientName = rs("txtName") & ""
                    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientName") Then Resume 0
    Resume CleanExit


End Function

Public Function MenuCheck(sMenu As String)

    If sMenu = "1" Or frmStockWatch.grdMenu.Tag = "True" Then
    ' ok to begin new stock take or
    ' access all other buttons on menu
    
        MenuCheck = True
    
    End If

End Function
Public Function CreateClientFolder(sFolder As String, sName As String, sDate As String)
Dim objfile As Object
Dim sClientFolder As String

    On Error GoTo ErrorHandler
    
    sClientFolder = sDBLoc & "\" & Trim$(sName)
    Set objfile = CreateObject("Scripting.FileSystemObject")
    If Not objfile.FolderExists(sClientFolder) Then
    ' check for existance of new client name folder
    ' if it doesnt exists then create it ...
        objfile.CreateFolder (sClientFolder)

    End If
                    
    sFolder = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate)
    Set objfile = CreateObject("Scripting.FileSystemObject")
    If Not objfile.FolderExists(sFolder) Then
    ' check for existance of new client name folder
    ' if it doesnt exists then create it ...
        objfile.CreateFolder (sFolder)

    End If
    
    CreateClientFolder = True
    
CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("CreateClientFolder") Then Resume 0
    Resume CleanExit

End Function

Public Function VerifyClientFolderFile(sFile As String, sName As String, sDate As String, sRep As String) As Boolean
Dim objfile As Object
Dim sClientFolder As String
Dim sDateFolder As String
    On Error GoTo ErrorHandler
    
    sClientFolder = sDBLoc & "\" & Trim$(sName)
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    
    If Not objfile.FolderExists(sClientFolder) Then
    ' check for existance of client name folder
    ' if it doesnt exists then error
        
        GoTo CleanExit

    End If
                    
                    
    sDateFolder = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate)
    
    If Not objfile.FolderExists(sDateFolder) Then
    ' check for existance of client name folder
    ' if it doesnt exists then error
        GoTo CleanExit

    End If
    
    
    sFile = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\" & Trim$(sRep)
    
    
    
    
    If Not objfile.FileExists(sFile) Then
    ' check for existance of client name folder
    ' if it doesnt exists then error
        
        ' ver 542
        ' if its product deficit report then try
        '   summary analysis report instead
        
        If sRep = "Product Deficit.Doc" Then
        
            sRep = "Summary Analysis.Doc"
            
        
            sFile = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\" & Trim$(sRep)
        
            If Not objfile.FileExists(sFile) Then
        
                GoTo CleanExit
            End If
            
        Else
            GoTo CleanExit
        End If
    
    End If
    
    
    VerifyClientFolderFile = True
    
CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("VerifyClientFolderFile") Then Resume 0
    Resume CleanExit

End Function

Public Function RepGroupTotals()
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbSellExVat As Double
'Dim dbProfitMargin As Double
Dim dbSalesExVat As Double
Dim dbCostOfSales As Double
Dim dbGrossProfit As Double
Dim dbRetailValue As Double
Dim dbClosingValueTotal As Double
Dim dbRetailValueTotal As Double
Dim dbSalesExVatTotal As Double
Dim dbCostOfSalesTotal As Double
Dim dbGrossProfitTotal As Double

Dim dbCalcSalesGrand As Double
Dim dbCostOfSalesGrand As Double
Dim dbGrossProfitGrand As Double
'Dim dbTotalGPPerCent As Double
Dim dbCalcAmount As Double

Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double

Dim curCost As Currency

Dim lLastProd As Long
Dim iRow As Integer
Dim iGrpCount As Integer
Dim dblLastQty As Double
Dim curFix As Currency
Dim iIssue As Integer

    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With frmStockWatch
    
    .grdCount.ForeColor = sBlack
    
    .grdCount.Rows = 1
    .grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "", "Group"
    SetupCountField frmStockWatch, "Calc Sales", "Value"
    SetupCountField frmStockWatch, "% of", "Total"
    SetupCountField frmStockWatch, "Cost of", "Sales"
    SetupCountField frmStockWatch, "% of", "Total"
    SetupCountField frmStockWatch, "Gross", "Profit"
    SetupCountField frmStockWatch, "G.P.", "P'cent"
    
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

    dbCalcSalesGrand = 0
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lSelClientID) & " Order By cboGroups, PLUnumber, tblProducts.ID, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        iGrpCount = 1
        
        Do
            
            If iLastGroup <> rs("cboGroups") Then
                
                .grdCount.AddItem rs("tblProductGroup.txtDescription")
                ' Show group name
                
                iLastGroup = rs("cboGroups")
                
                dbClosingValueTotal = 0
                dbRetailValueTotal = 0
                dbGrossProfitTotal = 0
                
                dbSalesExVatTotal = 0
                dbCostOfSalesTotal = 0
                dbGrossProfitTotal = 0

            End If
            
            If lLastProd <> rs("tblProducts.ID") Then

                lLastProd = rs("tblProducts.ID")
                
                dbSellExVat = Format(rs("SellPrice") / (1 + (sngvatrate / 100)), "0.00")
                ' Sell Ex-Vat
                
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
                ' Deliveries
                
'ver 560
'                If Not IsNull(rs("FullQty")) Then

                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    
                    dbRetailValue = Format((((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
                    
                    dbRetailValueTotal = dbRetailValueTotal + (dbRetailValue * rs("SellPrice"))
                    ' Retail Value
                
                
    
'ver 560
'                    Dim lLastPLU As Long
'                    If rs("plunumber") <> lLastPLU Then
'                        lLastPLU = rs("plunumber")
                    
                        ' VER 530 - Glass measures --------------

                        iIssue = getGlasses(rs("PLUGroupID"))

                        If iIssue > 0 And dbRetailValue > 0 Then

                            Dim curSellPrice As Currency
                            Dim curGlassPrice As Currency
                            Dim curSellPriceDP As Currency
                            Dim curGlassPriceDP As Currency
                            Dim lGlassQty As Long
                            Dim lGlassQtyDP As Long
                            Dim lSalesQtyDP As Long
    
                        ' 25Jun removed this already one above
                        '    iIssue = getGlasses(rs("PLUGroupID"))
                            If IsNull(rs("SellPrice")) Then curSellPrice = 0 Else curSellPrice = rs("SellPrice")
                            If IsNull(rs("GlassPrice")) Then curGlassPrice = 0 Else curGlassPrice = rs("GlassPrice")
                            If IsNull(rs("SellPriceDP")) Then curSellPriceDP = 0 Else curSellPriceDP = rs("SellPriceDP")
                            If IsNull(rs("GlassPriceDP")) Then curGlassPriceDP = 0 Else curGlassPriceDP = rs("GlassPriceDP")
    
                            If IsNull(rs("GlassQty")) Then lGlassQty = 0 Else lGlassQty = rs("GlassQty")
                            If IsNull(rs("GlassQtyDP")) Then lGlassQtyDP = 0 Else lGlassQtyDP = rs("GlassQtyDP")
                            If IsNull(rs("SalesQtyDP")) Then lSalesQtyDP = 0 Else lSalesQtyDP = rs("SalesQtyDP")

                            dbRetailValueTotal = dbRetailValueTotal + GlassPriceValueFix(iIssue, curSellPrice, curGlassPrice, curSellPriceDP, curGlassPriceDP, lGlassQty, lGlassQtyDP, lSalesQtyDP)
                                                                                                       ' pint           glass           pint 2nd price      glass 2nd price     glass qty       glass qty @ 2nd price    Pint 2nd qty
                        End If

                        '----------------------------------------


                    ' ver 452 This added here 5th MAr
                        If bDualPrice And dbRetailValue > 0 Then

'                            Dim lLastPLU As Long
'                            If rs("plunumber") <> lLastPLU Then
'                               lLastPLU = rs("plunumber")
                               dbRetailValueTotal = dbRetailValueTotal + DualPriceValueFix(rs("SellPrice"), rs("SellPriceDP"), rs("SalesQtyDP"))
'                            End If
                        End If
'                    End If
                    
                    
                    dbSalesExVat = Format(((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits"), "0") * dbSellExVat
                    dbSalesExVatTotal = dbSalesExVatTotal + dbSalesExVat
                    ' Sales Ex-Vat
            
                    dbCostOfSales = (dblLastQty * rs("PurchasePrice")) + (dblCostDel * curCost) - (dbCalcAmount * rs("PurchasePrice"))
                    dbCostOfSalesTotal = dbCostOfSalesTotal + dbCostOfSales
                    ' Cost of Sales
            
                    
' ver 452 This removed here 5th MAr
         
     '               dbCalcSalesGrand = dbCalcSalesGrand + (dbRetailValue * rs("SellPrice"))
                    
                    dbCostOfSalesGrand = dbCostOfSalesGrand + dbCostOfSales
'
'               End If

          End If
            
            rs.MoveNext
            
            If rs.EOF Then
            
                ' ver 452 This added here 5th MAr
                dbCalcSalesGrand = dbCalcSalesGrand + dbRetailValueTotal
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CalcSalesValue")) = Format(dbRetailValueTotal, "#,##0.00")
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CostOfSales")) = Format(dbCostOfSalesTotal, "#,##0.00")
                dbGrossProfitTotal = (dbRetailValueTotal / (1 + (sngvatrate / 100))) - dbCostOfSalesTotal
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GrossProfit")) = Format(dbGrossProfitTotal, "#,##0.00")
                    
                If dbSalesExVatTotal > 0 Then
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, 6) = Format((dbGrossProfitTotal / dbSalesExVatTotal) * 100, "#,##0.00")
                End If
                
                iGrpCount = iGrpCount + 1
            
            ElseIf iLastGroup <> rs("cboGroups") Then
                    
                ' ver 452 This added here 5th MAr
                dbCalcSalesGrand = dbCalcSalesGrand + dbRetailValueTotal
                
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CalcSalesValue")) = Format(dbRetailValueTotal, "#,##0.00")
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("CostOfSales")) = Format(dbCostOfSalesTotal, "#,##0.00")
                dbGrossProfitTotal = (dbRetailValueTotal / (1 + (sngvatrate / 100))) - dbCostOfSalesTotal
                .grdCount.Cell(flexcpText, .grdCount.Rows - 1, .grdCount.ColIndex("GrossProfit")) = Format(dbGrossProfitTotal, "#,##0.00")
                
                If dbSalesExVatTotal <> 0 Then
                    .grdCount.Cell(flexcpText, .grdCount.Rows - 1, 6) = Format((dbGrossProfitTotal / dbSalesExVatTotal) * 100, "#,##0.00")
                End If
                iGrpCount = iGrpCount + 1
                
            End If
        
        
        Loop While Not rs.EOF
        
        .grdCount.AddItem ""
        
        If dbCostOfSalesGrand > 0 Or dbGrossProfitGrand > 0 Then
            
            dbGrossProfitGrand = (dbCalcSalesGrand / (1 + (sngvatrate / 100))) - dbCostOfSalesGrand
'            .grdCount.AddItem "Totals" & vbTab & Format(dbCalcSalesGrand, "Currency") & _
'                                     vbTab & vbTab & Format(dbCostOfSalesGrand, "Currency") & _
'                                     vbTab & vbTab & Format(dbGrossProfitGrand, "Currency") & _
'                                     vbTab & Format((dbGrossProfitGrand / (dbCostOfSalesGrand + dbGrossProfitGrand)) * 100, "#,##0.00")
        
            If dbGrossProfitGrand - dbGrossProfitGrand <> 0 Then
                .grdCount.AddItem "Totals" & vbTab & Format(dbCalcSalesGrand, "Currency") & _
                                     vbTab & vbTab & Format(dbCostOfSalesGrand, "Currency") & _
                                     vbTab & vbTab & Format(dbGrossProfitGrand, "Currency") & _
                                     vbTab & Format((dbGrossProfitGrand / (dbCostOfSalesGrand + dbGrossProfitGrand)) * 100, "#,##0.00")
        
            Else
                .grdCount.AddItem "Totals" & vbTab & Format(dbCalcSalesGrand, "Currency") & _
                                     vbTab & vbTab & Format(dbCostOfSalesGrand, "Currency") & _
                                     vbTab & vbTab & Format(dbGrossProfitGrand, "Currency")
                                     
            End If
        
        
        End If
        

        .grdCount.AddItem vbTab & "Includes Vat" & _
                          vbTab & vbTab & "Excludes Vat"
        
        For iRow = 2 To .grdCount.Rows - 4 ' ignore last 3 rows 1 blank and the other the totals
        
            If dbCalcSalesGrand > 0 Then
                .grdCount.Cell(flexcpText, iRow, 2) = Format((.grdCount.Cell(flexcpTextDisplay, iRow, 1) / dbCalcSalesGrand) * 100, "#,##0.00")
            End If
            
            If dbCostOfSalesGrand > 0 Then
                .grdCount.Cell(flexcpText, iRow, 4) = Format((.grdCount.Cell(flexcpTextDisplay, iRow, 3) / dbCostOfSalesGrand) * 100, "#,##0.00")
            End If
     
        Next
        
        .grdCount.AddItem ""
    
    End If
    
        
    .grdCount.ScrollBars = flexScrollBarBoth
    .grdCount.AutoSize 0, 6

    End With
    
    gbOk = SetReportSize("")
    
    bHourGlass False
    
    RepGroupTotals = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepGroupTotals") Then Resume 0
    Resume CleanExit


End Function

Public Function GetCalculatedSalesTOTAL() As Double
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbSellExVat As Double
Dim dbRetailValue As Double
Dim dbRetailValueTotal As Double
Dim dbCalcSalesGrand As Double
Dim dbCalcAmount As Double
Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double
Dim curCost As Currency
Dim lLastProd As Long
Dim dblLastQty As Double
Dim iIssue As Integer
'Dim lLastPLU As Long
Dim curSellPrice As Currency
Dim curGlassPrice As Currency
Dim curSellPriceDP As Currency
Dim curGlassPriceDP As Currency
Dim lGlassQty As Long
Dim lGlassQtyDP As Long
Dim lSalesQtyDP As Long

    On Error GoTo ErrorHandler
    
    ' THIS IS A NEW FUNCTION ADDED IN VER 557
    ' ITS A DIRECT COPY OF GROUP TOTAL FUNCTION STRIPPED BACK
    
    dbCalcSalesGrand = 0
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lSelClientID) & " Order By cboGroups, PLUnumber, tblProducts.ID, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
'            If rs("PLUnumber") = 3 Then Stop
            
            If iLastGroup <> rs("cboGroups") Then
                iLastGroup = rs("cboGroups")
                dbRetailValueTotal = 0
            End If
            
            If lLastProd <> rs("tblProducts.ID") Then


'        If rs("tblProductGroup.txtDescription") = "Liquers" Then Stop
        

                lLastProd = rs("tblProducts.ID")
                
                dbSellExVat = Format(rs("SellPrice") / (1 + (sngvatrate / 100)), "0.00")
                ' Sell Ex-Vat
                
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
                ' Deliveries
                
'ver 560
'                If Not IsNull(rs("FullQty")) Then

                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    dbRetailValue = Format((((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
                    dbRetailValueTotal = dbRetailValueTotal + (dbRetailValue * rs("SellPrice"))
                    ' Retail Value
    
'ver 560
'                    If rs("plunumber") <> lLastPLU Then
'                        lLastPLU = rs("plunumber")
                    
                        iIssue = getGlasses(rs("PLUGroupID"))

                        If iIssue > 0 And dbRetailValue > 0 Then

                            If IsNull(rs("SellPrice")) Then curSellPrice = 0 Else curSellPrice = rs("SellPrice")
                            If IsNull(rs("GlassPrice")) Then curGlassPrice = 0 Else curGlassPrice = rs("GlassPrice")
                            If IsNull(rs("SellPriceDP")) Then curSellPriceDP = 0 Else curSellPriceDP = rs("SellPriceDP")
                            If IsNull(rs("GlassPriceDP")) Then curGlassPriceDP = 0 Else curGlassPriceDP = rs("GlassPriceDP")
                            If IsNull(rs("GlassQty")) Then lGlassQty = 0 Else lGlassQty = rs("GlassQty")
                            If IsNull(rs("GlassQtyDP")) Then lGlassQtyDP = 0 Else lGlassQtyDP = rs("GlassQtyDP")
                            If IsNull(rs("SalesQtyDP")) Then lSalesQtyDP = 0 Else lSalesQtyDP = rs("SalesQtyDP")

                            dbRetailValueTotal = dbRetailValueTotal + GlassPriceValueFix(iIssue, curSellPrice, curGlassPrice, curSellPriceDP, curGlassPriceDP, lGlassQty, lGlassQtyDP, lSalesQtyDP)
                                                                                                       ' pint           glass           pint 2nd price      glass 2nd price     glass qty       glass qty @ 2nd price    Pint 2nd qty
                        End If

                        If bDualPrice And dbRetailValue > 0 Then
                            dbRetailValueTotal = dbRetailValueTotal + DualPriceValueFix(rs("SellPrice"), rs("SellPriceDP"), rs("SalesQtyDP"))
                        End If
'                    End If
'                Else
'                Debug.Print rs("tblProductGroup.txtDescription")
'                End If
            End If
            
            rs.MoveNext
            
            If rs.EOF Then
                dbCalcSalesGrand = dbCalcSalesGrand + dbRetailValueTotal
            ElseIf iLastGroup <> rs("cboGroups") Then
                dbCalcSalesGrand = dbCalcSalesGrand + dbRetailValueTotal
            End If
            
        Loop While Not rs.EOF
        
        GetCalculatedSalesTOTAL = dbCalcSalesGrand
    
    End If
    
    rs.Close

CleanExit:
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetCalculatedSalesTOTAL") Then Resume 0
    Resume CleanExit


End Function



Public Function GetActualAndAllowances(lID As Long, _
                                        curActual As Currency, _
                                        curStaff As Currency, _
                                        curCompDrinks As Currency, _
                                        curWastage As Currency, _
                                        curOverRings As Currency, _
                                        curPromotions As Currency, _
                                        curOffSales As Currency, _
                                        curVouchers As Currency, _
                                        curKitchen As Currency, _
                                        curOther As Currency, _
                                        sOther As String, _
                                        curSurplus As Currency, _
                                        sSurplus As String)
                                        
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        curActual = Format(rs("Actual"), "0.00")
        curStaff = Format(rs("Staff"), "0.00")
        curCompDrinks = Format(rs("Complimentary"), "0.00")
        curWastage = Format(rs("Wastage"), "0.00")
        curOverRings = Format(rs("OverRings"), "0.00")
        curPromotions = Format(rs("Promotions"), "0.00")
        curOffSales = Format(rs("OffLicense"), "0.00")
        curVouchers = Format(rs("VoucherSales"), "0.00")
        curKitchen = Format(rs("Kitchen"), "0.00")
        If Val(rs("Other")) > 0 Then
            sOther = rs("OtherTitle") & ""
            curOther = Format(rs("Other"), "0.00")
        End If
        
        If Not IsNull(rs("Surplus")) Then
            If Val(rs("Surplus")) > 0 Then
                curSurplus = Format(rs("Surplus"), "0.00")
                sSurplus = rs("SurplusTitle") & ""
            End If
        Else
            curSurplus = "0.00"
            sSurplus = ""
        End If
    End If
    
    GetActualAndAllowances = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetActualAndAllowances") Then Resume 0
    Resume CleanExit

End Function


Public Function bPrintGrid(fname As Form, sCtrl As String, iCharCount As Integer)
Dim iRow As Integer
Dim iCol As Integer
Dim sOut As String
Dim iWidth As Integer
Dim iFactor As Integer
Dim iNextPos As Integer
Dim bPrintOnSeperateLine As Boolean

    On Error GoTo ErrorHandler
    
    iFactor = 90
    For iRow = 0 To fname.Controls(sCtrl).Rows - 1
        'If fname.Controls(sCtrl).Cell(flexcpTextDisplay, iRow, 1) <> "" Then
        ' as long as there's something on the Grid row...
        
            sOut = String(iCharCount, " ")
            
            'If Not fname.Controls(sCtrl).IsSubtotal(iRow) Then
            ' as long as its not a subtotal row
            
                For iCol = 0 To fname.Controls(sCtrl).Cols - 1
                
                    If Not fname.Controls(sCtrl).ColHidden(iCol) Then
                        iNextPos = Int(1 + fname.Controls(sCtrl).ColPos(iCol) / iFactor)
                        iWidth = Int(fname.Controls(sCtrl).ColWidth(iCol) / iFactor)
                        
                        If fname.Controls(sCtrl).Cell(flexcpChecked, iRow, iCol) = 2 Then
                            Mid(sOut, iNextPos, iWidth) = ""
                        ElseIf fname.Controls(sCtrl).Cell(flexcpChecked, iRow, iCol) = 1 Then
                            Mid(sOut, iNextPos, iWidth) = "Y"
                        Else
                            Mid(sOut, iNextPos, iWidth) = fname.Controls(sCtrl).Cell(flexcpTextDisplay, iRow, iCol)
                        End If
                    End If
                Next
            If iRow = 0 Or (fname.Controls(sCtrl).IsSubtotal(iRow)) Then
            ' highlight a header line
                Printer.Print " "
                Printer.FontBold = True
                Printer.Print sOut
            
            ElseIf iRow < fname.Controls(sCtrl).FixedRows Then
                Printer.FontBold = True
                Printer.Print sOut
            
            Else
            ' standard data line
                Printer.FontBold = False
                Printer.Print sOut
            
                If bPrintOnSeperateLine Then
                    Printer.Print ""
                    bPrintOnSeperateLine = False
                End If
            
            End If
            
        'End If
        
    Next
    Printer.Print ""
    Printer.EndDoc
        
    bPrintGrid = True
    
CleanExit:
    Exit Function

ErrorHandler:
    If CheckDBError("bPrintGrid") Then Resume CleanExit

End Function
Public Function PrintDisplay(fname As Form, sCtrl As String, bVertical As Boolean)
Dim iCopies As Integer
Dim iCnt As Integer
Dim iCharCount As Integer

    On Error GoTo ErrorHandler
  
    
    With fname.ComDlg
  
        .CancelError = True
        .Copies = 1
    
    ' Display the Print dialog box
        Printer.Orientation = bVertical + 2
        .ShowPrinter
        Printer.FontName = "Courier New"
        Printer.FontSize = 10
        Printer.Orientation = fname.ComDlg.Orientation
        iCharCount = 400
        
    ' Get user-selected values from the dialog box
  
        iCopies = .Copies
    
    ' Put code here to send data to the printer

    For iCnt = 1 To iCopies
        
        gbOk = bPrintHeader(fname, iCharCount)
        If gbOk Then gbOk = bPrintGrid(fname, sCtrl, iCharCount)
    Next
    
    End With
    
CleanExit:
    Exit Function

ErrorHandler:
    If Err = 32755 Then
        Resume CleanExit
    ElseIf CheckDBError("PrintDisplay") Then
        Resume CleanExit
    End If
    
End Function
Public Function bPrintHeader(fname As Form, iCharCount As Integer)
Dim sOut As String
    
    On Error GoTo ErrorHandler
    
    
    sOut = String(iCharCount, " ")
    Mid(sOut, 1, iCharCount) = fname.Caption
    Mid(sOut, 80, 17) = Format(Now, "dd/mm/yy hh:mm:ss")
    Printer.Print sOut
    Printer.Print " "
    
    Printer.FontBold = True
    sOut = String(iCharCount, " ")
    Mid(sOut, (90 - Len(fname.labelTitle.Caption)) / 2, Len(fname.labelTitle.Caption)) = fname.labelTitle.Caption
    Printer.Print sOut
    
    bPrintHeader = True
    
CleanExit:
    Exit Function

ErrorHandler:
    If CheckDBError("bPrintHeader") Then Resume CleanExit
    
End Function

Public Function SaveTillDifference(lCLId As Long, lDtID As Long)
Dim rs As Recordset

'Dim lStockSales As Long
Dim dblStockSales As Long
' ver543

Dim rsTill As Recordset
Dim lSalesTotQty As Double
Dim iIssue As Integer

    On Error GoTo ErrorHandler
    
    Set rsTill = SWdb.OpenRecordset("tblTillDifference")
    rsTill.Index = "PrimaryKey"
    
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) LEFT JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lCLId) & " ORDER BY tblClientProductPLUs.PLUGroupID, tblPLUs.txtDescription", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
    
        Do
            
                iIssue = getGlasses(rs("PLUGroupID"))
'(rs("txtIssueUnits") * iIssue)
                
                dblStockSales = GetStockSales(rs("PLUNumber"))
                
                'DP
                lSalesTotQty = 0    ' Init
            
                If Not IsNull(rs("SalesQty")) Then
                    lSalesTotQty = rs("SalesQty")
                End If
            
                If Not IsNull(rs("SalesQtyDP")) Then
                    lSalesTotQty = lSalesTotQty + rs("SalesQtyDP")
                End If
                ' Total Sales units
                

                If Not IsNull(rs("GlassQty")) Then
                    If rs("GlassQty") > 0 Then
                        lSalesTotQty = lSalesTotQty + rs("GlassQty") / iIssue
                    End If
                    
                End If
            
                If Not IsNull(rs("GlassQtyDP")) Then
                    If rs("GlassQtyDP") > 0 Then
                        lSalesTotQty = lSalesTotQty + rs("GlassQtyDP") / iIssue
                    End If
                    
                End If






'---------------------------------------------
' Ver 442
' This line removed as Kate reported 'gaps' appearing in the till reconcile report
' The line was introduced with Dual pricing but only checked Sales Total it should
' have also checked Stock sales as it did previously

                'DP
'                If Not lSalesTotQty = 0 Then
'---------------------------------------------
                    If (dblStockSales <> 0) Or (lSalesTotQty <> 0) Then
                    
                        rsTill.AddNew
                        rsTill("DatesID") = lDtID
                        rsTill("PLUID") = rs("PLUID")
                        If dblStockSales <= lSalesTotQty Then
                            rsTill("Difference") = "+" & Trim$(lSalesTotQty - dblStockSales)
                    
                        Else
                            rsTill("Difference") = "-" & Trim$(dblStockSales - lSalesTotQty)
                    
                        End If
                        rsTill.Update
                        
                    End If
'                End If
            
            
            
            rs.MoveNext
        
        Loop While Not rs.EOF

    End If
    
    SaveTillDifference = True
    
CleanExit:
    Exit Function

ErrorHandler:
    
    If Err = 3022 Then
        rsTill("id") = rsTill("id") + 1
        Resume 0
    
    Else
    
        If CheckDBError("SaveTillDifference") Then Resume CleanExit
    
        Resume 0
    End If
    
End Function

Public Function CreateReport(sClientFolder As String, sTitle As String, iLinesperpage As Integer, iHdrLines As Integer)
Dim objfile As Object
'Dim bFileExists As Boolean
Dim Reportdoc
Dim sReport As String
Dim sNextPage As String

Dim iTotPages As Integer
Dim sAddr As String
Dim sClient As String
Dim iRow As Integer
Dim iCol As Integer
Dim sFrom As String
Dim sTo As String
Dim iRows As Integer
Dim iCnt As Integer
Dim NextPageDoc
Dim iDataPointer As Integer
'Dim iLinesperpage As Integer
Dim iRowCounter As Integer
Dim iCols As Integer
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer
Dim sDP As String ' dual price template Name

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    bHourGlass True
    
'    iLinesPerPage = iLinesPerPage - 1
'    ' Reduce by 1 since there is one blank line at to bottom of each report generated
'    ' and this is causing a header only to be printed
    
    
    frmStockWatch.labelTitle.Caption = "Generating the " & sTitle & " Report"
    
    sTitle = Replace(sTitle, "Print ", "")
    ' incase its a print sheet thats looked for
    
    
    gbOk = TerminateWINWORD()
    
    sReport = sClientFolder & "/" & sTitle & ".Doc"
    sNextPage = sClientFolder & "/" & "NextPage.Doc"
    
    On Error GoTo CantDeleteReport
            
    Kill sReport
    
    
deleteNextPage:
    On Error GoTo CantDeleteNextPage
    
    Kill sNextPage
    ' DELETE Report HERE!
        '
ResumeHere:
    On Error GoTo ErrorHandler
    
    
    Set WriteWord = New Word.Application
    WriteWord.Visible = False
    
'    iRows = frmStockWatch.grdCount.Rows - 1
    
    iRows = frmStockWatch.grdCount.Rows


    
    
'Ver 4.0.2
    If iRows = 0 Then
        iTotPages = 1

    ElseIf (iRows - iHdrLines) Mod (iLinesperpage - iHdrLines) = 0 Then
        iTotPages = Int((iRows - iHdrLines) / (iLinesperpage - iHdrLines))
    
    Else
        iTotPages = (Int((iRows - iHdrLines) / (iLinesperpage - iHdrLines))) + 1
    End If
    
    
    'Ver310
    sDP = "" ' default
    

' Ver 542 fix for Product deficit report which has different number of rows for 1st
' page vs other pages:     25 vs 40

    If sTitle = "Product Deficit" Then
    
        If iRows < 26 Then
            iTotPages = 1
    
        ElseIf (iRows - iHdrLines - 26) Mod (50 - iHdrLines) = 0 Then
        iTotPages = Int((iRows - iHdrLines - 26) / (50 - iHdrLines)) + 1
    
        Else
            iTotPages = (Int((iRows - iHdrLines - 26) / (50 - iHdrLines))) + 2
        End If
    
        
    ElseIf sTitle = "Stock Analysis" Then
    'DP - to pcik up the proper template if
    ' working on a dual price client
        
    
        ' ver 536
        sDP = "DPandMEASURE"
    
    End If


    iDataPointer = frmStockWatch.grdCount.FixedRows
    ' Initialize the Row index pointer to point to the first line of data past the header.
 
    For iCnt = 1 To iTotPages
  '      If iCnt = 4 Then Stop
        
         bHourGlass True
         
         If iCnt = 1 Then
            Set Reportdoc = WriteWord.Documents.Add(sDBLoc & "\Templates\" & sTitle & sDP & ".dot")
         
         ElseIf sTitle = "Product Deficit" Then
         
            Set NextPageDoc = WriteWord.Documents.Add(sDBLoc & "\Templates\" & sTitle & sDP & " NextPage.dot")
            iLinesperpage = 50
            
         Else
            Set NextPageDoc = WriteWord.Documents.Add(sDBLoc & "\Templates\" & sTitle & sDP & ".dot")
         End If
         
         WriteWord.Visible = False
         
         WriteWord.ActiveDocument.UndoClear
        
         With WriteWord.ActiveDocument.Bookmarks
        
            ' HEADER
            
            ' Program Name & revision
            ' app.Title app.Revision app.Major app.minor
        
            sClient = GetClientName(lSelClientID, sAddr, True)
            .Item("Client").Range.Text = sClient
            .Item("Address").Range.Text = sAddr
            ' Client Name & Address
    
            .Item("Date").Range.Text = Format(Now, sDMMYY)
            ' Date & Time
        
            gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
            .Item("From").Range.Text = sFrom
            .Item("To").Range.Text = sTo
            ' Stock Take Dates From & To
        
            .Item("ReportName").Range.Text = sTitle & " Report"
            ' Report title
            
            .Item("PageNo").Range.Text = Trim$(iCnt) & " of " & Trim$(iTotPages)
            ' Page No
            
            WriteWord.Visible = False

            If sTitle = frmStockWatch.grdMenu.Cell(flexcpTextDisplay, frmStockWatch.grdMenu.FindRow("D"), 1) And frmStockWatch.grdCount.Cols > 8 Then
                iCols = 9
            Else
                iCols = frmStockWatch.grdCount.Cols
            End If
            
            For iRow = 0 To frmStockWatch.grdCount.FixedRows - 1
                
                ' ver310
                If sTitle = "Product Deficit" Then
                    For iCol = 0 To iCols - 2   ' Ignore Select Column
                        .Item("Cell" & "R" & Trim$(iRow) & "C" & Trim$(iCol)).Range.Text = Trim$(frmStockWatch.grdCount.Cell(flexcpTextDisplay, iRow, iCol))
                    Next
                Else
                    For iCol = 0 To iCols - 1
                        .Item("Cell" & "R" & Trim$(iRow) & "C" & Trim$(iCol)).Range.Text = Trim$(frmStockWatch.grdCount.Cell(flexcpTextDisplay, iRow, iCol))
                    Next
                End If
                ' COLUMN TITLES
            Next
            ' There maybe 2 fixed rows in the title so we need to allow for that
            ' A corresponding No of shaded rows are in the Template
            
            
            ' DATA
            
            iRowCounter = frmStockWatch.grdCount.FixedRows
            ' For each page init row counter
            
            If iRows > iRowCounter Then
            ' do a check here incase there are no records and just want to show a blank page
                Do
                
                    'ver310
                    If sTitle = "Product Deficit" Then
                    ' Ignore the "Sel" Column at the right of the report
                        For iCol = 0 To iCols - 2
                            
                            If frmStockWatch.grdCount.Cell(flexcpChecked, iDataPointer, frmStockWatch.grdCount.ColIndex("Sel")) = flexChecked Then
                                .Item("Cell" & "R" & Trim$(iRowCounter) & "C" & Trim$(iCol)).Range.Text = frmStockWatch.grdCount.Cell(flexcpTextDisplay, iDataPointer, iCol)
                            End If
                        Next
                    
                        If frmStockWatch.grdCount.Cell(flexcpChecked, iDataPointer, frmStockWatch.grdCount.ColIndex("Sel")) = flexChecked Then
                            iRowCounter = iRowCounter + 1
                        End If
                        iDataPointer = iDataPointer + 1


                    Else
                        For iCol = 0 To iCols - 1
                            .Item("Cell" & "R" & Trim$(iRowCounter) & "C" & Trim$(iCol)).Range.Text = frmStockWatch.grdCount.Cell(flexcpTextDisplay, iDataPointer, iCol)
                        Next
                    
                        iRowCounter = iRowCounter + 1
                        iDataPointer = iDataPointer + 1
                    
                    End If
                    
                    WriteWord.Visible = False

            
                    
                Loop While (iRowCounter < iLinesperpage) And (iDataPointer < frmStockWatch.grdCount.Rows - 1)
            
            End If
            
            
            If iCnt = 1 Then
                
                ' ADVISORY NOTE ON SUMMARY ANALYSIS NOW CALLED : Product Deficit
                If sTitle = "Product Deficit" Then
                    .Item("Note").Range.Text = GetSummaryDetails(lDatesID, True)
                    .Item("Total").Range.Text = Replace(frmStockWatch.labelTotal.Caption, "Total: ", "")
                
                End If
                
                WriteWord.Visible = False
            
                WriteWord.ActiveDocument.UndoClear

                WriteWord.Application.NormalTemplate.Saved = True
               
                WriteWord.Application.ActiveDocument.SaveAs (sReport)
                ' Save it as the file name thats required
            
                Reportdoc.Close False
                ' close first page
                
            Else
                WriteWord.Visible = False
                WriteWord.ActiveDocument.UndoClear
            
                WriteWord.Application.NormalTemplate.Saved = True
                WriteWord.Visible = False
        
                WriteWord.Application.ActiveDocument.SaveAs (sNextPage)
                ' Save it as the file name thats required
            
                NextPageDoc.Close False
                ' close next page
                
                WriteWord.Visible = False
                
                WriteWord.Application.Documents.Open sReport
                ' open 1st page again
                
                WriteWord.Application.NormalTemplate.Saved = True

                WriteWord.Selection.EndKey wdStory
                WriteWord.Selection.InsertBreak wdSectionBreakNextPage
                WriteWord.Visible = False
                ' go to end section of first page and set a break
                ' page to go to new page
                
                WriteWord.ActiveDocument.UndoClear
                
                WriteWord.Selection.InsertFile sNextPage
                
                WriteWord.Visible = False
                ' Merge each page into top page
                
                Kill sNextPage
                Set NextPageDoc = Nothing

                
            End If

        End With
    
    
    Next
    
    
'    On Error GoTo FinishUpProblem

'Ver 2.1.0
' add this to print



CloseWordStuff:
    
    bHourGlass True
    
    On Error Resume Next
    
'Ver 555 ---------------
    If bEvaluation Then
        gbOk = CopyValuation(frmStockWatch.lblClient.Tag, Replace(frmStockWatch.labelTo.Tag, "/", "-"))
    End If
'Ver 555 ---------------
    
    WriteWord.Application.NormalTemplate.Saved = True
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    
    Set objfile = Nothing
    Set Reportdoc = Nothing
    Set NextPageDoc = Nothing
    
    
    '-----------------------------------
    'KILL NEXT FILE
    Dim sFile, sf
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object
    
    ' NOW CHECK TO SEE IF THERE'S A NEW SWIAgent PROGRAM.
    
    If objfile.FileExists(sNextPage) Then
    
        On Error Resume Next
        
        Set sFile = objfile.GetFile(sNextPage)
        sf = sFile.Delete
        ' kill old
    End If
    
    
'    Kill sNextPage
    

CleanExit:
'
    
    bHourGlass True
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("CreateReport") Then Resume 0
    Resume CloseWordStuff
    
CantDeleteReport:
    LogMsg frmStockWatch, "", sReport & " Can't delete report: " & Trim$(Error)
    Resume deleteNextPage
    
CantDeleteNextPage:
    LogMsg frmStockWatch, "", sReport & " Can't delete next page: " & Trim$(Error)
    Resume ResumeHere
     

End Function

Public Function GetStockTakeDates(lDtID As Long, sFrom As String, sTo As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lDtID
    If Not rs.NoMatch Then
    
        sFrom = Trim$(Format(rs("From"), "dd mmm yy"))
        sTo = Trim$(Format(rs("To"), "dd mmm yy"))
        
        GetStockTakeDates = True
    
    End If
    
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetStockTakeDates") Then Resume 0
    Resume CleanExit

End Function

Public Function ShowReport(sClientFolder As String, sTitle As String)
Dim objfile As Object
'Dim bFileExists As Boolean
Dim sReport As String

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    sReport = sTitle
   
    If objfile.FileExists(sReport) Then
    ' see if file already exists

        Set WriteWord = New Word.Application
        WriteWord.Visible = False
    
        WriteWord.Application.Documents.Open sReport, , vbTrue
        ' open report
 
        WriteWord.Visible = True
    
        WriteWord.Application.NormalTemplate.Saved = True
        
        ShowReport = True
        
    Else
        ShowReport = False
    End If
        

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
 '   If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowReport") Then Resume 0
    Resume CleanExit

End Function

Public Function PrintReportCover(sName As String, sDate As String)
Dim objfile As Object
'Dim bFileExists As Boolean
Dim Reportdoc
Dim sReport As String
Dim sAddr As String
Dim sClient As String
Dim sFrom As String
Dim sTo As String

Dim sFranNAme As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    bHourGlass True
    
    frmStockWatch.labelTitle.Caption = "Generating the Report Cover"
   
    sReport = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\ReportCover.Doc"
   
    On Error Resume Next
            
    Kill sReport
    
    On Error GoTo ErrorHandler
    
    WriteWord.Visible = False
    
    Set Reportdoc = WriteWord.Documents.Add(sDBLoc & "\Templates\Report Cover.dot")
    
    With WriteWord.ActiveDocument.Bookmarks
        
        sClient = GetClientName(lSelClientID, sAddr, True)
        .Item("Client").Range.Text = sClient
        .Item("Address").Range.Text = sAddr
        ' Client Name & Address
    
        .Item("Date").Range.Text = Format(Now, "dd mmm yy hh:mm")
        ' Date & Time
        
        gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
        .Item("From").Range.Text = sFrom
        .Item("To").Range.Text = sTo
        ' Stock Take Dates From & To
        
        If GetLicenseInfo(sFranNAme, sAddress, sPhone, sEmail, dtExpiry, iDays, iwarn) Then
            
            .Item("FranName").Range.Text = sFranNAme
            .Item("FranAddress").Range.Text = sAddress
            .Item("FranPhone").Range.Text = sPhone
            .Item("FranRegion").Range.Text = gbRegion
                
        End If

        WriteWord.Application.NormalTemplate.Saved = True

        WriteWord.Application.ActiveDocument.SaveAs (sReport)

        WriteWord.Application.PrintOut -1
    
'        WriteWord.Quit
        
        PrintReportCover = True
    End With
    

CloseWordStuff:
    
    On Error Resume Next
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    Set objfile = Nothing
    Set Reportdoc = Nothing

CleanExit:
'
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("PrintReportCover") Then Resume 0
    Resume CloseWordStuff

End Function

Public Function PrintCountSheets(sDate As String, lID As Long, bStockCountSheet As Boolean, iCopies As Integer)
Dim rs As Recordset
Dim iLastGroup As Integer
Dim sOut As String
Dim sClient As String
Dim sAddr As String
Dim sFrom As String
Dim sTo As String
Dim bVertical As Boolean
Dim iLinesperpage As Integer
Dim iLinesCounter As Integer
Dim iPageNo As Integer
Dim sPages As String
Dim iTotPages As Integer
Dim iCopy As Integer
Dim iGrp As Integer

    On Error GoTo ErrorHandler
    
    iLinesperpage = 60
    ' page = 63 lines
    ' header = 6 lines
    
    
'    Dim frmCD As New frmCD
'    frmCD.Move (Screen.Width - frmCD.Width) / 2, (Screen.Height - frmCD.Height) / 2
    
    With frmStockWatch
        .ComDlg.Copies = iCopies
        .ComDlg.CancelError = True
        .ComDlg.ShowPrinter
    End With
    
    bVertical = True
    Printer.Orientation = bVertical + 2
'    frmStockWatch.comdlg.ShowPrinter
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
'    Printer.FontSize = 11
    'ver 304
    
    ' Display the Print dialog box
        
'    iCopies = frmStockWatch.comdlg.Copies
    
    sDate = Format(Now, "dd mmm yy hh:mm")
    
    For iCopy = 1 To iCopies

        iPageNo = 0
        iLinesCounter = 0
        iLastGroup = 0
        iGrp = 0
        sOut = ""
        
        'ver522
        Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND Active = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
        ' SQl return list of plus for this clinet sorted by plu#
        If Not (rs.EOF And rs.BOF) Then
            
            rs.MoveFirst
            
            ' GROUP COUNTER
            Do
                If iLastGroup <> rs("cboGroups") Then
                    iGrp = iGrp + 1
                    iLastGroup = rs("cboGroups")
                End If
                rs.MoveNext
            Loop While Not rs.EOF
            ' get number of groups
            
            rs.MoveLast
            
            If ((rs.RecordCount + (iGrp * 2)) Mod (iLinesperpage - 6)) = 0 Then
                iTotPages = (rs.RecordCount + (iGrp * 2)) / (iLinesperpage - 6)
            Else
                iTotPages = Int((rs.RecordCount + (iGrp * 2)) / (iLinesperpage - 6)) + 1
            End If
            ' Get how many pages of a report are going to be produced
            ' also include the numer of groups times 2 cause there's 2 lines for a group header
            
            rs.MoveFirst
            
            sClient = GetClientName(lSelClientID, sAddr, True)
            gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
            
            Do
                
                If iLinesCounter = 0 Or iLinesCounter >= iLinesperpage Then
                ' New Page - print header
                
                    iPageNo = iPageNo + 1
                    sPages = Trim$(iPageNo) & " of " & Trim$(iTotPages)
                    
                    ' PRINT HEADER
                    
                    If iPageNo > 1 Then Printer.NewPage
                
                    If bStockCountSheet Then
                        iLinesCounter = PrintHeader(sClient, sAddr, sFrom, sTo, sDate, sPages, "Stock Count Sheet")
                        Printer.Print "Description                   Size <----------- Counted -----------> < Total >"
                    Else
                        iLinesCounter = PrintHeader(sClient, sAddr, sFrom, sTo, sDate, sPages, "Delivery Sheet")
                        Printer.Print "Description                   Size           Delivery Notes          < Total >"
                    End If
                    iLinesCounter = iLinesCounter + 1
    
                End If
                
                
                sOut = String$(80, " ")
                
                If iLastGroup <> rs("cboGroups") Then
    
                    Printer.FontUnderline = False
                    Printer.FontBold = True
                    Printer.Print ""
                    Printer.Print " >>>  " & vbTab & rs("cboGroups") & "  " & rs("tblProductGroup.txtDescription")
                    iLastGroup = rs("cboGroups")
                    iLinesCounter = iLinesCounter + 2
                End If
                    
                Printer.FontBold = False
                Printer.FontUnderline = True
                
                Mid(sOut, 1, 30) = rs("tblProducts.txtDescription")
                Mid(sOut, 31, 7) = rs("txtSize")
                Mid(sOut, 69, 1) = "|"
                Printer.Print sOut
                iLinesCounter = iLinesCounter + 1
                
                        
                rs.MoveNext
            Loop While Not rs.EOF
        
        End If
        
        Printer.EndDoc
    
    Next
    
    PrintCountSheets = True
    rs.Close


CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If Err = 32755 Then Resume CleanExit
    
    If CheckDBError("PrintCountSheets") Then Resume 0
    Resume CleanExit

End Function

Public Function PrintHeader(sClient As String, sAddr As String, sFrom As String, sTo As String, sDate As String, sPages As String, sTitle As String)
Dim sOut As String

    Printer.FontUnderline = False
        
    sOut = String$(80, " ")
    Mid(sOut, 1, 20) = "Stock Watch"
    Mid(sOut, 40 - (Len(sClient)) / 2, 30) = sClient
    Mid(sOut, 65, 15) = sDate
    Printer.Print sOut
    ' Program Name & revision
        
    sOut = String$(80, " ")
    Mid(sOut, 40 - (Len(sAddr)) / 2, 50) = sAddr
    Printer.Print sOut
    Printer.Print ""
    ' Address
        
    sOut = String$(80, " ")
    Mid(sOut, 1, 20) = sTitle
    Mid(sOut, 24, 36) = "From  " & sFrom & "  To  " & sTo
    Mid(sOut, 68, 13) = "Pages " & sPages
    Printer.Print sOut
        
    Printer.Print String$(80, "=")
        
    PrintHeader = 5
        
End Function

Public Function PrintPLUCountSheet(sDate As String, lID As Long, iCopies As Integer)
Dim rs As Recordset
Dim iLastPLU As Integer
Dim sOut As String
Dim sClient As String
Dim sAddr As String
Dim sFrom As String
Dim sTo As String
Dim bVertical As Boolean
Dim iLinesperpage As Integer
Dim iLinesCounter As Integer
Dim iPageNo As Integer
Dim sPages As String
Dim iTotPages As Integer
Dim iCnt As Integer
Dim iCopy As Integer
Dim iRecLineCount As Integer
Dim sGlsRetail As String

    On Error GoTo ErrorHandler
    
    iLinesperpage = 62
    ' page = 63 lines
    ' header = 6 lines
    
    With frmStockWatch
        .ComDlg.Copies = iCopies
        .ComDlg.CancelError = True
        .ComDlg.ShowPrinter
    End With
    
    bVertical = True
    Printer.Orientation = bVertical + 2
    Printer.FontName = "Courier New"
    Printer.FontSize = 12

    sDate = Format(Now, "dd mmm yy hh:mm")
    
    For iCopy = 1 To iCopies
    
        iPageNo = 0
        iLinesCounter = 0
        iLastPLU = 0
        iRecLineCount = 0 'DP
        sOut = ""
    
        'ver522
        Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND Active = true ORDER BY PLUNumber", dbOpenSnapshot)
        ' SQl return list of plus for this clinet sorted by plu#
        If Not (rs.EOF And rs.BOF) Then
        
            'DP ----------------------------------------------------------------
            
            ' CALCULATE THE NUMBER OF PAGES
            
            ' Loop thru recordset
            rs.MoveFirst
            
            Do
                ' IF THERES A DUAL PRICE THEN COUNT ANOTHER LINE
                
                If bDualPrice Then
                
                    iRecLineCount = iRecLineCount + 2
                Else
                    iRecLineCount = iRecLineCount + 1
                    ' normal line count
                End If
                
                rs.MoveNext
            Loop While Not rs.EOF
            
            If (iRecLineCount Mod iLinesperpage) = 0 Then
                iTotPages = iRecLineCount / iLinesperpage
            Else
                iTotPages = Int(iRecLineCount / iLinesperpage) + 1
            End If
            ' Get how many pages of a report are going to be produced
            
            '--------------------------------------------------------------------
            
            rs.MoveFirst
                
            sClient = GetClientName(lSelClientID, sAddr, True)
            gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
            Do
                    
                If iLinesCounter = 0 Or iLinesCounter >= iLinesperpage Then
                ' New Page - print header
                
                    iPageNo = iPageNo + 1
                    sPages = Trim$(iPageNo) & " of " & Trim$(iTotPages)
                    
                    ' PRINT HEADER
                    
                    If iPageNo > 1 Then Printer.NewPage
                    
                    iLinesCounter = PrintHeader(sClient, sAddr, sFrom, sTo, sDate, sPages, "PLU Sales Worksheet")
'                   Printer.Print "No Grp Description               Retail                Sales               Total"
' Ver 537
                    Printer.Print "No Grp Description                Retail          Sales            | Units | Gls"
                    iLinesCounter = iLinesCounter + 1
    
                End If
                
                sOut = String$(80, " ")
                    
                If iLastPLU <> rs("PLUNumber") Then
    
                    Mid(sOut, 1, 4) = rs("PLUNumber")
                    Mid(sOut, 5, 2) = rs("txtGroupNumber")
                    Mid(sOut, 8, 30) = rs("tblPLUs.txtDescription")
'                    Mid(sOut, 33, 7) = Right("       " & Format(rs("SellPrice"), "0.00"), 7)
'                    Mid(sOut, 71, 1) = "|"
'Ver 537
'Ver 538 change / to | between retail prices

                    sGlsRetail = ""
                    If Not IsNull(rs("GlassPrice")) Then
                        If rs("GlassPrice") > 0 Then
                            sGlsRetail = "|" & Format(rs("GlassPrice"), "0.00")
                        End If
                    End If
                    
                    Mid(sOut, 32, 14) = Right("       " & Format(rs("SellPrice"), "0.00"), 7) & sGlsRetail
                    Mid(sOut, 68, 1) = "|"
                    Mid(sOut, 76, 1) = "|"
                    
                    'DP
                    If bDualPrice Then
                    
                        Mid(sOut, 46, 22) = String$(31, "_")
                        Mid(sOut, 69, 7) = String$(9, "_")
                        Mid(sOut, 77, 4) = String$(9, "_")
                        Printer.FontUnderline = False
                        Printer.Print sOut
                        iLinesCounter = iLinesCounter + 1
                        ' 1st line of dual price printed
                        
'Ver 537
                        sGlsRetail = ""
                        If Not IsNull(rs("GlassPriceDP")) Then
                            If rs("GlassPriceDP") > 0 Then
                                sGlsRetail = "|" & Format(rs("GlassPriceDP"), "0.00")
                            End If
                        End If
                        
                        
                        sOut = String$(80, " ")
                        Mid(sOut, 32, 14) = Right("       " & Format(rs("SellPriceDP"), "0.00"), 7) & sGlsRetail
                        Mid(sOut, 68, 1) = "|"
                        Mid(sOut, 76, 1) = "|"
                        ' setup for 2nd line to be printed
                        
                    End If
                    
                    Printer.FontUnderline = True
                    Printer.Print sOut
                    iLinesCounter = iLinesCounter + 1
                    
                    iLastPLU = rs("PLUNumber")
                    iCnt = 1
                    
                Else
                    iCnt = iCnt + 1
                End If
                    
                rs.MoveNext
            Loop While Not rs.EOF
    
        
        End If
            
        Printer.EndDoc
    
    Next
    
    PrintPLUCountSheet = True
    rs.Close



CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    

    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If Err = 32755 Then Resume CleanExit
    
    If CheckDBError("PrintPLUCountSheet") Then Resume 0
    Resume CleanExit

    

End Function

Public Function GetCalculatedSales(lID As Long)
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbRetailValue As Double
Dim dbCalcSalesGrand As Double
Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double
Dim curCost As Currency
Dim lLastProd As Long
Dim dbCalcAmount As Double
Dim dblLastQty As Double

    On Error GoTo ErrorHandler
    
    ' VAT INCLUSIVE
    
' Ver 546 ------------------------------
' Per phone call with Don/Kate figures were wrong on Calculated sales cash entry screen.
' This was copied from Group Totals to make them match

    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lSelClientID) & " Order By cboGroups, PLUnumber, tblProducts.ID, tblProducts.txtDescription, txtSize", dbOpenSnapshot)

'    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lID) & " Order By cboGroups, tblProducts.txtDescription", dbOpenSnapshot)
'---------------------------------------



    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
            If iLastGroup <> rs("cboGroups") Then
                
                iLastGroup = rs("cboGroups")
            
            End If
            
            If lLastProd <> rs("tblProducts.ID") Then

                lLastProd = rs("tblProducts.ID")
                
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
                ' Deliveries
                
    '            If Not IsNull(rs("FullQty")) And Not IsNull(rs("LastQty")) Then
    'Ver 4.0.1
                    
                If Not IsNull(rs("FullQty")) Then

                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    
    ' ver 4.0.2
    '                dbRetailValue = Format((((rs("LastQty") + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
                    dbRetailValue = Format((((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
                    ' Retail Value
                
                    dbCalcSalesGrand = dbCalcSalesGrand + (dbRetailValue * rs("SellPrice"))
                
                End If

            End If
            
            rs.MoveNext
        
            
        Loop While Not rs.EOF
        
    End If
    
    GetCalculatedSales = dbCalcSalesGrand
    
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("getCalculatedSales") Then Resume 0
    Resume CleanExit

End Function

Public Function RemoveClientDates(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        rs.Delete
    End If
    
    RemoveClientDates = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RemoveClientDates") Then Resume 0
    Resume CleanExit

End Function
Public Function GetPurchasePrice(lID As Long) As Currency

Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        GetPurchasePrice = rs("PurchasePrice") + 0
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
'

    Exit Function

ErrorHandler:
    If CheckDBError("GetPurchasePrice") Then Resume 0
    Resume CleanExit

End Function

Public Function GetDates(lID As Long, sFromDate As String, sToDate As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        sFromDate = Format(rs("From"), sDMY)
        sToDate = Format(rs("To"), sDMY)
    End If

    rs.Close

CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("GetDates") Then Resume 0
    Resume CleanExit


End Function


Public Function Encrypt(ByVal Plain As String, _
  sEncKey As String) As String
    '*********************************************************
    'Coded WhiteKnight 6-1-00
    'This Encrypts A string by converting it to its ASCII number
    'but the difference is it uses a Key String it converts the
    'keystring to ASCII and adds it to the first ASCII Value the
    'key is needed to decrypt the text.  I do plan on changing
    'this some what but For Now its ok.  I've only seen it
    'cause an error when the wrong Key was entered while
     'decrypting.
    
    'Note That If you use the same letter more then 3 times in a
    'row then each letter after it if still the same is ignored
    '(ie aaa = aaaaaaaaa but aaa <> aaaza)
    'If anyone Can figure out a way to fix this please e-mail me
  '*********************************************************
    Dim encrypted2 As String
    Dim LenLetter As Integer
    Dim Letter As String
    Dim KeyNum As String
    Dim encstr As String
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops
    
    If sEncKey = "" Then sEncKey = sKey
    'Sets the Encryption Key if one is not set
    ReDim encKEY(1 To Len(sEncKey))
    
    'starts the values for the Encryption Key
        
    For i = 1 To Len(sEncKey$)
     KeyNum = Mid$(sEncKey$, i, 1) 'gets the letter at index i
     encKEY(i) = Asc(KeyNum) 'sets the the Array value
                             'to ASC number for the letter

           'This is the first letter so just hold the value
        If i = 1 Then Math = encKEY(i): GoTo nextone

        'compares the value to the previous value and then
        'either adds/subtracts the value to the Math total
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)

        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) >= _
           encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
          >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    For i = 1 To Len(Plain) 'Now for the String to be encrypted
        Letter = Mid$(Plain, i, 1) 'sets Letter to
                                   'the letter at index i
        LenLetter = Asc(Letter) + Math 'Now it adds the Asc
                                       'value of Letter to Math

'checks and corrects the format then adds a space to separate them frm each other
        If LenLetter >= 100 Then encstr = _
             encstr & Asc(Letter) + Math & " "

         'checks and corrects the format then adds a space
        'to separate them frm each other
        If LenLetter <= 99 Then encstr$ = encstr & "0" & _
          Asc(Letter) + Math & " "
    Next i


    'This is part of what i'm doing to convert the encrypted
    'numbers to Letters so it sort of encrypts the
    'encrypted message.
    temp$ = encstr 'hold the encrypted data
    temp$ = TrimSpaces(temp) 'get rid of the spaces
    itempnum% = Mid(temp, 1, 2) 'grab the first 2 numbers
    temp2$ = Chr(itempnum% + 100) 'Now add 100 so it
                                   'will be a valid char

    'If its a 2 digit number hold it and continue
    If Len(itempnum%) >= 2 Then itempstr$ = Str(itempnum%)
 
   'If the number is a single digit then add a '0' to the front
   'then hold it
    If Len(itempnum%) = 1 Then itempstr$ = "0" & _
        TrimSpaces(Str(itempnum%))
    
    encrypted2$ = temp2 'set the encrypted message
    
    For i = 3 To Len(temp) Step 2
        itempnum% = Mid(temp, i, 2) 'grab the next 2 numbers
  
      ' add 100 so it will be a valid char
        temp2$ = Chr(itempnum% + 100)

      'if its the last number we only want to hold it we
       'don't want to add a '0' even if its a single digit
        If i = Len(temp) Then itempstr$ = _
         Str(itempnum%): GoTo itsdone

'If its a 2 digit number hold it and continue
        If Len(itempnum%) = 2 Then itempstr$ = _
            Str(itempnum%)

        'If the number is a single digit then add a '0'
        'to the front then hold it
        If Len(TrimSpaces(Str(itempnum))) = 1 Then _
      itempstr$ = "0" & TrimSpaces(Str(itempnum%))

        'Now check to see if a - number was created
        'if so cause an error message
        If Left(TrimSpaces(Str(itempnum)), 1) = "-" Then _
          Err.Raise 20000, , "Unexpected Error"
           

itsdone:
           'Set The Encrypted message
        encrypted2$ = encrypted2 & temp2$
    Next i


    'Encrypt = encstr 'Returns the First Encrypted String
    Encrypt = encrypted2 'Returns the Second Encrypted String
    Exit Function 'We are outta Here
oops:
    Debug.Print "Error description", Err.Description
End Function

Public Function Decrypt(ByVal Encrypted As String, _
    sEncKey As String) As String

    Dim NewEncrypted As String
    Dim Letter As String
    Dim KeyNum As String
    Dim EncNum As String
    Dim encbuffer As Long
    Dim strDecrypted As String
    Dim Kdecrypt As String
    Dim lastTemp As String
    Dim LenTemp As Integer
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops

    If sEncKey = "" Then sEncKey = sKey

    ReDim encKEY(1 To Len(sEncKey))
    
    'Convert The Key For Decryption
    For i = 1 To Len(sEncKey$)
        KeyNum = Mid$(sEncKey$, i, 1) 'Get Letter i% in the Key
        encKEY(i) = Asc(KeyNum) 'Convert Letter i to Asc value
 
'if it the first letter just hold it
       If i = 1 Then Math = encKEY(i): GoTo nextone
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
               <= encKEY(i - 1) Then Math = Math - encKEY(i)
               'compares the value to the previous value and
               'then either adds/subtracts the value to the
               'Math total
        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
              <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    'This is part of what i'm doing to convert the encrypted
    'numbers to  Letters so it sort of encrypts the encrypted
    'message.
    temp$ = Encrypted 'hold the encrypted data


    For i = 1 To Len(temp)
        itempstr = TrimSpaces(Str(Asc(Mid(temp, i, 1)) - _
           100)) 'grab the next 2 numbers
           'If its a 2 digit number hold it and continue
        If Len(itempstr$) = 2 Then itempstr$ = itempstr$
          If i = Len(temp) - 2 Then LenTemp% = _
               Len(Mid(temp2, Len(temp2) - 3))
          If i = Len(temp) Then itempstr$ = _
              TrimSpaces(itempstr$): GoTo itsdone
          'If the number is a single digit then add a '0' to the
          'front then hold it
        If Len(TrimSpaces(itempstr$)) = 1 Then _
             itempstr$ = "0" & TrimSpaces(itempstr$)
        'Now check to see if a - number was created if so
        'cause an error message
        If Left(TrimSpaces(itempstr$), 1) = "-" Then _
             Err.Raise 20000, , "Unexpected Error"
           

itsdone:
        temp2$ = temp2$ & itempstr 'hold the first decryption
    Next i
    
    
    Encrypted = TrimSpaces(temp2$) 'set the encrypted data


    For i = 1 To Len(Encrypted) Step 3
        'Format the encrypted string for the second decryption
        NewEncrypted = NewEncrypted & _
            Mid(Encrypted, CLng(i), 3) & " "
    Next i

' Hold the last set of numbers to check it its the correct format
    lastTemp$ = TrimSpaces(Mid(NewEncrypted, _
         Len(NewEncrypted$) - 3))
         
         If Len(lastTemp$) = 2 Then
' If it = 2 then its not the Correct format and we need to fix it
        lastTemp$ = Mid(NewEncrypted, _
           Len(NewEncrypted$) - 1) 'Holds Last Number so a '0'
                                    'Can be added between them
'set it to the new format
        Encrypted = Mid(NewEncrypted, 1, _
           Len(NewEncrypted) - 2) & "0" & lastTemp$
Else
        Encrypted$ = NewEncrypted$ 'set the new format

    End If
    'The Actual Decryption
    For i = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i, 1) 'Hold Letter at index i
        EncNum = EncNum & Letter 'Hold the letters
        If Letter = " " Then 'we have a letter to decrypt
            encbuffer = CLng(Mid(EncNum, 1, _
              Len(EncNum) - 1)) 'Convert it to long and
                                 'get the number minus the " "
            strDecrypted$ = strDecrypted & Chr(encbuffer - _
               Math) 'Store the decrypted string
            EncNum = "" 'clear if it is a space so we can get
                        'the next set of numbers
        End If
    Next i

    Decrypt = strDecrypted

    Exit Function
oops:
    Debug.Print "Error description", Err.Description
Err.Raise 20001, , "You have entered the wrong encryption string"
    Resume 0
End Function

Private Function TrimSpaces(strstring As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strstring$, " ")
        DoEvents
         lngpos& = InStr(1&, strstring$, " ")
         strstring$ = Left$(strstring$, (lngpos& - 1&)) & _
            Right$(strstring$, Len(strstring$) - _
               (lngpos& + Len(" ") - 1&))
    Loop
     TrimSpaces$ = strstring$
End Function
Public Function WindowAppear(frm As Form, iOpacityStart As Integer, _
                                        iNewTop As Integer, _
                                        iNewLeft As Integer, _
                                        iSpeed As Integer, _
                                        bWithSound As Boolean)
Dim iCnt As Integer

        
        frm.Top = iNewTop
        frm.Left = iNewLeft
        
        SetTranslucent frm.hwnd, 0
        DoEvents
        
        
        For iCnt = 1 To 255 Step iSpeed
    
            SetTranslucent frm.hwnd, iCnt
                
        Next

        SetTranslucent frm.hwnd, 255


End Function
Sub SetTranslucent(ThehWnd As Long, nTrans As Integer)
    On Error GoTo ErrorRtn

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    
'    If General_ChkTranslucent Then
    
        attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
        SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
        SetLayeredWindowAttributes ThehWnd, RGB(255, 255, 0), nTrans, LWA_ALPHA
'    End If
    
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
    
End Sub

Public Function Cuint(ByVal v As Integer) As Long
Dim r As Long

If (v > 0) Then
r = v
Else
r = 65536 + v
End If
Cuint = r
End Function

Public Function Uni2Ac(ByRef buffer() As Byte, ByVal Leng As Integer)
Dim i As Integer

For i = 0 To (Leng / 2 - 1)               ' 2 bytes-> 1 bytes, remove '00'
    buffer(i) = buffer(2 * i)
Next

buffer(Leng / 2) = 0                  ' end of buffer

End Function


Public Sub SetUpActiveList(sWhich As String)

    With frmCtrl
    
    .cboActive.Clear
    
        Select Case sWhich
        
            Case "Clients"
             .cboActive.AddItem "Active Clients"
             .cboActive.AddItem "InActive Clients"
    
            Case "Products"
             .cboActive.AddItem "Active Products"
             .cboActive.AddItem "InActive Products"
    
            Case "PLUs"
             .cboActive.AddItem "Active PLUs"
             .cboActive.AddItem "InActive PLUs"
    
            Case "Groups"
             .cboActive.AddItem "Active Groups"
             .cboActive.AddItem "InActive Groups"
    
            Case "Product/PLUs"
             .cboActive.AddItem "Client PLUs and Stock Products"
             .cboActive.AddItem "Client PLUs"
             .cboActive.AddItem "Client Stock Products"
        
        End Select

    End With

End Sub

Public Function SetReportSize(sWhich As String)


    With frmStockWatch
   
        .picSummary.Visible = False
        
        Select Case sWhich
            
            Case "Till"
            
             If .grdCount.Rows * .grdCount.RowHeight(0) > .fraPrint.Height - 1500 Then
    
                .grdCount.Height = .fraPrint.Height - 1500
                
                If .grdCount.Cols > 15 Then
                    .grdCount.Width = .Width - .picStatus.Width - 400
                    .grdCount.ScrollBars = flexScrollBarBoth

                Else
                    .grdCount.Width = .grdCount.ColPos(.grdCount.Cols - 1) + .grdCount.ColWidth(.grdCount.Cols - 1) + 300
                End If
    
             Else
        
                .grdCount.Height = .grdCount.Rows * .grdCount.RowHeight(0) + 30
                
                If .grdCount.Cols > 15 Then
                    .grdCount.Width = .Width - .picStatus.Width - 370
                    .grdCount.ScrollBars = flexScrollBarBoth
            
                Else
                    .grdCount.Width = .grdCount.ColPos(.grdCount.Cols - 1) + .grdCount.ColWidth(.grdCount.Cols - 1) + 30
                End If
            
             End If
            
             .imgReport.Height = .grdCount.Height + 320
             .imgReport.Width = .grdCount.Width + 260
    
             .imgReport.Left = (.fraPrint.Width - .grdCount.Width) / 2
             .grdCount.Left = .picStatus.Width + (.Width - .picStatus.Width - .grdCount.Width) / 2 + 140
        
             .imgReport.Top = (.fraPrint.Height - .grdCount.Height) / 2
             .grdCount.Top = (.Height - .picStatus.Top - .grdCount.Height) / 2 + 1540
            
             .imgReport.Visible = True
             .grdCount.Visible = True
            
            Case "Analysis"
            
             .picSummary.Top = 810
             .grdCount.Top = .picStatus.Top + 910
             
             .picSummary.Height = 8985  ' fixed
             .picSummary.Width = 9495   ' fixed
             
             .grdCount.Width = .picSummary.Width - 300
             
             
' Ver 543 chenged from 12 to 14

             .grdCount.Height = 14 * .grdCount.RowHeight(0) + 30
             .grdCount.ScrollBars = flexScrollBarVertical
             .grdCount.SelectionMode = flexSelectionByRow
             
'             .grdCount.Height = .grdCount.Rows * .grdCount.RowHeight(0) + 30
             '------------------
    
             .picSummary.Left = (.fraPrint.Width - .grdCount.Width) / 2
             .grdCount.Left = .picStatus.Width + (.Width - .picStatus.Width - .grdCount.Width) / 2 + 140
        
             .imgReport.Visible = False
             .grdCount.Visible = True
             .picSummary.Visible = True
            
            
            Case Else
             
             If .grdCount.Rows * .grdCount.RowHeight(0) > .fraPrint.Height - 1500 Then
    
                .grdCount.Height = .fraPrint.Height - 1500
                .grdCount.Width = .grdCount.ColPos(.grdCount.Cols - 1) + .grdCount.ColWidth(.grdCount.Cols - 1) + 300
    
             Else
        
                .grdCount.Height = .grdCount.Rows * .grdCount.RowHeight(0) + 30
                .grdCount.Width = .grdCount.ColPos(.grdCount.Cols - 1) + .grdCount.ColWidth(.grdCount.Cols - 1) + 30
        
             End If
        
             .imgReport.Height = .grdCount.Height + 320
             .imgReport.Width = .grdCount.Width + 260
    
             .imgReport.Left = (.fraPrint.Width - .grdCount.Width) / 2
             .grdCount.Left = .picStatus.Width + (.Width - .picStatus.Width - .grdCount.Width) / 2 + 140
        
             .imgReport.Top = (.fraPrint.Height - .grdCount.Height) / 2
             .grdCount.Top = (.Height - .picStatus.Top - .grdCount.Height) / 2 + 1540
        
             .imgReport.Visible = True
             .grdCount.Visible = True
        
        End Select
        
    End With

End Function
Public Function RepSummaryAnalysis()
Dim rs As Recordset

'Dim lStockSales As Long
Dim dblStockSales As Double

Dim iRow As Integer
Dim iLastPLUNo As Integer
Dim sngSalesTotQty As Double
Dim curLowestPrice As Currency
'ver 531
Dim sngGlass As Double
Dim sngGlassDP As Double
Dim iGlass As Integer

    ' THIS IS NOW THE PRODUCT DEFICIT REPORT

    On Error GoTo ErrorHandler
    
    
    bHourGlass True
    
    frmStockWatch.btnShow.Caption = "Show Selected"
    
    frmStockWatch.grdCount.Rows = 2
    frmStockWatch.grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "", "Description"
    SetupCountField frmStockWatch, "Till", "Sales"
    SetupCountField frmStockWatch, "Stock", "Sales"
    SetupCountField frmStockWatch, "", "- Diff"
    SetupCountField frmStockWatch, "", "+ Diff"
    SetupCountField frmStockWatch, "", "Value"

'ver 310
    SetupCountField frmStockWatch, "", "Sel "
    
    DoEvents
    
    frmStockWatch.grdCount.Cols = frmStockWatch.grdCount.Cols + 1
    frmStockWatch.grdCount.ColHidden(frmStockWatch.grdCount.Cols - 1) = True
    frmStockWatch.grdCount.ColKey(frmStockWatch.grdCount.Cols - 1) = "abs"
    ' Special hidden column which will hold the negative and posative values without the sign
    ' so easier for sorting
    
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblclientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) LEFT JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " ORDER BY tblPLUs.txtDescription, tblClientProductPLUs.PLUGroupID", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
           
'           If rs("tblPLUs.txtDescription") = "SCOTCH" Then Stop
           
           
            'DP
            sngSalesTotQty = 0    ' Init
            
            If Not IsNull(rs("SalesQty")) Then
                sngSalesTotQty = rs("SalesQty")
            End If
            
            If Not IsNull(rs("SalesQtyDP")) Then
                sngSalesTotQty = sngSalesTotQty + rs("SalesQtyDP")
            End If
            ' Total Sales units
            '--------------
            
            'Ver351 Glasses
            
            If Not IsNull(rs("Glass")) Then
                iGlass = rs("Glass")
                If iGlass > 0 Then
                    If Not IsNull(rs("GlassQty")) Then
                        sngSalesTotQty = sngSalesTotQty + rs("GlassQty") / iGlass
                    End If
                    
                    If Not IsNull(rs("GlassQtyDP")) Then
                        sngSalesTotQty = sngSalesTotQty + rs("GlassQtyDP") / iGlass
                    End If
                End If
            End If
            
            
            If bDualPrice Then
            
                If Not IsNull(rs("SellPriceDP")) Then
                    If rs("SellPriceDP") > rs("SellPrice") Then
                        curLowestPrice = rs("SellPrice")
                    Else
                        curLowestPrice = rs("SellPriceDP")
                    End If
                Else
                    curLowestPrice = rs("SellPrice")
                End If
            Else
                curLowestPrice = rs("SellPrice")
            End If
            
'            If rs("PLUNumber") = 43 Then Stop
            
            If iLastPLUNo <> Val(rs("PLUNumber")) Then
                iLastPLUNo = Val(rs("PLUNumber"))
                dblStockSales = GetStockSales(rs("PLUNumber"))
                frmStockWatch.grdCount.AddItem rs("tblPLUs.txtDescription") & vbTab & sngSalesTotQty & vbTab & dblStockSales
                
                If dblStockSales <= sngSalesTotQty Then
            
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("+Diff")) = "+" & Trim$(sngSalesTotQty - dblStockSales)
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Value")) = "+" & Format(curLowestPrice * Abs(Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("+Diff")))), "0.00")
                    
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("abs")) = Abs(Format(sngSalesTotQty - dblStockSales, "0.00"))
                    ' add absolute value to sorting cloumn
' ver 548 -------------
' Sort by value instead
'                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("abs")) = Abs(Format(curLowestPrice * Abs(Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("+Diff")))), "0.00"))
                
                Else
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("-Diff")) = "-" & Trim$(dblStockSales - sngSalesTotQty)
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Value")) = "-" & Format(curLowestPrice * Abs(Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("-Diff")))), "0.00")
                
                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("abs")) = Abs(Format(dblStockSales - sngSalesTotQty, "0.00"))
                    ' add absolute value to sorting cloumn
                
'                    frmStockWatch.grdCount.Cell(flexcpText, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("abs")) = Abs(Format(curLowestPrice * Abs(Val(frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("-Diff")))), "0.00"))
'---------------------


                End If
                
                frmStockWatch.grdCount.RowData(frmStockWatch.grdCount.Rows - 1) = Trim$(rs("tblPLUs.ID") + 0)
                
                frmStockWatch.grdCount.Cell(flexcpData, frmStockWatch.grdCount.Rows - 1, 0) = rs("PLUGroupID") + 0
                ' save the group ID here for counting group totals later
                
            End If
            
            rs.MoveNext
        
        Loop While Not rs.EOF
    
    
    End If
    rs.Close
    
    iRow = 2
    With frmStockWatch
        Do
            
            If .grdCount.RowData(iRow) > 0 Then
            ' check for data row
    
                If (Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("TillSales"))) = 0) And _
                    (Val(.grdCount.Cell(flexcpTextDisplay, iRow, .grdCount.ColIndex("StockSales"))) = 0) _
                Then
                        .grdCount.RemoveItem iRow
                Else
                'Debug.Print .grdCount.Cell(flexcpTextDisplay, iRow, 0)
                
                    iRow = iRow + 1
                End If
    
            Else
                iRow = iRow + 1
            End If
    
        Loop While iRow < .grdCount.Rows

        
        If .grdCount.Rows > 2 Then
        
            .grdCount.Select 2, .grdCount.ColIndex("abs"), .grdCount.Rows - 1, .grdCount.ColIndex("abs")
            .grdCount.Sort = flexSortNumericDescending
            .grdCount.Select 0, 0, 0, 0
            ' Sort the grid and remove selection afterwards
        
        End If
        
    
    
    End With
    
    bHourGlass False
    
    gbOk = SetReportSize("Analysis")
    
    frmStockWatch.grdCount.Cols = frmStockWatch.grdCount.Cols - 1
    ' Remove the Sorting column here not needed anymore and it affects the sorting on
    ' the next line
    
    gbOk = SetColWidths(frmStockWatch, "grdCount", "Description", True)
    
'Ver 548 removed again!!!!!!

'    If frmStockWatch.grdCount.Rows > 2 Then
'
'        frmStockWatch.grdCount.Select 2, frmStockWatch.grdCount.ColIndex("Value"), frmStockWatch.grdCount.Rows - 1, frmStockWatch.grdCount.ColIndex("Value")
'        frmStockWatch.grdCount.Sort = flexSortNumericDescending
'        frmStockWatch.grdCount.Select 0, 0
'    End If
    
    
    
    frmStockWatch.labelTitle.Caption = "Product Deficit"
    frmStockWatch.labelTitle.Tag = "G"
    RepSummaryAnalysis = True


CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepSummaryAnalysis") Then Resume 0
    Resume CleanExit

End Function


Public Function GetSummaryDetails(lID As Long, bShowSelected As Boolean)
Dim rs As Recordset
Dim iPntr As Integer
Dim sPicked As String
Dim sPLUID As String
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        If InStr(rs("SummaryNote") & "", "~") > 0 Then
            GetSummaryDetails = Mid(rs("SummaryNote"), InStr(rs("SummaryNote"), "~") + 1, Len(rs("SummaryNote")))
        End If
    End If
    
    If bShowSelected Then
    
        If InStr(rs("SummaryNote"), "|") <> 0 Then
    
            iPntr = 0
            
            sPicked = Left(rs("SummaryNote"), InStr(rs("SummaryNote"), "~") - 1)
            
'            For iPntr = 0 To Int(Len(sPicked) / 3) - 1
'
'                frmStockWatch.grdCount.Cell(flexcpChecked, Val(Mid(sPicked, iPntr * 3 + 2, 2)), frmStockWatch.grdCount.ColIndex("Sel")) = flexChecked
'
'            Next
    
             ' |132|333|457|23|5680|234|7|
             
            If Left(sPicked, 1) = "|" Then sPicked = Mid(sPicked, 2, Len(sPicked))
            ' remove first pipe
             
            Do
                iPntr = InStr(sPicked, "|") ' get next pipe
                
                If iPntr > 0 And Len(sPicked) > 1 Then
                ' as long as there is a pipe and its not the last char
                   
                    sPLUID = Left(sPicked, InStr(sPicked, "|") - 1)
                    ' get the PLU ID
             
                    If frmStockWatch.grdCount.FindRow(sPLUID) > -1 Then
                        frmStockWatch.grdCount.Cell(flexcpChecked, frmStockWatch.grdCount.FindRow(sPLUID), frmStockWatch.grdCount.ColIndex("Sel")) = flexChecked
                    End If
                    
                    sPicked = Mid(sPicked, InStr(sPicked, "|") + 1)
                End If
                
            Loop While Len(sPicked) > 1

        End If
    End If
    rs.Close
    
    If frmStockWatch.grdCount.Cell(flexcpTextDisplay, frmStockWatch.grdCount.Rows - 1, 0) <> "" Then
        frmStockWatch.grdCount.Rows = frmStockWatch.grdCount.Rows + 1
    End If
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetSummaryDetails") Then Resume 0
    Resume CleanExit

End Function

Public Function SetEvaluationFlag(lSelClID As Long, iEval As Integer)

    SWdb.Execute "UPDATE tblDates SET CountStep = " & Trim$(iEval) & " WHERE ClientID = " & Trim$(lSelClID) & " AND InProgress = true"
    
    ' Using this field since it was already there and didnt have to add a new one
    
    
End Function

Public Function GetClientDefaultFee(lCLId As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClients")
    rs.Index = "ID"
    rs.Seek "=", lCLId
    If Not rs.NoMatch Then
                    
        
        GetClientDefaultFee = rs("tedFee")
                    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientDefaultFee") Then Resume 0
    Resume CleanExit



End Function

Public Function GetAuditOnDate(lDatesID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    If lDatesID > 0 Then
        rs.Seek "=", lDatesID
        If Not rs.NoMatch Then
            GetAuditOnDate = Format(rs("on"), sDMMYY)
        End If
    End If
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetAuditOnDate") Then Resume 0
    Resume CleanExit

End Function

Public Function GetDefaultInvoiceText()
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblInvoiceDefaults")
    rs.Index = "PrimaryKey"
    If Not rs.EOF Then
        rs.MoveFirst
    
        GetDefaultInvoiceText = rs("InvoiceText")
    End If
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetDefaultInvoiceText") Then Resume 0
    Resume CleanExit



End Function

Public Function GetReport(lDatesID As Long, sTitle As String)
Dim sReport As String
Dim sFrom As String
Dim sTo As String

    ' Changes made for Ver 2.5.1
        
        'Here we use the lDatesID to get the To Date of the report and this will point to
        ' the folder for the client.
        
    gbOk = GetDates(lDatesID, sFrom, sTo)
    
    If IsDate(sTo) Then
    ' First check for valid TO Date
    
        If VerifyClientFolderFile(sReport, frmStockWatch.lblClient.Tag, Replace(sTo, "/", "-"), Replace(sTitle, "/", "_") & ".Doc") Then
        ' Client Name \ To Date \ Report Title
        ' THEn look for valid folder name with TO Date....
        
            gbOk = ShowReport(frmStockWatch.lblClient.Tag, sReport)
            ' Show report
        
' Ver304 Shouldnt need this anymore

        Else
            LogMsg frmStockWatch, sReport & " Does Not Exist ", "Client: " & frmStockWatch.lblClient.Tag
            MsgBox "Report: " & sReport & "  does not exist"
        
        End If
    
    End If
    
End Function

Public Function GetLicenseInfo(sName As String, _
                                sAddress As String, _
                                sPhone As String, _
                                sEmail As String, _
                                dtExpiry As Date, _
                                iDays As Integer, _
                                iwarn As Integer)

Dim rs As Recordset
Dim sLic As String

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblFranchisee")
    rs.Index = "PrimaryKey"
    
    If Not rs.EOF Then
        
        sLic = Decrypt(rs("franchiseDetails"), sKey)
        ' unbundle it

        sLic = Replace(sLic, "@@", vbCrLf)
        ' fix up address lines

        sName = Mid(sLic, InStr(1, sLic, "<Name>") + 6, InStr(1, sLic, "/<Name>") - InStr(1, sLic, "<Name>") - 6)
        sAddress = Mid(sLic, InStr(1, sLic, "<Address>") + 9, InStr(1, sLic, "/<Address>") - InStr(1, sLic, "<Address>") - 9)
        sPhone = Mid(sLic, InStr(1, sLic, "<Phone>") + 7, InStr(1, sLic, "/<Phone>") - InStr(1, sLic, "<Phone>") - 7)
        sEmail = Mid(sLic, InStr(1, sLic, "<Email>") + 7, InStr(1, sLic, "/<Email>") - InStr(1, sLic, "<Email>") - 7)
        dtExpiry = Mid(sLic, InStr(1, sLic, "<Expiry>") + 8, InStr(1, sLic, "/<Expiry>") - InStr(1, sLic, "<Expiry>") - 8)
        iDays = Val(Mid(sLic, InStr(1, sLic, "<Days>") + 6, InStr(1, sLic, "/<Days>") - InStr(1, sLic, "<Days>") - 6))
        iwarn = Val(Mid(sLic, InStr(1, sLic, "<Warn>") + 6, InStr(1, sLic, "/<Warn>") - InStr(1, sLic, "<Warn>") - 6))
        
        GetLicenseInfo = True
    
    End If
    
    rs.Close
    

Leave:
    
    Exit Function


ErrorHandler:

    MsgBox "Not Properly Licensed"

    Resume Leave


End Function

Public Function SetAuditFileSent(sInv As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT INVNumber, InvNotYetSentToStockWatch FROM tblDates WHERE INVNumber = " & sInv)
    If Not (rs.EOF And rs.BOF) Then
        rs.Edit
        rs("InvNotYetSentToStockWatch") = False
        rs.Update
    
        SetAuditFileSent = True
    
    End If
    rs.Close
    

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("SetAuditFileSent") Then Resume 0
    Resume 0

End Function

Public Function DeleteLocalSummaryFile(sInvoiceSummaryFile As String)

    On Error GoTo ErrorHandler
    
    Kill sInvoiceSummaryFile
    
    DeleteLocalSummaryFile = True
    
Leave:
    Exit Function
    
ErrorHandler:
'    MsgBox "error: " & Error
    Resume Leave

End Function

Public Function ExtendExpiryDate(bEXP As Boolean)
Dim rs As Recordset
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer
Dim sLic As String

    On Error GoTo ErrorHandler
    
    
    If (gbRegion <> "SW1") And (gbRegion <> "SW2") Then
    ' as long as its not head office and not development mode
    
        If GetLicenseInfo(sName, sAddress, sPhone, sEmail, dtExpiry, iDays, iwarn) Then
        ' unbundle license
        
            
            Set rs = SWdb.OpenRecordset("SELECT [on] from tblDates WHERE INVNumber > 0 ORDER BY [ON] DESC")
            If Not (rs.EOF And rs.BOF) Then
        
                rs.MoveFirst
                
                dtExpiry = DateAdd("d", iDays, Format(rs("on"), sDMY))
                ' get latest date that an audit was carried out
            Else
                dtExpiry = DateAdd("d", iDays, Format(Now, sDMY))
                ' update date
            End If
            
            
            If bEXP Then
                sLic = "<Name>" & sName & "/<Name>" & _
                "<Address>" & Trim$(Replace(sAddress, vbCrLf, "@@")) & "/<Address>" & _
                "<Phone>" & sPhone & "/<Phone>" & _
                "<Email>" & sEmail & "/<Email>" & _
                "<Expiry>" & Trim$(dtExpiry) & "/<Expiry>" & _
                "<Days>" & Trim$(iDays) & "/<Days>" & _
                "<Warn>" & Trim$(iwarn) & "/<Warn>"
    
            Else
                sLic = "<Name>" & "LICENSE EXPIRED" & "/<Name>" & _
                "<Address>" & "Please Contact Stock Watch On" & "/<Address>" & _
                "<Phone>" & "091 442987" & "/<Phone>" & _
                "<Email>" & "---" & "/<Email>" & _
                "<Expiry>" & "---" & "/<Expiry>" & _
                "<Days>" & "0" & "/<Days>" & _
                "<Warn>" & "0" & "/<Warn>"
            
            End If
            
            Set rs = SWdb.OpenRecordset("tblFranchisee")
            rs.Index = "PrimaryKey"
            If Not rs.EOF Then
                rs.MoveFirst
                rs.Edit
            
                ' ENCRYPTED LICENSE
                rs("FranchiseDetails") = Encrypt(sLic, sKey)
                ' rebundle license
            
                rs.Update
            
            End If
            ' save
            
            rs.Close
        
            ExtendExpiryDate = True
    
        End If
     
    End If
     
     
CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("ExtendExpiryDate") Then Resume 0
    Resume 0


End Function

Public Function getRegionEmailAndXferLocation(sReg As String, sEmail As String, SW1 As Boolean)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblFranchisee")
    rs.Index = "PrimaryKey"
    
    If Not rs.EOF Then
        
        rs.MoveFirst
        sReg = rs("region") & ""
        sEmail = rs("SWEmail") & ""
        
    End If
    rs.Close
    
    If CurDir$ = "C:\Program Files\Microsoft Visual Studio\VB98" Then
    ' development environment
    
        SW1 = (sReg = "SW2")
        ' If true set the headoffice flag (test environment mode)
    
    Else
    ' live mode...
        SW1 = (sReg = "SW1")
    
    End If
    
    getRegionEmailAndXferLocation = True

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("getRegionEmailAndXferLocation") Then Resume 0
    Resume 0


End Function

Public Function SetOkState(bOkButtonEnabled As Boolean)

    SetOkState = (bOkButtonEnabled Or SW1)
    
End Function
Public Function GetBankInfo(sBank As String, _
                                sNameOnAccount As String, _
                                sAccountNo As String, _
                                sSortCode, _
                                sBIC, _
                                sIBAN)

Dim rs As Recordset
Dim sBnk As String

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblFranchisee")
    rs.Index = "PrimaryKey"
    
    If Not rs.EOF Then
        
        rs.MoveFirst
        
        sBnk = Decrypt(rs("BankInfo"), sKey)
        ' unbundle it

        sBank = Mid(sBnk, InStr(1, sBnk, "<Bank>") + 6, InStr(1, sBnk, "/<Bank>") - InStr(1, sBnk, "<Bank>") - 6)
        sNameOnAccount = Mid(sBnk, InStr(1, sBnk, "<NameOnAccount>") + 15, InStr(1, sBnk, "/<NameOnAccount>") - InStr(1, sBnk, "<NameOnAccount>") - 15)
        sAccountNo = Mid(sBnk, InStr(1, sBnk, "<AccountNo>") + 11, InStr(1, sBnk, "/<AccountNo>") - InStr(1, sBnk, "<AccountNo>") - 11)
        sSortCode = Mid(sBnk, InStr(1, sBnk, "<SortCode>") + 10, InStr(1, sBnk, "/<SortCode>") - InStr(1, sBnk, "<SortCode>") - 10)
        sBIC = Mid(sBnk, InStr(1, sBnk, "<BIC>") + 5, InStr(1, sBnk, "/<BIC>") - InStr(1, sBnk, "<BIC>") - 5)
        sIBAN = Mid(sBnk, InStr(1, sBnk, "<IBAN>") + 6, InStr(1, sBnk, "/<IBAN>") - InStr(1, sBnk, "<IBAN>") - 6)
        
        GetBankInfo = True
    
    End If
    
    rs.Close
    

Leave:
    
    Exit Function


ErrorHandler:

    Resume Leave


End Function

Public Function PrintTillNotes(sDate As String, lID As Long, bStockCountSheet As Boolean, iCopies As Integer)
Dim rs As Recordset
Dim iLastGroup As Integer
Dim sClient As String
Dim sAddr As String
Dim sFrom As String
Dim sTo As String
Dim bVertical As Boolean
Dim iLinesperpage As Integer
Dim iLinesCounter As Integer
Dim iPageNo As Integer
Dim sPages As String
Dim iTotPages As Integer
Dim iCopy As Integer
Dim iGrp As Integer

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblClients WHERE ID = " & Trim$(lSelClientID), dbOpenSnapshot)
    ' SQl return list of plus for this clinet sorted by plu#
    If Not (rs.EOF And rs.BOF) Then
        
      If rs("txtNotes") & "" <> "" Then
        
        sClient = GetClientName(lSelClientID, sAddr, True)
        gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
        
        iLinesperpage = 60
        ' page = 63 lines
        ' header = 6 lines
        
        With frmStockWatch
            .ComDlg.Copies = iCopies
            .ComDlg.CancelError = True
            .ComDlg.ShowPrinter
        End With
        
        bVertical = True
        Printer.Orientation = bVertical + 2
        Printer.FontName = "Courier New"
        Printer.FontSize = 12
        
        ' Display the Print dialog box
            
        sDate = Format(Now, "dd mmm yy hh:mm")
    
        iPageNo = 0
        iLinesCounter = 0
        iLastGroup = 0
        iGrp = 0
    
        
        Printer.FontUnderline = False
        Printer.FontBold = False
        
        ' PRINT HEADER
                    
                    
        iLinesCounter = PrintHeader(sClient, sAddr, sFrom, sTo, sDate, sPages, "Till Information")
        
        rs.MoveFirst
        
        Printer.Print ""
                
        Printer.Print rs("txtNotes")
    
        Printer.EndDoc
    
        PrintTillNotes = True
      
      End If
      
    End If
    
    
    rs.Close


CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If Err = 32755 Then Resume CleanExit
    
    If CheckDBError("PrintTillNotes") Then Resume 0
    Resume CleanExit

End Function

Public Function SendOutLookMail(sSubject As String, sSendTo As String, sMessage As String, sFile As String, bMultipleattachments As Boolean)

 Dim iRow As Integer
 Dim iPos As Integer

'         sSendTo = txtEmailTo
'        sSubject = txtSubj
'        sMessage = txtText
'        sAttachPath = sDBLoc
'

 'KB113033 How to Send a Mail Message Using Visual Basic MAPI Controls
 'MAPI constants from CONSTANT.TXT file:
 Const ATTACHTYPE_DATA = 0
 Const RECIPTYPE_TO = 1
 Const RECIPTYPE_CC = 2

 On Error GoTo errh

 'Open up a MAPI session:
 frmEmail.MAPISession1.DownLoadMail = False 'revised - sending mail
 frmEmail.MAPISession1.SignOn
 'Point the MAPI messages control to the open MAPI session:
 frmEmail.MAPIMessages1.SessionID = frmEmail.MAPISession1.SessionID

 'Create a new message
 frmEmail.MAPIMessages1.MsgIndex = -1 'revised
 frmEmail.MAPIMessages1.Compose

 'Set the subject of the message:
 frmEmail.MAPIMessages1.MsgSubject = sSubject
 'Set the message content:
 frmEmail.MAPIMessages1.MsgNoteText = sMessage

 'The following four lines of code add an attachment to the message,
 'and set the character position within the MsgNoteText where the
 'attachment icon will appear. A value of 0 means the attachment will
 'replace the first character in the MsgNoteText. You must have at
 'least one character in the MsgNoteText to be able to attach a file.

 iPos = 0

 If bMultipleattachments Then
    For iRow = 0 To frmEmail.grdReps.Rows - 1

           If frmEmail.grdReps.Cell(flexcpChecked, iRow, 2) = 1 Then


               frmEmail.MAPIMessages1.AttachmentIndex = iPos
               frmEmail.MAPIMessages1.AttachmentPosition = iPos
               'Set the type of attachment:
               frmEmail.MAPIMessages1.AttachmentType = ATTACHTYPE_DATA
               'Set the icon title of attachment:
    '            Me.MAPIMessages1.AttachmentName = GetReportName(cboReportDate.ItemData(frmEmail.cboReportDate.ListIndex), grdReps.Cell(flexcpTextDisplay, iRow, 1))
               'Set the path and file name of the attachment:
               frmEmail.MAPIMessages1.AttachmentPathName = GetReportName(frmEmail.cboReportDate.ItemData(frmEmail.cboReportDate.ListIndex), frmEmail.grdReps.Cell(flexcpTextDisplay, iRow, 1))

               iPos = iPos + 1

           End If

       Next

 Else
    ' ONE FILE
               frmEmail.MAPIMessages1.AttachmentIndex = iPos
               'Set the type of attachment:
               frmEmail.MAPIMessages1.AttachmentType = ATTACHTYPE_DATA
               'Set the icon title of attachment:
    '            Me.MAPIMessages1.AttachmentName = GetReportName(cboReportDate.ItemData(frmEmail.cboReportDate.ListIndex), grdReps.Cell(flexcpTextDisplay, iRow, 1))
               'Set the path and file name of the attachment:
               frmEmail.MAPIMessages1.AttachmentPathName = sFile

 End If


 'Set the recipients
 frmEmail.MAPIMessages1.RecipIndex = 0
 frmEmail.MAPIMessages1.RecipType = RECIPTYPE_TO
 frmEmail.MAPIMessages1.RecipDisplayName = sSendTo '4/22/03
 'Me.MAPImessages1.RecipAddress = sSendTo

 'MESSAGE_RESOLVENAME checks to ensure the recipient is valid and puts
 'the recipient address in MapiMessages1.RecipAddress
 'If the E-Mail name is not valid, a trappable error will occur.
 'Me.MAPImessages1.ResolveName 'comment out due to receiptent error w/ GW6.5 1/5/04
 'Send the message:
 frmEmail.MAPIMessages1.Send True 'revised


xit:
 'Close MAPI mail session:
 frmEmail.MAPISession1.SignOff

xit2:

 MsgBox "Email is Queued. Open Outlook and press Send/Receive to Send Email"


 SendOutLookMail = True

 Screen.MousePointer = 0
 Exit Function

errh:
 If Err.Number = 32053 Then Resume xit2
 MsgBox Err.Description, vbCritical, Err.Number
 Resume xit

 End Function

Public Function GetReportName(lDatesID As Long, sTitle As String)
Dim sReport As String
Dim sFrom As String
Dim sTo As String
        
        'Here we use the lDatesID to get the From Date of the report and this will point to
        ' the folder for the client.
        
    gbOk = GetDates(lDatesID, sFrom, sTo)
    
    
    If IsDate(sTo) Then
    ' First check for valid TO Date
    
        If VerifyClientFolderFile(sReport, frmStockWatch.lblClient.Tag, Replace(sTo, "/", "-"), Replace(sTitle, "/", "_") & ".Doc") Then
        ' Client Name \ To Date \ Report Title
        ' THEn look for valid folder name with TO Date....
        
            GetReportName = sReport
            ' return report name
        
' ver441 Remove this check since its now only the to date in use.

'''        ElseIf IsDate(sFrom) Then
'''        ' Otherwise check for valid from date
'''
'''            If VerifyClientFolderFile(sReport, frmStockWatch.lblClient.Tag, Replace(sFrom, "/", "-"), Replace(sTitle, "/", "_") & ".Doc") Then
'''            ' Client Name \ From Date \ Report Title
'''            ' look for valid folder name with FROM date
'''
'''                GetReportName = sReport
'''                ' return report name
'''
'''            End If
'''
        
        Else
            LogMsg frmStockWatch, sReport & " Does Not Exist ", "Client: " & frmStockWatch.lblClient.Tag
            MsgBox "Report: " & sReport & "  does not exist"
        
        End If
    
    End If
    
    
End Function

Public Function GetEmailDefaults()

        gbSMTP = GetSetting(App.Title, "Email", "smtp")
        gbPort = Val(GetSetting(App.Title, "Email", "port"))
        gbEmailfromAddress = GetSetting(App.Title, "Email", "from")

        gbSSL = Val(GetSetting(App.Title, "Email", "ssl"))
        gbUsername = GetSetting(App.Title, "Email", "username")
        gbPassword = GetSetting(App.Title, "Email", "password")
        gbSWEmail = GetSetting(App.Title, "Email", "swEmail")
    
End Function

Public Function SendInvoiceBySMTP()
Dim rs As Recordset
Dim iRow As Integer
Dim lobj_cdomsg As CDO.Message
Dim sInvoice As String
Dim sOnDate As String
Dim sFolderToDate As String
Dim sClient As String
Dim sTotalFee As String
Dim sFolder As String

    On Error GoTo ErrorHandler
    
    ' FIRST GET ALL EMAILS WHICH HAVE NOT BEEN SENT YET
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblDates INNER JOIN tblCLients ON tblDates.ClientID = tblClients.ID WHERE InvSMTPEmailNotSentYet = true")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Screen.MousePointer = 11
        
        On Error GoTo SendMail_Error:
        
        Do
        
            sInvoice = ""
            sOnDate = ""
            sClient = ""
            sFolderToDate = ""
            sTotalFee = ""
            sFolder = ""
            
            sInvoice = gbRegion & "_" & Trim$(rs("InvNumber"))
            sOnDate = Trim$(Format(rs("on"), "dd/mmm/yyyy"))
            sFolderToDate = Trim$(Format(rs("to"), "dd-mm-yy"))
            sClient = Replace(Trim$(rs("txtName")), " ", "_")
            sTotalFee = Trim$(rs("INVTotal") & "")
                
            sFolder = sDBLoc & "\" & Trim$(sClient) & "\" & sFolderToDate
                
            Set lobj_cdomsg = New CDO.Message
            lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = gbSMTP    ' email setting
            lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = gbPort ' email setting
            lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = gbSSL     ' email setting
            lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
            If gbUsername <> "" Then
                lobj_cdomsg.Configuration.Fields(cdoSendUserName) = gbUsername ' email setting
            End If
            
            If gbPassword <> "" Then
                lobj_cdomsg.Configuration.Fields(cdoSendPassword) = gbPassword ' email setting
            End If
            
            lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
            lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
            lobj_cdomsg.Configuration.Fields.Update
            lobj_cdomsg.To = gbSWEmail   ' globally set
            lobj_cdomsg.From = sFranchiseEmail  ' from License
            lobj_cdomsg.Subject = sInvoice & " " & sOnDate & " " & sClient & " " & Trim$(Format(sTotalFee, "0.00"))   ' passed
            lobj_cdomsg.TextBody = "Invoice Details"   ' passed
            lobj_cdomsg.AddAttachment (sFolder & "\Invoice.Doc")    ' passed
                
            lobj_cdomsg.Send
            ' If Not net connection then error trap to SendMail_Error
            ' on this line
            
            gbOk = ClearSMTPFlag(rs("tblDates.ID"))
    
    
            rs.MoveNext
            
        Loop While Not rs.EOF
        
        rs.Close
    
    End If

Leave:
    Set lobj_cdomsg = Nothing
    SendInvoiceBySMTP = True

    Screen.MousePointer = 0

    Exit Function
          
ErrorHandler:
    If CheckDBError("SendInvoiceBySMTP") Then Resume 0
    Resume Leave

SendMail_Error:
    Resume Leave

End Function

Public Function ClearSMTPFlag(lID As Long)
Dim rs As Recordset

    ' need to clear smtp flag
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        rs.Edit
        rs("InvSMTPEmailNotSentYet") = False
        rs.Update
    End If
    rs.Close
    
Leave:
    Exit Function
          
ErrorHandler:
    If CheckDBError("ClearSMTPFlag") Then Resume 0
    Resume Leave

End Function

Public Function TerminateWINWORD()
On Error Resume Next

    Dim Process As Object
'    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'WINWORD.exe'")
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'WINWORD.exe'")
        Process.Terminate
    Next

    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'WINWORD.exe *32'")
        Process.Terminate
    Next

End Function

Public Function TrimNull$(ByVal s$)
    Dim pos%
    
    pos = InStr(s, Chr$(0))
    If pos Then s = Left$(s, pos - 1)
    
    TrimNull = Trim$(s)
End Function

Public Function RestartAgentProgram()
Dim Process As Object
    
    Set Process = GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'SWIAgent.exe'")

    If Process.Count = 0 Then
    
        Shell sDBLoc & "\SWIAgent.exe", vbNormalNoFocus
    End If
    
    
End Function

''''Public Function DualPriceFix(lID As Long, curFix As Currency)
''''Dim rs As Recordset
''''
''''
''''    ' VER 427 FIX
''''
''''    'from Kate's email:
''''    'What about the price on level 1 take away price on level 2 multiply by sales on level 2.
''''    ' Then on the discrepancy report add it to calculated sales!!!!!!! Not sure.
''''    '
''''
''''    ' as per conversation with Kate, 2nd price level will always be greater than 1st level.
''''
''''
''''    ' as per talk with Kate 3/7/13 Geraldine has client whos 1st prices are higher than 2nd
''''    ' so this is fix ver 428 below
''''
''''    curFix = 0
''''
''''    Set rs = SWdb.OpenRecordset("Select * FROM tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.productID = tblProducts.ID WHERE tblClientProductPLUs.ID = " & Trim$(lID))
''''    If Not (rs.EOF And rs.BOF) Then
''''        rs.MoveFirst
''''
''''        Do
''''
''''            If Not IsNull(rs("SellPrice")) Then
''''                If Not IsNull(rs("SellPriceDP")) Then
''''                    If Not IsNull(rs("SalesQtyDP")) Then
''''
''''''' VER 429 - 433 was involved in fixing wrong figures on the profit discrepancy report
''''' if there was more than one product added to the same plu .
'''''Also catered for here was if the level 1 price > level 2 price
''''
'''''''
''''
''''                        If rs("SellPriceDP") > rs("SellPrice") Then ' check diff anyway
''''                            curFix = curFix + (rs("SellPriceDP") - rs("SellPrice")) * rs("SalesQtyDP")
''''
''''                            DualPriceFix = True
''''
''''                        ' VER 428 FIX ----------------------------------------
''''
''''
''''                        ElseIf rs("SellPrice") > rs("SellPriceDP") Then
''''
''''                        ' Ver 429 fix changes the plus to a minus ------------
''''
'''''ORIGINAL LINE========================================================================================
''''                            curFix = curFix - (rs("SellPrice") - rs("SellPriceDP")) * rs("SalesQtyDP")
'''''=====================================================================================================
''''
''''                            DualPriceFix = True
''''
''''                        Else
'''''VER 430 removed this ... caused a problem when both prices were the same.
'''''                            curfix=0
''''' VER 431
''''                            curFix = curFix
''''
''''                            DualPriceFix = False
''''' VER 432
''''                        End If
''''                        '-----------------------------------------------------
''''
''''                    End If
''''                End If
''''            End If
''''
''''            rs.MoveNext
''''
''''        Loop While Not rs.EOF
''''
''''    End If
''''    rs.Close
''''
''''
''''
''''Leave:
''''    Exit Function
''''
''''ErrorHandler:
''''    If CheckDBError("DualPriceFix") Then Resume 0
''''    Resume Leave
''''
''''
''''End Function

Public Function GenerateReportCover(sName As String, sDate As String)
Dim objfile As Object
'Dim bFileExists As Boolean
Dim Reportdoc
Dim sReport As String
Dim sAddr As String
Dim sClient As String
Dim sFrom As String
Dim sTo As String

Dim sFranNAme As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer

    On Error GoTo ErrorHandler
    
    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    gbOk = TerminateWINWORD()
    
    Set WriteWord = New Word.Application
    
    bHourGlass True
    
    frmStockWatch.labelTitle.Caption = "Generating the Report Cover"
    
    sReport = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\Report Cover.Doc"
   
    On Error Resume Next
            
    Kill sReport
    
    On Error GoTo ErrorHandler
    
    WriteWord.Visible = False
    
    Set Reportdoc = WriteWord.Documents.Add(sDBLoc & "\Templates\Report Cover.dot")
    
    With WriteWord.ActiveDocument.Bookmarks
        
        sClient = GetClientName(lSelClientID, sAddr, True)
        .Item("Client").Range.Text = sClient
        .Item("Address").Range.Text = sAddr
        ' Client Name & Address
    
        .Item("Date").Range.Text = Format(Now, "dd mmm yy hh:mm")
        ' Date & Time
        
        gbOk = GetStockTakeDates(lDatesID, sFrom, sTo)
        .Item("From").Range.Text = sFrom
        .Item("To").Range.Text = sTo
        ' Stock Take Dates From & To
        
        If GetLicenseInfo(sFranNAme, sAddress, sPhone, sEmail, dtExpiry, iDays, iwarn) Then
            
            .Item("FranName").Range.Text = sFranNAme
            .Item("FranAddress").Range.Text = sAddress
            .Item("FranPhone").Range.Text = sPhone
            .Item("FranRegion").Range.Text = gbRegion
                
        End If

        WriteWord.Application.NormalTemplate.Saved = True

        WriteWord.Application.ActiveDocument.SaveAs (sReport)

        
        GenerateReportCover = True
    End With
    

CloseWordStuff:
    
    On Error Resume Next
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    Set objfile = Nothing
    Set Reportdoc = Nothing

CleanExit:
'
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GenerateReportCover") Then Resume 0
    Resume CloseWordStuff

End Function


Public Function DualPriceValueFix(curSell As Currency, curSellDP As Currency, curQtyDP As Variant)

' ver 452 This function added here 5th MAr
' to fix calc sales totals in group report
        
    If Not IsNull(curSell) Then
          If Not IsNull(curSellDP) Then
              If Not IsNull(curQtyDP) Then


                If curSellDP > curSell Then ' check diff anyway

                   DualPriceValueFix = (curSellDP - curSell) * curQtyDP

                ElseIf curSell > curSellDP Then
                
                    DualPriceValueFix = -(curSell - curSellDP) * curQtyDP
                
                Else
                
                    DualPriceValueFix = 0
                End If
          End If
      End If
    End If

End Function

Public Function TabletImportFileFound(lID As Long, sDrive As String)
    
    sDrive = voldrive("SWCount")
    
    If Dir$(sDrive & "SWCount_" & Trim$(lID) & "*.csv") <> "" Then


        TabletImportFileFound = True

    End If

End Function
Public Function ImportTabletFile(lID As Long, sDrive As String)
Dim lInf As Long
Dim objfile
Dim sIn As String
Dim sFile, sf
Dim sFileName As String
Dim rs As Recordset
Dim lClientProdPLUID As Long
Dim dblFullQty As Double
Dim dblOpen As Double
Dim dblWeight As Double
Dim rsBar As Recordset
Dim lBarID As Long


'Dim lClientID As Long
'Dim sClient As String
    
    Set objfile = CreateObject("Scripting.FileSystemObject")

    '     New function in version 500

    lInf = FreeFile
    
    If Not bMultipleBars Then
    ' check if multibar then....
    
    ' No
    
    ' so just try open basic file
    
'        On Error GoTo MainFileFoundError
    
        sFileName = Dir$(sDrive & "SWCount_" & Trim$(lID) & "_0_*.csv")
        
        If sFileName <> "" Then
        ' This is main file with no individual bar counts
        
            
'            On Error GoTo CantOpenMainFileError
            
            Open sDrive & sFileName For Input As #lInf
            ' open new file

            ' READ IN NEW COUNT FIGURES
            
GetNextLine:
            Line Input #lInf, sIn
            If Left(sIn, 5) <> "!Item" Then GoTo GetNextLine
            
            Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
            rs.Index = "PrimaryKey"
            
            Do
                sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                lClientProdPLUID = Val(Left(sIn, InStr(sIn, ",") - 1))
                
                sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                dblFullQty = Val(Left(sIn, InStr(sIn, ",") - 1))
                
                sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                dblOpen = Val(Left(sIn, InStr(sIn, ",") - 1))
                 
                sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                dblWeight = Val(sIn)
                ' Pick off values
            
                rs.Seek "=", lClientProdPLUID
                If Not rs.NoMatch Then
                    rs.Edit
                    rs("FullQty") = dblFullQty
                    rs("Open") = dblOpen
                    rs("Weight") = dblWeight
                    rs.Update
                End If
                ' Get the record in question
                
                Line Input #lInf, sIn
                If sIn = "!END" Then Exit Do
    
            Loop
        
            Close #lInf
        End If
    
        ImportTabletFile = True
    
    Else
    ' Yes
    
'        'FIRST CLEAR BAR COUNTS
'
'        SWdb.Execute "DELETE * FROM tblBarCount WHERE ClientID = " & Trim$(lID)
        
        Set rsBar = SWdb.OpenRecordset("SELECT * FROM tblBars")
'        Set rsBar = SWdb.OpenRecordset("SELECT * FROM tblBars WHERE ClientID = " & Trim$(lID))
        If Not (rsBar.EOF And rsBar.BOF) Then
            rsBar.MoveFirst
            Do
                
                sFileName = Dir$(sDrive & "SWCount_" & Trim$(lID) & "_" & rsBar("ID") & "_*.csv")
            
                If sFileName <> "" Then
                ' This is main file with no individual bar counts
                
    '            On Error GoTo CantOpenBarFileError
                
                    Open sDrive & sFileName For Input As #lInf
                    ' open new file
    
GetNextBarLine:
                    Line Input #lInf, sIn
                    If Left(sIn, 4) <> "!Bar" Then GoTo GetNextBarLine
                    
                    sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                    lBarID = Val(Left(sIn, InStr(sIn, ",") - 1))
                    
                    'FIRST CLEAR THE BAR COUNT
        
                    SWdb.Execute "DELETE * FROM tblBarCount WHERE ClientID = " & Trim$(lID) & " AND BARID = " & Trim$(lBarID)
    
GetNextItemLine:
                    If sIn = "!END" Then GoTo close_resume_next_file
                    Line Input #lInf, sIn
                    If Left(sIn, 5) <> "!Item" Then GoTo GetNextItemLine
                    
                    Do
                        sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                        lClientProdPLUID = Val(Left(sIn, InStr(sIn, ",") - 1))
                        
                        sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                        dblFullQty = Val(Left(sIn, InStr(sIn, ",") - 1))
                        
                        sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                        dblOpen = Val(Left(sIn, InStr(sIn, ",") - 1))
                         
                        sIn = Mid(sIn, InStr(sIn, ",") + 1, Len(sIn))
                        dblWeight = Val(sIn)
                        ' Pick off values
                    
                        Set rs = SWdb.OpenRecordset("SELECT * FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(lClientProdPLUID) & " AND BarID = " & Trim$(lBarID) & " AND ClientID = " & Trim$(lID))
                        If Not (rs.EOF And rs.BOF) Then
                    
                            rs.MoveFirst
                            rs.Edit
                            rs("BarFullQty") = dblFullQty + rs("BarFullQty")
                            rs("BarOpen") = dblOpen + rs("BarOpen")
                            rs("BarWeight") = dblWeight + rs("BarWeight")
                        
                        Else
                            rs.AddNew
                            rs("ClientProdPLUID") = lClientProdPLUID
                            rs("BarID") = lBarID
                            rs("ClientID") = lID
                            rs("BarFullQty") = dblFullQty
                            rs("BarOpen") = dblOpen
                            rs("BarWeight") = dblWeight
                        End If
                        
                        rs.Update
                        ' Get the record in question
                        
                        Line Input #lInf, sIn
                        If sIn = "!END" Then Exit Do
            
                    Loop

close_resume_next_file:
                    Close #lInf
                End If
                rsBar.MoveNext
            Loop While Not rsBar.EOF
        End If
        rsBar.Close
        rs.Close
        
    
        ' Add Up bar counts and add to clientproductsplus table
        
        Set rs = SWdb.OpenRecordset("SELECT ID FROM tblClientProductPLUs WHERE ClientID = " & Trim$(lID))
        If Not (rs.EOF And rs.BOF) Then
        ' First loop on all bars for the client
            rs.MoveFirst
            Do
            
                Set rsBar = SWdb.OpenRecordset("SELECT SUM(BarFullQty) AS FullQtyTotal, sum(BarOpen) as OpenTotal, sum(BarWeight) as weighttotal FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(rs("ID")))
                If Not (rsBar.EOF And rsBar.BOF) Then
                    rsBar.MoveFirst
                    Do
                        If Not IsNull(rsBar("FullQtyTotal")) Then
                            SWdb.Execute "UPDATE tblClientProductPLUs Set FullQty = " & Trim$(rsBar("FullQtyTotal")) & " WHERE ID = " & Trim$(rs("ID"))
                        
                        Else
                            SWdb.Execute "UPDATE tblClientProductPLUs Set FullQty = '' WHERE ID = " & Trim$(rs("ID"))
                        End If
            
'#########################
' these lines added for test

                        If Not IsNull(rsBar("OpenTotal")) Then
                            SWdb.Execute "UPDATE tblClientProductPLUs Set Open = " & Trim$(rsBar("OpenTotal")) & " WHERE ID = " & Trim$(rs("ID"))
                        
                        Else
                            SWdb.Execute "UPDATE tblClientProductPLUs Set Open = '' WHERE ID = " & Trim$(rs("ID"))
                        End If
                        
                        If Not IsNull(rsBar("WeightTotal")) Then
                            SWdb.Execute "UPDATE tblClientProductPLUs Set Weight = " & Trim$(rsBar("WeightTotal")) & " WHERE ID = " & Trim$(rs("ID"))
                        
                        Else
                            SWdb.Execute "UPDATE tblClientProductPLUs Set Weight = '' WHERE ID = " & Trim$(rs("ID"))
                        End If
                        
'#########################
                        rsBar.MoveNext
                    Loop While Not rsBar.EOF
                End If
            
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    
        ImportTabletFile = True
    
    End If
    
End Function


Function voldrive(targetvol As String) As String
Dim r  As Long, allDrives As String, JustOneDrive As String, pos As Integer, DriveType As Long
Dim CDfound As Integer, aronedrive() As String, d As Integer
allDrives = Space(64)
r = GetLogicalDriveStrings(Len(allDrives), allDrives)
allDrives = Left(allDrives, r)
aronedrive = Split(allDrives, vbNullChar)
For d = UBound(aronedrive) To 1 Step -1
    If getdrivetype(aronedrive(d)) = DRIVE_REMOVABLE Then
        If InStr(1, Dir(aronedrive(d), vbVolume), targetvol, vbTextCompare) > 0 Then
            voldrive = aronedrive(d)
            Exit Function
        End If
    End If
Next
voldrive = ""
End Function

Public Function ImportNewProductsFound(lID As Long, sDrive As String)
    
    sDrive = voldrive("SWCount")
    
    If Dir$(sDrive & "SWCount_" & Trim$(lID) & "NEW_PRODUCTS.csv") <> "" Then


        ImportNewProductsFound = True

    End If
    

End Function

Public Function DeleteImportedFiles(lID As Long, sDrive As String)


    On Error Resume Next

    Kill sDrive & "SWCount_" & Trim$(lID) & "_*.csv"

    DeleteImportedFiles = True
    
End Function

Public Function GetCostOfSales(lID As Long) As Double
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbCostOfSales As Double
Dim dbCostOfSalesTotal As Double
Dim dbCostOfSalesGrand As Double
Dim dbCalcAmount As Double

Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double

Dim curCost As Currency

Dim lLastProd As Long
Dim dblLastQty As Double
Dim curFix As Currency

    ' NEW function added 31/12/2014 as per Kates phone call 30/12 on cost of sales being different
    ' in group vs discrepancy reports   for the G hotel and some others.
    
    ' Dont know why this is needed since reports look correct for all clients I tried here.
    

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE ClientID = " & Trim$(lID) & " Order By cboGroups, PLUnumber, tblProducts.ID, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
            If iLastGroup <> rs("cboGroups") Then
                
                iLastGroup = rs("cboGroups")
                
                dbCostOfSalesTotal = 0

            End If
            
            If lLastProd <> rs("tblProducts.ID") Then

                lLastProd = rs("tblProducts.ID")
                
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
                ' Deliveries
                
                If Not IsNull(rs("FullQty")) Then

                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    
                    dbCostOfSales = (dblLastQty * rs("PurchasePrice")) + (dblCostDel * curCost) - (dbCalcAmount * rs("PurchasePrice"))
                    dbCostOfSalesTotal = dbCostOfSalesTotal + dbCostOfSales
                    ' Cost of Sales
            
                    
                    dbCostOfSalesGrand = dbCostOfSalesGrand + dbCostOfSales

                End If

            End If
            
            rs.MoveNext
            
        Loop While Not rs.EOF
        
    End If
        
    GetCostOfSales = dbCostOfSalesGrand
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetCostOfSales") Then Resume 0
    Resume CleanExit


End Function

Public Function getGlasses(lPLUID As Long) As Integer
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblPLUGroup")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lPLUID
    If Not rs.NoMatch Then
        If Not IsNull(rs("Glass")) Then
            getGlasses = rs("Glass") + 0
        Else
            getGlasses = 0
        End If
        
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("getGlasses") Then Resume 0
    Resume CleanExit

End Function
Public Function CheckForEvaluation(lID As Long) As Boolean
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT CountStep FROM tblDates WHERE ClientID = " & Trim$(lID) & " AND InProgress = true", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        
        rs.MoveFirst
        
        If rs("CountStep") <> 0 Then
        
            SetMenuEvaluation True
            
            CheckForEvaluation = True
        End If
        
    End If
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing

    Exit Function

ErrorHandler:
    If CheckDBError("CheckForEvaluation") Then Resume 0
    Resume CleanExit



End Function
Public Sub SetMenuEvaluation(bhow As Boolean)

    frmStockWatch.grdMenu.Cell(flexcpForeColor, 3, 0, 3, 1) = &HD1C7C5
    
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 6, 0, 6, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 7, 0, 7, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 8, 0, 8, 1) = &HD1C7C5
    
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 10, 0, 10, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 11, 0, 11, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 12, 0, 12, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 13, 0, 13, 1) = &HD1C7C5
    
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 15, 0, 15, 1) = &HD1C7C5
    frmStockWatch.grdMenu.Cell(flexcpForeColor, 16, 0, 16, 1) = &HD1C7C5

End Sub


Public Function GlassPriceValueFix(iGlass As Integer, curSell As Currency, curGlass As Currency, curSellDP As Currency, curGlassDP As Currency, GlassQty As Long, GlassQtyDP As Long, SellQtyDP As Long)
                               '   iIssue            rs("SellPrice")         rs("GlassPrice"),     rs("SellPriceDP"),     rs("GlassPriceDP"),    rs("GlassQty"),    rs("GlassQtyDP")
    
Dim GlassFix As Currency
Dim PintFix As Currency

    ' GLASS AT 1sT Level Price
    
    If Not IsNull(curGlass) Then
          
        If Not IsNull(GlassQty) Then
          
            If (iGlass * curGlass) > curSell Then
          
                   GlassFix = ((iGlass * curGlass) - curSell) * (GlassQty / iGlass)
            End If
        End If
    End If
    
    GlassPriceValueFix = GlassFix
          
          
    ' Glass at 2nd Level Price
    
    GlassFix = 0
    If Not IsNull(curGlassDP) Then
          
        If Not IsNull(GlassQtyDP) Then
          
            If (iGlass * curGlassDP) > curSellDP Then
          
                   GlassFix = ((iGlass * curGlassDP) - curSellDP) * (GlassQtyDP / iGlass)
            End If
        End If
    End If
    
    GlassPriceValueFix = GlassPriceValueFix + GlassFix
          
     ' Pint at 2nd Level Price
    
    PintFix = 0
    If Not IsNull(curSellDP) Then
          
        If Not IsNull(SellQtyDP) Then
          
            If (curSellDP) > curSell Then
          
                   PintFix = ((curSellDP) - curSell) * (SellQtyDP)
            End If
        End If
    End If
    
    GlassPriceValueFix = GlassPriceValueFix + PintFix
         
          
          
          
          
          
End Function

Public Function RepProfitDiscrepance()
Dim rs As Recordset
Dim iLastGroup As Integer
Dim dbSellExVat As Double
'Dim dbSalesExVat As Double
Dim dbCostOfSales As Double
Dim dbRetailValue As Double
Dim dbRetailValueTotal As Double
Dim dbCostOfSalesTotal As Double
Dim dbGrossProfitTotal As Double
Dim dbCalcSalesGrand As Double
Dim dbCostOfSalesGrand As Double

Dim dblDeliveries As Double
Dim dblCostDel As Double
Dim dblFreeDel As Double
Dim curCost As Currency

Dim lLastProd As Long

Dim curActual As Currency
Dim curStaff As Currency
Dim curCompDrinks As Currency
Dim curWastage As Currency
Dim curOverRings As Currency
Dim curPromotions As Currency
Dim curOffSales As Currency
Dim curVouchers As Currency
Dim curKitchen As Currency
Dim curOther As Currency
Dim sOther As String
Dim curProjectedCash As Currency
Dim curSubTotal As Currency
Dim sSign As String
Dim dbCalcAmount As Double
Dim dblLastQty As Double
Dim iIssue As Integer
Dim curSurplus As Currency
Dim sSurplus As String
Dim curCashDiff As Currency
Dim dbPotentialSalesGrand As Double

    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With frmStockWatch
    
    .grdCount.ForeColor = sBlack
    
    .grdCount.Rows = 1
    .grdCount.Cols = 0
    
    SetupCountField frmStockWatch, "Vat Inclusive", ""
    SetupCountField frmStockWatch, "Calculated", ""
    SetupCountField frmStockWatch, "Potential", ""
    SetupCountField frmStockWatch, "Actual", ""
    
    ' VAT INCLUSIVE
    
    frmStockWatch.btnCloseFraPrint.Left = frmStockWatch.fraPrint.Width - 350

   Set rs = SWdb.OpenRecordset("SELECT * FROM (((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " Order By cboGroups, PLUnumber, tblProducts.ID, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            

            
            If iLastGroup <> rs("cboGroups") Then
                
                iLastGroup = rs("cboGroups")
            
                dbRetailValueTotal = 0
                dbCostOfSalesTotal = 0
                dbGrossProfitTotal = 0

            End If
            
            If lLastProd <> rs("tblProducts.ID") Then

                lLastProd = rs("tblProducts.ID")
                
                dbSellExVat = Format(rs("SellPrice") / (1 + (sngvatrate / 100)), "Currency")
                ' Sell Ex-Vat
                
                dblDeliveries = GetDeliveries(rs("tblClientProductPLUs.ID"), dblCostDel, dblFreeDel, curCost)
                ' Deliveries

                If Not IsNull(rs("FullQty")) Then
                    
                    If Not IsNull(rs("LastQty")) Then dblLastQty = rs("lastQty") Else dblLastQty = 0
                    
                    dbCalcAmount = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    
                    dbRetailValue = Format((((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits")), "0")
                    
                    dbRetailValueTotal = dbRetailValueTotal + (dbRetailValue * rs("SellPrice"))
                    ' Retail Value
                
 '                   dbSalesExVat = Format(((dblLastQty + dblDeliveries) - dbCalcAmount) * rs("txtIssueUnits"), "0") * dbSellExVat
                    ' Sales Ex-Vat
            
                    dbCostOfSales = (dblLastQty * rs("PurchasePrice")) + (dblCostDel * curCost) - (dbCalcAmount * rs("PurchasePrice"))
                    dbCostOfSalesTotal = dbCostOfSalesTotal + dbCostOfSales
                    ' Cost of Sales
            
                    
                End If
            
                dbCostOfSalesGrand = dbCostOfSalesGrand + dbCostOfSales
                
            End If
            
            rs.MoveNext
        
            
        Loop While Not rs.EOF

        ' COST OF SALES
        
        gbOk = GetActualAndAllowances(lDatesID, _
                                        curActual, _
                                        curStaff, _
                                        curCompDrinks, _
                                        curWastage, _
                                        curOverRings, _
                                        curPromotions, _
                                        curOffSales, _
                                        curVouchers, _
                                        curKitchen, _
                                        curOther, _
                                        sOther, _
                                        curSurplus, _
                                        sSurplus)

        .grdCount.AddItem ""
        
'------------------------------------------------------
'ver 560 and whereever dbPotentialSalesGrand appears below
         dbPotentialSalesGrand = curActual + (curStaff + curCompDrinks + curWastage + curOverRings + curPromotions + curOffSales + curVouchers + curKitchen + curOther)
'------------------------------------------------------
        
'------------------------------------------------------
'Ver 557
        dbCalcSalesGrand = GetCalculatedSalesTOTAL()
'------------------------------------------------------

        .grdCount.AddItem "Sales" & vbTab & Format(dbCalcSalesGrand, "Currency") & _
                                    vbTab & Format(dbPotentialSalesGrand, "Currency") & _
                                    vbTab & Format(curActual, "Currency")
        

'-----------------------------------------------------

'Ver524 Add in function to retrieve the Cost Of Sales. This new function is exactly same as
'       what generates the group report.

        dbCostOfSalesGrand = GetCostOfSales(lSelClientID)
        
'-----------------------------------------------------

        .grdCount.AddItem "Cost of Sales" & _
                                    vbTab & Format(dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)), "Currency") & _
                                    vbTab & Format(dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)), "Currency") & _
                                    vbTab & Format(dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)), "Currency")

        .grdCount.AddItem "Gross Profit" & _
                                    vbTab & Format(dbCalcSalesGrand - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100))), "Currency") & _
                                    vbTab & Format(dbPotentialSalesGrand - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100))), "Currency") & _
                                    vbTab & Format(curActual - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100))), "Currency")
        
        If dbCalcSalesGrand <> 0 And curActual <> 0 Then
            .grdCount.AddItem "Gross Profit %" & _
                                    vbTab & Format(((dbCalcSalesGrand - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)))) / (dbCalcSalesGrand)) * 100, "0.00") & _
                                    vbTab & Format(((dbPotentialSalesGrand - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)))) / (dbPotentialSalesGrand)) * 100, "0.00") & _
                                    vbTab & Format(((curActual - (dbCostOfSalesGrand + (dbCostOfSalesGrand * (sngvatrate / 100)))) / (curActual)) * 100, "0.00")
        Else
            .grdCount.AddItem "Gross Profit %"
        End If
        
        
        .grdCount.AddItem ""
        .grdCount.AddItem ""
        
        .grdCount.AddItem "Calculated Sales" & vbTab & vbTab & vbTab & Format(dbCalcSalesGrand, "Currency")

        .grdCount.AddItem ""
        .grdCount.AddItem "Allowances -                     "
        .grdCount.AddItem ""
        .grdCount.AddItem "Staff Drinks" & vbTab & vbTab & Format(curStaff, "Currency")
        .grdCount.AddItem "Comp. Drinks" & vbTab & vbTab & Format(curCompDrinks, "Currency")
        .grdCount.AddItem "Wastage" & vbTab & vbTab & Format(curWastage, "Currency")
        .grdCount.AddItem "Mistakes" & vbTab & vbTab & Format(curOverRings, "Currency")
        .grdCount.AddItem "Promotions" & vbTab & vbTab & Format(curPromotions, "Currency")
        .grdCount.AddItem "Off Sales" & vbTab & vbTab & Format(curOffSales, "Currency")
        .grdCount.AddItem "Vouchers" & vbTab & vbTab & Format(curVouchers, "Currency")
        .grdCount.AddItem "Kitchen" & vbTab & vbTab & Format(curKitchen, "Currency")
        If Val(curOther) > 0 Then
            .grdCount.AddItem sOther & vbTab & vbTab & Format(curOther, "Currency")
        End If
        
        curSubTotal = curStaff + curCompDrinks + curWastage + curOverRings + curPromotions + curOffSales + curVouchers + curKitchen + curOther
        
        .grdCount.AddItem "Sub Total" & vbTab & vbTab & vbTab & Format(curSubTotal, "Currency")
        .grdCount.AddItem vbTab & vbTab & vbTab & "============"
        
        curProjectedCash = dbCalcSalesGrand - curSubTotal
        
        .grdCount.AddItem "Projected Cash Receipts" & vbTab & vbTab & vbTab & Format(curProjectedCash, "Currency")
        .grdCount.AddItem ""
        .grdCount.AddItem "Actual Cash Receipts" & vbTab & vbTab & vbTab & Format(curActual, "Currency")
        .grdCount.AddItem vbTab & vbTab & vbTab & "============"
        
'Ver 547 Surplus -----------------------------------------------------------
        

        If curProjectedCash > curActual Then
            curCashDiff = curProjectedCash - curActual
            sSign = "-"
            .grdCount.AddItem vbTab & vbTab & vbTab & sSign & Format(curCashDiff, "Currency")
        ElseIf curActual > curProjectedCash Then
            curCashDiff = curActual - curProjectedCash
            sSign = ""
            .grdCount.AddItem vbTab & vbTab & vbTab & Format(curCashDiff, "Currency")
        Else
            curCashDiff = 0
            sSign = ""
            .grdCount.AddItem vbTab & vbTab & vbTab & Format(curCashDiff, "Currency")
        End If

        .grdCount.AddItem ""
        .grdCount.AddItem "Surplus " & sSurplus & vbTab & vbTab & vbTab & Format(curSurplus, "Currency")
        .grdCount.AddItem ""

'Ver549 TEST - per email from Kate
'        .grdCount.AddItem "Cash" & vbTab & vbTab & sSign & Format(curCashDiff - curSurplus, "Currency")
'        If (dbCostOfSalesGrand > 0) And curActual > 0 Then
'            .grdCount.AddItem "Gross Variance %" & vbTab & vbTab & "% " & Format(((curCashDiff - curSurplus) / (curActual)) * 100, "0.00")

        .grdCount.AddItem "Cash" & vbTab & vbTab & vbTab & Format((sSign & curCashDiff) - curSurplus, "Currency")
        If (dbCostOfSalesGrand > 0) And curActual > 0 Then
            .grdCount.AddItem "Gross Variance %" & vbTab & vbTab & vbTab & "% " & Format((((sSign & curCashDiff) - curSurplus) / (curActual)) * 100, "0.00")
        ElseIf (dbCostOfSalesGrand > 0) And curActual > 0 Then

        Else
            .grdCount.AddItem "Gross Variance %" & vbTab & vbTab & vbTab & "% 0"

        End If

        .grdCount.AddItem ""
    
    
        
'Ver 547 'Surplus' -----------------------------------------------
        
'''        .grdCount.AddItem vbTab & vbTab & Format(curSurplus, "Currency")
'''''
'''''
'''''
'''''
'''        .grdCount.AddItem "Discrepancy -"
'''        .grdCount.AddItem ""
'''        .grdCount.AddItem "Cash" & vbTab & vbTab & Format(curActual - curProjectedCash - curSurplus, "Currency")
'''
'''        If curActual > curProjectedCash And (dbCostOfSalesGrand > 0) And curActual > 0 Then
'''            .grdCount.AddItem "Gross Profit" & vbTab & vbTab & "% " & Format(((curActual - curProjectedCash - curSurplus) / (curActual)) * 100, "0.00")
'''        ElseIf (dbCostOfSalesGrand > 0) And curActual > 0 Then
'''            .grdCount.AddItem "Gross Profit" & vbTab & vbTab & "% " & Format(((curProjectedCash - curActual - curSurplus) / (curActual)) * 100, "0.00")
'''        Else
'''            .grdCount.AddItem "Gross Profit" & vbTab & vbTab & "% 0"
'''        End If
 '-------------------------------------------------------
    End If

        
    .grdCount.ScrollBars = flexScrollBarBoth
    .grdCount.AutoSize 0, 2

    End With
    
    gbOk = SetReportSize("")
    
    bHourGlass False
    
    RepProfitDiscrepance = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RepProfitDiscrepance") Then Resume 0
    Resume CleanExit

    
End Function

Public Function CopyValuation(sName As String, sDate As String)
Dim sReport As String

    ' NEW IN Ver 555

    On Error GoTo ErrorHandler:
    
    sReport = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\Closing Stock.Doc"

    FileCopy sReport, sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sDate) & "\Stock Valuation.Doc"

Leave:
    Exit Function

ErrorHandler:
    MsgBox Trim$(Error)
    Resume Leave

End Function

