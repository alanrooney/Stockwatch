VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmEmail 
   BackColor       =   &H00DCD5BC&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Email Reports"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   Icon            =   "frmEmail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEmail.frx":1CCA
   ScaleHeight     =   5550
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8LCtl.VSFlexGrid grdReps 
      Height          =   2610
      Left            =   540
      TabIndex        =   3
      Top             =   1770
      Width           =   3765
      _cx             =   6641
      _cy             =   4604
      Appearance      =   2
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
      MousePointer    =   54
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12157534
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   14932961
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEmail.frx":83B8
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   1.2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   2
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
      BackColorFrozen =   12157534
      ForeColorFrozen =   12157534
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   6600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":84AA
               Key             =   "button"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":87A0
               Key             =   "inp"
               Object.Tag             =   "inp"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":8AB1
               Key             =   "correct"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":8C52
               Key             =   "yellowflag"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":8DB4
               Key             =   "redflag"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmail.frx":8F2E
               Key             =   "Tick"
            EndProperty
         EndProperty
      End
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1755
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   870
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   5190
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2850
      Width           =   5115
   End
   Begin VB.TextBox txtSubj 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5190
      MaxLength       =   100
      TabIndex        =   8
      Top             =   2010
      Width           =   5115
   End
   Begin VB.TextBox txtEmailTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9EEF3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5190
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1080
      Width           =   5115
   End
   Begin VB.ComboBox cboReportDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   570
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3795
   End
   Begin VB.CheckBox chkAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E7DEE6&
      Caption         =   "&All Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2835
      TabIndex        =   4
      Top             =   1470
      Width           =   1335
   End
   Begin MyCommandButton.MyButton cmdSend 
      Height          =   495
      Left            =   9165
      TabIndex        =   11
      Top             =   4770
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   873
      BackColor       =   13748165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   6805503
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14472636
      Caption         =   "&Send"
      CaptionPosition =   4
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton cmdCancel 
      Height          =   495
      Left            =   7860
      TabIndex        =   12
      Top             =   4770
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   873
      BackColor       =   13748165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14472636
      Caption         =   "&Cancel"
      CaptionPosition =   4
      ForeColorDisabled=   8421504
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   10350
      TabIndex        =   13
      Top             =   135
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
      BackColor       =   13748165
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   15200231
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   8421504
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14472636
      Caption         =   "X"
      CaptionPosition =   4
      ForeColorDisabled=   8421504
      ForeColorOver   =   192
      ForeColorFocus  =   13003064
      ForeColorDown   =   192
      PictureAlignment=   4
      GradientType    =   2
      TextFadeToColor =   8388608
   End
   Begin VB.Label lblGlow 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      TabIndex        =   14
      Top             =   1050
      Visible         =   0   'False
      Width           =   5190
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "&Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   9
      Top             =   2580
      Width           =   2745
   End
   Begin VB.Label lblSubj 
      BackStyle       =   0  'Transparent
      Caption         =   "Su&bject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   7
      Top             =   1740
      Width           =   2745
   End
   Begin VB.Label lblEmailTo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Email To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5250
      TabIndex        =   5
      Top             =   825
      Width           =   2745
   End
   Begin VB.Label lblReports 
      BackStyle       =   0  'Transparent
      Caption         =   "Select &Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1500
      Width           =   2595
   End
   Begin VB.Label lblReportDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Report &Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   0
      Top             =   825
      Width           =   2715
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSendTo As String
Dim sSubject As String, sMessage As String
Dim sAttachName As String, sAttachPath As String

Private Sub btnClose_Click()

    cmdCancel_Click

End Sub

Private Sub cboReportDate_Click()

    gbOk = GetSubject()

End Sub

Private Sub cboReportDate_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboReportDate_LostFocus()
    lblReportDate.ForeColor = sBlack

End Sub

Private Sub txtQty_Change()

End Sub

Private Sub chkAll_Click()

    gbOk = SelectReports(chkAll)

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdSend_Click()

    If CheckFieldsOK() Then

        gbOk = SendOutLookMail(txtSubj, txtEmailTo.Text, txtText, "", True)
    
    End If

End Sub

Private Sub Form_Activate()
    If txtEmailTo.Text = "" Then
        txtEmailTo.Locked = False
        txtEmailTo.BackColor = vbWhite
        bSetFocus Me, "txtEmailTo"
    Else
''        If InStr(txtText, "Customer") <> 0 Then
''            txtText.SelStart = 6
''            txtText.SelLength = 8
            bSetFocus Me, "txtText"
'        Else
'            bSetFocus Me, "cmdSend"
'        End If
        
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        cmdCancel_Click
        
    End If

End Sub

Private Sub Form_Load()

    gbOk = GetReportDates(lSelClientID)
    ' Get report Date List
    
    If cboReportDate.ListCount > 0 Then cboReportDate.ListIndex = 0
    ' default to last date
    
    chkAll.Value = 1
    ' set defaults
    
    gbOk = GetEmailAddressAndContact(lSelClientID)

    gbOk = SetupMenu()

End Sub

Private Sub grdReps_Click()
'    If grdReps.Col = 1 Then grdReps.Col = 0

End Sub

Private Sub grdReports_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub grdReports_LostFocus()
    lblReports.ForeColor = sBlack

End Sub

Private Sub grdReps_KeyPress(KeyAscii As Integer)
    grdReps.Row = grdReps.FindRow(UCase(Chr$(KeyAscii)), , 0)
        
    setClearSelection grdReps.Row, 2
        
End Sub


Public Sub setClearSelection(iRow As Integer, iCol As Integer)
Dim sRep As String
    
    If iRow > -1 And iCol = 2 Then
        
        
        If grdReps.Cell(flexcpChecked, iRow, iCol) = 2 Then
            grdReps.Cell(flexcpText, iRow, iCol) = True
            
        
        Else
            grdReps.Cell(flexcpText, iRow, iCol) = False
        
        End If
    End If


End Sub

Private Sub grdReps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setClearSelection grdReps.Row, 2

End Sub

Private Sub txtEmailTo_DblClick()
    
    SetUpEditEmail True

End Sub

Private Sub txtEmailTo_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtEmailTo_LostFocus()
    lblEmailTo.ForeColor = sBlack

    SetUpEditEmail False

End Sub

Private Sub txtSubj_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSubj_LostFocus()
    lblSubj.ForeColor = sBlack

End Sub

Private Sub txtText_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtText_LostFocus()
    lblText.ForeColor = sBlack

End Sub
Public Function GetReportDates(lCLId As Long)
Dim rs As Recordset
Dim iRow As Integer

    On Error GoTo ErrorHandler
    
    cboReportDate.Clear
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblDates WHERE ClientID = " & Trim$(lCLId) & " ORDER By From DESC", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        
        rs.MoveFirst
        
        Do
            
            cboReportDate.AddItem Trim$(Format(rs("To"), "ddd dd mmm yy"))
            cboReportDate.ItemData(cboReportDate.NewIndex) = rs("ID") + 0
            
            rs.MoveNext
        
        Loop While Not rs.EOF
    
    End If

    GetReportDates = True
    
CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("GetReportDates") Then Resume 0
    Resume CleanExit

End Function

Public Function SelectReports(bAll As Boolean)
Dim iRow As Integer

    For iRow = 0 To grdReps.Rows - 1

        grdReps.Cell(flexcpChecked, iRow, 2) = bAll

    Next
    

End Function

Public Function GetEmailAddressAndContact(lCLId As Long)
Dim rs As Recordset
Dim iRow As Integer

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClients")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lCLId
    
    If Not rs.NoMatch Then
        
        txtEmailTo.Text = Trim$(rs("txtEmail") & "")
        txtEmailTo.Tag = Trim$(rs("txtEmail") & "")
    
        txtText = Replace(txtText, "Customer", Trim$(rs("txtContact") & ""))
    
    End If
    rs.Close
    
CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("GetEmailAddress") Then Resume 0
    Resume CleanExit





End Function

Public Function CheckFieldsOK()

    If cboReportDate.ListIndex > -1 Then
    ' Check week selected
    
        If ReportSelected() Then
        ' check reports selected
        
            If EmailValid() Then
            ' check email valid
            
                If SubjOK() Then
                ' check subject
            
                    CheckFieldsOK = True
                
                Else
                    bSetFocus Me, "txtSubj"
                End If
    
            Else
                MsgBox "Please enter a valid Email Address (eg: name@service.ie)"
                SetUpEditEmail True
            End If
        Else
            MsgBox "Please select a report(s) to Email"
        End If
        
    Else
        If cboReportDate.ListCount = 0 Then
            MsgBox "there are no Reports Generated for this Client"
        Else
            MsgBox "Please Select a Report Date to Email"
            bSetFocus Me, "cboReportdate"
        End If
        
    End If
    
    
    
End Function


Public Function ReportSelected()
Dim iRow As Integer

    For iRow = 0 To grdReps.Rows - 1
        If grdReps.Cell(flexcpChecked, iRow, 2) = 1 Then
            ReportSelected = True
            Exit Function
        End If

    Next
End Function

Public Function EmailValid()

    If Len(Trim$(txtEmailTo)) > 0 Then
    
        If InStr(txtEmailTo, "@") > 0 Then
        
            If txtEmailTo.Text <> txtEmailTo.Tag Then
            
                If MsgBox("Update Client with New Email address?", vbYesNo + vbDefaultButton1 + vbQuestion, "Email Address Changed") = vbYes Then
                
                    gbOk = UpdateEmailAddress(lSelClientID)
            
                End If
            
            End If
            EmailValid = True
            
        End If
    End If
    
End Function

Public Sub SetUpEditEmail(bhow As Boolean)

    Select Case bhow
        
        Case True
         lblEmailTo.ForeColor = &HFF0000
         txtEmailTo.Locked = False
         txtEmailTo.BackColor = vbWhite
         bSetFocus Me, "txtEmailTo"
    
        Case Else
         txtEmailTo.Locked = True
         txtEmailTo.BackColor = sVryLtgGrey
    End Select
    
End Sub

Public Function UpdateEmailAddress(lCLId As Long)
Dim rs As Recordset
Dim iRow As Integer

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClients")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lCLId
    
    rs.Edit
    rs("txtEmail") = Trim$(txtEmailTo.Text)
    rs.Update
        
    rs.Close
    
CleanExit:
    Exit Function
    
ErrorHandler:
    If CheckDBError("UpdateEmailAddress") Then Resume 0
    Resume CleanExit

End Function

Public Function SubjOK()
    
    If txtSubj.Text = "" Then
        If MsgBox("No Subject, Continue Sending Email?", vbDefaultButton2 + vbYesNo + vbQuestion, "No Subject") = vbYes Then
            SubjOK = True
        End If
    Else
        SubjOK = True
    End If
    
End Function


Public Function GetSubject()

    txtSubj.Text = App.Title & " Reports For " & frmStockWatch.lblClient.Tag & " " & cboReportDate


End Function
Public Function SetupMenu()
Dim iRow As Integer

    For iRow = 0 To 8
        grdReps.Cell(flexcpPicture, iRow, 0) = frmStockWatch.imgList.ListImages("button").Picture
        grdReps.Cell(flexcpPictureAlignment, iRow, 0) = flexPicAlignCenterCenter

    Next

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveX = X
    MoveY = Y
    
    SetTranslucent Me.hwnd, 200
    
    bAllowMove = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bAllowMove Then
        Me.Move Me.Left + (X - MoveX), Me.Top + (Y - MoveY)
    End If
    



End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bAllowMove = False
    
    SetTranslucent Me.hwnd, 255

End Sub
