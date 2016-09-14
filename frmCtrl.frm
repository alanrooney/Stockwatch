VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmCtrl 
   BorderStyle     =   0  'None
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   ControlBox      =   0   'False
   Icon            =   "frmCtrl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCtrl.frx":1CCA
   ScaleHeight     =   9300
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboByGroup 
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
      ItemData        =   "frmCtrl.frx":8D48
      Left            =   7080
      List            =   "frmCtrl.frx":8D4F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.ComboBox cboActive 
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
      ItemData        =   "frmCtrl.frx":8D56
      Left            =   1920
      List            =   "frmCtrl.frx":8D5D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.OptionButton optProduct 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9FAFA&
      Caption         =   "&Product Groups"
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
      Height          =   345
      Left            =   2610
      TabIndex        =   12
      Top             =   840
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.OptionButton optPLU 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9FAFA&
      Caption         =   "P&LUs Groups"
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
      Height          =   345
      Left            =   4590
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00E9FAFA&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   9915
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2445
      Begin VB.TextBox txtSearch 
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
         Left            =   690
         TabIndex        =   7
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Searc&h"
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
         Left            =   15
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   998
      ImageHeight     =   1053
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCtrl.frx":8D64
            Key             =   "PLU"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCtrl.frx":FFA7
            Key             =   "Client"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCtrl.frx":17035
            Key             =   "Product/PLUs"
         EndProperty
      EndProperty
   End
   Begin MyCommandButton.MyButton cmdNew 
      Height          =   435
      Left            =   240
      TabIndex        =   14
      Top             =   795
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   767
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
      TransparentColor=   14215660
      Caption         =   "&New"
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
   Begin VSFlex8LCtl.VSFlexGrid grdList 
      Height          =   7500
      Left            =   240
      TabIndex        =   10
      Top             =   1590
      Visible         =   0   'False
      Width           =   14490
      _cx             =   25559
      _cy             =   13229
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
      MousePointer    =   54
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   11454186
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
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
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   14550
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
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
      TransparentColor=   14215660
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
   Begin MyCommandButton.MyButton btnShowAll 
      Height          =   360
      Left            =   5430
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
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
      TransparentColor=   14215660
      Caption         =   "Show All"
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
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
      Left            =   1830
      TabIndex        =   13
      Top             =   60
      Width           =   2805
   End
   Begin VB.Label lblByGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "&By Group"
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
      Left            =   6210
      TabIndex        =   4
      Top             =   885
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
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
      Left            =   12570
      TabIndex        =   8
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label labelCount 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Left            =   13140
      TabIndex        =   9
      Top             =   840
      Width           =   765
   End
   Begin VB.Label lblActive 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&View"
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
      Left            =   780
      TabIndex        =   1
      Top             =   885
      Width           =   1095
   End
End
Attribute VB_Name = "frmCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bAddNewPLUProduct As Boolean
Public bFormIsShown As Boolean
Public bClientProducts As Boolean

Private Sub btnClose_Click()

    Unload Me

End Sub

Private Sub btnShowAll_Click()
' ver440 added this button

    txtSearch.Text = ""
    DoEvents
    
    gbOk = ShowProducts(cboActive.ListIndex, "")

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If Not bFormIsShown Then
    
        gbOk = SetBackGroundColour(sMenuCtrl)

        If bClientProducts Then
            gbOk = WindowAppear(Me, 0, frmStockWatch.picStatus.Top + 200, frmStockWatch.picStatus.Width + 500, 3, False)
         
        Else
        
            gbOk = WindowAppear(Me, 0, (Screen.Height - Me.Height) / 2, (Screen.Width - Me.Width) / 2, 3, False)
        End If
        
        ' Ver 3.0.6 (8) Moved this inside if statement. click on save btn in plu/product form would move to next product ok
        ' but Alt S would refocus on search box ... this fixes it.
        
        Me.Top = 1690
        
        If cboActive.ListIndex = 0 Then bSetFocus Me, "txtSearch"
        
        bFormIsShown = True
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Public Sub cboActive_Click()

    grdList.Rows = 1
    DoEvents
    ' refresh quicker
    
         lblTitle.Caption = "- " & sMenuCtrl
    
    Select Case sMenuCtrl
    
        Case "Clients"
         
         btnShowAll.Visible = False
         grdList.BackColorAlternate = &HCEEDF4
         gbOk = ShowClients(cboActive.ListIndex)
         bSetFocus Me, "grdList"
        
        Case "Products"
         grdList.BackColorAlternate = &HC6DDD6
' Ver 440
'         gbOk = ShowProducts(cboActive.ListIndex, txtSearch)
         btnShowAll.Visible = True
         fraSearch.Visible = True
         bSetFocus Me, "txtSearch"
        
        Case "PLUs"
         btnShowAll.Visible = False
         grdList.BackColorAlternate = &HD2DCCF
         gbOk = ShowPLUs(cboActive.ListIndex, txtSearch)
         fraSearch.Visible = True
         bSetFocus Me, "txtSearch"
        
        Case "Groups"
         btnShowAll.Visible = False
         grdList.BackColorAlternate = 13498878
         gbOk = ShowGroups(cboActive.ListIndex)
         bSetFocus Me, "grdList"
         
        Case "PLUClient"
         btnShowAll.Visible = False
         grdList.BackColorAlternate = &HEFF2EE
         gbOk = ShowPLUGroups(cboActive.ItemData(cboActive.ListIndex))
         bSetFocus Me, "grdList"
         
        Case "Product/PLUs"
         btnShowAll.Visible = False
         gbOk = GetGroupList(cboActive.ListIndex, lSelClientID)
         cboByGroup.ListIndex = 0
         fraSearch.Visible = True
         bSetFocus Me, "txtSearch"
         
    End Select
    
    bHourGlass False

End Sub

Private Sub cboByGroup_Click()

    gbOk = ShowProductPLUs(cboActive.ListIndex, cboByGroup.ListIndex, txtSearch)

End Sub
Private Sub cmdNew_Click()
    
    ShowDetail 0

    bHourGlass False

    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    bClientProducts = False
    bFormIsShown = False
    
    frmStockWatch.picSelect.Visible = True

End Sub

Private Sub grdList_Click()
    
    If grdList.Row > -1 And grdList.Row < grdList.Rows Then
    
        ShowDetail grdList.RowData(grdList.Row)
        
    End If
    
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        If grdList.Row = 1 Then
            bSetFocus Me, "txtSearch"
        End If
    End If

End Sub

Private Sub grdList_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ShowDetail grdList.RowData(grdList.Row)
 
End Sub
Private Sub optPLU_Click()
    SetupCtrlGroup False

End Sub

Private Sub optProduct_Click()

    SetupCtrlGroup True
End Sub

Private Sub txtSearch_GotFocus()

    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        grdList.Row = 1
        bSetFocus Me, "grdList"
        
    End If
    

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'    KeyAscii = CharOk(KeyAscii, 2, " *-&%./") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case sMenuCtrl
        Case "Products"
         If txtSearch <> "" Then
         'ver440 added this check to make sure we're searching for something
            gbOk = ShowProducts(cboActive.ListIndex, txtSearch)
         End If
         
        Case "PLUs"
         gbOk = ShowPLUs(cboActive.ListIndex, txtSearch)
        
        Case "Product/PLUs"
         gbOk = ShowProductPLUs(cboActive.ListIndex, cboByGroup.ListIndex, txtSearch)
        
        Case Else
        
    End Select
    

End Sub

Private Sub txtSearch_LostFocus()
    lblSearch.ForeColor = sBlack

End Sub

Public Function ShowClients(iWhichClients As Integer)
Dim rs As Recordset
Dim sSql As String
    
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    grdList.Cols = 0
    SetupHeaderField frmCtrl, "Name"
    SetupHeaderField frmCtrl, "Address"
    SetupHeaderField frmCtrl, "Phone"
    SetupHeaderField frmCtrl, "Contact"
    SetupHeaderField frmCtrl, "Mobile"
    SetupHeaderField frmCtrl, "Fee"
    SetupHeaderField frmCtrl, "Email"
    SetupHeaderField frmCtrl, "Notes"
    ' setup grid passing which grid and which list

    
    If iWhichClients = 0 Then
        sSql = "WHERE chkActive = true"
    Else
        sSql = "WHERE chkActive = false"
    End If
    
    grdList.Rows = 1
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblClients " & sSql & " ORDER BY txtName", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("txtName") & vbTab & _
                            Replace(rs("rtfAddress"), vbCrLf, " ") & vbTab & _
                            rs("txtPhone") & vbTab & _
                            rs("txtContact") & vbTab & _
                            rs("txtMobile") & vbTab & _
                            Format(rs("tedFee"), "currency") & vbTab & _
                            rs("txtEmail") & vbTab & _
                            rs("txtNotes")

            
            grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    labelCount = ShowCount(Me, "grdList")
    
'    grdList.Width = 14500

'    grdList.Width = grdList.ColPos(grdList.Cols - 1) + grdList.ColWidth(grdList.Cols - 1) + 320

    gbOk = SetColWidths(frmCtrl, "grdList", "Notes", False)
'    grdList.Height = 9800
    
    bHourGlass False
    
    ShowClients = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowClients") Then Resume 0
    Resume CleanExit

End Function


Public Function ShowProducts(iWhichProducts As Integer, sSearch As String)
Dim rs As Recordset
Dim sSql As String
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    grdList.Cols = 0
    SetupHeaderField frmCtrl, "Code"
    SetupHeaderField frmCtrl, "Group"
    SetupHeaderField frmCtrl, "Description"
    SetupHeaderField frmCtrl, "Size"
    SetupHeaderField frmCtrl, "Issue Units"
    SetupHeaderField frmCtrl, "Full Weight"
    SetupHeaderField frmCtrl, "Empty Weight"
    SetupHeaderField frmCtrl, "Vat Rate"
'    SetupHeaderField frmCtrl, "Full Verify"
'    SetupHeaderField frmCtrl, "Empty Verify"
    ' setup grid passing which grid and which list

    
'    gbOk = SetColWidths(frmCtrl, "grdList", "Description", False)
    
    If iWhichProducts = 0 Then
        sSql = "WHERE chkActive = true"
    Else
        sSql = "WHERE chkActive = false"
    End If
    
    grdList.Rows = 1
    
    If sSearch = "" Then
        Set rs = SWdb.OpenRecordset("Select * FROM tblProducts " & sSql & " ORDER BY cboGroups, txtDescription", dbOpenSnapshot)
    Else

        Set rs = SWdb.OpenRecordset("Select * FROM tblProducts " & sSql & " AND txtDescription LIKE " & """" & sSearch & "*""" & " ORDER BY cboGroups, txtDescription", dbOpenSnapshot)
    End If
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("txtCode") & vbTab & _
                            GetGroup(rs("cboGroups")) & vbTab & _
                            rs("txtDescription") & vbTab & _
                            rs("txtSize") & "" & vbTab & _
                            rs("txtIssueUnits") & vbTab & _
                            rs("txtFullWeight") & vbTab & _
                            rs("txtEmptyWeight") & vbTab & _
                            rs("cboVatTable")
'                            rs("chkFullVerify") & vbTab & _
'                            rs("chkEmptyVerify")
            
            grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    labelCount = ShowCount(Me, "grdList")
    
'    grdList.Width = grdList.ColPos(grdList.Cols - 1) + grdList.ColWidth(grdList.Cols - 1) + 320
'    grdList.Height = 9800

    gbOk = SetColWidths(frmCtrl, "grdList", "Description", True)
    
    
    bHourGlass False
    
    ShowProducts = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowProducts") Then Resume 0
    Resume CleanExit


End Function

Public Function ShowProductPLUs(iWhich As Integer, iGrp As Integer, sSearch As String)
Dim rs As Recordset
Dim sSql As String
Dim lThisRec As Long
Dim sGrp As String
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    grdList.Rows = 1
    grdList.Cols = 0

'    sSql = "WHERE chkActive = true"
        
    Select Case iWhich
        Case 0
         
         SetupHeaderField frmCtrl, "PLU #"
         SetupHeaderField frmCtrl, "Key Group"
         SetupHeaderField frmCtrl, "Key Description"
         
         SetupHeaderField frmCtrl, "Sell 1"
         SetupHeaderField frmCtrl, "Gls 1"
         
         SetupHeaderField frmCtrl, "Sell 2"
         SetupHeaderField frmCtrl, "Gls 2"
         
'         SetupHeaderField frmCtrl, "--"
         SetupHeaderField frmCtrl, "Code"
         SetupHeaderField frmCtrl, "Group"
         SetupHeaderField frmCtrl, "Description"
         SetupHeaderField frmCtrl, "Size"
         SetupHeaderField frmCtrl, "Cost"
         SetupHeaderField frmCtrl, "A"
        
         frmCtrl.grdList.ColHidden(frmCtrl.grdList.ColIndex("Sell2")) = Not bDualPrice
         frmCtrl.grdList.ColHidden(frmCtrl.grdList.ColIndex("Gls2")) = Not bDualPrice
         If iGrp > 0 Then
            If Left(cboByGroup, 3) = "Plu" Then
                sGrp = " AND tblPLUGroup.ID = " & Trim$(cboByGroup.ItemData(cboByGroup.ListIndex))
            Else
                sGrp = " AND tblProductGroup.ID = " & Trim$(cboByGroup.ItemData(cboByGroup.ListIndex))
            End If
         Else
            sGrp = ""
         End If
         
         If Trim$(txtSearch.Text) <> "" Then
            'sSearch = " AND ((tblProducts.txtDescription LIKE '" & sSearch & "*' " & ") OR (tblPLUs.txtDescription LIKE '" & sSearch & "*'))"
            sSearch = " AND ((tblProducts.txtDescription LIKE " & """" & sSearch & "*""" & ") OR (tblPLUs.txtDescription LIKE " & """" & sSearch & "*""" & "))"
         Else
            sSearch = ""
         End If
          
         sSql = "SELECT Active, glassPrice, glassPriceDP, tblClientProductPLUs.ID, txtGroupNumber, tblClientProductPLUs.PLUNumber, tblClientProductPLUs.PLUGroupNo, tblPLUs.txtDescription, tblClientProductPLUs.SellPrice, tblClientProductPLUs.SellPriceDP, tblProducts.txtCode, tblProductGroup.txtDescription, tblProducts.txtDescription, tblProducts.txtSize, tblClientProductPLUs.PurchasePrice, tblClientProductPLUs.Active, tblProductGroup.txtDescription, tblPLUGroup.txtDescription FROM ((((tblClientProductPLUs INNER JOIN tblClients ON tblClientProductPLUs.ClientID = tblClients.ID) INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE (((tblClients.ID)= " & Trim$(lSelClientID) & ")) " & sGrp & sSearch & " ORDER BY PLUNumber;"
        
        Case 1
         
         SetupHeaderField frmCtrl, "PLU #"
         SetupHeaderField frmCtrl, "Group"
         SetupHeaderField frmCtrl, "Key Description"
         SetupHeaderField frmCtrl, "Sell"
'         SetupHeaderField frmStockWatch, "Stock Connection"
         
        
         If iGrp > 0 Then
             sGrp = " AND tblPLUGroup.ID = " & Trim$(cboByGroup.ItemData(cboByGroup.ListIndex))
         Else
            sGrp = ""
         End If
        
         If Trim$(txtSearch.Text) <> "" Then
            sSearch = " AND tblPLUs.txtDescription LIKE " & """" & sSearch & "*"""
         Else
            sSearch = ""
         End If
          
         
         sSql = "SELECT DISTINCTROW tblClientProductPLUs.ID, txtGroupNumber, tblClientProductPLUs.Active, tblClientProductPLUs.PLUNumber, tblClientProductPLUs.PLUGroupNo, tblPLUs.txtDescription, tblClientProductPLUs.SellPrice, tblClientProductPLUs.ClientID, tblPLUGroup.txtDescription FROM (tblPLUs INNER JOIN tblClientProductPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE (((tblClientProductPLUs.ClientID)= " & Trim$(lSelClientID) & ")) " & sGrp & sSearch & " ORDER BY PLUNumber;"
        
        Case 2
         
         SetupHeaderField frmCtrl, "Code"
         SetupHeaderField frmCtrl, "Group"
         SetupHeaderField frmCtrl, "Description"
         SetupHeaderField frmCtrl, "Size"
         SetupHeaderField frmCtrl, "Cost"
         SetupHeaderField frmCtrl, "Active"
        
         If iGrp > 0 Then
             sGrp = " AND tblProductGroup.ID = " & Trim$(cboByGroup.ItemData(cboByGroup.ListIndex))
         Else
            sGrp = ""
         End If
    
         If Trim$(txtSearch.Text) <> "" Then
            sSearch = " AND tblProducts.txtDescription LIKE " & """" & sSearch & "*"""
         Else
            sSearch = ""
         End If
          
         
         sSql = "SELECT DISTINCT tblClientProductPLUs.ID, tblClientProductPLUs.PLUNumber, tblClientProductPLUs.PLUGroupNo, tblClientProductPLUs.SellPrice, tblClientProductPLUs.SellPriceDP, tblProducts.txtCode, tblProductGroup.txtDescription, tblProducts.txtDescription, tblProducts.txtSize, tblClientProductPLUs.PurchasePrice, tblClientProductPLUs.Active FROM ((tblClientProductPLUs INNER JOIN tblClients ON tblClientProductPLUs.ClientID = tblClients.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE (((tblClients.ID)= " & Trim$(lSelClientID) & ")) " & sGrp & sSearch & " ORDER BY tblProductGroup.txtDescription;"
    
    End Select
    
    grdList.Rows = 1
    
    Set rs = SWdb.OpenRecordset(sSql, dbOpenSnapshot)

    Select Case iWhich
        Case 0
         If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                grdList.AddItem rs("PLUNumber") & vbTab & _
                            rs("txtGroupNumber") & "  " & rs("tblPLUGroup.txtDescription") & vbTab & _
                            rs("tblPLUs.txtDescription") & vbTab & _
                            Format(rs("SellPrice"), "0.00") & vbTab & _
                            Format(rs("GlassPrice"), "0.00") & vbTab & _
                            Format(rs("SellPriceDP"), "0.00") & vbTab & _
                            Format(rs("GlassPriceDP"), "0.00") & vbTab & _
                            rs("txtCode") & vbTab & _
                            rs("tblProductGroup.txtDescription") & vbTab & _
                            rs("tblProducts.txtDescription") & vbTab & _
                            rs("txtSize") & vbTab & _
                            Format(rs("PurchasePrice"), "0.00") & vbTab & _
                            rs("tblClientProductPLUs.Active")
                grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
                rs.MoveNext
            Loop While Not rs.EOF
            
            grdList.AutoSize 0, 5
            gbOk = SetColWidths(frmCtrl, "grdList", "Description", False)
            
         End If
    
        Case 1
         If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                If lThisRec <> rs("PLUNumber") Then
                    grdList.AddItem rs("PLUNumber") & vbTab & _
                            rs("txtGroupNumber") & "  " & rs("tblPLUGroup.txtDescription") & vbTab & _
                            rs("tblPLUs.txtDescription") & vbTab & _
                            Format(rs("SellPrice"), "0.00")
            
                    lThisRec = rs("PLUNumber")
                    
                    grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
                
                End If
                
                rs.MoveNext
            Loop While Not rs.EOF
            gbOk = SetColWidths(frmCtrl, "grdList", "KeyDescription", False)
         
         End If
    
        Case 2
         If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                
                grdList.AddItem rs("txtCode") & vbTab & _
                            rs("tblProductGroup.txtDescription") & vbTab & _
                            rs("tblProducts.txtDescription") & vbTab & _
                            rs("txtSize") & vbTab & _
                            Format(rs("PurchasePrice"), "0.00") & vbTab & _
                            rs("Active")
            
                grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
                rs.MoveNext
            Loop While Not rs.EOF
            gbOk = SetColWidths(frmCtrl, "grdList", "Description", False)

         End If

    End Select
    
'    grdList.Height = Screen.Height - fraCtrl.Height - picSelect.Height - 800
    
    labelCount = ShowCount(Me, "grdList")

    ShowProductPLUs = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowProductPLUs") Then Resume 0
    Resume CleanExit



End Function

Public Function ShowPLUGroups(lID As Long)
Dim rs As Recordset
Dim sSql As String

    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    grdList.Cols = 0
    SetupHeaderField frmCtrl, "Client"
    SetupHeaderField frmCtrl, "Group Number"
    SetupHeaderField frmCtrl, "Description"
    SetupHeaderField frmCtrl, "Active"

    If lID > 0 Then
        sSql = " WHERE ClientID = " & Trim$(lID)
        grdList.ColHidden(grdList.ColIndex("Client")) = True
        cmdNew.Enabled = True
    Else
        sSql = ""
        grdList.ColHidden(grdList.ColIndex("Client")) = False
        cmdNew.Enabled = False
    End If
    
    grdList.Rows = 1
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblPLUGroup INNER JOIN tblClients ON tblPLUGroup.ClientID = tblClients.ID " & sSql & " ORDER BY txtGroupNumber", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("txtName") & vbTab & _
                            rs("txtGroupNumber") & vbTab & _
                            rs("txtDescription") & vbTab & _
                            rs("tblPLUGroup.chkActive")

            grdList.Cell(flexcpData, grdList.Rows - 1, 0) = rs("tblClients.ID") + 0
            grdList.RowData(grdList.Rows - 1) = rs("tblPLUGroup.ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    labelCount = ShowCount(Me, "grdList")
    
'    grdList.Width = 14500
    
    gbOk = SetColWidths(frmCtrl, "grdList", "Description", True)
    
    
'    grdList.ColWidth(0) = grdList.Width
'    grdList.Height = 9800
    
    bHourGlass False
    
    ShowPLUGroups = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowPLUGroups") Then Resume 0
    Resume CleanExit


End Function
Public Function ShowPLUs(iWhich As Integer, sSearch As String)
Dim rs As Recordset
Dim sSql As String
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    grdList.Cols = 0
    SetupHeaderField frmCtrl, "Description"

'    gbOk = SetColWidths(frmCtrl, "grdList", "Description", False)
    
        
    sSql = "WHERE chkActive = " & Trim$(iWhich - 1)
    
    grdList.Rows = 1
    
    If sSearch = "" Then
        Set rs = SWdb.OpenRecordset("Select * FROM tblPLUs " & sSql & " ORDER BY txtDescription", dbOpenSnapshot)
    Else
        Set rs = SWdb.OpenRecordset("Select * FROM tblPLUs " & sSql & " AND txtDescription LIKE " & """" & sSearch & "*""" & " ORDER BY txtDescription", dbOpenSnapshot)
    End If
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("txtDescription")
            
            grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    labelCount = ShowCount(Me, "grdList")
    
    grdList.ColWidth(0) = grdList.Width
    ShowPLUs = True

CleanExit:
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit


End Function


Public Function ShowGroups(iWhich As Integer)
Dim rs As Recordset
Dim sSql As String
    
    On Error GoTo ErrorHandler
    
    cmdNew.Enabled = True
    
    bHourGlass True
    
    grdList.Cols = 0
    SetupHeaderField frmCtrl, "Group ID"
    SetupHeaderField frmCtrl, "Description"
    SetupHeaderField frmCtrl, "Open Item"

'    grdList.Width = 14500
    gbOk = SetColWidths(frmCtrl, "grdList", "Description", False)
    
        
    sSql = "WHERE chkActive = " & Trim$(iWhich - 1)
    
    grdList.Rows = 1
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblProductGroup " & sSql & " ", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdList.AddItem rs("ID") & vbTab & _
                            rs("txtDescription") & vbTab & _
                            rs("chkOpenItem")
            
            grdList.RowData(grdList.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    labelCount = ShowCount(Me, "grdList")
    
    
    
'    grdList.ColWidth(1) = grdList.Width
'    grdList.Height = 9800
    
    
    bHourGlass False
    
    ShowGroups = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowGroups") Then Resume 0
    Resume CleanExit


End Function

Public Function GetGroupList(iWhich As Integer, lCLId As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    cboByGroup.Clear
    
    cboByGroup.AddItem "All"
    
    If iWhich = 0 Or iWhich = 1 Then
    
        Set rs = SWdb.OpenRecordset("Select * FROM tblPLUGroup WHERE ClientID = " & Trim$(lCLId), dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                cboByGroup.AddItem "Plu - " & rs("txtDescription")
                cboByGroup.ItemData(cboByGroup.NewIndex) = rs("ID") + 0
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    End If
        
    If iWhich = 0 Or iWhich = 2 Then
        
        Set rs = SWdb.OpenRecordset("Select * FROM tblProductGroup WHERE chkActive = true", dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                cboByGroup.AddItem "Stock - " & rs("txtDescription")
                cboByGroup.ItemData(cboByGroup.NewIndex) = rs("ID") + 0
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUGroupID") Then Resume 0
    Resume CleanExit


End Function

Public Sub ShowDetail(lID As Long)
    
    Select Case sMenuCtrl
        Case "Clients"
         frmClientDetail.lClientID = lID
         frmClientDetail.sOldName = ""  ' just make sure its empty only needed in a rename
         frmClientDetail.Show vbModal
         gbOk = ShowClients(cboActive.ListIndex)

        Case "Products"
         frmProductDetail.lProductID = lID
         frmProductDetail.Show vbModal

        Case "PLUs"
         frmPLUDetail.lPLUID = lID
         frmPLUDetail.Show vbModal

        Case "Groups"
         frmGroupDetail.lProductGroupID = lID
         frmGroupDetail.Show vbModal
         gbOk = ShowGroups(cboActive.ListIndex)
    
        Case "PLUClient"
         frmPLUGroupDetail.lPLUGroupID = lID
         
         If lID > 0 Then
         
             frmPLUGroupDetail.lCLId = grdList.Cell(flexcpData, grdList.Row, 0)
             frmPLUGroupDetail.labelClient = grdList.Cell(flexcpTextDisplay, grdList.Row, 0)
         ElseIf cboActive.ListIndex > -1 Then
             frmPLUGroupDetail.lCLId = cboActive.ItemData(cboActive.ListIndex)
             frmPLUGroupDetail.labelClient = cboActive
         
         
         End If
         ' Name and ID of Client
         
         If cboActive.ListIndex > -1 Then
            frmPLUGroupDetail.Show vbModal
            gbOk = ShowPLUGroups(cboActive.ItemData(cboActive.ListIndex))
         End If
         
        Case "Product/PLUs"
         
         If cboActive.ListIndex <> 0 Then

            MsgBox "Please Select View: 'Client PLUs & Stock Products' first"
         
         Else
             'lPLUProductID = lID
             
'             frmPLUProductDetail.lPLUProductID = lPLUProductID
             frmPLUProductDetail.lPLUProductID = lID
             
             If lID = 0 Then
                frmPLUProductDetail.bNewProdPLUInProgress = True
             Else
                frmPLUProductDetail.bNewProdPLUInProgress = False
             End If
             ' send some numbers to the form before displaying it
             ' and deside if its an edit or a new entry
             
             frmPLUProductDetail.Show vbModal
             
             If lID > 0 Then
             ' its an edit so just update the line
        
                 gbOk = UpdateView(lID)
             
             
             ElseIf bAddNewPLUProduct Then
'             ' its a new one so show the lot again and position at the bottom of the list
                 gbOk = ShowProductPLUs(cboActive.ListIndex, cboByGroup.ListIndex, txtSearch)
        
                 grdList.TopRow = grdList.Rows - 1
                 ' make sure last row is visible
            
                 bAddNewPLUProduct = False
             
             End If
        
             bSetFocus Me, "grdList"
        
         End If
    
    End Select
    

End Sub

Public Function UpdateView(lID As Long)
Dim rs As Recordset
Dim sSql As String
Dim lThisRec As Long
Dim sGrp As String
    
    On Error GoTo ErrorHandler
    
    If lID > 0 Then
    
        Set rs = SWdb.OpenRecordset("SELECT tblClientProductPLUs.ID, txtGroupNumber, tblClientProductPLUs.PLUNumber, tblClientProductPLUs.PLUGroupNo, tblPLUs.txtDescription, tblClientProductPLUs.SellPrice, tblProducts.txtCode, tblProductGroup.txtDescription, tblProducts.txtDescription, tblProducts.txtSize, tblClientProductPLUs.PurchasePrice, tblClientProductPLUs.Active, tblProductGroup.txtDescription, tblPLUGroup.txtDescription FROM ((((tblClientProductPLUs INNER JOIN tblClients ON tblClientProductPLUs.ClientID = tblClients.ID) INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE tblClientProductPLUs.ID = " & Trim$(lID), dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            
            If grdList.FindRow(lID) > 0 Then
            
                If grdList.ColIndex("PLU#") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("PLU#")) = rs("PLUNumber")
                If grdList.ColIndex("KeyGroup") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("KeyGroup")) = rs("txtGroupNumber") & "  " & rs("tblPLUGroup.txtDescription")
                If grdList.ColIndex("KeyDescription") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("KeyDescription")) = rs("tblPLUs.txtDescription")
                If grdList.ColIndex("Sell") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Sell")) = Format(rs("SellPrice"), "0.00")
                
                If grdList.ColIndex("Code") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Code")) = rs("txtCode")
                If grdList.ColIndex("Group") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Group")) = rs("tblProductGroup.txtDescription")
                If grdList.ColIndex("Description") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Description")) = rs("tblProducts.txtDescription")
                If grdList.ColIndex("Size") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Size")) = rs("txtSize")
                If grdList.ColIndex("Cost") > -1 Then grdList.Cell(flexcpText, grdList.FindRow(lID), grdList.ColIndex("Cost")) = Format(rs("PurchasePrice"), "0.00")
            
            End If
            
        End If
    
    
    End If
    
    UpdateView = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("UpdateView") Then Resume 0
    Resume CleanExit


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

Public Function SetupCtrlGroup(bProducts As Boolean)

    If bProducts Then
    ' Product Groups

        sMenuCtrl = "Groups"
        
        lblActive.Caption = "View"
        
        SetUpActiveList "Groups"
        ' fill it
    
        cboActive.ListIndex = 0
        ' default to Active
        
    Else
        lblActive.Caption = "By Client"
         
        sMenuCtrl = "PLUClient"
    
        gbOk = GetClientList(cboActive)
        ' get list of clients and put them in cboactive list box
    
        grdList.Rows = 1
    
        If lSelClientID <> 0 Then

            gbOk = PointToEntry(Me, "cboActive", Val(Trim$(lSelClientID)), False)
        End If
    
    End If


End Function


Public Function GetClientList(cboBox As ComboBox)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With cboBox
    
        .Clear
        .AddItem "All"
        
        Set rs = SWdb.OpenRecordset("Select * FROM tblClients WHERE chkActive = true ", dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                .AddItem rs("txtName")
                .ItemData(.NewIndex) = rs("ID") + 0
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    
    End With
    
    GetClientList = True

CleanExit:
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientList") Then Resume 0
    Resume CleanExit


End Function


Public Function SetBackGroundColour(sWhich As String)

    Select Case sWhich
    
        Case "Clients", "Groups"
         Me.Picture = imgList.ListImages("Client").Picture
        
        Case "Products", "PLUs"
         Me.Picture = imgList.ListImages("PLU").Picture
        
        Case "PLUClient"
         
        Case "Product/PLUs"
         Me.Picture = imgList.ListImages(sWhich).Picture
         
        Case Else
        
    End Select

End Function
