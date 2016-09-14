VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmProductDetail 
   BackColor       =   &H00C6DDD6&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Product Detail"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   Icon            =   "frmProductDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProductDetail.frx":1CCA
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEmptyVerify 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00A4B88D&
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   19
      Top             =   3780
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CheckBox chkFullVerify 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B6C6A4&
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   18
      Top             =   3150
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox cboVatTable 
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
      Left            =   4770
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F2F3ED&
      Caption         =   "Acti&ve"
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
      Left            =   5790
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   870
      Width           =   915
   End
   Begin VB.ComboBox cboGroups 
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
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1410
      Width           =   2025
   End
   Begin VB.TextBox txtIssueUnits 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   13
      ToolTipText     =   "Client or Company Name"
      Top             =   4350
      Width           =   945
   End
   Begin VB.TextBox txtEmptyWeight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   11
      ToolTipText     =   "Client or Company Name"
      Top             =   3750
      Width           =   945
   End
   Begin VB.TextBox txtFullWeight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   9
      ToolTipText     =   "Client or Company Name"
      Top             =   3150
      Width           =   945
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   7
      ToolTipText     =   "Client or Company Name"
      Top             =   2550
      Width           =   945
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   5
      ToolTipText     =   "Client or Company Name"
      Top             =   1950
      Width           =   4665
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
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
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   1
      ToolTipText     =   "Client or Company Name"
      Top             =   1410
      Width           =   945
   End
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6195
      TabIndex        =   20
      Top             =   5130
      Width           =   840
      _ExtentX        =   1482
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
      TransparentColor=   13032918
      Caption         =   "&Ok"
      CaptionPosition =   4
      ForeColorDisabled=   -2147483632
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton cmdQuit 
      Height          =   495
      Left            =   5190
      TabIndex        =   21
      Top             =   5130
      Width           =   840
      _ExtentX        =   1482
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
      TransparentColor=   13032918
      Caption         =   "&Quit"
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
      Left            =   6930
      TabIndex        =   23
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
      TransparentColor=   13032918
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
      Left            =   2070
      TabIndex        =   22
      Top             =   1380
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label labelVatRate 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5850
      TabIndex        =   17
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblGroups 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Group"
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
      Height          =   240
      Left            =   4140
      TabIndex        =   2
      Top             =   1470
      Width           =   555
   End
   Begin VB.Label lblVatTable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vat &Table"
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
      Height          =   240
      Left            =   3840
      TabIndex        =   14
      Top             =   2580
      Width           =   885
   End
   Begin VB.Label lblIssueUnits 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Issue Units"
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
      Height          =   240
      Left            =   1050
      TabIndex        =   12
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label lblEmptyWeight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Empty Weight"
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
      Height          =   240
      Left            =   810
      TabIndex        =   10
      Top             =   3780
      Width           =   1245
   End
   Begin VB.Label lblFullWeight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Full Weight"
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
      Height          =   240
      Left            =   1050
      TabIndex        =   8
      Top             =   3195
      Width           =   990
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Size"
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
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   2580
      Width           =   390
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Description"
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
      Height          =   240
      Left            =   1035
      TabIndex        =   4
      Top             =   1995
      Width           =   1020
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Code"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   1455
      Width           =   495
   End
End
Attribute VB_Name = "frmProductDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public bInProgress As Boolean
Public lProductID As Long
Public bQuickAdd As Boolean
Public bFranchiseOwned As Boolean

Private Sub btnClose_Click()
    cmdQuit_Click

End Sub

Private Sub cboGroups_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboGroups_KeyPress(KeyAscii As Integer)

    cmdOk.Enabled = True

End Sub

Private Sub cboGroups_LostFocus()
    lblGroups.ForeColor = sBlack

End Sub

Private Sub cboVatTable_Click()

    labelVatRate.Caption = GetVatRate(cboVatTable)

End Sub

Private Sub cboVatTable_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboVatTable_LostFocus()
    lblVatTable.ForeColor = sBlack

End Sub

Private Sub chkEmptyVerify_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then bSetFocus Me, "cmdOk"


End Sub

Private Sub cmdOk_Click()
    
    If FieldsCheckOut() Then
        
        
        
        gbOk = SaveProduct(lProductID)
        
        LogMsg frmStockWatch, "Product Added/Modified " & txtDescription, "Code:" & txtCode & " Group:" & cboGroups & " Size:" & txtSize & " Full:" & txtFullWeight & " Empty:" & txtEmptyWeight & " Issue Units:" & txtIssueUnits
        
        bInProgress = False
    
        If Not bQuickAdd Then
        ' as long as we're not doing a quick add then re-show this list
            If frmCtrl.txtSearch.Text <> "" Then gbOk = frmCtrl.ShowProducts(frmCtrl.cboActive.ListIndex, frmCtrl.txtSearch)
        ' Ver440 changed the click to txtsearch_click
        
        End If
        
        cmdQuit_Click
    
    End If


End Sub

Private Sub cmdQuit_Click()
    ' check if entry in progress and warn
    If bInProgress Then
        If MsgBox("Quit Entering/Modifying Product Details", vbQuestion + vbYesNo, "Cancel Enter Customer") = vbYes Then
            Unload Me
            bFranchiseOwned = False
        End If
    Else
        Unload Me
        bFranchiseOwned = False
    End If


End Sub

Private Sub Form_Load()
    
    gbOk = InitProductDetail()
    ' init panel
    

End Sub

Public Function InitProductDetail()

    txtCode.Text = ""
    txtDescription.Text = ""
    txtSize.Text = ""
    txtFullWeight.Text = ""
    txtEmptyWeight.Text = ""
    txtIssueUnits.Text = ""
    gbOk = GetVatRates()
    gbOk = PointToEntry(Me, "cboVatTable", "S", True)
    
    gbOk = GetGroups(Me)
    ' true = Product groups
    
    cboGroups.ListIndex = -1
    chkActive.Value = 1
    

End Function
Private Sub Form_Activate()

    If lProductID > 0 Then
        
        gbOk = ReadDB(Me, "Products", lProductID, 12, _
                    chkActive, _
                    txtCode, _
                    txtDescription, _
                    txtSize, _
                    txtFullWeight, _
                    txtEmptyWeight, _
                    txtIssueUnits, _
                    chkFullVerify, _
                    chkEmptyVerify, _
                    cboVatTable, _
                    cboGroups, _
                    cmdOk)

        bFranchiseOwned = cmdOk.Enabled
        ' Save it here
        
        cmdOk.Enabled = SetOkState(bFranchiseOwned)
        ' Set the ok button to true if its either Stock watch head office
        ' or the product was already added by the franchisee
        
'        SetFeildsLocked Not cmdOk.Enabled
        
        bSetFocus Me, "cmdQuit"
    
    Else
        txtCode = GetNextCode()
        cmdOk.Enabled = True
        bSetFocus Me, "cboGroups"
        
    
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        cmdQuit_Click
        
    End If

End Sub

Public Function FieldsCheckOut()
    
    If Val(txtCode) > 999 And Val(txtCode) < 1000000 Then
        If Not CodeInUse(Val(txtCode), lProductID) Then
            If cboGroups.ListIndex > -1 Then
                If Len(txtDescription) > 2 Then
                    If cboVatTable.ListIndex > -1 Then
                        If Len(txtSize.Text) > 1 Then
                            If Val(txtFullWeight.Text) > 0 Then
                                If Val(txtEmptyWeight.Text) > 0 Then
                                    If Val(txtIssueUnits.Text) > 0 Then
                            
                                        FieldsCheckOut = True
                
                                    Else
                                        MsgBox "Please set the Issue Units Value"
                                        bSetFocus Me, "txtIssueUnits"
                                    End If
                                Else
                                    MsgBox "Please enter the Empty Weight Value"
                                    bSetFocus Me, "txtEmptyWeight"
                                End If
                            Else
                                MsgBox "Please enter the Full Weight Value"
                                bSetFocus Me, "txtFullWeight"
                            End If
                        Else
                            MsgBox "Please enter a size"
                            bSetFocus Me, "txtSize"
                        End If
                    Else
                        MsgBox "Please Select a valid Vat Table"
                        bSetFocus Me, "cboVatTable"
                    End If
                    
                Else
                    MsgBox "Please enter a valid description. (Must be at least 3 Characters in length)"
                    bSetFocus Me, "txtDescription"
                End If
            Else
                MsgBox "Please Select a Group for the product"
                bSetFocus Me, "cboGroups"
            End If
        Else
            MsgBox "Code already in use"
            bSetFocus Me, "txtCode"
        End If
    Else
        MsgBox "Please enter a Valid Code Number (1000 - 999999)"
        bSetFocus Me, "txtCode"
    End If
    
End Function

Public Function GetVatRates()
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    cboVatTable.Clear
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblVat", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        
        rs.MoveFirst
        
        Do
            cboVatTable.AddItem rs("txtCode") & ""
        
            rs.MoveNext
        Loop While Not rs.EOF
        
    End If
    rs.Close
    GetVatRates = True
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetVatRates ") Then Resume 0
    Resume CleanExit

End Function


Public Function GetNextCode()
Dim rs As Recordset
Dim iMax As Integer
    
    On Error GoTo ErrorHandler
    
    ' HAd to modify this rountine to get max number stored in string: txtcode
    ' Someone had added 999 and the old routine kept showing 1000 as the next number!
    
    ' its a bit ineffecent but it works - saved having to change the field in the db to an integer type.
    
    Set rs = SWdb.OpenRecordset("SELECT ID, txtCode FROM tblProducts ORDER BY ID DESC", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        
        rs.MoveFirst
        Do
        
            If Val(rs("txtCode")) > iMax Then
                iMax = Val(rs("txtCode"))
            End If
            
            rs.MoveNext
        Loop While Not rs.EOF
        
        GetNextCode = iMax + 1
        
    End If
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetNextCode ") Then Resume 0
    Resume CleanExit


End Function

Public Function SaveProduct(lID As Long)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblProducts")
    rs.Index = "PrimaryKey"
    If lID <> 0 Then
        rs.Seek "=", lID
        If Not rs.NoMatch Then
            rs.Edit
        Else
            rs.AddNew
        End If
    Else
        rs.AddNew
    End If
    
    rs("txtCode") = Val(txtCode.Text)
    rs("txtDescription") = Trim$(txtDescription.Text)
    rs("txtSize") = Trim$(txtSize.Text)
    rs("cboGroups") = cboGroups.ItemData(cboGroups.ListIndex)
    rs("txtIssueUnits") = Val(txtIssueUnits.Text)
    rs("txtFullWeight") = Val(txtFullWeight.Text)
    rs("txtEmptyWeight") = Val(txtEmptyWeight.Text)
    rs("cboVatTable") = cboVatTable.Text
    rs("chkFullVerify") = chkFullVerify
    rs("chkEmptyVerify") = chkEmptyVerify
    rs("chkActive") = chkActive
    
    If lID = 0 Then rs("cmdOk") = Not SW1
    ' make sure ok button is going to be enabled if this product
    ' added by a franchisee
    
    rs.Update
    
    rs.Bookmark = rs.LastModified
    lID = rs("ID") + 0
    
    SaveProduct = True
    
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    
    If Err = 3022 Then
        rs("id") = rs("id") + 1
        Resume 0
    
    Else

       If CheckDBError("SaveProduct ") Then Resume 0
        Resume CleanExit
    End If
    


End Function

Private Sub txtCode_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

    cmdOk.Enabled = True


    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtCode_LostFocus()
    lblCode.ForeColor = sBlack
End Sub

Private Sub txtDescription_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    
'    txtDescription.Locked = False
    cmdOk.Enabled = True
    
    KeyAscii = CharOk(KeyAscii, 2, " *-&%./'") ' 0 = no only, 1 = char only, 2 = both

    ' Ver 310
    ' Allow name change only



End Sub

Private Sub txtDescription_LostFocus()
    lblDescription.ForeColor = sBlack

End Sub

Private Sub txtEmptyWeight_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtEmptyWeight_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
    
        cmdOk.Enabled = True
        
    ElseIf InStr(txtEmptyWeight, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtEmptyWeight_LostFocus()
    lblEmptyWeight.ForeColor = sBlack

End Sub

Private Sub txtFullWeight_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtFullWeight_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
        cmdOk.Enabled = True
    
    ElseIf InStr(txtFullWeight, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtFullWeight_LostFocus()
    lblFullWeight.ForeColor = sBlack

End Sub

Private Sub txtIssueUnits_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtIssueUnits_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then
        
        cmdOk.Enabled = True
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
        If KeyAscii = 13 Then bSetFocus Me, "cmdOk"
    
    ElseIf InStr(txtIssueUnits, ".") > 0 Then
            
        KeyAscii = 0
    
    End If

End Sub

Public Function CodeInUse(iCode As Integer, lProdID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblProducts")
    rs.Index = "txtcode"
    rs.Seek "=", iCode
    If rs.NoMatch Then
        ' not in use so ok ...
        CodeInUse = False
        
    ElseIf rs("ID") <> lProdID Then
        
        CodeInUse = True
    
    Else
        CodeInUse = False
    
    End If
    
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("CodeInUse") Then Resume 0
    Resume CleanExit


End Function

Private Sub txtIssueUnits_LostFocus()
    lblIssueUnits.ForeColor = sBlack

End Sub

Private Sub txtSize_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    cmdOk.Enabled = True

End Sub

Private Sub txtSize_LostFocus()
    lblSize.ForeColor = sBlack

End Sub
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

'Public Function SetFeildsLocked(bhow As Boolean)
'
'' only enable buttons other than the name box
'' so that name can only be edited
'
'    chkActive.Enabled = Not bhow
'
'    cboGroups.Locked = bhow
'    cboVatTable.Locked = bhow
'    txtCode.Locked = bhow
'    txtDescription.Locked = bhow
'    txtSize.Locked = bhow
'    txtFullWeight.Locked = bhow
'    txtEmptyWeight.Locked = bhow
'    txtIssueUnits.Locked = bhow
'
'End Function
