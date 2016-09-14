VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmPLUGroupDetail 
   BackColor       =   &H00C6DDD6&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch PLU Group Detail"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmPLUGroupDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPLUGroupDetail.frx":1CCA
   ScaleHeight     =   4200
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGlass 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmPLUGroupDetail.frx":7CC0
      Left            =   4020
      List            =   "frmPLUGroupDetail.frx":7CC2
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox cboDescription 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmPLUGroupDetail.frx":7CC4
      Left            =   2040
      List            =   "frmPLUGroupDetail.frx":7CC6
      TabIndex        =   7
      Text            =   "cboDescription"
      Top             =   2520
      Width           =   4305
   End
   Begin VB.TextBox txtGroupNumber 
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
      Left            =   2070
      MaxLength       =   100
      TabIndex        =   3
      ToolTipText     =   "Client or Company Name"
      Top             =   1935
      Width           =   615
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEF1E2&
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
      Left            =   6330
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   900
      Width           =   915
   End
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6510
      TabIndex        =   9
      Top             =   3420
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
      Left            =   5520
      TabIndex        =   10
      Top             =   3420
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
      Left            =   7230
      TabIndex        =   11
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
   Begin VB.Label lblGlass 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Measure"
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
      Left            =   3180
      TabIndex        =   4
      Top             =   1965
      Width           =   795
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
      Left            =   2040
      TabIndex        =   12
      Top             =   1905
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label labelClient 
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
      Height          =   375
      Left            =   2025
      TabIndex        =   1
      Top             =   1290
      Width           =   4275
   End
   Begin VB.Label lblClient 
      BackStyle       =   0  'Transparent
      Caption         =   "&Client"
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
      Left            =   1485
      TabIndex        =   0
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label lblGroupNumber 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Group Number"
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
      Left            =   660
      TabIndex        =   2
      Top             =   1980
      Width           =   1320
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
      Left            =   975
      TabIndex        =   6
      Top             =   2580
      Width           =   1020
   End
End
Attribute VB_Name = "frmPLUGroupDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lPLUGroupID As Long
Public lCLId As Long
Public bInProgress As Boolean



Private Sub btnClose_Click()

    cmdQuit_Click
    

End Sub

Private Sub cboDescription_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboDescription_KeyPress(KeyAscii As Integer)

    cboDescription.ListIndex = -1

    If KeyAscii <> 13 Then cmdOk.Enabled = True


End Sub

Private Sub cboDescription_LostFocus()
    lblDescription.ForeColor = sBlack
    
End Sub

Private Sub cboGlass_Click()
    cmdOk.Enabled = True
End Sub

Private Sub cboGlass_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboGlass_LostFocus()
    lblGlass.ForeColor = sBlack
End Sub

Private Sub cmdOk_Click()
            
    
    If CheckPLUGroupFieldsOK() Then
    
        gbOk = SaveClientPLUGroup(lPLUGroupID)
    
        LogMsg frmStockWatch, "PLU/Group Added/Modified for " & labelClient, "grp No:" & txtGroupNumber & " Desc:" & cboDescription & " Act:" & Trim$(chkActive)
    
        bInProgress = False
        Unload Me
    
    End If

End Sub

Private Sub cmdQuit_Click()
    ' check if entry in progress and warn
    If bInProgress Then
        If MsgBox("Quit Entering/Modifying Group Details", vbQuestion + vbYesNo, "Cancel Enter Customer") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If



End Sub

Private Sub Form_Activate()
    
    ' ver530 glass included here.
    
    If lPLUGroupID > 0 Then
        
        gbOk = GetClientPLUGroup(lPLUGroupID)
        bSetFocus Me, "cmdQuit"

    Else
            
        bSetFocus Me, "txtGroupNumber"
    
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

Private Sub Form_Load()
    
    InitPLUGroup
    
    gbOk = GetPLUGroupList()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lPLUGroupID = 0

End Sub

Private Sub txtgroupNumber_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtGroupNumber_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then cmdOk.Enabled = True

End Sub

Private Sub txtgroupNumber_LostFocus()
    lblClient.ForeColor = sBlack

End Sub

Public Function GetPLUGroupList()
Dim rs As Recordset
Dim iCnt As Integer

    On Error GoTo ErrorHandler
    
    cboDescription.Clear
    
    Set rs = SWdb.OpenRecordset("SELECT DISTINCT txtDescription from tblPLUGroup ORDER BY txtDescription", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            For iCnt = 0 To cboDescription.ListCount - 1


' Ver 530
' While making changes here for ver530 (glass measures) I noticed the line below and
' where this line might have worked for having only one group of Draught Beers or Draught Beer to
' choose from it was bad for spirits group.
' The group 1/2 Bot spirits appeared on the list but the group Spirits itself was missing!!!!
'
' And the reason Whiskey and White Spirits was removed from list is that Spirits is the group to use.


'                If InStr(cboDescription.List(iCnt), rs("txtDescription")) <> 0 Then
                If cboDescription.List(iCnt) = rs("txtDescription") Then
                    ' found it...
                    GoTo GetNext
                
                ElseIf rs("txtDescription") = "Whiskeys" Or rs("txtDescription") = "White Spirits" Then
                ' Ver 205
                    GoTo GetNext
                
                End If
            Next
    
            cboDescription.AddItem rs("txtDescription")
            
GetNext:
        rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    
    GetPLUGroupList = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUGroupList") Then Resume 0
    Resume CleanExit



End Function


Public Function GetClientPLUGroup(lID As Long)
Dim rs As Recordset


    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblPLUGroup")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
    
        txtGroupNumber = rs("txtGroupNumber")
        
                
        
        txtGroupNumber.Tag = rs("txtGroupNumber") ' save it here for checking later
        
        gbOk = PointToEntry(Me, "cboDescription", rs("txtDescription"), True)
        
        If Not IsNull(rs("Glass")) Then
            gbOk = PointToEntry(Me, "cboGlass", rs("Glass"), False)
        End If
        
        chkActive = Abs(rs("chkActive"))
    
        cmdOk.Enabled = SetOkState(rs("CmdOk"))
        ' Set the ok button to true if its either Stock watch head office
        ' or the product was already added by the franchisee
        
        GetClientPLUGroup = True
    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientPLUGroup") Then Resume 0
    Resume CleanExit


End Function

Public Function SaveClientPLUGroup(lID As Long)
Dim rs As Recordset


    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblPLUGroup")
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
    
    rs("ClientID") = lCLId
    rs("txtGroupNumber") = txtGroupNumber
    rs("txtDescription") = cboDescription
    If cboGlass.ListIndex > -1 Then
        rs("Glass") = cboGlass.ItemData(cboGlass.ListIndex)
    End If
    rs("chkActive") = chkActive
    rs("CmdOk") = Not SW1
    rs.Update
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If Err = 3022 Then
        rs("id") = rs("id") + 1
        Resume 0
    
    Else
    
        If CheckDBError("SaveClientPLUGroup") Then Resume 0
        Resume CleanExit
    End If


End Function

Public Sub InitPLUGroup()


    labelClient.Caption = ""
    txtGroupNumber.Text = ""
    txtGroupNumber.Tag = ""
    cboDescription.Text = ""
    cboDescription.ListIndex = -1
    
    
    cboGlass.Clear
    cboGlass.AddItem "0 - Not Applicable"
    cboGlass.ItemData(cboGlass.NewIndex) = 0
    cboGlass.AddItem "2 - Draught Beer"
    cboGlass.ItemData(cboGlass.NewIndex) = 2
    cboGlass.AddItem "4 - Open Wine"
    cboGlass.ItemData(cboGlass.NewIndex) = 4
    cboGlass.AddItem "5 - Champagne / Sparkling"
    cboGlass.ItemData(cboGlass.NewIndex) = 5
    
    
    cboGlass.ListIndex = 0    ' default
    chkActive.Value = 1
    
End Sub

Public Function CheckPLUGroupFieldsOK()

    If lCLId <> 0 Then
        If GroupNoOK() Then

            If DescriptionOK() Then
            
                If cboGlass.ListIndex > -1 Then
                
                    If chkActive Then
                    
                        CheckPLUGroupFieldsOK = True
                
                    ElseIf MsgBox("Are you sure you want to disable this PLU group?", vbDefaultButton1 + vbYesNo + vbQuestion, "Disable Group") = vbYes Then
                            CheckPLUGroupFieldsOK = True
                        
                    End If
                Else
                    MsgBox "Please Select a measure quantity"
                    bSetFocus Me, "cboGlass"
                End If
                
            Else
                MsgBox "Please Select a description for the PLU Group"
                bSetFocus Me, "cboDescription"
            End If
            
        End If
        
    Else
        MsgBox "No Client Specified"
    End If
    

End Function

Public Function GroupNoOK()

    If lPLUGroupID = 0 Then
    ' new group

        If ActiveClientgroupexists(txtGroupNumber) Then
            MsgBox "Group No" & txtGroupNumber & " is already an active group number for this client"
        Else
            GroupNoOK = True
        End If
    
    Else
    ' modify a group
    
        If txtGroupNumber.Text = txtGroupNumber.Tag Then
            GroupNoOK = True
        
        ElseIf ActiveClientgroupexists(txtGroupNumber) Then
            MsgBox "Group No" & txtGroupNumber & " is already an active group number for this client"
        Else
            GroupNoOK = True
        
        
        End If
        

    End If

End Function

Public Function ActiveClientgroupexists(sGrpNo As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT chkActive FROM tblPLUGroup WHERE ClientID = " & lSelClientID & " AND txtGroupNumber = " & Val(sGrpNo), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        ActiveClientgroupexists = rs("chkActive")
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ActiveClientgroupexists") Then Resume 0
    Resume CleanExit

End Function

Public Function DescriptionOK()

    If cboDescription.ListIndex > -1 Then
        If cboDescription = cboDescription.List(cboDescription.ListIndex) Then
            DescriptionOK = True
        Else
            If MsgBox("Do you wish to create a new group description for Group Number " & txtGroupNumber, vbDefaultButton1 + vbYesNo + vbQuestion, "New Group Description") = vbYes Then
                DescriptionOK = True
            End If
        End If
    
    ElseIf cboDescription <> "" Then
        If MsgBox("Do you wish to create a new group description for Group Number " & txtGroupNumber, vbDefaultButton1 + vbYesNo + vbQuestion, "New Group Description") = vbYes Then
            DescriptionOK = True
        End If
    
    Else
        MsgBox ("Please Select a Group Description from the List")
        bSetFocus Me, "cboDescription"
    End If

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

