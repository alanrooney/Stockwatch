VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmGroupDetail 
   BackColor       =   &H00CDF9FE&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Product Group Detail"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ControlBox      =   0   'False
   Icon            =   "frmGroupDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGroupDetail.frx":1CCA
   ScaleHeight     =   3885
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D6F5F5&
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
      Left            =   5520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   915
   End
   Begin VB.CheckBox chkOpenItem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D6F5F5&
      Caption         =   "Open Item Check"
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
      Left            =   510
      TabIndex        =   4
      Top             =   2100
      Width           =   1935
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
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
      Height          =   390
      Left            =   2250
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Client or Company Name"
      Top             =   1530
      Width           =   4185
   End
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6180
      TabIndex        =   5
      Top             =   3120
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
      TransparentColor=   13498878
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
      TabIndex        =   6
      Top             =   3120
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
      TransparentColor=   13498878
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
      TabIndex        =   9
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
      TransparentColor=   13498878
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
      Height          =   465
      Left            =   2220
      TabIndex        =   8
      Top             =   1500
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Label ID 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   1155
      TabIndex        =   2
      Top             =   1590
      Width           =   1020
   End
   Begin VB.Label lblGroupID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID"
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
      Left            =   1410
      TabIndex        =   0
      Top             =   1050
      Width           =   795
   End
End
Attribute VB_Name = "frmGroupDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bInProgress As Boolean
Public lProductGroupID As Long

Private Sub btnClose_Click()
    
    cmdQuit_Click

End Sub

Private Sub chkOpenItem_Click()

    cmdOk.Enabled = True


End Sub

Private Sub cmdOk_Click()
    If Len(Trim$(txtDescription)) > 2 Then
        
        If lProductGroupID = -1 Then lProductGroupID = 0
        
        cmdOk.Enabled = Not SW1
        ' make sure ok button is going to be enabled if this product
        ' added by a franchisee. The state of it is added in the WriteDB
        ' Yes its a little confusing but it works!
        
        gbOk = WriteDB(Me, "ProductGroup", lProductGroupID, False, 4, _
                chkOpenItem, _
                chkActive, _
                txtDescription, _
                cmdOk)
        
        LogMsg frmStockWatch, "Group Added/Modified " & txtDescription, "ID:" & ID.Caption & " Chk Opn Item:" & Trim$(chkOpenItem) & " Act:" & Trim$(chkActive)
        
        bInProgress = False
        cmdQuit_Click
    
    Else
        MsgBox "Please enter Description"
        bSetFocus Me, "txtDescription"
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
    
    If lProductGroupID > 0 Then
    
        gbOk = ReadDB(Me, "ProductGroup", lProductGroupID, 5, _
            ID, _
            txtDescription, _
            chkOpenItem, _
            chkActive, _
            cmdOk)
    
        cmdOk.Enabled = SetOkState(cmdOk.Enabled)
        ' Set the ok button to true if its either Stock watch head office
        ' or the product was already added by the franchisee
        
        bSetFocus Me, "cmdQuit"
    
    Else
        ID.Caption = GetNextGroupID()
        bSetFocus Me, "txtDescription"
    
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
    gbOk = InitGroupDetail()
    ' init panel
    
    ' list vars that can be set externally

End Sub

Public Function InitGroupDetail()
    
    ID.Caption = ""
    txtDescription.Text = ""
    chkOpenItem.Value = 0
    chkActive.Value = 1

End Function

Private Sub Form_Unload(Cancel As Integer)

    lProductGroupID = 0
    
End Sub

Private Sub txtDescription_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, " /") ' 0 = no only, 1 = char only, 2 = both

    If KeyAscii <> 13 Then cmdOk.Enabled = True


End Sub

Private Sub txtDescription_LostFocus()
    lblDescription.ForeColor = sBlack

End Sub

Public Function GetNextGroupID()
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT ID FROM tblProductGroup Order by ID DESC", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        GetNextGroupID = rs("ID") + 1
    End If
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetNextGroupID") Then Resume 0
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

