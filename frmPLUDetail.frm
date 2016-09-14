VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmPLUDetail 
   BackColor       =   &H00D2DCCF&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch PLUs"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   Icon            =   "frmPLUDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmPLUDetail.frx":1CCA
   ScaleHeight     =   2985
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5610
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   930
      Width           =   915
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
      Height          =   315
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Client or Company Name"
      Top             =   1380
      Width           =   4845
   End
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6210
      TabIndex        =   4
      Top             =   2250
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   873
      BackColor       =   13748165
      Enabled         =   0   'False
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
      TransparentColor=   13819087
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
      Left            =   5220
      TabIndex        =   5
      Top             =   2250
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
      TransparentColor=   13819087
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
      Left            =   6960
      TabIndex        =   6
      Top             =   120
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
      TransparentColor=   13819087
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
      Height          =   375
      Left            =   1650
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   4920
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
      Left            =   570
      TabIndex        =   2
      Top             =   1380
      Width           =   1020
   End
End
Attribute VB_Name = "frmPLUDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bInProgress As Boolean
Public lPLUID As Long
Public bQuickAdd As Boolean

Private Sub btnClose_Click()
    
    cmdQuit_Click

End Sub

Private Sub chkActive_Click()
    cmdOk.Enabled = True

End Sub

Private Sub cmdOk_Click()

    If Len(Trim$(txtDescription)) > 2 Then
        
      If UniquePLU(Trim$(txtDescription)) Then
        
        If lPLUID = -1 Then lPLUID = 0
           
        cmdOk.Enabled = Not SW1
        ' make sure ok button is going to be enabled if this product
        ' added by a franchisee. The state of it is added in the WriteDB
        ' Yes its a little confusing but it works!
        
        gbOk = WriteDB(Me, "PLUs", lPLUID, False, 3, _
                    chkActive, _
                    txtDescription, _
                    cmdOk)
        
        LogMsg frmStockWatch, "PLU Added/Modified " & txtDescription, " Act:" & Trim$(chkActive)
        
        bInProgress = False
        cmdQuit_Click
    
      Else
        MsgBox "Another PLU of the same name exists!"
        bSetFocus Me, "txtDescription"
      End If
      
    Else
        MsgBox "Please enter Description"
        bSetFocus Me, "txtDescription"
    End If

End Sub

Private Sub cmdQuit_Click()
    ' check if entry in progress and warn
    If bInProgress Then
        If MsgBox("Quit Entering/Modifying PLU Details", vbQuestion + vbYesNo, "Cancel Enter Customer") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Activate()
    
    If lPLUID > 0 Then
        
        gbOk = ReadDB(Me, "PLUs", lPLUID, 3, _
                txtDescription, _
                chkActive, _
                cmdOk)
        
  '      cmdOk.Enabled = SetOkState(cmdOk.Enabled)
        ' Set the ok button to true if its either Stock watch head office
        ' or the product was already added by the franchisee
        
        bSetFocus Me, "cmdQuit"
        
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    KeyAscii = CharOk(KeyAscii, 2, " '*-&%./") ' 0 = no only, 1 = char only, 2 = both
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        cmdQuit_Click
        
    End If

End Sub

Private Sub Form_Load()
    gbOk = InitPLUDetail()
    ' init panel
    
    ' list vars that can be set externally


End Sub
Public Function InitPLUDetail()
    
    txtDescription.Text = ""
    chkActive.Value = 1

End Function

Private Sub txtDescription_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        bSetFocus Me, "cmdOk"
    Else
        cmdOk.Enabled = True
    End If
    
End Sub

Private Sub txtDescription_LostFocus()
    lblDescription.ForeColor = sBlack

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


Public Function UniquePLU(sPLU As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("Select ID from tblPLUS WHERE txtDescription = '" & sPLU & "'", dbOpenSnapshot)
    If (rs.EOF And rs.BOF) Then
        UniquePLU = True
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    
    If CheckDBError("UniquePLU") Then Resume 0
    Resume CleanExit

End Function
