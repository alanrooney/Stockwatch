VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00A0724B&
   BorderStyle     =   0  'None
   ClientHeight    =   5970
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":1CCA
   ScaleHeight     =   5970
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmailAddress 
      BackColor       =   &H00EFE7E0&
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
      Height          =   240
      Left            =   705
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4110
      Width           =   3180
   End
   Begin VB.TextBox txtRegion 
      BackColor       =   &H00EFE7E0&
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
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   3525
      Width           =   570
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00EFE7E0&
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
      Height          =   240
      Left            =   705
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3795
      Width           =   3180
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00EFE7E0&
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
      Height          =   870
      Left            =   705
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2865
      Width           =   3165
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00EFE7E0&
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
      Height          =   345
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2475
      Width           =   3135
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   7830
      TabIndex        =   4
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
      TransparentColor=   10515019
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
   Begin VB.Label labelExpDate 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3270
      TabIndex        =   7
      Top             =   4710
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   5760
      Width           =   2715
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Rev: 5.0.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5895
      TabIndex        =   0
      Top             =   1350
      Width           =   1875
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFormIsShown As Boolean
Public bShowSplash As Boolean

'''Private Sub btnCheckForUpdate_Click()
'''
'''    Screen.MousePointer = 11
'''
'''    frmAutoUpdate.Show vbModal
'''
'''    Screen.MousePointer = 0
'''
'''
'''End Sub

Private Sub btnClose_Click()

    Unload Me

End Sub


Private Sub btnSaveText_Click()
End Sub


Private Sub Form_Activate()
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sCount As String
'Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer
    
    NoLicense = True    ' Default
    
    If GetLicenseInfo(sName, sAddress, sPhone, sFranchiseEmail, dtExpiry, iDays, iwarn) Then
        NoLicense = False
    End If
    
    
    txtName = sName
    txtAddress = sAddress
    txtPhone = sPhone
    txtRegion = gbRegion
    txtEmailAddress = sFranchiseEmail
    If (gbRegion <> "SW1") And (gbRegion <> "SW2") Then
    ' If its StockWatch Head Office then dont check/extend expiry date
    
        labelExpDate = Format(dtExpiry)
        
        If DateValue(dtExpiry) <= DateValue(Format(Now, sDMY)) Then
            
            gbOk = ExtendExpiryDate(False)
            LogMsg Me, "", "Expiry Date of License: " & Trim$(dtExpiry) & " - License has expired"
            MsgBox "License has expired. Please contact StockWatch Ireland on 091 442987.", vbOKOnly, "License Expired"
            End
        
        ElseIf DateValue(dtExpiry - iwarn) <= DateValue(Now) Then
            LogMsg Me, "", "Expiry Date of License: " & Trim$(dtExpiry) & " - Warning that license will expire"
            MsgBox "License will expiry in " & Trim$(DateDiff("d", Now, dtExpiry)) & " day(s)", vbOKOnly, "License Expiring"
        End If
    
    
    End If
    
    bSetFocus Me, "btnClose"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
    
'    labelDebug.Visible = bDebug
    
'    lblMaint.BackStyle = 0  ' set maint label transparent
    
    lblVersion.Caption = "Current Rev: " & App.Major & "." & App.Minor & "." & App.Revision

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

'Private Sub lblMaint_Click()
'    ' This is used for doing any little updates required
'
'    ' An extra dot after the revision indicates the procedure is finished running
'
'    '====================================================================================
'    ' Ver 551
'    ' Set the 'Include in History' check box same as active flag
'    SWdb.Execute "UPDATE tblClientProductPLUs SET chkHistory = true WHERE Active = true"
'    '====================================================================================
'
'    lblVersion.Caption = lblVersion.Caption & "."
'
'
'End Sub


'''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'''
'''    If KeyCode = Asc("D") Then
'''        If Shift = 2 Then
'''
'''            bDebug = Not bDebug
'''            ' toggle
'''
'''            labelDebug.Visible = bDebug
'''
'''        End If
'''    End If
'''
'''End Sub

Private Sub txtAddress_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.BackColor = &HEFE7E0

End Sub


Private Sub txtName_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = &HEFE7E0
End Sub

Private Sub txtPhone_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtPhone_LostFocus()
    txtPhone.BackColor = &HEFE7E0

End Sub

