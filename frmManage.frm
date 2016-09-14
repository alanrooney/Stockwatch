VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmManage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmManage.frx":0000
   ScaleHeight     =   4215
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView Cal 
      Height          =   2820
      Left            =   1065
      TabIndex        =   6
      Top             =   1170
      Visible         =   0   'False
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   5606001
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   16777218
      TitleBackColor  =   13342061
      TitleForeColor  =   16777215
      CurrentDate     =   39972
   End
   Begin VB.OptionButton optSetExpiryDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4EFE1&
      Caption         =   "Set Expiry Date"
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
      Left            =   795
      TabIndex        =   3
      Top             =   2325
      Width           =   1710
   End
   Begin VB.OptionButton optExpireLicense 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4EFE1&
      Caption         =   "Terminate License"
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
      Left            =   495
      TabIndex        =   2
      Top             =   1830
      Width           =   1995
   End
   Begin VB.ComboBox cboRegions 
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
      Left            =   2325
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1215
      Width           =   1425
   End
   Begin MSMask.MaskEdBox tedDate 
      Height          =   375
      Left            =   2790
      TabIndex        =   4
      Top             =   2310
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MyCommandButton.MyButton btnCall 
      Height          =   360
      Left            =   3780
      TabIndex        =   5
      Top             =   2325
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmManage.frx":4B7A
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmManage.frx":7F7F
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   4740
      TabIndex        =   9
      Top             =   45
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
   Begin MyCommandButton.MyButton btnActivate 
      Height          =   540
      Left            =   3735
      TabIndex        =   7
      Top             =   3465
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   953
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
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Execute"
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
   Begin MyCommandButton.MyButton btnQuit 
      Height          =   540
      Left            =   2655
      TabIndex        =   8
      Top             =   3465
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   953
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
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Stockwatch - Manage Franchise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   10
      Top             =   45
      Width           =   3600
   End
   Begin VB.Label lblRegions 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Region"
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
      Left            =   960
      TabIndex        =   0
      Top             =   1245
      Width           =   1485
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnActivate_Click()
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sRegion As String
Dim sCount As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iWarn As Integer

    If FieldsOK() Then
            
        If MsgBox("Are you sure you want to perform this action", vbYesNo + vbDefaultButton2 + vbQuestion, "License Termination/Timeout") = vbYes Then
        ' Confirm Action
            
            If DoExpiry() Then
            
                MsgBox "New Expiry Date sent to " & cboRegions
            End If
            
        End If
    End If
    
    Unload Me

End Sub

Private Sub btnCall_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedDate) Then
            Cal.Value = tedDate
        Else
            Cal.Value = Format(Now, "dd/mm/yy")
        End If
        Cal.Top = cboRegions.Top
        Cal.Left = btnCall.Left - Cal.Width
        Cal.Visible = True
    End If

End Sub

Private Sub btnClose_Click()

    btnQuit_Click
    

End Sub


Private Sub btnQuit_Click()

    Unload Me

End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)
    
    tedDate.Text = Format(Cal.Value, "dd/mm/yy")
    Cal.Visible = False

    setActivateBtn


End Sub

Private Sub cboRegions_Click()
    setActivateBtn

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    ' GET REGIONS
    
    gbOk = GetRegions(Me)

    



End Sub

Private Sub optExpireLicense_Click()

    setActivateBtn

End Sub

Public Sub setActivateBtn()

    If cboRegions.ListIndex <> -1 Then
    
        If optExpireLicense.Value Then
            btnActivate.Enabled = True
        ElseIf optSetExpiryDate.Value Then
            If IsDate(tedDate) Then
                If DateValue(tedDate) > DateValue(Now) Then
                    btnActivate.Enabled = True
                Else
                    btnActivate.Enabled = False
                End If
            Else
                btnActivate.Enabled = False
            End If
        Else
            btnActivate.Enabled = False
        End If
    End If
    
End Sub

Private Sub optSetExpiryDate_Click()
    setActivateBtn

End Sub

Public Function FieldsOK()

    If cboRegions.ListIndex <> -1 Then
    
        If optExpireLicense Or optSetExpiryDate Then
        
            If optExpireLicense Then
                FieldsOK = True
                
            ElseIf optSetExpiryDate Then
            
                If IsDate(tedDate) Then
                    FieldsOK = True
                
                End If
            End If
        
        Else
            MsgBox "No Action Selected"
            bSetFocus Me, "optSetExpiryDate"
        End If
        
    Else
        MsgBox "No Region Selected"
        bSetFocus Me, "cboRegions"
    End If
    
End Function
'''Public Function GetLicenseInfo(sName As String, _
'''                                sAddress As String, _
'''                                sPhone As String, _
'''                                sRegion As String, _
'''                                dtExpiry As Date, _
'''                                iDays As Integer, _
'''                                iWarn As Integer)
'''
'''Dim rs As Recordset
'''Dim sExpiry As String
'''
'''    On Error GoTo ErrorHandler
'''
'''    Set rs = swDB.OpenRecordset("tblFranchisee")
'''    rs.Index = "PrimaryKey"
'''
'''    If Not rs.EOF Then
'''
'''        sName = rs("name") & ""
'''        sAddress = rs("address") & ""
'''        sPhone = rs("phone") & ""
'''        sRegion = rs("region") & ""
'''        sExpiry = rs("Expiry") & ""
'''
'''        iDays = 0
'''        iWarn = 0
'''        If sExpiry <> "" Then
'''            sExpiry = Decrypt(sExpiry, sKey)
'''
'''            If InStr(sExpiry, "~") Then
'''                If IsDate(Left(sExpiry, InStr(sExpiry, "~") - 1)) Then
'''                    dtExpiry = Left(sExpiry, InStr(sExpiry, "~") - 1)
'''
'''                    If InStr(sExpiry, "@") Then
'''                        iDays = Val(Mid(sExpiry, InStr(sExpiry, "~") + 1, InStr(sExpiry, "@") - InStr(sExpiry, "~") - 1))
'''
'''                        If InStr(sExpiry, "|") Then
'''                            iWarn = Mid(sExpiry, InStr(sExpiry, "@") + 1, InStr(sExpiry, "|") - InStr(sExpiry, "@") - 1)
'''                        End If
'''                    End If
'''
'''                End If
'''
'''            Else
'''                dtExpiry = Format(Now - 1, "dd/mm/yy")
'''            End If
'''        Else
'''            dtExpiry = Format(Now - 1, "dd/mm/yy")
'''        End If
'''
'''        sFranchiseEmail = rs("Email") & ""
'''        sSWEmail = rs("SWEmail") & ""
'''
'''    End If
'''
'''    rs.Close
'''
'''End Function


Public Function DoExpiry()
Dim sDropBoxLoc As String
Dim filenum As Long
Dim dtExpiry As Date
Dim sExpiry As String

    On Error GoTo ErrorHandler

    sDropBoxLoc = "C:\Documents and Settings\" & VBA.Environ("USERNAME") & "\My Documents\My Dropbox\StockWatch\Franchise" & "\" & cboRegions
    ' get dropbox folder

    filenum = FreeFile
    Open sDBLoc & "\" & cboRegions & "_Maint.log" For Output As #filenum

    If optExpireLicense Then
        dtExpiry = "01/01/1970"
    
    ElseIf optSetExpiryDate Then
        dtExpiry = tedDate
    End If
    
    sExpiry = dtExpiry & "~" & "30" & "@" & "7" & "|"
    sExpiry = Encrypt(sExpiry, sKey)

    Print #filenum, sExpiry
    
    Close #filenum

    FileCopy sDBLoc & cboRegions & "_Maint.log", sDropBoxLoc & "\" & cboRegions & "_Maint.log"

LeavE:
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Trim$(Error)
    Resume LeavE
    
End Function
