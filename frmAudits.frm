VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmAudits 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmAudits.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncludeVat 
      BackColor       =   &H00F7F3EA&
      Caption         =   "Include Vat"
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
      Left            =   7800
      TabIndex        =   16
      Top             =   1005
      Width           =   1290
   End
   Begin MSComCtl2.MonthView Cal 
      Height          =   2820
      Left            =   2955
      TabIndex        =   13
      Top             =   1890
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
      StartOfWeek     =   16384002
      TitleBackColor  =   13342061
      TitleForeColor  =   16777215
      CurrentDate     =   39972
   End
   Begin VB.ComboBox cboClients 
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
      Left            =   2235
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4005
   End
   Begin VB.CheckBox chkAllClients 
      BackColor       =   &H00F7F3EA&
      Caption         =   "All"
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
      Left            =   6480
      TabIndex        =   2
      Top             =   1005
      Width           =   540
   End
   Begin VB.CheckBox chkAllDates 
      BackColor       =   &H00FAF8F1&
      Caption         =   "All"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   1575
      Width           =   540
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   9360
      TabIndex        =   0
      Top             =   105
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
   Begin MSMask.MaskEdBox tedFrom 
      Height          =   375
      Left            =   2955
      TabIndex        =   5
      Top             =   1515
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
   Begin MyCommandButton.MyButton btnCallFrom 
      Height          =   360
      Left            =   3900
      TabIndex        =   6
      Top             =   1530
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
      Picture         =   "frmAudits.frx":77A8
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmAudits.frx":ABAD
      ShowFocus       =   -1  'True
   End
   Begin VSFlex8LCtl.VSFlexGrid grdAud 
      Height          =   6540
      Left            =   330
      TabIndex        =   12
      Top             =   2340
      Width           =   9105
      _cx             =   16060
      _cy             =   11536
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   16052193
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16052193
      GridColor       =   -2147483633
      GridColorFixed  =   16052193
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   60
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAudits.frx":AECF
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin MyCommandButton.MyButton btnCalTo 
      Height          =   360
      Left            =   5955
      TabIndex        =   9
      Top             =   1515
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmAudits.frx":AF9C
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmAudits.frx":E3A1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton btnGo 
      Height          =   360
      Left            =   8565
      TabIndex        =   11
      Top             =   1500
      Width           =   540
      _ExtentX        =   953
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
      Caption         =   "Go >"
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
   Begin MSMask.MaskEdBox tedTo 
      Height          =   375
      Left            =   4995
      TabIndex        =   8
      Top             =   1515
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MyCommandButton.MyButton btnPrint 
      Height          =   360
      Left            =   1875
      TabIndex        =   17
      Top             =   9090
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
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
      Caption         =   "Print"
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
   Begin MyCommandButton.MyButton btnCreate 
      Height          =   360
      Left            =   315
      TabIndex        =   18
      Top             =   9090
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
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
      Caption         =   "Create Report"
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
   Begin VB.Label labelTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.0"
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
      Left            =   8775
      TabIndex        =   20
      Top             =   9135
      Width           =   255
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7245
      TabIndex        =   19
      Top             =   9135
      Width           =   465
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
      Left            =   2205
      TabIndex        =   15
      Top             =   915
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.Label lblClients 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Client"
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
      Left            =   1065
      TabIndex        =   14
      Top             =   1005
      Width           =   1110
   End
   Begin VB.Label lblDates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Audits"
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
      Left            =   1020
      TabIndex        =   3
      Top             =   1545
      Width           =   1170
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   2445
      TabIndex        =   4
      Top             =   1545
      Width           =   465
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   4695
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "frmAudits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFromDate As Boolean
Public sRepName As String

Private Sub btnClose_Click()

    Unload Me
    
End Sub

Private Sub btnCreate_Click()

    If grdAud.Rows > 1 Then
    
        sRepName = GetSetting(App.Title, "REPORTNAME", Key:=App.Title & "REPORTNAME")
        
        sRepName = InputBox("Ente File Name or <Enter> for default", "Audit Report File Name", sRepName)
        
        If sRepName <> "" Then
            SaveSetting appname:=App.Title, Section:="REPORTNAME", Key:=App.Title & "REPORTNAME", Setting:=sRepName
            
            If InStr(sRepName, ".") <> 0 Then
                If InStr(sRepName, ".doc") = 0 Then
                    sRepName = sRepName & ".doc"
                End If
            Else
                sRepName = Replace(sRepName, ".", "") & ".doc"
            End If
            ' Add a .doc and clear a dot if its there already
            
            grdAud.AddItem ""
            grdAud.AddItem lblTotal.Caption & vbTab & vbTab & vbTab & vbTab & vbTab & labelTotal
            grdAud.AddItem ""
            ' Include total here before report runs
            
            If CreateReport(sRepName) Then
                btnPrint.Enabled = True
            End If
        
        End If
    End If
    

End Sub

Private Sub btnGo_Click()
    
    If DatesOk() Then
    
        If DoAuditReport() Then
            btnCreate.Enabled = True

        End If
        
    End If


End Sub

Private Sub btnPrint_Click()
Dim objfile As Object

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object
    
    Set WriteWord = New Word.Application
    
    WriteWord.Visible = False
    
    WriteWord.Application.Documents.Open sRepName, , vbTrue
    ' open report
    
    WriteWord.Application.NormalTemplate.Saved = True
    
    WriteWord.Application.PrintOut -1
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    Set objfile = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    MsgBox "Problem Printing report " & Trim(Error)


End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)

    If bFromDate Then
         tedFrom.Text = Format(Cal.Value, sDMY)
    Else
         tedTo.Text = Format(Cal.Value, sDMY)
    End If
    
    Cal.Visible = False

End Sub

Private Sub cboClients_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboClients_LostFocus()
    lblClients.ForeColor = sBlack

End Sub

Private Sub chkAllClients_Click()

    If chkAllClients.Value = 1 Then
        cboClients.ForeColor = sLightGrey
    Else
        cboClients.ForeColor = sBlack
    End If

End Sub

Private Sub chkAllDates_Click()

    If chkAllDates.Value = 1 Then
        
        tedFrom.ForeColor = sLightGrey
        tedTo.ForeColor = sLightGrey
    Else
        tedFrom.ForeColor = sBlack
        tedTo.ForeColor = sBlack
    
    End If


End Sub

Private Sub Form_Load()

    gbOk = GetClientList()

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        btnClose_Click
        
    End If

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


Private Sub lblRegion_Click()

End Sub
Private Sub btnCallFrom_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedFrom) Then
            Cal.Value = tedFrom
        Else
            Cal.Value = Format(Now, sDMY)
        End If
        bFromDate = True
        Cal.Top = tedFrom.Top + tedFrom.Height
        Cal.Left = tedFrom.Left
        Cal.Visible = True
        bSetFocus Me, "Cal"
    End If

End Sub

Private Sub btnCalTo_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedTo) Then
            Cal.Value = tedTo
        Else
            Cal.Value = Format(Now, sDMY)
        End If
        bFromDate = False
        Cal.Top = tedTo.Top + tedTo.Height
        Cal.Left = tedTo.Left
        Cal.Visible = True
        bSetFocus Me, "Cal"
    End If
    

End Sub

Public Function DatesOk()

    If chkAllDates.Value = 1 Then
        DatesOk = 1
    ElseIf IsDate(tedFrom) Then
        
        If IsDate(tedTo) Then
        
            If DateValue(tedFrom) <= DateValue(tedTo) Then
                DatesOk = 1
            Else
            
                MsgBox "From Date must be Older than To Date"
                bSetFocus Me, "tedFrom"
            End If
        Else
            bSetFocus Me, "tedTo"
        End If
    Else
        bSetFocus Me, "tedFrom"
    End If

End Function

Public Function GetClientList()
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    bHourGlass True
    
    With cboClients
    
        .Clear
    
        Set rs = SWdb.OpenRecordset("Select * FROM tblClients WHERE chkActive = True", dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do
                .AddItem rs("txtName")
                .ItemData(.NewIndex) = rs("ID") + 0
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    
    End With
    
    bHourGlass False
    
    GetClientList = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientList") Then Resume 0
    Resume CleanExit


End Function


Public Function DoAuditReport()
Dim rs As Recordset
Dim sSql As String
Dim sWhereAnd As String
Dim curTotal As Currency

    On Error GoTo ErrorHandler
    
    sWhereAnd = " WHERE "
    grdAud.Rows = 1
    
    'Clients
    If chkAllClients.Value = 1 Or cboClients = "" Then
    
        sSql = ""
        
    ElseIf cboClients.ListIndex <> -1 Then
        sSql = sWhereAnd & "tblDates.ClientID = " & Trim$(cboClients.ItemData(cboClients.ListIndex))
        sWhereAnd = " AND "
    End If
    
    'DATES
    If chkAllDates.Value = 1 Then
    
    ElseIf IsDate(tedFrom) And IsDate(tedTo) Then
        sSql = sSql & sWhereAnd & "[On] >= #" & Format(tedFrom, "mm/dd/yy") & "# AND [On] <= #" & Format(tedTo, "mm/dd/yy") & "#"
        sWhereAnd = " AND "
    End If
    
    curTotal = 0
    
    Set rs = SWdb.OpenRecordset("SELECT * from tblDates INNER JOIN tblClients ON tblDates.ClientID = tblClients.ID " & sSql & " Order By [On] DESC")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
        
            
            Select Case chkIncludeVat
            
                Case 1
                    grdAud.AddItem rs("txtName") & vbTab & rs("On") & vbTab & rs("From") & vbTab & rs("To") & vbTab & rs("InvNumber") & vbTab & Format(rs("INVTotal"), "0.00")
                    If Not IsNull(rs("INVTotal")) Then
                        curTotal = curTotal + rs("INVTotal")
                    End If
                
                Case 0
                    grdAud.AddItem rs("txtName") & vbTab & rs("On") & vbTab & rs("From") & vbTab & rs("To") & vbTab & rs("InvNumber") & vbTab & Format(rs("INVAmount"), "0.00")
                    If Not IsNull(rs("INVAmount")) Then
                        curTotal = curTotal + rs("INVAmount")
                    End If
            End Select
            
            grdAud.RowData(grdAud.Rows - 1) = rs("tblDates.ID") + 0
            
            
            rs.MoveNext
        Loop While Not rs.EOF
    
        DoAuditReport = True
        
    End If
    
    rs.Close
    
    Select Case chkIncludeVat
        Case 1
         lblTotal.Caption = "Total (Incl Vat)"
         labelTotal.Caption = Format(curTotal, "Currency")
         
        Case 0
         lblTotal.Caption = "Total (Excl Vat)"
         labelTotal.Caption = Format(curTotal, "Currency")
         
    End Select
        
Leave:
    Exit Function

ErrorHandler:
    
    MsgBox "Error: " & Trim$(Error)
    Resume Leave

    

End Function

Private Sub tedFrom_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedFrom_LostFocus()
    lblFrom.ForeColor = sBlack

End Sub

Private Sub tedTo_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedTo_LostFocus()
    lblTo.ForeColor = sBlack

End Sub
Public Function CreateReport(sName As String)
Dim objfile As Object
Dim Reportdoc
Dim sReport As String
Dim sNextPage As String

Dim iTotPages As Integer
Dim iRow As Integer
Dim iCol As Integer
Dim iRows As Integer
Dim iCnt As Integer
Dim NextPageDoc
Dim iDataPointer As Integer
Dim iLinesperpage As Integer
Dim iRowCounter As Integer
Dim iCols As Integer

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    bHourGlass True
    
    iLinesperpage = 41
    
'    sReport = sDBLoc & "\" & sName
'    sNextPage = sDBLoc & "\NextPage.Doc"
   
    sReport = sName
    sNextPage = sDBLoc & "\NextPage.Doc"
    
    On Error Resume Next
            
    Kill sReport
    Kill sNextPage
    ' DELETE Report HERE!
        '
    On Error GoTo ErrorHandler
    
    Set WriteWord = New Word.Application
    WriteWord.Visible = False
    
    iRows = grdAud.Rows

    If iRows = 0 Then
        iTotPages = 1

    ElseIf (iRows Mod (iLinesperpage)) = 0 Then
        iTotPages = iRows / iLinesperpage

    Else
        iTotPages = Int(iRows / (iLinesperpage)) + 1
    End If
    ' Get how many pages of a report are going to be produced
    ' -1 = trailer line
    
    iDataPointer = grdAud.FixedRows
    ' Initialize the Row index pointer to point to the first line of data past the header.
 
    For iCnt = 1 To iTotPages
  '      If iCnt = 4 Then Stop
        
         bHourGlass True
         
         If iCnt = 1 Then
            Set Reportdoc = WriteWord.Documents.Add(sDBLoc & "\Templates\AuditReport.dot")
         Else
            Set NextPageDoc = WriteWord.Documents.Add(sDBLoc & "\Templates\AuditReport.dot")
         End If
         WriteWord.Visible = False
         
         WriteWord.ActiveDocument.UndoClear
        
         With WriteWord.ActiveDocument.Bookmarks
        
            ' HEADER
            
            If chkAllClients Then
                .Item("Client").Range.Text = "All Clients"
            Else
                .Item("Client").Range.Text = cboClients
            End If
            
            .Item("Date").Range.Text = Format(Now, sDMMYY)
            
            .Item("Region").Range.Text = gbRegion
            ' Region
            
            If chkAllDates Then
                .Item("From").Range.Text = "  All"
                .Item("To").Range.Text = "  All"
            Else
                .Item("From").Range.Text = tedFrom
                .Item("To").Range.Text = tedTo
            End If
            
            .Item("PageNo").Range.Text = Trim$(iCnt) & " of " & Trim$(iTotPages)
            ' Page No
            
            WriteWord.Visible = False

            iCols = grdAud.Cols
            
            ' GRID HEADER
            For iCol = 1 To iCols - 1
                .Item("Cell" & "R" & Trim$(iRow) & "C" & Trim$(iCol)).Range.Text = Trim$(grdAud.Cell(flexcpTextDisplay, iRow, iCol))
            Next
            
            ' DATA
            
            iRowCounter = grdAud.FixedRows
            ' For each page init row counter
            
            If iRows > iRowCounter Then
            ' do a check here incase there are no records and just want to show a blank page
                Do
                
                    For iCol = 0 To iCols - 1
                        .Item("Cell" & "R" & Trim$(iRowCounter) & "C" & Trim$(iCol)).Range.Text = grdAud.Cell(flexcpTextDisplay, iDataPointer, iCol)
                    Next
                    
                    WriteWord.Visible = False

                    iRowCounter = iRowCounter + 1
                    iDataPointer = iDataPointer + 1
            
                Loop While (iRowCounter < iLinesperpage) And (iDataPointer < grdAud.Rows - 1)
            
            End If
            
            If iCnt = 1 Then
                
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
                
            End If

        End With
    Next
    
    CreateReport = True

CloseWordStuff:
    
    bHourGlass True
    
    On Error Resume Next
    
    WriteWord.Application.NormalTemplate.Saved = True
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    
    Set objfile = Nothing
    Set Reportdoc = Nothing
    Set NextPageDoc = Nothing
    
    Kill sNextPage

    MsgBox "Report Created (" & sReport & ")"

CleanExit:
    
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    ' If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("CreateReport") Then Resume 0
    Resume CloseWordStuff

    
End Function

