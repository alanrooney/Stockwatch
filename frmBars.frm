VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmBars 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmBars.frx":0000
   ScaleHeight     =   4830
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBar 
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
      Left            =   765
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3450
      Width           =   2400
   End
   Begin VSFlex8LCtl.VSFlexGrid grdBar 
      Height          =   2670
      Left            =   360
      TabIndex        =   2
      Top             =   765
      Width           =   3135
      _cx             =   5530
      _cy             =   4710
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   9929356
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   15658209
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBars.frx":3DC0
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   0
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
   Begin MyCommandButton.MyButton btnQuit 
      Height          =   495
      Left            =   2625
      TabIndex        =   3
      Top             =   4035
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
   Begin MyCommandButton.MyButton btnCloseInvoice 
      Height          =   255
      Left            =   3525
      TabIndex        =   4
      Top             =   75
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
   Begin MyCommandButton.MyButton btnDelete 
      Height          =   360
      Left            =   3165
      TabIndex        =   5
      Top             =   3465
      Width           =   330
      _ExtentX        =   582
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
      Caption         =   "X"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Watch - Bars"
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
      Left            =   165
      TabIndex        =   6
      Top             =   60
      Width           =   1695
   End
   Begin VB.Label lblBar 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bar"
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
      Left            =   405
      TabIndex        =   0
      Top             =   3510
      Width           =   2745
   End
End
Attribute VB_Name = "frmBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lBrID As Long
Public lClientID As Long


Private Sub cmdOk_Click()

    Unload Me
    

End Sub

Private Sub cmdQuit_Click()

    Unload Me
    

End Sub

Private Sub btnCloseInvoice_Click()

    btnQuit_Click
    

End Sub

Private Sub btnDelete_Click()

    If MsgBox("This will remove all counts for this bar and should only be done after a stock take is closed. Continue with Delete?", vbDefaultButton2 + vbYesNo + vbQuestion, "Delete Bar") = vbYes Then
    
        Screen.MousePointer = 11
    
        SWdb.Execute "DELETE from tblBars WHERE ID = " & Trim$(lBrID)
    
        SWdb.Execute "DELETE FROM tblBarCount WHERE ClientID = " & Trim$(lClientID) & " AND BarID = " & Trim$(lBrID)
    
        Screen.MousePointer = 0
    
    ' also remove counts for this bar
    
    End If
    
    gbOk = GetBars()
    lBrID = 0
    txtBar.Text = ""
    
    bSetFocus Me, "txtBar"
    
End Sub

Private Sub btnQuit_Click()

    Unload Me
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me
    

End Sub

Private Sub Form_Load()

    ' Ver 440 This is an all new form
    

    gbOk = GetBars()


End Sub

Public Function GetBars()
Dim rs As Recordset


    grdBar.Rows = 0
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblbars WHERE ClientID = " & Trim$(lClientID))
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            grdBar.AddItem rs("Bar") & ""
            grdBar.RowData(grdBar.Rows - 1) = rs("ID") + 0
        
            rs.MoveNext
        Loop While Not rs.EOF
        
    End If
    rs.Close
    GetBars = True
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetBars ") Then Resume 0
    Resume CleanExit

    
End Function

Private Sub Form_Unload(Cancel As Integer)

    If grdBar.Rows = 1 Then
    
        If MsgBox("This is used for multi bar stock counts. Use of this is not recommended for only one bar. Continue?", vbDefaultButton2 + vbYesNo + vbExclamation, "Multiple Bars") = vbNo Then
        
            Cancel = True
        End If
    End If

End Sub

Private Sub grdBar_Click()

    If grdBar.Row > -1 Then
        txtBar = grdBar.Text
        lBrID = grdBar.RowData(grdBar.Row)
        btnDelete.Enabled = True
    
        bSetFocus Me, "txtBar"
    End If
    
    
End Sub

Private Sub txtBar_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtBar_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
    
     If Len(Trim$(txtBar)) > 0 Then
      
      If BarUnique(lBrID) Then
        
        If AddBar(lBrID) Then
        
            lBrID = 0
            gbOk = GetBars()
            txtBar.Text = ""
            btnDelete.Enabled = False

        End If
        
        bSetFocus Me, "txtBar"
      Else
        MsgBox "Bar already listed"
      End If
      
     Else
        MsgBox "Please enter Bar Location"
        bSetFocus Me, "txtBar"
     End If
     
    End If
    
End Sub

Public Function AddBar(lID As Long)
Dim rs As Recordset


    Set rs = SWdb.OpenRecordset("SELECT * FROM tblbars WHERE ID = " & Trim$(lID) & " AND ClientID = " & Trim$(lClientID))
    If Not (rs.EOF And rs.BOF) Then
        rs.Edit
    Else
        rs.AddNew
    End If
    
    rs("ClientID") = lClientID + 0
    rs("Bar") = txtBar
    rs.Update
    rs.Close
    AddBar = True
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("AddBar ") Then Resume 0
    Resume CleanExit

    

End Function

Public Function BarUnique(lID As Long)
Dim iRow As Integer

    
    For iRow = 0 To grdBar.Rows - 1
    
        If lID = 0 Or (lID <> grdBar.RowData(iRow)) Then
            If grdBar.Cell(flexcpTextDisplay, iRow, 0) = txtBar Then Exit Function
        End If
        
'
    Next
  
  BarUnique = True

End Function

Private Sub txtBar_LostFocus()
    lblBar.ForeColor = sBlack

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

