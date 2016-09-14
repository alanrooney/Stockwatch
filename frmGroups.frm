VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLU Groups"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
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
      Left            =   3090
      TabIndex        =   6
      Top             =   6450
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   465
      Left            =   3120
      TabIndex        =   5
      Top             =   450
      Width           =   705
   End
   Begin VB.TextBox txtGroupName 
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
      Left            =   1140
      TabIndex        =   3
      Top             =   480
      Width           =   1785
   End
   Begin VSFlex8LCtl.VSFlexGrid grdGroups 
      Height          =   5925
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   2895
      _cx             =   5106
      _cy             =   10451
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGroups.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Label Label1 
      Caption         =   "Name As:"
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
      Left            =   180
      TabIndex        =   4
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label labelGroupNo 
      BackColor       =   &H8000000E&
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
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   30
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Group No                - NOT Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   2895
   End
End
Attribute VB_Name = "frmGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSelectGroup As String
Public iGroupNo As Integer
Public lCLid As Long
Public bStop As Boolean


Private Sub cmdAdd_Click()

    sSelectGroup = txtGroupName

    gbOk = AddGroup(iGroupNo, sSelectGroup)

    Unload Me
    
End Sub

Private Sub cmdStop_Click()

    bStop = True

    Unload Me


End Sub

Private Sub Form_Load()

    gbOk = ShowGroups()
    DoEvents
    bStop = False
    
    sSelectGroup = ""
    labelGroupNo = Trim$(iGroupNo)
    Screen.MousePointer = 0
    
End Sub

Public Function ShowGroups()
Dim rs As Recordset

    grdGroups.Rows = 0
    
    Set rs = SWdb.OpenRecordset("SELECT * from tblPLUGroup ORDER BY txtGroupNumber")
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
        
        Do

            If grdGroups.FindRow(rs("txtDescription"), , 0) < 0 Then
                grdGroups.AddItem rs("txtDescription")
            End If
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
   Exit Function

ErrorHandler:
    MsgBox Trim$(Error)
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit

End Function

Private Sub Form_Unload(Cancel As Integer)
    iGroupNo = 0

End Sub

Private Sub grdGroups_Click()

    sSelectGroup = grdGroups.Cell(flexcpTextDisplay, grdGroups.Row, 0)
    
    gbOk = AddGroup(iGroupNo, sSelectGroup)
    
    
    Unload Me
    
End Sub

Public Function AddGroup(iGpNo As Integer, sGroup As String)
Dim rs As Recordset

    Set rs = SWdb.OpenRecordset("SELECT * FROM tblPLUGroup WHERE ClientID = " & Trim$(lCLid) & " AND txtGroupNumber = " & Trim$(iGpNo))
    If (rs.EOF And rs.BOF) Then
    
        Set rs = SWdb.OpenRecordset("tblPLUGroup")
        rs.Index = "PrimaryKey"
        rs.AddNew
        rs("ClientID") = lCLid
        rs("txtGroupNumber") = iGpNo
        rs("txtDescription") = sGroup
        rs("chkActive") = True
        rs.Update
    
    End If
    
CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
   Exit Function

ErrorHandler:
    MsgBox Trim$(Error)
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit


End Function
