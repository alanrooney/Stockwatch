VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmCopyUpdate 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmCopyUpdate.frx":0000
   ScaleHeight     =   1650
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   4005
      TabIndex        =   1
      Top             =   135
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
   Begin VB.Label labelMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restarting Stockwatch ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   45
      TabIndex        =   0
      Top             =   825
      Width           =   4425
   End
End
Attribute VB_Name = "frmCopyUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gbOk As Boolean
Public lAudf As Long
Public endofpause As Double

Private Sub btnClose_Click()

    If MsgBox("Are you sure you want to interrupt the Update Process?", vbDefaultButton2 + vbYesNo + vbQuestion, "Update In Progress") = vbYes Then
    
        End
    
    End If
    
End Sub

Private Sub Form_Activate()
Dim sDBLoc As String
Dim objFile As Object
Dim sFile, sFilesCol, sf
Dim bFileExists As Boolean
Dim sNewRev As String
Dim sFileName As String

    
    sNewRev = Trim$(Command$())
    
    If sNewRev <> "" Then
    
        Print #lAudf, "New Revision: " & sNewRev
        
        ' copy source to destination trapping for overwrite
    
        Set objFile = CreateObject("Scripting.FileSystemObject")
        ' create object
    
            
        sDBLoc = "" & GetSetting("Stockwatch", "DB", "StockwatchDB") & ""
        ' get the DB Location from the registry
        
        ' OPEN AUDIT FILE
        
        Print #lAudf, "Checking for file: " & sDBLoc & "/" & "Stockwatch Ver" & sNewRev & ".exe"
        bFileExists = objFile.FileExists(sDBLoc & "/" & "Stockwatch Ver" & sNewRev & ".exe")
        ' Make sure New fies exists
        
        If bFileExists Then
        
        
            ' DELETE OLD    carry on if none
            
            On Error GoTo NoOldFile
            
            sFileName = sDBLoc & "/" & "StockwatchOLD.exe"
            Set sFile = objFile.GetFile(sFileName)
            sf = sFile.Delete
            
            
            On Error GoTo NoPrevFileToRename
            
            ' RENAME PREV AS OLD    trap error
            
            Name sDBLoc & "/" & "StockwatchPREV.exe" As sDBLoc & "/" & "StockwatchOLD.exe"
            
            On Error GoTo NoCurFileToRename
            
            ' RENAME CUR AS PREV    trap error
            
            Name sDBLoc & "/" & "Stockwatch.exe" As sDBLoc & "/" & "StockwatchPREV.exe"
            
            ' COPY NEW VER TO STOCKWATCH.EXE trap error
            
            On Error GoTo NoCopy
            FileCopy sDBLoc & "/" & "Stockwatch Ver" & sNewRev & ".exe", sDBLoc & "/" & "Stockwatch.exe"
            
            ' SPAWN StockWatch
        
            Print #lAudf, "Update " & sNewRev & " Complete"
            
Leave:
            Pause 3000
            
            Shell sDBLoc & "/" & "Stockwatch.exe", vbNormalFocus
        
            Print #lAudf, "Restart stockwatch ..."
            
            Close #lAudf
            
            
            
        Else
            Print #lAudf, "Update File does not exist: " & sDBLoc & "/" & "Stockwatch Ver" & sNewRev & ".exe"
            
            MsgBox "Update file Stockwatch Ver" & sNewRev & ".exe Does not exist. Exiting Update"
            Close #lAudf
            End
        End If
    
    End If
    
    End
    
NoOldFile:
    Resume Next
    
NoPrevFileToRename:
    Resume Next
    
NoCurFileToRename:
    Print #lAudf, "No Current file: " & sDBLoc & "/" & "Stockwatch.exe"
    Resume Next
    
NoCopy:
    Print #lAudf, "New file not copied: " & sDBLoc & "/" & "StockwatchVer " & sNewRev & ".exe"
    Resume Leave

    

End Sub

Public Function OpenAuditFile(tim As Date)

On Error GoTo ErrorHandler

    lAudf = FreeFile
    Open CurDir & "\SW" & Trim$(Format(DateValue(Now) - (Format(tim, "w") - 2), "ddmmyy")) & ".CSV" For Append As #lAudf
    OpenAuditFile = True
    ' use mondays date of this week for the file name
    
    Print #lAudf, Format(Now, "dd/mm/yy hh:mm:ss") & "," & "CopyUpdate Started"
    ' send out a small header
    
    OpenAuditFile = True

    Exit Function

ErrorHandler:
  
End Function

Private Sub Form_Load()
    gbOk = OpenAuditFile(Now)

End Sub
Public Sub Pause(millisec)
Dim X As Integer

  endofpause# = Timer + millisec / 1000
  Do
      X% = DoEvents()
  Loop While Timer < endofpause#

End Sub

