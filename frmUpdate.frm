VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUpdate 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StockWatch Update Program"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProblem 
      Caption         =   "Send StockWatch.mdb >>>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   9
      Top             =   2340
      Width           =   2565
   End
   Begin VB.TextBox txtProgLocation 
      BackColor       =   &H8000000F&
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
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   450
      Width           =   8025
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   10740
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtProgram 
      BackColor       =   &H8000000F&
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "StockWatch"
      Top             =   450
      Width           =   2715
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Check For Update >>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8670
      TabIndex        =   0
      Top             =   2340
      Width           =   2565
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   1
      Top             =   2340
      Width           =   1665
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C5DDFE&
      Height          =   495
      Left            =   2310
      TabIndex        =   8
      Top             =   2340
      Width           =   6015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   270
      X2              =   11220
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   270
      X2              =   11220
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label lblProgLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "&Location"
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
      Height          =   285
      Left            =   3210
      TabIndex        =   7
      Top             =   180
      Width           =   1785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Program"
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
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   180
      Width           =   1785
   End
   Begin VB.Label labelLocation 
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
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   1620
      Width           =   10995
   End
   Begin VB.Label lblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "&DropBox Location"
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
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   1350
      Width           =   1785
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sDBLoc As String
Public SWdb As Database
Public gbOK As Boolean
Public lAUDf As Long


Private Sub cmdChange_Click()
                
    labelLocation = GetDropBoxLocation()
            
    SaveSetting appname:=txtProgram, Section:="PRG", Key:=txtProgram & "DropBox", Setting:=labelLocation

End Sub

Private Sub cmdProblem_Click()
Dim objFile As Object
Dim sFileName As String
Dim sMDBName As String

    
    If MsgBox("Please make sure " & txtProgram & " is not running. Continue with Copy?", vbDefaultButton1 + vbYesNo + vbQuestion, "Exit StockWatch") = vbYes Then
    
        On Error GoTo DropBoxLocationProblem
        
        sMDBName = txtProgram & ".mdb"
        
        sFileName = txtProgLocation & "\" & sMDBName
        ' file name without the dropbox folder name
        
        If MsgBox("Send Mdb file to dropbox?", vbDefaultButton1 + vbYesNo + vbQuestion, "Send Mdb File") = vbYes Then
        
            Set objFile = CreateObject("Scripting.FileSystemObject")
            
            '=========================================================
            ' COPY MDB FILE TO DROPBOX
            
            On Error GoTo CopyFileToDropBoxProblem
            
            LogMsg Me, "Copying Mdb File to DropBox", sFileName
            
            objFile.copyfile sFileName, labelLocation & "\" & sMDBName, True
            
            '=========================================================
            ' COPY AUDIT FILE ALSO
            
            On Error Resume Next
            
            Close #lAUDf
            
            objFile.copyfile txtProgLocation & "\SW" & Trim$(Format(DateValue(Now) - (Format(Now, "w") - 2), "ddmmyy")) & ".CSV", labelLocation & "\SW" & Trim$(Format(DateValue(Now) - (Format(Now, "w") - 2), "ddmmyy")) & ".CSV", True
            
            
            '=========================================================
            ' ALL DONE
            
            MsgBox "Copy Started .. This will take a few minutes"
            LogMsg Me, "Copy Complete ", ""
        
        End If
        
    End If
    
Leave:
    
    Exit Sub
    
DropBoxLocationProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

CopyFileToDropBoxProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

End Sub

Private Sub cmdUpdate_Click()
Dim objFile As Object
Dim sFileName As String
Dim sProgName As String

    
    If MsgBox("Please make sure " & txtProgram & " is not running. Continue with Update?", vbDefaultButton1 + vbYesNo + vbQuestion, "Exit StockWatch") = vbYes Then
    
    
    
        ' Using DropBox Location and file name - check for updates
        On Error GoTo DropBoxLocationProblem
        
        sProgName = txtProgram & ".exe"
        
        ComDlg.InitDir = labelLocation
    '    ComDlg.Filter = txtProgram & " Ver"
        ComDlg.ShowSave
        sFileName = Replace(ComDlg.FileName, labelLocation & "\", "")
        ' file name without the dropbox folder name
        
        If Left(sFileName, Len(txtProgram)) = txtProgram Then
        
            If MsgBox("Ready to Update Program?", vbDefaultButton1 + vbYesNo + vbQuestion, "Update Program") = vbYes Then
            
                Set objFile = CreateObject("Scripting.FileSystemObject")
                
                '============================================================
                ' IF NEW FILE ALREADY EXISTS - DELETE IT
                
                On Error Resume Next
                
                If objFile.FileExists(txtProgLocation & "\" & sFileName) Then
                
                    LogMsg Me, "Deleting old file ", txtProgLocation & "\" & sFileName
                    
                    objFile.Deletefile txtProgLocation & "\" & sFileName, True
                    ' Delete any preexisting program with same name
                
                End If
                
                '=============================================================
                ' COPY FILE FROM DROPBOX TO CLIENT FOLDER
                
                On Error GoTo CopyFileFromDropBoxProblem
                
                LogMsg Me, "Copying New program from DropBox ", labelLocation & "\" & sFileName
                
                objFile.copyfile labelLocation & "\" & sFileName, txtProgLocation & "\" & sFileName, True
            
                '============================================================
                ' IF OLD PROGRAM FILE EXISTS - DELETE IT
                
                On Error Resume Next
                
                If objFile.FileExists(txtProgLocation & "\" & sProgName) Then
                    
                    LogMsg Me, "Deleting Old program  ", txtProgLocation & "\" & sProgName
                    
                    objFile.Deletefile txtProgLocation & "\" & sProgName
                    ' Delete any preexisting program with same name
                End If
                ' Delete Old Program
                
                '============================================================
                ' COPY NEW FILE (StockWatch Ver130.exe) TO OLD FILE NAME (StockWatch.exe)
                
                On Error GoTo CopyNewProgramToTempProblem
                
                LogMsg Me, "Copying New Program to ", txtProgLocation & "\" & sProgName
    
                objFile.copyfile txtProgLocation & "\" & sFileName, txtProgLocation & "\" & sProgName
                ' Copy New Program to Temp Program Name
                
                LogMsg Me, "Update Complete ", ""
            
            End If
        
        Else
            MsgBox "You must Select the " & txtProgram & " Program to do the Update"
        End If
    
    End If
    
Leave:
    
    Exit Sub
    
DropBoxLocationProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

CopyFileFromDropBoxProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

CopyNewProgramToTempProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

MoveFileFromTempProblem:
    LogMsg Me, Trim$(Error), ""
    Resume Leave

End Sub

Private Sub Form_Load()
    
    If Not App.PrevInstance Then
    ' as long as its not running already...
    
        txtProgLocation = "" & GetSetting(txtProgram, "DB", txtProgram & "DB") & ""
        ' get the DB Location from the registry
        
        gbOK = OpenAuditFile(Now)
        ' open audit file early
        
        If txtProgLocation = "" Then
            MsgBox "No Program Location"
            End
        
        Else
            
            labelLocation = GetSetting(txtProgram, "PRG", txtProgram & "DropBox") & ""
        
            If labelLocation = "" Then
                    
                labelLocation = GetDropBoxLocation()
            
                SaveSetting appname:=txtProgram, Section:="PRG", Key:=txtProgram & "DropBox", Setting:=labelLocation
            
            End If
        End If
        
    Else
        MsgBox App.Title & " already running"
        End
    
    End If

End Sub


Public Sub bSetFocus(fname As Form, sCtrl As String)

    If fname.Controls(sCtrl).Enabled And fname.Controls(sCtrl).Visible Then
        fname.Controls(sCtrl).SetFocus
    End If

End Sub

Public Function GetDropBoxLocation()


    MsgBox "Locate and Open TESTFILE.txt in the DropBox Folder on your system in the next screen.(Try under My Documents)"
    
    ComDlg.Action = 1
    ComDlg.ShowSave
    GetDropBoxLocation = Replace(ComDlg.FileName, "\TESTFILE.txt", "")
    

End Function
Public Function OpenAuditFile(tim As Date)

On Error GoTo ErrorHandler

    lAUDf = FreeFile
    Open txtProgLocation & "\SW" & Trim$(Format(DateValue(Now) - (Format(tim, "w") - 2), "ddmmyy")) & ".CSV" For Append As #lAUDf
    OpenAuditFile = True
    ' use mondays date of this week for the file name
    
    Print #lAUDf, Format(Now, "dd/mm/yy hh:mm:ss") & "," & App.Title & " Started"
    ' send out a small header
    
    OpenAuditFile = True

    Exit Function

ErrorHandler:
  
End Function
Public Sub LogMsg(Frm As Form, sMsg As String, sAudMsg As String)

    On Error GoTo ErrorHandler
    
    If (sMsg & sAudMsg = "") Or sMsg <> "" Then
        Frm.lblMsg.Caption = sMsg
        DoEvents
    
    End If
    ' display it on the form
    
    If sAudMsg <> "" Then
        Print #lAUDf, Format$(Now, "dd/mm/yy hh:mm:ss") & "," & Screen.ActiveForm.Name & "," & sMsg & "," & sAudMsg
    End If
    ' save it in the log file
    
    Exit Sub
    
ErrorHandler:
    Print #lAUDf, Format$(Now, "dd/mm/yy hh:mm:ss") & "," & sMsg & "," & sAudMsg
   
End Sub

