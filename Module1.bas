Attribute VB_Name = "Module1"
Option Explicit

Public gbOk As Boolean
Public endofpause As Double
Public lAUDf As Long
Public SWdb As Database
Public iErrCount As Integer
Public Const sBlack = "&H80000012"
Public gbCnt As Integer
Public DBGf As Integer

Function gbOpenDB(mainfrm As Form) As Boolean
Dim sDBLoc As String
    
    On Error GoTo ErrorHandler
    
    sDBLoc = "" & GetSetting(App.Title, "DB", App.Title & "DB") & ""
    ' get the DB Location from the registry
    
    ' remove the extension since customer windows expl might not be showing extensions
    
    If sDBLoc = "" Then
        sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & ")", "Invalid Database Location: " & sDBLoc)
        If sDBLoc = "exit" Then End
        SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
    End If
    
OpenDB:
    Set SWdb = OpenDatabase("" & sDBLoc, False, False)
'    LogMsg frmSubMan, "DataBase Opened", " File: " & sDBLoc
    gbOpenDB = True

CleanExit:
    Exit Function
    
ErrorHandler:
    sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & "\" & App.Title & ".Mdb)", "Invalid Database Location: " & sDBLoc)
    If sDBLoc = "exit" Then End
    SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
    Resume OpenDB
    
End Function

