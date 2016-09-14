Attribute VB_Name = "FranchiseMod"
Option Explicit

Public DBGf As Long
Public gbCnt As Integer
Public gbOk As Boolean
Public sDBLoc As String
Public swDB As Database
Public bAllowMove As Boolean
Public MoveX As Integer
Public MoveY As Integer
Public Const sKey = "Stockwatch Ireland Version 3.0"

Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public gbDropBox As String
Public iErrCount As Integer
Public gbsw1 As String
Public Const sBlack = "&H80000012"
Public sHORegion As String ' Head Office Region SW1


Public Function bSetupControl(fname As Form)
Dim iPreviousTab As Integer
Dim iTab As Integer
Dim sLastCtrl As String
    
    On Error GoTo ErrorHandler
    
    If Not fname.Controls(fname.ActiveControl.Name).Locked Then
    
        Select Case TypeName(fname.Controls(fname.ActiveControl.Name))
            Case "TextBox", "RichTextBox"
             If fname.Controls(fname.ActiveControl.Name).MultiLine Then
                 fname.Controls(fname.ActiveControl.Name).SelStart = Len(fname.Controls(fname.ActiveControl.Name).Text)
             Else
                 fname.Controls(fname.ActiveControl.Name).SelStart = 0
                 fname.Controls(fname.ActiveControl.Name).SelLength = Len(fname.Controls(fname.ActiveControl.Name).Text)
             End If
             
             fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
             
             fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
             
             
             If sLastCtrl <> "" Then
               
        ' grey if not filled
               
               fname.Controls("lbl" & Mid(sLastCtrl, 4, 50)).ForeColor = vbBlack
             End If
    
            Case "MaskEdBox", "DriveListBox"
             fname.Controls(fname.ActiveControl.Name).SelStart = 0
             fname.Controls(fname.ActiveControl.Name).SelLength = Len(fname.Controls(fname.ActiveControl.Name).Text)
             fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
    
             fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
             
             If sLastCtrl <> "" Then
                
        ' grey if not filled
                
                fname.Controls("lbl" & Mid(sLastCtrl, 4, 50)).ForeColor = vbBlack
             End If
            
            Case "HScrollBar", "ComboBox", "VSFlexGrid"
                fname.Controls("lbl" & Mid(fname.ActiveControl.Name, 4, 50)).ForeColor = &HFF0000
                fname.Controls(fname.ActiveControl.Name).BackColor = vbWhite
            
            Case ""
            Case Else
        End Select
        
             fname.lblGlow.Visible = True
             fname.lblGlow.Left = fname.ActiveControl.Left - 8
             fname.lblGlow.Top = fname.ActiveControl.Top - 8
             fname.lblGlow.Width = fname.ActiveControl.Width + 28
             fname.lblGlow.Height = fname.ActiveControl.Height + 30

'        fname.Controls(fname.ActiveControl.BackColor) = vbWhite
        
        sLastCtrl = fname.ActiveControl.Name
    
    End If
    
CleanExit:
    Exit Function
    
ErrorHandler:
    
    If Err = 340 Or Err = 730 Or Err = 438 Then
        Resume Next
        Resume 0
        
    Else
        Resume CleanExit
    End If
    
End Function

Public Sub bSetFocus(fname As Form, sCtrl As String)

    If fname.Controls(sCtrl).Enabled And fname.Controls(sCtrl).Visible Then
        fname.Controls(sCtrl).SetFocus
    End If

End Sub
Public Function GetRegions(frm As Form)
Dim rs As Recordset

    frm.cboRegions.Clear
    
    Set rs = swDB.OpenRecordset("SELECT Region FROM tblfranchisees ORDER BY REGION ASC")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            frm.cboRegions.AddItem rs("Region") & ""
        
            rs.MoveNext
        
        Loop While Not rs.EOF
    End If
    GetRegions = True
    
CleanExit:
    DBEngine.Idle dbRefreshCache
     ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
'    If CheckDBError("GetRegions") Then Resume 0
    Resume CleanExit


End Function
Public Function Decrypt(ByVal Encrypted As String, _
    sEncKey As String) As String

    Dim NewEncrypted As String
    Dim Letter As String
    Dim KeyNum As String
    Dim EncNum As String
    Dim encbuffer As Long
    Dim strDecrypted As String
    Dim Kdecrypt As String
    Dim lastTemp As String
    Dim LenTemp As Integer
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops

    If sEncKey = "" Then sEncKey = "WhiteKnight"

    ReDim encKEY(1 To Len(sEncKey))
    
    'Convert The Key For Decryption
    For i = 1 To Len(sEncKey$)
        KeyNum = Mid$(sEncKey$, i, 1) 'Get Letter i% in the Key
        encKEY(i) = Asc(KeyNum) 'Convert Letter i to Asc value
 
'if it the first letter just hold it
       If i = 1 Then Math = encKEY(i): GoTo nextone
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
               <= encKEY(i - 1) Then Math = Math - encKEY(i)
               'compares the value to the previous value and
               'then either adds/subtracts the value to the
               'Math total
        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
              <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    'This is part of what i'm doing to convert the encrypted
    'numbers to  Letters so it sort of encrypts the encrypted
    'message.
    temp$ = Encrypted 'hold the encrypted data


    For i = 1 To Len(temp)
        itempstr = TrimSpaces(Str(Asc(Mid(temp, i, 1)) - _
           100)) 'grab the next 2 numbers
           'If its a 2 digit number hold it and continue
        If Len(itempstr$) = 2 Then itempstr$ = itempstr$
          If i = Len(temp) - 2 Then LenTemp% = _
               Len(Mid(temp2, Len(temp2) - 3))
          If i = Len(temp) Then itempstr$ = _
              TrimSpaces(itempstr$): GoTo itsdone
          'If the number is a single digit then add a '0' to the
          'front then hold it
        If Len(TrimSpaces(itempstr$)) = 1 Then _
             itempstr$ = "0" & TrimSpaces(itempstr$)
        'Now check to see if a - number was created if so
        'cause an error message
        If Left(TrimSpaces(itempstr$), 1) = "-" Then _
             Err.Raise 20000, , "Unexpected Error"
           

itsdone:
        temp2$ = temp2$ & itempstr 'hold the first decryption
    Next i
    
    
    Encrypted = TrimSpaces(temp2$) 'set the encrypted data


    For i = 1 To Len(Encrypted) Step 3
        'Format the encrypted string for the second decryption
        NewEncrypted = NewEncrypted & _
            Mid(Encrypted, CLng(i), 3) & " "
    Next i

' Hold the last set of numbers to check it its the correct format
    lastTemp$ = TrimSpaces(Mid(NewEncrypted, _
         Len(NewEncrypted$) - 3))
         
         If Len(lastTemp$) = 2 Then
' If it = 2 then its not the Correct format and we need to fix it
        lastTemp$ = Mid(NewEncrypted, _
           Len(NewEncrypted$) - 1) 'Holds Last Number so a '0'
                                    'Can be added between them
'set it to the new format
        Encrypted = Mid(NewEncrypted, 1, _
           Len(NewEncrypted) - 2) & "0" & lastTemp$
Else
        Encrypted$ = NewEncrypted$ 'set the new format

    End If
    'The Actual Decryption
    For i = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i, 1) 'Hold Letter at index i
        EncNum = EncNum & Letter 'Hold the letters
        If Letter = " " Then 'we have a letter to decrypt
            encbuffer = CLng(Mid(EncNum, 1, _
              Len(EncNum) - 1)) 'Convert it to long and
                                 'get the number minus the " "
            strDecrypted$ = strDecrypted & Chr(encbuffer - _
               Math) 'Store the decrypted string
            EncNum = "" 'clear if it is a space so we can get
                        'the next set of numbers
        End If
    Next i

    Decrypt = strDecrypted

    Exit Function
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
Err.Raise 20001, , "You have entered the wrong encryption string"

End Function

Public Function Encrypt(ByVal Plain As String, _
  sEncKey As String) As String
    '*********************************************************
    'Coded WhiteKnight 6-1-00
    'This Encrypts A string by converting it to its ASCII number
    'but the difference is it uses a Key String it converts the
    'keystring to ASCII and adds it to the first ASCII Value the
    'key is needed to decrypt the text.  I do plan on changing
    'this some what but For Now its ok.  I've only seen it
    'cause an error when the wrong Key was entered while
     'decrypting.
    
    'Note That If you use the same letter more then 3 times in a
    'row then each letter after it if still the same is ignored
    '(ie aaa = aaaaaaaaa but aaa <> aaaza)
    'If anyone Can figure out a way to fix this please e-mail me
  '*********************************************************
    Dim encrypted2 As String
    Dim LenLetter As Integer
    Dim Letter As String
    Dim KeyNum As String
    Dim encstr As String
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops
    
    If sEncKey = "" Then sEncKey = sKey
    'Sets the Encryption Key if one is not set
    ReDim encKEY(1 To Len(sEncKey))
    
    'starts the values for the Encryption Key
        
    For i = 1 To Len(sEncKey$)
     KeyNum = Mid$(sEncKey$, i, 1) 'gets the letter at index i
     encKEY(i) = Asc(KeyNum) 'sets the the Array value
                             'to ASC number for the letter

           'This is the first letter so just hold the value
        If i = 1 Then Math = encKEY(i): GoTo nextone

        'compares the value to the previous value and then
        'either adds/subtracts the value to the Math total
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)

        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) >= _
           encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
          >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    For i = 1 To Len(Plain) 'Now for the String to be encrypted
        Letter = Mid$(Plain, i, 1) 'sets Letter to
                                   'the letter at index i
        LenLetter = Asc(Letter) + Math 'Now it adds the Asc
                                       'value of Letter to Math

'checks and corrects the format then adds a space to separate them frm each other
        If LenLetter >= 100 Then encstr = _
             encstr & Asc(Letter) + Math & " "

         'checks and corrects the format then adds a space
        'to separate them frm each other
        If LenLetter <= 99 Then encstr$ = encstr & "0" & _
          Asc(Letter) + Math & " "
    Next i


    'This is part of what i'm doing to convert the encrypted
    'numbers to Letters so it sort of encrypts the
    'encrypted message.
    temp$ = encstr 'hold the encrypted data
    temp$ = TrimSpaces(temp) 'get rid of the spaces
    itempnum% = Mid(temp, 1, 2) 'grab the first 2 numbers
    temp2$ = Chr(itempnum% + 100) 'Now add 100 so it
                                   'will be a valid char

    'If its a 2 digit number hold it and continue
    If Len(itempnum%) >= 2 Then itempstr$ = Str(itempnum%)
 
   'If the number is a single digit then add a '0' to the front
   'then hold it
    If Len(itempnum%) = 1 Then itempstr$ = "0" & _
        TrimSpaces(Str(itempnum%))
    
    encrypted2$ = temp2 'set the encrypted message
    
    For i = 3 To Len(temp) Step 2
        itempnum% = Mid(temp, i, 2) 'grab the next 2 numbers
  
      ' add 100 so it will be a valid char
        temp2$ = Chr(itempnum% + 100)

      'if its the last number we only want to hold it we
       'don't want to add a '0' even if its a single digit
        If i = Len(temp) Then itempstr$ = _
         Str(itempnum%): GoTo itsdone

'If its a 2 digit number hold it and continue
        If Len(itempnum%) = 2 Then itempstr$ = _
            Str(itempnum%)

        'If the number is a single digit then add a '0'
        'to the front then hold it
        If Len(TrimSpaces(Str(itempnum))) = 1 Then _
      itempstr$ = "0" & TrimSpaces(Str(itempnum%))

        'Now check to see if a - number was created
        'if so cause an error message
        If Left(TrimSpaces(Str(itempnum)), 1) = "-" Then _
          Err.Raise 20000, , "Unexpected Error"
           

itsdone:
           'Set The Encrypted message
        encrypted2$ = encrypted2 & temp2$
    Next i


    'Encrypt = encstr 'Returns the First Encrypted String
    Encrypt = encrypted2 'Returns the Second Encrypted String
    Exit Function 'We are outta Here
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Function
Private Function TrimSpaces(strstring As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strstring$, " ")
        DoEvents
         lngpos& = InStr(1&, strstring$, " ")
         strstring$ = Left$(strstring$, (lngpos& - 1&)) & _
            Right$(strstring$, Len(strstring$) - _
               (lngpos& + Len(" ") - 1&))
    Loop
     TrimSpaces$ = strstring$
End Function


Sub SetTranslucent(ThehWnd As Long, nTrans As Integer)
    On Error GoTo ErrorRtn

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    
'    If General_ChkTranslucent Then
    
        attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
        SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
        SetLayeredWindowAttributes ThehWnd, RGB(255, 255, 0), nTrans, LWA_ALPHA
'    End If
    
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
    
End Sub

Public Function CharOk(iAsc As Integer, i012 As Integer, sExp As String)
Dim iCnt As Integer

    Select Case i012
        Case 0  ' No Only
         If (iAsc > 47 And iAsc < 58) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
         
        Case 1  ' Alpha Only
         If (iAsc > 64 And iAsc < 91) Or (iAsc > 96 And iAsc < 123) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
         
        Case 2  ' Either Only
         If (iAsc > 47 And iAsc < 58) Or (iAsc > 64 And iAsc < 91) Or (iAsc > 96 And iAsc < 123) Or iAsc = 8 Or iAsc = 13 Or iAsc = 27 Or InStr(sExp, Chr$(iAsc)) > 0 Then
            CharOk = iAsc
         End If
    End Select

End Function


Public Function CopyFileToDropBox(sSource As String, sDest As String)
Dim objFile As Object

    On Error GoTo RenameError
            
    Set objFile = CreateObject("Scripting.FileSystemObject")
    
    If objFile.FileExists(sDest) Then
    ' See if destination file already exists and remove it
    
        Kill sDest
    End If
    
    Name sSource As sDest

'    Name sSource As "test.csv"

    CopyFileToDropBox = True

Leave:
    Exit Function

RenameError:
    MsgBox "Error: " & Trim$(Error) & " - Problem copying file from " & sSource & " to " & sDest
    Resume Leave

End Function

Public Sub SaveDebug(sSect As String, dbgmsg As String)

    On Error GoTo ErrorHandler

    DBGf = FreeFile
    
    Open CurDir & "Debug.CSV" For Append As #DBGf
    'local debug file
 
    With Screen.ActiveForm
        Print #DBGf, Format(Now, "dd/mm/yy hh:mm:ss") & ",Form:" & .Name & ",Object:" & Screen.ActiveControl.Name & ", Routine: " & sSect & ", "; dbgmsg
    End With
    ' send out a small header

    Close #DBGf

Leave:
Exit Sub

ErrorHandler:
  
  MsgBox "Warning: Cannot save debug information"
  Resume Leave

End Sub

