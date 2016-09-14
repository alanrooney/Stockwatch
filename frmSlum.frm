VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmSlum 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCount 
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
      Height          =   375
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3630
      Width           =   855
   End
   Begin VB.TextBox txtName 
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
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   3435
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   3435
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   960
      TabIndex        =   1
      Top             =   2940
      Width           =   3465
   End
   Begin MyCommandButton.MyButton btnSave 
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4290
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
      Caption         =   "&Save"
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
      Height          =   495
      Left            =   2610
      TabIndex        =   9
      Top             =   4290
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
   Begin MyCommandButton.MyButton btnZero 
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Width           =   1740
      _ExtentX        =   3069
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
      Caption         =   "&Zero Audit Count"
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
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   270
      TabIndex        =   10
      Top             =   5010
      Width           =   2415
   End
   Begin VB.Label lblCount 
      Caption         =   "New Audit Count"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   3690
      Width           =   2925
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
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
      Left            =   150
      TabIndex        =   6
      Top             =   660
      Width           =   645
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
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
      Left            =   150
      TabIndex        =   5
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label lblPhone 
      Caption         =   "Phone"
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
      Left            =   330
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Stockwatch  License Utility Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   30
      Width           =   3855
   End
End
Attribute VB_Name = "frmSlum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnQuit_Click()

    Unload Me

End Sub

Private Sub btnSave_Click()
Dim iErr As Long

    
    lblStatus.Caption = "Writing License Dongle..."
    ' Show Status
    
    If DonglePresent(iErr) Then
    ' Check for Dongle
    
        If OpenDongle(iErr) Then
        ' login to dongle
    
            If WriteName(iErr) Then
            ' save name
        
                If WriteAddress(iErr) Then
                ' save address
    
                    If WritePhone(iErr) Then
                    'save phone
            
                        If WriteAuditCount(iErr) Then
                        ' Read audit count
                        
                            If CloseDongle(iErr) Then
                        
                                lblStatus.Caption = "Write to Dongle Successful"
                                ' Show Status
                        
                                    MsgBox "License Dongle Updated Successfully"
                        
                            Else
                                MsgBox "Error " & Trim$(iErr) & " Cannot Close Dongle"
                            End If
                            
                        Else
                            MsgBox "Error " & Trim$(iErr) & " Cannot Save Audit Count to Dongle"
                        End If
                
                    Else
                        MsgBox "Error " & Trim$(iErr) & " Cannot Save Phone to Dongle"
                    End If
                Else
                    MsgBox "Error " & Trim$(iErr) & " Cannot Save Address to Dongle"
                End If
            Else
                MsgBox "Error " & Trim$(iErr) & " Cannot Save Name to Dongle"
            End If
        
        Else
            MsgBox "Error " & Trim$(iErr) & " Cannot Open/Log Into Dongle"
        End If
    Else
        MsgBox "Error " & Trim$(iErr) & " Cannot Find Dongle"
    End If
    


End Sub

Private Sub btnZero_Click()

    txtCount.Text = "0"

    ' Call Zero function and reshow details
    

End Sub

Private Sub Form_Activate()
Dim iErr As Long

    DoEvents
    
    lblStatus.Caption = "Reading License Dongle..."
    ' Show Status

    If DonglePresent(iErr) Then
    ' Check for Dongle
    
        If OpenDongle(iErr) Then
        ' login to dongle
    
            If ReadName(iErr) Then
            ' Read Name
    
                If ReadAddress(iErr) Then
                ' Read Address
                
                    If ReadPhone(iErr) Then
                    ' Read Phone
                
                        If ReadAuditCount(iErr) Then
                        ' Read audit count
    
                            ' read audit reminder count

                            ' read audit max count
        
                            
                            If CloseDongle(iErr) Then
                                
                                lblStatus.Caption = "Dongle Read Successfully..."
                                ' Show Status

                            
                                If txtName.Text = "" Then
                                    bSetFocus Me, "txtName"
                                Else
                                    bSetFocus Me, "btnQuit"
                                End If
                            Else
                                MsgBox "Error " & Trim$(iErr) & " Cannot Close Dongle"
                            End If
                            
                
                        Else
                            MsgBox "Error " & Trim$(iErr) & " Cannot read Audit Count"
                        End If
            
                    Else
                        MsgBox "Error " & Trim$(iErr) & " Cannot read Phone"
                    End If
                Else
                    MsgBox "Error " & Trim$(iErr) & " Cannot read Address"
                End If
            Else
                MsgBox "Error " & Trim$(iErr) & " Cannot read Name"
            End If
        
        Else
            MsgBox "Error " & Trim$(iErr) & " Cannot Open/Log Into Dongle"
        End If
    Else
        MsgBox "Error " & Trim$(iErr) & " Cannot Find Dongle"
    End If
    
    
    ' focus name


End Sub

Private Sub Form_Load()

    ' Set small size for password


End Sub

Private Sub txtPass_Change()

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    ' on enter check password with whats on dongle
    
    ' if bad warn
    
    ' if good open up panel
    
    '   show name, address, phone no
    
    '   Show audits since last zeroed
    
    ' default to quit button
    

End Sub

Private Sub txtCount_DblClick()

    txtCount.Locked = False

End Sub

Public Function DonglePresent(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

'find dongle
retcode = UniKey_Find(handle, lp1, lp2)
If (retcode = 0) Then
    DonglePresent = True
End If

End Function

Public Function OpenDongle(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

retcode = UniKey_User_Logon(handle, p1, p2)
If (retcode = 0) Then
    OpenDongle = True
End If

End Function

Public Function ReadName(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    'read dongle memory
    ' NAME
    p1 = 1
    p2 = 100
    
    buffer = String(255, " ")
'    buffer = "                          "
    retcode = UniKey_Read_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
    
        str = StrConv(buffer(0), vbUnicode)
    
        i = 0
        Do
            sOut = sOut & Chr$(buffer(i))
            i = i + 1
    
        Loop While buffer(i) <> 0
    
        txtName.Text = sOut
    
        ReadName = True
    End If
    


End Function

Public Function ReadAddress(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    'read dongle memory
    ' ADDRESS
    p1 = 101
    p2 = 200
    
    buffer = String(255, " ")
'    buffer = "                          "
    retcode = UniKey_Read_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
    
        str = StrConv(buffer(0), vbUnicode)
    
        i = 0
        Do
            sOut = sOut & Chr$(buffer(i))
            i = i + 1
    
        Loop While buffer(i) <> 0
    
        txtAddress.Text = sOut
    
        ReadAddress = True
    End If
    


End Function

Public Function ReadPhone(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    'read dongle memory
    ' PHONE
    p1 = 302
    p2 = 20
    
    buffer = String(255, " ")
'    buffer = "                          "
    retcode = UniKey_Read_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
    
        str = StrConv(buffer(0), vbUnicode)
    
        i = 0
        Do
            sOut = sOut & Chr$(buffer(i))
            i = i + 1
    
        Loop While buffer(i) <> 0
    
        txtPhone.Text = sOut
    
        ReadPhone = True
    End If
    


End Function

Public Function ReadAuditCount(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)
    
    
    'AUDIT COUNT
    sOut = ""
    i = 0
    
    p1 = 330
    p2 = 5
    
    buffer = String(255, " ")
    
    retcode = UniKey_Read_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
        str = StrConv(buffer(0), vbUnicode)
    
        i = 0
        Do
            sOut = sOut & Chr$(buffer(i))
            i = i + 1
    
        Loop While buffer(i) <> 0
    
        txtCount.Text = sOut
    
        ReadAuditCount = True
    
    End If
    
End Function

Public Function CloseDongle(retcode As Long)
Dim handle As Integer
    
    
    ' close dongle
    retcode = UniKey_Logoff(handle)
    If (retcode = 0) Then
        CloseDongle = True
    End If

End Function
Public Function WriteName(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    ' NAME
    p1 = 1
    p2 = 100
    str = Trim$(txtName.Text)
    buffer = str
    tmp = Uni2Ac(buffer, LenB(str))
    str = StrConv(buffer, vbUnicode)
    'write memory
    retcode = UniKey_Write_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
        WriteName = True
    End If
    
End Function

Public Function WriteAddress(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    ' ADDRESS
    p1 = 101
    p2 = 200
    str = Trim$(txtAddress.Text)
    buffer = str
    tmp = Uni2Ac(buffer, LenB(str))
    str = StrConv(buffer, vbUnicode)
    'write memory
    retcode = UniKey_Write_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
        WriteAddress = True
    End If

End Function

Public Function WritePhone(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    ' Phone
    p1 = 302
    p2 = 20
    str = Trim$(txtPhone.Text)
    buffer = str
    tmp = Uni2Ac(buffer, LenB(str))
    str = StrConv(buffer, vbUnicode)
    'write memory
    retcode = UniKey_Write_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
        WritePhone = True
    End If
End Function

Public Function WriteAuditCount(retcode As Long)
Dim handle As Integer
Dim p1, p2, p3, p4, i, j
Dim lp1, lp2, v As Long
Dim buffer() As Byte
Dim rc(0 To 3) As Long
Dim curline As Integer
Dim str As String
Dim tmp As Integer
Dim sOut As String

p1 = 1234       'passwords
p2 = 1234
p3 = 1234
p4 = 1234

buffer = Space(4096)

    ' Audit Count
    p1 = 330
    p2 = 5
    str = Trim$(txtCount.Text)
    buffer = str
    tmp = Uni2Ac(buffer, LenB(str))
    str = StrConv(buffer, vbUnicode)
    'write memory
    retcode = UniKey_Write_Memory(handle, p1, p2, buffer(0))
    If (retcode = 0) Then
        WriteAuditCount = True
    End If

End Function

