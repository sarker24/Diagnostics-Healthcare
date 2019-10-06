VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogIn 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2685
   ClientLeft      =   3315
   ClientTop       =   3885
   ClientWidth     =   5385
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   2685
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   405
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtTime 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Unique Diagnostic Center"
      Top             =   600
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Left            =   4380
      Top             =   2160
   End
   Begin VB.TextBox Txtpass 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "01712560276"
      Top             =   1665
      Width           =   2115
   End
   Begin VB.TextBox Txtuserid 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2235
      TabIndex        =   0
      Text            =   "md"
      Top             =   1200
      Width           =   2115
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   3240
      Picture         =   "frmLogin.frx":2A76
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdEnter 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   2280
      Picture         =   "frmLogin.frx":3340
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2160
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker ExpiryDate 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57737217
      CurrentDate     =   41037
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   1215
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim password     As String
Dim Passlen      As Integer
Dim Passtot      As Integer
Dim Passnum      As Double
Dim Finalpass    As String
Dim Flen         As Integer
Dim mssg         As String
Dim emp As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub CmdEnter_Click()
On Error Resume Next
If CmdEnter.Caption = "" Then
    If Txtuserid = "" And Txtpass = "" Then
        Unload Me
        Exit Sub
    End If
'----------------------------------------------
    If Txtuserid = Empty Or Txtpass = Empty Then
        mssg = MsgBox("Please Enter User ID and Password.", vbOKOnly + vbExclamation, "Confirmation")
    Else
        password = LTrim(RTrim(Txtpass))
        Passlen = Len(password)
        Passtot = 0
    
        Select Case Passlen
        Case 1
            Passtot = Passtot + Asc(password)
            Passnum = Val(Mid("12345678901234123456789123456789", 15, 9)) + Passtot
            Finalpass = "12345678901234" + LTrim(RTrim(CStr(Passnum))) + "123456789"
        Case 2
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234567812345", 9, 5)) + Passtot
            Finalpass = "12345678" + LTrim(RTrim(CStr(Passnum)))
        Case 3
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("12312123456", 3)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "12123456"
        Case 4
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456112345", 4, 6)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "112345"
        Case 5
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456781234567890", 4, 8)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "1234567890"
        Case 6
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123456123456789", 6)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456789"
        Case 7
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234561212345", 9, 5)) + Passtot
            Finalpass = "12345612" + LTrim(RTrim(CStr(Passnum)))
        Case 8
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("12345612345123456", 7, 5)) + Passtot
            Finalpass = "123456" + LTrim(RTrim(CStr(Passnum))) + "123456"
        Case 9
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123456781231234561234", 12, 6)) + Passtot
            Finalpass = "12345678123" + LTrim(RTrim(CStr(Passnum))) + "1234"
        Case 10
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("12345678912312345678", 13, 8)) + Passtot
            Finalpass = "123456789123" + LTrim(RTrim(CStr(Passnum)))
        Case 11
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123456123456712345", 6)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456712345"
        Case 12
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234512345612345", 12, 5)) + Passtot
            Finalpass = "12345123456" + LTrim(RTrim(CStr(Passnum)))
        Case 13
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("12345123456123456", 5)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456123456"
        Case 14
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123451234", 5)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "1234"
        Case 15
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456789012312", 4, 10)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "12312"
        End Select
        
        Adodc1.connectionstring = strcn.Connection
        'emp = "select user_pass from micropass where emp_id='" + RTrim(Txtuserid) + "' and cancel='0'"
        emp = "select user_pass from micropass where u_id='" + RTrim(Txtuserid) + "' and cancel='0'"
        Adodc1.RecordSource = emp
        Adodc1.Refresh
        If Adodc1.Recordset.EOF = False Then
            If IsNull(Adodc1.Recordset!user_pass) Then
                mssg = MsgBox("Password not found, please call <Database administrator>.", _
                vbOKOnly + vbExclamation, "Confirmation")
                If mssg = vbOK Then
                End If
            Else
                If RTrim(Adodc1.Recordset!user_pass) = RTrim(Finalpass) Then
                    u_id = LTrim(RTrim(Txtuserid))
                    upass = RTrim(Finalpass)
                    Unload Me
                    frmMAIN.Show vbModal
''                    mnuHide (0)
                Else
                   
                    MsgBox "Invalid password, please try again.", vbOKOnly + vbExclamation, "Confirmation"
                    Txtpass = ""
                    Txtuserid.SetFocus
                End If
            End If
        Else
            
           MsgBox "Invalid employee ID and Password.", vbOKOnly + vbExclamation, "Confirmation"
           Txtuserid = ""
           Txtpass = ""
           Txtuserid.SetFocus
        End If
    End If
    
    Locate_Booth 'for BOOTH STATUS
    'MsgBox BoothN
'    MsgBox u_id
End If

End Sub


Private Sub Form_Load()
Locate_Booth

Timer1.Enabled = True
Timer1.Interval = 1000

Call TimeExpired
End Sub

Private Sub TimeExpired()
Dim CurrentDate As Date
Dim ExpiryDate
CurrentDate = Now
ExpiryDate = "12-31-2015"
If CurrentDate > ExpiryDate Then
'MsgBox ("Database Connection Fail.")
'MsgBox ("Your system has Expired")
'frmLogIn.Hide
Unload Me
End If

End Sub


Private Sub Timer1_Timer()
txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
     txtCTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    If Txtuserid = "" And Txtpass = "" Then
        Unload Me
        Exit Sub
    End If
'----------------------------------------------
    If Txtuserid = Empty Or Txtpass = Empty Then
        mssg = MsgBox("Please Enter User ID and Password.", vbOKOnly + vbExclamation, "Confirmation")
    Else
        password = LTrim(RTrim(Txtpass))
        Passlen = Len(password)
        Passtot = 0
    
        Select Case Passlen
        Case 1
            Passtot = Passtot + Asc(password)
            Passnum = Val(Mid("12345678901234123456789123456789", 15, 9)) + Passtot
            Finalpass = "12345678901234" + LTrim(RTrim(CStr(Passnum))) + "123456789"
        Case 2
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234567812345", 9, 5)) + Passtot
            Finalpass = "12345678" + LTrim(RTrim(CStr(Passnum)))
        Case 3
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("12312123456", 3)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "12123456"
        Case 4
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456112345", 4, 6)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "112345"
        Case 5
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456781234567890", 4, 8)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "1234567890"
        Case 6
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123456123456789", 6)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456789"
        Case 7
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234561212345", 9, 5)) + Passtot
            Finalpass = "12345612" + LTrim(RTrim(CStr(Passnum)))
        Case 8
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("12345612345123456", 7, 5)) + Passtot
            Finalpass = "123456" + LTrim(RTrim(CStr(Passnum))) + "123456"
        Case 9
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123456781231234561234", 12, 6)) + Passtot
            Finalpass = "12345678123" + LTrim(RTrim(CStr(Passnum))) + "1234"
        Case 10
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("12345678912312345678", 13, 8)) + Passtot
            Finalpass = "123456789123" + LTrim(RTrim(CStr(Passnum)))
        Case 11
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123456123456712345", 6)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456712345"
        Case 12
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("1234512345612345", 12, 5)) + Passtot
            Finalpass = "12345123456" + LTrim(RTrim(CStr(Passnum)))
        Case 13
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("12345123456123456", 5)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456123456"
        Case 14
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Left("123451234", 5)) + Passtot
            Finalpass = LTrim(RTrim(CStr(Passnum))) + "1234"
        Case 15
            For Flen = 1 To Passlen
                Passtot = Passtot + Asc(Mid(password, Flen, 1))
            Next
            Passnum = Val(Mid("123123456789012312", 4, 10)) + Passtot
            Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "12312"
        End Select
        
        Adodc1.connectionstring = strcn.Connection
        'emp = "select user_pass from micropass where emp_id='" + RTrim(Txtuserid) + "' and cancel='0'"
        emp = "select user_pass from micropass where u_id='" + RTrim(Txtuserid) + "' and cancel='0'"
        Adodc1.RecordSource = emp
        Adodc1.Refresh
        If Adodc1.Recordset.EOF = False Then
            If IsNull(Adodc1.Recordset!user_pass) Then
                mssg = MsgBox("Password not found, please call <Database administrator>.", _
                vbOKOnly + vbExclamation, "Confirmation")
                If mssg = vbOK Then
                End If
            Else
                If RTrim(Adodc1.Recordset!user_pass) = RTrim(Finalpass) Then
                    u_id = LTrim(RTrim(Txtuserid))
                    upass = RTrim(Finalpass)
                    Unload Me
                    frmMAIN.Show vbModal
''                    mnuHide (0)
                Else
                   
                    MsgBox "Invalid password, please try again.", vbOKOnly + vbExclamation, "Confirmation"
                    Txtpass = ""
                    Txtuserid.SetFocus
                End If
            End If
        Else
            
           MsgBox "Invalid employee ID and Password.", vbOKOnly + vbExclamation, "Confirmation"
           Txtuserid = ""
           Txtpass = ""
           Txtuserid.SetFocus
        End If
    End If
    
    Locate_Booth 'for BOOTH STATUS
    'MsgBox BoothN
'    MsgBox u_id
End If
End Sub

Private Sub Txtuserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Txtpass.SetFocus
End If
End Sub


