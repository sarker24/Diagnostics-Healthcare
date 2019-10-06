VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChange_Password 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Change Password"
   ClientHeight    =   3615
   ClientLeft      =   3150
   ClientTop       =   1485
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChange Password.frx":0000
   ScaleHeight     =   3615
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   4740
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
   Begin VB.TextBox txtU_ID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   1035
      Width           =   2475
   End
   Begin VB.TextBox Txtnewpass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2190
      Width           =   3060
   End
   Begin VB.TextBox Txtconfpass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2625
      Width           =   3060
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   330
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3150
      Width           =   870
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      Height          =   1005
      Left            =   2070
      Top             =   900
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   1050
      Left            =   2070
      Top             =   2025
      Width           =   4695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   225
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2205
      TabIndex        =   8
      Top             =   1050
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2205
      TabIndex        =   7
      Top             =   1500
      Width           =   405
   End
   Begin VB.Label Lblname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2205
      TabIndex        =   6
      Top             =   2205
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2205
      TabIndex        =   5
      Top             =   2670
      Width           =   1350
   End
End
Attribute VB_Name = "frmChange_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim newlen          As Integer
Dim npass           As String
Dim Conflen         As Integer
Dim cpass           As String
Dim newtot          As Integer
Dim newpass         As Integer
Dim conftot         As Integer
Dim Confpass        As Integer
'Dim con             As New ADODB.Connection --
'Dim cmd             As New ADODB.Command --
Dim myrst           As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
'Pflag = "99"
'Tflag = False
'Form24.Show (1)
End Sub

Private Sub Command2_GotFocus()
'If Tflag = True Then
'    If pick <> "" Then
'        txtU_ID = pick
''        Lblname.Caption = terget_emp
'        Tflag = False
'        Txtnewpass.Enabled = True
'        Txtnewpass.SetFocus
'    End If
'End If
'pick = ""
'terget_emp = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub Txtconfpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txtconfpass = Empty Then
        MsgBox "Information Incomplete.", vbOKOnly + vbExclamation, "Confirmation"
    Else
        Conflen = Len(LTrim(RTrim(Txtconfpass)))
        conftot = 0
        Select Case Conflen
        Case 1
            conftot = conftot + Asc(Txtconfpass)
            cpass = "12345678901234" + _
            LTrim(RTrim(CStr(123456789 + conftot))) + "123456789"
        Case 2
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "12345678" + LTrim(RTrim(CStr(12345 + conftot)))
        Case 3
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = LTrim(RTrim(CStr(123 + conftot))) + "12123456"
         Case 4
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "123" + LTrim(RTrim(CStr(123456 + conftot))) + "112345"
         Case 5
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "123" + LTrim(RTrim(CStr(12345678 + conftot))) + "1234567890"
         Case 6
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = LTrim(RTrim(CStr(123456 + conftot))) + "123456789"
        Case 7
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "12345612" + LTrim(RTrim(CStr(12345 + conftot)))
        Case 8
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "123456" + LTrim(RTrim(CStr(12345 + conftot))) + "123456"
        Case 9
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "12345678123" + LTrim(RTrim(CStr(123456 + conftot))) + "1234"
        Case 10
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "123456789123" + LTrim(RTrim(CStr(12345678 + conftot)))
        Case 11
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = LTrim(RTrim(CStr(123456 + conftot))) + "123456712345"
        Case 12
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "12345123456" + LTrim(RTrim(CStr(12345 + conftot)))
        Case 13
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = LTrim(RTrim(CStr(12345 + conftot))) + "123456123456"
        Case 14
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = LTrim(RTrim(CStr(12345 + conftot))) + "1234"
        Case 15
            For Confpass = 1 To Conflen
                conftot = conftot + _
                Asc(Mid(LTrim(RTrim(Txtconfpass)), Confpass, 1))
            Next
            cpass = "123" + LTrim(RTrim(CStr(1234567890 + conftot))) + "12312"
        End Select
        If npass <> cpass Then
            MsgBox "Password confirmation does not match.", vbOKOnly + vbExclamation, "Confirmatiom"
            Txtconfpass.Enabled = False
        Else
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con
            cmd.CommandText = "exec pro_micropass 0,'" _
            + txtU_ID + "',' ','" + cpass + "','" + u_id + "','2000-01-01',0,'E'"
            cmd.Execute
            con.Close
            Txtnewpass = ""
            Txtconfpass = ""
        End If
    End If
End If
End Sub

Private Sub txtU_ID_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtU_ID_LostFocus()
    User_List_Mode = "frmChange_Pasword_Name"
    If Len(Trim(txtU_ID.Text)) = 0 Then Exit Sub
       
               Adodc1.connectionstring = strcn.Connection
               Adodc1.RecordSource = "exec pro_name_SELECT 5,'" & Trim(txtU_ID.Text) & "'"
               Adodc1.Refresh
        
               If Adodc1.Recordset.RecordCount > 0 Then
                   txtU_ID = Adodc1.Recordset!u_id
                   Lblname = Adodc1.Recordset!u_name
               Else
                    frmUser_List.Show
               End If
    
'-------------------

'    If Adodc4.Recordset.Fields(0) = "N" Then
'        User_List_Mode = "frmChange_Pasword_Name"
'        txtS_Name = ""
'        nbrTest_Rate = 0
'        frmUser_List.Show
'    Else
End Sub

Private Sub Txtemp_id_Change()

End Sub

Private Sub Txtemp_id_LostFocus()

End Sub

Private Sub Txtnewpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Txtnewpass = Empty Then
        MsgBox "Blank not allowed.", vbOKOnly + vbExclamation, "Confirmation"
    Else
        newlen = Len(LTrim(RTrim(Txtnewpass)))
        newtot = 0
        Select Case newlen
        Case 1
            newtot = newtot + Asc(Txtnewpass)
            npass = "12345678901234" + _
            LTrim(RTrim(CStr(123456789 + newtot))) + "123456789"
        Case 2
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "12345678" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 3
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123 + newtot))) + "12123456"
         Case 4
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(123456 + newtot))) + "112345"
         Case 5
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(12345678 + newtot))) + "1234567890"
         Case 6
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123456 + newtot))) + "123456789"
        Case 7
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "12345612" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 8
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "123456" + LTrim(RTrim(CStr(12345 + newtot))) + "123456"
        Case 9
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "12345678123" + LTrim(RTrim(CStr(123456 + newtot))) + "1234"
        Case 10
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "123456789123" + LTrim(RTrim(CStr(12345678 + newtot)))
        Case 11
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123456 + newtot))) + "123456712345"
        Case 12
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "12345123456" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 13
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(12345 + newtot))) + "123456123456"
        Case 14
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(12345 + newtot))) + "1234"
        Case 15
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(Txtnewpass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(1234567890 + newtot))) + "12312"
        End Select
        Txtconfpass.Enabled = True
        Txtconfpass.SetFocus
    End If
End If
End Sub




