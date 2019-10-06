VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChange_Password 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change User Password"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmChange Password1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Password Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4935
      Begin VB.TextBox Txtnewpass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   660
         Width           =   3120
      End
      Begin VB.TextBox Txtconfpass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1095
         Width           =   3120
      End
      Begin VB.TextBox TxtOldPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   240
         Width           =   3120
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   135
         TabIndex        =   12
         Top             =   675
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   135
         TabIndex        =   11
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      Begin VB.TextBox txtU_ID 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   975
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   375
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   705
         Width           =   585
      End
      Begin VB.Label Lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   975
         TabIndex        =   3
         Top             =   765
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1290
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   240
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5175
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
Dim myrst           As New ADODB.Recordset

Private Sub cmdClose_Click()

Unload Me

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtU_ID = frmCreate_User.Text1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF0000
    Label2.ForeColor = &HFF0000
    Label3.ForeColor = &HFF0000
    Label4.ForeColor = &HFF0000
    Label5.ForeColor = &HFF0000
    Label6.ForeColor = &HFF0000
    
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &H8000&
    Label2.ForeColor = &H8000&
    Label3.ForeColor = &H8000&
    Label4.ForeColor = &H8000&
    Label5.ForeColor = &H8000&
    Label6.ForeColor = &H8000&
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
'        If npass <> "" Then
           If npass <> cpass And npass <> "" Then
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
                cmdclose.SetFocus
          End If
'      End If
    End If
End If
End Sub

Private Sub TxtOldPass_KeyPress(KeyAscii As Integer)
    
If KeyAscii = 13 Then
    If TxtOldPass = Empty Then
        MsgBox "Blank not allowed.", vbOKOnly + vbExclamation, "Confirmation"
    Else
        newlen = Len(LTrim(RTrim(TxtOldPass)))
        newtot = 0
        Select Case newlen
        Case 1
            newtot = newtot + Asc(TxtOldPass)
            npass = "12345678901234" + _
            LTrim(RTrim(CStr(123456789 + newtot))) + "123456789"
        Case 2
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "12345678" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 3
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123 + newtot))) + "12123456"
         Case 4
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(123456 + newtot))) + "112345"
         Case 5
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(12345678 + newtot))) + "1234567890"
         Case 6
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123456 + newtot))) + "123456789"
        Case 7
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "12345612" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 8
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "123456" + LTrim(RTrim(CStr(12345 + newtot))) + "123456"
        Case 9
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "12345678123" + LTrim(RTrim(CStr(123456 + newtot))) + "1234"
        Case 10
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "123456789123" + LTrim(RTrim(CStr(12345678 + newtot)))
        Case 11
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(123456 + newtot))) + "123456712345"
        Case 12
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "12345123456" + LTrim(RTrim(CStr(12345 + newtot)))
        Case 13
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(12345 + newtot))) + "123456123456"
        Case 14
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = LTrim(RTrim(CStr(12345 + newtot))) + "1234"
         Case 15
            For newpass = 1 To newlen
                newtot = newtot + _
                Asc(Mid(LTrim(RTrim(TxtOldPass)), newpass, 1))
            Next
            npass = "123" + LTrim(RTrim(CStr(1234567890 + newtot))) + "12312"
        End Select
        Txtconfpass.Enabled = True
        Txtconfpass.SetFocus
        '-----------------------------
        'Dim st As String
        Adodc1.connectionstring = strcn.Connection
        Adodc1.RecordSource = "select user_pass from micropass where u_id='" & Trim(txtU_ID.Text) & "'"
        
        Adodc1.Refresh
        If IsNull(Adodc1.Recordset.RecordCount) = False Then
            Dim OldPass As String
            OldPass = Adodc1.Recordset!user_pass
        End If
        
        If OldPass <> npass Then
            MsgBox "Password confirmation does not match.", vbOKOnly + vbExclamation, "Confirmatiom"
            'Txtconfpass.Enabled = False
            TxtOldPass = ""
            TxtOldPass.SetFocus
            Exit Sub
        End If
        
        '-----------------------------
        Txtnewpass.SetFocus
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
                   lblName = Adodc1.Recordset!U_Name
               Else
'                    frmUser_List.Show vbModal
               End If

'-------------------
End Sub
Private Sub Txtnewpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

'-------------------------------
        Adodc1.connectionstring = strcn.Connection
        Adodc1.RecordSource = "select user_pass from micropass where u_id='" & Trim(txtU_ID.Text) & "'"
        Adodc1.Refresh
        
        If Adodc1.Recordset.RecordCount > 0 Then
            If IsNull(Adodc1.Recordset!user_pass) = False Then
                If Trim(TxtOldPass.Text) = "" Then
                    MsgBox "First You Have to entry Old Password.", vbOKOnly + vbExclamation, "Confirmation"
                    Txtnewpass = ""
                    TxtOldPass.SetFocus
                    Exit Sub
                End If
            End If
        End If
'---------------------------------

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
