VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLeave_as_Cash 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   DrawWidth       =   2
   Icon            =   "Leave_as_Cash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2505
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
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
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2505
      Width           =   1050
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New"
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
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2505
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2505
      Width           =   1050
   End
   Begin VB.TextBox txtEmp_Name 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3810
   End
   Begin VB.TextBox txtEmp_ID 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   510
      TabIndex        =   0
      Top             =   1050
      Width           =   1245
   End
   Begin VB.ComboBox comLeave_Type 
      Height          =   315
      Left            =   510
      TabIndex        =   2
      Top             =   1740
      Width           =   1995
   End
   Begin VB.TextBox nbrLeave 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2730
      TabIndex        =   3
      Top             =   1800
      Width           =   1110
   End
   Begin MSComCtl2.DTPicker Cash_date 
      Height          =   285
      Left            =   4470
      TabIndex        =   4
      Top             =   1740
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   37367
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   300
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave as Cash"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   510
      TabIndex        =   14
      Top             =   210
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   510
      TabIndex        =   13
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1980
      TabIndex        =   12
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Type"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   510
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave For Cash"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2700
      TabIndex        =   10
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4470
      TabIndex        =   9
      Top             =   1440
      Width           =   750
   End
End
Attribute VB_Name = "frmLeave_as_Cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Search_Leave_as_cash 1,'" & Trim(txtEmp_ID.Text) & "','" & Trim(comLeave_Type) & "','" & Format(Cash_date, "yyyy-mm-dd") & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Leave_Trans
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Leave_Trans
        MsgBox "inserted successfully"
    End If
        
    txtEmp_ID.Text = ""
    txtEmp_Name.Text = ""
    comLeave_Type.Text = ""
    nbrLeave = ""
    Cash_date.value = Date
    txtEmp_ID.SetFocus
    
End Sub
Private Sub Leave_Trans()
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Leave_As_Cash_IUD '" + Mode + "','" + txtEmp_ID.Text + _
    "','" + comLeave_Type.Text + _
    "'," + nbrLeave + _
    ",'" + Format(Cash_date, "yyyy-mm-dd") + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
    
End Sub

Private Sub Search_Leave_Type() 'to leave type

    'comLeave_Type.AddItem
    'comLeave_Type.Refresh
    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Leave_Type 1", con
    'My_Rst.Open "exec pro_name_select '7',''", con
    If My_Rst.EOF = False Then
    Do Until My_Rst.EOF
    comLeave_Type.AddItem My_Rst!Leave_Type
    My_Rst.MoveNext
    Loop
    End If
    con.Close
    
    'comLeave_Type.Refresh
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    Cash_date.value = Date
    Search_Leave_Type
End Sub
Private Sub Search_Emp_Info() 'search EMPLOYEE NAME
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '9','" + txtEmp_ID.Text + "'", con
    If My_Rst.EOF = False Then
        txtEmp_Name.Text = My_Rst!Emp_Name
    Else
        MsgBox "Invalid Employee ID, Try again...."
        txtEmp_ID = ""
        txtEmp_Name.Text = ""
        txtEmp_ID.SetFocus
    End If
    con.Close
End Sub
Private Sub txtEmp_ID_LostFocus()
    If txtEmp_ID = "" Then Exit Sub
    Search_Emp_Info
End Sub
