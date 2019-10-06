VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rLeave_Balance 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   DrawWidth       =   2
   Icon            =   "rLeave_Balance.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   330
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1455
      Width           =   960
   End
   Begin VB.ComboBox comEmp_Name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   3435
   End
   Begin VB.TextBox txtEmp_ID 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2445
      TabIndex        =   3
      Top             =   645
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre&view"
      Height          =   330
      Left            =   2895
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1455
      Width           =   960
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   555
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   1785
   End
End
Attribute VB_Name = "rLeave_Balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdPreview_Click()
    If Trim(comEmp_Name) = "" Then Exit Sub

    CRViewer1_MODE = 37 'TO USE VIEWER1 FOR Employee Leave Info
    Viewer1.Show vbModal
End Sub
Private Sub comEmp_Name_Click()
    txtEmp_ID.Text = ""
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_select '13','" & Trim(comEmp_Name.Text) & "'"
    Adodc1.Refresh
     
    If Adodc1.Recordset.RecordCount > 0 Then
       txtEmp_ID = Adodc1.Recordset!emp_id
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub
Private Sub Form_Load()
   
    Show_Emp_Name
End Sub
Private Sub Show_Emp_Name()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Leave_Type 2", con
    If My_Rst.EOF = False Then
    Do Until My_Rst.EOF
    comEmp_Name.AddItem My_Rst!Emp_Name

    My_Rst.MoveNext
    Loop
    End If
    con.Close
    
End Sub
Private Sub Option1_Click()
'    If Option1.value = True Then
'        comEmp_Name.Enabled = False
'    End If
End Sub
'Private Sub Option2_Click()
'    If Option2.value = True Then
'        comEmp_Name.Enabled = True
'    End If
'End Sub

