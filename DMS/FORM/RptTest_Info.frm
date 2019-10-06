VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RptTest_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "RptTest_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific"
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
      Height          =   240
      Left            =   1725
      TabIndex        =   1
      Top             =   945
      Width           =   1050
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All"
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
      Height          =   240
      Left            =   1725
      TabIndex        =   0
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   330
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   960
   End
   Begin VB.ComboBox comM_Name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1725
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1305
      Width           =   3075
   End
   Begin VB.TextBox txtM_Code 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2895
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre&view"
      Height          =   330
      Left            =   2895
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   960
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2640
      Top             =   270
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   5355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   510
      TabIndex        =   7
      Top             =   240
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   555
      TabIndex        =   6
      Top             =   1350
      Width           =   1095
   End
End
Attribute VB_Name = "RptTest_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdPreview_Click()
'    vew_ven.Show VBMODAL
    CRViewer1_MODE = 1 'TO USE VIEWER1 FOR RptTest_Info
    Viewer.Show vbModal
End Sub
'Private Sub comCust_Name_Click()
'    txtM_Code.Text = ""
'    Dim St As String
'    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "select m_code from test_info_main where m_name='" & Trim(comM_Name.Text) & "'"
'    Adodc1.Refresh
'
'    If Adodc1.Recordset.RecordCount > 0 Then
'       txtM_Code = Adodc1.Recordset!m_code
'    End If
'End Sub
Private Sub comM_Name_Click()
    txtM_Code.text = ""
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select m_code from test_info_main where m_name='" & Trim(comM_Name.text) & "'"
    Adodc1.Refresh
        
    If Adodc1.Recordset.RecordCount > 0 Then
       txtM_Code = Adodc1.Recordset!m_code
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
    
    Show_Ven
End Sub
Private Sub Show_Ven()
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select m_name from test_info_main order by m_code"
    Adodc1.Refresh
        
    If Adodc1.Recordset.RecordCount > 0 Then
       Do Until Adodc1.Recordset.EOF
          comM_Name.AddItem Adodc1.Recordset!m_name
       Adodc1.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub Option1_Click()
    If Option1.value = True Then
        comM_Name.Enabled = False
    End If
End Sub
Private Sub Option2_Click()
    If Option2.value = True Then
        comM_Name.Enabled = True
    End If
End Sub

