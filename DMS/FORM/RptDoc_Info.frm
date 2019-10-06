VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RptDoctor_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "RptDoc_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Preview"
      Height          =   330
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   960
   End
   Begin VB.TextBox txtRefer_Code 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2355
      TabIndex        =   5
      Top             =   915
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox comDoc_Name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1230
      Width           =   3435
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   330
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1725
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1320
      TabIndex        =   0
      Top             =   630
      Width           =   1050
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1320
      TabIndex        =   1
      Top             =   900
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2670
      Top             =   825
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   210
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   1275
      Width           =   945
   End
End
Attribute VB_Name = "RptDoctor_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdPreview_Click()
    CRViewer1_MODE = 2 'TO USE VIEWER1 FOR RptDoc_Info
    Viewer.Show vbModal
End Sub
Private Sub comDoc_Name_Click()
    txtRefer_Code.text = ""
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select refer_code from doctor_info where doc_name='" & Trim(ComDoc_Name.text) & "'"
    Adodc1.Refresh
        
    If Adodc1.Recordset.RecordCount > 0 Then
       txtRefer_Code = Adodc1.Recordset!refer_code
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
    
    Show_Doc_Name
End Sub
Private Sub Show_Doc_Name()
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select doc_name from doctor_info order by refer_code"
    Adodc1.Refresh
        
    If Adodc1.Recordset.RecordCount > 0 Then
       Do Until Adodc1.Recordset.EOF
          ComDoc_Name.AddItem Adodc1.Recordset!doc_name
       Adodc1.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub Option1_Click()
    If Option1.value = True Then
        ComDoc_Name.Enabled = False
    End If
End Sub
Private Sub Option2_Click()
    If Option2.value = True Then
        ComDoc_Name.Enabled = True
    End If
End Sub
