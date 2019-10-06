VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RptConsultant 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   DrawWidth       =   2
   Icon            =   "rPat_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6885
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
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1590
      Width           =   930
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre&view"
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
      TabIndex        =   2
      Top             =   1590
      Width           =   960
   End
   Begin VB.TextBox txtCons_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1350
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   645
   End
   Begin VB.TextBox txtDoc_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   2190
      TabIndex        =   4
      ToolTipText     =   "Delevary Time"
      Top             =   750
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   52297730
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   1230
      TabIndex        =   5
      Top             =   750
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   52297731
      CurrentDate     =   37306
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   5220
      TabIndex        =   6
      ToolTipText     =   "Delevary Time"
      Top             =   750
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   52297730
      UpDown          =   -1  'True
      CurrentDate     =   37163.9999884259
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   4260
      TabIndex        =   7
      Top             =   750
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   52297731
      CurrentDate     =   37337.9993055556
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   1560
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
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
      Caption         =   "Consultant Name"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   1560
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Consultant Performance Statement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consultant ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3960
      TabIndex        =   8
      Top             =   810
      Width           =   240
   End
End
Attribute VB_Name = "RptConsultant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPreview_Click()
'    If txtCons_Code = "" Then Exit Sub
        
    CRViewer1_MODE = 39
    Viewer1.Show vbModal
End Sub
Private Sub Search_Emp_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con

    My_Rst.Open "exec pro_name_SELECT '6','" + Me.txtCons_Code + "'", con
    If My_Rst.EOF = False Then
        'txtItem_Code.Text = My_Rst!item_code
        txtDoc_Name.text = My_Rst!doc_name
    Else
        txtDoc_Name.text = ""
    End If

    con.Close

End Sub

Private Sub edDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub edDT_TM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
'    Search_Emp_Name
    End If
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    edDt = Date
    stDt = Date
'    stDT_TM = Now
'    edDT_TM = Now
End Sub
Private Sub stDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub stDT_TM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub txtCons_Code_GotFocus()
txtCons_Code.BackColor = &HFFFFC0
End Sub

Private Sub txtCons_Code_LostFocus()
'    If Trim(txtRefer_Code) = "" Then Exit Sub
'    Search_Emp_Name
If Trim(txtCons_Code.text) = 0 Then Exit Sub
    
    Dim st As String
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec Pro_FLUSH2 1,'" & Trim(txtCons_Code.text) & "'"
    Adodc2.Refresh
    
    
    If Adodc2.Recordset.RecordCount > 0 Then
      txtDoc_Name.text = Adodc2.Recordset!doc_name
        
        End If
        
    Exit Sub

End Sub
