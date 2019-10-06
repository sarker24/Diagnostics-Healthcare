VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rDoc_Due_Pat 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   DrawWidth       =   2
   Icon            =   "rDoc_Due_Pat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   330
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1710
      Width           =   930
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre&view"
      Height          =   330
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1710
      Width           =   960
   End
   Begin VB.TextBox txtRefer_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1320
      Width           =   1125
   End
   Begin VB.TextBox txtDoc_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Delevary Time"
      Top             =   870
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   67633154
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   870
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   67633155
      CurrentDate     =   37306
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   5310
      TabIndex        =   3
      ToolTipText     =   "Delevary Time"
      Top             =   870
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   67633154
      UpDown          =   -1  'True
      CurrentDate     =   37163.9999884259
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   4350
      TabIndex        =   2
      Top             =   870
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   67633155
      CurrentDate     =   37337.9993055556
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Due Patient"
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
      Left            =   330
      TabIndex        =   9
      Top             =   240
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's ID"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   330
      TabIndex        =   8
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day between"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   900
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      TabIndex        =   6
      Top             =   930
      Width           =   195
   End
End
Attribute VB_Name = "rDoc_Due_Pat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPreview_Click()
    If txtRefer_Code = "" Then Exit Sub
        
    CRViewer1_MODE = 40
    Viewer1.Show vbModal
End Sub
Private Sub Search_Doc_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '6','" + Me.txtRefer_Code + "'", con
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

Private Sub txtRefer_Code_LostFocus()
    If Trim(txtRefer_Code) = "" Then Exit Sub
    Search_Doc_Name
End Sub
