VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rStock_Status 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   DrawWidth       =   2
   Icon            =   "rStock_Status.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Stock Present Balance"
      Height          =   225
      Left            =   1950
      TabIndex        =   10
      Top             =   600
      Width           =   2085
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Stock Status"
      Height          =   225
      Left            =   450
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton CmdPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1110
      Width           =   1260
   End
   Begin VB.ComboBox comItem_Name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.TextBox txtItem_Code 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1650
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1050
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   37306
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   2940
      TabIndex        =   1
      Top             =   1050
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   37337
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Status"
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
      Height          =   330
      Left            =   390
      TabIndex        =   8
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2940
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "rStock_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPreview_Click()
    'If Option1.value = True Then
    If Option1.value = True Then
        CRViewer1_MODE = 38
        Viewer1.Show vbModal
    End If
    
    If Option2.value = True Then
        CRViewer1_MODE = 48
        Viewer1.Show vbModal
    End If
    
    'End If
    'If Option2.value = True Then
    'CRViewer1_MODE = 32
    'Viewer.Show VBMODAL
    'End If
    
End Sub

Private Sub comItem_Name_Click()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '15','" + comItem_Name + "'", con
    If My_Rst.EOF = False Then
        txtItem_Code.Text = My_Rst!item_code
    End If
    con.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
    
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Show_Item_Name()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Leave_Type 4", con
    
    If My_Rst.EOF = False Then
    Do Until My_Rst.EOF
    comItem_Name.AddItem My_Rst!item_name
    My_Rst.MoveNext
    Loop
    End If
    con.Close

End Sub

Private Sub Form_Load()
    stDt = Date
    edDt = Date
    Show_Item_Name
End Sub

Private Sub Option1_Click()
'    If Option1.value = True Then
        'Me.comItem_Name = ""
 '       Me.comItem_Name.Enabled = False
 '   End If
    
End Sub

Private Sub Option2_Click()
    'If Option2.value = True Then
        
    '    comItem_Name.Enabled = True
    '    End If
End Sub
