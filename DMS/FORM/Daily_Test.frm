VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rDaily_Test 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   DrawWidth       =   2
   Icon            =   "Daily_Test.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtM_Name 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3135
   End
   Begin VB.TextBox txtM_Code 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   390
      TabIndex        =   1
      Top             =   1350
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   1650
      TabIndex        =   6
      Top             =   870
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   870
      Value           =   -1  'True
      Width           =   525
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
      Height          =   285
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1080
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
      Height          =   285
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1080
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   1290
      TabIndex        =   8
      ToolTipText     =   "Delevary Time"
      Top             =   2040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   63176706
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   330
      TabIndex        =   2
      Top             =   2040
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   63176707
      CurrentDate     =   37306
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "Delevary Time"
      Top             =   2040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   63176706
      UpDown          =   -1  'True
      CurrentDate     =   37163.9993055556
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   63176707
      CurrentDate     =   37337.9993055556
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date and Time"
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
      Left            =   3240
      TabIndex        =   12
      Top             =   1710
      Width           =   2325
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date and Time"
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
      Left            =   360
      TabIndex        =   11
      Top             =   1710
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Statement ( Reagent)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   405
      Left            =   420
      TabIndex        =   10
      Top             =   180
      Width           =   3780
   End
End
Attribute VB_Name = "rDaily_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPreview_Click()


    If Option1.value = True Then
        txtM_Code = ""
        txtM_Name = ""
    End If
       
    If Option2.value = True Then
    
        If Trim(txtM_Code.text) = "" Then
        MsgBox "Test Name mandatory"
        txtM_Code.SetFocus
        Exit Sub
        End If
        
    End If
    
    CRViewer1_MODE = 34
    Viewer.Show vbModal
    
    
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
    If KeyAscii = 27 Then
    Unload Me
    End If
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    stDt = Date
    edDt = Date
    
    Group_Name
    
End Sub
Private Sub Group_Name() 'to search Test Group name

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "select distinct m_name from test_info_main where m_code= '" & Me.txtM_Code & "'", con
    
    If My_Rst.EOF = False Then
    'Do Until My_Rst.EOF
    'comTest_Group.AddItem My_Rst!m_name
    txtM_Name.text = My_Rst!m_name
    'My_Rst.MoveNext
    'Loop
    End If
    con.Close
    
End Sub

Private Sub Option1_Click()
    If Option2.value = True Then
        comTest_Group.Enabled = True
    End If
    
    If Option1.value = True Then
'        comTest_Group = ""
'        comTest_Group.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.value = True Then
'        comTest_Group.Enabled = True
    End If
    
    If Option1.value = True Then
 '       comTest_Group = ""
        comTest_Group.Enabled = False
    End If
End Sub

Private Sub stDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub

Private Sub txtM_Code_LostFocus()
    If Me.txtM_Code = "" Then Exit Sub
    Group_Name
End Sub
