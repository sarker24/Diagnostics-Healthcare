VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rBooth_User_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "rBooth_Usr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBooth 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1380
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   330
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific"
      Height          =   195
      Left            =   4020
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All"
      Height          =   195
      Left            =   3450
      TabIndex        =   12
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtU_Name 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   300
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   405
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1830
      Width           =   840
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
      Height          =   405
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1830
      Width           =   960
   End
   Begin VB.TextBox txtUID 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Delevary Time"
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   64487426
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1380
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   64487427
      CurrentDate     =   37114
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      ToolTipText     =   "Delevary Time"
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   64487426
      UpDown          =   -1  'True
      CurrentDate     =   37163.9993055556
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   2700
      TabIndex        =   3
      Top             =   1380
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   64487427
      CurrentDate     =   37114
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booth No."
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
      Left            =   270
      TabIndex        =   14
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booth wise Collection"
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
      Left            =   270
      TabIndex        =   10
      Top             =   -30
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2700
      TabIndex        =   9
      Top             =   1110
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1110
      Width           =   1140
   End
End
Attribute VB_Name = "rBooth_User_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPreview_Click()
 '   If Trim(txtUID.Text) = "" Then Exit Sub
     If Trim(txtBooth) = "" Then
        MsgBox "Booth No Required"
        txtBooth.SetFocus
        Exit Sub
    End If
    
    CRViewer1_MODE = 18
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
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
    
    If KeyAscii = 27 Then
    Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    stDt.value = Now
'    stDT_TM.value = Now
    edDt.value = Now
'    edDT_TM.value = Now
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label2.ForeColor = &HFFFFFF
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label2.ForeColor = &HFF0000
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

Private Sub txtUID_LostFocus()
    If Trim(txtUID.text) = "" Then Exit Sub
    U_Name
End Sub
Private Sub U_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set My_Rst.ActiveConnection = con
    My_Rst.Open "exec Doc_SELECT 2,'" & txtUID & "'", con
    If My_Rst.EOF = False Then
    txtU_Name = My_Rst!U_Name
    
    Else
    txtU_Name = ""
    
    End If
    My_Rst.Close
    con.Close
End Sub
