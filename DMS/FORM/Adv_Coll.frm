VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rAdv_Coll 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection Reports"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   DrawWidth       =   2
   Icon            =   "Adv_Coll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   6735
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
         Height          =   420
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   810
         Width           =   1080
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
         Height          =   420
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   810
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker stDT_TM 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Delevary Time"
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   62455810
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin MSComCtl2.DTPicker stDt 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   450
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         _Version        =   393216
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   62455811
         CurrentDate     =   37306
      End
      Begin MSComCtl2.DTPicker edDT_TM 
         Height          =   285
         Left            =   4020
         TabIndex        =   5
         ToolTipText     =   "Delevary Time"
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   62455810
         UpDown          =   -1  'True
         CurrentDate     =   37163.9993055556
      End
      Begin MSComCtl2.DTPicker edDt 
         Height          =   285
         Left            =   3060
         TabIndex        =   4
         Top             =   450
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         _Version        =   393216
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   62455811
         CurrentDate     =   37337.9993055556
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
         Left            =   270
         TabIndex        =   19
         Top             =   150
         Width           =   2400
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
         Left            =   3060
         TabIndex        =   18
         Top             =   120
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Due Patients Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   870
         TabIndex        =   15
         Top             =   1350
         Width           =   2595
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Due Collection Information (Specific)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   870
         TabIndex        =   14
         Top             =   825
         Width           =   3525
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance Collection Information (Specific)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   870
         TabIndex        =   0
         Top             =   240
         Width           =   4125
      End
      Begin VB.TextBox txtU_ID 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   960
         TabIndex        =   1
         Top             =   1770
         Width           =   1455
      End
      Begin VB.TextBox txtU_Name 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1770
         Width           =   3465
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance Collection Information (All)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   870
         TabIndex        =   12
         Top             =   525
         Width           =   3525
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Due Collection Information All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   870
         TabIndex        =   11
         Top             =   1080
         Width           =   2925
      End
      Begin VB.CommandButton cmdDaily_State 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Daily Statement"
         Height          =   315
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1305
      End
      Begin VB.CommandButton cmdProcess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1350
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1740
         Width           =   675
      End
   End
End
Attribute VB_Name = "rAdv_Coll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrStdt As String
Dim StrSttime As String
Dim StDate_TM As String
Dim StDate_TM1 As String
Dim StrEddt As String
Dim StrEdtime As String
Dim EdDate_TM As String
Dim EdDate_TM1 As String



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDaily_State_Click()
    strScr_Name = "rDaily_Statement"
        Authority
   If strAllow = "YES" Then
        rDaily_Statement.Show vbModal
   End If
End Sub

Private Sub CmdPreview_Click()
    If Option1.value = True Then
        If Me.txtU_ID = "" Then
            MsgBox "User ID Required"
            Me.txtU_ID.SetFocus
            Exit Sub
        End If
        CRViewer1_MODE = 31
        Viewer1.Show vbModal
    End If
    
    If Option4.value = True Then
        CRViewer1_MODE = 41
        Viewer1.Show vbModal
    End If
    
    If Option2.value = True Then
        If Me.txtU_ID = "" Then
            MsgBox "User ID Required"
            txtU_ID.SetFocus
            Exit Sub
        End If
        CRViewer1_MODE = 32
        Viewer1.Show vbModal
    End If
    
    If Option3.value = True Then
    
        CRViewer1_MODE = 33
        Viewer1.Show vbModal
        
    End If
    
   If Option5.value = True Then
        CRViewer1_MODE = 42
        Viewer1.Show vbModal
   End If

    
End Sub

Private Sub cmdProcess_Click()

Call Process_Due

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
    stDt = Date
'    stDT_TM = Now
    edDt = Date
'    edDT_TM = Now
End Sub
Private Sub Search_User_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '5','" + Me.txtU_ID + "'", con
    If My_Rst.EOF = False Then
        
        txtU_Name.text = My_Rst!U_Name
    Else
        MsgBox "Invalid ID, Try arain........"
        txtU_Name.text = ""
        txtU_ID.SetFocus
    End If
    
    con.Close

End Sub



Private Sub Option1_Click()
cmdProcess.Visible = False

End Sub

Private Sub Option2_Click()
cmdProcess.Visible = False
End Sub

Private Sub Option3_Click()
cmdProcess.Visible = True
End Sub

Private Sub Option4_Click()
cmdProcess.Visible = False
End Sub

Private Sub Option5_Click()
    txtU_ID = ""
    txtU_Name = ""
    cmdProcess.Visible = False
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

Private Sub txtU_ID_LostFocus()
    If Trim(txtU_ID.text) = "" Then Exit Sub
    Search_User_Name
End Sub

Public Sub Process_Due()
    
   
             StrStdt = Trim(Format(rAdv_Coll.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime
             StDate_TM1 = StrStdt + Space(3) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
             
             StrEddt = Trim(Format(rAdv_Coll.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rAdv_Coll.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++


    
    con.connectionstring = strcn.Connection
    con.ConnectionTimeout = 0
    con.Open
    cmd.CommandTimeout = 0
    
    Set cmd.ActiveConnection = con
    'cmd.CommandText = "exec Rpt_doc_pay2 1,'" & txtRefer_Code & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") & "','" & Format(strEd_date, "yyyy-mm-dd hh:mm AMPM") & "'"
    cmd.CommandText = "exec due_pat '" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'"
     
    Set rs = cmd.Execute
    MsgBox rs!MSG
    con.Close
        
End Sub

