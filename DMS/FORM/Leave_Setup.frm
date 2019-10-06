VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLeave_Setup 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Leave Setup"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Leave_Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox comLeave_Type 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Leave_Setup.frx":000C
         Left            =   1740
         List            =   "Leave_Setup.frx":000E
         TabIndex        =   8
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox nbrCelling 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1740
         TabIndex        =   7
         Top             =   1170
         Width           =   630
      End
      Begin MSComCtl2.DTPicker St_Year 
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   57737219
         UpDown          =   -1  'True
         CurrentDate     =   37257
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2940
         Top             =   1170
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   810
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1170
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Starting Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   390
         Width           =   1410
      End
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
      Height          =   300
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   900
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
      Height          =   300
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   900
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
      Height          =   300
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   900
   End
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
      Height          =   300
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   60
   End
End
Attribute VB_Name = "frmLeave_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
If Me.comLeave_Type = "" Then Exit Sub

    Mode = "D"
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        Leave_Setup
    Me.comLeave_Type = ""
    nbrCelling = ""
    End If
End Sub

Private Sub cmdNew_Click()
    Me.comLeave_Type = ""
    Me.nbrCelling = ""
    Search_Leave_Type
End Sub

Private Sub cmdSave_Click()

    If comLeave_Type = "" Then
        MsgBox "Leave Type Mandatory"
        comLeave_Type.SetFocus
        Exit Sub
    End If

    If Trim(nbrCelling) = "" Or Trim(nbrCelling) = "0" Then
        MsgBox "Leave mandatory"
        nbrCelling.SetFocus
        Exit Sub
    End If
    
    If comLeave_Type = "" Then
        MsgBox "Leave Mandatory"
        comLeave_Type.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '10','" & Trim(comLeave_Type.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Leave_Setup
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Leave_Setup
        MsgBox "inserted successfully"
    End If
    
    comLeave_Type = ""
    nbrCelling = ""
    comLeave_Type.SetFocus
    
    'Search_Leave_Type
End Sub
Private Sub Leave_Setup()
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_Leave_setup_IUD '" + Mode + "','" + Trim(comLeave_Type.Text) + _
    "'," + nbrCelling.Text + ",'" + Format(St_Year.value, "yyyy/mm/dd") + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
    
End Sub

Private Sub comLeave_Type_LostFocus()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '10','" + comLeave_Type + "'", con
    If My_Rst.EOF = False Then
        nbrCelling = My_Rst!Celing
    Else
        nbrCelling = ""
    End If
    
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
    'St_Year = My_Rst!lv_st_year
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
    
    If KeyAscii = 27 Then
    Unload Me
    End If
    
End Sub
Private Sub Form_Load()
   Search_Leave_Type
End Sub
