VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSoftware_Priviliege 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Software privilege"
   ClientHeight    =   5235
   ClientLeft      =   780
   ClientTop       =   1020
   ClientWidth     =   10065
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSoftware privilege.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Index           =   2
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   330
      Index           =   1
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      Height          =   330
      Index           =   4
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4635
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSoftware privilege.frx":302D
      Height          =   2610
      Index           =   1
      Left            =   5130
      TabIndex        =   6
      Top             =   1845
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4604
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16382975
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "scr_no"
         Caption         =   "Screen no."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descript"
         Caption         =   "Discription"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5625.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSoftware privilege.frx":3045
      Height          =   2610
      Index           =   0
      Left            =   3285
      TabIndex        =   4
      Top             =   1845
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4604
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   15923199
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "software"
         Caption         =   "List"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2534.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   0
      Left            =   6840
      Picture         =   "frmSoftware privilege.frx":305D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   955
      Width           =   390
   End
   Begin VB.TextBox Txtemp_id 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4635
      TabIndex        =   0
      Top             =   990
      Width           =   2160
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   270
      Index           =   0
      Left            =   3060
      Top             =   4650
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   476
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
      Appearance      =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   1
      Left            =   4230
      Top             =   4635
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Software Privileges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   3330
      TabIndex        =   12
      Top             =   270
      Width           =   2445
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Privileges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   5895
      TabIndex        =   7
      Top             =   1485
      Width           =   690
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   990
      Width           =   525
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   3330
      TabIndex        =   3
      Top             =   1485
      Width           =   1290
   End
   Begin VB.Label txtname 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   7560
      TabIndex        =   2
      Top             =   990
      Width           =   405
   End
End
Attribute VB_Name = "frmSoftware_Priviliege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim st    As String
''Dim cn    As New ADODB.Connection --
''Dim comm  As New ADODB.Command --
'
'Private Sub Command1_Click(Index As Integer)
'Select Case Index
'Case 0
'    Pflag = "99"
'    Tflag = False
'    Form24.Show (1)
'Case 1
'    Txtemp_id = ""
'    Txtemp_id.Refresh
'Case 2
'    Unload Me
'Case 3
'    con.connectionstring = strcn.Connection
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandText = "exec size_pmt " + CStr(Adodc1(1).Recordset!code)
'    cmd.Execute
'    con.Close
'    Call my_process
'Case 4
'    If Trim(Txtemp_id) = "" Then Exit Sub
'    terget_emp = Txtemp_id
'    soft = Adodc1(0).Recordset!software
'    Form23.Show (1)
'End Select
'End Sub
'
'Private Sub Command1_GotFocus(Index As Integer)
'If Index = 4 Then
'    Call my_process
'End If
'If Index = 0 Then
'    If Tflag = True Then
'        If pick <> "" Then
'            Txtemp_id = pick
'            txtname(1).Caption = terget_emp
'            Tflag = False
'        End If
'    End If
'    pick = " "
'    terget_emp = " "
'End If
'End Sub
'
'Private Sub DataGrid1_DblClick(Index As Integer)
'If Index = 0 Then
'Call my_process
'End If
'End Sub
'
'Private Sub Form_Load()
'Adodc1(0).connectionstring = strcn.Connection
'Adodc1(0).RecordSource = "select software from soft_bag group by software"
'Adodc1(0).Refresh
''--------------------
'Adodc1(1).connectionstring = strcn.Connection
'Adodc1(1).RecordSource = "select scr_no,descript from soft_bag where code=0"
'Adodc1(1).Refresh
'End Sub
'
'Public Sub my_process()
'    If RTrim(Txtemp_id) <> "" Then
'        st = "select scr_no,descript,soft_bag.code from " _
'        & "soft_bag,permit where soft_bag.code = permit.code And permit.emp_id ='" + LTrim(RTrim(Txtemp_id)) + "' and software='" + Adodc1(0).Recordset!software + "'"
'        Adodc1(1).RecordSource = st
'        Adodc1(1).Refresh
'        DataGrid1(1).Refresh
'    Else
'        st = "select scr_no,descript,soft_bag.code from " _
'        & "soft_bag,permit where soft_bag.code = permit.code And permit.emp_id ='' and software='" + Adodc1(0).Recordset!software + "'"
'        Adodc1(1).RecordSource = st
'        Adodc1(1).Refresh
'        DataGrid1(1).Refresh
'        MsgBox "Select an user.", vbOKOnly, "Attention"
'    End If
'    If Adodc1(1).Recordset.RecordCount = 0 Then
'        Command1(3).Enabled = False
'    Else
'        Command1(3).Enabled = True
'    End If
'End Sub
'
'Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'
'End Sub
'
'Private Sub Txtemp_id_Change()
'Call my_process
'End Sub
'
'
'
'
'
