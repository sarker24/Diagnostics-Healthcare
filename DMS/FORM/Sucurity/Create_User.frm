VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCreate_User 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Informations"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "Create_User.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   7215
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Create_User.frx":000C
         Height          =   3225
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12629161
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            MarqueeStyle    =   2
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   7215
      Begin VB.CommandButton cmdSet_Pass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Password"
         Height          =   345
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   420
         Width           =   1680
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   780
         Width           =   5280
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6C4C1&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   5460
         TabIndex        =   4
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   735
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   330
      Index           =   0
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6765
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C&lear"
      Height          =   330
      Index           =   1
      Left            =   5145
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6765
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   330
      Index           =   2
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6765
      Width           =   1005
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   6720
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "2Grid"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2280
      Top             =   6720
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Adodc2"
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
   Begin VB.Label lblUName 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Create New User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmCreate_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim con    As New ADODB.Connection
'Dim cmd    As New ADODB.Command
Dim myrst  As New ADODB.Recordset
Dim perm   As String
Dim psl    As String

Private Sub cmdSet_Pass_Click()
    If Trim(Text1) = "" Then
        MsgBox "User ID Mandatory"
        Text1.SetFocus
        Exit Sub
    End If

    frmChange_Password.Show vbModal

End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Text1 = Empty Or Text2 = Empty Then
        MsgBox "Incomplete Information", vbOKOnly, "Attention"
        Exit Sub
    End If
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    If opr = "I" Then
'        cmd.CommandText = "Select * from micropass where emp_id='" + RTrim(Text1) + "'"
        cmd.CommandText = "Select * from micropass where u_id='" + RTrim(Text1) + "'"
        Set myrst = cmd.Execute
        If myrst.EOF = False Then
            con.Close
            MsgBox "Record allready exist.", vbOKOnly, "Caution"
            Exit Sub
        End If
    End If
    '-------------------------------

'    cmd.CommandText = "exec pro_micropass 0,'" + Text1 + "','" + Text2 + "','','" _
'    + u_id + "','2000-01-01'," + CStr(Check1.value) + ",'" + opr + "'"

If Command1(0).Caption = "Save" Then
      cmd.CommandText = "exec pro_micropass 0,'" + Text1 + "','" + Text2 + "','','" _
      + u_id + "','2000-01-01'," + CStr(Check1.value) + ",'I'"
Else

     cmd.CommandText = "exec pro_micropass 0,'" + Text1 + "','" + Text2 + "','','" _
     + u_id + "','2000-01-01'," + CStr(Check1.value) + ",'U'"
End If

'    Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
'    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "Select * from micropass where u_id='" & Trim(Text1) & "'"
'   Adodc1.RecordSource = "select * from Pat_Info_main where pat_id='" & Trim(txtPat_ID.Text) & "'"
    
'    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "exec pro_name_SELECT 5,'" & Trim(Text1.Text) & "'"
    Adodc1.Refresh
    
    
    Adodc1.Refresh
    Text1 = ""
    Text2 = ""
    Check1.value = 0
    If opr = "U" Then
        opr = "I"
        Text1.Enabled = True
        Command1(0).Caption = "Save"
        Text1.SetFocus
    End If
    Text1.SetFocus
Case 1
    Text1 = ""
    Text2 = ""
    Check1.value = 0
    If opr = "I" Then Text1.SetFocus
    If opr = "U" Then
        opr = "I"
        Text1.Enabled = True
        Command1(0).Caption = "Save"
        Text1.SetFocus
    End If
Case 2
    Unload Me
End Select
Grid_Width
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If Adodc1.Recordset.EOF Then Exit Sub

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Doc_SELECT 5,'" & Trim(Text1.Text) & "'"
    Adodc1.Refresh
    
If Adodc1.Recordset.EOF Then Exit Sub
'Text1 = Adodc1.Recordset!emp_id
Text1 = Adodc1.Recordset!u_id
Text2 = Adodc1.Recordset!U_Name
If Adodc1.Recordset!Cancel = True Then
    Check1.value = 1
Else
    Check1.value = 0
End If
Text1.Enabled = False
Command1(0).Caption = "Edit"
opr = "U"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
'Adodc1.connectionstring = strcn.Connection
'''''Adodc1.RecordSource = "select emp_id,u_name,'perm' = "
''Adodc1.RecordSource = "select u_id,u_name,'perm' = " _
''& "case when cancel=0 then 'Yes' else 'No' end,cancel from micropass " _
''& "where u_id<>'administrator'"
''''& "where emp_id<>'administrator'"
'Adodc1.RecordSource = "Select * from micropass"
''
''Adodc1.Refresh
''
'Set DataGrid1.DataSource = Adodc1
'DataGrid1.Refresh
''opr = "I"
''    Grid_Width

Adodc1.connectionstring = strcn.Connection
Adodc1.RecordSource = "SELECT u_id As UserID ,u_name,Cancel From micropass ORDER BY u_id"
Adodc1.Refresh
    
'If Adodc1.Recordset.RecordCount > 0 Then
'   Do Until Adodc1.Recordset.EOF
'      Combo1.AddItem Adodc1.Recordset!Doctor
'   Adodc1.Recordset.MoveNext
'   Loop
'End If

Grid_Width



End Sub



Private Sub lblUName_Click()
lblUName.Caption = Text2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()

    If Trim(Text1.Text) = "" Then Exit Sub
    Dim StrUid As String
    Dim StrUid1 As String
    StrUid = Text1.Text
    StrUid1 = Mid(StrUid, 1, 1)
    If StrUid1 = "0" Then
        MsgBox "First letter '0' not allow"
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Me.Command1(0).SetFocus
'End If
End Sub
Private Sub Grid_Width()
     Adodc2.connectionstring = strcn.Connection
     Adodc2.RecordSource = "exec Pro_User 1,''"
'      Adodc2.RecordSource = "select u_id, u_name  from micropass order by u_id"
     Adodc2.Refresh

    DataGrid1.Columns(0).Width = 700
    DataGrid1.Columns(1).Width = 6100
'    DataGrid1.Columns(2).Width = 400
'    DataGrid1.Columns(2).Width = 1395.213
'    DataGrid1.Columns(3).Width = 200
    End Sub

Private Sub DataGrid1_Click()
    Text1 = DataGrid1.Columns(0)
    Text2 = DataGrid1.Columns(1)
    
End Sub


