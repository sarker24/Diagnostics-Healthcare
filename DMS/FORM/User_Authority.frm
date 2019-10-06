VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUser_Authority 
   BackColor       =   &H00C0B4A9&
   Caption         =   "User Permission Setup"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   DrawWidth       =   2
   Icon            =   "User_Authority.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Height          =   5775
      Left            =   5520
      TabIndex        =   16
      Top             =   2520
      Width           =   7935
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5385
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   9499
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   10874778
         HeadLines       =   1
         RowHeight       =   15
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   5775
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   5295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5355
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   9446
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   15
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
      Caption         =   "User Permission Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   13335
      Begin VB.TextBox txtUser_Name 
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   3630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   4230
      End
      Begin VB.TextBox txtScreen_Descrip 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   4260
      End
      Begin VB.ComboBox txtUser_ID 
         Height          =   315
         Left            =   1170
         TabIndex        =   8
         Top             =   270
         Width           =   2385
      End
      Begin VB.ComboBox txtScreen_Name 
         Height          =   315
         ItemData        =   "User_Authority.frx":000C
         Left            =   960
         List            =   "User_Authority.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1590
         Width           =   2415
      End
      Begin VB.ComboBox ComUsr_Type 
         Height          =   315
         ItemData        =   "User_Authority.frx":0010
         Left            =   1170
         List            =   "User_Authority.frx":001D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   2385
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Select"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7980
         TabIndex        =   5
         Top             =   1650
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   930
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1380
      Top             =   8460
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3450
      Top             =   8460
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8430
      Width           =   1050
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
      Height          =   330
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8430
      Width           =   1050
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
      Height          =   330
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8430
      Width           =   1050
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
      Height          =   330
      Left            =   7590
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8430
      Width           =   1050
   End
End
Attribute VB_Name = "frmUser_Authority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChkVal As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
              cmd.CommandText = "exec pro_Soft_Security 'D','" + txtUser_ID.Text + _
              "','" + ComUsr_Type + _
              "','" + txtScreen_Name + _
              "','" + ChkVal + "'"
              cmd.Execute
    con.Close
    
    GetGridData
End Sub

Private Sub cmdSave_Click()

    If Trim(txtUser_ID) = "" Then
        MsgBox " User ID Mandatory"
        txtUser_ID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtUser_Name.Text) = "" Then
        MsgBox "Invalid User ID"
        txtUser_ID = ""
        txtUser_ID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtScreen_Name) = "" Then
        MsgBox " Screen Name Mandatory"
        txtScreen_Name.SetFocus
        Exit Sub
    End If
    
    If Trim(ComUsr_Type) = "" Then
        MsgBox " User Type Mandatory"
        ComUsr_Type.SetFocus
        Exit Sub
    End If
                   
    Allow_Screen
    
'    Dim StrUid As String
'    Dim StrUid1 As String
'    StrUid = txtUser_ID.Text
'    StrUid1 = Mid(StrUid, 1, 1)
'    If StrUid1 = "0" Then
'        MsgBox "not allow"
'    End If
    
    
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Select_Soft_Sucurity 1,'" & Trim(txtUser_ID.Text) & "','" & Trim(txtScreen_Name.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        UpdSoft_Security
        
    Else
        InsSoft_Security

    End If
    
    GetGridData
End Sub

Private Sub DataGrid1_DblClick()
    txtScreen_Name.Text = DataGrid1.Columns(0).value
    txtScreen_Descrip.Text = DataGrid1.Columns(1).value
End Sub

Private Sub DataGrid2_Click()
    txtUser_ID.Text = DataGrid2.Columns(0).value
    txtUser_Name.Text = DataGrid2.Columns(1).value
    txtScreen_Name.Text = DataGrid2.Columns(2).value
    txtScreen_Descrip.Text = DataGrid2.Columns(3).value
    Dim ChVal As String
     ChVal = DataGrid2.Columns(4).value
     If ChVal = "YES" Then
        Check1.value = 1
     End If
     If ChVal = "NO" Then
        Check1.value = 0
     End If
     
    ComUsr_Type.Text = DataGrid2.Columns(5).value
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
    ComUsr_Type = "User"
    'strScr_Name = "frmUser_Authority"
    'Authority
    'If strAllow = "NO" Then
    'Unload Me
    'End If

    User_ID 'for search software user name
   
    Screen_Name 'for select screen name
    GetGridDataAll
    GetGridData 'get grid data from "soft_security"
End Sub


Private Sub txtScreen_Name_LostFocus()
    Screen_Describ
End Sub

Private Sub txtUser_ID_LostFocus()
    If Trim(txtUser_ID.Text) = "" Then Exit Sub
    User_Name
    GetGridData
End Sub
Private Sub User_ID() 'to search user name

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    
    My_Rst.Open "exec Select_Soft_Sucurity 2,'',''", con

    If My_Rst.EOF = False Then
    Do Until My_Rst.EOF
    txtUser_ID.AddItem My_Rst!u_id
    My_Rst.MoveNext
    Loop
    End If
    con.Close
    
End Sub
Private Sub User_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'My_Rst.Open "select u_id from micropass", con
    My_Rst.Open "exec pro_name_select '7','" + Trim(txtUser_ID.Text) + "'", con
    If My_Rst.EOF = False Then
    txtUser_Name.Text = My_Rst!U_Name
    Else
    txtUser_ID = ""
    txtUser_Name.Text = ""
    End If
    con.Close
End Sub
Private Sub Screen_Name() 'to search screen name

    Dim My_Rst1 As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    'My_Rst1.Open "select scr_no from soft_bag", con
    My_Rst1.Open "exec Select_Soft_Sucurity 3,'',''", con
    'My_Rst.Open "exec pro_name_select '7',''", con
    If My_Rst1.EOF = False Then
    Do Until My_Rst1.EOF
    txtScreen_Name.AddItem My_Rst1!scr_no
    My_Rst1.MoveNext
    Loop
    End If
    con.Close
    
End Sub
Private Sub Screen_Describ()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'My_Rst.Open "select u_id from micropass", con
    My_Rst.Open "exec pro_name_select '8','" + Trim(txtScreen_Name.Text) + "'", con
    If My_Rst.EOF = False Then
    txtScreen_Descrip.Text = My_Rst!descript
    End If
    con.Close
End Sub
Private Sub GetGridDataAll()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_Test_Info_FLUSH '3',''"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Width = 2
    DataGrid1.Columns(1).Width = 4500
    
End Sub
Private Sub GetGridData()
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec pro_GetAuthority '4','','" & txtUser_ID.Text & "'"
    Adodc2.Refresh
    
    Set DataGrid2.DataSource = Adodc2.Recordset
    
    DataGrid2.Columns(0).Width = 1000
    DataGrid2.Columns(1).Width = 2500
    DataGrid2.Columns(2).Width = 0
    DataGrid2.Columns(3).Width = 2500
    DataGrid2.Columns(4).Width = 500
    DataGrid2.Columns(5).Width = 0
    
End Sub
Private Sub InsSoft_Security()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
              cmd.CommandText = "exec pro_Soft_Security 'I','" + txtUser_ID + _
              "','" + ComUsr_Type + _
              "','" + txtScreen_Name + _
              "','" + ChkVal + "'"
              cmd.Execute
              
              'Debug.Print cmd.CommandText
              
    con.Close

End Sub
Private Sub UpdSoft_Security()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
              cmd.CommandText = "exec pro_Soft_Security 'U','" + txtUser_ID + _
              "','" + ComUsr_Type + _
              "','" + txtScreen_Name + _
              "','" + ChkVal + "'"
              cmd.Execute
    con.Close

End Sub
Private Sub Allow_Screen()
    If Check1.value = 1 Then
        ChkVal = "YES"
    End If
    If Check1.value = 0 Then
        ChkVal = "NO"
    End If
End Sub
