VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUser_Authority 
   BackColor       =   &H00800000&
   Caption         =   "Prime Diagnostic Ltd."
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   Icon            =   "User_Authority.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5460
      Top             =   180
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Left            =   7020
      Top             =   180
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800000&
      Caption         =   "Select"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8940
      TabIndex        =   13
      Top             =   2460
      Width           =   795
   End
   Begin VB.ComboBox ComUsr_Type 
      Height          =   315
      ItemData        =   "User_Authority.frx":030A
      Left            =   2130
      List            =   "User_Authority.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1350
      Width           =   2385
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1065
      Left            =   2100
      TabIndex        =   11
      Top             =   2850
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   1879
      _Version        =   393216
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
   Begin VB.ComboBox txtScreen_Name 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox txtUser_ID 
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   750
      Width           =   2385
   End
   Begin VB.TextBox txtScreen_Descrip 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2430
      Width           =   4260
   End
   Begin VB.TextBox txtUser_Name 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   4590
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   4230
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
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5910
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
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5910
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5910
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
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5910
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1875
      Left            =   870
      TabIndex        =   14
      Top             =   3960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3307
      _Version        =   393216
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   750
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   2430
      Width           =   975
   End
End
Attribute VB_Name = "frmUser_Authority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChkVal As String
Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Trim(txtUser_ID) = "" Then
        MsgBox " User ID Mandatory"
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
    
End Sub
Private Sub Form_Load()
    
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
End Sub
Private Sub User_ID() 'to search user name

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "select u_id from micropass", con
    'My_Rst.Open "exec pro_name_select '7',''", con
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
    End If
    con.Close
End Sub
Private Sub Screen_Name() 'to search screen name

    Dim My_Rst1 As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst1.Open "select scr_no from soft_bag", con
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
    
    DataGrid1.Columns(0).Width = 2000
    DataGrid1.Columns(1).Width = 4500
    
End Sub
Private Sub GetGridData()
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec pro_Test_Info_FLUSH '4',''"
    Adodc2.Refresh
    
    Set DataGrid2.DataSource = Adodc2.Recordset
    
    DataGrid2.Columns(0).Width = 1000
    DataGrid2.Columns(1).Width = 2500
    DataGrid2.Columns(2).Width = 2000
    DataGrid2.Columns(3).Width = 2500
    DataGrid2.Columns(4).Width = 500
    DataGrid2.Columns(5).Width = 600
    
End Sub
Private Sub InsSoft_Security()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
              cmd.CommandText = "exec pro_Soft_Security 'I'," + txtUser_ID + _
              ",'" + ComUsr_Type + _
              "','" + txtScreen_Name + _
              "','" + ChkVal + "'"
              cmd.Execute
    con.Close

End Sub
Private Sub UpdSoft_Security()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
              cmd.CommandText = "exec pro_Soft_Security 'U'," + txtUser_ID + _
              ",'" + ComUsr_Type + _
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
