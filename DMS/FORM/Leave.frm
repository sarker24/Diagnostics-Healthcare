VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLeave 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   DrawWidth       =   2
   Icon            =   "Leave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow_All 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   540
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1665
      Left            =   540
      TabIndex        =   20
      Top             =   2790
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   2937
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
         BeginProperty Column00 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   870.236
         EndProperty
      EndProperty
   End
   Begin VB.TextBox nbrPresent_Leave 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1530
      TabIndex        =   6
      Top             =   2370
      Width           =   630
   End
   Begin VB.ComboBox comLeave_Type 
      Height          =   315
      ItemData        =   "Leave.frx":000C
      Left            =   1530
      List            =   "Leave.frx":000E
      TabIndex        =   4
      Text            =   "comLeave_Type"
      Top             =   1770
      Width           =   1845
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
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4695
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
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4695
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4695
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
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4695
      Width           =   1050
   End
   Begin VB.TextBox txtEmp_Name 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   870
      Width           =   3270
   End
   Begin VB.TextBox txtEmp_ID 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1590
      TabIndex        =   0
      Top             =   870
      Width           =   1245
   End
   Begin VB.TextBox nbrTot_Day 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4950
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   630
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4950
      Top             =   90
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin MSComCtl2.DTPicker StDate 
      Height          =   285
      Left            =   1530
      TabIndex        =   2
      Top             =   1275
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   6755585
      CalendarForeColor=   16777215
      Format          =   56360961
      CurrentDate     =   37114
   End
   Begin MSComCtl2.DTPicker EdDate 
      Height          =   285
      Left            =   4890
      TabIndex        =   3
      Top             =   1260
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   6755585
      CalendarForeColor=   16777215
      Format          =   56360961
      CurrentDate     =   37114
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "day"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2310
      TabIndex        =   19
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "day"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   5820
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave for"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   2340
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave"
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
      Height          =   360
      Left            =   570
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Type"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   15
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   14
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave from "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   13
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4440
      TabIndex        =   12
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4410
      TabIndex        =   11
      Top             =   1290
      Width           =   195
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrBalance As String
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    txtEmp_ID = ""
    txtEmp_Name = ""
    StDate = Now
    EdDate = Now
    comLeave_Type = "Casual"
    Me.nbrTot_Day = 0
    txtEmp_ID.SetFocus
End Sub

Private Sub cmdSave_Click()
If Trim(txtEmp_ID.Text) = "" Then Exit Sub
If comLeave_Type = "" Then Exit Sub
If Val(nbrTot_Day) = 0 Then Exit Sub

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec SELECT_Leave '1','" & Trim(txtEmp_ID.Text) & "','" & Format(StDate, "yyyy-mm-dd") & "','" & Format(EdDate, "yyyy-mm-dd") & "','" & Trim(comLeave_Type.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Leave_Transection
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Leave_Transection
        MsgBox "Inserted successfully"
    End If
'    txtEmp_ID = ""
    txtEmp_Name = ""
    StDate = Date
    EdDate = Date
    comLeave_Type = "Casual"
'    nbrTot_Day = ""
    nbrPresent_Leave = ""

    GetGridData
    txtEmp_ID.SetFocus
End Sub

Private Sub Leave_Transection()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_Leave_IUD '" + Mode + "','" + txtEmp_ID.Text + _
    "','" + Format(StDate.value, "yyyy-mm-dd") + _
    "','" + Format(EdDate.value, "yyyy-mm-dd") + _
    "','" + comLeave_Type + _
    "'," + nbrPresent_Leave + ""
    cmd.Execute
'    MsgBox cmd.Execute
    con.Close
End Sub

Private Sub cmdShow_All_Click()
    GetGridDataAll
End Sub

Private Sub comLeave_Type_Click()
    Tot_Leave
'    Dim My_Rst As New ADODB.Recordset
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'
'    My_Rst.Open "exec pro_name_SELECT '10','" + comLeave_Type + "'", con
'    If My_Rst.EOF = False Then
'        nbrTot_Day = My_Rst!Celing
'        StrBalance = My_Rst!Celing
'    Else
'        nbrTot_Day = ""
'    End If
'    con.Close
End Sub

'    cmd.Execute
''    MsgBox cmd.Execute
'    con.Close
'
'End Sub
Private Sub comLeave_Type_LostFocus()
    Tot_Leave
'    Dim My_Rst As New ADODB.Recordset
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'
'    My_Rst.Open "exec pro_name_SELECT '10','" + comLeave_Type + "'", con
'    If My_Rst.EOF = False Then
'        nbrTot_Day = My_Rst!Celing
'        StrBalance = My_Rst!Celing
'    Else
'        nbrTot_Day = ""
'    End If
'    con.Close
End Sub

Private Sub EdDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
    StDate = Date
    EdDate = Date
    
    Search_Leave_Type
    'Me.comLeave_Type.AddItem "Casual"
        Me.comLeave_Type.Text = "Casual"
    'GetGridDataAll
End Sub

Private Sub Total_Leave()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '10','" + comLeave_Type + "'", con
    If My_Rst.EOF = False Then
        nbrTot_Day = My_Rst!Celing
        StrBalance = My_Rst!Celing
    Else
        nbrTot_Day = ""
    End If
    con.CloseEnd
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
    My_Rst.MoveNext
    Loop
    End If
    con.Close
    'comLeave_Type.Refresh
End Sub

Private Sub Search_Emp_Info() 'search EMPLOYEE NAME
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '9','" + txtEmp_ID.Text + "'", con
    If My_Rst.EOF = False Then
        txtEmp_ID.Text = My_Rst!emp_id
        txtEmp_Name.Text = My_Rst!Emp_Name
    Else
        MsgBox "Invalid Employee ID, Try again...."
        txtEmp_Name.Text = ""
        txtEmp_ID.SetFocus
    End If
    con.Close
End Sub

Private Sub StDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub txtEmp_ID_LostFocus()
    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    
    Search_Emp_Info 'search EMPLOYEE NAME
    
    GetGridData
End Sub

Private Sub GetGridData()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec leave_balance 1,'" + Me.txtEmp_ID + "'"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 2500
    DataGrid1.Columns(2).Width = 700
    DataGrid1.Columns(3).Width = 500
    DataGrid1.Columns(4).Width = 500
    DataGrid1.Columns(5).Width = 500
End Sub

Private Sub GetGridDataAll()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec leave_balance1 1"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 2500
    DataGrid1.Columns(2).Width = 700
    DataGrid1.Columns(3).Width = 500
    DataGrid1.Columns(4).Width = 500
    DataGrid1.Columns(5).Width = 500
End Sub

Private Sub Tot_Leave()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '10','" + comLeave_Type + "'", con
    If My_Rst.EOF = False Then
        nbrTot_Day = My_Rst!Celing
        StrBalance = My_Rst!Celing
    Else
        nbrTot_Day = ""
    End If
    con.Close
End Sub
