VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStock_Out 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Item Issues Information"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   DrawWidth       =   2
   Icon            =   "Stock_Out.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   2415
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   10215
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   3625
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   8775
      Begin VB.TextBox txtOut_No 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   11
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox txtEmp_ID 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   10
         Top             =   690
         Width           =   1260
      End
      Begin VB.TextBox nbrItem_Qty 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   9
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtItem_Name 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4170
      End
      Begin VB.TextBox txtItem_Code 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   7
         Top             =   1110
         Width           =   1245
      End
      Begin VB.TextBox txtNotes 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   6
         Top             =   1935
         Width           =   5490
      End
      Begin VB.TextBox txtEmp_Name 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   4200
      End
      Begin MSComCtl2.DTPicker issu_date 
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47382529
         CurrentDate     =   37365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   240
         TabIndex        =   16
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issu Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   2370
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   1950
         Width           =   375
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
      Height          =   330
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
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
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
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
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1050
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   540
      Top             =   5400
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2430
      Top             =   5040
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   540
      Top             =   5040
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Adodc3"
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
End
Attribute VB_Name = "frmStock_Out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    txtItem_Code = ""
    txtItem_Name = ""
    txtEmp_ID = ""
    txtEmp_Name = ""
    Me.nbrItem_Qty = ""
    Me.issu_date = Date
    txtNotes = ""
    txtOut_No.SetFocus
End Sub

Private Sub cmdSave_Click()
    
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Inv_NO_SELECT '2','" & Trim(txtOut_No.Text) & "','" & Trim(txtItem_Code.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Stock_Out_Trans
        MsgBox "Updated Successfully"
    Else
        Mode = "I"
        Stock_Out_Trans
        MsgBox "Inserted Successfully"
    End If
    
    GetGridData
    
    txtItem_Code = ""
    txtItem_Name = ""
    txtEmp_ID = ""
    txtEmp_Name = ""
    Me.nbrItem_Qty = ""
    Me.issu_date = Date
    txtNotes = ""
    txtOut_No.SetFocus
    
    
End Sub
Private Sub Stock_Out_Trans()
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Stock_Out_IUD '" + Mode + "','" + txtOut_No + _
    "','" + txtItem_Code.Text + _
    "'," + nbrItem_Qty + _
    ",'" + txtEmp_ID + _
    "','" + txtNotes + _
    "','" + Format(issu_date, "yyyy-mm-dd") + _
    "','" + u_id + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub
Private Sub DataGrid1_DblClick()
On Error Resume Next
    txtOut_No = DataGrid1.Columns(0).value
    txtEmp_ID = DataGrid1.Columns(1).value
    txtEmp_Name = DataGrid1.Columns(2).value
    txtItem_Code = DataGrid1.Columns(3).value
    txtItem_Name = DataGrid1.Columns(4).value
    nbrItem_Qty = DataGrid1.Columns(5).value
    txtNotes = DataGrid1.Columns(7).value
    issu_date.value = DataGrid1.Columns(6).value
        
        
End Sub

'Private Sub Search_Item_Name()
'    Dim My_Rst As New ADODB.Recordset
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'
'    My_Rst.Open "exec pro_name_SELECT '11','" + Me.txtItem_Code + "'", con
'    If My_Rst.EOF = False Then
'        txtItem_Code.Text = My_Rst!item_code
'        txtItem_Name.Text = My_Rst!item_name
'    Else
'        'txtItem_Name.Text = ""
'        frmItem_List.Show vbModal
'        Exit Sub
'    End If
'
'    con.Close
'
'End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
    
    If KeyAscii = 27 Then
    Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    issu_date.value = Date
End Sub

Private Sub txtEmp_ID_LostFocus()

    If txtEmp_ID = "" Then Exit Sub
    
    Emp_List_MODE = "frmStock_Out"
        
    Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = "exec pro_name_SELECT '9','" & Trim(txtEmp_ID.Text) & "'"
    Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtEmp_ID.Text = Adodc3.Recordset!emp_id
        txtEmp_Name.Text = Adodc3.Recordset!Emp_Name
        'MsgBox "11"
    Else
       frmEmp_List.Show vbModal
       Exit Sub
    End If
        
    
    'Search_Emp_Name
    GetGridData
End Sub

Private Sub txtItem_Code_LostFocus()
    If txtItem_Code = "" Then Exit Sub
    
    Item_List_MODE = "frmStock_Out"
        
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec pro_name_SELECT '11','" & Trim(txtItem_Code.Text) & "'"
    Adodc2.Refresh
    
    If Adodc2.Recordset.RecordCount > 0 Then
        txtItem_Code.Text = Adodc2.Recordset!item_code
        txtItem_Name.Text = Adodc2.Recordset!item_name
        
    Else
       frmItem_List.Show vbModal
       Exit Sub
    End If
    
    'Search_Item_Name
    
End Sub
Private Sub Search_Emp_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '9','" + txtEmp_ID.Text + "'", con
    If My_Rst.EOF = False Then
    
        txtEmp_ID.Text = My_Rst!emp_id
        txtEmp_Name.Text = My_Rst!Emp_Name
    Else
        txtEmp_ID.Text = ""
        txtEmp_Name.Text = ""
               
    End If
    con.Close

End Sub
Private Sub GetGridData()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec leave_balance 3,'" & txtEmp_ID & "'"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Width = 500
    DataGrid1.Columns(1).Width = 1000
    DataGrid1.Columns(2).Width = 2000
    DataGrid1.Columns(3).Width = 1000
    DataGrid1.Columns(4).Width = 2000
    DataGrid1.Columns(5).Width = 700
    DataGrid1.Columns(6).Width = 1000
    DataGrid1.Columns(8).Width = 0
    
    
End Sub

