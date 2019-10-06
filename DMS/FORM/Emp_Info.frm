VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmp_Info 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   DrawWidth       =   2
   Icon            =   "Emp_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Employee Details Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   27
      Top             =   4920
      Width           =   9735
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3525
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   6218
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
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
      Caption         =   "Employee Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   9735
      Begin VB.TextBox txtPre_Add 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   4350
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   2415
         Width           =   3570
      End
      Begin VB.TextBox txtEmp_Name 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1740
         TabIndex        =   30
         Top             =   630
         Width           =   6180
      End
      Begin VB.TextBox txtEmail 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4365
         TabIndex        =   29
         Top             =   3750
         Width           =   3555
      End
      Begin VB.TextBox txtTitle 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4350
         TabIndex        =   10
         Top             =   1290
         Width           =   1680
      End
      Begin VB.TextBox txtPhone 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   360
         TabIndex        =   19
         Top             =   3750
         Width           =   3675
      End
      Begin VB.TextBox txtEmp_ID 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   360
         TabIndex        =   7
         Top             =   630
         Width           =   1245
      End
      Begin VB.ComboBox ComDesig 
         Height          =   315
         ItemData        =   "Emp_Info.frx":000C
         Left            =   1800
         List            =   "Emp_Info.frx":0013
         TabIndex        =   9
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox nbrSalary 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6180
         TabIndex        =   12
         Top             =   1290
         Width           =   1740
      End
      Begin VB.TextBox txtPer_Add 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2415
         Width           =   3675
      End
      Begin VB.TextBox nbrAge 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4365
         TabIndex        =   14
         Top             =   1740
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   6
         Top             =   1740
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   5
         Top             =   1740
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker Join_date 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   1290
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   57344003
         CurrentDate     =   37367
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4350
         TabIndex        =   25
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4365
         TabIndex        =   24
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phones"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   23
         Top             =   3495
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary (Tk.)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6120
         TabIndex        =   20
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4365
         TabIndex        =   18
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   1740
         Width           =   270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age (Years)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3360
         TabIndex        =   13
         Top             =   1740
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3120
      Top             =   8880
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8880
      Width           =   930
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8880
      Width           =   930
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0FF&
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8880
      Width           =   930
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8880
      Width           =   930
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   8880
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmEmp_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sex As String
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Mode = "D"
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        Emp_Transaction
        txtEmp_ID = ""
    Clearscreen
    End If
End Sub

Private Sub cmdNew_Click()
    txtEmp_ID = ""
    Clearscreen
End Sub

Private Sub cmdSave_Click()
    If Trim(txtEmp_ID.text) = "" Then
        MsgBox "Employee ID Required"
        txtEmp_ID.SetFocus
        Exit Sub
    End If
       
    If txtEmp_Name = "" Then
        MsgBox "Employee Name Required"
        txtEmp_Name.SetFocus
        Exit Sub
    End If
    
    If ComDesig.text = "" Then
        MsgBox "Employee Designation Required"
        ComDesig.SetFocus
        Exit Sub
    End If
    
    If txtPer_Add.text = "" Then
        MsgBox "Employee's Present Address Required"
        txtPer_Add.SetFocus
        Exit Sub
    End If
    
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '9','" & Trim(txtEmp_ID.text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Emp_Transaction
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Emp_Transaction
        MsgBox "inserted successfully"
    End If
        
    txtEmp_ID.text = ""
    txtEmp_Name.text = ""
    Join_date.value = Date
    ComDesig = ""
    txtTitle = ""
    nbrSalary = ""
    Option1.value = True
    nbrAge.text = ""
    txtPer_Add = ""
    txtPre_Add = ""
    txtPhone = ""
    txtEmail = ""
    txtEmp_ID.SetFocus
    
    GetGridData
End Sub

Private Sub Emp_Transaction()
    If Option1.value = True Then
        Sex = "Male"
    Else
        Sex = "Female"
    End If
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_Emp_Info '" + Mode + "','" + txtEmp_ID.text + _
    "','" + txtEmp_Name.text + _
    "','" + Format(Join_date.value, "yyyy-mm-dd") + _
    "','" + ComDesig + _
    "','" + txtTitle + _
    "'," + nbrSalary + _
    ",'" + Sex + _
    "','" + Trim(nbrAge.text) + _
    "','" + txtPre_Add.text + _
    "','" + txtPer_Add.text + _
    "','" + txtPhone.text + _
    "','" + txtEmail.text + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub ComDesig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub DataGrid1_Click()
    txtEmp_ID.text = DataGrid1.Columns(0).value
    txtEmp_Name.text = DataGrid1.Columns(1).value
    Join_date.value = DataGrid1.Columns(2).value
    ComDesig.text = DataGrid1.Columns(3).value
    txtTitle.text = DataGrid1.Columns(4).value
    nbrSalary.text = DataGrid1.Columns(5).value
    Dim StrMF As String
    
    StrMF = DataGrid1.Columns(6).value
    If StrMF = "Male" Then
        Option1.value = True
    End If
    If StrMF = "Female" Then
        Option2.value = True
    End If
    
    nbrAge.text = DataGrid1.Columns(7).value
    txtPer_Add.text = DataGrid1.Columns(8).value
    txtPre_Add.text = DataGrid1.Columns(9).value
    txtPhone.text = DataGrid1.Columns(10).value
    txtEmail.text = DataGrid1.Columns(11).value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    nbrSalary.text = "0"
    Join_date.value = Date
    GetGridData
End Sub

Private Sub Join_date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub nbrAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub nbrSalary_Change()
    If Not IsNumeric(nbrSalary.text) Then
        'MsgBox "Only Numaric value allow"
        nbrSalary = "0"
        nbrSalary.SelStart = 0
        nbrSalary.SelLength = Len(nbrSalary)
        nbrSalary.SetFocus
    End If
End Sub

Private Sub nbrSalary_GotFocus()
    nbrSalary.SelStart = 0
    nbrSalary.SelLength = Len(nbrSalary)
End Sub

Private Sub nbrSalary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
    
    If KeyAscii = 44 Then KeyAscii = 0
End Sub

Private Sub nbrTot_Leave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtEmp_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtEmp_ID_LostFocus()
    If txtEmp_ID.text = "" Then Exit Sub
    Search_Emp_Info
End Sub

Private Sub Search_Emp_Info()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '9','" + txtEmp_ID.text + "'", con
    If My_Rst.EOF = False Then
    
        txtEmp_ID.text = My_Rst!emp_id
        txtEmp_Name.text = My_Rst!Emp_Name
        Join_date.value = My_Rst!Join_date
        ComDesig = My_Rst!Emp_Desig
        txtTitle = My_Rst!Title
        nbrSalary = My_Rst!Salary
        Dim StrSex As String
        
        StrSex = My_Rst!Salary
        
        If StrSex = "Male" Then
            Option1.value = True
        End If
                
        If StrSex = "Female" Then
            Option2.value = True
        End If
        
        nbrAge = My_Rst!age
        txtPer_Add = My_Rst!Emp_Per_Add
        txtPre_Add = My_Rst!Emp_Pre_Add
        txtPhone = My_Rst!Emp_Phone
        txtEmail = My_Rst!Emp_Email
        Else
            txtEmp_Name.text = ""
            Join_date.value = Date
            ComDesig = ""
            txtTitle = ""
            nbrSalary = "0"
            Option1.value = True
            nbrAge = ""
            txtPer_Add = ""
            txtPre_Add = ""
            txtPhone = ""
            txtEmail = ""
    End If
    con.Close
End Sub

Private Sub txtEmp_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub Clearscreen()
    txtEmp_Name.text = ""
    Join_date.value = Date
    ComDesig = ""
    txtTitle = ""
    nbrSalary = ""
    Option1.value = True
    nbrAge.text = ""
    txtPer_Add = ""
    txtPre_Add = ""
    txtPhone = ""
    txtEmail = ""
    txtEmp_ID.SetFocus
End Sub

Private Sub GetGridData()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec leave_balance 4,''"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
'    DataGrid1.Columns(0).Width = 0
'    DataGrid1.Columns(1).Width = 1000
'    DataGrid1.Columns(2).Width = 2000
'    DataGrid1.Columns(3).Width = 1000
'    DataGrid1.Columns(4).Width = 2000
'    DataGrid1.Columns(5).Width = 700
'    DataGrid1.Columns(6).Width = 1000
'    DataGrid1.Columns(8).Width = 0
End Sub
