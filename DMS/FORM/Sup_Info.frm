VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSup_Info 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   DrawWidth       =   2
   Icon            =   "Sup_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Supplier Details Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   8895
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2715
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   4789
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   0   'False
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
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3869.858
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Supplier Informations"
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   8895
      Begin VB.TextBox txtSup_Name 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   8
         Top             =   750
         Width           =   6120
      End
      Begin VB.TextBox txtSup_ID 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   7
         Top             =   390
         Width           =   1245
      End
      Begin VB.TextBox txtSup_Add 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   6120
      End
      Begin VB.TextBox txtSup_Phone 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   5
         Top             =   1650
         Width           =   5160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   11
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   1650
         Width           =   450
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
      Height          =   420
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5925
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
      Height          =   420
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5925
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
      Height          =   420
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5925
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
      Height          =   420
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5925
      Width           =   900
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5880
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
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
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmSup_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtSup_ID.Text = "" Then Exit Sub
    Mode = "D"
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        Sup_Info_IUD
        txtSup_ID = ""
        txtSup_Name = ""
    End If
    
End Sub

Private Sub cmdNew_Click()
    txtSup_ID = ""
    txtSup_Name = ""
    txtSup_ID.SetFocus
    
End Sub
Private Sub cmdSave_Click()
    If txtSup_ID = "" Then Exit Sub
    If txtSup_Name = "" Then Exit Sub
'    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '12','" & txtSup_ID & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Sup_Info_IUD
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Sup_Info_IUD
        MsgBox "inserted successfully"
    End If
    
    txtSup_ID = ""
    txtSup_Name = ""
    txtSup_Add = ""
    txtSup_Phone = ""
    txtSup_ID.SetFocus
    Show_Grid
End Sub
Private Sub Sup_Info_IUD()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Sup_Info_IUD '" & Mode & "','" & txtSup_ID & _
    "','" & txtSup_Name.Text & _
    "','" & txtSup_Add.Text & _
    "','" & txtSup_Phone.Text & _
    "','" & u_id & "'"
    
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
    txtSup_ID = DataGrid1.Columns(0).value
    txtSup_Name = DataGrid1.Columns(1).value
    txtSup_Add = DataGrid1.Columns(2).value
    txtSup_Phone = DataGrid1.Columns(3).value
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Show_Grid
End Sub

Private Sub txtSup_ID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtSup_ID_LostFocus()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '12','" + txtSup_ID + "'", con
    If My_Rst.EOF = False Then
        Me.txtSup_ID.Text = My_Rst!sup_id
        txtSup_Name.Text = My_Rst!sup_name
    Else
        txtSup_Name.Text = ""
    End If
    
    con.Close
    
End Sub
Private Sub Show_Grid()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '16',''"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(0).Width = 1275.024
    DataGrid1.Columns(1).Width = 3869.858
        
End Sub

Private Sub txtSup_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtSup_Phone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub
