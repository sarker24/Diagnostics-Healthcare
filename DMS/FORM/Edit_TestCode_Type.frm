VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEdit_TestCode_Type 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Test Modification"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   DrawWidth       =   2
   Icon            =   "Edit_TestCode_Type.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   8535
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
         Height          =   315
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox txtPat_Name 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1380
         TabIndex        =   12
         Top             =   810
         Width           =   4080
      End
      Begin VB.TextBox txtPat_ID 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   810
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox nbrAdv 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   5490
         TabIndex        =   10
         Top             =   810
         Width           =   765
      End
      Begin VB.TextBox nbrColl_Fee 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6270
         TabIndex        =   9
         Top             =   810
         Width           =   795
      End
      Begin VB.CommandButton cmdPatOld 
         BackColor       =   &H00FFFFFF&
         Caption         =   "O&ld"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   540
      End
      Begin VB.CommandButton cmdPatNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ne&w"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtPat_ID1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   810
         Width           =   1230
      End
      Begin VB.TextBox txtType 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   7080
         TabIndex        =   5
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton Command1 
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
         Height          =   315
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1395
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2461
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   570
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5490
         TabIndex        =   17
         Top             =   570
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6330
         TabIndex        =   16
         Top             =   570
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   15
         Top             =   570
         Width           =   435
      End
   End
   Begin VB.TextBox txtDummy_Pat_ID 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4230
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   2970
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtUnique_ID 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   1245
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
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2745
      Width           =   930
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5580
      Top             =   2760
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
End
Attribute VB_Name = "frmEdit_TestCode_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String

Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double

Private Sub GetGridData()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Pat_Info_SELECT 10," & txtPat_ID & ""
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 1280
    DataGrid1.Columns(2).Width = 4080
    DataGrid1.Columns(3).Width = 800
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 0
    
End Sub
Private Sub GetGridDataTest()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Pat_Info_SELECT 13," & txtPat_ID & ""
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 1280
    DataGrid1.Columns(2).Width = 4080
    DataGrid1.Columns(3).Width = 800
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 1000
    DataGrid1.Columns(6).Width = 0
    
End Sub
Private Sub GetGridDataTestAll()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Pat_Info_SELECT 14,''"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 1280
    DataGrid1.Columns(2).Width = 4080
    DataGrid1.Columns(3).Width = 800
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 1000
    DataGrid1.Columns(6).Width = 0
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPatNew_Click()
    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = True
    txtPat_ID.Visible = False
    
    txtPat_ID1.SetFocus
End Sub

Private Sub cmdPatOld_Click()
    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = False
    txtPat_ID.Visible = True
    txtPat_ID.SetFocus
End Sub

Private Sub cmdSave_Click()
    If txtPat_Name = "" Or txtPat_Name = "0" Then Exit Sub
    If txtUnique_ID = "" Then Exit Sub

    
'    If u_id <> "md" Then
'        MsgBox "You are not Authorized person, please contact to Mr. Bashar  ", vbCritical
'        txtPat_ID.Text = "0"
'        txtPat_Name = "0"
'        nbrAdv = "0"
'        nbrColl_Fee = "0"
'        txtUnique_ID = ""
'
'        Exit Sub
'    End If
    
        
'        If Option1.value = True Then
'
'            Mode = "U"
'            UPat_Pay
'            MsgBox "Successfully Updated"
'            GetGridData
'
'        End If
        
'        If Option2.value = True Then
        
            Mode = "UT"
            UPat_Test_Code
            
            GetGridDataTest
            
'        End If
        
        nbrAdv = "0"
        nbrColl_Fee = "0"
        txtUnique_ID = ""
        
End Sub

Private Sub Command1_Click()
GetGridDataTestAll
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next

'If Option1.value = True Then
'    txtPat_ID.Text = DataGrid1.Columns(0).value
'    txtPat_Name = DataGrid1.Columns(2).value
'    nbrAdv = DataGrid1.Columns(3).value
'    nbrColl_Fee = DataGrid1.Columns(4).value
'    txtUnique_ID = DataGrid1.Columns(5).value
'
'    nbrAdv.SetFocus
'End If
    
    
'If Option2.value = True Then

    txtPat_ID.Text = DataGrid1.Columns(0).value
    txtPat_ID1.Text = DataGrid1.Columns(1).value
    txtPat_Name = DataGrid1.Columns(2).value
    nbrAdv = DataGrid1.Columns(3).value
    nbrColl_Fee = DataGrid1.Columns(4).value
    txtType.Text = DataGrid1.Columns(5).value
    txtUnique_ID = DataGrid1.Columns(6).value
    nbrAdv.SetFocus
    
'End If

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

    txtPat_ID.Text = ""
    txtPat_Name = ""
    nbrAdv = "0"
    nbrColl_Fee = "0"
    txtUnique_ID = ""
    
End Sub

Private Sub nbrAdv_Change()
    If Not IsNumeric(nbrAdv.Text) Then
        MsgBox "Only Numaric value allow"
        nbrAdv = 0
        nbrAdv.SelStart = 0
        nbrAdv.SelLength = Len(nbrAdv)
        nbrAdv.SetFocus
    End If
    
End Sub
Private Sub nbrAdv_GotFocus()
        nbrAdv.SelStart = 0
        nbrAdv.SelLength = Len(nbrAdv)
End Sub
Private Sub nbrColl_Fee_Change()
    If Not IsNumeric(nbrColl_Fee.Text) Then
        MsgBox "Only Numaric value allow"
        nbrColl_Fee = 0
        nbrColl_Fee.SelStart = 0
        nbrColl_Fee.SelLength = Len(nbrColl_Fee)
        
    End If
    
End Sub
Private Sub nbrColl_Fee_GotFocus()
    nbrColl_Fee.SelStart = 0
    nbrColl_Fee.SelLength = Len(nbrColl_Fee)
End Sub

'Private Sub Option1_Click()
'Label4.Caption = "Payment"
'Label5.Caption = "Collection Fee"
'End Sub

'Private Sub Option2_Click()
'Label4.Caption = "Sub Code"
'Label5.Caption = "Main Code"
'txtPat_ID1.SetFocus
'
'End Sub

Private Sub txtPat_ID_Change()
If txtPat_ID.Visible = True Then
    If Not IsNumeric(txtPat_ID.Text) Then
'        MsgBox "Only Numaric value allow"
        txtPat_ID = ""
        txtPat_ID.SelStart = 0
        txtPat_ID.SelLength = Len(txtPat_ID)
        txtPat_ID.SetFocus
    End If
End If
End Sub

Private Sub txtPat_ID_GotFocus()
    txtPat_ID.SelStart = 0
    txtPat_ID.SelLength = Len(txtPat_ID)
End Sub

Private Sub txtPat_ID_LostFocus()
    If Trim(txtPat_ID) = "" Then Exit Sub
'    If Option1.value = True Then
'        GetGridData
'    End If
End Sub
Private Sub UPat_Test_Code()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec U_PAT_Test_Code '" & Mode & "','" & nbrAdv & _
    "','" & nbrColl_Fee & "','" & txtType.Text & "'," & txtUnique_ID & ""
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub txtPat_ID1_LostFocus()

'If Option1.value = True Then

    If txtPat_ID1 = "" Then Exit Sub
    
    Search_Patient_Type
    
    If StrRow_Count > "1" Then
        
            Dim Patmsg As String
            Patmsg = MsgBox("Do you want to update Inside Patient's information ? ", vbQuestion + vbYesNo)
            If Patmsg = vbYes Then
                StrPat_Type = "0"
                Srch_Pat_ID
            Else
                StrPat_Type = "1"
                Srch_Pat_ID
            End If
    Else
            Srch_Pat_ID1
    End If
    
   
    txtPat_ID = IntPat_ID
    
    txtDummy_Pat_ID.Text = IntPat_ID
    
    If IntPat_ID = 0 Then
        MsgBox "Invalid ID, Try again"
        txtPat_ID = ""
        txtPat_ID1 = ""
        txtDummy_Pat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If
    
'If Me.Option1.value = True Then
'    GetGridData
    
'End If

'If Option2.value = True Then

    GetGridDataTest
    
'End If

End Sub


Private Sub Search_Patient_Type()

    StrRow_Count = "1"
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_Type 1,'" & txtPat_ID1.Text & "'", con
    If My_Rst.EOF = False Then
    
        StrRow_Count = My_Rst!Row_Count
        'MsgBox StrRow_Count
    End If
    
    con.Close
    
End Sub
Private Sub Srch_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID 1,'" & txtPat_ID1.Text & "','" & StrPat_Type & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
  '      MsgBox IntPat_ID
    End If
    con.Close
    
End Sub
Private Sub Srch_Pat_ID1()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID1 1,'" & txtPat_ID1.Text & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
 '       MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Private Sub Flush_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Pat_Info_SELECT1 1,'" & txtPat_ID1.Text & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id
'        MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Private Sub UPat_Pay()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec U_PAT_Pay '" & Mode & "'," & nbrAdv & _
    "," & nbrColl_Fee & "," & txtUnique_ID & ""
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
'    nbrColl_Fee.SetFocus
End Sub
