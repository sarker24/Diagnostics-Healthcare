VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDisc_Edit 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Patient's Discount Modification"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   DrawWidth       =   2
   Icon            =   "Disc_Edit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtPat_Name 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   900
         Width           =   4050
      End
      Begin VB.TextBox txtPat_ID 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   900
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox nbrDisc 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   5610
         TabIndex        =   8
         Top             =   900
         Width           =   765
      End
      Begin VB.TextBox txtUnique_ID 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   390
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtDummy_Pat_ID 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   390
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txtPat_ID1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   270
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   900
         Width           =   1230
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
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   540
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1185
         Left            =   240
         TabIndex        =   11
         Top             =   1110
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   2090
         _Version        =   393216
         AllowUpdate     =   0   'False
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
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4110.236
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   270
         TabIndex        =   15
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1590
         TabIndex        =   13
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   660
         Width           =   630
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2595
      Width           =   885
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2595
      Width           =   930
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   2640
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
Attribute VB_Name = "frmDisc_Edit"
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
On Error GoTo err_sub

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Pat_Info_SELECT 12," & txtPat_ID & ""
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 1280
    DataGrid1.Columns(2).Width = 4080
    DataGrid1.Columns(3).Width = 800
    DataGrid1.Columns(4).Width = 0
    
Exit Sub
err_sub:
    MsgBox Err.Description, vbCritical
    Resume Next
    
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
    
    txtPat_Name = ""
    txtUnique_ID = ""
    
    txtPat_ID1.SetFocus
    
End Sub

Private Sub cmdPatOld_Click()
    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = False
    txtPat_ID.Visible = True
    txtPat_Name = ""
    txtUnique_ID = ""
    txtPat_ID.SetFocus
End Sub

Private Sub cmdSave_Click()
    If txtPat_Name = "" Or txtPat_Name = "0" Then Exit Sub
    If txtUnique_ID = "" Then Exit Sub

    
        Mode = "U"
        UPat_Pay
        MsgBox "Successfully Updated"
        GetGridData
        nbrDisc = "0"
        txtUnique_ID = ""
        txtDummy_Pat_ID = ""
        txtPat_ID = ""
        txtPat_ID1 = ""
        txtPat_Name = ""
        
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
    txtPat_ID.Text = DataGrid1.Columns(0).value
    txtPat_Name = DataGrid1.Columns(2).value
    nbrDisc = DataGrid1.Columns(3).value
    txtUnique_ID = DataGrid1.Columns(4).value
    nbrDisc.SetFocus
    
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
    nbrDisc = "0"
    txtUnique_ID = ""
    
End Sub

Private Sub nbrDisc_Change()
    If Not IsNumeric(nbrDisc.Text) Then
        MsgBox "Only Numaric value allow"
        nbrDisc = 0
        nbrDisc.SelStart = 0
        nbrDisc.SelLength = Len(nbrDisc)
        nbrDisc.SetFocus
    End If
    
End Sub
Private Sub nbrDisc_GotFocus()
        nbrDisc.SelStart = 0
        nbrDisc.SelLength = Len(nbrDisc)
End Sub
Private Sub txtPat_ID_Change()

If txtPat_ID.Visible = True Then
    If Not IsNumeric(txtPat_ID.Text) Then
       ' MsgBox "Only Numaric value allow"
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
    GetGridData
End Sub
Private Sub UPat_Pay()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec U_PAT_Disc '" & Mode & "'," & nbrDisc & "," & txtUnique_ID & ""
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub txtPat_ID1_LostFocus()

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

    GetGridData
    
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

