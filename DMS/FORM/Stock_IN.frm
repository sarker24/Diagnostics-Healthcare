VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStock_IN 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   Icon            =   "Stock_IN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox nbrAmt 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5910
      TabIndex        =   7
      Top             =   2280
      Width           =   1260
   End
   Begin VB.TextBox txtInv_No 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2040
      TabIndex        =   0
      Top             =   660
      Width           =   1245
   End
   Begin VB.TextBox nbrTest_per_Box 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3450
      TabIndex        =   6
      Top             =   2310
      Width           =   1260
   End
   Begin MSComCtl2.DTPicker Pur_Date 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   37365
   End
   Begin VB.TextBox txtSup_ID 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2040
      TabIndex        =   3
      Top             =   1740
      Width           =   1245
   End
   Begin VB.TextBox txtSup_Name 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3750
   End
   Begin VB.TextBox nbrNo_of_Box 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   1260
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
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3510
      Width           =   990
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
      Height          =   300
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3510
      Width           =   990
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
      Height          =   300
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3510
      Width           =   990
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3510
      Width           =   990
   End
   Begin VB.TextBox txtItem_Name 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1230
      Width           =   3750
   End
   Begin VB.TextBox txtItem_Code 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2040
      TabIndex        =   1
      Top             =   1260
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5340
      Top             =   150
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
   Begin MSComCtl2.DTPicker Exp_Date 
      Height          =   285
      Left            =   3450
      TabIndex        =   9
      Top             =   2820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   37365
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5280
      Top             =   150
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount (tk.)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5880
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   23
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Per Box"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3420
      TabIndex        =   22
      Top             =   2070
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock In"
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
      Left            =   510
      TabIndex        =   21
      Top             =   270
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expire Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3870
      TabIndex        =   20
      Top             =   2610
      Width           =   825
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Id"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   18
      Top             =   1530
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   17
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Box"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   16
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   14
      Top             =   2610
      Width           =   1065
   End
End
Attribute VB_Name = "frmStock_IN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    txtInv_No = ""
    txtItem_Code = ""
    txtItm_Name = ""
    txtSup_ID = ""
    txtSup_Name = ""
    nbrNo_of_Box = ""
    nbrTest_per_Box = ""
    Pur_Date.value = Date
    Exp_Date.value = Date
    txtInv_No.SetFocus
End Sub

Private Sub cmdSave_Click()

If txtInv_No.Text = "" Then Exit Sub

If txtItem_Name = "" Then
    MsgBox "Item Name Required"
    txtItem_Code.SetFocus
    Exit Sub
End If

If txtSup_Name = "" Then
    MsgBox "Supplier Name Requred"
    txtSup_ID.SetFocus
    Exit Sub
End If

If nbrNo_of_Box = "0" Then
    MsgBox "Quantity required"
    nbrNo_of_Box.SetFocus
    Exit Sub
End If



'    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Inv_NO_SELECT '1','" & Trim(txtInv_No.Text) & "','" & Trim(txtItem_Code.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Stock_Trans
        MsgBox "Updated Successfully"
    Else
        Mode = "I"
        Stock_Trans
        MsgBox "Inserted Successfully"
    End If
    
'    txtInv_No = ""
    txtItem_Code = ""
    txtItem_Name = ""
    txtSup_ID = ""
    txtSup_Name = ""
    nbrNo_of_Box = ""
    nbrTest_per_Box = ""
    nbrAmt.Text = "0"
    Pur_Date.value = Date
    Exp_Date.value = Date
    txtInv_No.SetFocus
    
End Sub
Private Sub Stock_Trans()
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Stock_In_IUD '" + Mode + "','" + txtInv_No.Text + _
    "','" + txtItem_Code.Text + _
    "','" + txtSup_ID.Text + _
    "'," + nbrNo_of_Box + _
    "," + nbrTest_per_Box + _
    "," + nbrAmt.Text + _
    ",'" + Format(Pur_Date, "yyyy-mm-dd") + _
    "','" + Format(Exp_Date, "yyyy-mm-dd") + _
    "','" + u_id + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
    
    If KeyAscii = 27 Then
    Unload Me
    End If
    
End Sub

'Private Sub Search_Item_Name()
'    Dim My_Rst As New ADODB.Recordset
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'
'    My_Rst.Open "exec pro_name_SELECT '11','" + txtItem_Code + "'", con
'    If My_Rst.EOF = False Then
'        txtItem_Code.Text = My_Rst!item_code
'        txtItem_Name.Text = My_Rst!item_name
'    Else
'        txtItem_Name.Text = ""
'    End If
'
'    con.Close
'End Sub

Private Sub Form_Load()

    Pur_Date = Date
    Exp_Date = Date
    
End Sub

Private Sub nbrAmt_Change()
    If Not IsNumeric(nbrAmt.Text) Then
        'MsgBox "Only Numaric value allow"
        nbrAmt = 0
        nbrAmt.SelStart = 0
        nbrAmt.SelLength = Len(nbrAmt)
        nbrAmt.SetFocus
    End If
End Sub

Private Sub nbrAmt_GotFocus()
    nbrAmt.SelStart = 0
    nbrAmt.SelLength = Len(nbrAmt)
End Sub

Private Sub nbrNo_of_Box_Change()
    If Not IsNumeric(nbrNo_of_Box.Text) Then
        'MsgBox "Only Numaric value allow"
        nbrNo_of_Box = 0
        nbrNo_of_Box.SelStart = 0
        nbrNo_of_Box.SelLength = Len(nbrNo_of_Box)
        nbrNo_of_Box.SetFocus
    End If
End Sub

Private Sub nbrNo_of_Box_GotFocus()

    nbrNo_of_Box.SelStart = 0
    nbrNo_of_Box.SelLength = Len(nbrNo_of_Box)
    
End Sub

Private Sub nbrTest_per_Box_Change()

        If Not IsNumeric(nbrTest_per_Box.Text) Then
        'MsgBox "Only Numaric value allow"
        nbrTest_per_Box = 0
        nbrTest_per_Box.SelStart = 0
        nbrTest_per_Box.SelLength = Len(nbrTest_per_Box)
        nbrTest_per_Box.SetFocus
    End If

End Sub

Private Sub nbrTest_per_Box_GotFocus()
    nbrTest_per_Box.SelStart = 0
    nbrTest_per_Box.SelLength = Len(nbrTest_per_Box)
End Sub

Private Sub txtItem_Code_LostFocus()
    If txtItem_Code = "" Then Exit Sub
    
    Item_List_MODE = "frmStock_IN"
    
    
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
Private Sub Search_Sup_Name()
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

Private Sub txtSup_ID_LostFocus()
    If txtSup_ID = "" Then Exit Sub
    Search_Sup_Name
End Sub
