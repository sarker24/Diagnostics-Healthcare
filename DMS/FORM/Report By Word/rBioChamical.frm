VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form rBio_Chamical 
   Caption         =   "Prime Diagnostic Ltd."
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "rBioChamical.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtU_Name 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   7890
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6690
      Width           =   3510
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rBioChamical.frx":030A
      Height          =   930
      Left            =   1440
      TabIndex        =   20
      Top             =   2670
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   1640
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.ComboBox ComFile_Name 
      Height          =   315
      ItemData        =   "rBioChamical.frx":031F
      Left            =   5910
      List            =   "rBioChamical.frx":0321
      TabIndex        =   12
      Top             =   2760
      Width           =   3825
   End
   Begin VB.TextBox txtDoc_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2001
      Width           =   8280
   End
   Begin VB.TextBox txtAddr 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1604
      Width           =   8280
   End
   Begin VB.TextBox txtSex 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1050
   End
   Begin VB.TextBox nbrAge 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1207
      Width           =   1050
   End
   Begin VB.TextBox txtPat_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Width           =   4155
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3345
      Left            =   300
      TabIndex        =   13
      Top             =   3255
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5900
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"rBioChamical.frx":0323
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   6120
      Width           =   11280
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6435
      TabIndex        =   16
      Top             =   8040
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5385
      TabIndex        =   15
      Top             =   8040
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9585
      TabIndex        =   19
      Top             =   8040
      Width           =   1050
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8535
      TabIndex        =   18
      Top             =   8040
      Width           =   1050
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7500
      TabIndex        =   17
      Top             =   8040
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9240
      Top             =   60
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "1"
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
   Begin VB.CommandButton cmdShow 
      Caption         =   "S&how"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9840
      TabIndex        =   11
      Top             =   2370
      Width           =   1050
   End
   Begin VB.TextBox txtS_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   6885
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1950
      TabIndex        =   9
      Top             =   2400
      Width           =   765
   End
   Begin VB.TextBox txtM_Code 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      TabIndex        =   8
      Top             =   2400
      Width           =   345
   End
   Begin VB.TextBox txtPat_ID 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1470
      TabIndex        =   0
      Top             =   810
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   8520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   25165825
      CurrentDate     =   37114
   End
   Begin MSComCtl2.DTPicker Delv_DATE 
      Height          =   285
      Left            =   8520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   25165825
      CurrentDate     =   37114
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   9240
      Top             =   360
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
      Connect         =   "Provider=SQLOLEDB.1;Password=564;Persist Security Info=True;User ID=sa;Initial Catalog=DIAGNOSTIC;Data Source=EJAZ"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=564;Persist Security Info=True;User ID=sa;Initial Catalog=DIAGNOSTIC;Data Source=EJAZ"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "1"
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   195
      Left            =   7440
      TabIndex        =   30
      Top             =   1230
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   270
      TabIndex        =   29
      Top             =   1650
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      Height          =   195
      Left            =   2610
      TabIndex        =   28
      Top             =   1230
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refered Name"
      Height          =   195
      Left            =   270
      TabIndex        =   27
      Top             =   1950
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      Height          =   195
      Left            =   270
      TabIndex        =   26
      Top             =   1260
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   2610
      TabIndex        =   25
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Test Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3630
      TabIndex        =   24
      Top             =   105
      Width           =   4695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   270
      TabIndex        =   23
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Top             =   810
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7440
      TabIndex        =   21
      Top             =   840
      Width           =   345
   End
End
Attribute VB_Name = "rBio_Chamical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim strM_Code As String
    Dim strS_Code As String
    Dim StrPat_ID As String
    Dim StrPat_ID1 As String
    Dim StrU_Name As String
    Dim strDel_Dt As String
    Dim strDel_Dt1 As String
    
Private Sub cmdClear_Click()
    
    txtPat_ID = ""
    txtPat_Name = ""
    nbrAge = ""
    txtSex = ""
    txtAddr = ""
    txtDoc_Name = ""
    txtS_Code = ""
    txtM_Code = ""
    txtS_Name = ""
    Dt.value = Now
    Delv_DATE.value = Now
    RichTextBox1 = ""
    txtPat_ID.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()

    If Trim(txtPat_ID.Text) = "" Then Exit Sub
    If Trim(txtM_Code.Text) = "" Then Exit Sub
    If Trim(txtS_Code.Text) = "" Then Exit Sub
    strM_Code = txtM_Code
    strS_Code = txtS_Code
    StrPat_ID = txtPat_ID.Text
    StrPat_ID1 = strM_Code + strS_Code + StrPat_ID


    If txtPat_ID.Text = "" Then Exit Sub
    
    Dim path As String
    path = Destin_Path + StrPat_ID1 + ".rtf"
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    Set Doc = WordApp.Documents.Open(path)
    Doc.PrintPreview
    
End Sub

Private Sub cmdPrint_Click()
    If Trim(txtPat_ID.Text) = "" Then Exit Sub
    If Trim(txtM_Code.Text) = "" Then Exit Sub
    If Trim(txtS_Code.Text) = "" Then Exit Sub

    strM_Code = txtM_Code
    strS_Code = txtS_Code
    StrPat_ID = txtPat_ID.Text
    StrPat_ID1 = strM_Code + strS_Code + StrPat_ID

    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    
    Source_Destin

    Set Doc = WordApp.Documents.Open(Destin_Path + StrPat_ID1 + ".rtf")
    Doc.PrintOut
    WordApp.Quit
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID) = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID.SetFocus
        Exit Sub
    End If
    
   
'-----end validation check--------------------------------------------------
    
    strM_Code = txtM_Code
    strS_Code = txtS_Code
    StrPat_ID = txtPat_ID.Text
    StrPat_ID1 = strM_Code + strS_Code + StrPat_ID


    Dim cnn    As New ADODB.Connection
    Dim comm  As New ADODB.Command
    Source_Destin
    
    Me.RichTextBox1.SaveFile Destin_Path + StrPat_ID1 + ".rtf"
'    Set Doc = WordApp.Documents.Open(Destin_Path + StrPat_ID1 + ".rtf")
    comm.CommandText = "exec pass_para 'I','" + txtPat_ID.Text + "','" + _
    txtM_Code + "','" + txtS_Code + "','" + Destin_Path + "'"
    comm.ActiveConnection = strcn.Connection
    comm.Execute
    MsgBox "Document Saved"

End Sub
Private Sub cmdShow_Click()
    If Trim(txtPat_ID.Text) = "" Then Exit Sub
    GetGridData
End Sub
Private Sub ComFile_Name_Click()
    Source_Destin
    'MsgBox Source_Path
    Me.RichTextBox1.LoadFile Source_Path + Me.ComFile_Name + ".rtf"
    'MsgBox "1"
    
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim count_space As Integer
    Dim add_space As String
'-------------------------------------------------------------

    pos1 = InStr(1, Me.RichTextBox1.Text, "Patient ID :")
    Me.RichTextBox1.SelStart = pos1 + 12
    pos2 = CStr(InStr(1, Me.RichTextBox1.Text, "Date :")) - 1
    pos3 = pos2 - (pos1 + 12)
    Me.RichTextBox1.SelLength = pos3

    Me.RichTextBox1.SetFocus
    count_space = pos3 - Len(Trim(Me.txtPat_ID))

    For i = 1 To (count_space - Len(Trim(Me.txtPat_ID)))
        add_space = add_space + Chr(32)
    Next
    Me.RichTextBox1.SelText = Trim(Me.txtPat_ID) + add_space


    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "Date :"))
    Me.RichTextBox1.SelStart = pos1 + 6
    Me.RichTextBox1.SelLength = 1
    Me.RichTextBox1.SetFocus
    Me.RichTextBox1.SelText = CStr(Me.Dt)

'-----------------------------------------------------------------
    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "Delivery Date :"))
    Me.RichTextBox1.SelStart = pos1 + 15
    Me.RichTextBox1.SelLength = 1
    Me.RichTextBox1.SetFocus
    'Me.RichTextBox1.SelText = CStr(Me.Delv_DATE)
    Me.RichTextBox1.SelText = CStr(strDel_Dt1)

'-----------------------------------------------------------------
    Dim nadd_space As String
    pos1 = InStr(1, Me.RichTextBox1.Text, "Patient Name :")
    Me.RichTextBox1.SelStart = pos1 + 14
    pos2 = CStr(InStr(1, Me.RichTextBox1.Text, "Age :")) - 1
    pos3 = pos2 - (pos1 + 14)
    Me.RichTextBox1.SelLength = pos3
    Me.RichTextBox1.SetFocus

    For j = 1 To 29 - Len(Me.txtPat_Name) + 30
       nadd_space = nadd_space + Chr(32)
    Next

    Me.RichTextBox1.SelText = Trim(Me.txtPat_Name) + nadd_space
    
'-----------------------------------------------------------------
    Dim madd_space As String
    pos1 = InStr(1, Me.RichTextBox1.Text, "Age :")
    Me.RichTextBox1.SelStart = pos1 + 5
    pos2 = CStr(InStr(1, Me.RichTextBox1.Text, "Sex :")) - 1
    pos3 = pos2 - (pos1 + 14)
    Me.RichTextBox1.SelLength = pos3
    Me.RichTextBox1.SetFocus
    
    For k = 1 To 6 - Len(Me.txtPat_Name) + 10
       madd_space = madd_space + Chr(32)
    Next
      
    Me.RichTextBox1.SelText = Trim(Me.nbrAge) + madd_space
        
'------------------------------------------------------------------
    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "Sex :"))
    Me.RichTextBox1.SelStart = pos1 + 5
    Me.RichTextBox1.SelLength = 1
    Me.RichTextBox1.SetFocus
    Me.RichTextBox1.SelText = CStr(Me.txtSex)

'------------------------------------------------------------------
'    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "Address :"))
'    Me.RichTextBox1.SelStart = pos1 + 9
'    Me.RichTextBox1.SelLength = 1
'    Me.RichTextBox1.SetFocus
'    Me.RichTextBox1.SelText = CStr(Me.txtAddr)
    
'-----------------------------------------------------------------
    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "Referred by :"))
    Me.RichTextBox1.SelStart = pos1 + 9
    Me.RichTextBox1.SelLength = 1
    Me.RichTextBox1.SetFocus
    Me.RichTextBox1.SelText = CStr(Me.txtDoc_Name)

'-----------------------------------------------------------------
'    pos1 = CStr(InStr(1, Me.RichTextBox1.Text, "D_NA"))
'    Me.RichTextBox1.SelStart = pos1 - 1
'    Me.RichTextBox1.SelLength = 4
'    Me.RichTextBox1.SetFocus
'    Me.RichTextBox1.SelText = StrU_Name

'-----------------------------------------------------------------
   
   Me.Refresh
   'Search_User_Name
End Sub

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rBio_Chamical.DataGrid1.Visible = True Then
        rBio_Chamical.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
ComFile_Name_Add
Search_User_Name
Me.Refresh
End Sub

Private Sub txtM_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtM_Code_LostFocus()

    If txtS_Code = "" Then Exit Sub
    If txtPat_ID.Text = "" Then
        MsgBox "Patient ID Required"
        txtPat_ID.SetFocus
    Exit Sub
    End If
    
    If txtM_Code.Text = "" Then
        MsgBox "Test Code Required"
        txtM_Code.SetFocus
    Exit Sub
    End If
    
    Search_S_Name

End Sub

Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub txtPat_ID_LostFocus()
If Trim(txtPat_ID.Text) = "" Then Exit Sub
Pat_Paid
If Trim(txtPat_ID.Text) = "" Then Exit Sub
'If Trim(txtPat_ID.Text) = "" Then Exit Sub
Search_Pat_Name
End Sub
Private Sub DataGrid1_DblClick()
    txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.Text = DataGrid1.Columns(1)
    txtS_Name.Text = DataGrid1.Columns(2)
    
    
    Adodc5.connectionstring = strcn.Connection
    Adodc5.RecordSource = "exec pass_para 'B','" + Trim(txtPat_ID.Text) + "','" + _
    Trim(txtM_Code) + "','" + Trim(txtS_Code) + "',''"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
        Me.RichTextBox1.LoadFile Adodc5.Recordset.Fields(3)
    End If
    ComFile_Name.SetFocus
    DataGrid1.Visible = False
End Sub

Private Sub txtS_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Pat_Paid()
    Adodc5.connectionstring = strcn.Connection
    Adodc5.RecordSource = "exec Select_Paid 1,'" + txtPat_ID + "'"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
    Else
        MsgBox "The Patient didn't paid"
        txtPat_ID = ""
        txtPat_Name.Text = ""
        Delv_DATE.value = Date
        nbrAge.Text = ""
        txtSex.Text = ""
        txtAddr.Text = ""
        txtDoc_Name.Text = ""
        txtM_Code.Text = ""
        txtS_Code.Text = ""
        txtDoc_Name.Text = ""
        txtPat_ID.SetFocus
    End If
    
End Sub
Private Sub Search_Pat_Name()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_select1 1," + Trim(txtPat_ID.Text) + "", con
    If My_Rst.EOF = False Then
        txtPat_Name.Text = My_Rst!pat_name
        Delv_DATE.value = My_Rst!Delv_Dt
'        Dim strDel_Dt As String
'        Dim strDel_Dt1 As String
        strDel_Dt = Delv_DATE.value
        strDel_Dt1 = Mid(strDel_Dt, 1, 10)
        'MsgBox strDel_Dt1
        
        nbrAge.Text = My_Rst!age
        txtSex.Text = My_Rst!Sex
        txtAddr.Text = My_Rst!addr
        txtDoc_Name.Text = My_Rst!doc_name
    Else
        txtPat_Name.Text = ""
        Delv_DATE.value = Date
        nbrAge.Text = ""
        txtSex.Text = ""
        txtAddr.Text = ""
        txtDoc_Name.Text = ""
        txtPat_ID.SetFocus
    End If
    
    con.Close
    
End Sub
Private Sub ComFile_Name_Add()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "select * from Test_name", con
    
    'If My_Rst.EOF = False Then
    My_Rst.MoveFirst
    While My_Rst.EOF = False
         ComFile_Name.AddItem My_Rst.Fields(0)
         My_Rst.MoveNext
    Wend
    'End If
    con.Close
    
End Sub
Private Sub GetGridData()

    DataGrid1.Visible = True
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_select1 2," + txtPat_ID.Text + ""
    Adodc1.Refresh
    'Set DataGrid1.DataSource = Adodc1.Recordset
    DataGrid1.Columns(0).Width = 500
    DataGrid1.Columns(1).Width = 500
    DataGrid1.Columns(2).Width = 4000
    
End Sub


Private Sub Search_User_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '5','" + u_id + "'", con
    If My_Rst.EOF = False Then
        
        StrU_Name = My_Rst!U_Name
        txtU_Name.Text = My_Rst!U_Name
'        MsgBox StrU_Name
    Else
        txtU_Name.Text = ""
    End If
    
    con.Close
End Sub

Private Sub Search_S_Name()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec S_Name_Select1 1," + txtPat_ID.Text + ",'" + txtM_Code.Text + "','" + txtS_Code.Text + "'", con
    If My_Rst.EOF = False Then
        txtS_Name = My_Rst!s_name
    Else
        txtS_Name = "Invalid Test Name"
    End If
    
    con.Close
End Sub

Private Sub txtS_Code_LostFocus()

    If txtS_Code = "" Then Exit Sub
    If txtPat_ID.Text = "" Then
        MsgBox "Patient ID Required"
        txtPat_ID.SetFocus
    Exit Sub
    End If
    
    If txtM_Code.Text = "" Then
        MsgBox "Test Code Required"
        txtM_Code.SetFocus
    Exit Sub
    End If
    
    Search_S_Name
    
    '-----------show seved document------
    Adodc5.connectionstring = strcn.Connection
    Adodc5.RecordSource = "exec pass_para 'B','" + Trim(txtPat_ID.Text) + "','" + _
    Trim(txtM_Code) + "','" + Trim(txtS_Code) + "',''"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
        Me.RichTextBox1.LoadFile Adodc5.Recordset.Fields(3)
    Else
        RichTextBox1 = ""
    End If
    '-----------------
    
End Sub
