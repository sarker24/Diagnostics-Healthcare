VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmp_List 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Diagnostic management system"
   ClientHeight    =   3630
   ClientLeft      =   5040
   ClientTop       =   4530
   ClientWidth     =   5235
   Icon            =   "Emp_List.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   5295
   End
   Begin VB.TextBox txtRefer_Code 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      MaxLength       =   10
      TabIndex        =   2
      Top             =   315
      Visible         =   0   'False
      Width           =   5190
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   1215
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
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
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3150
      Width           =   5235
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   225
      Top             =   2025
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
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
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   2370
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5235
   End
End
Attribute VB_Name = "frmEmp_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Form_Load()

    List.Clear
    
    Adodc1.connectionstring = strcn.Connection
    Select Case Emp_List_MODE
           Case "rPat_Info"
           strEmp_Code = "rPat_Info"
           strEmp_Code = txtSearch.text
'           Adodc1.RecordSource = "exec pro_name_SELECT 1,'" & StrRef_Code & "%'"
            Adodc1.RecordSource = "select Emp_name from Emp_Info where Emp_name like '" & txtSearch.text & "%'"
           MsgBox StrRef_Code
           Exit Sub
           Case "rPat_Info"
           Adodc1.RecordSource = "exec pro_name_SELECT 13,'" & Trim(rPat_Info.txtEmp_ID.text) & "%'"
        
    End Select
    Adodc1.Refresh
'    Adodc1.connectionstring = strcn.Connection
'    Select Case Emp_List_MODE
'
'           Case "frmEmp_Info"
'                strEmp_Code = ""
'                strEmp_Code = frmEmp_Info.txtEmp_ID.text
'                Adodc1.RecordSource = "exec pro_name_SELECT '9','" & strEmp_Code & "%'"
'
'    End Select
'    Adodc1.Refresh
       
    Do Until Adodc1.Recordset.EOF = True
        List.AddItem Adodc1.Recordset!Emp_Name
        Adodc1.Recordset.MoveNext
    Loop

End Sub
Private Sub List_DblClick()
    GET_Emp
End Sub
Private Sub List_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       GET_Emp
    End If
End Sub

Private Sub GET_Emp()
    
    If Len(Trim(List.text)) = 0 Then Exit Sub
    Adodc2.connectionstring = strcn.Connection
   
    Adodc2.RecordSource = "exec pro_name_SELECT '22','" & Trim(List.text) & "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        
        Select Case Emp_List_MODE
               'Case "frmStock_IN"
   
                '     frmStock_IN.txtItem_Code = Adodc2.Recordset!item_code
                '     frmStock_IN.txtItem_Name = Trim(List.Text)
                '     Unload Me
   
               Case "frmStock_Out"
                     
                     frmStock_Out.txtEmp_ID = Adodc2.Recordset!emp_id
                     frmStock_Out.txtEmp_Name = Trim(List.text)
                     Unload Me
                     
        End Select
   
   End If

End Sub
