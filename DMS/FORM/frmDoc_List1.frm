VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDoc_List1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commission Pay"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmDoc_List1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRefer_Code 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   240
      MaxLength       =   10
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   5190
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Picture         =   "frmDoc_List1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
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
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1035
   End
   Begin VB.ListBox List 
      Height          =   5715
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   675
      Top             =   5820
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   660
      Top             =   5550
      Visible         =   0   'False
      Width           =   2730
      _ExtentX        =   4815
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
End
Attribute VB_Name = "frmDoc_List1"
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
    Select Case Doc_List_MODE
           Case "frmPatient_Info"
           StrRef_Code = ""
'           StrRef_Code = frmPatient_Info.txtRefer_Code.Text
           StrRef_Code = frmPatient_Info.txtCons_Code.text
           Adodc1.RecordSource = "exec pro_name_SELECT 1,'" & StrRef_Code & "%'"
           'MsgBox StrRef_Code
           'Exit Sub
           Case "rDoc_Pay"
           Adodc1.RecordSource = "exec pro_name_SELECT 1,'" & Trim(rDoc_Pay.txtRefer_Code.text) & "%'"
        
    End Select
    Adodc1.Refresh
       
    Do Until Adodc1.Recordset.EOF = True
        List.AddItem Adodc1.Recordset!doc_name
        Adodc1.Recordset.MoveNext
    Loop

End Sub
Private Sub List_DblClick()
    If Len(Trim(List.text)) = 0 Then Exit Sub
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "SELECT REFER_CODE FROM DOCTOR_INFO WHERE DOC_NAME='" & Trim(List.text) & "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then

        Select Case Doc_List_MODE
               Case "frmPatient_Info"
                     'frmPatient_Info.txtRefer_Code = ""
                     frmPatient_Info.txtCons_Code = Adodc2.Recordset!Refer_code
                     frmPatient_Info.Text2 = Trim(List.text)
                     Unload Me
                     'StrRef_Code = ""
                     frmPatient_Info.txtM_Code.TabStop = True
                     frmPatient_Info.txtRefer_Code.SetFocus
                     'Exit Sub
               Case "rDoc_Pay"
                     rDoc_Pay.txtRefer_Code = ""
                     rDoc_Pay.txtRefer_Code = Adodc2.Recordset!Refer_code
                     rDoc_Pay.txtDoc_Name = Trim(List.text)
'                     rDoc_Pay.Text2 = Trim(List.Text)
                     Unload Me
        End Select
   ' Unload Me
   End If
End Sub
Private Sub List_GotFocus()
'    List.BackColor = &H80000018
End Sub
Private Sub List_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Len(Trim(List.text)) = 0 Then Exit Sub
       Adodc2.connectionstring = strcn.Connection
       Adodc2.RecordSource = "SELECT REFER_CODE FROM DOCTOR_INFO WHERE DOC_NAME='" & Trim(List.text) & "'"
       Adodc2.Refresh
       If Adodc2.Recordset.RecordCount > 0 Then
       Select Case Doc_List_MODE
               Case "frmPatient_Info"
                     'frmPatient_Info.txtRefer_Code = ""
                     frmPatient_Info.txtCons_Code = Adodc2.Recordset!Refer_code
                     frmPatient_Info.Text2 = Trim(List.text)
                     Unload Me
                     frmPatient_Info.txtM_Code.TabStop = True
                     frmPatient_Info.txtCons_Code.SetFocus
               Case "rDoc_Pay"
                     rDoc_Pay.txtRefer_Code = ""
                     rDoc_Pay.txtRefer_Code = Adodc2.Recordset!Refer_code
                     rDoc_Pay.txtDoc_Name = Trim(List.text)
'                     rDoc_Pay.Text2 = Trim(List.Text)
                     
                     Unload Me
       End Select
      ' Unload Me
       End If
       
    End If
End Sub

Private Sub txtSearch_Change()
cmdFind_Click
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub


Private Sub cmdFind_Click()
 List.Clear
    Adodc1.connectionstring = strcn.Connection
    Select Case Doc_List_MODE
           Case "frmPatient_Info"
           StrRef_Code = ""
           StrRef_Code = txtSearch.text
'           Adodc1.RecordSource = "exec pro_name_SELECT 1,'" & StrRef_Code & "%'"
            Adodc1.RecordSource = "select doc_name from doctor_info where doc_name like '" & txtSearch.text & "%'"
           'MsgBox StrRef_Code
           'Exit Sub
           Case "rDoc_Pay"
           Adodc1.RecordSource = "exec pro_name_SELECT 1,'" & Trim(rDoc_Pay.txtRefer_Code.text) & "%'"
        
    End Select
    Adodc1.Refresh
       
    Do Until Adodc1.Recordset.EOF = True
        List.AddItem Adodc1.Recordset!doc_name
        Adodc1.Recordset.MoveNext
    Loop
End Sub



