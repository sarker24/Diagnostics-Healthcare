VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTest_List 
   Caption         =   "Diagnostic management system"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   Icon            =   "Test_List.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1860
      Top             =   1665
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
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
   Begin VB.CommandButton cmdclose 
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
      Height          =   660
      Left            =   -120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4410
      Width           =   9990
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Test_List.frx":030A
      Height          =   4425
      Left            =   -45
      TabIndex        =   0
      Top             =   -45
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7805
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2849.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTest_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
       
        Select Case Test_List_Mode
              Case "frmPatient_Info_M"
                    frmPatient_Info.txtM_Code = frmTest_List.DataGrid1.Columns(0)
                    frmPatient_Info.txtS_Code = frmTest_List.DataGrid1.Columns(2)
                    frmPatient_Info.txtS_Name = frmTest_List.DataGrid1.Columns(3)
                    frmPatient_Info.nbrTest_Rate = frmTest_List.DataGrid1.Columns(4)
                    Unload Me
                    frmPatient_Info.nbrTest_Rate.SetFocus
              Case "frmPatient_Info_S"
                    frmPatient_Info.txtM_Code = frmTest_List.DataGrid1.Columns(0)
                    frmPatient_Info.txtS_Code = frmTest_List.DataGrid1.Columns(2)
                    frmPatient_Info.txtS_Name = frmTest_List.DataGrid1.Columns(3)
                    frmPatient_Info.nbrTest_Rate = frmTest_List.DataGrid1.Columns(4)
                    Unload Me
                    frmPatient_Info.nbrTest_Rate.SetFocus
        End Select
End Sub
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then
'Select Case Test_List_Mode
'
'              Case "frmPatient_Info_M"
'                    frmPatient_Info.txtM_Code = frmTest_List.DataGrid1.Columns(0)
'                    frmPatient_Info.txtS_Code = frmTest_List.DataGrid1.Columns(2)
'                    frmPatient_Info.txtS_Name = frmTest_List.DataGrid1.Columns(3)
'                    frmPatient_Info.nbrTest_Rate = frmTest_List.DataGrid1.Columns(4)
'
'                    Unload Me
'                    frmPatient_Info.nbrTest_Rate.SetFocus
'              Case "frmPatient_Info_S"
'                    frmPatient_Info.txtM_Code = frmTest_List.DataGrid1.Columns(0)
'                    frmPatient_Info.txtS_Code = frmTest_List.DataGrid1.Columns(2)
'                    frmPatient_Info.txtS_Name = frmTest_List.DataGrid1.Columns(3)
'                    frmPatient_Info.nbrTest_Rate = frmTest_List.DataGrid1.Columns(4)
'                    Unload Me
'                    frmPatient_Info.nbrTest_Rate.SetFocus
'        End Select
'
'End If
End Sub

Private Sub Form_Load()
    
        
    Adodc1.connectionstring = strcn.Connection
    Select Case Test_List_Mode
            Case "frmPatient_Info_M"
                  Adodc1.RecordSource = "exec pro_name_SELECT 2,'" & Trim(frmPatient_Info.txtM_Code.Text) & "%'"
            Case "frmPatient_Info_S"
                  Adodc1.RecordSource = "exec pro_name_SELECT 3,'" & Trim(frmPatient_Info.txtS_Code.Text) & "%'"
    End Select
   
    Adodc1.Refresh
    Do Until Adodc1.Recordset.EOF = True
        DataGrid1.Columns(0) = Adodc1.Recordset!m_code

        Adodc1.Recordset.MoveNext
    Loop
    DataGrid1.Refresh

    DataGrid1.Columns(0).Width = 600
    DataGrid1.Columns(1).Width = 4000
    DataGrid1.Columns(2).Width = 600
    DataGrid1.Columns(3).Width = 4000
    DataGrid1.Columns(4).Width = 500
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTest_List = Nothing
End Sub
'Private Sub List_GotFocus()
'List.BackColor = &H80000018
'End Sub
Private Sub List_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        If Len(Trim(List.Text)) = 0 Then Exit Sub
'
'        Dim intCnt As String
'        Dim intLen As String
'        Dim Prod_Code_LEN As String
'        Dim StrProd_Name As String
'
'        intCnt = SpaceX(List.Text)
'        Prod_Code_LEN = Trim(Left$(List.Text, intCnt))
'        intLen = Len(Trim(List.Text))
'        StrProd_Name = Trim(Right$(List.Text, (intLen - intCnt)))
'        '--------------------------------------------------------------------
'
'      Select Case Prod_src_mode
'            Case "frmGRN"
'                  frmGRN.txtProd_Code = ""
'                  frmGRN.txtProd_Code = Prod_Code_LEN
'                  frmGRN.comProd_Name = StrProd_Name
'                  Unload Me
'            Case "frmChallan"
'                  frmChallan.txtProd_Code = Prod_Code_LEN
'                  frmChallan.comProd_Name = StrProd_Name
'                  Unload Me
'
'        End Select
'
'    End If
End Sub

