VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCPaySearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Doctor Payment Search"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "frmCPaySearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Doctor Payment Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6015
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   10610
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
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4320
      Picture         =   "frmCPaySearch.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5400
      Picture         =   "frmCPaySearch.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6360
      Picture         =   "frmCPaySearch.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Available Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
End
Attribute VB_Name = "frmCPaySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
'If txtSearch.text = "Serial Number" Then

    Adodc1.connectionstring = cn.connectionstring
    Adodc1.RecordSource = "SELECT AID,AHName,Department,AHType  " & _
                 "FROM AccountsHead WHERE AccountsHead.AHName LIKE '" & txtSearch.text & "%'"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1.Recordset

      
'ElseIf txtSearch.text = "Date Search" Then
'
'    Adodc1.ConnectionString = cn.ConnectionString
'    Adodc1.RecordSource = "SELECT AID as [Bill No],strDate as Date,TableName,WaiterName,Name " & _
'                          "FROM AccountsHead WHERE AccountsHead.strDate = '" & dtDate.Date & "'"
'    Adodc1.Refresh
'    Set DataGrid1.DataSource = Adodc1.Recordset
'
'
'ElseIf txtSearch.text = "Table Number" Then
'
'    Adodc1.ConnectionString = cn.ConnectionString
'    Adodc1.RecordSource = "SELECT AID as [Bill No],strDate as Date,TableName,WaiterName,Name " & _
'                          "FROM AccountsHead WHERE AccountsHead.TableName LIKE '" & txtSearch.text & "%'"
'    Adodc1.Refresh
'    Set DataGrid1.DataSource = Adodc1.Recordset
'
'ElseIf txtSearch.text = "Waiter Name" Then
'
'    Adodc1.ConnectionString = cn.ConnectionString
'    Adodc1.RecordSource = "SELECT AID as [Bill No],strDate as Date,TableName,WaiterName,Name " & _
'                 "FROM AccountsHead WHERE AccountsHead.WaiterName LIKE '" & txtSearch.text & "%'"
'    Adodc1.Refresh
'    Set DataGrid1.DataSource = Adodc1.Recordset
'
'Else
'
'    Adodc1.ConnectionString = cn.ConnectionString
'    Adodc1.RecordSource = "SELECT AID as [Bill No],strDate as Date,TableName,WaiterName,Name " & _
'                 "FROM AccountsHead WHERE AccountsHead.Name LIKE '" & txtSearch.text & "%'"
'    Adodc1.Refresh
'    Set DataGrid1.DataSource = Adodc1.Recordset
'
'End If
End Sub

Private Sub cmdOk_Click()
    If DataGrid1.Row < 0 Then
        MsgBox "Please Select a Doctor Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me

Set frmCPaySearch = Nothing
End Sub

Private Sub DataGrid1_DblClick()
Dim adu As Integer
adu = DataGrid1.Row

'End If
cmdOk_Click
End Sub

Private Sub dtDate_click()
cmdFind_Click
End Sub

Private Sub Form_Load()
   
GetGridData
DataGrid1.MarqueeStyle = dbgHighlightRow
Adodc1.Visible = False
End Sub

Private Sub GetGridData()
     Adodc1.connectionstring = cn.connectionstring
    Adodc1.RecordSource = "SELECT SerialNo,PDate,refer_code,Doc_Name,Amount,Pay_To,UName,strTime FROM Commission_Pay"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1.Recordset
    
'    DataGrid1.Columns(4).Visible = True
    
    
    DataGrid1.Columns(0).Caption = "Serial No"
    DataGrid1.Columns(1).Caption = "Doctor Name"
    DataGrid1.Columns(2).Caption = "Amount"
    DataGrid1.Columns(3).Caption = "Payment to"
    
    DataGrid1.Columns(1).Width = 3000
    DataGrid1.Columns(2).Width = 2000
    DataGrid1.Columns(3).Width = 2000
    
       DataGrid1.BackColor = &HC0B4A9

End Sub
  
    Private Sub PopulateCompanySearch()
        If DataGrid1.Row > -1 Then
        

              frmCommissionPay.PopulateCnf DataGrid1.Columns(0).text
        End If
    End Sub
    
Private Sub txtSearch_Change()
cmdFind_Click
End Sub



