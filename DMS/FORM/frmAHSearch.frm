VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAHSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Accounts Head Details"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmAHSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Accounts Head Name Details"
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
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin MSDataGridLib.DataGrid fgExport 
         Height          =   6735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   11880
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            Caption         =   "Accounts ID"
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
            Caption         =   "Accounts Head Name"
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
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3495.118
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   7440
      Width           =   2055
   End
End
Attribute VB_Name = "frmAHSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private rsTemp                      As ADODB.Recordset
'Private rsExport                    As ADODB.Recordset
'Private rsfactory                   As New ADODB.Recordset
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdFind_Click()
''frmLedgerParty.Show vbModal
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'
''            rsTemp.CursorLocation = adUseClient
'     rsTemp.Open "SELECT AYID,AHName,AHType " & _
'                 "FROM AccountsHead WHERE AccountsHead.AHName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
'
''   End If
'
'         fgExport.Rows = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("AYID") & vbTab & rsTemp("AHName") & _
'         vbTab & rsTemp("AHType")
'
'        rsTemp.MoveNext
'        Wend
'
'End Sub
'
'Private Sub cmdOK_Click()
'    If fgExport.RowSel < 0 Then
'        MsgBox "Please Select a Menu Group From the List."
'        Exit Sub
'    End If
'
'     Call PopulateCompanySearch
'
'Unload Me
'Set frmAccountsHeadSearch = Nothing
'End Sub
'
'
'
'Private Sub fgExport_DblClick()
'    cmdOK_Click
'End Sub
'
'Private Sub Form_Load()
'     ModFunction.StartUpPosition Me
'     Set rsTemp = New ADODB.Recordset
'     rsTemp.CursorLocation = adUseClient
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'
''            rsTemp.CursorLocation = adUseClient
'     rsTemp.Open "SELECT AID,AHName,AHType FROM AccountsHead", cn, adOpenStatic, adLockReadOnly
'
''   End If
'
'         fgExport.Row = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("AYID") & vbTab & rsTemp("AHName") & _
'         vbTab & rsTemp("AHType")
'
'        rsTemp.MoveNext
'    Wend
'     GridCount fgExport
''     If fgExport.Rows = 1 Then fgExport.AddItem ""
'
'
'End Sub
'
'
'
'    Private Sub PopulateCompanySearch()
'        If fgExport.Row > 0 Then
'
'             frmAHSearch.PopulateIteam fgExport.TextMatrix(fgExport.Row, 1)
'        End If
'    End Sub
'
'Private Sub txtSearch_Change()
'cmdFind_Click
'End Sub
'
'
